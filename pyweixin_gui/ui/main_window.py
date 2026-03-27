from __future__ import annotations

import logging

from PySide6.QtCore import QThread, Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QInputDialog,
    QListWidget,
    QListWidgetItem,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSplitter,
    QStackedWidget,
    QStatusBar,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from ..error_handling import UiError
from ..executor import BatchExecutor, failed_rows_from_execution
from ..models import AppSettings, FileBatchRow, MessageBatchRow, TaskTemplate, TaskType, dataclass_from_json, dataclass_to_json
from ..settings_manager import SettingsManager
from ..storage import AppStorage
from ..worker import BatchWorker
from .widgets import BatchPage, DashboardPage, HistoryPage, SettingsPage, TemplatesPage


class MainWindow(QMainWindow):
    def __init__(
        self,
        adapter,
        storage: AppStorage,
        settings_manager: SettingsManager,
        settings: AppSettings,
        logger: logging.Logger,
    ):
        super().__init__()
        self.adapter = adapter
        self.storage = storage
        self.settings_manager = settings_manager
        self.settings = settings
        self.logger = logger
        self.executor = BatchExecutor(adapter)
        self.worker_thread: QThread | None = None
        self.worker: BatchWorker | None = None
        self._template_cache: list[TaskTemplate] = []
        self._execution_cache = []

        self.setWindowTitle("PyWeChat GUI 工作台")
        self.resize(self.settings.window_width, self.settings.window_height)
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        root = QWidget()
        root_layout = QHBoxLayout(root)
        root_layout.setContentsMargins(12, 12, 12, 12)
        root_layout.setSpacing(12)

        splitter = QSplitter()
        root_layout.addWidget(splitter)

        self.nav = QListWidget()
        self.nav.setObjectName("SideNav")
        self.nav.setFixedWidth(200)
        self.nav.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.stack = QStackedWidget()

        splitter.addWidget(self.nav)
        splitter.addWidget(self.stack)
        splitter.setStretchFactor(1, 1)
        self.setCentralWidget(root)

        self.dashboard_page = DashboardPage()
        self.message_page = BatchPage(TaskType.MESSAGE, "批量消息")
        self.file_page = BatchPage(TaskType.FILE, "批量文件")
        self.templates_page = TemplatesPage()
        self.history_page = HistoryPage()
        self.settings_page = SettingsPage()

        self._add_page("首页", self.dashboard_page)
        self._add_page("批量消息", self.message_page)
        self._add_page("批量文件", self.file_page)
        self._add_page("模板中心", self.templates_page)
        self._add_page("执行历史", self.history_page)
        self._add_page("设置", self.settings_page)
        self.nav.setCurrentRow(0)
        self.nav.currentRowChanged.connect(self.stack.setCurrentIndex)

        self._bind_events()
        self._load_settings_to_form()
        self.refresh_environment()
        self.refresh_templates()
        self.refresh_history()
        self._show_welcome_if_needed()
        self._set_running_state(False)

    def _add_page(self, title: str, page: QWidget) -> None:
        self.nav.addItem(QListWidgetItem(title))
        self.stack.addWidget(page)

    def _show_welcome_if_needed(self) -> None:
        if self.settings.first_run_risk_ack:
            return
        QMessageBox.information(
            self,
            "首次使用提示",
            "欢迎使用 PyWeChat 办公助手。\n\n"
            "推荐你按这个顺序操作：\n"
            "1. 首页检查微信状态\n"
            "2. 去批量页面导入或填写表格\n"
            "3. 先校验，再执行\n\n"
            "执行期间请不要手动操作微信。",
        )
        self.settings.first_run_risk_ack = True
        self.settings_manager.save(self.settings)

    def _bind_events(self) -> None:
        self.dashboard_page.refresh_requested.connect(self.refresh_environment)
        self.dashboard_page.open_message_requested.connect(lambda: self.nav.setCurrentRow(1))
        self.dashboard_page.open_file_requested.connect(lambda: self.nav.setCurrentRow(2))
        self.dashboard_page.open_templates_requested.connect(lambda: self.nav.setCurrentRow(3))
        self.message_page.run_requested.connect(lambda rows, src: self.start_batch(TaskType.MESSAGE, rows, src))
        self.file_page.run_requested.connect(lambda rows, src: self.start_batch(TaskType.FILE, rows, src))
        self.message_page.stop_requested.connect(self.stop_batch)
        self.file_page.stop_requested.connect(self.stop_batch)
        self.message_page.save_template_requested.connect(self.save_template)
        self.file_page.save_template_requested.connect(self.save_template)
        self.message_page.open_templates_requested.connect(self.open_templates_for)
        self.file_page.open_templates_requested.connect(self.open_templates_for)

        self.templates_page.refresh_button.clicked.connect(self.refresh_templates)
        self.templates_page.search_input.textChanged.connect(self.apply_template_filter)
        self.templates_page.load_button.clicked.connect(self.load_selected_template)
        self.templates_page.delete_button.clicked.connect(self.delete_selected_template)
        self.templates_page.rename_button.clicked.connect(self.rename_selected_template)
        self.templates_page.duplicate_button.clicked.connect(self.duplicate_selected_template)
        self.templates_page.table.itemDoubleClicked.connect(lambda _item: self.load_selected_template())

        self.history_page.refresh_button.clicked.connect(self.refresh_history)
        self.history_page.search_input.textChanged.connect(self.apply_history_filter)
        self.history_page.clear_button.clicked.connect(self.clear_history)
        self.history_page.execution_table.itemSelectionChanged.connect(self.show_execution_details)
        self.history_page.detail_table.itemSelectionChanged.connect(self.show_selected_row_diagnostic)
        self.history_page.copy_diag_button.clicked.connect(self.copy_selected_row_diagnostic)
        self.history_page.retry_button.clicked.connect(self.retry_selected_execution_failures)

        self.settings_page.save_button.clicked.connect(self.save_settings)

    def refresh_environment(self) -> None:
        status = self.adapter.inspect_environment()
        self.dashboard_page.set_environment(status)
        self.status_bar.showMessage(status.status_message, 5000)

    def refresh_templates(self) -> None:
        self._template_cache = self.storage.list_templates()
        self.apply_template_filter()

    def apply_template_filter(self) -> None:
        query = self.templates_page.search_input.text().strip().lower()
        templates = self._template_cache
        table = self.templates_page.table
        table.setRowCount(0)
        filtered = []
        for template in templates:
            type_label = "批量消息" if template.task_type is TaskType.MESSAGE else "批量文件"
            haystack = " ".join([template.name, type_label, template.updated_at or ""]).lower()
            if query and query not in haystack:
                continue
            filtered.append(template)
        for row_index, template in enumerate(filtered):
            table.insertRow(row_index)
            values = [
                str(template.id),
                template.name,
                "批量消息" if template.task_type is TaskType.MESSAGE else "批量文件",
                template.updated_at or "",
            ]
            for column_index, value in enumerate(values):
                table.setItem(row_index, column_index, QTableWidgetItem(value))
        table.resizeColumnsToContents()
        if filtered:
            self.templates_page.summary_label.setText(
                f"当前显示 {len(filtered)} / {len(templates)} 个模板。双击或选中后可加载到工作台。"
            )
        else:
            if templates:
                self.templates_page.summary_label.setText("没有匹配到模板，请换个关键词再试。")
            else:
                self.templates_page.summary_label.setText("还没有模板。你可以先去批量页面整理一份任务，再点“保存模板”。")

    def refresh_history(self) -> None:
        self._execution_cache = self.storage.list_executions()
        self.apply_history_filter()

    def apply_history_filter(self) -> None:
        query = self.history_page.search_input.text().strip().lower()
        executions = self._execution_cache
        table = self.history_page.execution_table
        table.setRowCount(0)
        filtered = []
        for execution in executions:
            type_label = "批量消息" if execution.task_type is TaskType.MESSAGE else "批量文件"
            haystack = " ".join(
                [
                    str(execution.id or ""),
                    type_label,
                    execution.started_at,
                    execution.status,
                    str(execution.success_count),
                    str(execution.failure_count),
                ]
            ).lower()
            if query and query not in haystack:
                continue
            filtered.append(execution)
        for row_index, execution in enumerate(filtered):
            table.insertRow(row_index)
            values = [
                str(execution.id),
                "批量消息" if execution.task_type is TaskType.MESSAGE else "批量文件",
                execution.started_at,
                execution.status,
                str(execution.row_count),
                str(execution.success_count),
                str(execution.failure_count),
            ]
            for column_index, value in enumerate(values):
                item = QTableWidgetItem(value)
                if column_index == 3 and execution.status == "completed" and execution.failure_count == 0:
                    item.setBackground(QColor("#dcfce7"))
                elif column_index == 3 and execution.failure_count > 0:
                    item.setBackground(QColor("#fef3c7"))
                table.setItem(row_index, column_index, item)
        table.resizeColumnsToContents()
        self.history_page.detail_table.setRowCount(0)
        self.history_page.diagnostic_text.clear()
        success_runs = sum(1 for execution in executions if execution.failure_count == 0 and execution.status == "completed")
        failed_runs = sum(1 for execution in executions if execution.failure_count > 0)
        self.history_page.total_runs_label.value_label.setText(str(len(executions)))  # type: ignore[attr-defined]
        self.history_page.success_runs_label.value_label.setText(str(success_runs))  # type: ignore[attr-defined]
        self.history_page.failed_runs_label.value_label.setText(str(failed_runs))  # type: ignore[attr-defined]
        if filtered:
            self.history_page.summary_label.setText(
                f"当前显示 {len(filtered)} / {len(executions)} 条执行记录。先选中一条，再看下方逐行结果。"
            )
        else:
            if executions:
                self.history_page.summary_label.setText("没有匹配到执行记录，请换个关键词再试。")
            else:
                self.history_page.summary_label.setText("还没有执行历史。第一次成功或失败执行后，这里会显示完整记录。")

    def _load_settings_to_form(self) -> None:
        page = self.settings_page
        page.is_maximize.setChecked(self.settings.is_maximize)
        page.close_weixin.setChecked(self.settings.close_weixin)
        page.clear.setChecked(self.settings.clear)
        page.search_pages.setValue(self.settings.search_pages)
        page.send_delay.setValue(self.settings.send_delay)
        page.window_width.setValue(self.settings.window_width)
        page.window_height.setValue(self.settings.window_height)
        page.import_encoding.setCurrentText(self.settings.import_encoding)
        page.theme.setCurrentText(self.settings.theme)

    def save_settings(self) -> None:
        page = self.settings_page
        self.settings = AppSettings(
            is_maximize=page.is_maximize.isChecked(),
            close_weixin=page.close_weixin.isChecked(),
            search_pages=page.search_pages.value(),
            send_delay=page.send_delay.value(),
            clear=page.clear.isChecked(),
            window_width=page.window_width.value(),
            window_height=page.window_height.value(),
            import_encoding=page.import_encoding.currentText(),
            theme=page.theme.currentText(),
            history_retention=self.settings.history_retention,
            first_run_risk_ack=True,
        )
        self.settings_manager.save(self.settings)
        self.resize(self.settings.window_width, self.settings.window_height)
        self.status_bar.showMessage("设置已保存", 4000)

    def save_template(self, task_type: TaskType, rows: list[MessageBatchRow] | list[FileBatchRow]) -> None:
        name, ok = QInputDialog.getText(self, "保存模板", "模板名称")
        if not ok or not name.strip():
            return
        template = TaskTemplate(name=name.strip(), task_type=task_type, rows_json=dataclass_to_json(rows))
        self.storage.save_template(template)
        self.refresh_templates()
        self.status_bar.showMessage("模板已保存，下次可以在模板中心直接加载。", 4000)

    def open_templates_for(self, task_type: TaskType) -> None:
        self.nav.setCurrentRow(3)
        table = self.templates_page.table
        for row_index in range(table.rowCount()):
            if table.item(row_index, 2).text() == ("批量消息" if task_type is TaskType.MESSAGE else "批量文件"):
                table.selectRow(row_index)
                break

    def _selected_template_id(self) -> int | None:
        table = self.templates_page.table
        selected = table.selectedItems()
        if not selected:
            return None
        return int(table.item(selected[0].row(), 0).text())

    def load_selected_template(self) -> None:
        template_id = self._selected_template_id()
        if template_id is None:
            QMessageBox.information(self, "模板中心", "请先选中一个模板。")
            return
        template = self.storage.get_template(template_id)
        rows = dataclass_from_json(template.task_type, template.rows_json)
        if template.task_type is TaskType.MESSAGE:
            self.message_page.load_rows(rows)
            self.nav.setCurrentRow(1)
        else:
            self.file_page.load_rows(rows)
            self.nav.setCurrentRow(2)
        self.status_bar.showMessage(f"已加载模板：{template.name}", 5000)

    def delete_selected_template(self) -> None:
        template_id = self._selected_template_id()
        if template_id is None:
            return
        answer = QMessageBox.question(self, "删除模板", "确定删除这个模板吗？删除后不可恢复。")
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.storage.delete_template(template_id)
        self.refresh_templates()
        self.status_bar.showMessage("模板已删除", 4000)

    def rename_selected_template(self) -> None:
        template_id = self._selected_template_id()
        if template_id is None:
            return
        template = self.storage.get_template(template_id)
        name, ok = QInputDialog.getText(self, "重命名模板", "模板名称", text=template.name)
        if not ok or not name.strip():
            return
        template.name = name.strip()
        self.storage.save_template(template)
        self.refresh_templates()

    def duplicate_selected_template(self) -> None:
        template_id = self._selected_template_id()
        if template_id is None:
            return
        template = self.storage.get_template(template_id)
        name, ok = QInputDialog.getText(self, "复制模板", "新模板名称", text=f"{template.name}-副本")
        if not ok or not name.strip():
            return
        self.storage.duplicate_template(template_id, name.strip())
        self.refresh_templates()

    def start_batch(
        self,
        task_type: TaskType,
        rows: list[MessageBatchRow] | list[FileBatchRow],
        source_execution_id: int | None = None,
    ) -> None:
        if self.worker_thread is not None:
            QMessageBox.warning(self, "任务进行中", "当前已有任务在执行，请等待完成或先停止。")
            return
        environment = self.adapter.inspect_environment()
        if environment.login_status != "已登录":
            self.nav.setCurrentRow(0)
            self._show_guidance_dialog(
                title="暂时不能执行任务",
                message=environment.status_message,
                suggestion="\n".join(environment.advice) if environment.advice else "请先解决首页提示的问题，再回来执行。",
            )
            return
        runtime_options = self.settings.runtime_options()
        self.worker_thread = QThread(self)
        self.worker = BatchWorker(self.executor, task_type, rows, runtime_options, source_execution_id)
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._handle_progress)
        self.worker.finished.connect(self._handle_finished)
        self.worker.failed.connect(self._handle_worker_failure)
        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.failed.connect(self.worker_thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.failed.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.worker_thread.finished.connect(self._cleanup_worker)
        self.worker_thread.start()
        self._set_running_state(True)
        self.status_bar.showMessage("任务已开始执行。执行期间请不要手动操作微信。", 5000)

    def stop_batch(self) -> None:
        if self.worker is None:
            return
        self.worker.request_stop()
        self.status_bar.showMessage("已请求停止，当前行完成后将不再继续。", 5000)

    def _handle_progress(self, current: int, total: int, session_name: str) -> None:
        self.status_bar.showMessage(f"正在执行 {current + 1}/{total}: {session_name}")

    def _handle_finished(self, record) -> None:
        saved = self.storage.save_execution(record)
        for row in saved.rows:
            if not row.success:
                self.logger.error(
                    "batch-row-failed task_type=%s row_index=%s session=%s error_code=%s error_message=%s raw_error=%s",
                    saved.task_type.value,
                    row.row_index,
                    row.session_name,
                    row.error_code,
                    row.error_message,
                    row.raw_error,
                )
        self.refresh_history()
        self.status_bar.showMessage(
            f"任务完成: 成功 {saved.success_count} 行，失败 {saved.failure_count} 行",
            7000,
        )
        if saved.failure_count:
            QMessageBox.warning(
                self,
                "任务已完成",
                f"任务执行完成，但有 {saved.failure_count} 行失败。可前往执行历史查看详情并重试失败项。",
            )
        else:
            QMessageBox.information(self, "任务完成", f"任务执行完成，共成功处理 {saved.success_count} 行。")

    def _handle_worker_failure(self, ui_error: UiError) -> None:
        self.logger.error("Batch worker failed: %s", ui_error.diagnostic_text)
        self._show_error_dialog(ui_error, "任务执行失败")

    def _cleanup_worker(self) -> None:
        self.worker = None
        self.worker_thread = None
        self._set_running_state(False)

    def _selected_execution_id(self) -> int | None:
        table = self.history_page.execution_table
        selected = table.selectedItems()
        if not selected:
            return None
        return int(table.item(selected[0].row(), 0).text())

    def show_execution_details(self) -> None:
        execution_id = self._selected_execution_id()
        if execution_id is None:
            return
        execution = self.storage.get_execution(execution_id)
        table = self.history_page.detail_table
        table.setRowCount(0)
        for row_index, result in enumerate(execution.rows):
            table.insertRow(row_index)
            values = [
                str(result.row_index + 1),
                result.session_name,
                "成功" if result.success else "失败",
                result.error_code or "",
                result.error_message or "",
            ]
            for column_index, value in enumerate(values):
                item = QTableWidgetItem(value)
                if not result.success:
                    item.setBackground(QColor("#fef2f2"))
                elif column_index == 2:
                    item.setBackground(QColor("#f0fdf4"))
                table.setItem(row_index, column_index, item)
        table.resizeColumnsToContents()
        self.history_page.diagnostic_text.clear()

    def show_selected_row_diagnostic(self) -> None:
        execution_id = self._selected_execution_id()
        if execution_id is None:
            return
        execution = self.storage.get_execution(execution_id)
        table = self.history_page.detail_table
        selected = table.selectedItems()
        if not selected:
            self.history_page.diagnostic_text.clear()
            return
        row_index = selected[0].row()
        if row_index >= len(execution.rows):
            self.history_page.diagnostic_text.clear()
            return
        raw_error = execution.rows[row_index].raw_error or ""
        self.history_page.diagnostic_text.setPlainText(raw_error)

    def copy_selected_row_diagnostic(self) -> None:
        text = self.history_page.diagnostic_text.toPlainText().strip()
        if not text:
            QMessageBox.information(self, "诊断信息", "当前所选结果没有可复制的诊断信息。")
            return
        QApplication.clipboard().setText(text)
        self.status_bar.showMessage("诊断信息已复制到剪贴板", 4000)

    def retry_selected_execution_failures(self) -> None:
        execution_id = self._selected_execution_id()
        if execution_id is None:
            QMessageBox.information(self, "执行历史", "请先选中一条执行记录。")
            return
        execution = self.storage.get_execution(execution_id)
        failed_rows = failed_rows_from_execution(execution)
        if not failed_rows:
            QMessageBox.information(self, "执行历史", "该执行记录没有失败项。")
            return
        if execution.task_type is TaskType.MESSAGE:
            rows = [MessageBatchRow.from_mapping(row) for row in failed_rows]
            self.message_page.load_rows(rows, source_execution_id=execution.id)
            self.nav.setCurrentRow(1)
        else:
            rows = [FileBatchRow.from_mapping(row) for row in failed_rows]
            self.file_page.load_rows(rows, source_execution_id=execution.id)
            self.nav.setCurrentRow(2)
        self.status_bar.showMessage("已基于失败项重建任务，请确认后重新执行。", 6000)

    def clear_history(self) -> None:
        answer = QMessageBox.question(
            self,
            "清理历史",
            "确定清理全部执行历史吗？模板不会受影响，但历史记录会被永久删除。",
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.storage.clear_history()
        self.refresh_history()
        self.status_bar.showMessage("历史记录已清理", 4000)

    def _set_running_state(self, is_running: bool) -> None:
        self.message_page.set_running_state(is_running)
        self.file_page.set_running_state(is_running)
        for button in [
            self.templates_page.refresh_button,
            self.templates_page.rename_button,
            self.templates_page.duplicate_button,
            self.templates_page.delete_button,
            self.templates_page.load_button,
            self.history_page.refresh_button,
            self.history_page.retry_button,
            self.history_page.clear_button,
            self.settings_page.save_button,
        ]:
            button.setEnabled(not is_running)

    def _show_guidance_dialog(self, title: str, message: str, suggestion: str) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.resize(520, 280)
        layout = QVBoxLayout(dialog)
        title_label = QLabel(title)
        title_label.setStyleSheet("font-size: 18px; font-weight: 700;")
        layout.addWidget(title_label)
        message_label = QLabel(message)
        message_label.setWordWrap(True)
        layout.addWidget(message_label)
        suggestion_label = QLabel(f"建议：{suggestion}")
        suggestion_label.setWordWrap(True)
        suggestion_label.setStyleSheet("background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:10px;")
        layout.addWidget(suggestion_label)
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        button_box.accepted.connect(dialog.accept)
        layout.addWidget(button_box)
        dialog.exec()

    def _show_error_dialog(self, ui_error: UiError, title: str) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.resize(620, 440)
        layout = QVBoxLayout(dialog)
        title_label = QLabel(ui_error.title)
        title_label.setStyleSheet("font-size: 18px; font-weight: 700;")
        layout.addWidget(title_label)
        summary_label = QLabel(ui_error.user_summary)
        summary_label.setWordWrap(True)
        layout.addWidget(summary_label)
        detail_box = QTextEdit()
        detail_box.setReadOnly(True)
        detail_box.setPlainText(ui_error.diagnostic_text)
        layout.addWidget(detail_box)
        actions = QHBoxLayout()
        copy_button = QPushButton("复制诊断信息")
        copy_button.clicked.connect(lambda: QApplication.clipboard().setText(ui_error.diagnostic_text))
        actions.addWidget(copy_button)
        actions.addStretch(1)
        layout.addLayout(actions)
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        button_box.accepted.connect(dialog.accept)
        layout.addWidget(button_box)
        dialog.exec()
