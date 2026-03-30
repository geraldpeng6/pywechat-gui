from __future__ import annotations

import logging

from PySide6.QtCore import QThread
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
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
from ..export_service import ChatExportService
from ..export_worker import ChatExportWorker
from ..import_export import dump_rows
from ..models import AppSettings, ChatBatchExportRequest, ChatBatchExportResult, ChatExportRequest, ChatExportResult, ExportHistoryRecord, FileBatchRow, GroupScanResult, MessageBatchRow, RelayCollectFilesRequest, RelayCollectionResult, RelayCollectTextRequest, RelaySendRequest, RelaySendResult, RelayValidationRequest, RelayValidationResult, ResourceExportRequest, ResourceExportResult, SessionScanRequest, SessionScanResult, TaskTemplate, TaskType, dataclass_from_json, dataclass_to_json
from ..presentation import execution_metrics, export_history_can_rerun, export_history_can_retry_failed, export_history_failed_sessions, filter_executions, filter_templates, format_export_history_detail, rebuild_export_request, serialize_export_detail, summarize_failures, template_metrics, template_type_label
from ..relay_service import RelayService
from ..relay_worker import RelayWorker
from ..resource_export_service import ResourceExportService, export_kind_label
from ..resource_export_worker import ResourceExportWorker
from ..session_tools_service import SessionToolsService
from ..session_tools_worker import SessionToolsWorker
from ..settings_manager import SettingsManager
from ..storage import AppStorage
from ..system_ops import open_path
from ..worker import BatchWorker
from .widgets import BatchPage, DashboardPage, ExportHistoryPage, ExportPage, HistoryPage, RelayWorkbenchPage, ResourceToolsPage, SessionToolsPage, SettingsPage, TemplatesPage


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
        self.export_service = ChatExportService(adapter)
        self.resource_export_service = ResourceExportService(adapter)
        self.session_tools_service = SessionToolsService(adapter)
        self.relay_service = RelayService(adapter)
        self.worker_thread: QThread | None = None
        self.worker = None
        self._template_cache: list[TaskTemplate] = []
        self._execution_cache = []
        self._export_history_cache: list[ExportHistoryRecord] = []
        self._last_deleted_template: TaskTemplate | None = None

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
        self.export_page = ExportPage()
        self.resource_page = ResourceToolsPage()
        self.session_tools_page = SessionToolsPage()
        self.relay_page = RelayWorkbenchPage()
        self.templates_page = TemplatesPage()
        self.history_page = HistoryPage()
        self.export_history_page = ExportHistoryPage()
        self.settings_page = SettingsPage()

        self._add_page("首页", self.dashboard_page)
        self._add_page("批量消息", self.message_page)
        self._add_page("批量文件", self.file_page)
        self._add_page("会话导出", self.export_page)
        self._add_page("资源导出", self.resource_page)
        self._add_page("会话与群工具", self.session_tools_page)
        self._add_page("转发工作台", self.relay_page)
        self._add_page("模板中心", self.templates_page)
        self._add_page("执行历史", self.history_page)
        self._add_page("导出历史", self.export_history_page)
        self._add_page("设置", self.settings_page)
        self.nav.setCurrentRow(0)
        self.nav.currentRowChanged.connect(self.stack.setCurrentIndex)

        self._bind_events()
        self._load_settings_to_form()
        self.refresh_environment()
        self.refresh_templates()
        self.refresh_history()
        self.refresh_export_history()
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
        self.dashboard_page.open_templates_requested.connect(lambda: self.nav.setCurrentRow(7))
        self.message_page.run_requested.connect(lambda rows, src: self.start_batch(TaskType.MESSAGE, rows, src))
        self.file_page.run_requested.connect(lambda rows, src: self.start_batch(TaskType.FILE, rows, src))
        self.message_page.stop_requested.connect(self.stop_batch)
        self.file_page.stop_requested.connect(self.stop_batch)
        self.export_page.export_requested.connect(self.start_export)
        self.export_page.batch_export_requested.connect(self.start_batch_export)
        self.export_page.stop_requested.connect(self.stop_batch)
        self.resource_page.export_requested.connect(self.start_resource_export)
        self.resource_page.stop_requested.connect(self.stop_batch)
        self.export_page.open_folder_button.clicked.connect(lambda: self._open_export_page_folder(self.export_page))
        self.resource_page.open_folder_button.clicked.connect(lambda: self._open_export_page_folder(self.resource_page))
        self.session_tools_page.scan_sessions_requested.connect(self.start_session_scan)
        self.session_tools_page.scan_groups_requested.connect(self.start_group_scan)
        self.session_tools_page.use_session_names_requested.connect(self.load_session_names_into_export_page)
        self.relay_page.collect_texts_requested.connect(self.start_relay_collect_texts)
        self.relay_page.collect_files_requested.connect(self.start_relay_collect_files)
        self.relay_page.validate_routes_requested.connect(self.start_relay_validate_routes)
        self.relay_page.test_send_requested.connect(self.start_relay_test_send)
        self.relay_page.send_requested.connect(self.start_relay_send)
        self.message_page.save_template_requested.connect(self.save_template)
        self.file_page.save_template_requested.connect(self.save_template)
        self.message_page.open_templates_requested.connect(self.open_templates_for)
        self.file_page.open_templates_requested.connect(self.open_templates_for)

        self.templates_page.refresh_button.clicked.connect(self.refresh_templates)
        self.templates_page.search_input.textChanged.connect(self.apply_template_filter)
        self.templates_page.create_message_requested.connect(lambda: self.nav.setCurrentRow(1))
        self.templates_page.create_file_requested.connect(lambda: self.nav.setCurrentRow(2))
        self.templates_page.load_button.clicked.connect(self.load_selected_template)
        self.templates_page.delete_button.clicked.connect(self.delete_selected_template)
        self.templates_page.restore_button.clicked.connect(self.restore_deleted_template)
        self.templates_page.rename_button.clicked.connect(self.rename_selected_template)
        self.templates_page.duplicate_button.clicked.connect(self.duplicate_selected_template)
        self.templates_page.table.itemDoubleClicked.connect(lambda _item: self.load_selected_template())

        self.history_page.refresh_button.clicked.connect(self.refresh_history)
        self.history_page.search_input.textChanged.connect(self.apply_history_filter)
        self.history_page.failed_only_checkbox.toggled.connect(self.apply_history_filter)
        self.history_page.open_message_requested.connect(lambda: self.nav.setCurrentRow(1))
        self.history_page.open_file_requested.connect(lambda: self.nav.setCurrentRow(2))
        self.history_page.clear_button.clicked.connect(self.clear_history)
        self.history_page.export_failed_button.clicked.connect(self.export_selected_execution_failures)
        self.history_page.execution_table.itemSelectionChanged.connect(self.show_execution_details)
        self.history_page.execution_table.itemSelectionChanged.connect(self._update_history_action_state)
        self.history_page.detail_table.itemSelectionChanged.connect(self.show_selected_row_diagnostic)
        self.history_page.copy_diag_button.clicked.connect(self.copy_selected_row_diagnostic)
        self.history_page.retry_button.clicked.connect(self.retry_selected_execution_failures)

        self.export_history_page.search_input.textChanged.connect(self.apply_export_history_filter)
        self.export_history_page.table.itemSelectionChanged.connect(self.show_export_history_detail)
        self.export_history_page.open_folder_button.clicked.connect(self.open_selected_export_folder)
        self.export_history_page.open_summary_button.clicked.connect(self.open_selected_export_summary)
        self.export_history_page.rerun_button.clicked.connect(self.rerun_selected_export_record)
        self.export_history_page.retry_failed_button.clicked.connect(self.retry_failed_export_sessions)
        self.export_history_page.clear_button.clicked.connect(self.clear_export_history)

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
        filtered = filter_templates(templates, query)
        for row_index, template in enumerate(filtered):
            table.insertRow(row_index)
            values = [
                str(template.id),
                template.name,
                template_type_label(template.task_type),
                template.updated_at or "",
            ]
            for column_index, value in enumerate(values):
                item = QTableWidgetItem(value)
                if column_index == 2:
                    if template.task_type is TaskType.MESSAGE:
                        item.setBackground(QColor("#dbeafe"))
                    else:
                        item.setBackground(QColor("#fef3c7"))
                table.setItem(row_index, column_index, item)
        table.resizeColumnsToContents()
        metrics = template_metrics(templates)
        self.templates_page.total_templates_label.value_label.setText(str(metrics["total"]))  # type: ignore[attr-defined]
        self.templates_page.message_templates_label.value_label.setText(str(metrics["message"]))  # type: ignore[attr-defined]
        self.templates_page.file_templates_label.value_label.setText(str(metrics["file"]))  # type: ignore[attr-defined]
        if filtered:
            self.templates_page.summary_label.setText(
                f"当前显示 {len(filtered)} / {len(templates)} 个模板。双击或选中后可加载到工作台。"
            )
        else:
            if templates:
                self.templates_page.summary_label.setText("没有匹配到模板，请换个关键词再试。")
            else:
                self.templates_page.summary_label.setText("还没有模板。你可以先去批量页面整理一份任务，再点“保存模板”。")
        has_templates = bool(templates)
        self.templates_page.create_message_button.setVisible(not has_templates)
        self.templates_page.create_file_button.setVisible(not has_templates)
        self.templates_page.restore_button.setEnabled(self._last_deleted_template is not None)

    def refresh_history(self) -> None:
        self._execution_cache = self.storage.list_executions()
        self.apply_history_filter()

    def refresh_export_history(self) -> None:
        self._export_history_cache = self.storage.list_export_records()
        self.apply_export_history_filter()

    def apply_history_filter(self) -> None:
        query = self.history_page.search_input.text().strip().lower()
        executions = self._execution_cache
        table = self.history_page.execution_table
        table.setRowCount(0)
        filtered = filter_executions(executions, query, self.history_page.failed_only_checkbox.isChecked())
        for row_index, execution in enumerate(filtered):
            table.insertRow(row_index)
            values = [
                str(execution.id),
                template_type_label(execution.task_type),
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
        self.history_page.failure_summary_label.setText("选中一条执行记录后，这里会展示失败原因摘要。")
        metrics = execution_metrics(executions)
        self.history_page.total_runs_label.value_label.setText(str(metrics["total"]))  # type: ignore[attr-defined]
        self.history_page.success_runs_label.value_label.setText(str(metrics["success"]))  # type: ignore[attr-defined]
        self.history_page.failed_runs_label.value_label.setText(str(metrics["failed"]))  # type: ignore[attr-defined]
        if filtered:
            self.history_page.summary_label.setText(
                f"当前显示 {len(filtered)} / {len(executions)} 条执行记录。先选中一条，再看下方逐行结果。"
            )
        else:
            if executions:
                self.history_page.summary_label.setText("没有匹配到执行记录，请换个关键词再试。")
            else:
                self.history_page.summary_label.setText("还没有执行历史。第一次成功或失败执行后，这里会显示完整记录。")
        has_history = bool(executions)
        self.history_page.open_message_button.setVisible(not has_history)
        self.history_page.open_file_button.setVisible(not has_history)
        self._update_history_action_state()

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
        self.nav.setCurrentRow(7)
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
        self._last_deleted_template = self.storage.get_template(template_id)
        self.storage.delete_template(template_id)
        self.refresh_templates()
        self.status_bar.showMessage("模板已删除。如有需要，可点“恢复刚删除”。", 5000)

    def restore_deleted_template(self) -> None:
        if self._last_deleted_template is None:
            QMessageBox.information(self, "恢复模板", "当前没有可恢复的模板。")
            return
        template = self._last_deleted_template
        template.id = None
        restored = self.storage.save_template(template)
        self._last_deleted_template = None
        self.refresh_templates()
        self.status_bar.showMessage(f"已恢复模板：{restored.name}", 5000)

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
        enabled_rows = [row for row in rows if row.enabled]
        if len(enabled_rows) >= 20 and source_execution_id is None:
            answer = QMessageBox.question(
                self,
                "确认执行大批量任务",
                f"当前准备执行 {len(enabled_rows)} 行任务。\n\n"
                "建议先确认会话名称和消息内容无误，再继续执行。\n"
                "是否现在开始？",
            )
            if answer != QMessageBox.StandardButton.Yes:
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

    def start_export(self, request: ChatExportRequest) -> None:
        if self.worker_thread is not None:
            QMessageBox.warning(self, "任务进行中", "当前已有任务在执行，请等待完成或先停止。")
            return
        environment = self.adapter.inspect_environment()
        if environment.login_status != "已登录":
            self.nav.setCurrentRow(0)
            self._show_guidance_dialog(
                title="暂时不能执行导出",
                message=environment.status_message,
                suggestion="\n".join(environment.advice) if environment.advice else "请先解决首页提示的问题，再回来导出。",
            )
            return
        self.worker_thread = QThread(self)
        self.worker = ChatExportWorker(self.export_service, request, self.settings.runtime_options())
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._handle_export_progress)
        self.worker.finished.connect(self._handle_export_finished)
        self.worker.failed.connect(self._handle_worker_failure)
        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.failed.connect(self.worker_thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.failed.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.worker_thread.finished.connect(self._cleanup_worker)
        self.worker_thread.start()
        self._set_running_state(True)
        self.export_page.result_text.clear()
        self.export_page.open_folder_button.setEnabled(False)
        self.export_page.summary_label.setText("导出任务已开始，请勿手动操作微信。")
        self.status_bar.showMessage("导出任务已开始，请等待完成。", 5000)

    def start_batch_export(self, request: ChatBatchExportRequest) -> None:
        if self.worker_thread is not None:
            QMessageBox.warning(self, "任务进行中", "当前已有任务在执行，请等待完成或先停止。")
            return
        environment = self.adapter.inspect_environment()
        if environment.login_status != "已登录":
            self.nav.setCurrentRow(0)
            self._show_guidance_dialog(
                title="暂时不能执行导出",
                message=environment.status_message,
                suggestion="\n".join(environment.advice) if environment.advice else "请先解决首页提示的问题，再回来导出。",
            )
            return
        self.worker_thread = QThread(self)
        self.worker = ChatExportWorker(self.export_service, request, self.settings.runtime_options())
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._handle_export_progress)
        self.worker.finished.connect(self._handle_export_finished)
        self.worker.failed.connect(self._handle_worker_failure)
        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.failed.connect(self.worker_thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.failed.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.worker_thread.finished.connect(self._cleanup_worker)
        self.worker_thread.start()
        self._set_running_state(True)
        self.export_page.result_text.clear()
        self.export_page.open_folder_button.setEnabled(False)
        self.export_page.summary_label.setText("批量会话导出已开始，请等待完成。")
        self.status_bar.showMessage("批量会话导出已开始。", 5000)

    def start_resource_export(self, request: ResourceExportRequest) -> None:
        if self.worker_thread is not None:
            QMessageBox.warning(self, "任务进行中", "当前已有任务在执行，请等待完成或先停止。")
            return
        environment = self.adapter.inspect_environment()
        if environment.login_status != "已登录":
            self.nav.setCurrentRow(0)
            self._show_guidance_dialog(
                title="暂时不能执行导出",
                message=environment.status_message,
                suggestion="\n".join(environment.advice) if environment.advice else "请先解决首页提示的问题，再回来导出。",
            )
            return
        self.worker_thread = QThread(self)
        self.worker = ResourceExportWorker(self.resource_export_service, request, self.settings.runtime_options())
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._handle_resource_export_progress)
        self.worker.finished.connect(self._handle_resource_export_finished)
        self.worker.failed.connect(self._handle_worker_failure)
        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.failed.connect(self.worker_thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.failed.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.worker_thread.finished.connect(self._cleanup_worker)
        self.worker_thread.start()
        self._set_running_state(True)
        self.resource_page.result_text.clear()
        self.resource_page.open_folder_button.setEnabled(False)
        self.resource_page.summary_label.setText("资源导出已开始，请等待完成。")
        self.status_bar.showMessage("资源导出已开始，请等待完成。", 5000)

    def start_session_scan(self, request: SessionScanRequest) -> None:
        self._start_session_tools_worker(
            action="scan_sessions",
            request=request,
            start_message="会话采集已开始，请等待完成。",
        )

    def start_group_scan(self) -> None:
        self._start_session_tools_worker(
            action="scan_groups",
            request=None,
            start_message="群聊采集已开始，请等待完成。",
        )

    def start_relay_collect_texts(self, request: RelayCollectTextRequest) -> None:
        self._start_relay_worker("collect_texts", request, "正在采集上游文本，请等待完成。")

    def start_relay_collect_files(self, request: RelayCollectFilesRequest) -> None:
        self._start_relay_worker("collect_files", request, "正在采集上游聊天文件，请等待完成。")

    def start_relay_validate_routes(self, request: RelayValidationRequest) -> None:
        self._start_relay_worker("validate_routes", request, "正在验证下游会话，请等待完成。")

    def start_relay_test_send(self, request: RelaySendRequest) -> None:
        self._start_relay_worker("send_package", request, "正在发送测试包到文件传输助手，请勿手动操作微信。")

    def start_relay_send(self, request: RelaySendRequest) -> None:
        answer = QMessageBox.question(
            self,
            "确认正式批量发送",
            "正式发送会按转发包当前顺序，向当前上游匹配到的全部下游逐个发送。\n\n是否现在开始？",
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self._start_relay_worker("send_package", request, "正在正式批量发送，请勿手动操作微信。")

    def _start_session_tools_worker(
        self,
        action: str,
        request: SessionScanRequest | None,
        start_message: str,
    ) -> None:
        if self.worker_thread is not None:
            QMessageBox.warning(self, "任务进行中", "当前已有任务在执行，请等待完成或先停止。")
            return
        environment = self.adapter.inspect_environment()
        if environment.login_status != "已登录":
            self.nav.setCurrentRow(0)
            self._show_guidance_dialog(
                title="暂时不能执行采集",
                message=environment.status_message,
                suggestion="\n".join(environment.advice) if environment.advice else "请先解决首页提示的问题，再回来执行。",
            )
            return
        self.worker_thread = QThread(self)
        self.worker = SessionToolsWorker(
            service=self.session_tools_service,
            action=action,
            runtime_options=self.settings.runtime_options(),
            request=request,
        )
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._handle_session_tools_progress)
        self.worker.finished.connect(self._handle_session_tools_finished)
        self.worker.failed.connect(self._handle_worker_failure)
        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.failed.connect(self.worker_thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.failed.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.worker_thread.finished.connect(self._cleanup_worker)
        self.worker_thread.start()
        self._set_running_state(True)
        self.status_bar.showMessage(start_message, 5000)

    def _start_relay_worker(
        self,
        action: str,
        request: RelayCollectTextRequest | RelayCollectFilesRequest | RelayValidationRequest | RelaySendRequest,
        start_message: str,
    ) -> None:
        if self.worker_thread is not None:
            QMessageBox.warning(self, "任务进行中", "当前已有任务在执行，请等待完成或先停止。")
            return
        environment = self.adapter.inspect_environment()
        if environment.login_status != "已登录":
            self.nav.setCurrentRow(0)
            self._show_guidance_dialog(
                title="暂时不能执行转发工作台任务",
                message=environment.status_message,
                suggestion="\n".join(environment.advice) if environment.advice else "请先解决首页提示的问题，再回来执行。",
            )
            return
        self.worker_thread = QThread(self)
        self.worker = RelayWorker(
            service=self.relay_service,
            action=action,
            runtime_options=self.settings.runtime_options(),
            request=request,
        )
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._handle_relay_progress)
        self.worker.finished.connect(self._handle_relay_finished)
        self.worker.failed.connect(self._handle_worker_failure)
        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.failed.connect(self.worker_thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.failed.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.worker_thread.finished.connect(self._cleanup_worker)
        self.worker_thread.start()
        self._set_running_state(True)
        self.status_bar.showMessage(start_message, 5000)

    def load_session_names_into_export_page(self, session_names: list[str]) -> None:
        if not session_names:
            QMessageBox.information(self, "会话与群工具", "当前没有可回填的会话名称。")
            return
        self.export_page.session_names_input.setPlainText("\n".join(session_names))
        self.export_page.summary_label.setText(f"已从会话与群工具回填 {len(session_names)} 个会话名称。")
        self.nav.setCurrentRow(3)

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

    def _handle_export_progress(self, message: str) -> None:
        self.export_page.summary_label.setText(message)
        self.status_bar.showMessage(message, 3000)

    def _handle_export_finished(self, result: ChatExportResult | ChatBatchExportResult) -> None:
        if isinstance(result, ChatBatchExportResult):
            request = self.worker.request if isinstance(self.worker, ChatExportWorker) else None
            lines = [
                f"批量导出目录：{result.export_root}",
                f"会话总数：{result.total_sessions}",
                f"成功数量：{result.success_count}",
                f"失败数量：{result.failure_count}",
            ]
            if result.summary_txt:
                lines.append(f"摘要文件：{result.summary_txt}")
            if result.failed_sessions:
                lines.append("")
                lines.append("失败会话：")
                lines.extend(f"- {item['session_name']}: {item['error']}" for item in result.failed_sessions[:10])
            self.export_page.result_text.setPlainText("\n".join(lines))
            self.export_page.summary_label.setText("批量导出完成。你可以直接打开导出目录查看结果。")
            self.export_page.open_folder_button.setEnabled(True)
            self.export_page.open_folder_button.setProperty("export_path", result.export_root)
            self._save_export_record(
                export_kind="chat_batch",
                title=f"批量会话导出 ({result.total_sessions} 个)",
                export_folder=result.export_root,
                exported_count=result.success_count,
                summary_path=result.summary_txt,
                detail_json=serialize_export_detail(
                    "chat_batch",
                    request if isinstance(request, ChatBatchExportRequest) else ChatBatchExportRequest(session_names=[], target_folder=result.export_root),
                    {
                        "success_count": result.success_count,
                        "failure_count": result.failure_count,
                        "failed_sessions": result.failed_sessions,
                    },
                ),
            )
            self.status_bar.showMessage("批量会话导出完成。", 5000)
            QMessageBox.information(self, "导出完成", f"已完成批量会话导出。\n目录：{result.export_root}")
            return

        request = self.worker.request if isinstance(self.worker, ChatExportWorker) else None
        lines = [
            f"会话名称：{result.session_name}",
            f"导出目录：{result.export_folder}",
            f"消息数量：{result.message_count}",
            f"文件数量：{result.file_count}",
        ]
        if result.messages_csv:
            lines.append(f"消息 CSV：{result.messages_csv}")
        if result.messages_json:
            lines.append(f"消息 JSON：{result.messages_json}")
        if result.files_folder:
            lines.append(f"文件目录：{result.files_folder}")
        if result.warnings:
            lines.append("")
            lines.append("注意事项：")
            lines.extend(f"- {warning}" for warning in result.warnings)
        self.export_page.result_text.setPlainText("\n".join(lines))
        self.export_page.summary_label.setText("导出完成。你可以直接打开导出目录查看结果。")
        self.export_page.open_folder_button.setEnabled(True)
        self.export_page.open_folder_button.setProperty("export_path", result.export_folder)
        self._save_export_record(
            export_kind="chat",
            title=result.session_name,
            export_folder=result.export_folder,
            exported_count=result.message_count + result.file_count,
            summary_path=result.summary_txt,
            detail_json=serialize_export_detail(
                "chat",
                request if isinstance(request, ChatExportRequest) else ChatExportRequest(session_name=result.session_name, target_folder=result.export_folder),
                {
                    "session_name": result.session_name,
                    "message_count": result.message_count,
                    "file_count": result.file_count,
                    "messages_csv": result.messages_csv,
                    "messages_json": result.messages_json,
                    "files_folder": result.files_folder,
                    "warnings": result.warnings,
                },
            ),
        )
        self.status_bar.showMessage("会话导出完成。", 5000)
        QMessageBox.information(self, "导出完成", f"已完成会话导出。\n目录：{result.export_folder}")

    def _handle_resource_export_progress(self, message: str) -> None:
        self.resource_page.summary_label.setText(message)
        self.status_bar.showMessage(message, 3000)

    def _handle_resource_export_finished(self, result: ResourceExportResult) -> None:
        request = self.worker.request if isinstance(self.worker, ResourceExportWorker) else None
        lines = [
            f"导出类型：{export_kind_label(result.export_kind)}",
            f"导出目录：{result.target_folder}",
            f"导出数量：{result.exported_count}",
        ]
        if result.summary_txt:
            lines.append(f"摘要文件：{result.summary_txt}")
        if result.exported_paths:
            lines.append("")
            lines.append("导出结果预览：")
            lines.extend(f"- {path}" for path in result.exported_paths[:10])
            if len(result.exported_paths) > 10:
                lines.append(f"- ... 共 {len(result.exported_paths)} 项")
        self.resource_page.result_text.setPlainText("\n".join(lines))
        self.resource_page.summary_label.setText("资源导出完成。你可以直接打开导出目录查看结果。")
        self.resource_page.open_folder_button.setEnabled(True)
        self.resource_page.open_folder_button.setProperty("export_path", result.target_folder)
        self._save_export_record(
            export_kind=result.export_kind.value,
            title=export_kind_label(result.export_kind),
            export_folder=result.target_folder,
            exported_count=result.exported_count,
            summary_path=result.summary_txt,
            detail_json=serialize_export_detail(
                result.export_kind.value,
                request if isinstance(request, ResourceExportRequest) else ResourceExportRequest(export_kind=result.export_kind, target_folder=result.target_folder),
                {
                    "export_kind": result.export_kind.value,
                    "exported_count": result.exported_count,
                    "exported_paths": result.exported_paths[:20],
                },
            ),
        )
        self.status_bar.showMessage("资源导出完成。", 5000)
        QMessageBox.information(self, "导出完成", f"已完成资源导出。\n目录：{result.target_folder}")

    def _handle_session_tools_progress(self, message: str) -> None:
        self.status_bar.showMessage(message, 3000)
        self.session_tools_page.session_summary_label.setText(message)

    def _handle_session_tools_finished(self, action: str, result: SessionScanResult | GroupScanResult) -> None:
        if action == "scan_sessions" and isinstance(result, SessionScanResult):
            self.session_tools_page.set_session_rows(result.rows)
            self.status_bar.showMessage(f"已采集 {len(result.rows)} 个会话。", 5000)
            return
        if action == "scan_groups" and isinstance(result, GroupScanResult):
            self.session_tools_page.set_group_rows(result.rows)
            self.status_bar.showMessage(f"已采集 {len(result.rows)} 个群聊。", 5000)
            return

    def _handle_relay_progress(self, message: str) -> None:
        self.status_bar.showMessage(message, 3000)
        self.relay_page.result_text.setPlainText(message)

    def _handle_relay_finished(
        self,
        action: str,
        result: RelayCollectionResult | RelayValidationResult | RelaySendResult,
    ) -> None:
        if action == "collect_texts" and isinstance(result, RelayCollectionResult):
            self.relay_page.append_package_rows(result.rows)
            self.relay_page.refresh_package_summary(
                result.warning or f"已采集 {len(result.rows)} 条文本内容，请勾选真正需要转发的内容。"
            )
            self.status_bar.showMessage("上游文本采集完成。", 5000)
            return
        if action == "collect_files" and isinstance(result, RelayCollectionResult):
            self.relay_page.append_package_rows(result.rows)
            self.relay_page.refresh_package_summary(
                result.warning or f"已采集 {len(result.rows)} 个文件/图片，请确认是否都是最新版本。"
            )
            self.status_bar.showMessage("上游文件采集完成。", 5000)
            return
        if action == "validate_routes" and isinstance(result, RelayValidationResult):
            summary = (
                f"已验证 {result.checked_count} 条下游路由："
                f"找到 {result.found_count} 条，未找到 {result.missing_count} 条。"
            )
            self.relay_page.apply_validation_result(result.route_rows, summary)
            self.status_bar.showMessage("下游验证完成。", 5000)
            return
        if action == "send_package" and isinstance(result, RelaySendResult):
            lines = [
                "测试发送已完成。" if result.test_only else "正式批量发送已完成。",
                f"来源会话：{result.source_session or '未填写'}",
                f"转发内容数：{result.item_count}",
                f"目标会话数：{result.target_count}",
                f"成功目标：{result.success_count}",
                f"失败目标：{result.failure_count}",
            ]
            if result.results:
                lines.append("")
                lines.append("结果明细：")
                for row in result.results[:20]:
                    status = "成功" if row.success else "失败"
                    detail = f"{row.target_session}：{status}，已发 {row.sent_count} 条"
                    if row.error_message:
                        detail += f"；{row.error_message}"
                    lines.append(f"- {detail}")
            self.relay_page.apply_send_result("\n".join(lines))
            self.status_bar.showMessage("测试发送完成。" if result.test_only else "正式发送完成。", 5000)
            return

    def _handle_worker_failure(self, ui_error: UiError) -> None:
        self.logger.error("Batch worker failed: %s", ui_error.diagnostic_text)
        if isinstance(self.worker, ChatExportWorker):
            self.export_page.summary_label.setText("任务执行失败，请查看错误详情。")
        elif isinstance(self.worker, ResourceExportWorker):
            self.resource_page.summary_label.setText("任务执行失败，请查看错误详情。")
        elif isinstance(self.worker, SessionToolsWorker):
            self.session_tools_page.session_summary_label.setText("采集失败，请查看错误详情。")
        elif isinstance(self.worker, RelayWorker):
            self.relay_page.result_text.setPlainText("任务执行失败，请查看错误详情。")
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
        self.history_page.failure_summary_label.setText(summarize_failures(execution.rows))

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
        self.export_page.set_running_state(is_running)
        self.resource_page.set_running_state(is_running)
        self.session_tools_page.set_running_state(is_running)
        self.relay_page.set_running_state(is_running)
        for button in [
            self.templates_page.refresh_button,
            self.templates_page.rename_button,
            self.templates_page.duplicate_button,
            self.templates_page.delete_button,
            self.templates_page.load_button,
            self.templates_page.restore_button,
            self.history_page.refresh_button,
            self.history_page.retry_button,
            self.history_page.export_failed_button,
            self.history_page.clear_button,
            self.export_history_page.open_folder_button,
            self.export_history_page.open_summary_button,
            self.export_history_page.rerun_button,
            self.export_history_page.retry_failed_button,
            self.export_history_page.clear_button,
            self.settings_page.save_button,
        ]:
            button.setEnabled(not is_running)
        self.history_page.failed_only_checkbox.setEnabled(not is_running)
        self.export_history_page.search_input.setEnabled(not is_running)
        if not is_running:
            self._update_history_action_state()
            self._update_export_history_action_state()

    def apply_export_history_filter(self) -> None:
        query = self.export_history_page.search_input.text().strip().lower()
        records = self._export_history_cache
        table = self.export_history_page.table
        table.setRowCount(0)
        filtered = []
        for record in records:
            haystack = " ".join([record.export_kind, record.title, record.export_folder, record.created_at or ""]).lower()
            if query and query not in haystack:
                continue
            filtered.append(record)
        for row_index, record in enumerate(filtered):
            table.insertRow(row_index)
            values = [str(record.id), record.export_kind, record.title, str(record.exported_count), record.created_at or ""]
            for column_index, value in enumerate(values):
                table.setItem(row_index, column_index, QTableWidgetItem(value))
        table.resizeColumnsToContents()
        self.export_history_page.summary_label.setText(
            f"当前显示 {len(filtered)} / {len(records)} 条导出记录。选中后可打开目录或摘要。"
            if records
            else "还没有导出历史。完成一次会话导出或资源导出后，这里会自动记录。"
        )
        self.export_history_page.open_folder_button.setEnabled(False)
        self.export_history_page.open_summary_button.setEnabled(False)
        self.export_history_page.rerun_button.setEnabled(False)
        self.export_history_page.retry_failed_button.setEnabled(False)
        self.export_history_page.detail_text.clear()

    def _selected_export_record(self) -> ExportHistoryRecord | None:
        selected = self.export_history_page.table.selectedItems()
        if not selected:
            return None
        record_id = int(self.export_history_page.table.item(selected[0].row(), 0).text())
        return next((item for item in self._export_history_cache if item.id == record_id), None)

    def show_export_history_detail(self) -> None:
        record = self._selected_export_record()
        if record is None:
            self.export_history_page.detail_text.clear()
            self._update_export_history_action_state()
            return
        self.export_history_page.detail_text.setPlainText(format_export_history_detail(record))
        self._update_export_history_action_state()

    def _update_export_history_action_state(self) -> None:
        record = self._selected_export_record()
        if record is None:
            self.export_history_page.open_folder_button.setEnabled(False)
            self.export_history_page.open_summary_button.setEnabled(False)
            self.export_history_page.rerun_button.setEnabled(False)
            self.export_history_page.retry_failed_button.setEnabled(False)
            return
        self.export_history_page.open_folder_button.setEnabled(True)
        self.export_history_page.open_folder_button.setProperty("export_path", record.export_folder)
        self.export_history_page.open_summary_button.setEnabled(bool(record.summary_path))
        self.export_history_page.open_summary_button.setProperty("export_path", record.summary_path or "")
        self.export_history_page.rerun_button.setEnabled(export_history_can_rerun(record))
        self.export_history_page.retry_failed_button.setEnabled(export_history_can_retry_failed(record))

    def rerun_selected_export_record(self) -> None:
        record = self._selected_export_record()
        if record is None:
            QMessageBox.information(self, "导出历史", "请先选中一条导出记录。")
            return
        payload = rebuild_export_request(record)
        if payload is None:
            QMessageBox.information(self, "导出历史", "这条记录缺少原始执行参数，暂时不能直接重新执行。")
            return
        kind, request = payload
        if kind == "chat":
            self.export_page.apply_single_request(request)
            self.nav.setCurrentRow(3)
            self.start_export(request)
            return
        if kind == "chat_batch":
            self.export_page.apply_batch_request(request)
            self.nav.setCurrentRow(3)
            self.start_batch_export(request)
            return
        self.resource_page.apply_request(request)
        self.nav.setCurrentRow(4)
        self.start_resource_export(request)

    def retry_failed_export_sessions(self) -> None:
        record = self._selected_export_record()
        if record is None:
            QMessageBox.information(self, "导出历史", "请先选中一条导出记录。")
            return
        failed_names = export_history_failed_sessions(record)
        if not failed_names:
            QMessageBox.information(self, "导出历史", "这条记录没有失败会话可重试。")
            return
        payload = rebuild_export_request(record, failed_only=True)
        if payload is None:
            QMessageBox.information(self, "导出历史", "当前无法基于失败会话重建导出任务。")
            return
        _, request = payload
        self.export_page.apply_batch_request(request)
        self.nav.setCurrentRow(3)
        self.start_batch_export(request)

    def open_selected_export_folder(self) -> None:
        path = self.export_history_page.open_folder_button.property("export_path")
        if not path:
            QMessageBox.information(self, "导出历史", "当前没有可打开的导出目录。")
            return
        try:
            open_path(path)
        except Exception as exc:
            QMessageBox.critical(self, "打开导出目录", str(exc))

    def open_selected_export_summary(self) -> None:
        path = self.export_history_page.open_summary_button.property("export_path")
        if not path:
            QMessageBox.information(self, "导出历史", "当前没有可打开的摘要文件。")
            return
        try:
            open_path(path)
        except Exception as exc:
            QMessageBox.critical(self, "打开摘要文件", str(exc))

    def clear_export_history(self) -> None:
        answer = QMessageBox.question(self, "清理导出历史", "确定清理全部导出历史吗？不会删除磁盘上的真实导出文件。")
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.storage.clear_export_history()
        self.refresh_export_history()
        self.status_bar.showMessage("导出历史已清理。", 4000)

    def _save_export_record(
        self,
        export_kind: str,
        title: str,
        export_folder: str,
        exported_count: int,
        summary_path: str | None,
        detail_json: str,
    ) -> None:
        self.storage.save_export_record(
            ExportHistoryRecord(
                export_kind=export_kind,
                title=title,
                export_folder=export_folder,
                exported_count=exported_count,
                summary_path=summary_path,
                detail_json=detail_json,
            )
        )
        self.refresh_export_history()

    def _open_export_page_folder(self, page) -> None:
        path = page.open_folder_button.property("export_path")
        if not path:
            QMessageBox.information(self, "打开导出目录", "当前没有可打开的导出目录。")
            return
        try:
            open_path(path)
        except Exception as exc:
            QMessageBox.critical(self, "打开导出目录", str(exc))

    def _update_history_action_state(self) -> None:
        selected_execution = self._selected_execution_id()
        if selected_execution is None:
            self.history_page.export_failed_button.setEnabled(False)
            self.history_page.retry_button.setEnabled(False)
            return
        execution = self.storage.get_execution(selected_execution)
        has_failures = execution.failure_count > 0
        self.history_page.export_failed_button.setEnabled(has_failures)
        self.history_page.retry_button.setEnabled(has_failures)

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

    def export_selected_execution_failures(self) -> None:
        execution_id = self._selected_execution_id()
        if execution_id is None:
            QMessageBox.information(self, "导出失败项", "请先选中一条执行记录。")
            return
        execution = self.storage.get_execution(execution_id)
        failed_rows = failed_rows_from_execution(execution)
        if not failed_rows:
            QMessageBox.information(self, "导出失败项", "该执行记录没有失败项可导出。")
            return
        path, _ = QFileDialog.getSaveFileName(
            self,
            "导出失败项",
            f"execution-{execution.id}-failed-rows.xlsx",
            "Excel (*.xlsx);;CSV (*.csv)",
        )
        if not path:
            return
        try:
            dump_rows(execution.task_type, failed_rows, path)
            self.status_bar.showMessage("失败项已导出，可用于二次处理。", 5000)
        except Exception as exc:
            QMessageBox.critical(self, "导出失败项", str(exc))
