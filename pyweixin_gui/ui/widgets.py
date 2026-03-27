from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Any

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor, QGuiApplication
from PySide6.QtWidgets import (
    QCheckBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMenu,
    QMessageBox,
    QPushButton,
    QHeaderView,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from ..import_export import dump_rows, load_rows
from ..models import FileBatchRow, MessageBatchRow, TaskType, clone_row


@dataclass
class ColumnSpec:
    key: str
    title: str
    kind: str = "text"


MESSAGE_COLUMNS = [
    ColumnSpec("enabled", "启用", "bool"),
    ColumnSpec("session_name", "会话名称"),
    ColumnSpec("message", "消息内容"),
    ColumnSpec("at_members", "@成员"),
    ColumnSpec("at_all", "@所有人", "bool"),
    ColumnSpec("clear_before_send", "清空输入框", "bool"),
    ColumnSpec("send_delay_sec", "发送间隔(秒)", "float"),
    ColumnSpec("remark", "备注"),
]

FILE_COLUMNS = [
    ColumnSpec("enabled", "启用", "bool"),
    ColumnSpec("session_name", "会话名称"),
    ColumnSpec("file_paths", "文件路径(|分隔)"),
    ColumnSpec("with_message", "附带消息", "bool"),
    ColumnSpec("message", "消息内容"),
    ColumnSpec("message_first", "消息先发", "bool"),
    ColumnSpec("remark", "备注"),
]


class CardFrame(QFrame):
    def __init__(self, title: str | None = None, parent: QWidget | None = None, hero: bool = False):
        super().__init__(parent)
        self.setObjectName("HeroCard" if hero else "Card")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)
        if title:
            label = QLabel(title)
            label.setProperty("role", "sectionTitle")
            layout.addWidget(label)
        self.body_layout = layout


class BatchTableWidget(QTableWidget):
    def __init__(self, task_type: TaskType, parent: QWidget | None = None):
        super().__init__(parent)
        self.task_type = task_type
        self.columns = MESSAGE_COLUMNS if task_type is TaskType.MESSAGE else FILE_COLUMNS
        self.setColumnCount(len(self.columns))
        self.setHorizontalHeaderLabels([column.title for column in self.columns])
        self.setAlternatingRowColors(True)
        self.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        self.verticalHeader().setVisible(False)
        self.horizontalHeader().setStretchLastSection(True)
        self.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.setMinimumHeight(360)
        self.setWordWrap(True)
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self._open_context_menu)
        self._configure_columns()

    def _configure_columns(self) -> None:
        tooltips = {
            "enabled": "取消勾选后，该行不会执行，适合临时跳过某项任务。",
            "session_name": "填写微信中的好友备注、群名或会话名称，需与微信内显示一致。",
            "message": "要发送的文本内容。",
            "at_members": "群聊里要 @ 的成员，使用 | 分隔，例如 张三|李四。",
            "at_all": "仅群聊可用，需要账号有权限。",
            "clear_before_send": "勾选后发送前会先清空输入框，更适合批量任务。",
            "send_delay_sec": "留空则使用全局默认值。",
            "file_paths": "多个文件使用 | 分隔。也可以先选中行，再点“选择文件”。",
            "with_message": "发送文件时同时发一条消息。",
            "message_first": "勾选后先发消息，再发文件。",
            "remark": "只给自己看的备注，不会发给对方。",
        }
        widths = {
            "enabled": 70,
            "session_name": 170,
            "message": 280,
            "at_members": 160,
            "at_all": 90,
            "clear_before_send": 100,
            "send_delay_sec": 110,
            "remark": 150,
            "file_paths": 340,
            "with_message": 90,
            "message_first": 90,
        }
        for index, column in enumerate(self.columns):
            header_item = self.horizontalHeaderItem(index)
            if header_item is not None:
                header_item.setToolTip(tooltips.get(column.key, column.title))
            width = widths.get(column.key)
            if width is not None:
                self.setColumnWidth(index, width)

    def add_row(self, data: dict[str, Any] | None = None) -> None:
        row_index = self.rowCount()
        self.insertRow(row_index)
        data = data or {}
        for column_index, column in enumerate(self.columns):
            item = self._item_for_value(column, data.get(column.key))
            self.setItem(row_index, column_index, item)

    def load_rows(self, rows: list[MessageBatchRow] | list[FileBatchRow]) -> None:
        self.setRowCount(0)
        for row in rows:
            self.add_row(asdict(row))
        if not rows:
            self.add_row()

    def remove_selected_rows(self) -> None:
        indices = sorted(self.selected_row_indices(), reverse=True)
        for row_index in indices:
            self.removeRow(row_index)
        if self.rowCount() == 0:
            self.add_row()

    def copy_selected_rows(self) -> None:
        indices = self.selected_row_indices()
        if not indices:
            return
        lines = []
        for row_index in indices:
            values = [self._cell_display_value(row_index, column_index) for column_index in range(self.columnCount())]
            lines.append("\t".join(values))
        QGuiApplication.clipboard().setText("\n".join(lines))

    def paste_rows(self) -> None:
        text = QGuiApplication.clipboard().text().strip()
        if not text:
            return
        lines = [line for line in text.splitlines() if line.strip()]
        for line in lines:
            values = line.split("\t")
            data = {column.key: values[index] if index < len(values) else "" for index, column in enumerate(self.columns)}
            self.add_row(data)

    def choose_files_for_selected_rows(self) -> None:
        if self.task_type is not TaskType.FILE:
            return
        selected_rows = self.selected_row_indices()
        if not selected_rows:
            QMessageBox.information(self, "选择文件", "请先选中至少一行。")
            return
        paths, _ = QFileDialog.getOpenFileNames(self, "选择文件")
        if not paths:
            return
        column_index = self._column_index("file_paths")
        joined = "|".join(paths)
        for row in selected_rows:
            self.item(row, column_index).setText(joined)

    def import_rows(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "导入表格", "", "Excel/CSV (*.xlsx *.csv)")
        if not path:
            return
        rows = load_rows(self.task_type, path)
        self.load_rows(rows)

    def export_rows(self, template_only: bool = False) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "导出表格", "", "Excel (*.xlsx);;CSV (*.csv)")
        if not path:
            return
        rows = [] if template_only else self.rows_as_dicts()
        dump_rows(self.task_type, rows, path)

    def rows_as_dicts(self) -> list[dict[str, Any]]:
        result: list[dict[str, Any]] = []
        for row_index in range(self.rowCount()):
            row_data: dict[str, Any] = {}
            for column_index, column in enumerate(self.columns):
                row_data[column.key] = self._cell_value(row_index, column_index, column)
            result.append(row_data)
        return result

    def highlight_errors(self, errors: dict[int, dict[str, str]]) -> None:
        normal = QColor("#ffffff")
        invalid = QColor("#fef2f2")
        for row_index in range(self.rowCount()):
            for column_index, column in enumerate(self.columns):
                item = self.item(row_index, column_index)
                item.setBackground(invalid if column.key in errors.get(row_index, {}) else normal)
                message = errors.get(row_index, {}).get(column.key, "")
                item.setToolTip(message)

    def focus_first_error(self, errors: dict[int, dict[str, str]]) -> None:
        if not errors:
            return
        first_row = min(errors.keys())
        first_column_key = next(iter(errors[first_row].keys()))
        first_column = self._column_index(first_column_key)
        self.setCurrentCell(first_row, first_column)
        self.scrollToItem(self.item(first_row, first_column))

    def _item_for_value(self, column: ColumnSpec, value: Any) -> QTableWidgetItem:
        item = QTableWidgetItem()
        if column.kind == "bool":
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            checked = Qt.CheckState.Checked if str(value).lower() not in {"false", "0", ""} else Qt.CheckState.Unchecked
            if value is None and column.key in {"enabled", "clear_before_send"}:
                checked = Qt.CheckState.Checked
            item.setCheckState(checked)
            item.setText("")
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            return item
        item.setText("" if value is None else str(value))
        return item

    def _cell_display_value(self, row_index: int, column_index: int) -> str:
        column = self.columns[column_index]
        item = self.item(row_index, column_index)
        if column.kind == "bool":
            return "true" if item.checkState() == Qt.CheckState.Checked else "false"
        return item.text()

    def _cell_value(self, row_index: int, column_index: int, column: ColumnSpec) -> Any:
        item = self.item(row_index, column_index)
        if column.kind == "bool":
            return item.checkState() == Qt.CheckState.Checked
        text = item.text().strip()
        if column.kind == "float":
            return text
        return text

    def _column_index(self, key: str) -> int:
        for index, column in enumerate(self.columns):
            if column.key == key:
                return index
        raise KeyError(key)

    def current_column_spec(self) -> ColumnSpec | None:
        current_item = self.currentItem()
        if current_item is None:
            return None
        return self.columns[current_item.column()]

    def selected_row_indices(self) -> list[int]:
        return sorted({index.row() for index in self.selectedIndexes()})

    def duplicate_selected_rows(self) -> None:
        selected = self.selected_row_indices()
        if not selected:
            return
        copied_rows = [self._row_to_mapping(row_index) for row_index in selected]
        for row in copied_rows:
            self.add_row(row)

    def copy_previous_row_to_selected(self) -> bool:
        selected = self.selected_row_indices()
        if not selected:
            return False
        applied = False
        for row_index in selected:
            if row_index <= 0:
                continue
            previous_row = self._row_to_mapping(row_index - 1)
            for column_index, column in enumerate(self.columns):
                self.setItem(row_index, column_index, self._item_for_value(column, previous_row.get(column.key)))
            applied = True
        return applied

    def set_selected_enabled(self, enabled: bool) -> None:
        selected = self.selected_row_indices()
        if not selected:
            return
        column_index = self._column_index("enabled")
        for row_index in selected:
            self.item(row_index, column_index).setCheckState(
                Qt.CheckState.Checked if enabled else Qt.CheckState.Unchecked
            )

    def clear_all_rows(self) -> None:
        self.setRowCount(0)
        self.add_row()

    def apply_current_cell_to_selected_rows(self) -> tuple[bool, str]:
        current_item = self.currentItem()
        if current_item is None:
            return False, "请先把光标放到要复用的单元格里。"
        selected = self.selected_row_indices()
        if len(selected) <= 1:
            return False, "请至少选中两行，才能把当前值套用到其它行。"
        current_row = current_item.row()
        current_column = current_item.column()
        column = self.columns[current_column]
        value = self._cell_value(current_row, current_column, column)
        for row_index in selected:
            if row_index == current_row:
                continue
            self.setItem(row_index, current_column, self._item_for_value(column, value))
        return True, f"已将“{column.title}”列的当前值套用到 {len(selected) - 1} 行。"

    def apply_value_to_selected_rows(self, value: Any) -> tuple[bool, str]:
        current_item = self.currentItem()
        if current_item is None:
            return False, "请先选中要批量填写的那一列中的任意单元格。"
        selected = self.selected_row_indices()
        if not selected:
            return False, "请先选中至少一行。"
        current_column = current_item.column()
        column = self.columns[current_column]
        for row_index in selected:
            self.setItem(row_index, current_column, self._item_for_value(column, value))
        return True, f"已将“{column.title}”列批量填写到 {len(selected)} 行。"

    def _row_to_mapping(self, row_index: int) -> dict[str, Any]:
        row_data: dict[str, Any] = {}
        for column_index, column in enumerate(self.columns):
            row_data[column.key] = self._cell_value(row_index, column_index, column)
        return row_data

    def _open_context_menu(self, position) -> None:
        menu = QMenu(self)
        duplicate_action = menu.addAction("复制选中行")
        enable_action = menu.addAction("启用选中行")
        disable_action = menu.addAction("停用选中行")
        delete_action = menu.addAction("删除选中行")
        selected_action = menu.exec(self.viewport().mapToGlobal(position))
        if selected_action == duplicate_action:
            self.duplicate_selected_rows()
        elif selected_action == enable_action:
            self.set_selected_enabled(True)
        elif selected_action == disable_action:
            self.set_selected_enabled(False)
        elif selected_action == delete_action:
            self.remove_selected_rows()


def example_rows_for(task_type: TaskType) -> list[MessageBatchRow] | list[FileBatchRow]:
    if task_type is TaskType.MESSAGE:
        return [
            MessageBatchRow(session_name="客户A", message="您好，资料已经整理好了，稍后发您。", remark="示例任务"),
            MessageBatchRow(session_name="项目群", message="大家下午 3 点开会，请提前 5 分钟上线。", remark="示例任务"),
        ]
    return [
        FileBatchRow(
            session_name="客户A",
            file_paths="C:/合同/报价单.pdf",
            with_message=True,
            message="您好，附件请查收。",
            remark="示例任务",
        ),
        FileBatchRow(
            session_name="财务群",
            file_paths="C:/报表/本周统计.xlsx|C:/报表/本周明细.xlsx",
            with_message=False,
            remark="示例任务",
        ),
    ]


class DashboardPage(QWidget):
    refresh_requested = Signal()
    open_message_requested = Signal()
    open_file_requested = Signal()
    open_templates_requested = Signal()

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setSpacing(16)

        self.hero_card = CardFrame(hero=True)
        title = QLabel("PyWeChat 办公助手")
        title.setProperty("role", "pageTitle")
        subtitle = QLabel("把常见的微信批量发送工作，整理成清楚、可重复、可追踪的办公流程。")
        subtitle.setProperty("role", "pageSubtitle")
        subtitle.setWordWrap(True)
        self.hero_card.body_layout.addWidget(title)
        self.hero_card.body_layout.addWidget(subtitle)
        self.hero_hint = QLabel("推荐顺序：先检查微信状态，再去批量页面导入表格，最后校验并执行。")
        self.hero_hint.setProperty("role", "good")
        self.hero_hint.setWordWrap(True)
        self.hero_card.body_layout.addWidget(self.hero_hint)
        quick_actions = QHBoxLayout()
        self.open_message_button = QPushButton("去批量消息")
        self.open_file_button = QPushButton("去批量文件")
        self.open_templates_button = QPushButton("查看模板")
        self.open_file_button.setProperty("variant", "secondary")
        self.open_templates_button.setProperty("variant", "ghost")
        quick_actions.addWidget(self.open_message_button)
        quick_actions.addWidget(self.open_file_button)
        quick_actions.addWidget(self.open_templates_button)
        quick_actions.addStretch(1)
        self.hero_card.body_layout.addLayout(quick_actions)
        self.open_message_button.clicked.connect(lambda: self.open_message_requested.emit())
        self.open_file_button.clicked.connect(lambda: self.open_file_requested.emit())
        self.open_templates_button.clicked.connect(lambda: self.open_templates_requested.emit())
        layout.addWidget(self.hero_card)

        self.env_card = CardFrame("环境自检")
        grid = QGridLayout()
        grid.setHorizontalSpacing(16)
        grid.setVerticalSpacing(12)
        self.labels: dict[str, QLabel] = {}
        items = [
            ("operating_system", "操作系统"),
            ("python_version", "Python"),
            ("gui_version", "GUI 版本"),
            ("wechat_running", "微信进程"),
            ("login_status", "登录状态"),
            ("wechat_path", "微信路径"),
            ("status_message", "连接结果"),
        ]
        for row_index, (key, title) in enumerate(items):
            grid.addWidget(QLabel(f"{title}:"), row_index, 0)
            value_label = QLabel("-")
            value_label.setWordWrap(True)
            self.labels[key] = value_label
            grid.addWidget(value_label, row_index, 1)
        self.refresh_button = QPushButton("连接微信并刷新状态")
        self.env_card.body_layout.addLayout(grid)
        self.env_card.body_layout.addWidget(self.refresh_button)
        self.refresh_button.clicked.connect(lambda: self.refresh_requested.emit())
        layout.addWidget(self.env_card)

        self.steps_card = CardFrame("三步开始")
        steps = QLabel(
            "1. 点击“连接微信并刷新状态”，确认环境正常。\n"
            "2. 在批量页面导入 Excel/CSV，或直接填写表格。\n"
            "3. 先点“校验”，没有高亮报错后再执行。"
        )
        steps.setProperty("role", "hint")
        steps.setWordWrap(True)
        self.steps_card.body_layout.addWidget(steps)
        layout.addWidget(self.steps_card)

        self.advice_card = CardFrame("风险提示")
        self.advice_label = QLabel()
        self.advice_label.setWordWrap(True)
        self.advice_card.body_layout.addWidget(self.advice_label)
        layout.addWidget(self.advice_card)
        layout.addStretch(1)

        self.advice_label.setText(
            "1. 自动化执行期间会接管鼠标与键盘，请勿手动操作微信。\n"
            "2. 若无法识别微信主界面，可能需要在微信登录前开启 Windows 讲述人或相关无障碍服务。\n"
            "3. 请保持微信已登录且桌面会话未锁屏。"
        )

    def set_environment(self, status) -> None:
        self.labels["operating_system"].setText(status.operating_system)
        self.labels["python_version"].setText(status.python_version)
        self.labels["gui_version"].setText(status.gui_version)
        self.labels["wechat_running"].setText("运行中" if status.wechat_running else "未运行")
        self.labels["login_status"].setText(status.login_status)
        self.labels["wechat_path"].setText(status.wechat_path or "-")
        self.labels["status_message"].setText(status.status_message)
        if status.advice:
            self.advice_label.setText("\n".join(f"{index + 1}. {item}" for index, item in enumerate(status.advice)))
        self.hero_hint.setProperty("role", "good" if status.wechat_running and status.login_status == "已登录" else "warn")
        self.hero_hint.setText(status.status_message)
        self.hero_hint.style().unpolish(self.hero_hint)
        self.hero_hint.style().polish(self.hero_hint)


class BatchPage(QWidget):
    run_requested = Signal(object, object)
    stop_requested = Signal()
    save_template_requested = Signal(object, object)
    open_templates_requested = Signal(object)

    def __init__(self, task_type: TaskType, title: str, parent: QWidget | None = None):
        super().__init__(parent)
        self.task_type = task_type
        self.pending_source_execution_id: int | None = None
        self._buttons: dict[str, QPushButton] = {}
        layout = QVBoxLayout(self)
        card = CardFrame(title, hero=True)
        subtitle = QLabel(
            "只需要维护“启用”的行。推荐顺序：导入或填写 -> 校验 -> 执行。"
            if task_type is TaskType.MESSAGE
            else "先选中行再点“选择文件”会更省事。多个文件支持用 | 分隔。"
        )
        subtitle.setProperty("role", "pageSubtitle")
        subtitle.setWordWrap(True)
        card.body_layout.addWidget(subtitle)
        helper = QLabel("“备注”列只给自己看，适合写客户分类、批次说明或提醒。")
        helper.setProperty("role", "hint")
        helper.setWordWrap(True)
        card.body_layout.addWidget(helper)
        required_hint = QLabel(
            "必填列：会话名称、消息内容。群聊 @ 成员请用 | 分隔。"
            if task_type is TaskType.MESSAGE
            else "必填列：会话名称、文件路径。多个文件用 | 分隔；附带消息时还需要填写消息内容。"
        )
        required_hint.setProperty("role", "warn")
        required_hint.setWordWrap(True)
        card.body_layout.addWidget(required_hint)
        toolbar = QHBoxLayout()
        button_specs = [
            ("新增行", self._add_row),
            ("载入示例", self._load_example_rows),
            ("复制上一行", self._copy_previous_row),
            ("套用当前值", self._apply_current_value),
            ("批量填写当前列", self._fill_current_column),
            ("启用选中", self._enable_selected_rows),
            ("停用选中", self._disable_selected_rows),
            ("删除选中", self._remove_rows),
            ("复制行", self._copy_rows),
            ("粘贴行", self._paste_rows),
            ("导入", self._import_rows),
            ("导出当前", self._export_rows),
            ("导出模板", self._export_template),
        ]
        if task_type is TaskType.FILE:
            button_specs.append(("选择文件", self._choose_files))
        button_specs.extend(
            [
                ("保存模板", self._save_template),
                ("加载模板", self._open_templates),
                ("校验", self.validate_rows),
                ("执行", self._request_run),
                ("清空表格", self._clear_table),
                ("停止", self.stop_requested.emit),
            ]
        )
        for text, callback in button_specs:
            button = QPushButton(text)
            if text in {"删除选中", "停止", "保存模板", "加载模板", "清空表格"}:
                button.setProperty("variant", "secondary")
            if text in {"载入示例", "导入", "导出当前", "导出模板", "复制上一行", "套用当前值", "批量填写当前列"}:
                button.setProperty("variant", "ghost")
            toolbar.addWidget(button)
            self._buttons[text] = button
            button.clicked.connect(lambda _checked=False, cb=callback: cb())
        toolbar.addStretch(1)
        card.body_layout.addLayout(toolbar)
        self.table = BatchTableWidget(task_type)
        card.body_layout.addWidget(self.table)
        self.summary_label = QLabel("准备就绪")
        self.summary_label.setProperty("role", "muted")
        self.summary_label.setWordWrap(True)
        card.body_layout.addWidget(self.summary_label)
        layout.addWidget(card)
        self.table.load_rows([])
        self.set_running_state(False)

    def rows(self) -> list[MessageBatchRow] | list[FileBatchRow]:
        mappings = self.table.rows_as_dicts()
        if self.task_type is TaskType.MESSAGE:
            return [MessageBatchRow.from_mapping(row) for row in mappings]
        return [FileBatchRow.from_mapping(row) for row in mappings]

    def load_rows(self, rows: list[MessageBatchRow] | list[FileBatchRow], source_execution_id: int | None = None) -> None:
        self.pending_source_execution_id = source_execution_id
        cloned_rows = [clone_row(row) for row in rows]
        self.table.load_rows(cloned_rows)
        self._update_summary(f"已加载 {len(cloned_rows)} 行任务。")

    def validate_rows(self) -> tuple[list[MessageBatchRow] | list[FileBatchRow], dict[int, dict[str, str]]]:
        rows = self.rows()
        all_errors: dict[int, dict[str, str]] = {}
        enabled_rows = 0
        for row_index, row in enumerate(rows):
            if not row.enabled:
                continue
            enabled_rows += 1
            errors = row.validate()
            if errors:
                all_errors[row_index] = errors
        self.table.highlight_errors(all_errors)
        if all_errors:
            self.table.focus_first_error(all_errors)
            self._update_summary(f"表格中将执行 {enabled_rows} 行，其中有 {len(all_errors)} 行需要先修复。")
        else:
            self._update_summary(f"校验通过。表格共 {len(rows)} 行，实际会执行 {enabled_rows} 行。")
        return rows, all_errors

    def _request_run(self) -> None:
        rows, errors = self.validate_rows()
        enabled_rows = [row for row in rows if row.enabled]
        if not enabled_rows:
            QMessageBox.information(self, "没有可执行任务", "当前没有勾选“启用”的任务行。请先勾选至少一行再执行。")
            self._update_summary("当前没有启用的任务行，请先勾选需要执行的内容。")
            return
        if errors:
            line_numbers = [str(index + 1) for index in sorted(errors.keys())[:5]]
            extra = " 等" if len(errors) > 5 else ""
            QMessageBox.warning(
                self,
                "校验失败",
                f"请先修复表格中的高亮字段。\n出错行：第 {', '.join(line_numbers)} 行{extra}",
            )
            return
        self.run_requested.emit(rows, self.pending_source_execution_id)
        self.pending_source_execution_id = None

    def _save_template(self) -> None:
        rows, errors = self.validate_rows()
        if errors:
            QMessageBox.warning(self, "无法保存模板", "请先修复校验错误后再保存模板。")
            return
        self.save_template_requested.emit(self.task_type, rows)

    def _open_templates(self) -> None:
        self.open_templates_requested.emit(self.task_type)

    def _add_row(self) -> None:
        self.table.add_row()

    def _remove_rows(self) -> None:
        self.table.remove_selected_rows()

    def _copy_rows(self) -> None:
        self.table.copy_selected_rows()

    def _paste_rows(self) -> None:
        self.table.paste_rows()

    def _enable_selected_rows(self) -> None:
        self.table.set_selected_enabled(True)
        self._update_summary("已启用选中行。")

    def _disable_selected_rows(self) -> None:
        self.table.set_selected_enabled(False)
        self._update_summary("已停用选中行。")

    def _import_rows(self) -> None:
        try:
            self.table.import_rows()
            self._update_summary(f"导入完成，共 {self.table.rowCount()} 行。建议先点“校验”。")
        except Exception as exc:
            QMessageBox.critical(self, "导入失败", str(exc))

    def _export_rows(self) -> None:
        try:
            self.table.export_rows(template_only=False)
        except Exception as exc:
            QMessageBox.critical(self, "导出失败", str(exc))

    def _export_template(self) -> None:
        try:
            self.table.export_rows(template_only=True)
        except Exception as exc:
            QMessageBox.critical(self, "导出模板失败", str(exc))

    def _choose_files(self) -> None:
        self.table.choose_files_for_selected_rows()

    def _load_example_rows(self) -> None:
        self.load_rows(example_rows_for(self.task_type))
        QMessageBox.information(self, "已载入示例", "示例数据已经放入表格，你可以直接改成自己的内容。")

    def _copy_previous_row(self) -> None:
        applied = self.table.copy_previous_row_to_selected()
        if not applied:
            QMessageBox.information(self, "复制上一行", "请先选中第 2 行及之后的行，才能复制上一行内容。")
            return
        self._update_summary("已把上一行内容复制到选中行。")

    def _apply_current_value(self) -> None:
        applied, message = self.table.apply_current_cell_to_selected_rows()
        if not applied:
            QMessageBox.information(self, "套用当前值", message)
            return
        self._update_summary(message)

    def _fill_current_column(self) -> None:
        column = self.table.current_column_spec()
        if column is None:
            QMessageBox.information(self, "批量填写当前列", "请先选中要填写的列中的任意单元格。")
            return
        if column.kind == "bool":
            label, ok = QInputDialog.getItem(self, "批量填写当前列", f"将“{column.title}”设置为：", ["是", "否"], editable=False)
            if not ok:
                return
            value = label == "是"
        else:
            text, ok = QInputDialog.getText(self, "批量填写当前列", f"为“{column.title}”输入统一内容：")
            if not ok:
                return
            value = text
        applied, message = self.table.apply_value_to_selected_rows(value)
        if not applied:
            QMessageBox.information(self, "批量填写当前列", message)
            return
        self._update_summary(message)

    def _clear_table(self) -> None:
        answer = QMessageBox.question(
            self,
            "清空表格",
            "确定清空当前表格吗？清空后会保留一行空白行，方便继续录入。",
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.table.clear_all_rows()
        self.pending_source_execution_id = None
        self._update_summary("表格已清空。")

    def _update_summary(self, text: str) -> None:
        self.summary_label.setText(text)

    def set_running_state(self, is_running: bool) -> None:
        for name, button in self._buttons.items():
            if name == "停止":
                button.setEnabled(is_running)
            elif name == "执行":
                button.setEnabled(not is_running)
            else:
                button.setEnabled(not is_running)
        self.table.setDisabled(is_running)


class TemplatesPage(QWidget):
    load_template_requested = Signal(object, object)

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        card = CardFrame("模板中心")
        toolbar = QHBoxLayout()
        self.refresh_button = QPushButton("刷新")
        self.rename_button = QPushButton("重命名")
        self.rename_button.setProperty("variant", "secondary")
        self.duplicate_button = QPushButton("复制")
        self.duplicate_button.setProperty("variant", "secondary")
        self.delete_button = QPushButton("删除")
        self.delete_button.setProperty("variant", "secondary")
        self.load_button = QPushButton("加载到工作台")
        for button in [self.refresh_button, self.rename_button, self.duplicate_button, self.delete_button, self.load_button]:
            toolbar.addWidget(button)
        toolbar.addStretch(1)
        card.body_layout.addLayout(toolbar)
        helper = QLabel("把常用任务保存成模板，下次可以直接加载，适合固定通知、固定资料发送。")
        helper.setProperty("role", "hint")
        helper.setWordWrap(True)
        card.body_layout.addWidget(helper)
        stats_layout = QHBoxLayout()
        self.total_templates_label = self._create_stat_card("模板总数", "0")
        self.message_templates_label = self._create_stat_card("消息模板", "0")
        self.file_templates_label = self._create_stat_card("文件模板", "0")
        for stat_card in [self.total_templates_label, self.message_templates_label, self.file_templates_label]:
            stats_layout.addWidget(stat_card)
        card.body_layout.addLayout(stats_layout)
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("搜索模板名称或类型，例如：消息、客户、通知")
        card.body_layout.addWidget(self.search_input)
        self.summary_label = QLabel("还没有模板。你可以先去批量页面整理一份任务，再点“保存模板”。")
        self.summary_label.setProperty("role", "muted")
        self.summary_label.setWordWrap(True)
        card.body_layout.addWidget(self.summary_label)
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["ID", "名称", "类型", "更新时间"])
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        card.body_layout.addWidget(self.table)
        layout.addWidget(card)

    @staticmethod
    def _create_stat_card(title: str, value: str) -> QWidget:
        card = CardFrame()
        title_label = QLabel(title)
        title_label.setProperty("role", "statTitle")
        value_label = QLabel(value)
        value_label.setProperty("role", "statValue")
        card.body_layout.addWidget(title_label)
        card.body_layout.addWidget(value_label)
        card.value_label = value_label  # type: ignore[attr-defined]
        return card


class HistoryPage(QWidget):
    retry_failed_requested = Signal(object, object, object)
    clear_history_requested = Signal()

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        layout = QVBoxLayout(self)

        history_card = CardFrame("执行历史")
        toolbar = QHBoxLayout()
        self.refresh_button = QPushButton("刷新")
        self.retry_button = QPushButton("基于失败项重建任务")
        self.export_failed_button = QPushButton("导出失败项")
        self.clear_button = QPushButton("清理历史")
        self.clear_button.setProperty("variant", "secondary")
        self.failed_only_checkbox = QCheckBox("只看含失败记录")
        for button in [self.refresh_button, self.retry_button, self.export_failed_button, self.clear_button]:
            toolbar.addWidget(button)
        toolbar.addWidget(self.failed_only_checkbox)
        toolbar.addStretch(1)
        history_card.body_layout.addLayout(toolbar)
        helper = QLabel("先看成功/失败数量，再点开失败项。失败项可以一键回填到批量页面重新执行。")
        helper.setProperty("role", "hint")
        helper.setWordWrap(True)
        history_card.body_layout.addWidget(helper)
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("搜索会话、状态或任务类型，例如：失败、客户A、文件")
        history_card.body_layout.addWidget(self.search_input)
        stats_layout = QHBoxLayout()
        self.total_runs_label = self._create_stat_card("总记录", "0")
        self.success_runs_label = self._create_stat_card("全成功", "0")
        self.failed_runs_label = self._create_stat_card("含失败", "0")
        for card_widget in [self.total_runs_label, self.success_runs_label, self.failed_runs_label]:
            stats_layout.addWidget(card_widget)
        history_card.body_layout.addLayout(stats_layout)
        self.summary_label = QLabel("还没有执行历史。第一次成功或失败执行后，这里会显示完整记录。")
        self.summary_label.setProperty("role", "muted")
        self.summary_label.setWordWrap(True)
        history_card.body_layout.addWidget(self.summary_label)
        self.failure_summary_label = QLabel("选中一条执行记录后，这里会展示失败原因摘要。")
        self.failure_summary_label.setProperty("role", "hint")
        self.failure_summary_label.setWordWrap(True)
        history_card.body_layout.addWidget(self.failure_summary_label)

        self.execution_table = QTableWidget(0, 7)
        self.execution_table.setHorizontalHeaderLabels(
            ["ID", "任务类型", "开始时间", "状态", "总行数", "成功", "失败"]
        )
        self.execution_table.verticalHeader().setVisible(False)
        self.execution_table.horizontalHeader().setStretchLastSection(True)
        self.execution_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.execution_table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        history_card.body_layout.addWidget(self.execution_table)

        detail_group = QGroupBox("逐行结果")
        detail_layout = QVBoxLayout(detail_group)
        self.detail_table = QTableWidget(0, 5)
        self.detail_table.setHorizontalHeaderLabels(["行号", "会话", "结果", "错误码", "错误信息"])
        self.detail_table.verticalHeader().setVisible(False)
        self.detail_table.horizontalHeader().setStretchLastSection(True)
        detail_layout.addWidget(self.detail_table)
        self.copy_diag_button = QPushButton("复制诊断信息")
        self.copy_diag_button.setProperty("variant", "secondary")
        detail_layout.addWidget(self.copy_diag_button)
        self.diagnostic_text = QTextEdit()
        self.diagnostic_text.setReadOnly(True)
        self.diagnostic_text.setPlaceholderText("选中失败行后，可在这里查看原始诊断信息。")
        detail_layout.addWidget(self.diagnostic_text)
        history_card.body_layout.addWidget(detail_group)
        layout.addWidget(history_card)

    @staticmethod
    def _create_stat_card(title: str, value: str) -> QWidget:
        card = CardFrame()
        title_label = QLabel(title)
        title_label.setProperty("role", "statTitle")
        value_label = QLabel(value)
        value_label.setProperty("role", "statValue")
        card.body_layout.addWidget(title_label)
        card.body_layout.addWidget(value_label)
        card.value_label = value_label  # type: ignore[attr-defined]
        return card


class SettingsPage(QWidget):
    save_requested = Signal(object)

    def __init__(self, parent: QWidget | None = None):
        from PySide6.QtWidgets import QCheckBox, QComboBox, QDoubleSpinBox, QFormLayout, QSpinBox

        super().__init__(parent)
        layout = QVBoxLayout(self)
        card = CardFrame("设置")
        helper = QLabel("一般情况下保持默认即可。如果你不确定某个选项的作用，建议先不要修改。")
        helper.setProperty("role", "hint")
        helper.setWordWrap(True)
        card.body_layout.addWidget(helper)
        form = QFormLayout()
        form.setSpacing(14)

        self.is_maximize = QCheckBox("自动最大化微信主界面")
        self.close_weixin = QCheckBox("任务完成后关闭微信")
        self.clear = QCheckBox("发送前清空输入区")
        self.search_pages = QSpinBox()
        self.search_pages.setRange(0, 50)
        self.send_delay = QDoubleSpinBox()
        self.send_delay.setRange(0.0, 10.0)
        self.send_delay.setSingleStep(0.1)
        self.window_width = QSpinBox()
        self.window_width.setRange(960, 3840)
        self.window_height = QSpinBox()
        self.window_height.setRange(720, 2160)
        self.import_encoding = QComboBox()
        self.import_encoding.addItems(["auto", "utf-8-sig", "gbk"])
        self.theme = QComboBox()
        self.theme.addItems(["light"])

        form.addRow("窗口最大化", self.is_maximize)
        form.addRow("完成后关闭微信", self.close_weixin)
        form.addRow("清空输入框", self.clear)
        form.addRow("搜索页数", self.search_pages)
        form.addRow("发送间隔(秒)", self.send_delay)
        form.addRow("默认窗口宽度", self.window_width)
        form.addRow("默认窗口高度", self.window_height)
        form.addRow("CSV 编码", self.import_encoding)
        form.addRow("主题", self.theme)
        card.body_layout.addLayout(form)
        self.save_button = QPushButton("保存设置")
        card.body_layout.addWidget(self.save_button)
        layout.addWidget(card)
        layout.addStretch(1)
