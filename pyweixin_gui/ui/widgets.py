from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Any

from PySide6.QtCore import QPoint, QRect, QSize, Qt, Signal
from PySide6.QtGui import QColor, QDragEnterEvent, QDropEvent, QGuiApplication
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
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
    QSpinBox,
    QLayout,
    QSizePolicy,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from ..import_export import dump_route_rows, dump_rows, dump_table, load_route_rows, load_rows, load_session_names
from ..models import ChatBatchExportRequest, ChatExportRequest, ExportHistoryRecord, FileBatchRow, GroupSummaryRow, MessageBatchRow, RelayCollectFilesRequest, RelayCollectTextRequest, RelayItemType, RelayPackageRow, RelayRouteRow, RelaySendRequest, RelayValidationRequest, ResourceExportKind, ResourceExportRequest, SessionScanRequest, SessionSummaryRow, TaskType, clone_row


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

RELAY_ITEM_LABELS = {
    RelayItemType.TEXT.value: "文本",
    RelayItemType.FILE.value: "文件",
    RelayItemType.IMAGE.value: "图片",
}


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


class FlowLayout(QLayout):
    def __init__(self, parent: QWidget | None = None, margin: int = 0, h_spacing: int = 8, v_spacing: int = 8):
        super().__init__(parent)
        self._items: list[Any] = []
        self._h_spacing = h_spacing
        self._v_spacing = v_spacing
        self.setContentsMargins(margin, margin, margin, margin)

    def addItem(self, item) -> None:  # type: ignore[override]
        self._items.append(item)

    def count(self) -> int:  # type: ignore[override]
        return len(self._items)

    def itemAt(self, index: int):  # type: ignore[override]
        return self._items[index] if 0 <= index < len(self._items) else None

    def takeAt(self, index: int):  # type: ignore[override]
        if 0 <= index < len(self._items):
            return self._items.pop(index)
        return None

    def expandingDirections(self):  # type: ignore[override]
        return Qt.Orientation(0)

    def hasHeightForWidth(self) -> bool:  # type: ignore[override]
        return True

    def heightForWidth(self, width: int) -> int:  # type: ignore[override]
        return self._do_layout(QRect(0, 0, width, 0), test_only=True)

    def setGeometry(self, rect: QRect) -> None:  # type: ignore[override]
        super().setGeometry(rect)
        self._do_layout(rect, test_only=False)

    def sizeHint(self) -> QSize:  # type: ignore[override]
        return self.minimumSize()

    def minimumSize(self) -> QSize:  # type: ignore[override]
        size = QSize()
        for item in self._items:
            size = size.expandedTo(item.minimumSize())
        margins = self.contentsMargins()
        size += QSize(margins.left() + margins.right(), margins.top() + margins.bottom())
        return size

    def _do_layout(self, rect: QRect, test_only: bool) -> int:
        margins = self.contentsMargins()
        effective_rect = rect.adjusted(margins.left(), margins.top(), -margins.right(), -margins.bottom())
        x = effective_rect.x()
        y = effective_rect.y()
        line_height = 0

        for item in self._items:
            hint = item.sizeHint()
            next_x = x + hint.width() + self._h_spacing
            if line_height > 0 and next_x - self._h_spacing > effective_rect.right() + 1:
                x = effective_rect.x()
                y += line_height + self._v_spacing
                next_x = x + hint.width() + self._h_spacing
                line_height = 0
            if not test_only:
                item.setGeometry(QRect(QPoint(x, y), hint))
            x = next_x
            line_height = max(line_height, hint.height())

        return y + line_height - rect.y() + margins.bottom()


def _compact_row(*widgets: QWidget, stretch_last: bool = True) -> QWidget:
    container = QWidget()
    layout = QHBoxLayout(container)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(8)
    for index, widget in enumerate(widgets):
        if stretch_last and index == len(widgets) - 1:
            layout.addWidget(widget, 1)
        else:
            layout.addWidget(widget)
    if not stretch_last:
        layout.addStretch(1)
    return container


def _flow_container(*widgets: QWidget, h_spacing: int = 8, v_spacing: int = 8) -> QWidget:
    container = QWidget()
    container.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
    layout = FlowLayout(container, margin=0, h_spacing=h_spacing, v_spacing=v_spacing)
    for widget in widgets:
        layout.addWidget(widget)
    return container


def _set_compact_width(widget: QWidget, width: int) -> None:
    widget.setMinimumWidth(width)
    widget.setMaximumWidth(width)


def _configure_form_layout(form: QFormLayout) -> None:
    form.setSpacing(14)
    form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
    form.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
    form.setLabelAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)


def _configure_data_table(table: QTableWidget, minimum_height: int | None = None) -> None:
    table.setAlternatingRowColors(True)
    table.setWordWrap(True)
    table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
    table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
    header = table.horizontalHeader()
    header.setDefaultAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
    header.setMinimumHeight(42)
    vertical_header = table.verticalHeader()
    vertical_header.setDefaultSectionSize(38)
    vertical_header.setMinimumSectionSize(34)
    if minimum_height is not None:
        table.setMinimumHeight(minimum_height)


class BatchTableWidget(QTableWidget):
    def __init__(self, task_type: TaskType, parent: QWidget | None = None):
        super().__init__(parent)
        self.task_type = task_type
        self.columns = MESSAGE_COLUMNS if task_type is TaskType.MESSAGE else FILE_COLUMNS
        self.setColumnCount(len(self.columns))
        self.setHorizontalHeaderLabels([column.title for column in self.columns])
        _configure_data_table(self, minimum_height=360)
        self.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        self.verticalHeader().setVisible(False)
        self.horizontalHeader().setStretchLastSection(True)
        self.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
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
        item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
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


class RelayPackageTableWidget(QTableWidget):
    files_dropped = Signal(list)

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.viewport().setAcceptDrops(True)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if self._dropped_file_paths(event):
            event.acceptProposedAction()
            return
        super().dragEnterEvent(event)

    def dragMoveEvent(self, event) -> None:  # type: ignore[override]
        if self._dropped_file_paths(event):
            event.acceptProposedAction()
            return
        super().dragMoveEvent(event)

    def dropEvent(self, event: QDropEvent) -> None:
        file_paths = self._dropped_file_paths(event)
        if file_paths:
            self.files_dropped.emit(file_paths)
            event.acceptProposedAction()
            return
        super().dropEvent(event)

    @staticmethod
    def _dropped_file_paths(event) -> list[str]:
        mime_data = event.mimeData()
        if mime_data is None or not mime_data.hasUrls():
            return []
        return [url.toLocalFile() for url in mime_data.urls() if url.isLocalFile()]


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
        title = QLabel("AutoWeChat 办公助手")
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
        self.open_message_button = QPushButton("去批量消息")
        self.open_file_button = QPushButton("去批量文件")
        self.open_templates_button = QPushButton("查看模板")
        self.open_file_button.setProperty("variant", "secondary")
        self.open_templates_button.setProperty("variant", "ghost")
        self.hero_card.body_layout.addWidget(
            _flow_container(self.open_message_button, self.open_file_button, self.open_templates_button)
        )
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
        action_buttons: list[QPushButton] = []
        for text, callback in button_specs:
            button = QPushButton(text)
            if text in {"删除选中", "停止", "保存模板", "加载模板", "清空表格"}:
                button.setProperty("variant", "secondary")
            if text in {"载入示例", "导入", "导出当前", "导出模板", "复制上一行", "套用当前值", "批量填写当前列"}:
                button.setProperty("variant", "ghost")
            self._buttons[text] = button
            button.clicked.connect(lambda _checked=False, cb=callback: cb())
            action_buttons.append(button)
        card.body_layout.addWidget(_flow_container(*action_buttons))
        self.table = BatchTableWidget(task_type)
        card.body_layout.addWidget(self.table)
        self.summary_label = QLabel("准备就绪")
        self.summary_label.setProperty("role", "muted")
        self.summary_label.setWordWrap(True)
        card.body_layout.addWidget(self.summary_label)
        layout.addWidget(card)
        layout.addStretch(1)
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


class ExportPage(QWidget):
    export_requested = Signal(object)
    batch_export_requested = Signal(object)
    stop_requested = Signal()

    def __init__(self, parent: QWidget | None = None):
        from PySide6.QtWidgets import QCheckBox, QFormLayout, QSpinBox

        super().__init__(parent)
        self._buttons: dict[str, QPushButton] = {}
        layout = QVBoxLayout(self)
        card = CardFrame("会话导出", hero=True)
        subtitle = QLabel("一键导出指定好友或群聊的文本记录和聊天文件，适合做留档、交接和二次处理。")
        subtitle.setProperty("role", "pageSubtitle")
        subtitle.setWordWrap(True)
        card.body_layout.addWidget(subtitle)
        helper = QLabel(
            "可导出聊天文字记录和聊天文件。若相关文件尚未保存在本机，请先在微信里打开或下载后再试。"
        )
        helper.setProperty("role", "warn")
        helper.setWordWrap(True)
        card.body_layout.addWidget(helper)

        form = QFormLayout()
        _configure_form_layout(form)
        self.session_name_input = QLineEdit()
        self.session_name_input.setPlaceholderText("填写微信群名、好友备注或会话名称")
        self.session_names_input = QTextEdit()
        self.session_names_input.setPlaceholderText("批量导出时，每行填写一个会话名称，例如：\n项目群\n客户A\n财务群")
        self.session_names_input.setMinimumHeight(96)
        self.target_folder_input = QLineEdit()
        self.target_folder_input.setPlaceholderText("选择导出目录")
        self.choose_folder_button = QPushButton("选择文件夹")
        self.choose_folder_button.setProperty("variant", "ghost")
        self.choose_folder_button.setMinimumWidth(104)
        target_row = QHBoxLayout()
        target_row.addWidget(self.target_folder_input)
        target_row.addWidget(self.choose_folder_button)

        self.export_messages_checkbox = QCheckBox("导出文本消息")
        self.export_messages_checkbox.setChecked(True)
        self.export_files_checkbox = QCheckBox("导出聊天文件")
        self.export_files_checkbox.setChecked(True)

        self.message_limit_spin = QSpinBox()
        self.message_limit_spin.setRange(1, 5000)
        self.message_limit_spin.setValue(100)
        _set_compact_width(self.message_limit_spin, 128)
        self.file_limit_spin = QSpinBox()
        self.file_limit_spin.setRange(1, 5000)
        self.file_limit_spin.setValue(50)
        _set_compact_width(self.file_limit_spin, 128)

        message_limit_label = QLabel("消息条数")
        message_limit_label.setProperty("role", "muted")
        file_limit_label = QLabel("文件数量")
        file_limit_label.setProperty("role", "muted")
        count_row = _flow_container(message_limit_label, self.message_limit_spin, file_limit_label, self.file_limit_spin, h_spacing=12)

        scope_row = _flow_container(self.export_messages_checkbox, self.export_files_checkbox, h_spacing=16)

        form.addRow("会话名称", self.session_name_input)
        form.addRow("批量会话", self.session_names_input)
        form.addRow("导出目录", target_row)
        form.addRow("导出数量", count_row)
        form.addRow("导出范围", scope_row)
        card.body_layout.addLayout(form)

        toolbar_buttons: list[QPushButton] = []
        for text, callback in [
            ("示例填充", self._load_example),
            ("导入会话名单", self._import_session_names),
            ("一键导出", self._request_export),
            ("批量导出会话", self._request_batch_export),
            ("停止", self.stop_requested.emit),
        ]:
            button = QPushButton(text)
            if text in {"示例填充", "批量导出会话", "导入会话名单"}:
                button.setProperty("variant", "ghost")
            if text == "停止":
                button.setProperty("variant", "secondary")
            self._buttons[text] = button
            button.clicked.connect(lambda _checked=False, cb=callback: cb())
            toolbar_buttons.append(button)
        card.body_layout.addWidget(_flow_container(*toolbar_buttons))

        self.summary_label = QLabel("准备就绪。建议先确认微信状态正常，再导出指定会话。")
        self.summary_label.setProperty("role", "muted")
        self.summary_label.setWordWrap(True)
        card.body_layout.addWidget(self.summary_label)

        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setPlaceholderText("导出完成后，这里会显示导出结果、目录和注意事项。")
        card.body_layout.addWidget(self.result_text)
        self.open_folder_button = QPushButton("打开导出目录")
        self.open_folder_button.setProperty("variant", "secondary")
        card.body_layout.addWidget(self.open_folder_button)
        layout.addWidget(card)
        layout.addStretch(1)

        self.choose_folder_button.clicked.connect(self._choose_folder)
        self.open_folder_button.setEnabled(False)
        self.set_running_state(False)

    def current_request(self) -> tuple[ChatExportRequest | None, dict[str, str]]:
        request = ChatExportRequest(
            session_name=self.session_name_input.text().strip(),
            target_folder=self.target_folder_input.text().strip(),
            export_messages=self.export_messages_checkbox.isChecked(),
            export_files=self.export_files_checkbox.isChecked(),
            export_images=False,
            message_limit=self.message_limit_spin.value(),
            file_limit=self.file_limit_spin.value(),
        )
        errors = request.validate()
        return (None, errors) if errors else (request, {})

    def _request_export(self) -> None:
        request, errors = self.current_request()
        if errors:
            first_error = next(iter(errors.values()))
            QMessageBox.warning(self, "导出参数不完整", first_error)
            self.summary_label.setText("导出参数不完整，请先补齐后再执行。")
            return
        self.export_requested.emit(request)

    def _request_batch_export(self) -> None:
        request = ChatBatchExportRequest(
            session_names=self.session_names_input.toPlainText().splitlines(),
            target_folder=self.target_folder_input.text().strip(),
            export_messages=self.export_messages_checkbox.isChecked(),
            export_files=self.export_files_checkbox.isChecked(),
            export_images=False,
            message_limit=self.message_limit_spin.value(),
            file_limit=self.file_limit_spin.value(),
        )
        errors = request.validate()
        if errors:
            QMessageBox.warning(self, "批量导出参数不完整", next(iter(errors.values())))
            self.summary_label.setText("批量导出参数不完整，请先补齐后再执行。")
            return
        self.batch_export_requested.emit(request)

    def _choose_folder(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "选择导出目录")
        if path:
            self.target_folder_input.setText(path)

    def _load_example(self) -> None:
        self.session_name_input.setText("项目群")
        self.session_names_input.setPlainText("项目群\n客户群A\n内部通知群")
        self.summary_label.setText("已填入示例会话名称，你可以直接改成自己的群名或好友备注。")

    def _import_session_names(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "导入会话名单", "", "文本或表格 (*.txt *.csv *.xlsx)")
        if not path:
            return
        try:
            names = load_session_names(path)
        except Exception as exc:
            QMessageBox.critical(self, "导入会话名单失败", str(exc))
            return
        if not names:
            QMessageBox.information(self, "导入会话名单", "文件里没有识别到可用的会话名称。")
            return
        self.session_names_input.setPlainText("\n".join(names))
        self.summary_label.setText(f"已导入 {len(names)} 个会话名称。")

    def apply_single_request(self, request: ChatExportRequest) -> None:
        self.session_name_input.setText(request.session_name)
        self.target_folder_input.setText(request.target_folder)
        self.export_messages_checkbox.setChecked(request.export_messages)
        self.export_files_checkbox.setChecked(request.export_files)
        self.message_limit_spin.setValue(request.message_limit)
        self.file_limit_spin.setValue(request.file_limit)
        self.summary_label.setText("已回填上次导出参数。确认无误后可重新执行。")

    def apply_batch_request(self, request: ChatBatchExportRequest) -> None:
        self.session_names_input.setPlainText("\n".join(request.session_names))
        self.target_folder_input.setText(request.target_folder)
        self.export_messages_checkbox.setChecked(request.export_messages)
        self.export_files_checkbox.setChecked(request.export_files)
        self.message_limit_spin.setValue(request.message_limit)
        self.file_limit_spin.setValue(request.file_limit)
        self.summary_label.setText("已回填批量导出参数。确认无误后可重新执行。")

    def set_running_state(self, is_running: bool) -> None:
        for name, button in self._buttons.items():
            if name == "停止":
                button.setEnabled(is_running)
            else:
                button.setEnabled(not is_running)
        self.choose_folder_button.setEnabled(not is_running)
        for widget in [
            self.session_name_input,
            self.session_names_input,
            self.target_folder_input,
            self.export_messages_checkbox,
            self.export_files_checkbox,
            self.message_limit_spin,
            self.file_limit_spin,
        ]:
            widget.setEnabled(not is_running)


class ResourceToolsPage(QWidget):
    export_requested = Signal(object)
    stop_requested = Signal()

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self._buttons: dict[str, QPushButton] = {}
        layout = QVBoxLayout(self)
        card = CardFrame("资源导出工具", hero=True)
        subtitle = QLabel("把微信已经保存在本地的聊天文件、最近文件、聊天视频快速归档出来。")
        subtitle.setProperty("role", "pageSubtitle")
        subtitle.setWordWrap(True)
        card.body_layout.addWidget(subtitle)
        helper = QLabel("这里整理的是已经保存在本机的内容。若微信里尚未下载到本地，导出结果可能为空。")
        helper.setProperty("role", "warn")
        helper.setWordWrap(True)
        card.body_layout.addWidget(helper)

        form = QFormLayout()
        _configure_form_layout(form)
        self.kind_combo = QComboBox()
        self.kind_combo.addItem("导出最近聊天文件", ResourceExportKind.RECENT_FILES)
        self.kind_combo.addItem("按年月导出微信聊天文件", ResourceExportKind.WXFILES)
        self.kind_combo.addItem("按年月导出微信聊天视频", ResourceExportKind.VIDEOS)
        self.target_folder_input = QLineEdit()
        self.target_folder_input.setPlaceholderText("选择导出目录")
        self.choose_folder_button = QPushButton("选择文件夹")
        self.choose_folder_button.setProperty("variant", "ghost")
        self.choose_folder_button.setMinimumWidth(104)
        target_row = QHBoxLayout()
        target_row.addWidget(self.target_folder_input)
        target_row.addWidget(self.choose_folder_button)
        self.year_spin = QSpinBox()
        self.year_spin.setRange(2020, 2100)
        self.year_spin.setValue(2026)
        _set_compact_width(self.year_spin, 120)
        self.month_spin = QSpinBox()
        self.month_spin.setRange(0, 12)
        self.month_spin.setSpecialValueText("全部月份")
        self.month_spin.setValue(0)
        _set_compact_width(self.month_spin, 120)
        year_label = QLabel("年份")
        year_label.setProperty("role", "muted")
        month_label = QLabel("月份")
        month_label.setProperty("role", "muted")
        time_row = _flow_container(year_label, self.year_spin, month_label, self.month_spin, h_spacing=12)
        form.addRow("导出类型", self.kind_combo)
        form.addRow("导出目录", target_row)
        form.addRow("时间范围", time_row)
        card.body_layout.addLayout(form)

        toolbar_buttons: list[QPushButton] = []
        for text, callback in [("示例填充", self._load_example), ("开始导出", self._request_export), ("停止", self.stop_requested.emit)]:
            button = QPushButton(text)
            if text == "示例填充":
                button.setProperty("variant", "ghost")
            if text == "停止":
                button.setProperty("variant", "secondary")
            self._buttons[text] = button
            button.clicked.connect(lambda _checked=False, cb=callback: cb())
            toolbar_buttons.append(button)
        card.body_layout.addWidget(_flow_container(*toolbar_buttons))

        self.summary_label = QLabel("可用于做资料归档、交接备份和排查本地已下载的素材。")
        self.summary_label.setProperty("role", "muted")
        self.summary_label.setWordWrap(True)
        card.body_layout.addWidget(self.summary_label)
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setPlaceholderText("导出完成后，这里会显示导出数量、目录和摘要文件。")
        card.body_layout.addWidget(self.result_text)
        self.open_folder_button = QPushButton("打开导出目录")
        self.open_folder_button.setProperty("variant", "secondary")
        self.open_folder_button.setEnabled(False)
        card.body_layout.addWidget(self.open_folder_button)
        layout.addWidget(card)
        layout.addStretch(1)

        self.choose_folder_button.clicked.connect(self._choose_folder)
        self.kind_combo.currentIndexChanged.connect(self._sync_form_state)
        self._sync_form_state()
        self.set_running_state(False)

    def current_request(self) -> tuple[ResourceExportRequest | None, dict[str, str]]:
        request = ResourceExportRequest(
            export_kind=self.kind_combo.currentData(),
            target_folder=self.target_folder_input.text().strip(),
            year=str(self.year_spin.value()),
            month="" if self.month_spin.value() == 0 else f"{self.month_spin.value():02d}",
        )
        errors = request.validate()
        return (None, errors) if errors else (request, {})

    def _request_export(self) -> None:
        request, errors = self.current_request()
        if errors:
            QMessageBox.warning(self, "导出参数不完整", next(iter(errors.values())))
            self.summary_label.setText("导出参数不完整，请先补齐后再执行。")
            return
        self.export_requested.emit(request)

    def _choose_folder(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "选择导出目录")
        if path:
            self.target_folder_input.setText(path)

    def _load_example(self) -> None:
        self.summary_label.setText("示例已填好。你可以直接改年份和导出目录。")

    def _sync_form_state(self) -> None:
        kind = self.kind_combo.currentData()
        show_year_month = kind in {ResourceExportKind.WXFILES, ResourceExportKind.VIDEOS}
        self.year_spin.setEnabled(show_year_month)
        self.month_spin.setEnabled(show_year_month)

    def apply_request(self, request: ResourceExportRequest) -> None:
        index = self.kind_combo.findData(request.export_kind)
        if index >= 0:
            self.kind_combo.setCurrentIndex(index)
        self.target_folder_input.setText(request.target_folder)
        self.year_spin.setValue(int(request.year))
        self.month_spin.setValue(int(request.month) if request.month else 0)
        self.summary_label.setText("已回填上次资源导出参数。确认无误后可重新执行。")

    def set_running_state(self, is_running: bool) -> None:
        for name, button in self._buttons.items():
            if name == "停止":
                button.setEnabled(is_running)
            else:
                button.setEnabled(not is_running)
        self.choose_folder_button.setEnabled(not is_running)
        self.kind_combo.setEnabled(not is_running)
        self.target_folder_input.setEnabled(not is_running)
        self.year_spin.setEnabled(not is_running and self.kind_combo.currentData() in {ResourceExportKind.WXFILES, ResourceExportKind.VIDEOS})
        self.month_spin.setEnabled(not is_running and self.kind_combo.currentData() in {ResourceExportKind.WXFILES, ResourceExportKind.VIDEOS})


class SessionToolsPage(QWidget):
    scan_sessions_requested = Signal(object)
    scan_groups_requested = Signal()
    use_session_names_requested = Signal(object)

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self._buttons: dict[str, QPushButton] = {}
        layout = QVBoxLayout(self)

        session_card = CardFrame("会话与群工具", hero=True)
        intro = QLabel("把微信里已经存在的会话和群聊名单直接采集出来，减少手工抄名字的出错率。")
        intro.setProperty("role", "pageSubtitle")
        intro.setWordWrap(True)
        session_card.body_layout.addWidget(intro)
        helper = QLabel("适合做批量导出前的名单整理和交接核对。执行时同样会占用微信界面。")
        helper.setProperty("role", "hint")
        helper.setWordWrap(True)
        session_card.body_layout.addWidget(helper)

        self.chatted_only_checkbox = QCheckBox("只采集聊过天的会话")
        session_card.body_layout.addWidget(_flow_container(self.chatted_only_checkbox))

        self.scan_sessions_button = QPushButton("采集会话列表")
        self.export_sessions_button = QPushButton("导出会话名单")
        self.export_sessions_button.setProperty("variant", "secondary")
        self.use_sessions_button = QPushButton("填入会话导出页")
        self.use_sessions_button.setProperty("variant", "ghost")
        session_card.body_layout.addWidget(
            _flow_container(self.scan_sessions_button, self.export_sessions_button, self.use_sessions_button)
        )

        self.session_summary_label = QLabel("先采集一次会话列表，再把名单填入“会话导出”页面或导出成表格。")
        self.session_summary_label.setProperty("role", "muted")
        self.session_summary_label.setWordWrap(True)
        session_card.body_layout.addWidget(self.session_summary_label)

        self.session_table = QTableWidget(0, 3)
        self.session_table.setHorizontalHeaderLabels(["会话名称", "最近时间", "最后一条消息"])
        _configure_data_table(self.session_table, minimum_height=280)
        self.session_table.verticalHeader().setVisible(False)
        self.session_table.horizontalHeader().setStretchLastSection(True)
        self.session_table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        session_card.body_layout.addWidget(self.session_table)
        layout.addWidget(session_card)

        group_card = CardFrame("群聊与成员")
        self.scan_groups_button = QPushButton("采集群聊列表")
        self.export_groups_button = QPushButton("导出群聊名单")
        self.export_groups_button.setProperty("variant", "secondary")
        self.use_groups_button = QPushButton("群聊填入会话导出页")
        self.use_groups_button.setProperty("variant", "ghost")
        group_card.body_layout.addWidget(
            _flow_container(self.scan_groups_button, self.export_groups_button, self.use_groups_button)
        )

        self.group_summary_label = QLabel("如果你记不清准确群名，可以先采集群聊列表，再直接回填到批量会话导出。")
        self.group_summary_label.setProperty("role", "muted")
        self.group_summary_label.setWordWrap(True)
        group_card.body_layout.addWidget(self.group_summary_label)

        self.group_table = QTableWidget(0, 1)
        self.group_table.setHorizontalHeaderLabels(["群聊名称"])
        _configure_data_table(self.group_table, minimum_height=220)
        self.group_table.verticalHeader().setVisible(False)
        self.group_table.horizontalHeader().setStretchLastSection(True)
        self.group_table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        group_card.body_layout.addWidget(self.group_table)
        layout.addWidget(group_card)
        layout.addStretch(1)

        self.scan_sessions_button.clicked.connect(self._request_scan_sessions)
        self.scan_groups_button.clicked.connect(lambda: self.scan_groups_requested.emit())
        self.export_sessions_button.clicked.connect(self._export_session_rows)
        self.export_groups_button.clicked.connect(self._export_group_rows)
        self.use_sessions_button.clicked.connect(lambda: self._emit_selected_names(self.session_table, 0))
        self.use_groups_button.clicked.connect(lambda: self._emit_selected_names(self.group_table, 0))

        self.export_sessions_button.setEnabled(False)
        self.export_groups_button.setEnabled(False)
        self.use_sessions_button.setEnabled(False)
        self.use_groups_button.setEnabled(False)
        self.set_running_state(False)

    def _request_scan_sessions(self) -> None:
        self.scan_sessions_requested.emit(SessionScanRequest(chatted_only=self.chatted_only_checkbox.isChecked()))

    def set_session_rows(self, rows: list[SessionSummaryRow]) -> None:
        self.session_table.setRowCount(0)
        for row_index, row in enumerate(rows):
            self.session_table.insertRow(row_index)
            self.session_table.setItem(row_index, 0, QTableWidgetItem(row.session_name))
            self.session_table.setItem(row_index, 1, QTableWidgetItem(row.last_time))
            self.session_table.setItem(row_index, 2, QTableWidgetItem(row.last_message))
        self.session_table.resizeColumnsToContents()
        self.export_sessions_button.setEnabled(bool(rows))
        self.use_sessions_button.setEnabled(bool(rows))
        self.session_summary_label.setText(f"已采集 {len(rows)} 个会话，可直接导出名单或回填到会话导出页。")

    def set_group_rows(self, rows: list[GroupSummaryRow]) -> None:
        self.group_table.setRowCount(0)
        for row_index, row in enumerate(rows):
            self.group_table.insertRow(row_index)
            self.group_table.setItem(row_index, 0, QTableWidgetItem(row.group_name))
        self.group_table.resizeColumnsToContents()
        self.export_groups_button.setEnabled(bool(rows))
        self.use_groups_button.setEnabled(bool(rows))
        self.group_summary_label.setText(f"已采集 {len(rows)} 个群聊，可直接导出群聊名单或回填到会话导出页。")

    def set_running_state(self, is_running: bool) -> None:
        for widget in [
            self.chatted_only_checkbox,
            self.scan_sessions_button,
            self.scan_groups_button,
        ]:
            widget.setEnabled(not is_running)
        if is_running:
            self.export_sessions_button.setEnabled(False)
            self.export_groups_button.setEnabled(False)
            self.use_sessions_button.setEnabled(False)
            self.use_groups_button.setEnabled(False)
        else:
            self.export_sessions_button.setEnabled(self.session_table.rowCount() > 0)
            self.export_groups_button.setEnabled(self.group_table.rowCount() > 0)
            self.use_sessions_button.setEnabled(self.session_table.rowCount() > 0)
            self.use_groups_button.setEnabled(self.group_table.rowCount() > 0)

    def selected_session_names(self) -> list[str]:
        return self._selected_names(self.session_table, 0)

    def selected_group_names(self) -> list[str]:
        return self._selected_names(self.group_table, 0)

    def _selected_names(self, table: QTableWidget, column_index: int) -> list[str]:
        rows = sorted({item.row() for item in table.selectedItems()})
        if not rows:
            rows = list(range(table.rowCount()))
        names: list[str] = []
        for row in rows:
            item = table.item(row, column_index)
            text = item.text().strip() if item else ""
            if text:
                names.append(text)
        return names

    def _emit_selected_names(self, table: QTableWidget, column_index: int) -> None:
        names = self._selected_names(table, column_index)
        if not names:
            QMessageBox.information(self, "会话与群工具", "当前没有可回填的名称。")
            return
        self.use_session_names_requested.emit(names)

    def _export_session_rows(self) -> None:
        rows = [
            {
                "会话名称": self.session_table.item(row, 0).text() if self.session_table.item(row, 0) else "",
                "最近时间": self.session_table.item(row, 1).text() if self.session_table.item(row, 1) else "",
                "最后一条消息": self.session_table.item(row, 2).text() if self.session_table.item(row, 2) else "",
            }
            for row in range(self.session_table.rowCount())
        ]
        self._export_generic_rows("导出会话名单", ["会话名称", "最近时间", "最后一条消息"], rows)

    def _export_group_rows(self) -> None:
        rows = [
            {
                "群聊名称": self.group_table.item(row, 0).text() if self.group_table.item(row, 0) else "",
            }
            for row in range(self.group_table.rowCount())
        ]
        self._export_generic_rows("导出群聊名单", ["群聊名称"], rows)

    def _export_generic_rows(self, title: str, headers: list[str], rows: list[dict[str, Any]]) -> None:
        if not rows:
            QMessageBox.information(self, title, "当前没有可导出的数据。")
            return
        path, _ = QFileDialog.getSaveFileName(self, title, "", "Excel (*.xlsx);;CSV (*.csv)")
        if not path:
            return
        dump_table(headers, rows, path)


class RelayWorkbenchPage(QWidget):
    collect_texts_requested = Signal(object)
    collect_files_requested = Signal(object)
    import_export_folder_requested = Signal(str)
    validate_routes_requested = Signal(object)
    test_send_requested = Signal(object)
    send_requested = Signal(object)

    PACKAGE_COLUMNS = [
        ("sequence", "顺序", "int"),
        ("item_type", "类型", "text"),
        ("content", "内容预览", "text"),
        ("file_path", "本地路径", "text"),
        ("collected_at", "采集时间", "text"),
    ]

    ROUTE_COLUMNS = [
        ("downstream_session", "下游会话", "text"),
        ("validation_status", "验证状态", "text"),
        ("validation_message", "说明", "text"),
    ]

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self._running = False
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        source_card = CardFrame("转发工作台", hero=True)
        intro = QLabel("先从上游采集候选内容，再人工勾选、调整顺序、验证下游，最后测试发送或正式批量发送。")
        intro.setProperty("role", "pageSubtitle")
        intro.setWordWrap(True)
        source_card.body_layout.addWidget(intro)
        helper = QLabel("先整理这次真正要转发的内容，再先发一遍测试消息确认无误，最后正式批量发送。")
        helper.setProperty("role", "hint")
        helper.setWordWrap(True)
        source_card.body_layout.addWidget(helper)

        form = QFormLayout()
        _configure_form_layout(form)
        self.source_session_input = QLineEdit()
        self.source_session_input.setPlaceholderText("填写上游 A 的会话名称")
        self.package_name_input = QLineEdit()
        self.package_name_input.setPlaceholderText("可选，例如：A客户-0422转发包")
        self.message_limit_spin = QSpinBox()
        self.message_limit_spin.setRange(1, 500)
        self.message_limit_spin.setValue(20)
        _set_compact_width(self.message_limit_spin, 120)
        self.file_limit_spin = QSpinBox()
        self.file_limit_spin.setRange(1, 200)
        self.file_limit_spin.setValue(10)
        _set_compact_width(self.file_limit_spin, 120)
        message_limit_label = QLabel("文本采集条数")
        message_limit_label.setProperty("role", "muted")
        file_limit_label = QLabel("文件采集数量")
        file_limit_label.setProperty("role", "muted")
        count_row = _flow_container(message_limit_label, self.message_limit_spin, file_limit_label, self.file_limit_spin, h_spacing=12)
        form.addRow("上游会话", self.source_session_input)
        form.addRow("转发包名称", self.package_name_input)
        form.addRow("采集范围", count_row)
        source_card.body_layout.addLayout(form)

        self.collect_texts_button = QPushButton("采集文本")
        self.collect_files_button = QPushButton("采集文件")
        self.import_export_folder_button = QPushButton("从导出目录导入")
        self.add_text_button = QPushButton("手动加文本")
        self.add_files_button = QPushButton("选择文件/图片")
        self.clear_package_button = QPushButton("清空转发包")
        self.clear_package_button.setProperty("variant", "secondary")
        source_card.body_layout.addWidget(
            _flow_container(
                self.collect_texts_button,
                self.collect_files_button,
                self.import_export_folder_button,
                self.add_text_button,
                self.add_files_button,
                self.clear_package_button,
            )
        )
        self.package_summary_label = QLabel("先整理真正要发送的内容。文件可通过按钮选择，也可以直接拖入下方表格。")
        self.package_summary_label.setProperty("role", "muted")
        self.package_summary_label.setWordWrap(True)
        source_card.body_layout.addWidget(self.package_summary_label)
        layout.addWidget(source_card)

        package_card = CardFrame("转发包编辑器")
        self.keep_latest_files_button = QPushButton("只保留最新文件")
        self.keep_latest_files_button.setProperty("variant", "ghost")
        self.move_up_button = QPushButton("上移")
        self.move_down_button = QPushButton("下移")
        self.remove_package_button = QPushButton("删除选中")
        self.remove_package_button.setProperty("variant", "secondary")
        package_card.body_layout.addWidget(
            _flow_container(self.keep_latest_files_button, self.move_up_button, self.move_down_button, self.remove_package_button)
        )
        tip = QLabel("每一行就是一条要发送的内容。不要发的内容直接删除，需要调整顺序时可用上移和下移。")
        tip.setProperty("role", "hint")
        tip.setWordWrap(True)
        package_card.body_layout.addWidget(tip)
        self.package_table = RelayPackageTableWidget()
        self.package_table.setColumnCount(len(self.PACKAGE_COLUMNS))
        self.package_table.setHorizontalHeaderLabels([title for _, title, _ in self.PACKAGE_COLUMNS])
        _configure_data_table(self.package_table, minimum_height=220)
        self.package_table.verticalHeader().setVisible(False)
        self.package_table.horizontalHeader().setStretchLastSection(True)
        self.package_table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        package_card.body_layout.addWidget(self.package_table)
        layout.addWidget(package_card)

        route_card = CardFrame("下游会话")
        self.import_routes_button = QPushButton("导入下游名单")
        self.export_route_template_button = QPushButton("导出下游模板")
        self.export_route_template_button.setProperty("variant", "ghost")
        self.add_route_button = QPushButton("新增一行")
        self.remove_route_button = QPushButton("删除选中")
        self.remove_route_button.setProperty("variant", "secondary")
        route_card.body_layout.addWidget(
            _flow_container(self.import_routes_button, self.export_route_template_button, self.add_route_button, self.remove_route_button)
        )
        route_tip = QLabel("下游名单支持 Excel/CSV 导入。每行写一个下游会话，也支持在“下游会话列表”列里用 | 一次写多个下游。")
        route_tip.setProperty("role", "hint")
        route_tip.setWordWrap(True)
        route_card.body_layout.addWidget(route_tip)
        self.route_summary_label = QLabel("还没有填写下游会话。")
        self.route_summary_label.setProperty("role", "muted")
        self.route_summary_label.setWordWrap(True)
        route_card.body_layout.addWidget(self.route_summary_label)
        self.route_table = QTableWidget(0, len(self.ROUTE_COLUMNS))
        self.route_table.setHorizontalHeaderLabels([title for _, title, _ in self.ROUTE_COLUMNS])
        _configure_data_table(self.route_table, minimum_height=200)
        self.route_table.verticalHeader().setVisible(False)
        self.route_table.horizontalHeader().setStretchLastSection(True)
        self.route_table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        route_card.body_layout.addWidget(self.route_table)
        self.validate_routes_button = QPushButton("验证下游会话")
        self.test_send_button = QPushButton("测试发送到文件传输助手")
        self.test_send_button.setProperty("variant", "secondary")
        self.send_button = QPushButton("正式批量发送")
        route_card.body_layout.addWidget(
            _flow_container(self.validate_routes_button, self.test_send_button, self.send_button)
        )
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setMinimumHeight(96)
        self.result_text.setMaximumHeight(140)
        self.result_text.setPlaceholderText("验证结果、测试发送结果和正式发送结果会显示在这里。")
        route_card.body_layout.addWidget(self.result_text)
        layout.addWidget(route_card)
        layout.addStretch(1)

        self.collect_texts_button.clicked.connect(self._request_collect_texts)
        self.collect_files_button.clicked.connect(self._request_collect_files)
        self.import_export_folder_button.clicked.connect(self._request_import_export_folder)
        self.add_text_button.clicked.connect(self._add_manual_text)
        self.add_files_button.clicked.connect(self._add_local_files)
        self.clear_package_button.clicked.connect(self.clear_package_rows)
        self.keep_latest_files_button.clicked.connect(self.keep_latest_file_rows)
        self.move_up_button.clicked.connect(self.move_selected_package_rows_up)
        self.move_down_button.clicked.connect(self.move_selected_package_rows_down)
        self.remove_package_button.clicked.connect(self.remove_selected_package_rows)
        self.package_table.files_dropped.connect(self._handle_dropped_files)
        self.package_table.itemChanged.connect(lambda _item: self.refresh_package_summary())
        self.import_routes_button.clicked.connect(self.import_route_rows)
        self.export_route_template_button.clicked.connect(self.export_route_template)
        self.add_route_button.clicked.connect(self.add_empty_route_row)
        self.remove_route_button.clicked.connect(self.remove_selected_route_rows)
        self.route_table.itemChanged.connect(lambda _item: self.refresh_route_summary())
        self.validate_routes_button.clicked.connect(self._request_validate_routes)
        self.test_send_button.clicked.connect(self._request_test_send)
        self.send_button.clicked.connect(self._request_send)
        self.source_session_input.textChanged.connect(self.refresh_route_summary)

        self._configure_package_columns()
        self._configure_route_columns()
        self.set_running_state(False)

    def _configure_package_columns(self) -> None:
        widths = {"sequence": 70, "item_type": 90, "content": 320, "file_path": 360, "collected_at": 150}
        for index, (key, _, _) in enumerate(self.PACKAGE_COLUMNS):
            if key in widths:
                self.package_table.setColumnWidth(index, widths[key])

    def _configure_route_columns(self) -> None:
        widths = {"downstream_session": 220, "validation_status": 100, "validation_message": 320}
        for index, (key, _, _) in enumerate(self.ROUTE_COLUMNS):
            if key in widths:
                self.route_table.setColumnWidth(index, widths[key])

    def _request_collect_texts(self) -> None:
        request = RelayCollectTextRequest(
            source_session=self.source_session_input.text().strip(),
            message_limit=self.message_limit_spin.value(),
        )
        errors = request.validate()
        if errors:
            QMessageBox.warning(self, "采集文本", next(iter(errors.values())))
            return
        self.collect_texts_requested.emit(request)

    def _request_collect_files(self) -> None:
        request = RelayCollectFilesRequest(
            source_session=self.source_session_input.text().strip(),
            file_limit=self.file_limit_spin.value(),
        )
        errors = request.validate()
        if errors:
            QMessageBox.warning(self, "采集文件", next(iter(errors.values())))
            return
        self.collect_files_requested.emit(request)

    def _add_manual_text(self) -> None:
        text, ok = QInputDialog.getMultiLineText(self, "手动添加文本", "要加入转发包的文本内容")
        if not ok or not text.strip():
            return
        row = RelayPackageRow(
            sequence=self.package_table.rowCount() + 1,
            item_type=RelayItemType.TEXT,
            source_session=self.source_session_input.text().strip(),
            content=text.strip(),
            collected_at="手动添加",
        )
        self.append_package_rows([row], "已手动加入 1 条文本。")

    def _add_local_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(self, "选择本地文件或图片")
        if not paths:
            return
        self._append_local_files(paths, f"已加入 {len(paths)} 个本地文件/图片。")

    def _handle_dropped_files(self, paths: list[str]) -> None:
        valid_paths = [path for path in paths if Path(path).is_file()]
        if not valid_paths:
            return
        self._append_local_files(valid_paths, f"已拖入 {len(valid_paths)} 个本地文件/图片。")

    def _append_local_files(self, paths: list[str], summary: str) -> None:
        rows = []
        for index, path in enumerate(paths, start=1):
            item_type = RelayItemType.IMAGE if Path(path).suffix.lower() in {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".tiff"} else RelayItemType.FILE
            rows.append(
                RelayPackageRow(
                    sequence=self.package_table.rowCount() + index,
                    item_type=item_type,
                    source_session=self.source_session_input.text().strip(),
                    content=Path(path).name,
                    file_path=path,
                    collected_at="本地补充",
                )
            )
        self.append_package_rows(rows, summary)

    def _request_import_export_folder(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "选择会话导出目录")
        if not path:
            return
        self.import_export_folder_requested.emit(path)

    def append_package_rows(self, rows: list[RelayPackageRow], summary: str = "") -> None:
        start = self.package_table.rowCount()
        for offset, row in enumerate(rows):
            self.package_table.insertRow(start + offset)
            self._set_table_row(self.package_table, start + offset, self.PACKAGE_COLUMNS, row.__dict__ | {"item_type": row.item_type.value})
        self._renumber_package_rows()
        if summary:
            self.package_summary_label.setText(summary)

    def set_route_rows(self, rows: list[RelayRouteRow], summary: str | None = None) -> None:
        self.route_table.setRowCount(0)
        for row_index, row in enumerate(rows):
            self.route_table.insertRow(row_index)
            self._set_table_row(self.route_table, row_index, self.ROUTE_COLUMNS, row.__dict__)
        self.refresh_route_summary(summary)

    def route_rows(self) -> list[RelayRouteRow]:
        rows: list[RelayRouteRow] = []
        for row_index in range(self.route_table.rowCount()):
            mapping = self._row_mapping(self.route_table, row_index, self.ROUTE_COLUMNS)
            row = RelayRouteRow.from_mapping(mapping)
            if row.downstream_session.strip():
                rows.append(row)
        return rows

    def package_rows(self) -> list[RelayPackageRow]:
        rows: list[RelayPackageRow] = []
        for row_index in range(self.package_table.rowCount()):
            mapping = self._row_mapping(self.package_table, row_index, self.PACKAGE_COLUMNS)
            rows.append(RelayPackageRow.from_mapping(mapping))
        return rows

    def selected_package_rows(self) -> list[RelayPackageRow]:
        return self.package_rows()

    def current_validation_request(self) -> RelayValidationRequest:
        return RelayValidationRequest(
            source_session=self.source_session_input.text().strip(),
            route_rows=self.route_rows(),
        )

    def current_send_request(self, test_only: bool) -> RelaySendRequest:
        return RelaySendRequest(
            source_session=self.source_session_input.text().strip(),
            package_rows=self.package_rows(),
            route_rows=self.route_rows(),
            test_only=test_only,
        )

    def _request_validate_routes(self) -> None:
        request = self.current_validation_request()
        errors = request.validate()
        if errors:
            QMessageBox.warning(self, "验证下游", next(iter(errors.values())))
            return
        self.validate_routes_requested.emit(request)

    def _request_test_send(self) -> None:
        request = self.current_send_request(test_only=True)
        errors = request.validate()
        if errors:
            QMessageBox.warning(self, "测试发送", next(iter(errors.values())))
            return
        self.test_send_requested.emit(request)

    def _request_send(self) -> None:
        request = self.current_send_request(test_only=False)
        errors = request.validate()
        if errors:
            QMessageBox.warning(self, "正式发送", next(iter(errors.values())))
            return
        self.send_requested.emit(request)

    def import_route_rows(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "导入下游名单", "", "Excel/CSV (*.xlsx *.csv)")
        if not path:
            return
        rows = load_route_rows(path)
        self.set_route_rows(rows, f"已导入 {len(rows)} 个下游会话。")

    def export_route_template(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "导出下游模板", "relay-targets-template.xlsx", "Excel (*.xlsx);;CSV (*.csv)")
        if not path:
            return
        dump_route_rows([], path)

    def add_empty_route_row(self) -> None:
        row_index = self.route_table.rowCount()
        self.route_table.insertRow(row_index)
        self._set_table_row(
            self.route_table,
            row_index,
            self.ROUTE_COLUMNS,
            {
                "downstream_session": "",
                "validation_status": "未验证",
                "validation_message": "",
            },
        )
        self.refresh_route_summary()

    def remove_selected_package_rows(self) -> None:
        self._remove_selected_rows(self.package_table)
        self._renumber_package_rows()
        self.refresh_package_summary()

    def remove_selected_route_rows(self) -> None:
        self._remove_selected_rows(self.route_table)
        self.refresh_route_summary()

    def move_selected_package_rows_up(self) -> None:
        self._move_selected_rows(self.package_table, self.PACKAGE_COLUMNS, direction=-1)
        self._renumber_package_rows()

    def move_selected_package_rows_down(self) -> None:
        self._move_selected_rows(self.package_table, self.PACKAGE_COLUMNS, direction=1)
        self._renumber_package_rows()

    def clear_package_rows(self) -> None:
        self.package_table.setRowCount(0)
        self.refresh_package_summary("转发包已清空。")

    def replace_package_rows(self, rows: list[RelayPackageRow], summary: str = "") -> None:
        self.package_table.setRowCount(0)
        self.append_package_rows(rows, summary)

    def keep_latest_file_rows(self) -> None:
        try:
            from ..relay_service import RelayService
        except Exception:
            return
        rows = self.package_rows()
        if not rows:
            QMessageBox.information(self, "转发包编辑器", "当前没有可处理的文件项。")
            return
        updated_rows = RelayService.keep_latest_file_rows(rows)
        removed_count = max(0, len(rows) - len(updated_rows))
        self.replace_package_rows(updated_rows, f"已移除 {removed_count} 个同名旧文件，只保留最新版本。")

    def apply_validation_result(self, rows: list[RelayRouteRow], summary: str) -> None:
        self.set_route_rows(rows, summary)
        self.result_text.setPlainText(summary)

    def apply_send_result(self, text: str) -> None:
        self.result_text.setPlainText(text)

    def refresh_package_summary(self, override: str | None = None) -> None:
        if override:
            self.package_summary_label.setText(override)
            return
        rows = self.package_rows()
        self.package_summary_label.setText(f"当前转发包共有 {len(rows)} 条内容。")

    def refresh_route_summary(self, override: str | None = None) -> None:
        if override:
            self.route_summary_label.setText(override)
            return
        rows = self.route_rows()
        if not rows:
            self.route_summary_label.setText("还没有填写下游会话。")
            return
        self.route_summary_label.setText(f"当前已填写 {len(rows)} 个下游会话。")

    def set_running_state(self, is_running: bool) -> None:
        self._running = is_running
        for widget in [
            self.source_session_input,
            self.package_name_input,
            self.message_limit_spin,
            self.file_limit_spin,
            self.collect_texts_button,
            self.collect_files_button,
            self.import_export_folder_button,
            self.add_text_button,
            self.add_files_button,
            self.clear_package_button,
            self.keep_latest_files_button,
            self.move_up_button,
            self.move_down_button,
            self.remove_package_button,
            self.import_routes_button,
            self.export_route_template_button,
            self.add_route_button,
            self.remove_route_button,
            self.validate_routes_button,
            self.test_send_button,
            self.send_button,
        ]:
            widget.setEnabled(not is_running)

    def _set_table_row(self, table: QTableWidget, row_index: int, columns: list[tuple[str, str, str]], values: dict[str, Any]) -> None:
        for column_index, (key, _, kind) in enumerate(columns):
            item = QTableWidgetItem()
            value = values.get(key, "")
            if kind == "bool":
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
                item.setCheckState(Qt.CheckState.Checked if bool(value) else Qt.CheckState.Unchecked)
                item.setText("")
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            elif key == "item_type":
                item_value = value.value if isinstance(value, RelayItemType) else str(value or RelayItemType.TEXT.value)
                item.setData(Qt.ItemDataRole.UserRole, item_value)
                item.setText(RELAY_ITEM_LABELS.get(item_value, item_value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            else:
                item.setText("" if value is None else str(value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            table.setItem(row_index, column_index, item)

    def _row_mapping(self, table: QTableWidget, row_index: int, columns: list[tuple[str, str, str]]) -> dict[str, Any]:
        mapping: dict[str, Any] = {}
        for column_index, (key, _, kind) in enumerate(columns):
            item = table.item(row_index, column_index)
            if kind == "bool":
                mapping[key] = item.checkState() == Qt.CheckState.Checked if item else False
            elif key == "item_type":
                raw_value = item.data(Qt.ItemDataRole.UserRole) if item else ""
                mapping[key] = str(raw_value or RelayItemType.TEXT.value)
            else:
                mapping[key] = item.text().strip() if item else ""
        return mapping

    @staticmethod
    def _remove_selected_rows(table: QTableWidget) -> None:
        rows = sorted({item.row() for item in table.selectedItems()}, reverse=True)
        for row_index in rows:
            table.removeRow(row_index)

    def _renumber_package_rows(self) -> None:
        sequence_column = next(index for index, (key, _, _) in enumerate(self.PACKAGE_COLUMNS) if key == "sequence")
        for row_index in range(self.package_table.rowCount()):
            sequence_item = self.package_table.item(row_index, sequence_column)
            if sequence_item is not None:
                sequence_item.setText(str(row_index + 1))
        self.refresh_package_summary()

    def _move_selected_rows(self, table: QTableWidget, columns: list[tuple[str, str, str]], direction: int) -> None:
        rows = sorted({item.row() for item in table.selectedItems()})
        if not rows:
            return
        if direction > 0:
            rows = list(reversed(rows))
        for row_index in rows:
            new_index = row_index + direction
            if new_index < 0 or new_index >= table.rowCount():
                continue
            current = self._row_mapping(table, row_index, columns)
            target = self._row_mapping(table, new_index, columns)
            self._set_table_row(table, row_index, columns, target)
            self._set_table_row(table, new_index, columns, current)
            table.selectRow(new_index)


class TemplatesPage(QWidget):
    load_template_requested = Signal(object, object)
    create_message_requested = Signal()
    create_file_requested = Signal()

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        card = CardFrame("模板中心")
        self.refresh_button = QPushButton("刷新")
        self.rename_button = QPushButton("重命名")
        self.rename_button.setProperty("variant", "secondary")
        self.duplicate_button = QPushButton("复制")
        self.duplicate_button.setProperty("variant", "secondary")
        self.delete_button = QPushButton("删除")
        self.delete_button.setProperty("variant", "secondary")
        self.restore_button = QPushButton("恢复刚删除")
        self.restore_button.setProperty("variant", "ghost")
        self.load_button = QPushButton("加载到工作台")
        card.body_layout.addWidget(
            _flow_container(
                self.refresh_button,
                self.rename_button,
                self.duplicate_button,
                self.delete_button,
                self.restore_button,
                self.load_button,
            )
        )
        helper = QLabel("把常用任务保存成模板，下次可以直接加载，适合固定通知、固定资料发送。")
        helper.setProperty("role", "hint")
        helper.setWordWrap(True)
        card.body_layout.addWidget(helper)
        self.total_templates_label = self._create_stat_card("模板总数", "0")
        self.message_templates_label = self._create_stat_card("消息模板", "0")
        self.file_templates_label = self._create_stat_card("文件模板", "0")
        card.body_layout.addWidget(
            _flow_container(self.total_templates_label, self.message_templates_label, self.file_templates_label, h_spacing=12, v_spacing=12)
        )
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("搜索模板名称或类型，例如：消息、客户、通知")
        card.body_layout.addWidget(self.search_input)
        self.summary_label = QLabel("还没有模板。你可以先去批量页面整理一份任务，再点“保存模板”。")
        self.summary_label.setProperty("role", "muted")
        self.summary_label.setWordWrap(True)
        card.body_layout.addWidget(self.summary_label)
        self.create_message_button = QPushButton("新建消息任务")
        self.create_file_button = QPushButton("新建文件任务")
        self.create_file_button.setProperty("variant", "secondary")
        card.body_layout.addWidget(_flow_container(self.create_message_button, self.create_file_button))
        self.create_message_button.clicked.connect(lambda: self.create_message_requested.emit())
        self.create_file_button.clicked.connect(lambda: self.create_file_requested.emit())
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["ID", "名称", "类型", "更新时间"])
        _configure_data_table(self.table, minimum_height=280)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        card.body_layout.addWidget(self.table)
        layout.addWidget(card)
        layout.addStretch(1)

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
        card.setMinimumWidth(160)
        return card


class ExportHistoryPage(QWidget):
    open_folder_requested = Signal()
    open_summary_requested = Signal()
    rerun_requested = Signal()
    retry_failed_requested = Signal()
    clear_history_requested = Signal()

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        card = CardFrame("导出历史")
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("搜索导出类型、标题或目录")
        card.body_layout.addWidget(self.search_input)
        self.open_folder_button = QPushButton("打开目录")
        self.open_folder_button.setProperty("variant", "secondary")
        self.open_summary_button = QPushButton("打开摘要")
        self.open_summary_button.setProperty("variant", "ghost")
        self.rerun_button = QPushButton("重新执行")
        self.rerun_button.setProperty("variant", "secondary")
        self.retry_failed_button = QPushButton("仅重试失败会话")
        self.retry_failed_button.setProperty("variant", "ghost")
        self.clear_button = QPushButton("清理导出历史")
        self.clear_button.setProperty("variant", "secondary")
        card.body_layout.addWidget(
            _flow_container(
                self.open_folder_button,
                self.open_summary_button,
                self.rerun_button,
                self.retry_failed_button,
                self.clear_button,
            )
        )
        self.summary_label = QLabel("最近的会话导出和资源导出都会记录在这里，方便回看和重新打开目录。")
        self.summary_label.setProperty("role", "hint")
        self.summary_label.setWordWrap(True)
        card.body_layout.addWidget(self.summary_label)
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["ID", "类型", "标题", "数量", "时间"])
        _configure_data_table(self.table, minimum_height=240)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        card.body_layout.addWidget(self.table)
        self.detail_text = QTextEdit()
        self.detail_text.setReadOnly(True)
        self.detail_text.setPlaceholderText("选中一条导出历史后，这里会显示目录、摘要路径和更多细节。")
        card.body_layout.addWidget(self.detail_text)
        layout.addWidget(card)
        layout.addStretch(1)

        self.open_folder_button.setEnabled(False)
        self.open_summary_button.setEnabled(False)
        self.rerun_button.setEnabled(False)
        self.retry_failed_button.setEnabled(False)


class HistoryPage(QWidget):
    retry_failed_requested = Signal(object, object, object)
    clear_history_requested = Signal()
    open_message_requested = Signal()
    open_file_requested = Signal()

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        layout = QVBoxLayout(self)

        history_card = CardFrame("执行历史")
        self.refresh_button = QPushButton("刷新")
        self.retry_button = QPushButton("基于失败项重建任务")
        self.export_failed_button = QPushButton("导出失败项")
        self.clear_button = QPushButton("清理历史")
        self.clear_button.setProperty("variant", "secondary")
        self.failed_only_checkbox = QCheckBox("只看含失败记录")
        history_card.body_layout.addWidget(
            _flow_container(self.refresh_button, self.retry_button, self.export_failed_button, self.clear_button, self.failed_only_checkbox)
        )
        helper = QLabel("先看成功/失败数量，再点开失败项。失败项可以一键回填到批量页面重新执行。")
        helper.setProperty("role", "hint")
        helper.setWordWrap(True)
        history_card.body_layout.addWidget(helper)
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("搜索会话、状态或任务类型，例如：失败、客户A、文件")
        history_card.body_layout.addWidget(self.search_input)
        self.total_runs_label = self._create_stat_card("总记录", "0")
        self.success_runs_label = self._create_stat_card("全成功", "0")
        self.failed_runs_label = self._create_stat_card("含失败", "0")
        history_card.body_layout.addWidget(
            _flow_container(self.total_runs_label, self.success_runs_label, self.failed_runs_label, h_spacing=12, v_spacing=12)
        )
        self.summary_label = QLabel("还没有执行历史。第一次成功或失败执行后，这里会显示完整记录。")
        self.summary_label.setProperty("role", "muted")
        self.summary_label.setWordWrap(True)
        history_card.body_layout.addWidget(self.summary_label)
        self.failure_summary_label = QLabel("选中一条执行记录后，这里会展示失败原因摘要。")
        self.failure_summary_label.setProperty("role", "hint")
        self.failure_summary_label.setWordWrap(True)
        history_card.body_layout.addWidget(self.failure_summary_label)
        self.open_message_button = QPushButton("去批量消息")
        self.open_file_button = QPushButton("去批量文件")
        self.open_file_button.setProperty("variant", "secondary")
        history_card.body_layout.addWidget(_flow_container(self.open_message_button, self.open_file_button))
        self.open_message_button.clicked.connect(lambda: self.open_message_requested.emit())
        self.open_file_button.clicked.connect(lambda: self.open_file_requested.emit())

        self.execution_table = QTableWidget(0, 7)
        self.execution_table.setHorizontalHeaderLabels(
            ["ID", "任务类型", "开始时间", "状态", "总行数", "成功", "失败"]
        )
        _configure_data_table(self.execution_table, minimum_height=240)
        self.execution_table.verticalHeader().setVisible(False)
        self.execution_table.horizontalHeader().setStretchLastSection(True)
        self.execution_table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        history_card.body_layout.addWidget(self.execution_table)

        detail_group = QGroupBox("逐行结果")
        detail_layout = QVBoxLayout(detail_group)
        self.detail_table = QTableWidget(0, 5)
        self.detail_table.setHorizontalHeaderLabels(["行号", "会话", "结果", "错误码", "错误信息"])
        _configure_data_table(self.detail_table, minimum_height=220)
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
        layout.addStretch(1)

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
        card.setMinimumWidth(160)
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
        _configure_form_layout(form)

        self.is_maximize = QCheckBox("自动最大化微信主界面")
        self.close_weixin = QCheckBox("任务完成后关闭微信")
        self.clear = QCheckBox("发送前清空输入区")
        self.search_pages = QSpinBox()
        self.search_pages.setRange(0, 50)
        _set_compact_width(self.search_pages, 120)
        self.send_delay = QDoubleSpinBox()
        self.send_delay.setRange(0.0, 10.0)
        self.send_delay.setSingleStep(0.1)
        _set_compact_width(self.send_delay, 120)
        self.window_width = QSpinBox()
        self.window_width.setRange(960, 3840)
        _set_compact_width(self.window_width, 120)
        self.window_height = QSpinBox()
        self.window_height.setRange(720, 2160)
        _set_compact_width(self.window_height, 120)
        self.import_encoding = QComboBox()
        self.import_encoding.addItems(["auto", "utf-8-sig", "gbk"])
        _set_compact_width(self.import_encoding, 150)
        self.theme = QComboBox()
        self.theme.addItems(["light"])
        _set_compact_width(self.theme, 150)

        behavior_row = _flow_container(self.is_maximize, self.close_weixin, self.clear, h_spacing=16)

        search_label = QLabel("搜索页数")
        search_label.setProperty("role", "muted")
        delay_label = QLabel("发送间隔(秒)")
        delay_label.setProperty("role", "muted")
        runtime_row = _flow_container(search_label, self.search_pages, delay_label, self.send_delay, h_spacing=12)

        width_label = QLabel("宽度")
        width_label.setProperty("role", "muted")
        height_label = QLabel("高度")
        height_label.setProperty("role", "muted")
        size_row = _flow_container(width_label, self.window_width, height_label, self.window_height, h_spacing=12)

        encoding_label = QLabel("CSV 编码")
        encoding_label.setProperty("role", "muted")
        theme_label = QLabel("主题")
        theme_label.setProperty("role", "muted")
        misc_row = _flow_container(encoding_label, self.import_encoding, theme_label, self.theme, h_spacing=12)

        form.addRow("执行行为", behavior_row)
        form.addRow("发送设置", runtime_row)
        form.addRow("窗口大小", size_row)
        form.addRow("其他设置", misc_row)
        card.body_layout.addLayout(form)
        self.save_button = QPushButton("保存设置")
        card.body_layout.addWidget(self.save_button)
        layout.addWidget(card)
        layout.addStretch(1)
        layout.addStretch(1)
