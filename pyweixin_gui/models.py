from __future__ import annotations

import dataclasses
import json
from dataclasses import asdict, dataclass, field
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Any


class TaskType(str, Enum):
    MESSAGE = "message"
    FILE = "file"


def _parse_bool(value: Any, default: bool = False) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    text = str(value).strip().lower()
    if text in {"1", "true", "yes", "y", "on", "是"}:
        return True
    if text in {"0", "false", "no", "n", "off", "否", ""}:
        return False
    return default


def _parse_optional_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, str) and not value.strip():
        return None
    return float(value)


def _normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).replace("\r\n", "\n").strip()


def split_pipe_values(value: str) -> list[str]:
    return [part.strip() for part in value.split("|") if part.strip()]


@dataclass
class MessageBatchRow:
    enabled: bool = True
    session_name: str = ""
    message: str = ""
    at_members: str = ""
    at_all: bool = False
    clear_before_send: bool = True
    send_delay_sec: float | None = None
    remark: str = ""

    @classmethod
    def from_mapping(cls, mapping: dict[str, Any]) -> "MessageBatchRow":
        return cls(
            enabled=_parse_bool(mapping.get("enabled"), True),
            session_name=_normalize_text(mapping.get("session_name")),
            message=_normalize_text(mapping.get("message")),
            at_members=_normalize_text(mapping.get("at_members")),
            at_all=_parse_bool(mapping.get("at_all")),
            clear_before_send=_parse_bool(mapping.get("clear_before_send"), True),
            send_delay_sec=_parse_optional_float(mapping.get("send_delay_sec")),
            remark=_normalize_text(mapping.get("remark")),
        )

    def validate(self) -> dict[str, str]:
        errors: dict[str, str] = {}
        if not self.session_name:
            errors["session_name"] = "会话名称不能为空"
        if not self.message:
            errors["message"] = "消息内容不能为空"
        if self.send_delay_sec is not None and self.send_delay_sec < 0:
            errors["send_delay_sec"] = "发送间隔不能为负数"
        return errors

    def at_member_list(self) -> list[str]:
        return split_pipe_values(self.at_members)


@dataclass
class FileBatchRow:
    enabled: bool = True
    session_name: str = ""
    file_paths: str = ""
    with_message: bool = False
    message: str = ""
    message_first: bool = False
    remark: str = ""

    @classmethod
    def from_mapping(cls, mapping: dict[str, Any]) -> "FileBatchRow":
        return cls(
            enabled=_parse_bool(mapping.get("enabled"), True),
            session_name=_normalize_text(mapping.get("session_name")),
            file_paths=_normalize_text(mapping.get("file_paths")),
            with_message=_parse_bool(mapping.get("with_message")),
            message=_normalize_text(mapping.get("message")),
            message_first=_parse_bool(mapping.get("message_first")),
            remark=_normalize_text(mapping.get("remark")),
        )

    def validate(self) -> dict[str, str]:
        errors: dict[str, str] = {}
        if not self.session_name:
            errors["session_name"] = "会话名称不能为空"
        if not self.file_paths:
            errors["file_paths"] = "至少需要一个文件路径"
        resolved_paths = self.path_list()
        if not resolved_paths:
            errors["file_paths"] = "至少需要一个有效文件路径"
        else:
            missing = [path for path in resolved_paths if not Path(path).is_file()]
            if missing:
                errors["file_paths"] = f"存在无效文件: {missing[0]}"
        if self.with_message and not self.message:
            errors["message"] = "启用附带消息时，消息内容不能为空"
        return errors

    def path_list(self) -> list[str]:
        return split_pipe_values(self.file_paths)


@dataclass
class RuntimeOptions:
    is_maximize: bool = False
    close_weixin: bool = False
    search_pages: int = 5
    send_delay: float = 0.2
    clear: bool = True
    window_size: tuple[int, int] = (1000, 1000)


@dataclass
class TaskTemplate:
    name: str
    task_type: TaskType
    rows_json: str
    id: int | None = None
    created_at: str | None = None
    updated_at: str | None = None


@dataclass
class ExecutionRowResult:
    row_index: int
    session_name: str
    success: bool
    error_code: str | None = None
    error_message: str | None = None
    raw_error: str | None = None
    row_payload_json: str = ""

    @property
    def payload(self) -> dict[str, Any]:
        if not self.row_payload_json:
            return {}
        return json.loads(self.row_payload_json)


@dataclass
class ExecutionRecord:
    task_type: TaskType
    started_at: str
    finished_at: str | None
    status: str
    row_count: int
    success_count: int
    failure_count: int
    source_template_id: int | None = None
    source_execution_id: int | None = None
    id: int | None = None
    rows: list[ExecutionRowResult] = field(default_factory=list)

    @classmethod
    def new(cls, task_type: TaskType, row_count: int, source_execution_id: int | None = None) -> "ExecutionRecord":
        return cls(
            task_type=task_type,
            started_at=datetime.utcnow().isoformat(timespec="seconds"),
            finished_at=None,
            status="running",
            row_count=row_count,
            success_count=0,
            failure_count=0,
            source_execution_id=source_execution_id,
        )


@dataclass
class AppSettings:
    is_maximize: bool = False
    close_weixin: bool = False
    search_pages: int = 5
    send_delay: float = 0.2
    clear: bool = True
    window_width: int = 1280
    window_height: int = 860
    import_encoding: str = "auto"
    theme: str = "light"
    history_retention: str = "forever"
    first_run_risk_ack: bool = False

    @classmethod
    def from_mapping(cls, mapping: dict[str, Any]) -> "AppSettings":
        return cls(
            is_maximize=_parse_bool(mapping.get("is_maximize")),
            close_weixin=_parse_bool(mapping.get("close_weixin")),
            search_pages=int(mapping.get("search_pages", 5)),
            send_delay=float(mapping.get("send_delay", 0.2)),
            clear=_parse_bool(mapping.get("clear"), True),
            window_width=int(mapping.get("window_width", 1280)),
            window_height=int(mapping.get("window_height", 860)),
            import_encoding=str(mapping.get("import_encoding", "auto")),
            theme=str(mapping.get("theme", "light")),
            history_retention=str(mapping.get("history_retention", "forever")),
            first_run_risk_ack=_parse_bool(mapping.get("first_run_risk_ack")),
        )

    def runtime_options(self) -> RuntimeOptions:
        return RuntimeOptions(
            is_maximize=self.is_maximize,
            close_weixin=self.close_weixin,
            search_pages=self.search_pages,
            send_delay=self.send_delay,
            clear=self.clear,
            window_size=(self.window_width, self.window_height),
        )

    def to_json(self) -> str:
        return json.dumps(asdict(self), ensure_ascii=False, indent=2)


@dataclass
class EnvironmentStatus:
    operating_system: str
    python_version: str
    gui_version: str
    wechat_running: bool
    wechat_path: str
    login_status: str
    status_message: str
    advice: list[str] = field(default_factory=list)


@dataclass
class ChatExportRequest:
    session_name: str
    target_folder: str
    export_messages: bool = True
    export_files: bool = True
    export_images: bool = False
    message_limit: int = 100
    file_limit: int = 50

    def validate(self) -> dict[str, str]:
        errors: dict[str, str] = {}
        if not self.session_name.strip():
            errors["session_name"] = "会话名称不能为空"
        if not self.target_folder.strip():
            errors["target_folder"] = "请选择导出文件夹"
        if not self.export_messages and not self.export_files:
            errors["export_scope"] = "请至少勾选一种导出内容"
        if self.message_limit <= 0:
            errors["message_limit"] = "消息条数必须大于 0"
        if self.file_limit <= 0:
            errors["file_limit"] = "文件数量必须大于 0"
        return errors


@dataclass
class ChatBatchExportRequest:
    session_names: list[str]
    target_folder: str
    export_messages: bool = True
    export_files: bool = True
    export_images: bool = False
    message_limit: int = 100
    file_limit: int = 50

    def validate(self) -> dict[str, str]:
        errors: dict[str, str] = {}
        if not [name.strip() for name in self.session_names if name.strip()]:
            errors["session_names"] = "请至少填写一个会话名称"
        if not self.target_folder.strip():
            errors["target_folder"] = "请选择导出文件夹"
        if not self.export_messages and not self.export_files:
            errors["export_scope"] = "请至少勾选一种导出内容"
        if self.message_limit <= 0:
            errors["message_limit"] = "消息条数必须大于 0"
        if self.file_limit <= 0:
            errors["file_limit"] = "文件数量必须大于 0"
        return errors


@dataclass
class ChatExportResult:
    session_name: str
    export_folder: str
    message_count: int = 0
    file_count: int = 0
    messages_csv: str | None = None
    messages_json: str | None = None
    files_folder: str | None = None
    summary_json: str | None = None
    summary_txt: str | None = None
    warnings: list[str] = field(default_factory=list)


@dataclass
class ChatBatchExportResult:
    export_root: str
    total_sessions: int
    success_count: int
    failure_count: int
    session_results: list[ChatExportResult] = field(default_factory=list)
    failed_sessions: list[dict[str, str]] = field(default_factory=list)
    summary_txt: str | None = None


class ResourceExportKind(str, Enum):
    RECENT_FILES = "recent_files"
    WXFILES = "wxfiles"
    VIDEOS = "videos"


@dataclass
class ResourceExportRequest:
    export_kind: ResourceExportKind
    target_folder: str
    year: str = datetime.now().strftime("%Y")
    month: str = ""

    def validate(self) -> dict[str, str]:
        errors: dict[str, str] = {}
        if not self.target_folder.strip():
            errors["target_folder"] = "请选择导出文件夹"
        if not self.year.isdigit() or len(self.year) != 4:
            errors["year"] = "年份必须是 4 位数字，例如 2026"
        if self.month and (not self.month.isdigit() or not 1 <= int(self.month) <= 12):
            errors["month"] = "月份必须是 1-12 之间的数字"
        return errors


@dataclass
class ResourceExportResult:
    export_kind: ResourceExportKind
    target_folder: str
    exported_count: int
    exported_paths: list[str] = field(default_factory=list)
    summary_txt: str | None = None


@dataclass
class ExportHistoryRecord:
    export_kind: str
    title: str
    export_folder: str
    exported_count: int
    summary_path: str | None = None
    detail_json: str = "{}"
    id: int | None = None
    created_at: str | None = None


@dataclass
class SessionScanRequest:
    chatted_only: bool = False


@dataclass
class SessionSummaryRow:
    session_name: str
    last_time: str = ""
    last_message: str = ""


@dataclass
class SessionScanResult:
    rows: list[SessionSummaryRow] = field(default_factory=list)


@dataclass
class GroupSummaryRow:
    group_name: str
    member_count: str = ""


@dataclass
class GroupScanResult:
    rows: list[GroupSummaryRow] = field(default_factory=list)


class RelayItemType(str, Enum):
    TEXT = "text"
    FILE = "file"
    IMAGE = "image"


@dataclass
class RelayPackageRow:
    enabled: bool = True
    sequence: int = 1
    item_type: RelayItemType = RelayItemType.TEXT
    source_session: str = ""
    content: str = ""
    file_path: str = ""
    collected_at: str = ""
    remark: str = ""

    @classmethod
    def from_mapping(cls, mapping: dict[str, Any]) -> "RelayPackageRow":
        item_type = str(mapping.get("item_type", RelayItemType.TEXT.value) or RelayItemType.TEXT.value).strip().lower()
        if item_type not in {member.value for member in RelayItemType}:
            item_type = RelayItemType.TEXT.value
        sequence = mapping.get("sequence", 1)
        try:
            sequence_value = int(sequence)
        except (TypeError, ValueError):
            sequence_value = 1
        return cls(
            enabled=_parse_bool(mapping.get("enabled"), True),
            sequence=max(1, sequence_value),
            item_type=RelayItemType(item_type),
            source_session=_normalize_text(mapping.get("source_session")),
            content=_normalize_text(mapping.get("content")),
            file_path=_normalize_text(mapping.get("file_path")),
            collected_at=_normalize_text(mapping.get("collected_at")),
            remark=_normalize_text(mapping.get("remark")),
        )

    def validate(self) -> dict[str, str]:
        errors: dict[str, str] = {}
        if self.item_type is RelayItemType.TEXT:
            if not self.content:
                errors["content"] = "文本项不能为空"
        else:
            if not self.file_path:
                errors["file_path"] = "文件或图片路径不能为空"
            elif not Path(self.file_path).is_file():
                errors["file_path"] = "文件或图片路径不存在"
        return errors

    def preview_text(self) -> str:
        if self.item_type is RelayItemType.TEXT:
            return self.content
        if self.file_path:
            return Path(self.file_path).name
        return self.content


@dataclass
class RelayRouteRow:
    enabled: bool = True
    upstream_session: str = ""
    downstream_session: str = ""
    validation_status: str = "未验证"
    validation_message: str = ""
    remark: str = ""

    @classmethod
    def from_mapping(cls, mapping: dict[str, Any]) -> "RelayRouteRow":
        return cls(
            enabled=_parse_bool(mapping.get("enabled"), True),
            upstream_session=_normalize_text(mapping.get("upstream_session")),
            downstream_session=_normalize_text(mapping.get("downstream_session")),
            validation_status=_normalize_text(mapping.get("validation_status")) or "未验证",
            validation_message=_normalize_text(mapping.get("validation_message")),
            remark=_normalize_text(mapping.get("remark")),
        )

    def validate(self) -> dict[str, str]:
        errors: dict[str, str] = {}
        if not self.upstream_session:
            errors["upstream_session"] = "上游会话不能为空"
        if not self.downstream_session:
            errors["downstream_session"] = "下游会话不能为空"
        return errors


@dataclass
class RelayCollectTextRequest:
    source_session: str
    message_limit: int = 20

    def validate(self) -> dict[str, str]:
        errors: dict[str, str] = {}
        if not self.source_session.strip():
            errors["source_session"] = "请先填写上游会话"
        if self.message_limit <= 0:
            errors["message_limit"] = "消息条数必须大于 0"
        return errors


@dataclass
class RelayCollectFilesRequest:
    source_session: str
    file_limit: int = 10

    def validate(self) -> dict[str, str]:
        errors: dict[str, str] = {}
        if not self.source_session.strip():
            errors["source_session"] = "请先填写上游会话"
        if self.file_limit <= 0:
            errors["file_limit"] = "文件数量必须大于 0"
        return errors


@dataclass
class RelayCollectionResult:
    source_session: str
    rows: list[RelayPackageRow] = field(default_factory=list)
    warning: str = ""


@dataclass
class RelayValidationRequest:
    source_session: str
    route_rows: list[RelayRouteRow] = field(default_factory=list)

    def validate(self) -> dict[str, str]:
        if not self.source_session.strip():
            return {"source_session": "请先填写当前上游会话"}
        return {}


@dataclass
class RelayValidationResult:
    source_session: str
    route_rows: list[RelayRouteRow] = field(default_factory=list)
    checked_count: int = 0
    found_count: int = 0
    missing_count: int = 0


@dataclass
class RelaySendRequest:
    source_session: str
    package_rows: list[RelayPackageRow] = field(default_factory=list)
    route_rows: list[RelayRouteRow] = field(default_factory=list)
    test_only: bool = False

    def validate(self) -> dict[str, str]:
        errors: dict[str, str] = {}
        if not self.package_rows:
            errors["package_rows"] = "请先准备至少一条转发内容"
        if not self.test_only and not self.source_session.strip():
            errors["source_session"] = "正式发送前请先填写上游会话"
        return errors


@dataclass
class RelayTargetResult:
    target_session: str
    success: bool
    sent_count: int
    error_message: str = ""


@dataclass
class RelaySendResult:
    source_session: str
    test_only: bool
    item_count: int
    target_count: int
    success_count: int
    failure_count: int
    results: list[RelayTargetResult] = field(default_factory=list)


def dataclass_to_json(items: list[MessageBatchRow] | list[FileBatchRow]) -> str:
    return json.dumps([asdict(item) for item in items], ensure_ascii=False)


def dataclass_from_json(task_type: TaskType, payload: str) -> list[MessageBatchRow] | list[FileBatchRow]:
    rows = json.loads(payload or "[]")
    if task_type is TaskType.MESSAGE:
        return [MessageBatchRow.from_mapping(row) for row in rows]
    return [FileBatchRow.from_mapping(row) for row in rows]


def clone_row(row: MessageBatchRow | FileBatchRow) -> MessageBatchRow | FileBatchRow:
    return dataclasses.replace(row)
