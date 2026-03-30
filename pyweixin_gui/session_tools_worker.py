from __future__ import annotations

from PySide6.QtCore import QObject, Signal

from .error_handling import map_exception
from .models import GroupMembersRequest, RuntimeOptions, SessionScanRequest
from .session_tools_service import SessionToolsService


class SessionToolsWorker(QObject):
    progress = Signal(str)
    finished = Signal(str, object)
    failed = Signal(object)

    def __init__(
        self,
        service: SessionToolsService,
        action: str,
        runtime_options: RuntimeOptions,
        request: SessionScanRequest | GroupMembersRequest | None = None,
    ):
        super().__init__()
        self.service = service
        self.action = action
        self.runtime_options = runtime_options
        self.request = request

    def request_stop(self) -> None:
        return

    def run(self) -> None:
        try:
            if self.action == "scan_sessions":
                result = self.service.scan_sessions(
                    request=self.request if isinstance(self.request, SessionScanRequest) else SessionScanRequest(),
                    runtime_options=self.runtime_options,
                    on_progress=self.progress.emit,
                )
            elif self.action == "scan_groups":
                result = self.service.scan_groups(
                    runtime_options=self.runtime_options,
                    on_progress=self.progress.emit,
                )
            elif self.action == "load_group_members":
                if not isinstance(self.request, GroupMembersRequest):
                    raise ValueError("缺少群成员查询参数")
                result = self.service.load_group_members(
                    request=self.request,
                    runtime_options=self.runtime_options,
                    on_progress=self.progress.emit,
                )
            else:
                raise ValueError(f"不支持的工具动作: {self.action}")
            self.finished.emit(self.action, result)
        except Exception as exc:  # pragma: no cover - GUI runtime safety
            self.failed.emit(map_exception(exc))
