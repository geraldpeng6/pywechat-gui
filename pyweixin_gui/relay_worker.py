from __future__ import annotations

from PySide6.QtCore import QObject, Signal

from .error_handling import map_exception
from .models import RelayCollectFilesRequest, RelayCollectMediaRequest, RelayCollectTextRequest, RelaySendRequest, RelayValidationRequest, RuntimeOptions
from .relay_service import RelayService


class RelayWorker(QObject):
    progress = Signal(str)
    finished = Signal(str, object)
    failed = Signal(object)

    def __init__(
        self,
        service: RelayService,
        action: str,
        runtime_options: RuntimeOptions,
        request: RelayCollectTextRequest | RelayCollectFilesRequest | RelayCollectMediaRequest | RelayValidationRequest | RelaySendRequest | None = None,
    ):
        super().__init__()
        self.service = service
        self.action = action
        self.runtime_options = runtime_options
        self.request = request
        self._stop_requested = False

    def request_stop(self) -> None:
        self._stop_requested = True

    def run(self) -> None:
        try:
            if self.action == "collect_texts":
                if not isinstance(self.request, RelayCollectTextRequest):
                    raise ValueError("缺少文本采集参数")
                result = self.service.collect_text_rows(self.request, self.runtime_options, self.progress.emit)
            elif self.action == "collect_files":
                if not isinstance(self.request, RelayCollectFilesRequest):
                    raise ValueError("缺少文件采集参数")
                result = self.service.collect_file_rows(self.request, self.runtime_options, self.progress.emit)
            elif self.action == "collect_media":
                if not isinstance(self.request, RelayCollectMediaRequest):
                    raise ValueError("缺少图片/视频采集参数")
                result = self.service.collect_media_rows(self.request, self.runtime_options, self.progress.emit)
            elif self.action == "validate_routes":
                if not isinstance(self.request, RelayValidationRequest):
                    raise ValueError("缺少路由验证参数")
                result = self.service.validate_routes(self.request, self.runtime_options, self.progress.emit)
            elif self.action == "send_package":
                if not isinstance(self.request, RelaySendRequest):
                    raise ValueError("缺少发送参数")
                result = self.service.send_package(
                    self.request,
                    self.runtime_options,
                    on_progress=self.progress.emit,
                    should_stop=lambda: self._stop_requested,
                )
            else:
                raise ValueError(f"不支持的转发动作: {self.action}")
            self.finished.emit(self.action, result)
        except Exception as exc:  # pragma: no cover - GUI runtime safety
            self.failed.emit(map_exception(exc))
