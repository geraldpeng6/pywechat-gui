from __future__ import annotations

from PySide6.QtCore import QObject, Signal

from .export_service import ChatExportService
from .error_handling import map_exception
from .models import ChatBatchExportRequest, ChatExportRequest, RuntimeOptions


class ChatExportWorker(QObject):
    progress = Signal(str)
    finished = Signal(object)
    failed = Signal(object)

    def __init__(self, service: ChatExportService, request: ChatExportRequest | ChatBatchExportRequest, runtime_options: RuntimeOptions):
        super().__init__()
        self.service = service
        self.request = request
        self.runtime_options = runtime_options
        self._stop_requested = False

    def request_stop(self) -> None:
        self._stop_requested = True

    def run(self) -> None:
        try:
            result = self.service.export_chat_bundle(
                request=self.request,
                runtime_options=self.runtime_options,
                on_progress=self.progress.emit,
                should_stop=lambda: self._stop_requested,
            ) if isinstance(self.request, ChatExportRequest) else self.service.export_chat_batch(
                request=self.request,
                runtime_options=self.runtime_options,
                on_progress=self.progress.emit,
                should_stop=lambda: self._stop_requested,
            )
            self.finished.emit(result)
        except Exception as exc:  # pragma: no cover - GUI runtime safety
            self.failed.emit(map_exception(exc))
