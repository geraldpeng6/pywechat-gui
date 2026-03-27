from __future__ import annotations

from PySide6.QtCore import QObject, Signal

from .executor import BatchExecutor
from .error_handling import map_exception
from .models import FileBatchRow, MessageBatchRow, RuntimeOptions, TaskType


class BatchWorker(QObject):
    progress = Signal(int, int, str)
    finished = Signal(object)
    failed = Signal(object)

    def __init__(
        self,
        executor: BatchExecutor,
        task_type: TaskType,
        rows: list[MessageBatchRow] | list[FileBatchRow],
        runtime_options: RuntimeOptions,
        source_execution_id: int | None = None,
    ):
        super().__init__()
        self.executor = executor
        self.task_type = task_type
        self.rows = rows
        self.runtime_options = runtime_options
        self.source_execution_id = source_execution_id
        self._stop_requested = False

    def request_stop(self) -> None:
        self._stop_requested = True

    def run(self) -> None:
        try:
            if self.task_type is TaskType.MESSAGE:
                record = self.executor.run_message_batch(
                    rows=self.rows,
                    runtime_options=self.runtime_options,
                    on_progress=self._emit_progress,
                    should_stop=lambda: self._stop_requested,
                    source_execution_id=self.source_execution_id,
                )
            else:
                record = self.executor.run_file_batch(
                    rows=self.rows,
                    runtime_options=self.runtime_options,
                    on_progress=self._emit_progress,
                    should_stop=lambda: self._stop_requested,
                    source_execution_id=self.source_execution_id,
                )
            self.finished.emit(record)
        except Exception as exc:  # pragma: no cover - GUI runtime safety
            self.failed.emit(map_exception(exc))

    def _emit_progress(self, current: int, total: int, session_name: str) -> None:
        self.progress.emit(current, total, session_name)
