from __future__ import annotations

import json
from dataclasses import asdict
from datetime import UTC, datetime
from typing import Callable

from .adapter import PyWeixinAdapter
from .models import (
    ExecutionRecord,
    ExecutionRowResult,
    FileBatchRow,
    MessageBatchRow,
    RuntimeOptions,
    TaskType,
)


ProgressCallback = Callable[[int, int, str], None]
StopCallback = Callable[[], bool]


class BatchExecutor:
    def __init__(self, adapter: PyWeixinAdapter):
        self.adapter = adapter

    def run_message_batch(
        self,
        rows: list[MessageBatchRow],
        runtime_options: RuntimeOptions,
        on_progress: ProgressCallback | None = None,
        should_stop: StopCallback | None = None,
        source_execution_id: int | None = None,
    ) -> ExecutionRecord:
        return self._run(
            task_type=TaskType.MESSAGE,
            rows=rows,
            runtime_options=runtime_options,
            action=self.adapter.send_message,
            on_progress=on_progress,
            should_stop=should_stop,
            source_execution_id=source_execution_id,
        )

    def run_file_batch(
        self,
        rows: list[FileBatchRow],
        runtime_options: RuntimeOptions,
        on_progress: ProgressCallback | None = None,
        should_stop: StopCallback | None = None,
        source_execution_id: int | None = None,
    ) -> ExecutionRecord:
        return self._run(
            task_type=TaskType.FILE,
            rows=rows,
            runtime_options=runtime_options,
            action=self.adapter.send_file,
            on_progress=on_progress,
            should_stop=should_stop,
            source_execution_id=source_execution_id,
        )

    def _run(
        self,
        task_type: TaskType,
        rows: list[MessageBatchRow] | list[FileBatchRow],
        runtime_options: RuntimeOptions,
        action: Callable[[MessageBatchRow | FileBatchRow, RuntimeOptions], None],
        on_progress: ProgressCallback | None,
        should_stop: StopCallback | None,
        source_execution_id: int | None,
    ) -> ExecutionRecord:
        active_rows = [row for row in rows if row.enabled]
        record = ExecutionRecord.new(task_type=task_type, row_count=len(active_rows), source_execution_id=source_execution_id)
        for current, row in enumerate(active_rows, start=1):
            if should_stop and should_stop():
                record.status = "stopped"
                break
            if on_progress:
                on_progress(current - 1, len(active_rows), row.session_name)
            try:
                action(row, runtime_options)
                row_result = ExecutionRowResult(
                    row_index=current - 1,
                    session_name=row.session_name,
                    success=True,
                    row_payload_json=json.dumps(asdict(row), ensure_ascii=False),
                )
                record.success_count += 1
            except Exception as exc:  # pragma: no cover - exercised through adapter integration
                ui_error = self.adapter.map_runtime_exception(exc)
                row_result = ExecutionRowResult(
                    row_index=current - 1,
                    session_name=row.session_name,
                    success=False,
                    error_code=ui_error.code,
                    error_message=ui_error.message,
                    raw_error=ui_error.diagnostic_text,
                    row_payload_json=json.dumps(asdict(row), ensure_ascii=False),
                )
                record.failure_count += 1
            record.rows.append(row_result)
        else:
            record.status = "completed"
        if record.status == "running":
            record.status = "completed"
        record.finished_at = datetime.now(UTC).isoformat(timespec="seconds")
        return record


def failed_rows_from_execution(execution: ExecutionRecord) -> list[dict]:
    failed: list[dict] = []
    for row in execution.rows:
        if not row.success:
            failed.append(row.payload)
    return failed
