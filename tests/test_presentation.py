from __future__ import annotations

import unittest

from pyweixin_gui.models import ExecutionRecord, ExecutionRowResult, ExportHistoryRecord, TaskTemplate, TaskType
from pyweixin_gui.presentation import (
    execution_metrics,
    filter_executions,
    filter_export_records,
    filter_templates,
    summarize_failures,
    template_metrics,
    template_type_label,
)


class PresentationTestCase(unittest.TestCase):
    def test_filter_templates_and_metrics(self):
        templates = [
            TaskTemplate(name="本周资料发送", task_type=TaskType.RELAY_SEND, rows_json="[]", updated_at="2026-03-29"),
            TaskTemplate(name="客户通知", task_type=TaskType.MESSAGE, rows_json="[]", updated_at="2026-03-28"),
            TaskTemplate(name="财务附件", task_type=TaskType.FILE, rows_json="[]", updated_at="2026-03-27"),
        ]
        filtered = filter_templates(templates, "发送")
        metrics = template_metrics(templates)
        self.assertEqual(len(filtered), 1)
        self.assertEqual(filtered[0].name, "本周资料发送")
        self.assertEqual(metrics["total"], 3)
        self.assertEqual(metrics["send"], 1)
        self.assertEqual(metrics["message"], 1)
        self.assertEqual(metrics["file"], 1)
        self.assertEqual(metrics["legacy"], 2)

    def test_filter_executions_and_metrics(self):
        executions = [
            ExecutionRecord(
                id=1,
                task_type=TaskType.MESSAGE,
                started_at="2026-03-28T10:00:00",
                finished_at="2026-03-28T10:01:00",
                status="completed",
                row_count=2,
                success_count=2,
                failure_count=0,
            ),
            ExecutionRecord(
                id=2,
                task_type=TaskType.FILE,
                started_at="2026-03-28T11:00:00",
                finished_at="2026-03-28T11:02:00",
                status="completed",
                row_count=3,
                success_count=2,
                failure_count=1,
            ),
        ]
        filtered = filter_executions(executions, "文件", failed_only=True)
        metrics = execution_metrics(executions)
        self.assertEqual(len(filtered), 1)
        self.assertEqual(filtered[0].id, 2)
        self.assertEqual(metrics["total"], 2)
        self.assertEqual(metrics["success"], 1)
        self.assertEqual(metrics["failed"], 1)

        relay_filtered = filter_executions(executions, "", failed_only=False, task_type_filter="message")
        self.assertEqual(len(relay_filtered), 1)
        self.assertEqual(relay_filtered[0].id, 1)

    def test_relay_execution_uses_readable_label(self):
        executions = [
            ExecutionRecord(
                id=3,
                task_type=TaskType.RELAY_TEST_SEND,
                started_at="2026-04-02T10:00:00",
                finished_at="2026-04-02T10:01:00",
                status="completed",
                row_count=1,
                success_count=1,
                failure_count=0,
            )
        ]
        filtered = filter_executions(executions, "发送测试", failed_only=False)
        self.assertEqual(len(filtered), 1)
        self.assertEqual(template_type_label(TaskType.RELAY_VALIDATE), "收件人验证")
        self.assertEqual(template_type_label(TaskType.RELAY_SEND), "发送任务")

    def test_summarize_failures(self):
        rows = [
            ExecutionRowResult(row_index=0, session_name="A", success=False, error_code="SESSION_NOT_FOUND"),
            ExecutionRowResult(row_index=1, session_name="B", success=False, error_code="SESSION_NOT_FOUND"),
            ExecutionRowResult(row_index=2, session_name="C", success=False, error_message="网络不可用"),
        ]
        summary = summarize_failures(rows)
        self.assertIn("SESSION_NOT_FOUND x2", summary)
        self.assertIn("网络不可用 x1", summary)

    def test_filter_export_records_by_kind(self):
        records = [
            ExportHistoryRecord(export_kind="chat", title="项目群", export_folder="C:/exports/a", exported_count=3, created_at="2026-04-03 10:00:00"),
            ExportHistoryRecord(export_kind="relay_package", title="周报发送包", export_folder="C:/exports/b", exported_count=5, created_at="2026-04-03 11:00:00"),
        ]
        filtered = filter_export_records(records, "", export_kind_filter="relay_package")
        self.assertEqual(len(filtered), 1)
        self.assertEqual(filtered[0].title, "周报发送包")


if __name__ == "__main__":
    unittest.main()
