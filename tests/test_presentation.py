from __future__ import annotations

import unittest

from pyweixin_gui.models import ExecutionRecord, ExecutionRowResult, TaskTemplate, TaskType
from pyweixin_gui.presentation import (
    execution_metrics,
    filter_executions,
    filter_templates,
    summarize_failures,
    template_metrics,
)


class PresentationTestCase(unittest.TestCase):
    def test_filter_templates_and_metrics(self):
        templates = [
            TaskTemplate(name="客户通知", task_type=TaskType.MESSAGE, rows_json="[]", updated_at="2026-03-28"),
            TaskTemplate(name="财务附件", task_type=TaskType.FILE, rows_json="[]", updated_at="2026-03-27"),
        ]
        filtered = filter_templates(templates, "消息")
        metrics = template_metrics(templates)
        self.assertEqual(len(filtered), 1)
        self.assertEqual(filtered[0].name, "客户通知")
        self.assertEqual(metrics["total"], 2)
        self.assertEqual(metrics["message"], 1)
        self.assertEqual(metrics["file"], 1)

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

    def test_summarize_failures(self):
        rows = [
            ExecutionRowResult(row_index=0, session_name="A", success=False, error_code="SESSION_NOT_FOUND"),
            ExecutionRowResult(row_index=1, session_name="B", success=False, error_code="SESSION_NOT_FOUND"),
            ExecutionRowResult(row_index=2, session_name="C", success=False, error_message="网络不可用"),
        ]
        summary = summarize_failures(rows)
        self.assertIn("SESSION_NOT_FOUND x2", summary)
        self.assertIn("网络不可用 x1", summary)


if __name__ == "__main__":
    unittest.main()
