from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from pyweixin_gui.models import ExecutionRecord, ExecutionRowResult, TaskTemplate, TaskType
from pyweixin_gui.storage import AppStorage


class StorageTestCase(unittest.TestCase):
    def setUp(self) -> None:
        self.tempdir = tempfile.TemporaryDirectory()
        self.storage = AppStorage(Path(self.tempdir.name) / "app.db")

    def tearDown(self) -> None:
        self.tempdir.cleanup()

    def test_template_roundtrip(self):
        template = TaskTemplate(name="消息模板", task_type=TaskType.MESSAGE, rows_json="[]")
        saved = self.storage.save_template(template)
        listed = self.storage.list_templates()
        self.assertEqual(saved.id, listed[0].id)
        self.assertEqual(listed[0].name, "消息模板")

    def test_execution_roundtrip(self):
        record = ExecutionRecord.new(TaskType.MESSAGE, row_count=1)
        record.status = "completed"
        record.finished_at = "2025-01-01T00:00:00"
        record.success_count = 1
        record.rows.append(
            ExecutionRowResult(
                row_index=0,
                session_name="好友A",
                success=True,
                row_payload_json='{"session_name":"好友A"}',
            )
        )
        saved = self.storage.save_execution(record)
        loaded = self.storage.get_execution(saved.id)
        self.assertEqual(loaded.success_count, 1)
        self.assertEqual(len(loaded.rows), 1)
        self.assertEqual(loaded.rows[0].session_name, "好友A")


if __name__ == "__main__":
    unittest.main()
