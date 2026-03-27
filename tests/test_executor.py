from __future__ import annotations

import unittest

from pyweixin_gui.executor import BatchExecutor, failed_rows_from_execution
from pyweixin_gui.models import MessageBatchRow, RuntimeOptions


class FakeAdapter:
    def send_message(self, row, options):
        if row.session_name == "失败项":
            raise ValueError("boom")

    def send_file(self, row, options):
        return None

    @staticmethod
    def map_runtime_exception(exc):
        from pyweixin_gui.error_handling import map_exception

        return map_exception(exc)


class ExecutorTestCase(unittest.TestCase):
    def test_run_message_batch_keeps_going_after_failures(self):
        executor = BatchExecutor(FakeAdapter())
        rows = [
            MessageBatchRow(session_name="成功项", message="1"),
            MessageBatchRow(session_name="失败项", message="2"),
        ]
        record = executor.run_message_batch(rows, RuntimeOptions())
        self.assertEqual(record.success_count, 1)
        self.assertEqual(record.failure_count, 1)
        self.assertEqual(record.status, "completed")
        self.assertEqual(len(failed_rows_from_execution(record)), 1)

    def test_stop_before_next_row(self):
        executor = BatchExecutor(FakeAdapter())
        rows = [
            MessageBatchRow(session_name="成功项", message="1"),
            MessageBatchRow(session_name="成功项2", message="2"),
        ]
        called = {"count": 0}

        def should_stop():
            called["count"] += 1
            return called["count"] > 1

        record = executor.run_message_batch(rows, RuntimeOptions(), should_stop=should_stop)
        self.assertEqual(record.status, "stopped")
        self.assertEqual(record.success_count, 1)


if __name__ == "__main__":
    unittest.main()
