from __future__ import annotations

import unittest

from pyweixin_gui.models import FileBatchRow, MessageBatchRow


class ModelsTestCase(unittest.TestCase):
    def test_message_row_validation_requires_session_and_message(self):
        row = MessageBatchRow()
        errors = row.validate()
        self.assertIn("session_name", errors)
        self.assertIn("message", errors)

    def test_message_row_parses_at_members(self):
        row = MessageBatchRow(at_members="张三| 李四 |")
        self.assertEqual(row.at_member_list(), ["张三", "李四"])

    def test_file_row_validation_checks_paths(self):
        row = FileBatchRow(session_name="测试", file_paths="/tmp/not-exist.txt")
        errors = row.validate()
        self.assertIn("file_paths", errors)


if __name__ == "__main__":
    unittest.main()
