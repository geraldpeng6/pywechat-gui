from __future__ import annotations

import importlib.util
import tempfile
import unittest
from pathlib import Path

from pyweixin_gui.import_export import dump_rows, load_rows
from pyweixin_gui.models import TaskType


class ImportExportTestCase(unittest.TestCase):
    def test_csv_roundtrip(self):
        with tempfile.TemporaryDirectory() as tempdir:
            path = Path(tempdir) / "messages.csv"
            rows = [
                {
                    "enabled": True,
                    "session_name": "好友A",
                    "message": "你好",
                    "at_members": "",
                    "at_all": False,
                    "clear_before_send": True,
                    "send_delay_sec": 0.2,
                    "remark": "备注",
                }
            ]
            dump_rows(TaskType.MESSAGE, rows, path)
            loaded = load_rows(TaskType.MESSAGE, path)
            self.assertEqual(loaded[0].session_name, "好友A")
            self.assertEqual(loaded[0].message, "你好")

    @unittest.skipUnless(importlib.util.find_spec("openpyxl"), "openpyxl not installed")
    def test_xlsx_roundtrip(self):
        with tempfile.TemporaryDirectory() as tempdir:
            path = Path(tempdir) / "files.xlsx"
            rows = [
                {
                    "enabled": True,
                    "session_name": "群聊A",
                    "file_paths": "C:/test/a.txt|C:/test/b.txt",
                    "with_message": True,
                    "message": "请查收",
                    "message_first": False,
                    "remark": "",
                }
            ]
            dump_rows(TaskType.FILE, rows, path)
            loaded = load_rows(TaskType.FILE, path)
            self.assertEqual(loaded[0].session_name, "群聊A")
            self.assertTrue(loaded[0].with_message)


if __name__ == "__main__":
    unittest.main()
