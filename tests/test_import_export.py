from __future__ import annotations

import importlib.util
import tempfile
import unittest
from pathlib import Path

from pyweixin_gui.import_export import dump_route_rows, dump_rows, dump_table, load_route_rows, load_rows, load_session_names
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

    @unittest.skipUnless(importlib.util.find_spec("openpyxl"), "openpyxl not installed")
    def test_load_session_names_from_xlsx(self):
        with tempfile.TemporaryDirectory() as tempdir:
            path = Path(tempdir) / "sessions.xlsx"
            dump_table(
                ["session_name", "remark"],
                [
                    {"session_name": "项目群", "remark": "A"},
                    {"session_name": "客户群", "remark": "B"},
                ],
                path,
            )
            self.assertEqual(load_session_names(path), ["项目群", "客户群"])

    def test_load_route_rows_from_csv_with_chinese_headers(self):
        with tempfile.TemporaryDirectory() as tempdir:
            path = Path(tempdir) / "routes.csv"
            dump_table(
                ["下游会话"],
                [
                    {"下游会话": "下游B"},
                    {"下游会话": "下游C"},
                ],
                path,
            )
            rows = load_route_rows(path)
            self.assertEqual(rows[0].downstream_session, "下游B")
            self.assertEqual(rows[1].downstream_session, "下游C")

    @unittest.skipUnless(importlib.util.find_spec("openpyxl"), "openpyxl not installed")
    def test_dump_route_rows_template(self):
        with tempfile.TemporaryDirectory() as tempdir:
            path = Path(tempdir) / "route-template.xlsx"
            dump_route_rows([], path)
            self.assertTrue(path.exists())

    def test_load_route_rows_expands_multi_downstream_column(self):
        with tempfile.TemporaryDirectory() as tempdir:
            path = Path(tempdir) / "routes.csv"
            dump_table(
                ["下游会话列表"],
                [
                    {"下游会话列表": "下游B|下游C"},
                ],
                path,
            )
            rows = load_route_rows(path)
            self.assertEqual(len(rows), 2)
            self.assertEqual(rows[0].downstream_session, "下游B")
            self.assertEqual(rows[1].downstream_session, "下游C")


if __name__ == "__main__":
    unittest.main()
