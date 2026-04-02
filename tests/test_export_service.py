from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from pyweixin_gui.export_service import ChatExportService
from pyweixin_gui.models import ChatExportRequest, RuntimeOptions


class FakeAdapter:
    def dump_chat_history(self, session_name, number, options):
        return ["第一条消息", "第二条消息"], ["昨天 12:00", "昨天 12:01"]

    def save_chat_files(self, session_name, number, target_folder, options):
        folder = Path(target_folder)
        folder.mkdir(parents=True, exist_ok=True)
        (folder / "报告.pdf").write_text("demo", encoding="utf-8")
        return [str(folder / "报告.pdf")]


class ParseFailureAdapter(FakeAdapter):
    def save_chat_files(self, session_name, number, target_folder, options):
        raise AttributeError("'NoneType' object has no attribute 'group'")


class ExportServiceTestCase(unittest.TestCase):
    def test_export_bundle_writes_summary_and_messages(self):
        service = ChatExportService(FakeAdapter())
        with tempfile.TemporaryDirectory() as tempdir:
            request = ChatExportRequest(
                session_name="测试群",
                target_folder=tempdir,
                export_messages=True,
                export_files=True,
                export_images=False,
                message_limit=10,
                file_limit=10,
            )
            result = service.export_chat_bundle(request, RuntimeOptions())
            export_folder = Path(result.export_folder)
            self.assertTrue(export_folder.exists())
            self.assertTrue(Path(result.messages_csv).exists())
            self.assertTrue(Path(result.messages_json).exists())
            self.assertTrue(Path(result.summary_json).exists())
            self.assertTrue(Path(result.summary_txt).exists())
            self.assertEqual(result.message_count, 2)
            self.assertEqual(result.file_count, 1)
            summary = json.loads(Path(result.summary_json).read_text(encoding="utf-8"))
            self.assertEqual(summary["session_name"], "测试群")
            self.assertEqual(summary["warnings"], [])

    def test_export_bundle_keeps_messages_when_chat_file_parse_fails(self):
        service = ChatExportService(ParseFailureAdapter())
        with tempfile.TemporaryDirectory() as tempdir:
            request = ChatExportRequest(
                session_name="好友A",
                target_folder=tempdir,
                export_messages=True,
                export_files=True,
                export_images=False,
                message_limit=10,
                file_limit=10,
            )
            result = service.export_chat_bundle(request, RuntimeOptions())
            self.assertEqual(result.message_count, 2)
            self.assertEqual(result.file_count, 0)
            self.assertTrue(result.messages_csv)
            self.assertTrue(result.files_folder)
            self.assertTrue(any("聊天文件未整理到结果中" in warning for warning in result.warnings))


if __name__ == "__main__":
    unittest.main()
