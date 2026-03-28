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


class ExportServiceTestCase(unittest.TestCase):
    def test_export_bundle_writes_summary_and_messages(self):
        service = ChatExportService(FakeAdapter())
        with tempfile.TemporaryDirectory() as tempdir:
            request = ChatExportRequest(
                session_name="测试群",
                target_folder=tempdir,
                export_messages=True,
                export_files=True,
                export_images=True,
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
            self.assertTrue(summary["warnings"])


if __name__ == "__main__":
    unittest.main()
