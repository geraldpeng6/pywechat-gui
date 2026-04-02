from __future__ import annotations

import importlib.util
import tempfile
import unittest
from pathlib import Path

from pyweixin_gui.export_service import ChatExportService
from pyweixin_gui.models import ChatBatchExportRequest, RuntimeOptions


class FakeAdapter:
    def dump_chat_history(self, session_name, number, options):
        if session_name == "失败群":
            raise RuntimeError("会话不存在")
        return [f"{session_name}-消息"], ["昨天 12:00"]

    def save_chat_files(self, session_name, number, target_folder, options):
        folder = Path(target_folder)
        folder.mkdir(parents=True, exist_ok=True)
        (folder / f"{session_name}.txt").write_text("demo", encoding="utf-8")
        return [str(folder / f"{session_name}.txt")]


class BatchExportServiceTestCase(unittest.TestCase):
    @unittest.skipUnless(importlib.util.find_spec("openpyxl"), "openpyxl not installed")
    def test_batch_export_records_success_and_failure(self):
        service = ChatExportService(FakeAdapter())
        with tempfile.TemporaryDirectory() as tempdir:
            request = ChatBatchExportRequest(
                session_names=["项目群", "失败群", "客户群"],
                target_folder=tempdir,
                export_messages=True,
                export_files=True,
                message_limit=10,
                file_limit=10,
            )
            result = service.export_chat_batch(request, RuntimeOptions())
            self.assertEqual(result.total_sessions, 3)
            self.assertEqual(result.success_count, 2)
            self.assertEqual(result.failure_count, 1)
            self.assertEqual(Path(result.export_root).name.startswith("批量会话导出-"), True)


if __name__ == "__main__":
    unittest.main()
