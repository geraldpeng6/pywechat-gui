from __future__ import annotations

import unittest

from pyweixin_gui.models import ChatBatchExportRequest, ExportHistoryRecord, ResourceExportKind, ResourceExportRequest
from pyweixin_gui.presentation import export_history_can_rerun, export_history_can_retry_failed, export_history_failed_sessions, format_export_history_detail, rebuild_export_request, serialize_export_detail


class ExportHistoryPresentationTestCase(unittest.TestCase):
    def test_batch_export_record_can_retry_failed_sessions(self):
        request = ChatBatchExportRequest(
            session_names=["项目群", "客户群"],
            target_folder="C:/exports",
            export_messages=True,
            export_files=True,
            export_images=False,
            message_limit=100,
            file_limit=50,
        )
        record = ExportHistoryRecord(
            export_kind="chat_batch",
            title="批量会话导出",
            export_folder="C:/exports",
            exported_count=1,
            detail_json=serialize_export_detail(
                "chat_batch",
                request,
                {
                    "success_count": 1,
                    "failure_count": 1,
                    "failed_sessions": [{"session_name": "客户群", "error": "会话不存在"}],
                },
            ),
        )

        self.assertTrue(export_history_can_rerun(record))
        self.assertTrue(export_history_can_retry_failed(record))
        self.assertEqual(export_history_failed_sessions(record), ["客户群"])
        _, retry_request = rebuild_export_request(record, failed_only=True)
        self.assertEqual(retry_request.session_names, ["客户群"])
        self.assertIn("批量会话数", format_export_history_detail(record))

    def test_resource_export_record_rebuild(self):
        request = ResourceExportRequest(
            export_kind=ResourceExportKind.WXFILES,
            target_folder="C:/exports/wxfiles",
            year="2026",
            month="03",
        )
        record = ExportHistoryRecord(
            export_kind="wxfiles",
            title="微信聊天文件",
            export_folder="C:/exports/wxfiles",
            exported_count=10,
            detail_json=serialize_export_detail(
                "wxfiles",
                request,
                {"export_kind": "wxfiles", "exported_count": 10, "exported_paths": ["a.docx"]},
            ),
        )
        kind, rebuilt = rebuild_export_request(record)
        self.assertEqual(kind, "resource")
        self.assertEqual(rebuilt.export_kind, ResourceExportKind.WXFILES)
        self.assertEqual(rebuilt.month, "03")

    def test_legacy_record_without_request_cannot_rerun(self):
        record = ExportHistoryRecord(
            export_kind="chat",
            title="旧记录",
            export_folder="C:/exports",
            exported_count=1,
            detail_json='{"type":"chat","messages_csv":"a.csv"}',
        )
        self.assertFalse(export_history_can_rerun(record))
        self.assertIn("较早版本", format_export_history_detail(record))
        self.assertIsNone(rebuild_export_request(record))


if __name__ == "__main__":
    unittest.main()
