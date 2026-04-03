from __future__ import annotations

import unittest

from pyweixin_gui.models import ChatBatchExportRequest, ChatExportRequest, ExportHistoryRecord, ResourceExportKind, ResourceExportRequest
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

    def test_chat_export_record_shows_media_flags_and_counts(self):
        request = ChatExportRequest(
            session_name="项目群",
            target_folder="C:/exports",
            export_messages=True,
            export_files=False,
            export_images=True,
            message_limit=20,
            file_limit=8,
        )
        record = ExportHistoryRecord(
            export_kind="chat",
            title="项目群",
            export_folder="C:/exports/项目群",
            exported_count=5,
            detail_json=serialize_export_detail(
                "chat",
                request,
                {
                    "message_count": 3,
                    "file_count": 0,
                    "media_count": 2,
                    "media_folder": "C:/exports/项目群/聊天图片与视频",
                    "warnings": [],
                },
            ),
        )

        detail = format_export_history_detail(record)
        self.assertIn("导出图片/视频", detail)
        self.assertIn("图片/视频数量", detail)

    def test_relay_package_record_can_rebuild_to_folder_import(self):
        record = ExportHistoryRecord(
            export_kind="relay_package",
            title="周报发送包",
            export_folder="C:/exports/周报发送包",
            exported_count=4,
            detail_json=serialize_export_detail(
                "relay_package",
                {
                    "source_session": "项目群",
                    "package_name": "周报发送包",
                    "target_folder": "C:/exports",
                    "item_count": 4,
                },
                {
                    "package_folder": "C:/exports/周报发送包",
                    "item_count": 4,
                    "message_count": 2,
                    "file_count": 2,
                    "manifest_path": "C:/exports/周报发送包/发送清单.xlsx",
                    "files_folder": "C:/exports/周报发送包/素材文件",
                },
            ),
        )

        kind, payload = rebuild_export_request(record)
        detail = format_export_history_detail(record)
        self.assertEqual(kind, "relay_package")
        self.assertEqual(payload, "C:/exports/周报发送包")
        self.assertIn("发送清单", detail)
        self.assertIn("文件目录", detail)

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
