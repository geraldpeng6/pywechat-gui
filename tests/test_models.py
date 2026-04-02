from __future__ import annotations

import unittest

from pyweixin_gui.models import FileBatchRow, MessageBatchRow, RelayCollectFilesRequest, RelayCollectMode, RelayCollectTextRequest, RelayItemType, RelayPackageRow, RelayRecentRange, relay_template_from_json, relay_template_to_json


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

    def test_relay_package_row_keeps_enum_item_type(self):
        row = RelayPackageRow.from_mapping(
            {
                "sequence": 2,
                "item_type": RelayItemType.FILE,
                "content": "报价单.pdf",
                "file_path": "C:/demo/报价单.pdf",
            }
        )
        self.assertEqual(row.item_type, RelayItemType.FILE)

    def test_relay_template_roundtrip_keeps_file_and_image_types(self):
        payload = relay_template_to_json(
            source_session="上游A",
            package_name="测试任务",
            package_rows=[
                RelayPackageRow(sequence=1, item_type=RelayItemType.TEXT, content="你好"),
                RelayPackageRow(sequence=2, item_type=RelayItemType.FILE, content="报价单.pdf", file_path="C:/demo/报价单.pdf"),
                RelayPackageRow(sequence=3, item_type=RelayItemType.IMAGE, content="海报.png", file_path="C:/demo/海报.png"),
            ],
            route_rows=[],
        )
        loaded = relay_template_from_json(payload)
        package_rows = loaded["package_rows"]
        self.assertEqual(package_rows[0].item_type, RelayItemType.TEXT)
        self.assertEqual(package_rows[1].item_type, RelayItemType.FILE)
        self.assertEqual(package_rows[2].item_type, RelayItemType.IMAGE)

    def test_collect_text_request_accepts_period_mode(self):
        request = RelayCollectTextRequest(
            source_session="上游A",
            message_limit=20,
            collect_mode=RelayCollectMode.PERIOD,
            recent_range=RelayRecentRange.WEEK,
        )
        self.assertEqual(request.validate(), {})

    def test_collect_files_request_rejects_invalid_recent_range_in_period_mode(self):
        request = RelayCollectFilesRequest(
            source_session="上游A",
            file_limit=10,
            collect_mode="period",
            recent_range="invalid",
        )
        self.assertIn("recent_range", request.validate())

    def test_collect_request_sender_name_list(self):
        request = RelayCollectTextRequest(source_session="上游A", sender_names="张三| 李四 |")
        self.assertEqual(request.sender_name_list(), ["张三", "李四"])


if __name__ == "__main__":
    unittest.main()
