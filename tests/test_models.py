from __future__ import annotations

import unittest

from pyweixin_gui.models import FileBatchRow, MessageBatchRow, RelayItemType, RelayPackageRow, relay_template_from_json, relay_template_to_json


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


if __name__ == "__main__":
    unittest.main()
