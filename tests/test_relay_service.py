from __future__ import annotations

import importlib.util
import os
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from pyweixin_gui.models import (
    RelayCollectFilesRequest,
    RelayCollectMediaRequest,
    RelayCollectMode,
    RelayCollectTextRequest,
    RelayItemType,
    RelayPackageExportRequest,
    RelayPackageRow,
    RelayRouteRow,
    RelaySendRequest,
    RelayValidationRequest,
    RuntimeOptions,
)
from pyweixin_gui.relay_service import RelayService


class FakeAdapter:
    def __init__(self):
        self.called_recent_text = False
        self.called_recent_files = False

    def dump_chat_history(self, session_name, number, options):
        return ["较新的消息", "较早的消息"], ["今天 10:01", "今天 10:00"]

    def dump_recent_chat_history(self, session_name, recent_range, number, options):
        self.called_recent_text = True
        return ["今天的消息"], ["今天 09:00"]

    def dump_chat_history_items(self, session_name, number, options, recent_range=None):
        return [
            {"sender": "张三", "timestamp": "今天 10:01", "content": "张三的消息"},
            {"sender": "李四", "timestamp": "今天 10:00", "content": "李四的消息"},
        ]

    def save_chat_files(self, session_name, number, target_folder, options):
        folder = Path(target_folder)
        folder.mkdir(parents=True, exist_ok=True)
        (folder / "报价单.pdf").write_text("demo", encoding="utf-8")
        (folder / "产品图.png").write_text("png", encoding="utf-8")
        return [str(folder / "报价单.pdf"), str(folder / "产品图.png")]

    def save_recent_chat_files(self, session_name, recent_range, number, target_folder, options):
        self.called_recent_files = True
        folder = Path(target_folder)
        folder.mkdir(parents=True, exist_ok=True)
        (folder / "今天文件.pdf").write_text("demo", encoding="utf-8")
        return [str(folder / "今天文件.pdf")]

    def list_chat_file_items(self, session_name, number, options, recent_range=None):
        folder = Path(self.temp_root) / "sources"
        folder.mkdir(parents=True, exist_ok=True)
        first = folder / "张三文件.pdf"
        second = folder / "李四文件.pdf"
        first.write_text("demo", encoding="utf-8")
        second.write_text("demo", encoding="utf-8")
        return [
            {"sender": "张三", "timestamp": "今天 10:01", "source_path": str(first), "name": first.name},
            {"sender": "李四", "timestamp": "今天 10:00", "source_path": str(second), "name": second.name},
        ]

    def validate_session(self, session_name, options):
        if session_name == "找不到":
            raise RuntimeError("未找到会话")

    def save_chat_media(self, session_name, number, target_folder, options):
        folder = Path(target_folder)
        folder.mkdir(parents=True, exist_ok=True)
        (folder / "截图.png").write_text("png", encoding="utf-8")
        (folder / "演示视频.mp4").write_text("mp4", encoding="utf-8")
        return [str(folder / "截图.png"), str(folder / "演示视频.mp4")]

    def send_relay_item(self, target_session, item, options):
        if target_session == "发送失败":
            raise RuntimeError("发送失败")

    def map_runtime_exception(self, exc):
        class UiError:
            title = "错误"
            message = str(exc)

        return UiError()


class RelayServiceTestCase(unittest.TestCase):
    def setUp(self) -> None:
        self.adapter = FakeAdapter()
        self.service = RelayService(self.adapter)
        self.options = RuntimeOptions()
        self.tempdir = TemporaryDirectory()
        self.adapter.temp_root = self.tempdir.name
        self.env_patch = patch.dict(os.environ, {"AUTOWECHAT_HOME": self.tempdir.name})
        self.env_patch.start()

    def tearDown(self) -> None:
        self.env_patch.stop()
        self.tempdir.cleanup()

    def test_collect_text_rows(self):
        result = self.service.collect_text_rows(RelayCollectTextRequest(source_session="上游A", message_limit=2), self.options)
        self.assertEqual(len(result.rows), 2)
        self.assertEqual(result.rows[0].item_type, RelayItemType.TEXT)
        self.assertEqual(result.rows[0].sequence, 1)
        self.assertEqual(result.rows[0].content, "较早的消息")
        self.assertEqual(result.rows[1].sequence, 2)
        self.assertEqual(result.rows[1].content, "较新的消息")

    def test_collect_text_rows_skips_media_placeholders(self):
        class PlaceholderAdapter(FakeAdapter):
            def dump_chat_history(self, session_name, number, options):
                return ["图片", "正常文本", "视频"], ["今天 10:02", "今天 10:01", "今天 10:00"]

        service = RelayService(PlaceholderAdapter())
        result = service.collect_text_rows(RelayCollectTextRequest(source_session="上游A", message_limit=3), self.options)
        self.assertEqual(len(result.rows), 1)
        self.assertEqual(result.rows[0].content, "正常文本")
        self.assertIn("已自动跳过", result.warning)

    def test_collect_file_rows(self):
        result = self.service.collect_file_rows(RelayCollectFilesRequest(source_session="上游A", file_limit=5), self.options)
        self.assertEqual(len(result.rows), 2)
        self.assertIn(result.rows[0].item_type, {RelayItemType.FILE, RelayItemType.IMAGE})

    def test_collect_text_rows_by_period(self):
        result = self.service.collect_text_rows(
            RelayCollectTextRequest(source_session="上游A", message_limit=20, collect_mode=RelayCollectMode.PERIOD),
            self.options,
        )
        self.assertTrue(self.adapter.called_recent_text)
        self.assertEqual(result.rows[0].content, "今天的消息")

    def test_collect_text_rows_with_sender_filter(self):
        result = self.service.collect_text_rows(
            RelayCollectTextRequest(source_session="上游A", message_limit=20, sender_names="张三"),
            self.options,
        )
        self.assertEqual(len(result.rows), 1)
        self.assertEqual(result.rows[0].content, "张三的消息")

    def test_collect_file_rows_by_period(self):
        result = self.service.collect_file_rows(
            RelayCollectFilesRequest(source_session="上游A", file_limit=5, collect_mode=RelayCollectMode.PERIOD),
            self.options,
        )
        self.assertTrue(self.adapter.called_recent_files)
        self.assertEqual(result.rows[0].content, "今天文件.pdf")

    def test_collect_file_rows_with_sender_filter(self):
        result = self.service.collect_file_rows(
            RelayCollectFilesRequest(source_session="上游A", file_limit=5, sender_names="李四"),
            self.options,
        )
        self.assertEqual(len(result.rows), 1)
        self.assertEqual(result.rows[0].content, "李四文件.pdf")

    def test_collect_media_rows(self):
        result = self.service.collect_media_rows(RelayCollectMediaRequest(source_session="上游A", media_limit=5), self.options)
        self.assertEqual(len(result.rows), 2)
        self.assertEqual({row.item_type for row in result.rows}, {RelayItemType.IMAGE, RelayItemType.FILE})

    def test_validate_routes(self):
        rows = [
            RelayRouteRow(downstream_session="下游B"),
            RelayRouteRow(downstream_session="找不到"),
        ]
        result = self.service.validate_routes(RelayValidationRequest(source_session="上游A", route_rows=rows), self.options)
        self.assertEqual(result.checked_count, 2)
        self.assertEqual(result.found_count, 1)
        self.assertEqual(result.missing_count, 1)

    def test_send_package_in_test_mode(self):
        rows = [
            RelayPackageRow(sequence=1, item_type=RelayItemType.TEXT, content="文本1"),
            RelayPackageRow(sequence=2, item_type=RelayItemType.FILE, file_path=__file__, content="test"),
        ]
        result = self.service.send_package(RelaySendRequest(source_session="上游A", package_rows=rows, test_only=True), self.options)
        self.assertTrue(result.test_only)
        self.assertEqual(result.target_count, 1)
        self.assertEqual(result.success_count, 1)

    def test_send_package_to_routes(self):
        rows = [RelayPackageRow(sequence=1, item_type=RelayItemType.TEXT, content="文本1")]
        routes = [
            RelayRouteRow(downstream_session="下游B"),
            RelayRouteRow(downstream_session="发送失败"),
        ]
        result = self.service.send_package(
            RelaySendRequest(source_session="上游A", package_rows=rows, route_rows=routes, test_only=False),
            self.options,
        )
        self.assertEqual(result.target_count, 2)
        self.assertEqual(result.success_count, 1)
        self.assertEqual(result.failure_count, 1)

    def test_load_export_folder_rows(self):
        export_root = Path(self.tempdir.name) / "导出目录"
        export_root.mkdir(parents=True, exist_ok=True)
        (export_root / "export-summary.json").write_text('{"session_name":"上游A"}', encoding="utf-8")
        (export_root / "messages.json").write_text(
            '[{"index":1,"timestamp":"今天 10:00","message":"文本1"},{"index":2,"timestamp":"今天 10:01","message":"文本2"}]',
            encoding="utf-8",
        )
        files_dir = export_root / "files"
        files_dir.mkdir(parents=True, exist_ok=True)
        (files_dir / "报价单.pdf").write_text("demo", encoding="utf-8")
        result = self.service.load_export_folder_rows(export_root)
        self.assertEqual(result.source_session, "上游A")
        self.assertEqual(len(result.rows), 3)
        self.assertEqual(result.rows[0].item_type, RelayItemType.TEXT)
        self.assertEqual(result.rows[0].content, "文本2")
        self.assertEqual(result.rows[1].content, "文本1")

    @unittest.skipUnless(importlib.util.find_spec("openpyxl"), "openpyxl not installed")
    def test_export_package_folder_and_reload(self):
        with TemporaryDirectory() as tempdir:
            image_path = Path(tempdir) / "海报.png"
            image_path.write_text("png", encoding="utf-8")
            rows = [
                RelayPackageRow(sequence=1, item_type=RelayItemType.TEXT, content="第一条", collected_at="今天 09:00"),
                RelayPackageRow(sequence=2, item_type=RelayItemType.IMAGE, content=image_path.name, file_path=str(image_path), collected_at="今天 09:01"),
            ]
            request = RelayPackageExportRequest(
                source_session="上游A",
                package_name="测试转发包",
                target_folder=tempdir,
                package_rows=rows,
            )
            exported = self.service.export_package_folder(request)
            self.assertTrue(Path(exported.package_folder).exists())
            self.assertTrue(Path(exported.manifest_path).exists())
            reloaded = self.service.load_folder_rows(exported.package_folder)
            self.assertEqual(reloaded.source_session, "上游A")
            self.assertEqual(len(reloaded.rows), 2)
            self.assertEqual(reloaded.rows[0].content, "第一条")
            self.assertEqual(reloaded.rows[1].item_type, RelayItemType.IMAGE)
            self.assertTrue(Path(reloaded.rows[1].file_path).is_file())

    def test_keep_latest_file_rows(self):
        with TemporaryDirectory() as tempdir:
            folder = Path(tempdir)
            old_file = folder / "报价单(1).pdf"
            new_file = folder / "报价单.pdf"
            old_file.write_text("old", encoding="utf-8")
            new_file.write_text("new", encoding="utf-8")
            fixed_timestamp = 1_710_000_000
            os.utime(old_file, (fixed_timestamp, fixed_timestamp))
            os.utime(new_file, (fixed_timestamp, fixed_timestamp))
            rows = [
                RelayPackageRow(sequence=1, item_type=RelayItemType.TEXT, content="文本1"),
                RelayPackageRow(sequence=2, item_type=RelayItemType.FILE, file_path=str(old_file), content=old_file.name),
                RelayPackageRow(sequence=3, item_type=RelayItemType.FILE, file_path=str(new_file), content=new_file.name),
            ]
            updated = RelayService.keep_latest_file_rows(rows)
            self.assertEqual(len(updated), 2)
            self.assertEqual(updated[0].content, "文本1")
            self.assertEqual(updated[1].content, new_file.name)


if __name__ == "__main__":
    unittest.main()
