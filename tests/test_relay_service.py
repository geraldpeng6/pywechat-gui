from __future__ import annotations

import os
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from pyweixin_gui.models import (
    RelayCollectFilesRequest,
    RelayCollectTextRequest,
    RelayItemType,
    RelayPackageRow,
    RelayRouteRow,
    RelaySendRequest,
    RelayValidationRequest,
    RuntimeOptions,
)
from pyweixin_gui.relay_service import RelayService


class FakeAdapter:
    def dump_chat_history(self, session_name, number, options):
        return ["文本1", "文本2"], ["今天 10:00", "今天 10:01"]

    def save_chat_files(self, session_name, number, target_folder, options):
        folder = Path(target_folder)
        folder.mkdir(parents=True, exist_ok=True)
        (folder / "报价单.pdf").write_text("demo", encoding="utf-8")
        (folder / "产品图.png").write_text("png", encoding="utf-8")
        return [str(folder / "报价单.pdf"), str(folder / "产品图.png")]

    def validate_session(self, session_name, options):
        if session_name == "找不到":
            raise RuntimeError("未找到会话")

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
        self.service = RelayService(FakeAdapter())
        self.options = RuntimeOptions()
        self.tempdir = TemporaryDirectory()
        self.env_patch = patch.dict(os.environ, {"PYWEIXIN_GUI_HOME": self.tempdir.name})
        self.env_patch.start()

    def tearDown(self) -> None:
        self.env_patch.stop()
        self.tempdir.cleanup()

    def test_collect_text_rows(self):
        result = self.service.collect_text_rows(RelayCollectTextRequest(source_session="上游A", message_limit=2), self.options)
        self.assertEqual(len(result.rows), 2)
        self.assertEqual(result.rows[0].item_type, RelayItemType.TEXT)

    def test_collect_file_rows(self):
        result = self.service.collect_file_rows(RelayCollectFilesRequest(source_session="上游A", file_limit=5), self.options)
        self.assertEqual(len(result.rows), 2)
        self.assertIn(result.rows[0].item_type, {RelayItemType.FILE, RelayItemType.IMAGE})

    def test_validate_routes(self):
        rows = [
            RelayRouteRow(upstream_session="上游A", downstream_session="下游B"),
            RelayRouteRow(upstream_session="上游A", downstream_session="找不到"),
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
            RelayRouteRow(upstream_session="上游A", downstream_session="下游B"),
            RelayRouteRow(upstream_session="上游A", downstream_session="发送失败"),
        ]
        result = self.service.send_package(
            RelaySendRequest(source_session="上游A", package_rows=rows, route_rows=routes, test_only=False),
            self.options,
        )
        self.assertEqual(result.target_count, 2)
        self.assertEqual(result.success_count, 1)
        self.assertEqual(result.failure_count, 1)


if __name__ == "__main__":
    unittest.main()
