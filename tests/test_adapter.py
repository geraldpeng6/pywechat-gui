from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from pyweixin_gui.adapter import PyWeixinAdapter, UnsupportedPlatformError
from pyweixin_gui.models import RuntimeOptions


class AdapterTestCase(unittest.TestCase):
    def test_non_windows_environment_has_clear_status(self):
        adapter = PyWeixinAdapter()
        with patch("platform.system", return_value="Darwin"), patch("platform.release", return_value="24.0"), patch(
            "platform.python_version", return_value="3.9.0"
        ):
            status = adapter.inspect_environment()
        self.assertEqual(status.login_status, "当前平台不受支持")
        self.assertIn("Windows", status.advice[0])

    def test_loading_pyweixin_outside_windows_raises_supported_error(self):
        adapter = PyWeixinAdapter()
        with patch("platform.system", return_value="Linux"):
            with self.assertRaises(UnsupportedPlatformError):
                adapter._load_pyweixin()

    def test_dump_sessions_matches_pyweixin_signature_and_shape(self):
        class FakeMessages:
            @staticmethod
            def dump_sessions(chat_only=False, is_maximize=None, close_weixin=None):
                return [("项目群", "昨天", "已收到"), ("客户A", "今天", "请报价")]

        class FakePyWeixin:
            Messages = FakeMessages

        adapter = PyWeixinAdapter()
        with patch.object(adapter, "_load_pyweixin", return_value=FakePyWeixin()):
            rows = adapter.dump_sessions(chatted_only=True, options=RuntimeOptions())
        self.assertEqual(rows[0].session_name, "项目群")
        self.assertEqual(rows[1].last_message, "请报价")

    def test_dump_groups_accepts_name_list_result(self):
        class FakeContacts:
            @staticmethod
            def get_groups_info(is_maximize=None, close_weixin=None):
                return ["项目群", "通知群"]

        class FakePyWeixin:
            Contacts = FakeContacts

        adapter = PyWeixinAdapter()
        with patch.object(adapter, "_load_pyweixin", return_value=FakePyWeixin()):
            rows = adapter.dump_groups(RuntimeOptions())
        self.assertEqual(rows[0].group_name, "项目群")
        self.assertEqual(rows[0].member_count, "")

    def test_export_videos_recovers_from_upstream_unbound_local_error(self):
        class FakeFiles:
            @staticmethod
            def export_videos(year=None, month=None, target_folder=None):
                raise UnboundLocalError("cannot access local variable 'exported_videos' where it is not associated with a value")

        class FakePyWeixin:
            Files = FakeFiles

        adapter = PyWeixinAdapter()
        with tempfile.TemporaryDirectory() as tempdir:
            exported = Path(tempdir) / "sample.mp4"
            exported.write_text("demo", encoding="utf-8")
            with patch.object(adapter, "_load_pyweixin", return_value=FakePyWeixin()):
                rows = adapter.export_videos("2026", "04", tempdir)
        self.assertEqual(rows, [str(exported)])


if __name__ == "__main__":
    unittest.main()
