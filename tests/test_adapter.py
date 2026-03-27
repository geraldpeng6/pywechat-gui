from __future__ import annotations

import unittest
from unittest.mock import patch

from pyweixin_gui.adapter import PyWeixinAdapter, UnsupportedPlatformError


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


if __name__ == "__main__":
    unittest.main()
