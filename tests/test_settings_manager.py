from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from pyweixin_gui.models import AppSettings
from pyweixin_gui.settings_manager import SettingsManager


class SettingsManagerTestCase(unittest.TestCase):
    def test_roundtrip(self):
        with tempfile.TemporaryDirectory() as tempdir:
            path = Path(tempdir) / "settings.json"
            manager = SettingsManager(path)
            settings = AppSettings(is_maximize=True, search_pages=8, send_delay=0.5, theme="light", first_run_risk_ack=True)
            manager.save(settings)
            loaded = manager.load()
            self.assertTrue(loaded.is_maximize)
            self.assertEqual(loaded.search_pages, 8)
            self.assertEqual(loaded.send_delay, 0.5)
            self.assertTrue(loaded.first_run_risk_ack)


if __name__ == "__main__":
    unittest.main()
