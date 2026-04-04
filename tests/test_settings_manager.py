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
            settings = AppSettings(
                is_maximize=True,
                search_pages=8,
                send_delay=0.5,
                import_encoding="gbk",
                theme="light",
                history_retention="30d",
                first_run_risk_ack=True,
            )
            manager.save(settings)
            loaded = manager.load()
            self.assertTrue(loaded.is_maximize)
            self.assertEqual(loaded.search_pages, 8)
            self.assertEqual(loaded.send_delay, 0.5)
            self.assertEqual(loaded.import_encoding, "gbk")
            self.assertEqual(loaded.history_retention, "30d")
            self.assertTrue(loaded.first_run_risk_ack)

    def test_load_recovers_from_invalid_json_and_keeps_backup(self):
        with tempfile.TemporaryDirectory() as tempdir:
            path = Path(tempdir) / "settings.json"
            path.write_text("{invalid-json", encoding="utf-8")
            manager = SettingsManager(path)

            loaded = manager.load()

            self.assertEqual(loaded, AppSettings())
            self.assertFalse(path.exists())
            backups = list(Path(tempdir).glob("settings.invalid-*.json"))
            self.assertEqual(len(backups), 1)
            self.assertIn("{invalid-json", backups[0].read_text(encoding="utf-8"))

    def test_load_recovers_from_non_object_json(self):
        with tempfile.TemporaryDirectory() as tempdir:
            path = Path(tempdir) / "settings.json"
            path.write_text('["not-an-object"]', encoding="utf-8")
            manager = SettingsManager(path)

            loaded = manager.load()

            self.assertEqual(loaded, AppSettings())
            backups = list(Path(tempdir).glob("settings.invalid-*.json"))
            self.assertEqual(len(backups), 1)


if __name__ == "__main__":
    unittest.main()
