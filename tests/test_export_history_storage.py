from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from pyweixin_gui.models import ExportHistoryRecord
from pyweixin_gui.storage import AppStorage


class ExportHistoryStorageTestCase(unittest.TestCase):
    def setUp(self) -> None:
        self.tempdir = tempfile.TemporaryDirectory()
        self.storage = AppStorage(Path(self.tempdir.name) / "app.db")

    def tearDown(self) -> None:
        self.tempdir.cleanup()

    def test_export_record_roundtrip(self):
        record = ExportHistoryRecord(
            export_kind="chat",
            title="项目群",
            export_folder="/tmp/demo",
            exported_count=3,
            summary_path="/tmp/demo/export-summary.txt",
            detail_json='{"type":"chat"}',
        )
        saved = self.storage.save_export_record(record)
        listed = self.storage.list_export_records()
        self.assertEqual(saved.id, listed[0].id)
        self.assertEqual(listed[0].title, "项目群")


if __name__ == "__main__":
    unittest.main()
