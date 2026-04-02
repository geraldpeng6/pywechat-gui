from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from pyweixin_gui.models import ResourceExportKind, ResourceExportRequest, RuntimeOptions
from pyweixin_gui.resource_export_service import ResourceExportService


class FakeAdapter:
    def export_recent_files(self, target_folder, options):
        folder = Path(target_folder)
        folder.mkdir(parents=True, exist_ok=True)
        (folder / "recent.docx").write_text("x", encoding="utf-8")
        return [str(folder / "recent.docx")]

    def export_wxfiles(self, year, month, target_folder):
        folder = Path(target_folder)
        folder.mkdir(parents=True, exist_ok=True)
        (folder / "wx.xlsx").write_text("x", encoding="utf-8")
        return [str(folder / "wx.xlsx")]

    def export_videos(self, year, month, target_folder):
        folder = Path(target_folder)
        folder.mkdir(parents=True, exist_ok=True)
        (folder / "video.mp4").write_text("x", encoding="utf-8")
        return [str(folder / "video.mp4")]


class ResourceExportServiceTestCase(unittest.TestCase):
    def test_recent_files_export(self):
        service = ResourceExportService(FakeAdapter())
        with tempfile.TemporaryDirectory() as tempdir:
            request = ResourceExportRequest(export_kind=ResourceExportKind.RECENT_FILES, target_folder=tempdir, year="2026")
            result = service.run_export(request, RuntimeOptions())
            self.assertEqual(result.exported_count, 1)
            self.assertTrue(Path(result.summary_txt).exists())

    def test_video_export(self):
        service = ResourceExportService(FakeAdapter())
        with tempfile.TemporaryDirectory() as tempdir:
            request = ResourceExportRequest(export_kind=ResourceExportKind.VIDEOS, target_folder=tempdir, year="2026", month="03")
            result = service.run_export(request, RuntimeOptions())
            self.assertEqual(result.exported_count, 1)
            self.assertIn("video.mp4", result.exported_paths[0])

    def test_string_export_kind_is_normalized(self):
        service = ResourceExportService(FakeAdapter())
        with tempfile.TemporaryDirectory() as tempdir:
            request = ResourceExportRequest(export_kind="recent_files", target_folder=tempdir, year="2026")
            result = service.run_export(request, RuntimeOptions())
            self.assertEqual(result.export_kind, ResourceExportKind.RECENT_FILES)
            self.assertTrue(Path(result.summary_txt).name.startswith("recent_files-summary-"))

    def test_invalid_export_kind_is_rejected(self):
        request = ResourceExportRequest(export_kind="invalid_kind", target_folder="C:/exports", year="2026")
        self.assertIn("export_kind", request.validate())


if __name__ == "__main__":
    unittest.main()
