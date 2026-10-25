"""Protect workbook locations when distributing source, portable, and installed copies."""
from __future__ import annotations

import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import university_tracker_qt as tracker


class InstallerStorageTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix="university_tracker_storage_")
        self.addCleanup(temporary.cleanup)
        self.folder = Path(temporary.name)
        self.application = self.folder / "application"
        self.application.mkdir()
        self.local_data = self.folder / "local-data"
        self.root_patch = patch.object(tracker, "app_root", return_value=self.application)
        self.root_patch.start()
        self.addCleanup(self.root_patch.stop)
        self.environment_patch = patch.dict(os.environ, {"LOCALAPPDATA": str(self.local_data)})
        self.environment_patch.start()
        self.addCleanup(self.environment_patch.stop)

    def install_marker(self, root: Path | None = None):
        (root or self.application).joinpath("installed-mode.txt").write_text("installed", encoding="utf-8")

    def test_source_copy_keeps_existing_local_workbook_location(self):
        self.install_marker()
        with patch.object(sys, "frozen", False, create=True):
            self.assertEqual(tracker.data_root(), self.application)

    def test_portable_executable_keeps_local_workbook_location(self):
        with patch.object(sys, "frozen", True, create=True):
            self.assertEqual(tracker.data_root(), self.application)

    def test_installed_executable_uses_per_user_data(self):
        self.install_marker()
        with patch.object(sys, "frozen", True, create=True):
            self.assertEqual(tracker.data_root(), self.local_data / tracker.APP_NAME)

    def test_missing_localappdata_uses_user_profile_fallback(self):
        self.install_marker()
        with patch.dict(os.environ), patch.object(sys, "frozen", True, create=True), \
                patch.object(Path, "home", return_value=self.folder / "profile"):
            os.environ.pop("LOCALAPPDATA", None)
            self.assertEqual(tracker.data_root(), self.folder / "profile" / "AppData" / "Local" / tracker.APP_NAME)

    def test_installed_save_and_reload_never_write_beside_executable(self):
        self.install_marker()
        with patch.object(sys, "frozen", True, create=True):
            store = tracker.UniversityStore()
            self.assertEqual(store.records, [])
            record = {header: "" for header in tracker.HEADERS}
            record.update({"Country": "Italy", "University": "Test University", "Program Name": "Test Program"})
            store.records.append(record)
            store.save()
            self.assertEqual(tracker.UniversityStore().records, [record])
            self.assertTrue((self.local_data / tracker.APP_NAME / "Excel_Sheets" / tracker.FILE_NAME).is_file())
            self.assertFalse((self.application / "Excel_Sheets").exists())

    def test_new_application_location_reuses_the_installed_workbook(self):
        self.install_marker()
        with patch.object(sys, "frozen", True, create=True):
            store = tracker.UniversityStore()
            record = {header: "" for header in tracker.HEADERS}
            record.update({"Country": "Germany", "University": "Saved University", "Program Name": "Saved Program"})
            store.records.append(record)
            store.save()
            updated_application = self.folder / "updated-application"
            updated_application.mkdir()
            self.install_marker(updated_application)
            with patch.object(tracker, "app_root", return_value=updated_application), \
                    patch.object(tracker, "APP_VERSION", "2.0.0"):
                self.assertEqual(tracker.UniversityStore().records, [record])


if __name__ == "__main__":
    unittest.main()
