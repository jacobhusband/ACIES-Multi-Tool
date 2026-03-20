import os
import shutil
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest.mock import patch


def _ensure_google_genai_stub():
    try:
        from google import genai as _genai  # noqa: F401
        from google.genai import types as _types  # noqa: F401
        return
    except Exception:
        google_module = sys.modules.get("google")
        if google_module is None:
            google_module = types.ModuleType("google")
            google_module.__path__ = []
            sys.modules["google"] = google_module

        genai_module = types.ModuleType("google.genai")
        genai_types_module = types.ModuleType("google.genai.types")
        genai_module.types = genai_types_module
        google_module.genai = genai_module

        sys.modules["google.genai"] = genai_module
        sys.modules["google.genai.types"] = genai_types_module


def _ensure_webview_stub():
    try:
        import webview  # noqa: F401
        return
    except Exception:
        webview_module = types.ModuleType("webview")
        webview_module.windows = []
        webview_module.FOLDER_DIALOG = "folder"
        webview_module.OPEN_DIALOG = "open"
        webview_module.FileDialog = types.SimpleNamespace(SAVE="save")
        webview_module.create_window = lambda *args, **kwargs: None
        webview_module.start = lambda *args, **kwargs: None
        sys.modules["webview"] = webview_module


def _ensure_dotenv_stub():
    try:
        from dotenv import load_dotenv as _load_dotenv  # noqa: F401
        return
    except Exception:
        dotenv_module = types.ModuleType("dotenv")
        dotenv_module.load_dotenv = lambda *args, **kwargs: False
        sys.modules["dotenv"] = dotenv_module


def _ensure_requests_stub():
    try:
        import requests  # noqa: F401
        return
    except Exception:
        requests_module = types.ModuleType("requests")
        requests_module.get = lambda *args, **kwargs: None
        sys.modules["requests"] = requests_module


def _ensure_pydantic_stub():
    try:
        from pydantic import BaseModel as _BaseModel, Field as _Field  # noqa: F401
        return
    except Exception:
        pydantic_module = types.ModuleType("pydantic")

        class BaseModel:
            pass

        def Field(*args, **kwargs):
            if args:
                return args[0]
            return kwargs.get("default")

        pydantic_module.BaseModel = BaseModel
        pydantic_module.Field = Field
        sys.modules["pydantic"] = pydantic_module


def _ensure_pil_stub():
    try:
        from PIL import Image as _image  # noqa: F401
        return
    except Exception:
        pil_module = types.ModuleType("PIL")
        pil_image_module = types.ModuleType("PIL.Image")
        pil_module.Image = pil_image_module
        sys.modules["PIL"] = pil_module
        sys.modules["PIL.Image"] = pil_image_module


def _ensure_openpyxl_stub():
    try:
        import openpyxl  # noqa: F401
        from openpyxl.worksheet.copier import WorksheetCopy as _WorksheetCopy  # noqa: F401
        return
    except Exception:
        openpyxl_module = types.ModuleType("openpyxl")
        worksheet_module = types.ModuleType("openpyxl.worksheet")
        copier_module = types.ModuleType("openpyxl.worksheet.copier")

        class WorksheetCopy:
            pass

        copier_module.WorksheetCopy = WorksheetCopy
        worksheet_module.copier = copier_module
        openpyxl_module.worksheet = worksheet_module

        sys.modules["openpyxl"] = openpyxl_module
        sys.modules["openpyxl.worksheet"] = worksheet_module
        sys.modules["openpyxl.worksheet.copier"] = copier_module


_ensure_google_genai_stub()
_ensure_webview_stub()
_ensure_dotenv_stub()
_ensure_requests_stub()
_ensure_pydantic_stub()
_ensure_pil_stub()
_ensure_openpyxl_stub()

import main as main_module
from main import Api


REPO_ROOT = Path(__file__).resolve().parents[1]
INDEX_HTML_PATH = REPO_ROOT / "index.html"
SCRIPT_JS_PATH = REPO_ROOT / "script.js"


class BackupDrawingsBackendTests(unittest.TestCase):
    def setUp(self):
        self.api = Api.__new__(Api)

    def _settings(self, disciplines=None):
        return {"discipline": disciplines or ["Electrical"]}

    def _write_file(self, path, content):
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(content, encoding="utf-8")

    def test_backup_project_drawings_copies_configured_disciplines_and_xrefs(self):
        with tempfile.TemporaryDirectory(prefix="acies-backup-drawings-") as temp_dir:
            project_root = Path(temp_dir) / "260243 BofA - Eastport Plaza"
            self._write_file(project_root / "Electrical" / "Sheets" / "power.dwg", "power")
            self._write_file(project_root / "Mechanical" / "hvac.dwg", "hvac")
            self._write_file(project_root / "Xrefs" / "arch.dwg", "arch")

            with patch.object(self.api, "get_user_settings", return_value=self._settings(["Electrical", "Mechanical"])), patch.object(
                self.api,
                "_build_backup_drawings_timestamp",
                return_value="20260319_150812",
            ):
                result = self.api.backup_project_drawings(str(project_root))

            archive_root = project_root / "Archive"
            archive_path = archive_root / "20260319_150812"

            self.assertEqual("success", result["status"])
            self.assertEqual(os.path.normpath(str(archive_root)), os.path.normpath(result["archiveRootPath"]))
            self.assertEqual(os.path.normpath(str(archive_path)), os.path.normpath(result["archivePath"]))
            self.assertEqual(["Electrical", "Mechanical", "Xrefs"], result["copiedFolders"])
            self.assertEqual([], result["missingSourceFolders"])
            self.assertEqual(3, result["copiedFileCount"])
            self.assertTrue((archive_path / "Electrical" / "Sheets" / "power.dwg").exists())
            self.assertTrue((archive_path / "Mechanical" / "hvac.dwg").exists())
            self.assertTrue((archive_path / "Xrefs" / "arch.dwg").exists())

    def test_backup_project_drawings_preserves_workroom_resolution_details(self):
        with tempfile.TemporaryDirectory(prefix="acies-backup-drawings-") as temp_dir:
            project_root = Path(temp_dir) / "260243 BofA - Eastport Plaza"
            self._write_file(project_root / "Electrical" / "power.dwg", "power")
            self._write_file(project_root / "Xrefs" / "arch.dwg", "arch")

            with patch.object(self.api, "get_user_settings", return_value=self._settings()), patch.object(
                self.api,
                "_build_backup_drawings_timestamp",
                return_value="20260319_150812",
            ), patch.object(
                self.api,
                "_resolve_copy_project_source_path",
                return_value={
                    "status": "success",
                    "path": os.path.normpath(str(project_root)),
                    "resolvedFromWorkroom": True,
                    "resolutionMode": "project_id_ancestor",
                    "workroomProjectPath": os.path.normpath(str(project_root / "Electrical")),
                },
            ):
                result = self.api.backup_project_drawings(None, {"source": "workroom"})

            self.assertEqual("success", result["status"])
            self.assertTrue(result["resolvedFromWorkroom"])
            self.assertEqual("project_id_ancestor", result["resolutionMode"])
            self.assertEqual(
                os.path.normpath(str(project_root / "Electrical")),
                os.path.normpath(result["workroomProjectPath"]),
            )

    def test_backup_project_drawings_surfaces_manual_selection_required_for_workroom_resolution(self):
        with patch.object(self.api, "get_user_settings", return_value=self._settings()), patch.object(
            self.api,
            "_resolve_copy_project_source_path",
            return_value={
                "status": "error",
                "code": "manual_selection_required",
                "message": "Could not auto-resolve project folder from Project Workroom. Please select it manually.",
            },
        ):
            result = self.api.backup_project_drawings(None, {"source": "workroom"})

        self.assertEqual("error", result["status"])
        self.assertEqual("manual_selection_required", result["code"])

    def test_backup_project_drawings_reports_missing_folders_without_creating_placeholders(self):
        with tempfile.TemporaryDirectory(prefix="acies-backup-drawings-") as temp_dir:
            project_root = Path(temp_dir) / "260243 BofA - Eastport Plaza"
            self._write_file(project_root / "Electrical" / "power.dwg", "power")

            with patch.object(self.api, "get_user_settings", return_value=self._settings(["Electrical", "Mechanical"])), patch.object(
                self.api,
                "_build_backup_drawings_timestamp",
                return_value="20260319_150812",
            ):
                result = self.api.backup_project_drawings(str(project_root))

            archive_path = project_root / "Archive" / "20260319_150812"
            self.assertEqual("success", result["status"])
            self.assertEqual(["Mechanical", "Xrefs"], result["missingSourceFolders"])
            self.assertFalse((archive_path / "Mechanical").exists())
            self.assertFalse((archive_path / "Xrefs").exists())
            self.assertTrue((archive_path / "Electrical" / "power.dwg").exists())

    def test_backup_project_drawings_removes_empty_timestamp_folder_when_nothing_is_backed_up(self):
        with tempfile.TemporaryDirectory(prefix="acies-backup-drawings-") as temp_dir:
            project_root = Path(temp_dir) / "260243 BofA - Eastport Plaza"
            project_root.mkdir(parents=True)

            with patch.object(self.api, "get_user_settings", return_value=self._settings(["Electrical", "Mechanical"])), patch.object(
                self.api,
                "_build_backup_drawings_timestamp",
                return_value="20260319_150812",
            ):
                result = self.api.backup_project_drawings(str(project_root))

            archive_path = project_root / "Archive" / "20260319_150812"
            self.assertEqual("error", result["status"])
            self.assertEqual("nothing_to_backup", result["code"])
            self.assertFalse(archive_path.exists())

    def test_backup_project_drawings_tracks_failed_files_without_losing_successful_copies(self):
        with tempfile.TemporaryDirectory(prefix="acies-backup-drawings-") as temp_dir:
            project_root = Path(temp_dir) / "260243 BofA - Eastport Plaza"
            self._write_file(project_root / "Electrical" / "good.dwg", "good")
            self._write_file(project_root / "Electrical" / "bad.dwg", "bad")
            self._write_file(project_root / "Xrefs" / "xref.dwg", "xref")

            original_copy2 = shutil.copy2

            def flaky_copy2(src, dst, *args, **kwargs):
                if os.path.basename(src).lower() == "bad.dwg":
                    raise PermissionError("locked by AutoCAD")
                return original_copy2(src, dst, *args, **kwargs)

            with patch.object(self.api, "get_user_settings", return_value=self._settings()), patch.object(
                self.api,
                "_build_backup_drawings_timestamp",
                return_value="20260319_150812",
            ), patch.object(main_module.shutil, "copy2", side_effect=flaky_copy2):
                result = self.api.backup_project_drawings(str(project_root))

            archive_path = project_root / "Archive" / "20260319_150812"
            self.assertEqual("success", result["status"])
            self.assertEqual(1, result["failedFileCount"])
            self.assertEqual(2, result["copiedFileCount"])
            self.assertTrue((archive_path / "Electrical" / "good.dwg").exists())
            self.assertFalse((archive_path / "Electrical" / "bad.dwg").exists())
            self.assertTrue((archive_path / "Xrefs" / "xref.dwg").exists())
            self.assertIn("locked by AutoCAD", result["failedFiles"][0]["error"])

    def test_backup_project_drawings_uses_unique_suffix_when_timestamp_folder_already_exists(self):
        with tempfile.TemporaryDirectory(prefix="acies-backup-drawings-") as temp_dir:
            project_root = Path(temp_dir) / "260243 BofA - Eastport Plaza"
            self._write_file(project_root / "Electrical" / "power.dwg", "power")
            self._write_file(project_root / "Xrefs" / "arch.dwg", "arch")
            (project_root / "Archive" / "20260319_150812").mkdir(parents=True)

            with patch.object(self.api, "get_user_settings", return_value=self._settings()), patch.object(
                self.api,
                "_build_backup_drawings_timestamp",
                return_value="20260319_150812",
            ):
                result = self.api.backup_project_drawings(str(project_root))

            archive_path = project_root / "Archive" / "20260319_150812 (2)"
            self.assertEqual("success", result["status"])
            self.assertEqual(os.path.normpath(str(archive_path)), os.path.normpath(result["archivePath"]))
            self.assertTrue((archive_path / "Electrical" / "power.dwg").exists())


class BackupDrawingsUiTests(unittest.TestCase):
    def test_backup_drawings_markup_exists(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")

        self.assertIn('id="toolBackupDrawings"', html)
        self.assertIn("Backup Drawings", html)
        self.assertIn("Archive\\current-datetime", html)

    def test_backup_drawings_script_wiring_exists(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn('"toolBackupDrawings"', text)
        self.assertIn('window.updateToolStatus("toolBackupDrawings", "Select project folder...");', text)
        self.assertIn('window.updateToolStatus("toolBackupDrawings", "Resolving project folder...");', text)
        self.assertIn('window.updateToolStatus("toolBackupDrawings", "Creating archive backup...");', text)
        self.assertIn("window.pywebview.api.backup_project_drawings(", text)
        self.assertIn("result?.missingSourceFolders", text)
        self.assertIn('toast(`Missing folders: ${missingFolders.join(", ")}`', text)
        self.assertIn('toast(`Failed files: ${failedPreview.join(", ")}${suffix}`, 9000);', text)
        self.assertIn("await window.pywebview.api.open_path(result.archivePath);", text)
        self.assertIn('window.updateToolStatus("toolBackupDrawings", "DONE");', text)


if __name__ == "__main__":
    unittest.main()
