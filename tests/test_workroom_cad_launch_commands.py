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

from main import Api


AUTO_SELECTED_FILES_LIST = r"C:\Temp\workroom-dwgs.txt"
AUTOCAD_CORE_PATH = r"C:\Program Files\Autodesk\AutoCAD 2025\accoreconsole.exe"
FALLBACK_STATUS_MESSAGE = "Workroom auto-select unavailable. Opening file picker..."

TOOL_CASES = (
    {
        "method_name": "run_publish_script",
        "tool_id": "toolPublishDwgs",
        "script_name": "PlotDWGs.ps1",
        "settings": {
            "autocadPath": AUTOCAD_CORE_PATH,
            "publishDwgOptions": {
                "autoDetectPaperSize": False,
                "shrinkPercent": 85,
            },
        },
        "expected_args": (
            "-AcadCore",
            AUTOCAD_CORE_PATH,
            "-AutoDetectPaperSize",
            "0",
            "-ShrinkPercent",
            "85",
        ),
    },
    {
        "method_name": "run_freeze_layers_script",
        "tool_id": "toolFreezeLayers",
        "script_name": "FreezeLayersDWGs.ps1",
        "settings": {
            "autocadPath": AUTOCAD_CORE_PATH,
            "freezeLayerOptions": {
                "scanAllLayers": False,
            },
        },
        "expected_args": (
            "-AcadCore",
            AUTOCAD_CORE_PATH,
            "-ScanAllLayers",
            "0",
        ),
    },
    {
        "method_name": "run_thaw_layers_script",
        "tool_id": "toolThawLayers",
        "script_name": "ThawLayersDWGs.ps1",
        "settings": {
            "autocadPath": AUTOCAD_CORE_PATH,
            "thawLayerOptions": {
                "scanAllLayers": False,
            },
        },
        "expected_args": (
            "-AcadCore",
            AUTOCAD_CORE_PATH,
            "-ScanAllLayers",
            "0",
        ),
    },
)


class WorkroomCadLaunchCommandTests(unittest.TestCase):
    def setUp(self):
        self.api = Api.__new__(Api)
        self.api.test_mode = False
        self.api._workroom_cad_file_cache = {}

    def _run_tool(self, case, auto_selection, auto_select_enabled=True):
        captured = []
        with patch.object(self.api, "get_user_settings", return_value=case["settings"]), patch.object(
            self.api,
            "_resolve_workroom_auto_file_selection",
            return_value=auto_selection,
        ), patch.object(
            self.api,
            "_is_workroom_auto_select_enabled",
            return_value=auto_select_enabled,
        ), patch.object(
            self.api,
            "_resolve_workroom_context",
            return_value={
                "source": "workroom",
                "project_path": r"C:\Projects\123456\Sheets",
                "discipline": "Electrical",
                "discipline_source": "launch_context",
            },
        ), patch.object(
            self.api,
            "_run_script_with_progress",
            side_effect=lambda command, tool_id: captured.append((command, tool_id)),
        ), patch.object(self.api, "_notify_tool_status") as notify_mock:
            result = getattr(self.api, case["method_name"])({"source": "workroom"})
        return result, captured, notify_mock

    def test_workroom_cad_tools_pass_files_list_path_as_explicit_argv(self):
        auto_selection = {"files_list_path": AUTO_SELECTED_FILES_LIST}

        for case in TOOL_CASES:
            with self.subTest(method=case["method_name"]):
                result, captured, notify_mock = self._run_tool(case, auto_selection)

                self.assertEqual({"status": "success"}, result)
                self.assertEqual(1, len(captured))

                command, tool_id = captured[0]
                self.assertEqual(case["tool_id"], tool_id)
                self.assertIsInstance(command, list)
                self.assertGreaterEqual(len(command), 6)
                self.assertEqual("powershell.exe", command[0])
                self.assertEqual("-ExecutionPolicy", command[1])
                self.assertEqual("Bypass", command[2])
                self.assertEqual("-File", command[3])
                self.assertTrue(command[4].endswith(case["script_name"]))

                for expected_arg in case["expected_args"]:
                    self.assertIn(expected_arg, command)

                files_list_index = command.index("-FilesListPath")
                self.assertEqual(AUTO_SELECTED_FILES_LIST, command[files_list_index + 1])
                notify_mock.assert_not_called()

    def test_workroom_cad_tools_preserve_manual_fallback_when_auto_selection_is_empty(self):
        for case in TOOL_CASES:
            with self.subTest(method=case["method_name"]):
                result, captured, notify_mock = self._run_tool(case, auto_selection=None)

                self.assertEqual({"status": "success"}, result)
                self.assertEqual(1, len(captured))

                command, tool_id = captured[0]
                self.assertEqual(case["tool_id"], tool_id)
                self.assertIsInstance(command, list)
                self.assertNotIn("-FilesListPath", command)

                for expected_arg in case["expected_args"]:
                    self.assertIn(expected_arg, command)

                notify_mock.assert_called_once_with(case["tool_id"], FALLBACK_STATUS_MESSAGE)

    def test_resolve_workroom_auto_file_selection_prefers_explicit_launch_context_paths(self):
        with tempfile.TemporaryDirectory(prefix="acies-workroom-cad-") as temp_dir:
            temp_path = Path(temp_dir)
            first_dwg = temp_path / "first.dwg"
            second_dwg = temp_path / "second.dwg"
            invalid_txt = temp_path / "ignore.txt"
            missing_dwg = temp_path / "missing.dwg"
            first_dwg.write_text("", encoding="utf-8")
            second_dwg.write_text("", encoding="utf-8")
            invalid_txt.write_text("", encoding="utf-8")

            launch_context = {
                "source": "workroom",
                "cadFilePaths": [
                    str(first_dwg),
                    "",
                    str(invalid_txt),
                    str(second_dwg),
                    str(first_dwg),
                    str(missing_dwg),
                ],
            }

            with patch.object(
                self.api,
                "_is_workroom_auto_select_enabled",
                return_value=True,
            ), patch.object(
                self.api,
                "_resolve_workroom_context",
                return_value={
                    "source": "workroom",
                    "project_path": r"C:\Projects\123456",
                    "discipline": "Electrical",
                    "discipline_source": "launch_context",
                },
            ), patch.object(
                self.api,
                "_resolve_workroom_discipline_folder",
            ) as resolve_folder_mock, patch.object(
                self.api,
                "_list_base_level_dwgs",
            ) as list_dwgs_mock:
                result = self.api._resolve_workroom_auto_file_selection(
                    {},
                    launch_context,
                    "run_publish_script",
                )

            self.assertIsNotNone(result)
            self.assertEqual("launch_context_explicit_files", result["resolution_mode"])
            self.assertEqual(2, result["count"])
            self.assertEqual(str(temp_path), result["folder_path"])
            self.assertEqual(r"C:\Projects\123456", result["project_path"])
            self.assertEqual("Electrical", result["discipline"])

            files_list_path = Path(result["files_list_path"])
            try:
                self.assertTrue(files_list_path.exists())
                self.assertEqual(
                    [str(first_dwg), str(second_dwg)],
                    files_list_path.read_text(encoding="utf-8").splitlines(),
                )
            finally:
                files_list_path.unlink(missing_ok=True)

            resolve_folder_mock.assert_not_called()
            list_dwgs_mock.assert_not_called()

    def test_get_workroom_cad_files_caches_detected_dwg_paths(self):
        with tempfile.TemporaryDirectory(prefix="acies-workroom-cad-") as temp_dir:
            temp_path = Path(temp_dir)
            first_dwg = temp_path / "first.dwg"
            second_dwg = temp_path / "second.dwg"
            first_dwg.write_text("", encoding="utf-8")
            second_dwg.write_text("", encoding="utf-8")

            with patch.object(
                self.api,
                "get_user_settings",
                return_value={},
            ), patch.object(
                self.api,
                "_resolve_workroom_context",
                return_value={
                    "source": "workroom",
                    "project_path": r"C:\Projects\777777",
                    "discipline": "Electrical",
                    "discipline_source": "launch_context",
                },
            ), patch.object(
                self.api,
                "_resolve_workroom_discipline_folder",
                return_value={
                    "resolved_folder": str(temp_path),
                    "mode": "project_path_child_folder",
                    "discipline": "Electrical",
                    "candidates": [str(temp_path)],
                },
            ), patch.object(
                self.api,
                "_list_base_level_dwgs",
                return_value=[str(first_dwg), str(second_dwg)],
            ):
                result = self.api.get_workroom_cad_files({"source": "workroom"})

            self.assertEqual("success", result["status"])
            self.assertEqual(2, len(result["files"]))

            cache_entry = self.api._get_workroom_cad_file_cache_entry(
                r"C:\Projects\777777",
                "Electrical",
            )
            self.assertIsNotNone(cache_entry)
            self.assertEqual(r"C:\Projects\777777", cache_entry["project_path"])
            self.assertEqual("Electrical", cache_entry["discipline"])
            self.assertEqual(str(temp_path), cache_entry["folder_path"])
            self.assertEqual("project_path_child_folder", cache_entry["resolution_mode"])
            self.assertEqual(
                [str(first_dwg), str(second_dwg)],
                cache_entry["files"],
            )

    def test_resolve_workroom_auto_file_selection_prefers_cached_detection_over_explicit_paths(self):
        with tempfile.TemporaryDirectory(prefix="acies-workroom-cad-") as temp_dir:
            temp_path = Path(temp_dir)
            cached_dwg = temp_path / "cached.dwg"
            explicit_dwg = temp_path / "explicit.dwg"
            cached_dwg.write_text("", encoding="utf-8")
            explicit_dwg.write_text("", encoding="utf-8")
            self.api._set_workroom_cad_file_cache_entry(
                r"C:\Projects\123456",
                "Electrical",
                str(temp_path),
                [str(cached_dwg)],
                "project_path_child_folder",
            )

            launch_context = {
                "source": "workroom",
                "cadFilePaths": [str(explicit_dwg)],
            }

            with patch.object(
                self.api,
                "_is_workroom_auto_select_enabled",
                return_value=True,
            ), patch.object(
                self.api,
                "_resolve_workroom_context",
                return_value={
                    "source": "workroom",
                    "project_path": r"C:\Projects\123456",
                    "discipline": "Electrical",
                    "discipline_source": "launch_context",
                },
            ), patch.object(
                self.api,
                "_resolve_workroom_discipline_folder",
            ) as resolve_folder_mock, patch.object(
                self.api,
                "_list_base_level_dwgs",
            ) as list_dwgs_mock:
                result = self.api._resolve_workroom_auto_file_selection(
                    {},
                    launch_context,
                    "run_publish_script",
                )

            self.assertIsNotNone(result)
            self.assertEqual("workroom_cached_detection", result["resolution_mode"])
            self.assertEqual(1, result["count"])
            self.assertEqual(str(temp_path), result["folder_path"])
            self.assertEqual(r"C:\Projects\123456", result["project_path"])
            self.assertEqual("Electrical", result["discipline"])

            files_list_path = Path(result["files_list_path"])
            try:
                self.assertTrue(files_list_path.exists())
                self.assertEqual(
                    [str(cached_dwg)],
                    files_list_path.read_text(encoding="utf-8").splitlines(),
                )
            finally:
                files_list_path.unlink(missing_ok=True)

            resolve_folder_mock.assert_not_called()
            list_dwgs_mock.assert_not_called()

    def test_resolve_workroom_auto_file_selection_falls_back_to_explicit_paths_when_cached_detection_is_stale(self):
        with tempfile.TemporaryDirectory(prefix="acies-workroom-cad-") as temp_dir:
            temp_path = Path(temp_dir)
            stale_dwg = temp_path / "stale.dwg"
            explicit_dwg = temp_path / "explicit.dwg"
            explicit_dwg.write_text("", encoding="utf-8")
            self.api._set_workroom_cad_file_cache_entry(
                r"C:\Projects\222222",
                "Mechanical",
                str(temp_path),
                [str(stale_dwg)],
                "project_path_child_folder",
            )

            launch_context = {
                "source": "workroom",
                "cadFilePaths": [str(explicit_dwg)],
            }

            with patch.object(
                self.api,
                "_is_workroom_auto_select_enabled",
                return_value=True,
            ), patch.object(
                self.api,
                "_resolve_workroom_context",
                return_value={
                    "source": "workroom",
                    "project_path": r"C:\Projects\222222",
                    "discipline": "Mechanical",
                    "discipline_source": "launch_context",
                },
            ), patch.object(
                self.api,
                "_resolve_workroom_discipline_folder",
            ) as resolve_folder_mock, patch.object(
                self.api,
                "_list_base_level_dwgs",
            ) as list_dwgs_mock:
                result = self.api._resolve_workroom_auto_file_selection(
                    {},
                    launch_context,
                    "run_publish_script",
                )

            self.assertIsNotNone(result)
            self.assertEqual("launch_context_explicit_files", result["resolution_mode"])
            self.assertEqual(1, result["count"])
            self.assertEqual(str(temp_path), result["folder_path"])
            self.assertEqual(r"C:\Projects\222222", result["project_path"])
            self.assertEqual("Mechanical", result["discipline"])

            files_list_path = Path(result["files_list_path"])
            try:
                self.assertTrue(files_list_path.exists())
                self.assertEqual(
                    [str(explicit_dwg)],
                    files_list_path.read_text(encoding="utf-8").splitlines(),
                )
            finally:
                files_list_path.unlink(missing_ok=True)

            resolve_folder_mock.assert_not_called()
            list_dwgs_mock.assert_not_called()

    def test_resolve_workroom_auto_file_selection_falls_back_to_folder_scan_when_explicit_paths_invalid(self):
        with tempfile.TemporaryDirectory(prefix="acies-workroom-cad-") as temp_dir:
            temp_path = Path(temp_dir)
            invalid_txt = temp_path / "ignore.txt"
            fallback_dwg = temp_path / "fallback.dwg"
            invalid_txt.write_text("", encoding="utf-8")
            fallback_dwg.write_text("", encoding="utf-8")

            launch_context = {
                "source": "workroom",
                "cadFilePaths": [str(invalid_txt), str(temp_path / "missing.dwg"), ""],
            }

            with patch.object(
                self.api,
                "_is_workroom_auto_select_enabled",
                return_value=True,
            ), patch.object(
                self.api,
                "_resolve_workroom_context",
                return_value={
                    "source": "workroom",
                    "project_path": r"C:\Projects\654321",
                    "discipline": "Mechanical",
                    "discipline_source": "launch_context",
                },
            ), patch.object(
                self.api,
                "_resolve_workroom_discipline_folder",
                return_value={
                    "resolved_folder": str(temp_path),
                    "mode": "project_path_child_folder",
                    "discipline": "Mechanical",
                    "candidates": [str(temp_path)],
                },
            ) as resolve_folder_mock, patch.object(
                self.api,
                "_list_base_level_dwgs",
                return_value=[str(fallback_dwg)],
            ) as list_dwgs_mock:
                result = self.api._resolve_workroom_auto_file_selection(
                    {},
                    launch_context,
                    "run_publish_script",
                )

            self.assertIsNotNone(result)
            self.assertEqual("project_path_child_folder", result["resolution_mode"])
            self.assertEqual(1, result["count"])
            self.assertEqual(str(temp_path), result["folder_path"])
            self.assertEqual(r"C:\Projects\654321", result["project_path"])
            self.assertEqual("Mechanical", result["discipline"])

            files_list_path = Path(result["files_list_path"])
            try:
                self.assertTrue(files_list_path.exists())
                self.assertEqual(
                    [str(fallback_dwg)],
                    files_list_path.read_text(encoding="utf-8").splitlines(),
                )
            finally:
                files_list_path.unlink(missing_ok=True)

            resolve_folder_mock.assert_called_once_with(r"C:\Projects\654321", "Mechanical")
            list_dwgs_mock.assert_called_once_with(str(temp_path))


if __name__ == "__main__":
    unittest.main()
