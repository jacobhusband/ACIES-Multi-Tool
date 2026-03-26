import os
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

from main import Api


class PanelScheduleAiBackendTests(unittest.TestCase):
    def setUp(self):
        self.api = Api.__new__(Api)

    def test_process_panel_schedule_payload_uses_plural_path_lists(self):
        with tempfile.TemporaryDirectory(prefix="acies-panel-schedule-") as temp_dir:
            output_path = Path(temp_dir) / "panel_schedule.xlsx"
            output_path.write_text("placeholder", encoding="utf-8")
            captured = {}
            fake_panel_data = types.SimpleNamespace(panel_name="MDP")

            def fake_analyze(panel_name, breaker_paths, directory_paths):
                captured["panel_name"] = panel_name
                captured["breaker_paths"] = list(breaker_paths)
                captured["directory_paths"] = list(directory_paths)
                return fake_panel_data

            with patch.object(
                self.api,
                "_analyze_panel_schedule_images",
                side_effect=fake_analyze,
            ), patch.object(
                self.api,
                "_update_panel_schedule_workbook",
                return_value="MDP",
            ) as update_mock:
                result = self.api._process_panel_schedule_payload(
                    {
                        "outputMode": "existing",
                        "outputPath": str(output_path),
                        "panels": [
                            {
                                "panelId": "panel_1",
                                "panelName": "MDP",
                                "breakerPaths": [
                                    "breaker-top.jpg",
                                    "",
                                    "breaker-bottom.jpg",
                                ],
                                "directoryPaths": [
                                    "directory-1-42.jpg",
                                    "directory-43-84.jpg",
                                ],
                            }
                        ],
                    }
                )

            self.assertEqual("success", result["status"])
            self.assertEqual("MDP", captured["panel_name"])
            self.assertEqual(
                ["breaker-top.jpg", "breaker-bottom.jpg"],
                captured["breaker_paths"],
            )
            self.assertEqual(
                ["directory-1-42.jpg", "directory-43-84.jpg"],
                captured["directory_paths"],
            )
            update_mock.assert_called_once_with(
                fake_panel_data,
                os.path.normpath(str(output_path)),
            )

    def test_process_panel_schedule_payload_requires_at_least_one_breaker_and_directory_photo(self):
        with tempfile.TemporaryDirectory(prefix="acies-panel-schedule-") as temp_dir:
            output_path = Path(temp_dir) / "panel_schedule.xlsx"
            output_path.write_text("placeholder", encoding="utf-8")

            result = self.api._process_panel_schedule_payload(
                {
                    "outputMode": "existing",
                    "outputPath": str(output_path),
                    "panelName": "MDP",
                    "breakerPaths": [],
                    "directoryPaths": [],
                }
            )

        self.assertEqual("error", result["status"])
        self.assertEqual(0, result["successCount"])
        self.assertEqual(1, result["failureCount"])
        self.assertIn(
            "At least one breaker photo and at least one directory photo are required.",
            result["message"],
        )

    def test_build_panel_schedule_prompt_mentions_split_breakers_and_directories(self):
        prompt = self.api._build_panel_schedule_prompt("MDP", 3, 2)

        self.assertIn("upper half, middle, and bottom half images", prompt)
        self.assertIn("circuits 1-42 on one image and 43-84 on another", prompt)
        self.assertIn("Do not treat split photos as separate panels.", prompt)
        self.assertIn(
            "Use visible breaker positions and circuit numbering to merge split sections",
            prompt,
        )


if __name__ == "__main__":
    unittest.main()
