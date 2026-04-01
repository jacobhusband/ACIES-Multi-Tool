import os
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest.mock import patch

import fitz
from PIL import Image


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


_ensure_google_genai_stub()
_ensure_webview_stub()
_ensure_dotenv_stub()
_ensure_requests_stub()
_ensure_pydantic_stub()

import main as main_module
from main import Api


class UtilityPlanEditorBackendTests(unittest.TestCase):
    def setUp(self):
        self.api = Api.__new__(Api)

    def _make_pdf(self, output_path):
        doc = fitz.open()
        page = doc.new_page(width=400, height=300)
        page.draw_rect(fitz.Rect(40, 40, 360, 260), color=(0, 0, 0), width=1.5)
        page.insert_text((60, 90), "Main Switchboard", fontsize=18)
        page.insert_text((220, 180), "Panel LA", fontsize=16)
        doc.save(output_path)
        doc.close()

    def test_save_and_load_survey_report_draft_roundtrip(self):
        with tempfile.TemporaryDirectory(prefix="utility-plan-db-") as temp_dir:
            db_path = str(Path(temp_dir) / "survey_reports.db")
            payload = {
                "projectId": "260243",
                "utilityPlan": {
                    "sourcePdfPath": r"C:\Temp\plan.pdf",
                    "exportFolderPath": r"C:\Temp\exports",
                    "floors": [
                        {
                            "id": "floor_1",
                            "label": "Floor 1",
                            "order": 0,
                            "pageNumber": 0,
                            "cropRect": {"x": 10, "y": 20, "width": 200, "height": 150},
                            "callouts": [
                                {
                                    "id": "callout_1",
                                    "text": "MSB",
                                    "labelRect": {"x": 30, "y": 40, "width": 80, "height": 30},
                                    "targetPoint": {"x": 90, "y": 110},
                                }
                            ],
                        }
                    ],
                },
            }

            with patch.object(main_module, "SURVEY_REPORT_DB_FILE", db_path):
                saved = main_module._save_survey_report_draft("260243", payload, updated_by="test")
                loaded = main_module._get_survey_report_draft("260243")

            self.assertEqual("260243", saved["projectId"])
            self.assertEqual("260243", loaded["projectId"])
            self.assertEqual("Floor 1", loaded["utilityPlan"]["floors"][0]["label"])
            self.assertEqual("MSB", loaded["utilityPlan"]["floors"][0]["callouts"][0]["text"])
            self.assertEqual("test", loaded["updatedBy"])
            self.assertGreaterEqual(loaded["version"], 1)

    def test_get_utility_plan_page_preview_returns_png_data_url(self):
        with tempfile.TemporaryDirectory(prefix="utility-plan-preview-") as temp_dir:
            pdf_path = Path(temp_dir) / "floor-plan.pdf"
            self._make_pdf(pdf_path)

            result = self.api.get_utility_plan_page_preview(str(pdf_path), 0, 900)

            self.assertEqual("success", result["status"])
            data = result["data"]
            self.assertEqual(0, data["pageNumber"])
            self.assertEqual(1, data["pageCount"])
            self.assertTrue(str(data["dataUrl"]).startswith("data:image/png;base64,"))
            self.assertGreater(data["previewWidth"], 0)
            self.assertGreater(data["previewHeight"], 0)

    def test_export_utility_plan_pngs_persists_export_paths_and_sanitizes_labels(self):
        with tempfile.TemporaryDirectory(prefix="utility-plan-export-") as temp_dir:
            temp_path = Path(temp_dir)
            pdf_path = temp_path / "source-plan.pdf"
            export_dir = temp_path / "exports"
            export_dir.mkdir()
            db_path = str(temp_path / "survey_reports.db")
            self._make_pdf(pdf_path)

            payload = {
                "projectId": "260243",
                "utilityPlan": {
                    "sourcePdfPath": str(pdf_path),
                    "exportFolderPath": str(export_dir),
                    "floors": [
                        {
                            "id": "floor_1",
                            "label": "Floor/1",
                            "order": 0,
                            "pageNumber": 0,
                            "cropRect": {"x": 0, "y": 0, "width": 400, "height": 300},
                            "callouts": [
                                {
                                    "id": "callout_1",
                                    "text": "MSB",
                                    "labelRect": {"x": 20, "y": 20, "width": 100, "height": 40},
                                    "targetPoint": {"x": 120, "y": 120},
                                }
                            ],
                        },
                        {
                            "id": "floor_2",
                            "label": "Basement:East",
                            "order": 1,
                            "pageNumber": 0,
                            "cropRect": {"x": 20, "y": 20, "width": 200, "height": 160},
                            "callouts": [],
                        },
                    ],
                },
            }

            with patch.object(main_module, "SURVEY_REPORT_DB_FILE", db_path):
                result = self.api.export_utility_plan_pngs("260243", {"draft": payload})
                reloaded = main_module._get_survey_report_draft("260243")

            self.assertEqual("success", result["status"])
            self.assertEqual(2, len(result["exports"]))
            export_names = [Path(item["path"]).name for item in result["exports"]]
            self.assertEqual(
                ["Utility Plan - Floor_1.png", "Utility Plan - Basement_East.png"],
                export_names,
            )
            self.assertTrue(all(Path(item["path"]).exists() for item in result["exports"]))
            self.assertTrue(reloaded["utilityPlan"]["floors"][0]["exportPath"])
            self.assertTrue(reloaded["utilityPlan"]["floors"][1]["exportPath"])

            with Image.open(result["exports"][0]["path"]) as image:
                pixel = image.getpixel((110, 110))
                self.assertGreater(pixel[0], 200)
                self.assertGreater(pixel[1], 180)


if __name__ == "__main__":
    unittest.main()
