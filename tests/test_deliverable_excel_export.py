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


_ensure_google_genai_stub()
_ensure_webview_stub()
_ensure_dotenv_stub()
_ensure_requests_stub()
_ensure_pydantic_stub()

import main as main_module
from main import Api
from openpyxl import load_workbook


class DeliverableExcelExportTests(unittest.TestCase):
    def setUp(self):
        self.api = Api.__new__(Api)

    def test_export_deliverables_excel_orders_rows_by_due_date_descending(self):
        with tempfile.TemporaryDirectory(prefix="deliverable-export-") as temp_dir:
            output_path = Path(temp_dir) / "deliverables.xlsx"
            payload = {
                "filePath": str(output_path),
                "entries": [
                    {
                        "projectId": "1003",
                        "projectName": "No Date",
                        "deliverableName": "Undated",
                        "due": "",
                        "statusText": "Waiting",
                    },
                    {
                        "projectId": "1002",
                        "projectName": "Past Project",
                        "deliverableName": "Past Deliverable",
                        "due": "01/15/2026",
                        "statusText": "Complete",
                    },
                    {
                        "projectId": "1001",
                        "projectName": "Future Project",
                        "deliverableName": "Future Deliverable",
                        "due": "12/31/2026",
                        "statusText": "Working",
                    },
                    {
                        "projectId": "1004",
                        "projectName": "Middle Project",
                        "deliverableName": "Middle Deliverable",
                        "due": "2026-06-01",
                        "statusText": "Pending Review",
                    },
                ],
            }

            with patch.object(main_module.os, "startfile", create=True):
                result = self.api.export_deliverables_excel(payload)

            self.assertEqual("success", result["status"])
            self.assertEqual(str(output_path), result["path"])
            self.assertTrue(output_path.exists())

            workbook = load_workbook(output_path)
            self.addCleanup(workbook.close)
            worksheet = workbook["Deliverables"]

        self.assertEqual(
            ("Project ID", "Project Name", "Deliverable", "Due Date", "Status"),
            tuple(cell.value for cell in worksheet[1]),
        )
        self.assertEqual("1001", worksheet["A2"].value)
        self.assertEqual("Future Project", worksheet["B2"].value)
        self.assertEqual("Future Deliverable", worksheet["C2"].value)
        self.assertEqual("Working", worksheet["E2"].value)
        self.assertEqual("1004", worksheet["A3"].value)
        self.assertEqual("1002", worksheet["A4"].value)
        self.assertEqual("1003", worksheet["A5"].value)
        self.assertEqual("mm/dd/yyyy", worksheet["D2"].number_format)
        self.assertEqual("01/15/2026", worksheet["D4"].value.strftime("%m/%d/%Y"))
        self.assertIsNone(worksheet["D5"].value)


if __name__ == "__main__":
    unittest.main()
