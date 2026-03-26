import sys
import types
import unittest
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


class RevealPathBackendTests(unittest.TestCase):
    def setUp(self):
        self.api = Api.__new__(Api)

    def test_reveal_path_selects_existing_file_in_windows_explorer(self):
        target_path = r"C:\Projects\250940 Loro Piana\Task Lighting\sheet.pdf"
        parent_path = r"C:\Projects\250940 Loro Piana\Task Lighting"

        with patch.object(main_module.sys, "platform", "win32"), \
             patch.object(main_module.os.path, "normpath", side_effect=lambda value: value), \
             patch.object(main_module.os.path, "dirname", return_value=parent_path), \
             patch.object(main_module.os.path, "exists", side_effect=lambda value: value in {target_path, parent_path}), \
             patch.object(main_module.os.path, "isfile", side_effect=lambda value: value == target_path), \
             patch.object(main_module.os.path, "isdir", side_effect=lambda value: value == parent_path), \
             patch.object(main_module.os, "startfile", create=True) as mock_startfile, \
             patch.object(main_module.subprocess, "run") as mock_run:
            response = self.api.reveal_path(target_path)

        self.assertEqual("success", response["status"])
        self.assertEqual("select-file", response["mode"])
        self.assertEqual(target_path, response["path"])
        mock_run.assert_called_once_with(
            ["explorer.exe", "/select,", target_path],
            check=False,
        )
        mock_startfile.assert_not_called()

    def test_reveal_path_opens_existing_directory(self):
        target_path = r"C:\Projects\250940 Loro Piana"

        with patch.object(main_module.sys, "platform", "win32"), \
             patch.object(main_module.os.path, "normpath", side_effect=lambda value: value), \
             patch.object(main_module.os.path, "dirname", return_value=r"C:\Projects"), \
             patch.object(main_module.os.path, "exists", side_effect=lambda value: value == target_path), \
             patch.object(main_module.os.path, "isfile", return_value=False), \
             patch.object(main_module.os.path, "isdir", side_effect=lambda value: value == target_path), \
             patch.object(main_module.os, "startfile", create=True) as mock_startfile, \
             patch.object(main_module.subprocess, "run") as mock_run:
            response = self.api.reveal_path(target_path)

        self.assertEqual("success", response["status"])
        self.assertEqual("open-directory", response["mode"])
        self.assertEqual(target_path, response["path"])
        mock_startfile.assert_called_once_with(target_path)
        mock_run.assert_not_called()

    def test_reveal_path_falls_back_to_parent_when_target_missing(self):
        target_path = r"C:\Projects\250940 Loro Piana\missing\sheet.pdf"
        parent_path = r"C:\Projects\250940 Loro Piana\missing"

        with patch.object(main_module.sys, "platform", "win32"), \
             patch.object(main_module.os.path, "normpath", side_effect=lambda value: value), \
             patch.object(main_module.os.path, "dirname", return_value=parent_path), \
             patch.object(main_module.os.path, "exists", side_effect=lambda value: value == parent_path), \
             patch.object(main_module.os.path, "isfile", return_value=False), \
             patch.object(main_module.os.path, "isdir", side_effect=lambda value: value == parent_path), \
             patch.object(main_module.os, "startfile", create=True) as mock_startfile, \
             patch.object(main_module.subprocess, "run") as mock_run:
            response = self.api.reveal_path(target_path)

        self.assertEqual("success", response["status"])
        self.assertEqual("open-parent", response["mode"])
        self.assertEqual(parent_path, response["path"])
        mock_startfile.assert_called_once_with(parent_path)
        mock_run.assert_not_called()

    def test_reveal_path_errors_when_target_and_parent_are_missing(self):
        target_path = r"C:\Projects\250940 Loro Piana\missing\sheet.pdf"
        parent_path = r"C:\Projects\250940 Loro Piana\missing"

        with patch.object(main_module.sys, "platform", "win32"), \
             patch.object(main_module.os.path, "normpath", side_effect=lambda value: value), \
             patch.object(main_module.os.path, "dirname", return_value=parent_path), \
             patch.object(main_module.os.path, "exists", return_value=False), \
             patch.object(main_module.os.path, "isfile", return_value=False), \
             patch.object(main_module.os.path, "isdir", return_value=False), \
             patch.object(main_module.os, "startfile", create=True) as mock_startfile, \
             patch.object(main_module.subprocess, "run") as mock_run:
            response = self.api.reveal_path(target_path)

        self.assertEqual("error", response["status"])
        self.assertIn("Path and parent do not exist", response["message"])
        mock_startfile.assert_not_called()
        mock_run.assert_not_called()


if __name__ == "__main__":
    unittest.main()
