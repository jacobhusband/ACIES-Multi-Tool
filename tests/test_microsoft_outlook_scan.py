import datetime
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
        requests_module.post = lambda *args, **kwargs: None
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


class _FixedDateTime(datetime.datetime):
    @classmethod
    def now(cls, tz=None):
        base = cls(2026, 3, 23, 12, 0, 0, tzinfo=datetime.timezone.utc)
        return base.astimezone(tz) if tz else base


class DesktopOutlookScanTests(unittest.TestCase):
    def setUp(self):
        self.api = Api.__new__(Api)

    def test_get_outlook_scan_capability_reports_desktop_availability(self):
        with patch.object(
            self.api,
            "_get_desktop_outlook_availability",
            return_value=(True, ""),
        ):
            result = self.api.get_outlook_scan_capability()

        self.assertEqual(
            {
                "status": "success",
                "desktopAvailable": True,
                "desktopReason": "",
            },
            result,
        )

    def test_scan_outlook_inbox_returns_desktop_unavailable_error(self):
        with patch.object(
            self.api,
            "_get_desktop_outlook_availability",
            return_value=(False, "Classic Outlook is not installed on this machine."),
        ), patch.object(self.api, "_send_outlook_scan_progress") as progress_mock:
            result = self.api.scan_outlook_inbox(
                {"scanDate": "2026-03-20", "projectContext": []},
                "api-key",
                "Jacob Husband",
                ["Electrical"],
            )

        self.assertEqual("error", result["status"])
        self.assertEqual(main_module.OUTLOOK_SCAN_SOURCE_DESKTOP, result["source"])
        self.assertIn("Classic Outlook is not installed", result["message"])
        self.assertEqual("2026-03-20", result["scanDate"])
        self.assertEqual("week", result["timeframe"])
        self.assertEqual([], result["suggestions"])
        self.assertEqual([], result["skippedMessages"])
        self.assertIn("candidateCount", result)
        self.assertIn("scannedCount", result)
        self.assertTrue(
            any(call.args and call.args[0] == "error" for call in progress_mock.call_args_list)
        )

    def test_scan_outlook_inbox_returns_desktop_read_error_without_fallback(self):
        with patch.object(
            self.api,
            "_get_desktop_outlook_availability",
            return_value=(True, ""),
        ), patch.object(
            self.api,
            "_list_desktop_outlook_inbox_messages",
            side_effect=RuntimeError("MAPI unavailable"),
        ), patch.object(self.api, "_send_outlook_scan_progress"):
            result = self.api.scan_outlook_inbox(
                {"scanDate": "2026-03-20", "projectContext": []},
                "api-key",
                "Jacob Husband",
                ["Electrical"],
            )

        self.assertEqual("error", result["status"])
        self.assertEqual(main_module.OUTLOOK_SCAN_SOURCE_DESKTOP, result["source"])
        self.assertEqual(
            "Desktop Outlook could not be read: MAPI unavailable",
            result["message"],
        )
        self.assertNotIn("Microsoft", result["message"])

    def test_scan_outlook_inbox_uses_desktop_messages_when_available(self):
        payload = {
            "scanDate": "2026-03-20",
            "projectContext": [{"id": "proj-1", "deliverables": [{"name": "Submittal"}]}],
        }
        message_summaries = [{"id": "msg-1", "subject": "Need updated submittal"}]
        expected_result = {
            "status": "success",
            "source": main_module.OUTLOOK_SCAN_SOURCE_DESKTOP,
            "scanDate": "2026-03-20",
            "timeframe": "week",
            "suggestions": [],
            "skippedMessages": [],
        }

        with patch.object(
            self.api,
            "_get_desktop_outlook_availability",
            return_value=(True, ""),
        ), patch.object(
            self.api,
            "_list_desktop_outlook_inbox_messages",
            return_value=(message_summaries, False),
        ) as list_mock, patch.object(
            self.api,
            "_analyze_outlook_scan_batch",
            return_value=expected_result,
        ) as analyze_mock, patch.object(self.api, "_send_outlook_scan_progress"):
            result = self.api.scan_outlook_inbox(
                payload,
                "api-key",
                "Jacob Husband",
                ["Electrical"],
            )

        self.assertEqual(expected_result, result)
        list_mock.assert_called_once_with(
            timeframe="week",
            limit=main_module.OUTLOOK_SCAN_FETCH_LIMIT,
            scan_date="2026-03-20",
        )
        args, kwargs = analyze_mock.call_args
        self.assertEqual(message_summaries, args[0])
        self.assertEqual("proj-1", args[2][0]["id"])
        self.assertEqual("", args[2][0]["name"])
        self.assertEqual("Submittal", args[2][0]["deliverables"][0]["name"])
        self.assertEqual("api-key", args[3])
        self.assertEqual("Jacob Husband", args[4])
        self.assertEqual(["Electrical"], args[5])
        self.assertEqual("week", args[6])
        self.assertEqual(main_module.OUTLOOK_SCAN_SOURCE_DESKTOP, args[7])
        self.assertFalse(kwargs["has_more"])
        self.assertEqual("2026-03-20", kwargs["scan_date"])

    def test_scan_outlook_inbox_invalid_scan_date_defaults_to_today(self):
        with patch.object(
            main_module.datetime,
            "datetime",
            _FixedDateTime,
        ), patch.object(
            self.api,
            "_get_desktop_outlook_availability",
            return_value=(True, ""),
        ), patch.object(
            self.api,
            "_list_desktop_outlook_inbox_messages",
            return_value=([], False),
        ) as list_mock, patch.object(
            self.api,
            "_analyze_outlook_scan_batch",
            return_value={"status": "success", "suggestions": [], "skippedMessages": []},
        ), patch.object(self.api, "_send_outlook_scan_progress"):
            self.api.scan_outlook_inbox(
                {"scanDate": "not-a-date", "projectContext": []},
                "api-key",
                "Jacob Husband",
                ["Electrical"],
            )

        list_mock.assert_called_once_with(
            timeframe="week",
            limit=main_module.OUTLOOK_SCAN_FETCH_LIMIT,
            scan_date="2026-03-23",
        )

    def test_open_outlook_desktop_message_displays_mail_item(self):
        class _FakeMailItem:
            def __init__(self):
                self.displayed = False

            def Display(self):
                self.displayed = True

        class _FakeNamespace:
            def __init__(self, item):
                self.item = item
                self.last_id = None

            def GetItemFromID(self, message_id):
                self.last_id = message_id
                return self.item

        mail_item = _FakeMailItem()
        namespace = _FakeNamespace(mail_item)

        with patch.object(
            self.api,
            "_get_desktop_outlook_namespace",
            return_value=(None, namespace),
        ), patch.object(
            self.api,
            "_run_with_outlook_com",
            side_effect=lambda callback: callback(),
        ):
            result = self.api.open_outlook_desktop_message({"messageId": "abc123"})

        self.assertEqual({"status": "success"}, result)
        self.assertEqual("abc123", namespace.last_id)
        self.assertTrue(mail_item.displayed)


if __name__ == "__main__":
    unittest.main()
