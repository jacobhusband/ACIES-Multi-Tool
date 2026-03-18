import base64
import datetime
import json
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


class _FakeResponse:
    def __init__(self, status_code=200, payload=None, text=""):
        self.status_code = status_code
        self._payload = payload or {}
        self.text = text

    def json(self):
        return self._payload


def _build_fake_jwt(claims):
    def _encode(segment):
        raw = json.dumps(segment, separators=(",", ":")).encode("utf-8")
        return base64.urlsafe_b64encode(raw).rstrip(b"=").decode("ascii")

    return ".".join(
        [
            _encode({"alg": "none", "typ": "JWT"}),
            _encode(claims),
            "signature",
        ]
    )


class MicrosoftOutlookScanTests(unittest.TestCase):
    def setUp(self):
        self.api = Api.__new__(Api)

    def test_sign_in_with_microsoft_requires_client_id(self):
        with patch.object(self.api, "_get_microsoft_oauth_client_id", return_value=""):
            result = self.api.sign_in_with_microsoft()

        self.assertEqual("error", result["status"])
        self.assertIn(main_module.MICROSOFT_OAUTH_CLIENT_ID_ENV, result["message"])

    def test_refresh_microsoft_auth_record_persists_auth(self):
        captured = {}
        stale_expiry = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(minutes=10)
        existing_auth = {
            "provider": "microsoft",
            "subject": "subject-1",
            "email": "user@example.com",
            "displayName": "User Example",
            "signedInAt": "2026-01-01T00:00:00Z",
            "refreshToken": "refresh-token",
            "expiresAt": "2026-01-01T00:00:00Z",
            "tenantId": "tenant-old",
        }
        id_token = _build_fake_jwt(
            {
                "sub": "subject-2",
                "email": "refreshed@example.com",
                "name": "Refreshed User",
                "tid": "tenant-123",
            }
        )

        def fake_post(url, data=None, timeout=None):
            captured["url"] = url
            captured["data"] = dict(data or {})
            return _FakeResponse(
                payload={
                    "access_token": "refreshed-token",
                    "refresh_token": "refreshed-refresh-token",
                    "expires_in": 3600,
                    "token_type": "Bearer",
                    "scope": "openid profile email offline_access Mail.Read",
                    "id_token": id_token,
                }
            )

        with patch.object(self.api, "_get_microsoft_oauth_client_id", return_value="client-id"), patch.object(
            main_module, "parse_utc_iso", return_value=stale_expiry
        ), patch.object(main_module.requests, "post", side_effect=fake_post), patch.object(
            self.api, "_persist_microsoft_auth_record"
        ) as persist_auth:
            refreshed = self.api._refresh_microsoft_auth_record_if_needed(existing_auth)

        self.assertEqual("client-id", captured["data"]["client_id"])
        self.assertEqual("refresh_token", captured["data"]["grant_type"])
        self.assertEqual("refresh-token", captured["data"]["refresh_token"])
        self.assertEqual("microsoft", refreshed["provider"])
        self.assertEqual("tenant-123", refreshed["tenantId"])
        self.assertEqual("subject-2", refreshed["subject"])
        self.assertEqual("2026-01-01T00:00:00Z", refreshed["signedInAt"])
        persist_auth.assert_called_once_with(refreshed)

    def test_scan_outlook_inbox_builds_graph_requests(self):
        list_call = {}
        detail_call = {}

        def fake_get(url, headers=None, params=None, timeout=None):
            if url.endswith("/me/mailFolders/inbox/messages"):
                list_call["url"] = url
                list_call["headers"] = dict(headers or {})
                list_call["params"] = dict(params or {})
                return _FakeResponse(
                    payload={
                        "value": [
                            {
                                "id": "msg-1",
                                "subject": "250597 DD60 submittal",
                                "bodyPreview": "Please issue DD60 this week.",
                                "receivedDateTime": "2026-03-17T18:00:00Z",
                                "webLink": "https://outlook.office.com/mail/deeplink/read/msg-1",
                                "internetMessageId": "<msg-1@example.com>",
                                "hasAttachments": False,
                                "from": {
                                    "emailAddress": {
                                        "name": "Project Lead",
                                        "address": "lead@example.com",
                                    }
                                },
                            }
                        ]
                    }
                )
            if "/me/messages/msg-1" in url:
                detail_call["url"] = url
                detail_call["headers"] = dict(headers or {})
                detail_call["params"] = dict(params or {})
                return _FakeResponse(
                    payload={
                        "id": "msg-1",
                        "subject": "250597 DD60 submittal",
                        "body": {
                            "contentType": "text",
                            "content": "Please issue DD60 drawings this week.",
                        },
                        "bodyPreview": "Please issue DD60 this week.",
                        "receivedDateTime": "2026-03-17T18:00:00Z",
                        "webLink": "https://outlook.office.com/mail/deeplink/read/msg-1",
                        "internetMessageId": "<msg-1@example.com>",
                        "hasAttachments": False,
                        "from": {
                            "emailAddress": {
                                "name": "Project Lead",
                                "address": "lead@example.com",
                            }
                        },
                    }
                )
            raise AssertionError(f"Unexpected Graph URL: {url}")

        with patch.object(self.api, "_load_microsoft_auth_record", return_value={"accessToken": "access-token"}), patch.object(
            self.api, "_refresh_microsoft_auth_record_if_needed", side_effect=lambda auth: auth
        ), patch.object(
            self.api,
            "_extract_project_data_from_email_text",
            return_value={
                "id": "250597",
                "name": "Client Tower",
                "due": "03/21/2026",
                "path": "",
                "deliverable": "DD60",
                "tasks": [{"text": "Issue DD60 set", "done": False, "links": []}],
                "notes": "Issue DD60 drawings this week.",
            },
        ), patch.object(main_module.requests, "get", side_effect=fake_get):
            result = self.api.scan_outlook_inbox(
                {"timeframe": "week"},
                "api-key",
                "Jacob",
                ["Electrical"],
            )

        self.assertEqual("success", result["status"])
        self.assertEqual("Bearer access-token", list_call["headers"]["Authorization"])
        self.assertEqual("receivedDateTime desc", list_call["params"]["$orderby"])
        self.assertIn("receivedDateTime ge ", list_call["params"]["$filter"])
        self.assertIn("webLink", list_call["params"]["$select"])
        self.assertEqual(
            'outlook.body-content-type="text"',
            detail_call["headers"]["Prefer"],
        )
        self.assertIn("body", detail_call["params"]["$select"])
        self.assertEqual(1, result["candidateCount"])
        self.assertEqual(1, result["scannedCount"])
        self.assertEqual("success", result["items"][0]["analysisStatus"])

    def test_scan_outlook_inbox_returns_partial_results_when_one_message_fails(self):
        def fake_get(url, headers=None, params=None, timeout=None):
            if url.endswith("/me/mailFolders/inbox/messages"):
                return _FakeResponse(
                    payload={
                        "value": [
                            {
                                "id": "msg-1",
                                "subject": "250597 DD60 submittal",
                                "bodyPreview": "Please issue DD60 this week.",
                                "receivedDateTime": "2026-03-17T18:00:00Z",
                                "webLink": "https://outlook.office.com/mail/deeplink/read/msg-1",
                                "internetMessageId": "<msg-1@example.com>",
                                "hasAttachments": False,
                                "from": {
                                    "emailAddress": {
                                        "name": "Project Lead",
                                        "address": "lead@example.com",
                                    }
                                },
                            },
                            {
                                "id": "msg-2",
                                "subject": "250597 Bulletin request",
                                "bodyPreview": "Please issue a bulletin.",
                                "receivedDateTime": "2026-03-16T18:00:00Z",
                                "webLink": "https://outlook.office.com/mail/deeplink/read/msg-2",
                                "internetMessageId": "<msg-2@example.com>",
                                "hasAttachments": False,
                                "from": {
                                    "emailAddress": {
                                        "name": "Project Lead",
                                        "address": "lead@example.com",
                                    }
                                },
                            },
                        ]
                    }
                )
            if "/me/messages/msg-1" in url:
                return _FakeResponse(
                    payload={
                        "id": "msg-1",
                        "subject": "250597 DD60 submittal",
                        "body": {
                            "contentType": "text",
                            "content": "Please issue DD60 drawings this week.",
                        },
                        "bodyPreview": "Please issue DD60 this week.",
                        "receivedDateTime": "2026-03-17T18:00:00Z",
                        "webLink": "https://outlook.office.com/mail/deeplink/read/msg-1",
                        "internetMessageId": "<msg-1@example.com>",
                        "hasAttachments": False,
                        "from": {
                            "emailAddress": {
                                "name": "Project Lead",
                                "address": "lead@example.com",
                            }
                        },
                    }
                )
            if "/me/messages/msg-2" in url:
                raise RuntimeError("detail fetch failed")
            raise AssertionError(f"Unexpected Graph URL: {url}")

        with patch.object(self.api, "_load_microsoft_auth_record", return_value={"accessToken": "access-token"}), patch.object(
            self.api, "_refresh_microsoft_auth_record_if_needed", side_effect=lambda auth: auth
        ), patch.object(
            self.api,
            "_extract_project_data_from_email_text",
            return_value={
                "id": "250597",
                "name": "Client Tower",
                "due": "03/21/2026",
                "path": "",
                "deliverable": "DD60",
                "tasks": [{"text": "Issue DD60 set", "done": False, "links": []}],
                "notes": "Issue DD60 drawings this week.",
            },
        ), patch.object(main_module.requests, "get", side_effect=fake_get):
            result = self.api.scan_outlook_inbox(
                {"timeframe": "week"},
                "api-key",
                "Jacob",
                ["Electrical"],
            )

        self.assertEqual("success", result["status"])
        self.assertEqual(2, result["candidateCount"])
        self.assertEqual(2, result["scannedCount"])
        statuses = [item["analysisStatus"] for item in result["items"]]
        self.assertIn("success", statuses)
        self.assertIn("error", statuses)


if __name__ == "__main__":
    unittest.main()
