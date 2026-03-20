import base64
import datetime
import json
import sys
import types
import unittest
from urllib.parse import parse_qs, urlparse
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

    def test_open_url_prefers_edge_protocol_on_windows(self):
        with patch.object(main_module.sys, "platform", "win32"), patch.object(
            main_module.os, "startfile", create=True
        ) as startfile:
            result = self.api.open_url("https://login.microsoftonline.com/test", browser="edge")

        self.assertEqual("success", result["status"])
        startfile.assert_called_once_with(
            "microsoft-edge:https://login.microsoftonline.com/test"
        )

    def test_extract_microsoft_error_message_mentions_consent_policy_steps(self):
        response = _FakeResponse(
            status_code=400,
            payload={
                "error": "invalid_grant",
                "error_description": (
                    "AADSTS65001: The user or administrator has not consented to use "
                    "the application with ID ad3bdf05-66e2-4a1a-9660-4ca1ba237007."
                ),
            },
        )

        with patch.object(self.api, "_get_microsoft_oauth_client_id", return_value="ad3bdf05-66e2-4a1a-9660-4ca1ba237007"), patch.object(
            self.api, "_get_microsoft_oauth_tenant_id", return_value="417e5aea-af22-46eb-9b47-8667938cdcaf"
        ):
            message = self.api._extract_microsoft_error_message(response)

        self.assertIn("tenant consent policy", message)
        self.assertIn("Mail.Read", message)
        self.assertIn("Retry minimal consent", message)

    def test_request_microsoft_admin_consent_uses_adminconsent_endpoint(self):
        opened = {}

        class _FakeEvent:
            def wait(self, timeout):
                opened["wait_timeout"] = timeout
                return True

        class _FakeCallbackServer:
            def __init__(self):
                self.server_address = (main_module.MICROSOFT_OAUTH_CALLBACK_HOST, 43123)
                self.oauth_event = _FakeEvent()
                self.oauth_result = {
                    "admin_consent": "True",
                    "state": "expected-state",
                    "tenant": "417e5aea-af22-46eb-9b47-8667938cdcaf",
                }
                self.shutdown_called = False
                self.server_close_called = False

            def shutdown(self):
                self.shutdown_called = True

            def server_close(self):
                self.server_close_called = True

        callback_server = _FakeCallbackServer()

        def fake_open_url(url, browser=None):
            opened["url"] = url
            opened["browser"] = browser
            return {"status": "success"}

        with patch.object(self.api, "_get_microsoft_oauth_client_id", return_value="ad3bdf05-66e2-4a1a-9660-4ca1ba237007"), patch.object(
            self.api, "_get_microsoft_oauth_tenant_id", return_value="417e5aea-af22-46eb-9b47-8667938cdcaf"
        ), patch.object(
            self.api, "_start_microsoft_oauth_callback_server", return_value=callback_server
        ), patch.object(
            main_module.secrets, "token_urlsafe", return_value="expected-state"
        ), patch.object(
            self.api, "open_url", side_effect=fake_open_url
        ):
            result = self.api.request_microsoft_admin_consent()

        parsed = urlparse(opened["url"])
        params = parse_qs(parsed.query)
        self.assertEqual("success", result["status"])
        self.assertEqual("/417e5aea-af22-46eb-9b47-8667938cdcaf/v2.0/adminconsent", parsed.path)
        self.assertEqual("edge", opened["browser"])
        self.assertEqual(
            f"http://{main_module.MICROSOFT_OAUTH_CALLBACK_HOST}:43123/",
            params["redirect_uri"][0],
        )
        self.assertIn("https://graph.microsoft.com/Mail.Read", params["scope"][0])
        self.assertIn("openid", params["scope"][0])
        self.assertEqual(main_module.MICROSOFT_OAUTH_TIMEOUT_SECONDS, opened["wait_timeout"])
        self.assertTrue(callback_server.shutdown_called)
        self.assertTrue(callback_server.server_close_called)

    def test_diagnose_microsoft_consent_policy_uses_reduced_scopes(self):
        opened = {}

        class _FakeEvent:
            def wait(self, timeout):
                opened["wait_timeout"] = timeout
                return True

        class _FakeCallbackServer:
            def __init__(self):
                self.server_address = (main_module.MICROSOFT_OAUTH_CALLBACK_HOST, 43123)
                self.oauth_event = _FakeEvent()
                self.oauth_result = {"code": "auth-code", "state": "expected-state"}
                self.shutdown_called = False
                self.server_close_called = False

            def shutdown(self):
                self.shutdown_called = True

            def server_close(self):
                self.server_close_called = True

        callback_server = _FakeCallbackServer()
        id_token = _build_fake_jwt(
            {
                "sub": "subject-1",
                "email": "user@example.com",
                "name": "User Example",
                "tid": "tenant-123",
            }
        )

        def fake_open_url(url, browser=None):
            opened["url"] = url
            opened["browser"] = browser
            return {"status": "success"}

        def fake_exchange(client_id, code, code_verifier, redirect_uri):
            return {
                "access_token": "access-token",
                "expires_in": 3600,
                "token_type": "Bearer",
                "scope": "openid profile email User.Read",
                "id_token": id_token,
            }

        with patch.object(self.api, "_get_microsoft_oauth_client_id", return_value="client-id"), patch.object(
            self.api, "_start_microsoft_oauth_callback_server", return_value=callback_server
        ), patch.object(
            self.api, "_generate_google_pkce_pair", return_value=("verifier-123", "challenge-123")
        ), patch.object(
            main_module.secrets, "token_urlsafe", return_value="expected-state"
        ), patch.object(
            self.api, "open_url", side_effect=fake_open_url
        ), patch.object(
            self.api, "_exchange_microsoft_auth_code", side_effect=fake_exchange
        ):
            result = self.api.diagnose_microsoft_consent_policy()

        parsed = urlparse(opened["url"])
        params = parse_qs(parsed.query)
        self.assertEqual("success", result["status"])
        self.assertEqual("edge", opened["browser"])
        self.assertEqual("openid profile email User.Read", params["scope"][0])
        self.assertNotIn("Mail.Read", params["scope"][0])
        self.assertNotIn("offline_access", params["scope"][0])
        self.assertIn("Basic sign-in works", result["message"])
        self.assertEqual("user@example.com", result["diagnostic"]["email"])
        self.assertTrue(callback_server.shutdown_called)
        self.assertTrue(callback_server.server_close_called)

    def test_sign_in_with_microsoft_uses_localhost_redirect_uri(self):
        opened = {}
        exchanged = {}
        persisted = {}
        id_token = _build_fake_jwt(
            {
                "sub": "subject-1",
                "email": "user@example.com",
                "name": "User Example",
                "tid": "tenant-123",
            }
        )

        class _FakeEvent:
            def wait(self, timeout):
                opened["wait_timeout"] = timeout
                return True

        class _FakeCallbackServer:
            def __init__(self):
                self.server_address = (main_module.MICROSOFT_OAUTH_CALLBACK_HOST, 43123)
                self.oauth_event = _FakeEvent()
                self.oauth_result = {"code": "auth-code", "state": "expected-state"}
                self.shutdown_called = False
                self.server_close_called = False

            def shutdown(self):
                self.shutdown_called = True

            def server_close(self):
                self.server_close_called = True

        callback_server = _FakeCallbackServer()

        def fake_open_url(url, browser=None):
            opened["url"] = url
            opened["browser"] = browser
            return {"status": "success"}

        def fake_exchange(client_id, code, code_verifier, redirect_uri):
            exchanged["client_id"] = client_id
            exchanged["code"] = code
            exchanged["code_verifier"] = code_verifier
            exchanged["redirect_uri"] = redirect_uri
            return {
                "access_token": "access-token",
                "refresh_token": "refresh-token",
                "expires_in": 3600,
                "token_type": "Bearer",
                "scope": "openid profile email offline_access Mail.Read",
                "id_token": id_token,
            }

        with patch.object(self.api, "_get_microsoft_oauth_client_id", return_value="client-id"), patch.object(
            self.api, "_load_microsoft_auth_record", return_value=None
        ), patch.object(
            self.api, "_start_microsoft_oauth_callback_server", return_value=callback_server
        ), patch.object(
            self.api, "_generate_google_pkce_pair", return_value=("verifier-123", "challenge-123")
        ), patch.object(
            main_module.secrets, "token_urlsafe", return_value="expected-state"
        ), patch.object(
            self.api, "open_url", side_effect=fake_open_url
        ), patch.object(
            self.api, "_exchange_microsoft_auth_code", side_effect=fake_exchange
        ), patch.object(
            self.api, "_persist_microsoft_auth_record", side_effect=lambda auth: persisted.setdefault("auth", auth)
        ):
            result = self.api.sign_in_with_microsoft()

        parsed = urlparse(opened["url"])
        params = parse_qs(parsed.query)
        self.assertEqual("success", result["status"])
        self.assertEqual(
            f"http://{main_module.MICROSOFT_OAUTH_CALLBACK_HOST}:43123/",
            params["redirect_uri"][0],
        )
        self.assertEqual(
            f"http://{main_module.MICROSOFT_OAUTH_CALLBACK_HOST}:43123/",
            exchanged["redirect_uri"],
        )
        self.assertEqual("client-id", exchanged["client_id"])
        self.assertEqual("auth-code", exchanged["code"])
        self.assertEqual("verifier-123", exchanged["code_verifier"])
        self.assertEqual("edge", opened["browser"])
        self.assertEqual(main_module.MICROSOFT_OAUTH_TIMEOUT_SECONDS, opened["wait_timeout"])
        self.assertTrue(callback_server.shutdown_called)
        self.assertTrue(callback_server.server_close_called)
        self.assertEqual("user@example.com", persisted["auth"]["email"])
        self.assertEqual("tenant-123", persisted["auth"]["tenantId"])

    def test_sign_in_with_microsoft_reports_consent_policy_error(self):
        class _FakeEvent:
            def wait(self, timeout):
                return True

        class _FakeCallbackServer:
            def __init__(self):
                self.server_address = (main_module.MICROSOFT_OAUTH_CALLBACK_HOST, 43123)
                self.oauth_event = _FakeEvent()
                self.oauth_result = {
                    "error": "access_denied",
                    "error_description": (
                        "AADSTS65001: The user or administrator has not consented to use "
                        "the application with ID ad3bdf05-66e2-4a1a-9660-4ca1ba237007."
                    ),
                }

            def shutdown(self):
                return None

            def server_close(self):
                return None

        with patch.object(self.api, "_get_microsoft_oauth_client_id", return_value="ad3bdf05-66e2-4a1a-9660-4ca1ba237007"), patch.object(
            self.api, "_get_microsoft_oauth_tenant_id", return_value="417e5aea-af22-46eb-9b47-8667938cdcaf"
        ), patch.object(
            self.api, "_load_microsoft_auth_record", return_value=None
        ), patch.object(
            self.api, "_start_microsoft_oauth_callback_server", return_value=_FakeCallbackServer()
        ), patch.object(
            self.api, "_generate_google_pkce_pair", return_value=("verifier-123", "challenge-123")
        ), patch.object(
            main_module.secrets, "token_urlsafe", return_value="expected-state"
        ), patch.object(
            self.api, "open_url", return_value={"status": "success"}
        ):
            result = self.api.sign_in_with_microsoft()

        self.assertEqual("error", result["status"])
        self.assertIn("tenant consent policy", result["message"])
        self.assertIn("Mail.Read", result["message"])
        self.assertIn("Retry minimal consent", result["message"])

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
        ai_prompts = []
        progress_events = []

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
                                "conversationId": "conv-1",
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
                        "uniqueBody": {
                            "contentType": "text",
                            "content": "Only the new part of the reply.",
                        },
                        "body": {
                            "contentType": "text",
                            "content": (
                                "Only the new part of the reply.\n\n"
                                "-----Original Message-----\n"
                                "Please issue DD60 drawings this week."
                            ),
                        },
                        "bodyPreview": "Please issue DD60 this week.",
                        "receivedDateTime": "2026-03-17T18:00:00Z",
                        "webLink": "https://outlook.office.com/mail/deeplink/read/msg-1",
                        "internetMessageId": "<msg-1@example.com>",
                        "conversationId": "conv-1",
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

        def fake_batch(prompt, api_key):
            ai_prompts.append({"prompt": prompt, "api_key": api_key})
            return [
                {
                    "project": {
                        "id": "250597",
                        "name": "Client Tower",
                        "nick": "Tower",
                        "path": r"P:\250597 Client Tower",
                    },
                    "deliverable": {
                        "name": "DD90",
                        "due": "03/21/2026",
                        "notes": "Issue DD90 drawings this week.",
                        "tasks": ["Issue DD90 set"],
                    },
                    "supportingMessageRefs": ["E001"],
                }
            ]

        with patch.object(self.api, "_get_desktop_outlook_availability", return_value=(False, "")), patch.object(
            self.api, "_load_microsoft_auth_record", return_value={"accessToken": "access-token"}
        ), patch.object(
            self.api, "_refresh_microsoft_auth_record_if_needed", side_effect=lambda auth: auth
        ), patch.object(
            self.api, "_request_outlook_scan_batch_suggestions", side_effect=fake_batch
        ), patch.object(
            self.api,
            "_notify_outlook_scan_progress",
            side_effect=lambda payload: progress_events.append(dict(payload or {})),
        ), patch.object(main_module.requests, "get", side_effect=fake_get):
            result = self.api.scan_outlook_inbox(
                {
                    "scanDate": "2026-03-17",
                    "projectContext": [
                        {
                            "id": "250597",
                            "name": "Client Tower",
                            "nick": "Tower",
                            "path": r"P:\250597 Client Tower",
                            "deliverables": [
                                {
                                    "name": "DD60",
                                    "due": "03/17/2026",
                                    "status": "Working",
                                }
                            ],
                        }
                    ],
                },
                "api-key",
                "Jacob",
                ["Electrical"],
            )

        expected_range = self.api._build_outlook_scan_range(scan_date="2026-03-17")
        self.assertEqual("success", result["status"])
        self.assertEqual("batch", result["analysisMode"])
        self.assertEqual("2026-03-17", result["scanDate"])
        self.assertEqual("Bearer access-token", list_call["headers"]["Authorization"])
        self.assertEqual("receivedDateTime desc", list_call["params"]["$orderby"])
        self.assertIn("receivedDateTime ge ", list_call["params"]["$filter"])
        self.assertIn("receivedDateTime le ", list_call["params"]["$filter"])
        self.assertIn(
            expected_range["startUtc"].isoformat().replace("+00:00", "Z"),
            list_call["params"]["$filter"],
        )
        self.assertIn(
            expected_range["endUtc"].isoformat().replace("+00:00", "Z"),
            list_call["params"]["$filter"],
        )
        self.assertIn("webLink", list_call["params"]["$select"])
        self.assertIn("conversationId", list_call["params"]["$select"])
        self.assertEqual(
            'outlook.body-content-type="text"',
            detail_call["headers"]["Prefer"],
        )
        self.assertIn("body", detail_call["params"]["$select"])
        self.assertIn("uniqueBody", detail_call["params"]["$select"])
        self.assertIn("conversationId", detail_call["params"]["$select"])
        self.assertEqual(1, result["candidateCount"])
        self.assertEqual(1, result["scannedCount"])
        self.assertEqual(1, result["emailsIncludedCount"])
        self.assertEqual(1, result["deliverablesIncludedCount"])
        self.assertEqual(1, len(ai_prompts))
        self.assertIn("CURRENT_DELIVERABLES_IN_PERIOD", ai_prompts[0]["prompt"])
        self.assertIn('"deliverable":"DD60"', ai_prompts[0]["prompt"])
        self.assertIn('"conversationId":"conv-1"', ai_prompts[0]["prompt"])
        self.assertIn("03/17/2026", ai_prompts[0]["prompt"])
        self.assertIn("Only the new part of the reply.", ai_prompts[0]["prompt"])
        self.assertNotIn("Please issue DD60 drawings this week.", ai_prompts[0]["prompt"])
        self.assertEqual(1, len(result["suggestions"]))
        self.assertEqual("DD90", result["suggestions"][0]["deliverable"]["name"])
        self.assertEqual(["msg-1"], result["suggestions"][0]["supportingMessageIds"])
        self.assertEqual("msg-1", result["suggestions"][0]["relatedMessages"][0]["id"])
        stages = [event.get("stage") for event in progress_events]
        self.assertEqual("starting", stages[0])
        self.assertEqual("2026-03-17", progress_events[0].get("scanDate"))
        self.assertIn("listing", stages)
        self.assertIn("hydrating", stages)
        self.assertIn("preparing_ai", stages)
        self.assertIn("reviewing_ai", stages)
        self.assertIn("matching", stages)
        self.assertEqual("done", stages[-1])

    def test_scan_outlook_inbox_returns_partial_results_when_one_message_fails(self):
        ai_calls = []
        progress_events = []

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

        with patch.object(self.api, "_get_desktop_outlook_availability", return_value=(False, "")), patch.object(
            self.api, "_load_microsoft_auth_record", return_value={"accessToken": "access-token"}
        ), patch.object(
            self.api, "_refresh_microsoft_auth_record_if_needed", side_effect=lambda auth: auth
        ), patch.object(
            self.api,
            "_request_outlook_scan_batch_suggestions",
            side_effect=lambda prompt, api_key: ai_calls.append(prompt) or [],
        ), patch.object(
            self.api,
            "_notify_outlook_scan_progress",
            side_effect=lambda payload: progress_events.append(dict(payload or {})),
        ), patch.object(main_module.requests, "get", side_effect=fake_get):
            result = self.api.scan_outlook_inbox(
                {"timeframe": "week"},
                "api-key",
                "Jacob",
                ["Electrical"],
            )

        self.assertEqual("success", result["status"])
        self.assertEqual("batch", result["analysisMode"])
        self.assertEqual(1, result["candidateCount"])
        self.assertEqual(2, result["scannedCount"])
        self.assertEqual(1, result["emailsIncludedCount"])
        self.assertEqual(1, len(ai_calls))
        self.assertEqual([], result["suggestions"])
        self.assertEqual(1, len(result["skippedMessages"]))
        self.assertEqual("msg-2", result["skippedMessages"][0]["message"]["id"])
        self.assertIn("detail fetch failed", result["skippedMessages"][0]["reason"])
        hydrating_events = [
            event for event in progress_events if event.get("stage") == "hydrating"
        ]
        self.assertGreaterEqual(len(hydrating_events), 2)
        self.assertEqual(0, hydrating_events[0].get("processedEmails"))
        self.assertEqual(2, hydrating_events[-1].get("processedEmails"))

    def test_scan_outlook_inbox_defaults_missing_scan_date_to_today(self):
        today_iso = datetime.datetime.now().astimezone().date().isoformat()

        with patch.object(self.api, "_get_desktop_outlook_availability", return_value=(False, "")), patch.object(
            self.api, "_load_microsoft_auth_record", return_value={"accessToken": "access-token"}
        ), patch.object(
            self.api, "_refresh_microsoft_auth_record_if_needed", side_effect=lambda auth: auth
        ), patch.object(
            self.api, "_list_outlook_inbox_messages", return_value=([], False)
        ):
            result = self.api.scan_outlook_inbox(
                {},
                "api-key",
                "Jacob",
                ["Electrical"],
            )

        self.assertEqual("success", result["status"])
        self.assertEqual(today_iso, result["scanDate"])
        self.assertEqual(0, result["scannedCount"])

    def test_scan_outlook_inbox_invalid_scan_date_defaults_to_today(self):
        today_iso = datetime.datetime.now().astimezone().date().isoformat()

        with patch.object(self.api, "_get_desktop_outlook_availability", return_value=(False, "")), patch.object(
            self.api, "_load_microsoft_auth_record", return_value={"accessToken": "access-token"}
        ), patch.object(
            self.api, "_refresh_microsoft_auth_record_if_needed", side_effect=lambda auth: auth
        ), patch.object(
            self.api, "_list_outlook_inbox_messages", return_value=([], False)
        ) as list_messages:
            result = self.api.scan_outlook_inbox(
                {"scanDate": "not-a-date"},
                "api-key",
                "Jacob",
                ["Electrical"],
            )

        self.assertEqual("success", result["status"])
        self.assertEqual(today_iso, result["scanDate"])
        self.assertEqual(0, result["scannedCount"])
        self.assertEqual(today_iso, list_messages.call_args.kwargs["scan_date"])

    def test_list_desktop_outlook_inbox_messages_filters_to_exact_scan_date(self):
        scan_range = self.api._build_outlook_scan_range(scan_date="2026-03-17")
        start_utc = scan_range["startUtc"]
        end_utc = scan_range["endUtc"]

        class _FakeMailItem:
            def __init__(self, entry_id, received_time, subject):
                self.EntryID = entry_id
                self.ReceivedTime = received_time
                self.Subject = subject
                self.Body = subject
                self.MessageClass = "IPM.Note"
                self.Attachments = types.SimpleNamespace(Count=0)
                self.SenderName = "Project Lead"
                self.SenderEmailAddress = "lead@example.com"
                self.ConversationID = f"conv-{entry_id}"

        class _FakeItems:
            def __init__(self, items):
                self._items = list(items)
                self._index = -1

            def Sort(self, *_args, **_kwargs):
                return None

            def GetFirst(self):
                self._index = 0
                if self._index >= len(self._items):
                    return None
                return self._items[self._index]

            def GetNext(self):
                self._index += 1
                if self._index >= len(self._items):
                    return None
                return self._items[self._index]

        inbox_items = _FakeItems(
            [
                _FakeMailItem(
                    "msg-newer",
                    end_utc + datetime.timedelta(seconds=1),
                    "Newer than selected day",
                ),
                _FakeMailItem(
                    "msg-selected-1",
                    start_utc + datetime.timedelta(hours=12),
                    "Selected day message 1",
                ),
                _FakeMailItem(
                    "msg-selected-2",
                    start_utc + datetime.timedelta(hours=1),
                    "Selected day message 2",
                ),
                _FakeMailItem(
                    "msg-older",
                    start_utc - datetime.timedelta(seconds=1),
                    "Older than selected day",
                ),
            ]
        )
        fake_namespace = types.SimpleNamespace(
            GetDefaultFolder=lambda _folder_id: types.SimpleNamespace(Items=inbox_items)
        )

        with patch.object(
            self.api,
            "_run_with_outlook_com",
            side_effect=lambda callback: callback(),
        ), patch.object(
            self.api,
            "_get_desktop_outlook_namespace",
            return_value=(object(), fake_namespace),
        ), patch.object(
            self.api,
            "_get_desktop_outlook_internet_message_id",
            side_effect=lambda mail_item: f"<{mail_item.EntryID}@example.com>",
        ):
            messages, has_more = self.api._list_desktop_outlook_inbox_messages(
                scan_date="2026-03-17",
                limit=10,
            )

        self.assertFalse(has_more)
        self.assertEqual(
            ["msg-selected-1", "msg-selected-2"],
            [message["id"] for message in messages],
        )
        self.assertTrue(
            all(
                start_utc
                <= datetime.datetime.fromisoformat(
                    message["receivedDateTime"].replace("Z", "+00:00")
                )
                <= end_utc
                for message in messages
            )
        )

    def test_scan_outlook_inbox_prefers_desktop_outlook_without_microsoft_auth(self):
        desktop_summary = {
            "id": "desktop-msg-1",
            "subject": "250597 DD60 submittal",
            "bodyPreview": "Please issue DD60 this week.",
            "receivedDateTime": "2026-03-17T18:00:00Z",
            "webLink": "",
            "internetMessageId": "<desktop-msg-1@example.com>",
            "hasAttachments": False,
            "source": main_module.OUTLOOK_SCAN_SOURCE_DESKTOP,
            "from": {
                "name": "Project Lead",
                "address": "lead@example.com",
            },
        }

        batch_calls = []
        progress_events = []

        with patch.object(self.api, "_get_desktop_outlook_availability", return_value=(True, "")), patch.object(
            self.api,
            "_list_desktop_outlook_inbox_messages",
            return_value=([desktop_summary], False),
        ) as list_desktop, patch.object(
            self.api,
            "_get_desktop_outlook_message_body_text",
            return_value=(desktop_summary, "Please issue DD60 drawings this week."),
        ), patch.object(
            self.api,
            "_request_outlook_scan_batch_suggestions",
            side_effect=lambda prompt, api_key: batch_calls.append(prompt) or [
                {
                    "project": {
                        "id": "250597",
                        "name": "Client Tower",
                        "path": r"P:\250597 Client Tower",
                    },
                    "deliverable": {
                        "name": "DD60",
                        "due": "03/21/2026",
                        "notes": "Issue DD60 drawings this week.",
                        "tasks": ["Issue DD60 set"],
                    },
                    "supportingMessageRefs": ["E001"],
                }
            ],
        ), patch.object(
            self.api,
            "_notify_outlook_scan_progress",
            side_effect=lambda payload: progress_events.append(dict(payload or {})),
        ), patch.object(
            self.api, "_load_microsoft_auth_record"
        ) as load_auth:
            result = self.api.scan_outlook_inbox(
                {"timeframe": "week"},
                "api-key",
                "Jacob",
                ["Electrical"],
            )

        self.assertEqual("success", result["status"])
        self.assertEqual(main_module.OUTLOOK_SCAN_SOURCE_DESKTOP, result["source"])
        self.assertEqual(1, len(result["suggestions"]))
        self.assertEqual(1, len(batch_calls))
        self.assertEqual(1, result["candidateCount"])
        self.assertEqual(1, result["scannedCount"])
        list_desktop.assert_called_once()
        load_auth.assert_not_called()
        self.assertEqual(
            main_module.OUTLOOK_SCAN_SOURCE_DESKTOP,
            progress_events[-1]["source"],
        )

    def test_list_desktop_outlook_inbox_messages_limits_results_to_selected_day(self):
        class _FakeMailItem:
            def __init__(self, entry_id, received_time):
                self.EntryID = entry_id
                self.ReceivedTime = received_time
                self.MessageClass = "IPM.Note"

        class _FakeItems:
            def __init__(self, items):
                self._items = list(items)
                self._index = -1

            def Sort(self, *_args):
                self._items.sort(key=lambda item: item.ReceivedTime, reverse=True)

            def GetFirst(self):
                self._index = 0
                return self._items[0] if self._items else None

            def GetNext(self):
                self._index += 1
                return self._items[self._index] if self._index < len(self._items) else None

        fake_items = _FakeItems(
            [
                _FakeMailItem("msg-newer", datetime.datetime(2026, 3, 18, 9, 0, 0)),
                _FakeMailItem("msg-day-2", datetime.datetime(2026, 3, 17, 18, 0, 0)),
                _FakeMailItem("msg-day-1", datetime.datetime(2026, 3, 17, 9, 0, 0)),
                _FakeMailItem("msg-older", datetime.datetime(2026, 3, 16, 18, 0, 0)),
            ]
        )

        class _FakeNamespace:
            def GetDefaultFolder(self, _folder_id):
                return types.SimpleNamespace(Items=fake_items)

        with patch.object(
            self.api,
            "_run_with_outlook_com",
            side_effect=lambda func: func(),
        ), patch.object(
            self.api,
            "_get_desktop_outlook_namespace",
            return_value=(None, _FakeNamespace()),
        ), patch.object(
            self.api,
            "_build_desktop_outlook_message_summary",
            side_effect=lambda mail_item: {
                "id": mail_item.EntryID,
                "receivedDateTime": self.api._format_outlook_datetime(mail_item.ReceivedTime),
            },
        ):
            messages, has_more = self.api._list_desktop_outlook_inbox_messages(
                scan_date="2026-03-17",
                limit=10,
            )

        self.assertFalse(has_more)
        self.assertEqual(["msg-day-2", "msg-day-1"], [message["id"] for message in messages])

    def test_scan_outlook_inbox_falls_back_to_graph_when_desktop_outlook_fails(self):
        graph_message = {
            "id": "graph-msg-1",
            "subject": "250597 DD60 submittal",
            "bodyPreview": "Please issue DD60 this week.",
            "receivedDateTime": "2026-03-17T18:00:00Z",
            "webLink": "https://outlook.office.com/mail/deeplink/read/graph-msg-1",
            "internetMessageId": "<graph-msg-1@example.com>",
            "hasAttachments": False,
            "from": {
                "emailAddress": {
                    "name": "Project Lead",
                    "address": "lead@example.com",
                    }
                },
        }

        batch_calls = []
        progress_events = []

        with patch.object(self.api, "_get_desktop_outlook_availability", return_value=(True, "")), patch.object(
            self.api,
            "_list_desktop_outlook_inbox_messages",
            side_effect=RuntimeError("Desktop Outlook mailbox unavailable."),
        ), patch.object(
            self.api, "_load_microsoft_auth_record", return_value={"accessToken": "access-token"}
        ), patch.object(
            self.api, "_refresh_microsoft_auth_record_if_needed", side_effect=lambda auth: auth
        ), patch.object(
            self.api,
            "_list_outlook_inbox_messages",
            return_value=([graph_message], False),
        ), patch.object(
            self.api,
            "_get_outlook_message_body_text",
            return_value=(
                {
                    "id": "graph-msg-1",
                    "subject": "250597 DD60 submittal",
                    "bodyPreview": "Please issue DD60 this week.",
                    "receivedDateTime": "2026-03-17T18:00:00Z",
                    "webLink": "https://outlook.office.com/mail/deeplink/read/graph-msg-1",
                    "internetMessageId": "<graph-msg-1@example.com>",
                    "hasAttachments": False,
                    "source": main_module.OUTLOOK_SCAN_SOURCE_GRAPH,
                    "from": {
                        "name": "Project Lead",
                        "address": "lead@example.com",
                    },
                },
                "Please issue DD60 drawings this week.",
            ),
        ), patch.object(
            self.api,
            "_request_outlook_scan_batch_suggestions",
            side_effect=lambda prompt, api_key: batch_calls.append(prompt) or [
                {
                    "project": {
                        "id": "250597",
                        "name": "Client Tower",
                    },
                    "deliverable": {
                        "name": "DD60",
                        "due": "03/21/2026",
                        "notes": "Issue DD60 drawings this week.",
                    },
                    "supportingMessageRefs": ["E001"],
                }
            ],
        ), patch.object(
            self.api,
            "_notify_outlook_scan_progress",
            side_effect=lambda payload: progress_events.append(dict(payload or {})),
        ):
            result = self.api.scan_outlook_inbox(
                {"timeframe": "week"},
                "api-key",
                "Jacob",
                ["Electrical"],
            )

        self.assertEqual("success", result["status"])
        self.assertEqual(main_module.OUTLOOK_SCAN_SOURCE_GRAPH, result["source"])
        self.assertEqual(1, len(result["suggestions"]))
        self.assertEqual(1, len(batch_calls))
        stages = [event.get("stage") for event in progress_events]
        self.assertIn("fallback", stages)
        fallback_index = stages.index("fallback")
        listing_after_fallback = [
            event
            for event in progress_events[fallback_index + 1 :]
            if event.get("stage") == "listing"
        ]
        self.assertTrue(listing_after_fallback)
        self.assertEqual(
            main_module.OUTLOOK_SCAN_SOURCE_GRAPH,
            listing_after_fallback[0].get("source"),
        )

    def test_scan_outlook_inbox_reports_desktop_failure_without_graph_fallback(self):
        with patch.object(self.api, "_get_desktop_outlook_availability", return_value=(True, "")), patch.object(
            self.api,
            "_list_desktop_outlook_inbox_messages",
            side_effect=RuntimeError("Desktop Outlook mailbox unavailable."),
        ), patch.object(
            self.api, "_load_microsoft_auth_record", return_value=None
        ), patch.object(
            self.api, "_refresh_microsoft_auth_record_if_needed", return_value=None
        ), patch.object(
            self.api, "_notify_outlook_scan_progress"
        ) as notify_progress:
            result = self.api.scan_outlook_inbox(
                {"timeframe": "week"},
                "api-key",
                "Jacob",
                ["Electrical"],
            )

        self.assertEqual("error", result["status"])
        self.assertIn("Desktop Outlook is available but could not be read", result["message"])
        self.assertEqual("", result["source"])
        self.assertEqual("error", notify_progress.call_args_list[-1].args[0]["stage"])

    def test_analyze_outlook_scan_batch_trims_large_month_to_one_ai_call(self):
        message_summaries = [
            {
                "id": f"msg-{idx}",
                "subject": f"250597 DD{idx} submittal",
                "bodyPreview": "Please issue revised drawings.",
                "receivedDateTime": f"2026-03-{18 - idx:02d}T18:00:00Z",
                "webLink": f"https://outlook.office.com/mail/deeplink/read/msg-{idx}",
                "internetMessageId": f"<msg-{idx}@example.com>",
                "hasAttachments": True,
                "source": main_module.OUTLOOK_SCAN_SOURCE_GRAPH,
                "from": {
                    "name": "Project Lead",
                    "address": "lead@example.com",
                },
            }
            for idx in range(1, 5)
        ]
        project_context = [
            {
                "id": "250597",
                "name": "Client Tower",
                "nick": "Tower",
                "path": r"P:\250597 Client Tower",
                "deliverables": [
                    {
                        "name": f"Deliverable {idx}",
                        "due": f"03/{idx:02d}/2026",
                        "status": "Working",
                    }
                    for idx in range(1, 9)
                ],
            }
        ]
        ai_prompts = []

        def body_loader(summary):
            return summary, "Please review and issue revised drawings. " * 20

        def fake_batch(prompt, api_key):
            ai_prompts.append(prompt)
            return []

        with patch.object(main_module, "OUTLOOK_SCAN_PROMPT_EMAIL_BUDGET_CHARS", 1200), patch.object(
            main_module, "OUTLOOK_SCAN_PROMPT_DELIVERABLE_BUDGET_CHARS", 350
        ), patch.object(
            main_module, "OUTLOOK_SCAN_PROMPT_MAX_CHARS", 1700
        ), patch.object(
            self.api, "_request_outlook_scan_batch_suggestions", side_effect=fake_batch
        ):
            result = self.api._analyze_outlook_scan_batch(
                message_summaries,
                body_loader,
                project_context,
                "api-key",
                "Jacob",
                ["Electrical"],
                "month",
                main_module.OUTLOOK_SCAN_SOURCE_GRAPH,
                has_more=False,
            )

        self.assertEqual("success", result["status"])
        self.assertEqual(1, len(ai_prompts))
        self.assertTrue(result["promptTruncated"])
        self.assertLess(result["emailsIncludedCount"], len(message_summaries))
        self.assertLess(result["deliverablesIncludedCount"], len(project_context[0]["deliverables"]))
        self.assertGreaterEqual(len(result["skippedMessages"]), 1)

    def test_analyze_outlook_scan_batch_handles_zero_email_progress(self):
        progress_events = []

        with patch.object(
            self.api,
            "_notify_outlook_scan_progress",
            side_effect=lambda payload: progress_events.append(dict(payload or {})),
        ):
            result = self.api._analyze_outlook_scan_batch(
                [],
                lambda summary: (summary, ""),
                [],
                "api-key",
                "Jacob",
                ["Electrical"],
                "week",
                main_module.OUTLOOK_SCAN_SOURCE_GRAPH,
                has_more=False,
            )

        self.assertEqual("success", result["status"])
        self.assertEqual(0, result["scannedCount"])
        self.assertEqual("hydrating", progress_events[0]["stage"])
        self.assertEqual("done", progress_events[-1]["stage"])
        self.assertEqual(0, progress_events[-1]["totalEmails"])

    def test_build_desktop_outlook_message_summary_includes_conversation_id(self):
        class _FakeMailItem:
            Body = "Please review the updated drawings."
            EntryID = "desktop-msg-1"
            Subject = "250597 DD60 submittal"
            ReceivedTime = datetime.datetime(2026, 3, 17, 18, 0, tzinfo=datetime.timezone.utc)
            ConversationID = "desktop-conv-1"
            Attachments = types.SimpleNamespace(Count=0)
            SenderName = "Project Lead"
            SenderEmailAddress = "lead@example.com"

        with patch.object(
            self.api,
            "_get_desktop_outlook_internet_message_id",
            return_value="<desktop-msg-1@example.com>",
        ):
            summary = self.api._build_desktop_outlook_message_summary(_FakeMailItem())

        self.assertEqual("desktop-conv-1", summary["conversationId"])
        self.assertEqual("desktop-msg-1", summary["id"])

    def test_analyze_outlook_scan_batch_keeps_distinct_replies_but_removes_thread_history(self):
        message_summaries = [
            {
                "id": "desktop-msg-new",
                "subject": "250597 DD90 submittal",
                "bodyPreview": "Need DD90 by Friday.",
                "receivedDateTime": "2026-03-18T18:00:00Z",
                "webLink": "",
                "internetMessageId": "<desktop-msg-new@example.com>",
                "conversationId": "desktop-conv-1",
                "hasAttachments": False,
                "source": main_module.OUTLOOK_SCAN_SOURCE_DESKTOP,
                "from": {
                    "name": "Project Lead",
                    "address": "lead@example.com",
                },
            },
            {
                "id": "desktop-msg-old",
                "subject": "250597 DD60 submittal",
                "bodyPreview": "Please issue DD60 drawings this week.",
                "receivedDateTime": "2026-03-17T18:00:00Z",
                "webLink": "",
                "internetMessageId": "<desktop-msg-old@example.com>",
                "conversationId": "desktop-conv-1",
                "hasAttachments": False,
                "source": main_module.OUTLOOK_SCAN_SOURCE_DESKTOP,
                "from": {
                    "name": "Project Lead",
                    "address": "lead@example.com",
                },
            },
        ]
        ai_prompts = []
        body_map = {
            "desktop-msg-new": (
                "Need DD90 by Friday.\n\n"
                "-----Original Message-----\n"
                "Please issue DD60 drawings this week."
            ),
            "desktop-msg-old": "Please issue DD60 drawings this week.",
        }

        def body_loader(summary):
            return summary, body_map[summary["id"]]

        with patch.object(
            self.api,
            "_request_outlook_scan_batch_suggestions",
            side_effect=lambda prompt, api_key: ai_prompts.append(prompt) or [],
        ):
            result = self.api._analyze_outlook_scan_batch(
                message_summaries,
                body_loader,
                [],
                "api-key",
                "Jacob",
                ["Electrical"],
                "week",
                main_module.OUTLOOK_SCAN_SOURCE_DESKTOP,
                has_more=False,
            )

        self.assertEqual("success", result["status"])
        self.assertEqual(1, len(ai_prompts))
        self.assertEqual(1, result["threadsDetected"])
        self.assertEqual(1, result["dedupedEmailCount"])
        self.assertEqual(0, result["dedupeSkippedEmailCount"])
        self.assertEqual(2, result["emailsIncludedCount"])
        self.assertIn("Need DD90 by Friday.", ai_prompts[0])
        self.assertEqual(1, ai_prompts[0].count("Please issue DD60 drawings this week."))

    def test_analyze_outlook_scan_batch_skips_reply_without_unique_thread_content(self):
        message_summaries = [
            {
                "id": "msg-new",
                "subject": "250597 DD60 submittal",
                "bodyPreview": "Please issue DD60 drawings this week.",
                "receivedDateTime": "2026-03-18T18:00:00Z",
                "webLink": "https://outlook.office.com/mail/deeplink/read/msg-new",
                "internetMessageId": "<msg-new@example.com>",
                "conversationId": "conv-1",
                "hasAttachments": False,
                "source": main_module.OUTLOOK_SCAN_SOURCE_GRAPH,
                "from": {
                    "name": "Project Lead",
                    "address": "lead@example.com",
                },
            },
            {
                "id": "msg-old",
                "subject": "250597 DD60 submittal",
                "bodyPreview": "Please issue DD60 drawings this week.",
                "receivedDateTime": "2026-03-17T18:00:00Z",
                "webLink": "https://outlook.office.com/mail/deeplink/read/msg-old",
                "internetMessageId": "<msg-old@example.com>",
                "conversationId": "conv-1",
                "hasAttachments": False,
                "source": main_module.OUTLOOK_SCAN_SOURCE_GRAPH,
                "from": {
                    "name": "Project Lead",
                    "address": "lead@example.com",
                },
            },
        ]
        ai_prompts = []

        def body_loader(summary):
            return summary, "Please issue DD60 drawings this week."

        with patch.object(
            self.api,
            "_request_outlook_scan_batch_suggestions",
            side_effect=lambda prompt, api_key: ai_prompts.append(prompt) or [],
        ):
            result = self.api._analyze_outlook_scan_batch(
                message_summaries,
                body_loader,
                [],
                "api-key",
                "Jacob",
                ["Electrical"],
                "week",
                main_module.OUTLOOK_SCAN_SOURCE_GRAPH,
                has_more=False,
            )

        self.assertEqual("success", result["status"])
        self.assertEqual(1, len(ai_prompts))
        self.assertEqual(1, result["emailsIncludedCount"])
        self.assertEqual(1, result["dedupeSkippedEmailCount"])
        self.assertEqual(
            main_module.OUTLOOK_SCAN_DEDUPE_SKIP_REASON_THREAD,
            result["skippedMessages"][0]["reason"],
        )

    def test_analyze_outlook_scan_batch_without_conversation_id_only_dedupes_exact_content(self):
        message_summaries = [
            {
                "id": "msg-1",
                "subject": "Re: 250597 deliverable update",
                "bodyPreview": "Please issue DD60 drawings this week.",
                "receivedDateTime": "2026-03-18T18:00:00Z",
                "webLink": "https://outlook.office.com/mail/deeplink/read/msg-1",
                "internetMessageId": "<msg-1@example.com>",
                "conversationId": "",
                "hasAttachments": False,
                "source": main_module.OUTLOOK_SCAN_SOURCE_GRAPH,
                "from": {
                    "name": "Project Lead",
                    "address": "lead@example.com",
                },
            },
            {
                "id": "msg-2",
                "subject": "Re: 250597 deliverable update",
                "bodyPreview": "Please issue DD60 drawings this week.",
                "receivedDateTime": "2026-03-17T18:00:00Z",
                "webLink": "https://outlook.office.com/mail/deeplink/read/msg-2",
                "internetMessageId": "<msg-2@example.com>",
                "conversationId": "",
                "hasAttachments": False,
                "source": main_module.OUTLOOK_SCAN_SOURCE_GRAPH,
                "from": {
                    "name": "Project Lead",
                    "address": "lead@example.com",
                },
            },
            {
                "id": "msg-3",
                "subject": "Re: 250597 deliverable update",
                "bodyPreview": "Please issue DD90 drawings this week.",
                "receivedDateTime": "2026-03-16T18:00:00Z",
                "webLink": "https://outlook.office.com/mail/deeplink/read/msg-3",
                "internetMessageId": "<msg-3@example.com>",
                "conversationId": "",
                "hasAttachments": False,
                "source": main_module.OUTLOOK_SCAN_SOURCE_GRAPH,
                "from": {
                    "name": "Project Lead",
                    "address": "lead@example.com",
                },
            },
        ]
        ai_prompts = []
        body_map = {
            "msg-1": "Please issue DD60 drawings this week.",
            "msg-2": "Please issue DD60 drawings this week.",
            "msg-3": "Please issue DD90 drawings this week.",
        }

        def body_loader(summary):
            return summary, body_map[summary["id"]]

        with patch.object(
            self.api,
            "_request_outlook_scan_batch_suggestions",
            side_effect=lambda prompt, api_key: ai_prompts.append(prompt) or [],
        ):
            result = self.api._analyze_outlook_scan_batch(
                message_summaries,
                body_loader,
                [],
                "api-key",
                "Jacob",
                ["Electrical"],
                "week",
                main_module.OUTLOOK_SCAN_SOURCE_GRAPH,
                has_more=False,
            )

        self.assertEqual("success", result["status"])
        self.assertEqual(1, len(ai_prompts))
        self.assertEqual(0, result["threadsDetected"])
        self.assertEqual(2, result["emailsIncludedCount"])
        self.assertEqual(1, result["dedupeSkippedEmailCount"])
        self.assertEqual(
            main_module.OUTLOOK_SCAN_DEDUPE_SKIP_REASON_DUPLICATE,
            result["skippedMessages"][0]["reason"],
        )
        self.assertEqual(1, ai_prompts[0].count("Please issue DD60 drawings this week."))
        self.assertEqual(1, ai_prompts[0].count("Please issue DD90 drawings this week."))

    def test_normalize_outlook_scan_batch_suggestions_handles_empty_and_malformed_payloads(self):
        normalized = self.api._normalize_outlook_scan_batch_suggestions(
            {
                "suggestions": [
                    {
                        "projectId": "250597",
                        "projectName": "Client Tower",
                        "projectPath": r"P:\250597 Client Tower",
                        "deliverable": {
                            "name": "DD60",
                            "due": "03/21/2026",
                            "notes": "Issue DD60.",
                            "tasks": ["Issue DD60 set"],
                        },
                        "supportingMessageRefs": ["E001"],
                    }
                ]
            }
        )

        self.assertEqual([], self.api._normalize_outlook_scan_batch_suggestions({"suggestions": []}))
        self.assertEqual("250597", normalized[0]["project"]["id"])
        self.assertEqual("DD60", normalized[0]["deliverable"]["name"])
        self.assertEqual("Issue DD60 set", normalized[0]["deliverable"]["tasks"][0]["text"])
        with self.assertRaises(RuntimeError):
            self.api._normalize_outlook_scan_batch_suggestions({"foo": "bar"})

    def test_open_outlook_desktop_message_displays_mail_item(self):
        opened = {}

        class _FakeMailItem:
            def Display(self):
                opened["displayed"] = True

        class _FakeNamespace:
            def GetItemFromID(self, message_id):
                opened["message_id"] = message_id
                return _FakeMailItem()

        with patch.object(
            self.api,
            "_run_with_outlook_com",
            side_effect=lambda func: func(),
        ), patch.object(
            self.api,
            "_get_desktop_outlook_namespace",
            return_value=(object(), _FakeNamespace()),
        ):
            result = self.api.open_outlook_desktop_message("desktop-msg-1")

        self.assertEqual("success", result["status"])
        self.assertEqual("desktop-msg-1", opened["message_id"])
        self.assertTrue(opened["displayed"])


if __name__ == "__main__":
    unittest.main()
