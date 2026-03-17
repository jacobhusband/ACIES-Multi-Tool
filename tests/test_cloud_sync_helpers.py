import json
import os
import sys
import tempfile
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


class CloudSyncHelperTests(unittest.TestCase):
    def setUp(self):
        self.api = Api.__new__(Api)

    def test_build_default_user_settings_includes_cloud_sync_defaults(self):
        settings = main_module.build_default_user_settings()

        self.assertIn("cloudSync", settings)
        self.assertEqual(
            main_module.build_default_cloud_sync_settings(), settings["cloudSync"]
        )

    def test_build_google_auth_record_preserves_id_token(self):
        existing_auth = {"idToken": "existing-id-token", "refreshToken": "refresh-token"}
        token_payload = {
            "access_token": "access-token",
            "expires_in": 3600,
        }
        profile_payload = {"sub": "subject-1", "email": "user@example.com"}

        record = self.api._build_google_auth_record(
            token_payload, profile_payload, existing_auth=existing_auth
        )

        self.assertEqual("existing-id-token", record["idToken"])
        self.assertEqual("refresh-token", record["refreshToken"])
        self.assertEqual("access-token", record["accessToken"])

    def test_get_cloud_sync_config_reports_enabled_only_with_required_env(self):
        required_env = {
            main_module.FIREBASE_API_KEY_ENV: "api-key",
            main_module.FIREBASE_AUTH_DOMAIN_ENV: "project.firebaseapp.com",
            main_module.FIREBASE_PROJECT_ID_ENV: "project-id",
            main_module.FIREBASE_APP_ID_ENV: "app-id",
            main_module.FIREBASE_STORAGE_BUCKET_ENV: "",
            main_module.FIREBASE_MESSAGING_SENDER_ID_ENV: "",
        }
        with patch.dict(main_module.os.environ, required_env, clear=False):
            enabled_result = self.api.get_cloud_sync_config()

        self.assertEqual("success", enabled_result["status"])
        self.assertTrue(enabled_result["enabled"])
        self.assertEqual("api-key", enabled_result["config"]["apiKey"])

        missing_env = dict(required_env)
        missing_env[main_module.FIREBASE_PROJECT_ID_ENV] = ""
        with patch.dict(main_module.os.environ, missing_env, clear=False):
            disabled_result = self.api.get_cloud_sync_config()

        self.assertEqual("success", disabled_result["status"])
        self.assertFalse(disabled_result["enabled"])

    def test_get_google_sync_session_returns_tokens_and_auth_state(self):
        auth_record = {
            "provider": "google",
            "email": "user@example.com",
            "displayName": "User Example",
            "avatarUrl": "https://example.com/avatar.png",
            "signedInAt": "2026-01-01T00:00:00Z",
            "expiresAt": "2026-01-01T01:00:00Z",
            "idToken": "id-token",
            "accessToken": "access-token",
            "refreshToken": "refresh-token",
        }

        with patch.object(self.api, "_load_google_auth_record", return_value=auth_record), patch.object(
            self.api, "_refresh_google_auth_record_if_needed", return_value=auth_record
        ):
            result = self.api.get_google_sync_session()

        self.assertEqual("success", result["status"])
        self.assertTrue(result["signedIn"])
        self.assertTrue(result["firebaseReady"])
        self.assertEqual("id-token", result["idToken"])
        self.assertEqual("access-token", result["accessToken"])
        self.assertEqual("user@example.com", result["auth"]["email"])

    def test_create_cloud_sync_backup_copies_tracked_files_and_metadata(self):
        with tempfile.TemporaryDirectory() as tempdir:
            backups_dir = os.path.join(tempdir, "sync_backups")
            os.makedirs(backups_dir, exist_ok=True)
            tracked_files = {}
            for name in ("settings", "tasks", "notes"):
                file_path = os.path.join(tempdir, f"{name}.json")
                with open(file_path, "w", encoding="utf-8") as handle:
                    json.dump({"name": name}, handle)
                tracked_files[name] = file_path

            with patch.object(main_module, "SYNC_BACKUPS_DIR", backups_dir), patch.object(
                main_module, "SYNC_TRACKED_FILES", tracked_files
            ):
                result = main_module._create_cloud_sync_backup(
                    reason="remote overwrite",
                    metadata={"firebaseUid": "user-123"},
                )

            self.assertTrue(os.path.isdir(result["path"]))
            self.assertEqual(3, len(result["files"]))
            metadata_path = os.path.join(result["path"], "metadata.json")
            self.assertTrue(os.path.exists(metadata_path))

            with open(metadata_path, "r", encoding="utf-8") as handle:
                metadata = json.load(handle)

            self.assertEqual("remote overwrite", metadata["reason"])
            self.assertEqual("user-123", metadata["metadata"]["firebaseUid"])
            self.assertEqual(3, len(metadata["files"]))


if __name__ == "__main__":
    unittest.main()
