import datetime
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


def _summary(
    entry_id="entry-1",
    internet_message_id="<msg-1@example.com>",
    subject="RE: Panel schedule",
    received="2026-07-08T15:00:00Z",
    to_recipients=None,
    cc_recipients=None,
    body_preview="Please advise on the panel schedule.",
    sender_name="Jane Architect",
    sender_address="jane@arch.com",
):
    return {
        "id": entry_id,
        "subject": subject,
        "bodyPreview": body_preview,
        "receivedDateTime": received,
        "webLink": "",
        "internetMessageId": internet_message_id,
        "conversationId": "conv-1",
        "hasAttachments": False,
        "source": "desktop-outlook",
        "from": {"name": sender_name, "address": sender_address},
        "toRecipients": to_recipients or [],
        "ccRecipients": cc_recipients or [],
    }


OWNER = {
    "addresses": ["jacob@acies.com"],
    "names": ["Jacob Husband"],
}


class EmailCaptureDbTestCase(unittest.TestCase):
    def setUp(self):
        self._temp_dir = tempfile.TemporaryDirectory()
        self._db_patch = patch.object(
            main_module,
            "EMAIL_CAPTURE_DB_FILE",
            os.path.join(self._temp_dir.name, "email_capture.db"),
        )
        self._db_patch.start()
        self.addCleanup(self._db_patch.stop)
        self.addCleanup(self._temp_dir.cleanup)


class EmailCaptureSchemaTests(EmailCaptureDbTestCase):
    def test_schema_created_idempotently_with_version_meta(self):
        for _ in range(2):
            conn = main_module._open_email_capture_db()
            try:
                tables = {
                    row["name"]
                    for row in conn.execute(
                        "SELECT name FROM sqlite_master WHERE type = 'table'"
                    ).fetchall()
                }
                self.assertIn("captured_emails", tables)
                self.assertIn("follow_ups", tables)
                self.assertIn("email_capture_meta", tables)
                version = main_module._get_email_capture_meta(conn, "schema_version")
                self.assertEqual(
                    str(main_module.EMAIL_CAPTURE_SCHEMA_VERSION), version
                )
            finally:
                conn.close()

    def test_meta_set_and_get_round_trip(self):
        conn = main_module._open_email_capture_db()
        try:
            main_module._set_email_capture_meta(conn, "last_scan_status", "success")
            main_module._set_email_capture_meta(conn, "last_scan_status", "error")
            conn.commit()
            self.assertEqual(
                "error", main_module._get_email_capture_meta(conn, "last_scan_status")
            )
            self.assertEqual(
                "fallback",
                main_module._get_email_capture_meta(conn, "missing-key", "fallback"),
            )
        finally:
            conn.close()


class CapturedEmailUpsertTests(EmailCaptureDbTestCase):
    def test_same_internet_message_id_upserts_single_row(self):
        conn = main_module._open_email_capture_db()
        try:
            first = main_module._upsert_captured_email(
                conn, _summary(), "to", "2026-07-08T16:00:00Z"
            )
            second = main_module._upsert_captured_email(
                conn,
                _summary(entry_id="entry-1-moved"),
                "to",
                "2026-07-09T16:00:00Z",
            )
            conn.commit()
            self.assertEqual("inserted", first)
            self.assertEqual("updated", second)
            rows = conn.execute("SELECT * FROM captured_emails").fetchall()
            self.assertEqual(1, len(rows))
            self.assertEqual("entry-1-moved", rows[0]["entry_id"])
            self.assertEqual("2026-07-09T16:00:00Z", rows[0]["last_scan_at_utc"])
        finally:
            conn.close()

    def test_blank_internet_message_id_falls_back_to_entry_id(self):
        conn = main_module._open_email_capture_db()
        try:
            main_module._upsert_captured_email(
                conn,
                _summary(internet_message_id=""),
                "cc",
                "2026-07-08T16:00:00Z",
            )
            outcome = main_module._upsert_captured_email(
                conn,
                _summary(internet_message_id=""),
                "cc",
                "2026-07-09T16:00:00Z",
            )
            conn.commit()
            self.assertEqual("updated", outcome)
            rows = conn.execute("SELECT * FROM captured_emails").fetchall()
            self.assertEqual(1, len(rows))
        finally:
            conn.close()

    def test_reseen_email_never_regresses_status(self):
        conn = main_module._open_email_capture_db()
        try:
            main_module._upsert_captured_email(
                conn, _summary(), "to", "2026-07-08T16:00:00Z"
            )
            conn.execute(
                "UPDATE captured_emails SET status = 'accepted', ai_status = 'matched'"
            )
            main_module._upsert_captured_email(
                conn, _summary(), "to", "2026-07-09T16:00:00Z"
            )
            conn.commit()
            row = conn.execute("SELECT * FROM captured_emails").fetchone()
            self.assertEqual("accepted", row["status"])
            self.assertEqual("matched", row["ai_status"])
        finally:
            conn.close()


class DirectednessTests(unittest.TestCase):
    def test_owner_address_in_to_line(self):
        summary = _summary(to_recipients=[{"name": "", "address": "Jacob@ACIES.com"}])
        self.assertEqual(
            "to", main_module._classify_email_directedness(summary, OWNER)
        )

    def test_owner_display_name_in_to_line(self):
        summary = _summary(to_recipients=[{"name": "Jacob Husband", "address": ""}])
        self.assertEqual(
            "to", main_module._classify_email_directedness(summary, OWNER)
        )

    def test_first_name_in_body_counts_as_named(self):
        summary = _summary(
            to_recipients=[{"name": "Someone Else", "address": "other@x.com"}],
            body_preview="Hi Jacob, can you confirm the feeder sizes?",
        )
        self.assertEqual(
            "named", main_module._classify_email_directedness(summary, OWNER)
        )

    def test_cc_only_is_cc(self):
        summary = _summary(
            to_recipients=[{"name": "Someone Else", "address": "other@x.com"}],
            cc_recipients=[{"name": "", "address": "jacob@acies.com"}],
            body_preview="Team, see attached for coordination.",
        )
        self.assertEqual(
            "cc", main_module._classify_email_directedness(summary, OWNER)
        )

    def test_no_identity_match_is_unknown(self):
        summary = _summary(
            to_recipients=[{"name": "Someone Else", "address": "other@x.com"}],
            body_preview="Distribution list update.",
        )
        self.assertEqual(
            "unknown", main_module._classify_email_directedness(summary, OWNER)
        )

    def test_partial_name_token_does_not_match(self):
        summary = _summary(
            to_recipients=[],
            body_preview="The jacobean style facade is unchanged.",
        )
        self.assertEqual(
            "unknown", main_module._classify_email_directedness(summary, OWNER)
        )


class ScanWindowTests(EmailCaptureDbTestCase):
    def setUp(self):
        super().setUp()
        self.api = Api.__new__(Api)

    def test_first_scan_window_is_first_scan_days(self):
        conn = main_module._open_email_capture_db()
        try:
            start_utc, end_utc = self.api._resolve_email_capture_scan_window(conn)
        finally:
            conn.close()
        span = end_utc - start_utc
        self.assertAlmostEqual(
            main_module.EMAIL_CAPTURE_FIRST_SCAN_DAYS,
            span.total_seconds() / 86400,
            places=2,
        )

    def test_subsequent_window_starts_at_watermark_minus_overlap(self):
        conn = main_module._open_email_capture_db()
        try:
            watermark = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(
                days=1
            )
            main_module._set_email_capture_meta(
                conn,
                "last_scan_completed_at_utc",
                watermark.isoformat().replace("+00:00", "Z"),
            )
            conn.commit()
            start_utc, _ = self.api._resolve_email_capture_scan_window(conn)
        finally:
            conn.close()
        expected_start = watermark - datetime.timedelta(
            minutes=main_module.EMAIL_CAPTURE_SCAN_OVERLAP_MINUTES
        )
        self.assertAlmostEqual(
            0.0, abs((start_utc - expected_start).total_seconds()), places=0
        )

    def test_stale_watermark_clamped_to_max_window(self):
        conn = main_module._open_email_capture_db()
        try:
            watermark = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(
                days=90
            )
            main_module._set_email_capture_meta(
                conn,
                "last_scan_completed_at_utc",
                watermark.isoformat().replace("+00:00", "Z"),
            )
            conn.commit()
            start_utc, end_utc = self.api._resolve_email_capture_scan_window(conn)
        finally:
            conn.close()
        span = end_utc - start_utc
        self.assertLessEqual(
            span.total_seconds() / 86400,
            main_module.EMAIL_CAPTURE_MAX_WINDOW_DAYS + 0.01,
        )


class RunEmailCaptureScanTests(EmailCaptureDbTestCase):
    def setUp(self):
        super().setUp()
        self.api = Api.__new__(Api)
        self.api._email_capture_scan_running = False
        self.api._email_capture_scan_lock = main_module.threading.Lock()

    def _run_scan(
        self,
        summaries,
        api_key="",
        analysis=None,
        availability=(True, ""),
        list_error=None,
    ):
        def _list_messages(*args, **kwargs):
            if list_error is not None:
                raise list_error
            return summaries, False

        patches = [
            patch.object(
                self.api, "_get_desktop_outlook_availability", return_value=availability
            ),
            patch.object(
                self.api,
                "_list_desktop_outlook_inbox_messages",
                side_effect=_list_messages,
            ),
            patch.object(
                self.api, "_resolve_outlook_current_user_smtp", return_value="jacob@acies.com"
            ),
            patch.object(
                self.api,
                "get_user_settings",
                return_value={"userName": "Jacob Husband", "googleAuth": None},
            ),
            patch.object(self.api, "_send_email_capture_progress"),
        ]
        if analysis is not None:
            patches.append(
                patch.object(
                    self.api, "_analyze_outlook_scan_batch", return_value=analysis
                )
            )
        if not api_key:
            patches.append(
                patch.object(
                    self.api,
                    "_resolve_google_ai_api_key",
                    side_effect=RuntimeError("AI API key is not configured."),
                )
            )
        started = [p.start() for p in patches]
        try:
            return self.api.run_email_capture_scan(
                {"mode": "auto", "projectContext": []},
                api_key,
                "Jacob Husband",
                ["Electrical"],
            )
        finally:
            for p in patches:
                p.stop()
            del started

    def test_capture_persists_rows_and_advances_watermark(self):
        summaries = [
            _summary(
                entry_id="entry-1",
                internet_message_id="<msg-1@example.com>",
                to_recipients=[{"name": "", "address": "jacob@acies.com"}],
            ),
            _summary(
                entry_id="entry-2",
                internet_message_id="<msg-2@example.com>",
                subject="FYI meeting notes",
                cc_recipients=[{"name": "", "address": "jacob@acies.com"}],
                body_preview="For your awareness only.",
            ),
        ]
        result = self._run_scan(summaries)

        self.assertEqual("success", result["status"])
        self.assertEqual(2, result["newCount"])
        self.assertEqual(1, result["ccCount"])
        self.assertEqual(2, result["pendingAiCount"])

        conn = main_module._open_email_capture_db()
        try:
            rows = conn.execute(
                "SELECT * FROM captured_emails ORDER BY entry_id"
            ).fetchall()
            self.assertEqual(2, len(rows))
            self.assertEqual("to", rows[0]["directedness"])
            self.assertEqual("cc", rows[1]["directedness"])
            watermark = main_module._get_email_capture_meta(
                conn, "last_scan_completed_at_utc"
            )
            self.assertTrue(watermark)
        finally:
            conn.close()

    def test_listing_failure_does_not_advance_watermark(self):
        result = self._run_scan([], list_error=RuntimeError("COM exploded"))
        self.assertEqual("error", result["status"])
        conn = main_module._open_email_capture_db()
        try:
            watermark = main_module._get_email_capture_meta(
                conn, "last_scan_completed_at_utc"
            )
            self.assertEqual("", watermark)
        finally:
            conn.close()

    def test_unavailable_outlook_returns_quiet_error(self):
        result = self._run_scan(
            [], availability=(False, "Classic Outlook is not installed on this machine.")
        )
        self.assertEqual("error", result["status"])
        self.assertIn("Classic Outlook", result["message"])

    def test_concurrent_scan_returns_skipped(self):
        self.api._email_capture_scan_running = True
        result = self.api.run_email_capture_scan({"mode": "auto"}, "", "", [])
        self.assertEqual({"status": "skipped", "reason": "scan-in-progress"}, result)

    def test_ai_stage_maps_suggestions_to_rows(self):
        summaries = [
            _summary(
                entry_id="entry-1",
                internet_message_id="<msg-1@example.com>",
                to_recipients=[{"name": "", "address": "jacob@acies.com"}],
            ),
            _summary(
                entry_id="entry-2",
                internet_message_id="<msg-2@example.com>",
                subject="Spam-ish",
                body_preview="Nothing project related.",
            ),
            _summary(
                entry_id="entry-3",
                internet_message_id="<msg-3@example.com>",
                subject="Old thread reply",
            ),
        ]
        analysis = {
            "status": "success",
            "suggestions": [
                {
                    "project": {"id": "p1", "name": "Fire Station 12", "nick": "", "path": ""},
                    "deliverable": {
                        "name": "CD90 panel schedule",
                        "due": "2026-07-20",
                        "notes": "",
                        "tasks": [],
                    },
                    "supportingMessageIds": ["entry-1"],
                    "relatedMessages": [],
                }
            ],
            "skippedMessages": [
                {
                    "message": {"id": "entry-3"},
                    "reason": main_module.OUTLOOK_SCAN_DEDUPE_SKIP_REASON_THREAD,
                }
            ],
        }
        result = self._run_scan(summaries, api_key="test-key", analysis=analysis)

        self.assertEqual("success", result["status"])
        self.assertEqual(1, result["matchedCount"])
        conn = main_module._open_email_capture_db()
        try:
            statuses = {
                row["entry_id"]: row["ai_status"]
                for row in conn.execute("SELECT * FROM captured_emails").fetchall()
            }
            self.assertEqual("matched", statuses["entry-1"])
            self.assertEqual("no_match", statuses["entry-2"])
            self.assertEqual("skipped", statuses["entry-3"])
        finally:
            conn.close()

    def test_ai_failure_keeps_rows_pending_and_capture_succeeds(self):
        summaries = [_summary()]

        def _boom(*args, **kwargs):
            raise RuntimeError("AI timed out")

        with patch.object(self.api, "_analyze_outlook_scan_batch", side_effect=_boom):
            result = self._run_scan(summaries, api_key="test-key")

        self.assertEqual("success", result["status"])
        self.assertEqual("AI timed out", result.get("aiError"))
        conn = main_module._open_email_capture_db()
        try:
            row = conn.execute("SELECT * FROM captured_emails").fetchone()
            self.assertEqual("pending", row["ai_status"])
            watermark = main_module._get_email_capture_meta(
                conn, "last_scan_completed_at_utc"
            )
            self.assertTrue(watermark)
        finally:
            conn.close()

    def test_budget_trimmed_rows_stay_pending(self):
        summaries = [
            _summary(entry_id="entry-1", internet_message_id="<msg-1@example.com>"),
            _summary(entry_id="entry-2", internet_message_id="<msg-2@example.com>"),
        ]
        analysis = {
            "status": "success",
            "suggestions": [],
            "skippedMessages": [
                {
                    "message": {"id": "entry-2"},
                    "reason": main_module.OUTLOOK_SCAN_PROMPT_SKIP_REASON,
                }
            ],
        }
        self._run_scan(summaries, api_key="test-key", analysis=analysis)

        conn = main_module._open_email_capture_db()
        try:
            statuses = {
                row["entry_id"]: row["ai_status"]
                for row in conn.execute("SELECT * FROM captured_emails").fetchall()
            }
            self.assertEqual("no_match", statuses["entry-1"])
            self.assertEqual("pending", statuses["entry-2"])
        finally:
            conn.close()


class CapturedEmailCrudTests(EmailCaptureDbTestCase):
    def setUp(self):
        super().setUp()
        self.api = Api.__new__(Api)
        conn = main_module._open_email_capture_db()
        try:
            main_module._upsert_captured_email(
                conn,
                _summary(
                    to_recipients=[{"name": "", "address": "jacob@acies.com"}]
                ),
                "to",
                "2026-07-08T16:00:00Z",
            )
            main_module._upsert_captured_email(
                conn,
                _summary(
                    entry_id="entry-2",
                    internet_message_id="<msg-2@example.com>",
                    subject="FYI",
                ),
                "cc",
                "2026-07-08T16:00:00Z",
            )
            conn.commit()
        finally:
            conn.close()

    def test_list_captured_emails_filters_and_counts(self):
        result = self.api.list_captured_emails({"status": "new"})
        self.assertEqual("success", result["status"])
        self.assertEqual(2, len(result["items"]))
        self.assertEqual(2, result["counts"]["new"])
        self.assertEqual(1, result["counts"]["newDirected"])
        self.assertEqual(1, result["counts"]["newCc"])
        self.assertEqual(2, result["counts"]["pendingAi"])

        directed = self.api.list_captured_emails(
            {"status": "new", "directedness": "directed"}
        )
        self.assertEqual(1, len(directed["items"]))
        self.assertEqual("to", directed["items"][0]["directedness"])

    def test_status_transitions_enforced(self):
        listing = self.api.list_captured_emails({"status": "new"})
        record_id = listing["items"][0]["id"]

        accepted = self.api.update_captured_email_status(
            {"id": record_id, "status": "accepted", "acceptedProjectId": "p1"}
        )
        self.assertEqual("success", accepted["status"])
        self.assertEqual("accepted", accepted["item"]["status"])
        self.assertEqual("p1", accepted["item"]["acceptedProjectId"])

        invalid = self.api.update_captured_email_status(
            {"id": record_id, "status": "dismissed"}
        )
        self.assertEqual("error", invalid["status"])

    def test_dismiss_and_restore_round_trip(self):
        listing = self.api.list_captured_emails({"status": "new"})
        record_id = listing["items"][0]["id"]

        dismissed = self.api.update_captured_email_status(
            {"id": record_id, "status": "dismissed"}
        )
        self.assertEqual("dismissed", dismissed["item"]["status"])

        restored = self.api.update_captured_email_status(
            {"id": record_id, "status": "new"}
        )
        self.assertEqual("new", restored["item"]["status"])


class FollowUpTests(EmailCaptureDbTestCase):
    def setUp(self):
        super().setUp()
        self.api = Api.__new__(Api)
        conn = main_module._open_email_capture_db()
        try:
            main_module._upsert_captured_email(
                conn, _summary(), "to", "2026-07-08T16:00:00Z"
            )
            conn.commit()
            self.email_id = conn.execute(
                "SELECT id FROM captured_emails"
            ).fetchone()["id"]
        finally:
            conn.close()

    def test_create_follow_up_defaults_title_to_email_subject(self):
        result = self.api.create_follow_up(
            {"capturedEmailId": self.email_id, "waitingOn": "Jane Architect"}
        )
        self.assertEqual("success", result["status"])
        self.assertEqual("RE: Panel schedule", result["item"]["title"])
        self.assertEqual("Jane Architect", result["item"]["waitingOn"])
        self.assertEqual("open", result["item"]["status"])

    def test_create_custom_follow_up_requires_title(self):
        missing = self.api.create_follow_up({"kind": "custom"})
        self.assertEqual("error", missing["status"])

    def test_snooze_sets_snoozed_until(self):
        created = self.api.create_follow_up(
            {"capturedEmailId": self.email_id, "dueDate": "2026-07-01"}
        )
        follow_up_id = created["item"]["id"]

        snoozed = self.api.update_follow_up(
            {"id": follow_up_id, "action": "snooze", "days": 3}
        )
        expected = (datetime.date.today() + datetime.timedelta(days=3)).isoformat()
        self.assertEqual(expected, snoozed["item"]["snoozedUntil"])

    def test_resolve_and_reopen(self):
        created = self.api.create_follow_up({"capturedEmailId": self.email_id})
        follow_up_id = created["item"]["id"]

        resolved = self.api.update_follow_up({"id": follow_up_id, "action": "resolve"})
        self.assertEqual("resolved", resolved["item"]["status"])
        self.assertTrue(resolved["item"]["resolvedAtUtc"])

        reopened = self.api.update_follow_up({"id": follow_up_id, "action": "reopen"})
        self.assertEqual("open", reopened["item"]["status"])
        self.assertEqual("", reopened["item"]["resolvedAtUtc"])

    def test_list_follow_ups_joins_email_and_sorts_by_effective_due(self):
        self.api.create_follow_up(
            {"capturedEmailId": self.email_id, "title": "No due date"}
        )
        self.api.create_follow_up(
            {"kind": "custom", "title": "Due soon", "dueDate": "2026-07-10"}
        )
        self.api.create_follow_up(
            {"kind": "custom", "title": "Due later", "dueDate": "2026-08-01"}
        )

        result = self.api.list_follow_ups({"status": "open"})
        self.assertEqual("success", result["status"])
        titles = [item["title"] for item in result["items"]]
        self.assertEqual(["Due soon", "Due later", "No due date"], titles)
        email_backed = next(
            item for item in result["items"] if item["title"] == "No due date"
        )
        self.assertEqual("RE: Panel schedule", email_backed["emailSubject"])

    def test_badge_counts_include_overdue(self):
        yesterday = (datetime.date.today() - datetime.timedelta(days=1)).isoformat()
        tomorrow = (datetime.date.today() + datetime.timedelta(days=1)).isoformat()
        self.api.create_follow_up(
            {"kind": "custom", "title": "Overdue one", "dueDate": yesterday}
        )
        self.api.create_follow_up(
            {"kind": "custom", "title": "Future one", "dueDate": tomorrow}
        )

        counts = self.api.get_email_capture_badge_counts()
        self.assertEqual("success", counts["status"])
        self.assertEqual(1, counts["newDirected"])
        self.assertEqual(2, counts["openFollowUps"])
        self.assertEqual(1, counts["overdueFollowUps"])

    def test_snoozed_follow_up_not_overdue(self):
        yesterday = (datetime.date.today() - datetime.timedelta(days=1)).isoformat()
        created = self.api.create_follow_up(
            {"kind": "custom", "title": "Overdue but snoozed", "dueDate": yesterday}
        )
        self.api.update_follow_up(
            {"id": created["item"]["id"], "action": "snooze", "days": 5}
        )
        counts = self.api.get_email_capture_badge_counts()
        self.assertEqual(0, counts["overdueFollowUps"])


class RecipientExtractionTests(unittest.TestCase):
    def test_extract_recipients_splits_to_and_cc(self):
        class _Accessor:
            def __init__(self, smtp):
                self._smtp = smtp

            def GetProperty(self, _tag):
                return self._smtp

        class _Recipient:
            def __init__(self, name, smtp, recipient_type):
                self.Name = name
                self.Type = recipient_type
                self.Address = "/o=Exchange/cn=" + name
                self.PropertyAccessor = _Accessor(smtp)

        class _Recipients:
            def __init__(self, entries):
                self._entries = entries
                self.Count = len(entries)

            def Item(self, index):
                return self._entries[index - 1]

        class _MailItem:
            Recipients = _Recipients(
                [
                    _Recipient("Jacob Husband", "jacob@acies.com", 1),
                    _Recipient("Jane Architect", "jane@arch.com", 2),
                ]
            )

        api = Api.__new__(Api)
        to_recipients, cc_recipients = api._extract_desktop_outlook_recipients(_MailItem())
        self.assertEqual(
            [{"name": "Jacob Husband", "address": "jacob@acies.com"}], to_recipients
        )
        self.assertEqual(
            [{"name": "Jane Architect", "address": "jane@arch.com"}], cc_recipients
        )

    def test_extract_recipients_handles_com_failure(self):
        class _MailItem:
            @property
            def Recipients(self):
                raise RuntimeError("COM failure")

        api = Api.__new__(Api)
        to_recipients, cc_recipients = api._extract_desktop_outlook_recipients(_MailItem())
        self.assertEqual([], to_recipients)
        self.assertEqual([], cc_recipients)


class ScanStateTests(EmailCaptureDbTestCase):
    def test_scan_state_reports_watermark_and_running_flag(self):
        api = Api.__new__(Api)
        api._email_capture_scan_running = False
        api._email_capture_scan_lock = main_module.threading.Lock()

        conn = main_module._open_email_capture_db()
        try:
            main_module._set_email_capture_meta(
                conn, "last_scan_completed_at_utc", "2026-07-09T12:00:00Z"
            )
            main_module._set_email_capture_meta(conn, "last_scan_status", "success")
            conn.commit()
        finally:
            conn.close()

        state = api.get_email_capture_scan_state()
        self.assertEqual("success", state["status"])
        self.assertFalse(state["running"])
        self.assertEqual("2026-07-09T12:00:00Z", state["lastScanCompletedAtUtc"])
        self.assertEqual("success", state["lastScanStatus"])


if __name__ == "__main__":
    unittest.main()
