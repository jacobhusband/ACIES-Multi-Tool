from google.genai import types
from google import genai
import webview
import json
import os
import subprocess
import sys
import shutil
import csv
import io
import logging
import tempfile
import time
import sqlite3
import http.server
import html
from pathlib import Path
from copy import copy, deepcopy
from contextlib import closing
from dotenv import load_dotenv
import datetime
import threading
import secrets
import hashlib
import requests  # Added for GitHub API calls
import zipfile   # Added for extracting bundles
import math
import random
import re
import base64
from typing import List
import importlib
import uuid
from urllib.parse import parse_qs, urlencode, urlparse, unquote
import fitz

# Work around NumPy 2.x removing scalar aliases used by openpyxl.
_NUMPY_ALIAS_MAP = {
    "short": "int16",
    "ushort": "uint16",
    "intc": "int32",
    "uintc": "uint32",
    "int_": "int64",
    "uint": "uint64",
    "longlong": "int64",
    "ulonglong": "uint64",
    "half": "float16",
    "single": "float32",
    "double": "float64",
}
_NUMPY_ATTR_FALLBACKS = {
    "int8": int,
    "int16": int,
    "int32": int,
    "int64": int,
    "uint8": int,
    "uint16": int,
    "uint32": int,
    "uint64": int,
    "intp": int,
    "uintp": int,
    "float16": float,
    "float32": float,
    "float64": float,
    "longdouble": float,
    "bool_": bool,
    "floating": float,
    "integer": int,
}


def _patch_numpy_aliases():
    try:
        import numpy as _np  # noqa: F401
    except Exception:
        return False
    changed = False
    for _name, _fallback in _NUMPY_ATTR_FALLBACKS.items():
        if not hasattr(_np, _name):
            setattr(_np, _name, _fallback)
            changed = True
    for _alias, _target in _NUMPY_ALIAS_MAP.items():
        if not hasattr(_np, _alias):
            if hasattr(_np, _target):
                setattr(_np, _alias, getattr(_np, _target))
            else:
                setattr(_np, _alias, _NUMPY_ATTR_FALLBACKS.get(_target, int))
            changed = True
    return changed


_patch_numpy_aliases()

try:
    import openpyxl
except AttributeError as _exc:
    _msg = str(_exc)
    if "numpy" in _msg:
        _patch_numpy_aliases()
        for _name in ("openpyxl.compat.numbers", "openpyxl.compat", "openpyxl"):
            sys.modules.pop(_name, None)
        importlib.invalidate_caches()
        import openpyxl
    else:
        raise
from openpyxl.worksheet.copier import WorksheetCopy
from pydantic import BaseModel, Field
from PIL import Image as PILImage, ImageOps, ImageDraw, ImageFont, UnidentifiedImageError

# Helper functions for date parsing and status management
STATUS_CANON = ["Waiting", "Working",
                "Pending Review", "Complete", "Delivered"]
STATUS_PRIORITY = ['Delivered', 'Complete',
                   'Pending Review', 'Working', 'Waiting']
LABEL_TO_KEY = {
    "Waiting": "waiting",
    "Working": "working",
    "Pending Review": "pendingReview",
    "Complete": "complete",
    "Delivered": "delivered"
}
KEY_TO_LABEL = {v: k for k, v in LABEL_TO_KEY.items()}

APP_UPDATE_REPO = "jacobhusband/ACIES-Multi-Tool"
APP_INSTALLER_NAME = "acies-scheduler-setup.exe"
GITHUB_API_BASE = "https://api.github.com"
KNOWN_PLUGIN_BUNDLES = [
    "ElectricalCommands.AutoLispCommands.bundle",
    "ElectricalCommands.CleanCADCommands.bundle",
    "ElectricalCommands.GeneralCommands.bundle",
    "ElectricalCommands.GetAttributesCommands.bundle",
    "ElectricalCommands.PlotCommands.bundle",
    "ElectricalCommands.T24Commands.bundle",
    "ElectricalCommands.LFSCommands.bundle",
    "ElectricalCommands.TextCommands.bundle",
]
HIDDEN_PLUGIN_BUNDLES = {
    "ElectricalCommands.GetAttributesCommands.bundle",
}


def load_app_version():
    """Read version from VERSION file; fallback to '0.0.0'."""
    version_file = Path(__file__).resolve().parent / "VERSION"
    try:
        return version_file.read_text(encoding="utf-8").strip() or "0.0.0"
    except Exception as e:
        logging.warning(f"Could not read VERSION file: {e}")
        return "0.0.0"


APP_VERSION = load_app_version()
BUNDLE_RELEASE_TAG = f"v{APP_VERSION}"
GOOGLE_OAUTH_CLIENT_ID_ENV = "GOOGLE_OAUTH_CLIENT_ID"
GOOGLE_OAUTH_CLIENT_SECRET_ENV = "GOOGLE_CLIENT_SECRET"
GOOGLE_OAUTH_AUTH_ENDPOINT = "https://accounts.google.com/o/oauth2/v2/auth"
GOOGLE_OAUTH_TOKEN_ENDPOINT = "https://oauth2.googleapis.com/token"
GOOGLE_OAUTH_REVOKE_ENDPOINT = "https://oauth2.googleapis.com/revoke"
GOOGLE_OAUTH_USERINFO_ENDPOINT = "https://openidconnect.googleapis.com/v1/userinfo"
GOOGLE_OAUTH_SCOPES = ["openid", "email", "profile"]
GOOGLE_OAUTH_CALLBACK_PATH = "/"
GOOGLE_OAUTH_TIMEOUT_SECONDS = 180
OUTLOOK_SCAN_SOURCE_DESKTOP = "desktop-outlook"
OUTLOOK_MAPI_INBOX_FOLDER_ID = 6
OUTLOOK_MAPI_PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
OUTLOOK_SCAN_DEDUPE_SKIP_REASON_THREAD = (
    "Removed from the AI batch because this reply only repeated thread history already seen in newer emails."
)
OUTLOOK_SCAN_DEDUPE_SKIP_REASON_DUPLICATE = (
    "Removed from the AI batch because its content duplicated another email in the selected timeframe."
)
OUTLOOK_SCAN_FETCH_LIMIT = 200
OUTLOOK_SCAN_AI_LIMIT = 40
OUTLOOK_SCAN_PROMPT_MAX_CHARS = 120000
OUTLOOK_SCAN_PROMPT_EMAIL_BUDGET_CHARS = 90000
OUTLOOK_SCAN_PROMPT_DELIVERABLE_BUDGET_CHARS = 25000
OUTLOOK_SCAN_EMAIL_BODY_PROMPT_CHARS = 1200
OUTLOOK_SCAN_RETRY_AI_LIMIT = 20
OUTLOOK_SCAN_RETRY_PROMPT_EMAIL_BUDGET_CHARS = 45000
OUTLOOK_SCAN_RETRY_EMAIL_BODY_PROMPT_CHARS = 600
OUTLOOK_SCAN_PROMPT_SKIP_REASON = "Excluded from the AI batch prompt to keep the scan bounded."
OUTLOOK_SCAN_RETRY_SKIP_REASON = (
    "Excluded from the reduced AI retry batch after the first request timed out."
)
FIREBASE_API_KEY_ENV = "FIREBASE_API_KEY"
FIREBASE_AUTH_DOMAIN_ENV = "FIREBASE_AUTH_DOMAIN"
FIREBASE_PROJECT_ID_ENV = "FIREBASE_PROJECT_ID"
FIREBASE_APP_ID_ENV = "FIREBASE_APP_ID"
FIREBASE_STORAGE_BUCKET_ENV = "FIREBASE_STORAGE_BUCKET"
FIREBASE_MESSAGING_SENDER_ID_ENV = "FIREBASE_MESSAGING_SENDER_ID"


def build_default_cloud_sync_settings():
    return {
        'enabled': False,
        'firebaseUid': '',
        'lastSyncedAt': '',
        'migrationCompleted': False,
    }


def build_default_user_settings():
    return {
        'userName': '',
        'discipline': ['Electrical'],
        'apiKey': '',
        'autocadPath': '',
        'showSetupHelp': True,
        'cleanDwgOptions': {
            'stripXrefs': True,
            'setByLayer': True,
            'purge': True,
            'audit': True,
            'hatchColor': True
        },
        'publishDwgOptions': {
            'autoDetectPaperSize': True,
            'shrinkPercent': 100
        },
        'freezeLayerOptions': {
            'scanAllLayers': True
        },
        'thawLayerOptions': {
            'scanAllLayers': True
        },
        'workroomAutoSelectCadFiles': True,
        'enableUnderConstructionTools': False,
        'googleAuth': None,
        'cloudSync': build_default_cloud_sync_settings(),
    }


def _get_legacy_windows_app_data_dir():
    user_profile = os.getenv('USERPROFILE') or os.path.expanduser('~')
    appdata_base = os.getenv('APPDATA') or os.path.join(
        user_profile, 'AppData', 'Roaming')
    return os.path.join(appdata_base, 'ProjectManagementApp')


def _load_legacy_user_settings():
    if sys.platform != "win32":
        return None
    legacy_path = os.path.join(_get_legacy_windows_app_data_dir(), "settings.json")
    try:
        with open(legacy_path, 'r', encoding='utf-8') as f:
            payload = json.load(f)
        return payload if isinstance(payload, dict) else None
    except (FileNotFoundError, json.JSONDecodeError):
        return None


def _merge_missing_user_settings_from_legacy(current_settings, legacy_settings):
    current = current_settings if isinstance(current_settings, dict) else {}
    legacy = legacy_settings if isinstance(legacy_settings, dict) else {}
    if not legacy:
        return current, False

    merged = deepcopy(current)
    changed = False

    for key in (
        "userName",
        "apiKey",
        "autocadPath",
        "theme",
        "defaultPmInitials",
        "autocadVersion",
    ):
        current_value = str(merged.get(key) or "").strip()
        legacy_value = str(legacy.get(key) or "").strip()
        if not current_value and legacy_value:
            merged[key] = legacy_value
            changed = True

    current_discipline = merged.get("discipline")
    legacy_discipline = legacy.get("discipline")
    if (
        (not isinstance(current_discipline, list) or not current_discipline)
        and isinstance(legacy_discipline, list)
        and legacy_discipline
    ):
        merged["discipline"] = deepcopy(legacy_discipline)
        changed = True

    if (
        merged.get("showSetupHelp", True) is not False
        and legacy.get("showSetupHelp") is False
    ):
        merged["showSetupHelp"] = False
        changed = True

    if merged.get("autoPrimary") is not True and legacy.get("autoPrimary") is True:
        merged["autoPrimary"] = True
        changed = True

    current_templates = merged.get("lightingTemplates")
    legacy_templates = legacy.get("lightingTemplates")
    if (
        (not isinstance(current_templates, list) or not current_templates)
        and isinstance(legacy_templates, list)
        and legacy_templates
    ):
        merged["lightingTemplates"] = deepcopy(legacy_templates)
        changed = True

    for auth_key in ("googleAuth",):
        if not isinstance(merged.get(auth_key), dict) and isinstance(legacy.get(auth_key), dict):
            merged[auth_key] = deepcopy(legacy.get(auth_key))
            changed = True

    return merged, changed


def _sanitize_user_settings_payload(settings):
    normalized = deepcopy(settings) if isinstance(settings, dict) else {}
    changed = False

    if "microsoftAuth" in normalized:
        normalized.pop("microsoftAuth", None)
        changed = True

    defaults = build_default_user_settings()
    for key, value in defaults.items():
        if key not in normalized:
            normalized[key] = deepcopy(value)
            changed = True

    return normalized, changed


def utc_now_iso():
    return datetime.datetime.now(datetime.timezone.utc).replace(
        microsecond=0
    ).isoformat().replace("+00:00", "Z")


def parse_utc_iso(value):
    raw = str(value or "").strip()
    if not raw:
        return None
    try:
        return datetime.datetime.fromisoformat(raw.replace("Z", "+00:00"))
    except Exception:
        return None


def _normalize_version(raw):
    """Turn a version string like 'v1.2.3' into '1.2.3'."""
    return str(raw or '').strip().lstrip('vV')


def _version_tuple(raw):
    parts = []
    for chunk in _normalize_version(raw).split('.'):
        try:
            parts.append(int(chunk))
        except ValueError:
            break
    return tuple(parts or [0])


def _is_remote_newer(remote, current):
    return _version_tuple(remote) > _version_tuple(current)


def _bundle_name_from_asset_name(asset_name, release_tag):
    bundle_name = str(asset_name or "")
    if release_tag:
        bundle_name = bundle_name.replace(f"-{release_tag}.zip", ".bundle")
    if bundle_name.endswith(".zip"):
        bundle_name = bundle_name[:-4] + ".bundle"
    return bundle_name


def parse_due_str(s):
    """Parse due date string similar to JS parseDueStr."""
    if not s:
        return None
    s = s.strip()
    try:
        # Try ISO format first
        return datetime.datetime.fromisoformat(s.replace('Z', '+00:00'))
    except:
        pass
    s = s.replace('.', '/').replace(' ', '')
    parts = s.split('/')
    if len(parts) == 3:
        try:
            mm, dd, yy = parts
            if len(yy) == 2:
                yy = '20' + yy
            iso = f"{yy}-{mm.zfill(2)}-{dd.zfill(2)}T12:00:00"
            return datetime.datetime.fromisoformat(iso)
        except:
            pass
    try:
        return datetime.datetime.strptime(s, "%m/%d/%Y")
    except:
        pass
    return None


def sync_status_arrays(task):
    """Sync status arrays similar to JS syncStatusArrays."""
    if not isinstance(task.get('statuses'), list):
        task['statuses'] = []
    from_tags = task.get('statusTags', [])
    for key in from_tags:
        label = KEY_TO_LABEL.get(key)
        if label and label not in task['statuses']:
            task['statuses'].append(label)
    task['statuses'] = list({s for s in task['statuses'] if s in STATUS_CANON})
    task['statusTags'] = [LABEL_TO_KEY[s]
                          for s in task['statuses'] if LABEL_TO_KEY.get(s)]
    task['status'] = next(
        (label for label in STATUS_PRIORITY if label in task['statuses']), '')


def get_default_documents_dir():
    user_profile = os.getenv('USERPROFILE') or os.path.expanduser('~')
    onedrive_docs = os.path.join(
        user_profile, 'OneDrive - ACIES Engineering', 'Documents')
    if os.path.isdir(onedrive_docs):
        return onedrive_docs
    return os.path.join(user_profile, 'Documents')


_LEGACY_MIGRATION_DONE = False


def _get_windows_documents_dir():
    user_profile = os.getenv('USERPROFILE')
    if user_profile:
        return os.path.join(user_profile, 'Documents')
    return os.path.expanduser('~')


def _migrate_legacy_app_data():
    global _LEGACY_MIGRATION_DONE
    if _LEGACY_MIGRATION_DONE or sys.platform != "win32":
        return
    _LEGACY_MIGRATION_DONE = True

    user_profile = os.getenv('USERPROFILE') or os.path.expanduser('~')
    appdata_base = os.getenv('APPDATA') or os.path.join(
        user_profile, 'AppData', 'Roaming')
    old_dir = os.path.join(appdata_base, 'ProjectManagementApp')
    if not os.path.isdir(old_dir):
        return

    new_dir = os.path.join(_get_windows_documents_dir(), 'ProjectManagementApp')
    os.makedirs(new_dir, exist_ok=True)

    try:
        items = os.listdir(old_dir)
    except OSError as exc:
        logging.warning(f"Could not list legacy data directory: {exc}")
        return

    for name in items:
        src = os.path.join(old_dir, name)
        dst = os.path.join(new_dir, name)
        if os.path.exists(dst):
            logging.info(
                f"Skipping legacy data item because target exists: {dst}")
            continue
        try:
            if os.path.isdir(src):
                shutil.copytree(src, dst)
            else:
                shutil.copy2(src, dst)
            logging.info(f"Migrated legacy data item: {src} -> {dst}")
        except Exception as exc:
            logging.warning(f"Failed to migrate legacy data item {src}: {exc}")


def build_timesheet_filename(user_name, week_key):
    tokens = [t for t in str(user_name or '').strip().split() if t]
    if len(tokens) >= 2:
        initials = (tokens[0][0] + tokens[-1][0]).upper()
    elif tokens:
        initials = (tokens[0][:2]).upper()
        if len(initials) == 1:
            initials = initials + initials
    else:
        initials = "TS"

    week_start = None
    if week_key:
        try:
            week_start = datetime.date.fromisoformat(week_key)
        except ValueError:
            week_start = None
    if not week_start:
        week_start = datetime.date.today()

    monday = week_start + datetime.timedelta(
        days=(7 + 0 - week_start.weekday()) % 7
    )
    friday = monday + datetime.timedelta(days=4)

    return f"{initials}-TS-{monday.strftime('%m%d')}-{friday.strftime('%m%d%Y')}.xlsx"



TIMESHEET_DAY_ORDER = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]


def _coerce_float(value, default=0.0):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _get_timesheet_project_key(project_id="", project_name=""):
    normalized_id = str(project_id or "").strip()
    if normalized_id:
        return f"id:{normalized_id.casefold()}"
    normalized_name = str(project_name or "").strip()
    if normalized_name:
        return f"name:{normalized_name.casefold()}"
    return ""


def _get_timesheet_day_from_expense_date(value):
    if not value:
        return None
    try:
        return TIMESHEET_DAY_ORDER[datetime.date.fromisoformat(str(value)).weekday()]
    except ValueError:
        return None


def _build_export_mileage_by_row(entries, expense_projects):
    if not entries:
        return []

    mileage_by_row = [0.0] * len(entries)
    if not expense_projects:
        return mileage_by_row

    day_rows = {day: [] for day in TIMESHEET_DAY_ORDER}
    project_rows = {}
    project_day_rows = {}
    project_day_totals = {}

    for row_index, entry in enumerate(entries):
        project_key = _get_timesheet_project_key(
            entry.get("projectId", ""),
            entry.get("projectName", ""),
        )
        if project_key:
            project_rows.setdefault(project_key, []).append(row_index)

        hours = entry.get("hours") or {}
        for day in TIMESHEET_DAY_ORDER:
            if _coerce_float(hours.get(day, 0), 0.0) <= 0:
                continue
            day_rows[day].append(row_index)
            if project_key:
                project_day_rows.setdefault((project_key, day), []).append(row_index)

    for project in expense_projects:
        project_key = _get_timesheet_project_key(
            project.get("projectId", ""),
            project.get("projectName", ""),
        )
        for expense_entry in project.get("entries", []):
            day = _get_timesheet_day_from_expense_date(expense_entry.get("date"))
            mileage = _coerce_float(expense_entry.get("mileage", 0), 0.0)
            if not day or mileage <= 0:
                continue
            bucket = (project_key, day)
            project_day_totals[bucket] = project_day_totals.get(bucket, 0.0) + mileage

    for (project_key, day), mileage in project_day_totals.items():
        target_rows = []
        if project_key:
            target_rows = project_day_rows.get((project_key, day), [])
        if not target_rows:
            target_rows = day_rows.get(day, [])
        if not target_rows and project_key:
            target_rows = project_rows.get(project_key, [])
        if not target_rows:
            target_rows = [0]
        mileage_by_row[target_rows[0]] += mileage

    return mileage_by_row


def _format_long_date(date_value=None):
    """Format date like 'January 1, 2026'."""
    if date_value is None:
        date_value = datetime.date.today()
    if isinstance(date_value, datetime.datetime):
        date_value = date_value.date()
    return f"{date_value.strftime('%B')} {date_value.day}, {date_value.strftime('%Y')}"


def _replace_text_in_paragraph(paragraph, replacements):
    if not paragraph.text:
        return
    original = paragraph.text
    updated = original
    for find_text, replace_text in replacements.items():
        if find_text:
            updated = updated.replace(find_text, replace_text)
    if updated == original:
        return

    for run in paragraph.runs:
        for find_text, replace_text in replacements.items():
            if find_text and find_text in run.text:
                run.text = run.text.replace(find_text, replace_text)

    if paragraph.text != updated:
        paragraph.text = updated


def _delete_docx_paragraph(paragraph):
    element = paragraph._element
    element.getparent().remove(element)


def _remove_to_section_docx(doc):
    to_remove = []
    removing = False
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        lower = text.lower()
        if not removing and lower.startswith("to"):
            removing = True
            to_remove.append(paragraph)
            continue
        if removing:
            if "project" in lower or "*project name*" in lower:
                removing = False
                continue
            to_remove.append(paragraph)
    for paragraph in to_remove:
        _delete_docx_paragraph(paragraph)


def _replace_text_in_docx(doc, replacements):
    for paragraph in doc.paragraphs:
        _replace_text_in_paragraph(paragraph, replacements)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    _replace_text_in_paragraph(paragraph, replacements)
    for section in doc.sections:
        for paragraph in section.header.paragraphs:
            _replace_text_in_paragraph(paragraph, replacements)
        for paragraph in section.footer.paragraphs:
            _replace_text_in_paragraph(paragraph, replacements)


def _update_docx_file(file_path, replacements, remove_to_section=False):
    from docx import Document

    doc = Document(file_path)
    if remove_to_section:
        _remove_to_section_docx(doc)
    if replacements:
        _replace_text_in_docx(doc, replacements)
    doc.save(file_path)


def _remove_to_section_word(doc):
    removing = False
    idx = 1
    while idx <= doc.Paragraphs.Count:
        paragraph = doc.Paragraphs(idx)
        text = paragraph.Range.Text.replace("\r", "").strip()
        lower = text.lower()
        if not removing and lower.startswith("to"):
            removing = True
        if removing:
            if "project" in lower or "*project name*" in lower:
                removing = False
                idx += 1
                continue
            paragraph.Range.Delete()
            continue
        idx += 1


def _update_word_file_via_com(file_path, replacements, remove_to_section=False):
    import win32com.client

    # Launch an isolated Word instance so closing the template document does not
    # shut down any Word windows the user already has open.
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = None
    try:
        doc = word.Documents.Open(file_path)
        for find_text, replace_text in replacements.items():
            if not find_text:
                continue
            rng = doc.Range()
            rng.Find.ClearFormatting()
            rng.Find.Replacement.ClearFormatting()
            rng.Find.Execute(
                FindText=find_text,
                ReplaceWith=replace_text,
                Replace=2,  # wdReplaceAll
            )
        if remove_to_section:
            _remove_to_section_word(doc)
        doc.Save()
    finally:
        if doc is not None:
            doc.Close(False)
        word.Quit()


def _apply_template_context(file_path, context=None, options=None):
    context = context or {}
    options = options or {}
    template_key = options.get("templateKey")
    remove_to_section = bool(options.get("removeToSection"))

    replacements = {
        "*current date*": _format_long_date(),
    }
    deliverable_name = str(context.get("deliverableName") or "").strip()
    project_name = str(context.get("projectName") or "").strip()
    if deliverable_name:
        replacements["*deliverable*"] = deliverable_name
    if project_name:
        replacements["*project name*"] = project_name

    if not replacements and not remove_to_section:
        return

    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".docx":
        _update_docx_file(file_path, replacements, remove_to_section=remove_to_section)
    elif ext == ".doc":
        _update_word_file_via_com(file_path, replacements, remove_to_section=remove_to_section)


# --- Google GenAI (new client) ---
# Uses GOOGLE_API_KEY from environment/.env

# --- Helper function to get the application data path ---


def get_app_data_dir():
    """
    Determines the correct, cross-platform directory for storing user data.

    On Windows, we store data in the user's Documents folder to bypass
    Windows Store Python's file system virtualization, which redirects
    writes to AppData to a package-specific virtualized location.
    """
    if sys.platform == "win32":
        _migrate_legacy_app_data()
        base_dir = _get_windows_documents_dir()
    elif sys.platform == "darwin":
        base_dir = os.path.join(os.path.expanduser(
            '~'), 'Library', 'Application Support')
    else:
        base_dir = os.path.join(os.path.expanduser('~'), '.local', 'share')
    if not base_dir:
        base_dir = os.path.expanduser('~')
    app_data_dir = os.path.join(base_dir, 'ProjectManagementApp')
    os.makedirs(app_data_dir, exist_ok=True)
    return app_data_dir


def get_app_data_path(file_name="tasks.json"):
    """
    Determines the correct, cross-platform path for storing user data
    and returns the full path for the given file_name.
    """
    return os.path.join(get_app_data_dir(), file_name)


# --- Configuration ---
BASE_DIR = Path(__file__).resolve().parent
loaded = load_dotenv(BASE_DIR / '.env', override=True)  # <— force .env to win
logging.info(f".env loaded: {loaded} (override=True)")
TASKS_FILE = get_app_data_path("tasks.json")
NOTES_FILE = get_app_data_path("notes.json")
SETTINGS_FILE = get_app_data_path("settings.json")
TIMESHEETS_FILE = get_app_data_path("timesheets.json")
TEMPLATES_FILE = get_app_data_path("templates.json")
CAD_AUTO_SELECT_TRACE_FILE = get_app_data_path("cad_auto_select_trace.log")
CHECKLISTS_FILE = get_app_data_path("checklists.json")
SYNC_BACKUPS_DIR = os.path.join(get_app_data_dir(), "sync_backups")
LIGHTING_SCHEDULE_SYNC_FILE = "T24LightingFixtureSchedule.sync.json"
LIGHTING_SCHEDULE_DB_FILE = get_app_data_path("lighting_schedules.db")
SYNC_TRACKED_FILES = {
    "settings": SETTINGS_FILE,
    "tasks": TASKS_FILE,
    "notes": NOTES_FILE,
    "timesheets": TIMESHEETS_FILE,
    "templates": TEMPLATES_FILE,
    "checklists": CHECKLISTS_FILE,
}
LIGHTING_SCHEDULE_FIELDS = [
    "mark",
    "description",
    "manufacturer",
    "modelNumber",
    "mounting",
    "volts",
    "watts",
    "notes",
]
LIGHTING_SCHEDULE_DEFAULT_GENERAL_NOTES = "\n".join([
    "A.  VERIFY ALL CEILING TYPES PRIOR TO ORDERING FIXTURES.",
    "B.  VERIFY ALL OPERATING VOLTAGE PRIOR TO ORDERING FIXTURES.",
    "C.  COORDINATE THE HEIGHT OF ALL SUSPENDED FIXTURES WITH THE OWNER, ARCHITECT AND ENGINEER PRIOR TO INSTALLATION.",
    "D.  ALL LAMPS SHALL HAVE A COLOR TEMPERATURE OF 3500 DEG. KELVIN AND A CRI OF 85 UNLESS SPECIFICALLY NOTED.",
    'E.  FIXTURES DESIGNATED WITH A "1/2 SHADE" OR "FULL SHADE" AND ALL EXIT SIGNS SHALL BE CIRCUITED TO THE CENTRAL EMERGENCY BATTERY INVERTER OR PROVIDED WITH INTEGRAL BATTERY BACKUP AS REQUIRED. EM BACKUP POWER SHALL BE SUITABLE TO PROVIDE FULL POWER TO FIXTURES FOR A MINIMUM OF 90 MINUTES.',
])
LIGHTING_SCHEDULE_DEFAULT_NOTES = "\n".join([
    "1.  CONFIRM THE DRIVER TYPE WITH VENDOR. LIGHT DRIVER SHALL BE COMPATIBLE WITH THE DIMMER. REFER TO THE SENSOR SCHEDULE FOR DETAILS.",
    "2.  COORDINATE THE FINAL MANUFACTURER AND MODEL WITH ARCHITECT.",
])
LIGHTING_PROJECT_SEGMENT_REGEX = re.compile(
    r"^(\d{6})(?!\d)(?:\s*(?:[-_]\s*)?(.*))?$"
)


def _normalize_sync_dwg_path(dwg_path):
    raw = str(dwg_path or "").strip().strip('"').strip("'")
    if not raw:
        raise ValueError("DWG path is required.")
    normalized = os.path.normpath(raw)
    if not os.path.isabs(normalized):
        raise ValueError("DWG path must be an absolute path.")
    if os.path.splitext(normalized)[1].lower() != ".dwg":
        raise ValueError("DWG path must end with .dwg.")
    folder = os.path.dirname(normalized)
    if not folder:
        raise ValueError("DWG path must include a parent folder.")
    return normalized


def _resolve_lighting_schedule_sync_path(dwg_path):
    normalized_dwg = _normalize_sync_dwg_path(dwg_path)
    folder = os.path.dirname(normalized_dwg)
    sync_path = os.path.join(folder, LIGHTING_SCHEDULE_SYNC_FILE)
    return normalized_dwg, sync_path


def _normalize_t24_output_json_path(json_path):
    raw = str(json_path or "").strip().strip('"').strip("'")
    if not raw:
        raise ValueError("JSON path is required.")
    normalized = os.path.normpath(raw)
    if not os.path.isabs(normalized):
        raise ValueError("JSON path must be an absolute path.")
    if os.path.splitext(normalized)[1].lower() != ".json":
        raise ValueError("JSON path must end with .json.")
    if os.path.basename(normalized).lower() != "t24output.json":
        raise ValueError("Expected file name: T24Output.json.")
    return normalized


def _read_json_file_strict(path):
    try:
        with open(path, "r", encoding="utf-8-sig") as f:
            raw_text = f.read()
    except UnicodeDecodeError as exc:
        raise ValueError(f"Sync file '{path}' is not valid UTF-8 text: {exc}.")

    if raw_text.startswith("\ufeff"):
        raw_text = raw_text.lstrip("\ufeff")

    try:
        return json.loads(raw_text)
    except json.JSONDecodeError as exc:
        raise ValueError(
            f"Malformed JSON in sync file '{path}': {exc.msg} "
            f"(line {exc.lineno}, column {exc.colno})."
        )


def _atomic_write_json_file(path, payload):
    folder = os.path.dirname(path)
    os.makedirs(folder, exist_ok=True)
    fd, temp_path = tempfile.mkstemp(
        prefix="lfs_sync_",
        suffix=".tmp",
        dir=folder,
        text=True,
    )
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        os.replace(temp_path, path)
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass


def _get_file_modified_iso(path):
    if not path or not os.path.exists(path):
        return ""
    try:
        modified = datetime.datetime.fromtimestamp(
            os.path.getmtime(path),
            tz=datetime.timezone.utc,
        )
        return modified.replace(microsecond=0).isoformat().replace("+00:00", "Z")
    except OSError:
        return ""


def _get_local_sync_metadata():
    files = {}
    for key, path in SYNC_TRACKED_FILES.items():
        files[key] = {
            "path": os.path.normpath(path),
            "exists": os.path.exists(path),
            "modified": _get_file_modified_iso(path),
        }
    return files


def _create_cloud_sync_backup(reason="", metadata=None):
    created_at = utc_now_iso()
    safe_reason = re.sub(r"[^a-z0-9]+", "-", str(reason or "").strip().lower()).strip("-")
    suffix = safe_reason[:40] or "sync"
    backup_dir = os.path.join(
        SYNC_BACKUPS_DIR,
        f"{created_at.replace(':', '').replace('-', '')}_{suffix}_{uuid.uuid4().hex[:8]}",
    )
    os.makedirs(backup_dir, exist_ok=True)

    copied = []
    for key, source_path in SYNC_TRACKED_FILES.items():
        if not os.path.exists(source_path):
            continue
        destination_path = os.path.join(backup_dir, os.path.basename(source_path))
        shutil.copy2(source_path, destination_path)
        copied.append(
            {
                "key": key,
                "sourcePath": os.path.normpath(source_path),
                "backupPath": os.path.normpath(destination_path),
                "modified": _get_file_modified_iso(source_path),
            }
        )

    backup_metadata = {
        "createdAt": created_at,
        "reason": str(reason or "").strip(),
        "files": copied,
        "localSyncMetadata": _get_local_sync_metadata(),
    }
    if isinstance(metadata, dict) and metadata:
        backup_metadata["metadata"] = metadata

    _atomic_write_json_file(os.path.join(backup_dir, "metadata.json"), backup_metadata)
    return {
        "path": os.path.normpath(backup_dir),
        "createdAt": created_at,
        "files": copied,
    }


def _normalize_lighting_schedule_text(value):
    return str(value or "").replace("\r\n", "\n").replace("\r", "\n")


def _create_default_lighting_schedule_row(seed=None):
    seed = seed or {}
    row = {}
    for field in LIGHTING_SCHEDULE_FIELDS:
        row[field] = _normalize_lighting_schedule_text(seed.get(field))
    return row


def _create_default_lighting_schedule():
    return {
        "rows": [_create_default_lighting_schedule_row()],
        "generalNotes": LIGHTING_SCHEDULE_DEFAULT_GENERAL_NOTES,
        "notes": LIGHTING_SCHEDULE_DEFAULT_NOTES,
        "targetDwgPath": "",
    }


def _normalize_lighting_schedule_payload(schedule):
    normalized = (
        dict(schedule)
        if isinstance(schedule, dict)
        else _create_default_lighting_schedule()
    )
    rows = normalized.get("rows")
    if not isinstance(rows, list) or not rows:
        normalized["rows"] = [_create_default_lighting_schedule_row()]
    else:
        normalized["rows"] = [
            _create_default_lighting_schedule_row(row)
            for row in rows
            if isinstance(row, dict)
        ] or [_create_default_lighting_schedule_row()]

    normalized["generalNotes"] = _normalize_lighting_schedule_text(
        normalized.get("generalNotes")
        if normalized.get("generalNotes") is not None
        else LIGHTING_SCHEDULE_DEFAULT_GENERAL_NOTES
    )
    normalized["notes"] = _normalize_lighting_schedule_text(
        normalized.get("notes")
        if normalized.get("notes") is not None
        else LIGHTING_SCHEDULE_DEFAULT_NOTES
    )
    normalized["targetDwgPath"] = os.path.normpath(
        str(normalized.get("targetDwgPath") or "").strip()
    ) if str(normalized.get("targetDwgPath") or "").strip() else ""
    return normalized


def _extract_lighting_schedule_project_id_from_path(raw_path):
    normalized = str(raw_path or "").strip().replace("/", "\\")
    if not normalized:
        return ""

    parts = [part.strip() for part in normalized.split("\\") if part.strip()]
    for part in reversed(parts):
        match = LIGHTING_PROJECT_SEGMENT_REGEX.match(part)
        if match:
            return match.group(1)
    return ""


def _resolve_lighting_schedule_project_id(project_like):
    if isinstance(project_like, str):
        value = str(project_like).strip()
        return value

    if not isinstance(project_like, dict):
        return ""

    explicit_id = str(project_like.get("id") or "").strip()
    if explicit_id:
        return explicit_id

    for key in ("path", "localProjectPath", "workroomRootPath"):
        candidate = _extract_lighting_schedule_project_id_from_path(
            project_like.get(key)
        )
        if candidate:
            return candidate

    display_name = str(project_like.get("name") or project_like.get("nick") or "").strip()
    if display_name:
        return f"name:{display_name.lower()}"

    return ""


def _open_lighting_schedule_db():
    conn = sqlite3.connect(LIGHTING_SCHEDULE_DB_FILE, timeout=5)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode = WAL;")
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS lighting_schedule_projects (
            project_id TEXT PRIMARY KEY,
            schedule_json TEXT NOT NULL,
            target_dwg_path TEXT NOT NULL DEFAULT '',
            table_handle TEXT NOT NULL DEFAULT '',
            version INTEGER NOT NULL DEFAULT 1,
            updated_at_utc TEXT NOT NULL,
            updated_by TEXT NOT NULL DEFAULT 'unknown'
        );
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS lighting_schedule_links (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id TEXT NOT NULL,
            dwg_path TEXT NOT NULL UNIQUE,
            table_handle TEXT NOT NULL DEFAULT '',
            last_applied_version INTEGER NOT NULL DEFAULT 0,
            last_seen_at_utc TEXT NOT NULL DEFAULT '',
            FOREIGN KEY(project_id) REFERENCES lighting_schedule_projects(project_id)
                ON DELETE CASCADE
        );
        """
    )
    conn.execute(
        """
        CREATE INDEX IF NOT EXISTS idx_lighting_schedule_links_project_id
        ON lighting_schedule_links(project_id);
        """
    )
    return conn


def _normalize_lighting_schedule_record(row):
    if row is None:
        return None

    try:
        schedule_payload = json.loads(row["schedule_json"])
    except (TypeError, ValueError, KeyError):
        schedule_payload = _create_default_lighting_schedule()

    schedule = _normalize_lighting_schedule_payload(schedule_payload)
    schedule["targetDwgPath"] = os.path.normpath(
        str(row["target_dwg_path"] or "").strip()
    ) if str(row["target_dwg_path"] or "").strip() else ""

    return {
        "projectId": str(row["project_id"] or "").strip(),
        "version": int(row["version"] or 0),
        "updatedAtUtc": str(row["updated_at_utc"] or "").strip(),
        "updatedBy": str(row["updated_by"] or "").strip(),
        "targetDwgPath": schedule["targetDwgPath"],
        "tableHandle": str(row["table_handle"] or "").strip(),
        "schedule": {
            "rows": schedule["rows"],
            "generalNotes": schedule["generalNotes"],
            "notes": schedule["notes"],
        },
    }


def _get_lighting_schedule_record(project_id):
    resolved_id = _resolve_lighting_schedule_project_id(project_id)
    if not resolved_id:
        return None

    with _open_lighting_schedule_db() as conn:
        row = conn.execute(
            """
            SELECT project_id, schedule_json, target_dwg_path, table_handle, version,
                   updated_at_utc, updated_by
            FROM lighting_schedule_projects
            WHERE project_id = ?
            """,
            (resolved_id,),
        ).fetchone()
        return _normalize_lighting_schedule_record(row)


def _get_lighting_schedule_version(project_id):
    record = _get_lighting_schedule_record(project_id)
    if not record:
        return {
            "exists": False,
            "projectId": _resolve_lighting_schedule_project_id(project_id),
            "version": 0,
            "updatedAtUtc": "",
            "updatedBy": "",
        }

    return {
        "exists": True,
        "projectId": record["projectId"],
        "version": record["version"],
        "updatedAtUtc": record["updatedAtUtc"],
        "updatedBy": record["updatedBy"],
    }


def _get_lighting_schedule_links(project_id):
    resolved_id = _resolve_lighting_schedule_project_id(project_id)
    if not resolved_id:
        return []

    with _open_lighting_schedule_db() as conn:
        rows = conn.execute(
            """
            SELECT id, project_id, dwg_path, table_handle, last_applied_version, last_seen_at_utc
            FROM lighting_schedule_links
            WHERE project_id = ?
            ORDER BY dwg_path COLLATE NOCASE
            """,
            (resolved_id,),
        ).fetchall()

    links = []
    for row in rows:
        links.append(
            {
                "id": int(row["id"]),
                "projectId": str(row["project_id"] or "").strip(),
                "dwgPath": str(row["dwg_path"] or "").strip(),
                "tableHandle": str(row["table_handle"] or "").strip(),
                "lastAppliedVersion": int(row["last_applied_version"] or 0),
                "lastSeenAtUtc": str(row["last_seen_at_utc"] or "").strip(),
            }
        )
    return links


def _upsert_lighting_schedule_link(
    conn,
    project_id,
    dwg_path,
    table_handle="",
    last_applied_version=0,
):
    normalized_dwg = os.path.normpath(str(dwg_path or "").strip())
    if not normalized_dwg:
        return

    now_iso = datetime.datetime.utcnow().isoformat()
    existing = conn.execute(
        """
        SELECT table_handle, last_applied_version
        FROM lighting_schedule_links
        WHERE dwg_path = ?
        """,
        (normalized_dwg,),
    ).fetchone()

    next_table_handle = str(table_handle or "").strip()
    if not next_table_handle and existing:
        next_table_handle = str(existing["table_handle"] or "").strip()

    next_applied_version = int(last_applied_version or 0)
    if existing and next_applied_version <= 0:
        next_applied_version = int(existing["last_applied_version"] or 0)

    conn.execute(
        """
        INSERT INTO lighting_schedule_links (
            project_id, dwg_path, table_handle, last_applied_version, last_seen_at_utc
        )
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(dwg_path) DO UPDATE SET
            project_id = excluded.project_id,
            table_handle = excluded.table_handle,
            last_applied_version = excluded.last_applied_version,
            last_seen_at_utc = excluded.last_seen_at_utc
        """,
        (
            project_id,
            normalized_dwg,
            next_table_handle,
            next_applied_version,
            now_iso,
        ),
    )


def _save_lighting_schedule_record(
    project_id,
    payload,
    updated_by="desktop",
):
    resolved_id = _resolve_lighting_schedule_project_id(project_id)
    if not resolved_id:
        raise ValueError("Lighting schedule project ID is required.")
    if not isinstance(payload, dict):
        raise ValueError("Lighting schedule payload must be a JSON object.")

    normalized_schedule = _normalize_lighting_schedule_payload(payload.get("schedule"))
    target_dwg_path = payload.get("targetDwgPath")
    if target_dwg_path is None:
        target_dwg_path = normalized_schedule.get("targetDwgPath")
    normalized_target_dwg_path = (
        os.path.normpath(str(target_dwg_path or "").strip())
        if str(target_dwg_path or "").strip()
        else ""
    )
    table_handle = str(payload.get("tableHandle") or "").strip()
    expected_version = payload.get("expectedVersion", None)

    now_iso = datetime.datetime.utcnow().isoformat()
    canonical_schedule_json = json.dumps(
        {
            "rows": normalized_schedule["rows"],
            "generalNotes": normalized_schedule["generalNotes"],
            "notes": normalized_schedule["notes"],
        },
        ensure_ascii=False,
        separators=(",", ":"),
    )

    with _open_lighting_schedule_db() as conn:
        existing = conn.execute(
            """
            SELECT project_id, schedule_json, target_dwg_path, table_handle, version
            FROM lighting_schedule_projects
            WHERE project_id = ?
            """,
            (resolved_id,),
        ).fetchone()

        previous_version = int(existing["version"] or 0) if existing else 0
        next_version = previous_version + 1 if existing else 1
        current_table_handle = table_handle
        if existing and not current_table_handle:
            current_table_handle = str(existing["table_handle"] or "").strip()

        conn.execute(
            """
            INSERT INTO lighting_schedule_projects (
                project_id, schedule_json, target_dwg_path, table_handle,
                version, updated_at_utc, updated_by
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(project_id) DO UPDATE SET
                schedule_json = excluded.schedule_json,
                target_dwg_path = excluded.target_dwg_path,
                table_handle = excluded.table_handle,
                version = excluded.version,
                updated_at_utc = excluded.updated_at_utc,
                updated_by = excluded.updated_by
            """,
            (
                resolved_id,
                canonical_schedule_json,
                normalized_target_dwg_path,
                current_table_handle,
                next_version,
                now_iso,
                str(updated_by or "desktop").strip() or "desktop",
            ),
        )

        if normalized_target_dwg_path:
            _upsert_lighting_schedule_link(
                conn,
                resolved_id,
                normalized_target_dwg_path,
                current_table_handle,
                payload.get("lastAppliedVersion", 0),
            )

        conn.commit()

    record = _get_lighting_schedule_record(resolved_id)
    if record is None:
        raise RuntimeError("Lighting schedule save completed but record could not be reloaded.")

    record["conflict"] = (
        expected_version is not None
        and str(expected_version).strip() != ""
        and int(expected_version) != previous_version
    )
    record["previousVersion"] = previous_version
    return record


def _migrate_project_lighting_schedules(projects):
    if not isinstance(projects, list) or not projects:
        return

    with _open_lighting_schedule_db() as conn:
        for project in projects:
            if not isinstance(project, dict):
                continue

            project_id = _resolve_lighting_schedule_project_id(project)
            if not project_id:
                continue

            existing = conn.execute(
                "SELECT version FROM lighting_schedule_projects WHERE project_id = ?",
                (project_id,),
            ).fetchone()
            if existing:
                continue

            legacy_schedule = _normalize_lighting_schedule_payload(
                project.get("lightingSchedule")
            )
            target_dwg_path = legacy_schedule.get("targetDwgPath") or ""
            now_iso = datetime.datetime.utcnow().isoformat()
            conn.execute(
                """
                INSERT INTO lighting_schedule_projects (
                    project_id, schedule_json, target_dwg_path, table_handle,
                    version, updated_at_utc, updated_by
                )
                VALUES (?, ?, ?, '', 1, ?, 'migration')
                ON CONFLICT(project_id) DO NOTHING
                """,
                (
                    project_id,
                    json.dumps(
                        {
                            "rows": legacy_schedule["rows"],
                            "generalNotes": legacy_schedule["generalNotes"],
                            "notes": legacy_schedule["notes"],
                        },
                        ensure_ascii=False,
                        separators=(",", ":"),
                    ),
                    target_dwg_path,
                    now_iso,
                ),
            )
            if target_dwg_path:
                _upsert_lighting_schedule_link(
                    conn,
                    project_id,
                    target_dwg_path,
                    "",
                    0,
                )
        conn.commit()


def _overlay_projects_with_lighting_schedule_records(projects):
    if not isinstance(projects, list):
        return projects

    _migrate_project_lighting_schedules(projects)
    for project in projects:
        if not isinstance(project, dict):
            continue

        project_id = _resolve_lighting_schedule_project_id(project)
        if not project_id:
            continue

        record = _get_lighting_schedule_record(project_id)
        if not record:
            continue

        existing_schedule = _normalize_lighting_schedule_payload(
            project.get("lightingSchedule")
        )
        project["lightingSchedule"] = {
            **existing_schedule,
            "rows": record["schedule"]["rows"],
            "generalNotes": record["schedule"]["generalNotes"],
            "notes": record["schedule"]["notes"],
            "targetDwgPath": record["targetDwgPath"],
            "_storeVersion": record["version"],
            "_storeUpdatedAtUtc": record["updatedAtUtc"],
            "_storeUpdatedBy": record["updatedBy"],
            "_tableHandle": record["tableHandle"],
        }
    return projects


def get_bundled_template_path(template_name: str) -> Path:
    """
    Resolve template files for both dev runs and packaged desktop runs.
    """
    candidates = [BASE_DIR / "templates" / template_name]

    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        candidates.extend([
            exe_dir / "_internal" / "templates" / template_name,
            exe_dir / "templates" / template_name,
        ])

    for candidate in candidates:
        if candidate.exists():
            return candidate

    return candidates[0]

EXPENSE_SHEET_NAME = "Project Expense Sheet"
EXPENSE_SECTION_ROWS = 12
EXPENSE_HEADER_ROWS = 3
EXPENSE_DEFAULT_DATA_ROWS = 5
EXPENSE_FOOTER_ROWS = 4
EXPENSE_INSERTED_DATA_TEMPLATE_OFFSET = EXPENSE_HEADER_ROWS + 1
EXPENSE_IMAGE_ROW_OFFSET = 6


DEFAULT_TEMPLATES = [
    {
        "name": "Narrative of Changes",
        "discipline": "General",
        "fileType": "docx",
        "sourcePath": str(get_bundled_template_path("Narrative of Changes.docx")),
        "description": "Standard narrative of changes template with MEP sections"
    },
    {
        "name": "Plan Check Comments",
        "discipline": "General",
        "fileType": "doc",
        "sourcePath": str(get_bundled_template_path("PCC.doc")),
        "description": "Plan check comments response table template"
    }
]
TEMPLATE_KEY_BY_NAME = {
    "narrative of changes": "narrative",
    "plan check comments": "planCheck",
    "plan check response letter": "planCheck",
}
TEMPLATE_DEFAULT_FILENAME_BY_KEY = {
    "narrative": "Narrative of Changes",
    "planCheck": "Plan Check Comments",
}

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# --- Circuit Breaker AI helpers ---
CB_TEMPLATE_PATH = BASE_DIR / "CircuitBreakerAI" / "ElectricalPanels" / "Template.xlsx"
CB_RATE_LIMIT_MAX_REQUESTS = 2
CB_RATE_LIMIT_WINDOW = 61
_cb_api_call_timestamps = []

CB_COL_L_NOTE = "B"
CB_COL_L_TYPE = "C"
CB_COL_L_POLE = "D"
CB_COL_L_TRIP = "E"
CB_COL_L_DESC = "F"
CB_COL_L_KVA = "I"
CB_COL_VOLTAGE = "G"
CB_COL_R_KVA = "K"
CB_COL_R_DESC = "L"
CB_COL_R_POLE = "O"
CB_COL_R_TRIP = "P"
CB_COL_R_TYPE = "Q"
CB_COL_R_NOTE = "R"


class CircuitItem(BaseModel):
    circuit_number: int = Field(
        ..., description="The first circuit number of the breaker")
    description: str = Field(..., description="Load description")
    breaker_amps: str = Field(..., description="Amperage (e.g., '20')")
    poles: int = Field(1, description="Number of poles (1, 2, or 3)")
    load_type: str = Field(..., description="Code: 'C', 'D', 'G', 'K', 'M'")


class PanelData(BaseModel):
    panel_name: str = Field(..., description="Panel Name")
    voltage: str = Field(..., description="Voltage")
    bus_rating: str = Field(..., description="Bus Rating")
    phase: str = Field(..., description="Phase")
    wire: str = Field(..., description="Wire")
    mounting: str = Field(..., description="Mounting")
    enclosure: str = Field(..., description="Enclosure")
    circuits: List[CircuitItem] = Field(
        ..., description="List of detected breakers")


def cb_enforce_rate_limit():
    global _cb_api_call_timestamps
    now = time.time()
    _cb_api_call_timestamps = [
        t for t in _cb_api_call_timestamps if now - t < CB_RATE_LIMIT_WINDOW
    ]
    if len(_cb_api_call_timestamps) >= CB_RATE_LIMIT_MAX_REQUESTS:
        earliest_call = _cb_api_call_timestamps[0]
        wait_time = (earliest_call + CB_RATE_LIMIT_WINDOW) - now
        if wait_time > 0:
            time.sleep(wait_time)
    _cb_api_call_timestamps.append(time.time())


def cb_clean_text(text):
    if text is None:
        return ""
    return str(text).upper().strip()


def cb_calculate_estimated_load(amps_str, description, load_type):
    desc = cb_clean_text(description)

    if any(x in desc for x in ["SPARE", "SPACE", "UNUSED"]):
        return 0.00

    try:
        numeric_amps = float("".join(filter(str.isdigit, str(amps_str))))
        breaker_kva = (numeric_amps * 120) / 1000
    except Exception:
        numeric_amps = 0
        breaker_kva = 0

    if any(x in desc for x in ["A/C", "AC", "AIR COND", "CONDENSER", "HVAC", "COOLING", "COMPRESSOR"]):
        return round(breaker_kva * 0.70, 2)

    if any(x in desc for x in ["WATER HEATER", "WH", "HWH", "HOT WATER"]):
        return round(breaker_kva * 0.60, 2)

    if any(x in desc for x in ["FAN", "EXHAUST", "VENTILAT", "VENT FAN"]):
        return round(breaker_kva * 0.50, 2)

    if "ATM" in desc:
        return 0.60

    if any(x in desc for x in ["POLE LIGHT", "POLE LT", "PARKING LOT", "LOT LIGHT"]):
        return random.choice([1.0, 1.1, 1.2])

    if any(x in desc for x in ["EXT LIGHT", "EXTERIOR LIGHT", "OUTSIDE LIGHT", "OUTDOOR LIGHT",
                                "EXT LT", "EXTERIOR LT", "SIGN", "FACADE", "BUILDING LIGHT"]):
        return random.choice([0.6, 0.7, 0.8, 0.9, 1.0])

    if any(x in desc for x in ["LIGHT", "LT", "LTG", "LIGHTING", "LAMP"]):
        large_room_keywords = ["LOBBY", "HALL", "CONFERENCE", "MEETING", "AUDITORIUM", "GYM",
                               "WAREHOUSE", "SHOWROOM", "DINING", "MAIN", "OPEN", "COMMON"]
        medium_room_keywords = ["OFFICE", "BREAK", "STORAGE", "KITCHEN", "LUNCH"]

        if any(kw in desc for kw in large_room_keywords):
            return random.choice([0.7, 0.8, 0.9])
        if any(kw in desc for kw in medium_room_keywords):
            return random.choice([0.5, 0.6, 0.7])
        return random.choice([0.3, 0.4, 0.5])

    if any(x in desc for x in ["PLUG", "RECEP", "DUPLEX", "OUTLET", "REC", "RCPT"]):
        high_recep_keywords = ["KITCHEN", "BREAK", "LAB", "WORKSTATION", "COMPUTER", "SERVER",
                               "OFFICE", "CONF", "MEETING", "NURSE", "MEDICAL"]
        medium_recep_keywords = ["STORAGE", "UTILITY", "MECH", "CLOSET", "REST", "BATH"]

        if any(kw in desc for kw in high_recep_keywords):
            return random.choice([0.72, 0.9])
        if any(kw in desc for kw in medium_recep_keywords):
            return random.choice([0.36, 0.54])
        return random.choice([0.36, 0.54, 0.72, 0.9])

    l_type = cb_clean_text(load_type)
    factor = 0.50 if l_type in ['C', 'G'] else 0.80
    return round((numeric_amps * factor * 120) / 1000, 2)


def cb_fix_nema_type(raw_type):
    t = str(raw_type).upper()
    return "NEMA 3R" if any(x in t for x in ["3R", "OUT", "EXT", "WEATHER"]) else "NEMA 1"


def cb_resolve_ditto_marks(circuits: List[CircuitItem]) -> List[CircuitItem]:
    ditto_pattern = re.compile(r'^[\""\u2018\u2019\u201c\u201d\.]+$|^(SAME|DO)$', re.IGNORECASE)
    odds = sorted([c for c in circuits if c.circuit_number % 2 != 0], key=lambda x: x.circuit_number)
    evens = sorted([c for c in circuits if c.circuit_number % 2 == 0], key=lambda x: x.circuit_number)

    def process_column(col_list):
        last_desc = ""
        for ckt in col_list:
            if ditto_pattern.match(ckt.description.strip()):
                if last_desc:
                    ckt.description = last_desc
            elif ckt.description.strip() and ckt.description.strip() != "---":
                last_desc = ckt.description
        return col_list

    return process_column(odds) + process_column(evens)


def cb_make_unique_sheet_name(raw_name, existing_names):
    base = cb_clean_text(raw_name) or "PANEL"
    base = base.replace("/", "-").replace("\\", "-")
    base = base[:31] if base else "PANEL"
    if base not in existing_names:
        return base
    suffix = 2
    while True:
        suffix_str = f"-{suffix}"
        trimmed = base[: max(1, 31 - len(suffix_str))]
        candidate = f"{trimmed}{suffix_str}"
        if candidate not in existing_names:
            return candidate
        suffix += 1


def cb_ensure_template_sheet(wb):
    if "TEMPLATE" in wb.sheetnames:
        return
    if not CB_TEMPLATE_PATH.exists():
        raise ValueError("Panel schedule template not found.")
    template_wb = openpyxl.load_workbook(CB_TEMPLATE_PATH)
    try:
        source = template_wb["TEMPLATE"] if "TEMPLATE" in template_wb.sheetnames else template_wb.active
        target = wb.create_sheet("TEMPLATE")
        WorksheetCopy(source, target).copy_worksheet()
    finally:
        template_wb.close()


def cb_update_excel_workbook(panel_data: PanelData, workbook_path: str) -> str:
    wb = openpyxl.load_workbook(workbook_path)
    cb_ensure_template_sheet(wb)

    source = wb["TEMPLATE"]
    target = wb.copy_worksheet(source)
    target.sheet_view.showGridLines = False

    safe_name = cb_make_unique_sheet_name(panel_data.panel_name, wb.sheetnames)
    target.title = safe_name
    target["A3"] = f"(E) PANEL '{safe_name}'"

    target[f"{CB_COL_VOLTAGE}2"] = cb_clean_text(panel_data.voltage)
    target["G3"] = cb_clean_text(panel_data.bus_rating)
    target["K2"] = cb_clean_text(panel_data.wire)
    target["K3"] = cb_clean_text(panel_data.phase)
    target["K4"] = cb_fix_nema_type(panel_data.enclosure)
    target["N2"] = cb_clean_text(panel_data.mounting)

    start_row = 8
    max_row = 28
    occupied_slots = set()

    circuits = cb_resolve_ditto_marks(panel_data.circuits or [])
    sorted_ckts = sorted(circuits, key=lambda x: x.circuit_number)

    for ckt in sorted_ckts:
        try:
            c_num = int(ckt.circuit_number)
        except Exception:
            continue
        if c_num <= 0 or c_num > 42:
            continue

        row_idx = start_row + math.ceil(c_num / 2) - 1
        if row_idx > max_row:
            continue

        is_odd = (c_num % 2 != 0)
        side = "L" if is_odd else "R"
        if (side, row_idx) in occupied_slots:
            continue

        if is_odd:
            c_note, c_type, c_pole, c_trip, c_desc, c_kva = (
                CB_COL_L_NOTE, CB_COL_L_TYPE, CB_COL_L_POLE, CB_COL_L_TRIP, CB_COL_L_DESC, CB_COL_L_KVA
            )
        else:
            c_note, c_type, c_pole, c_trip, c_desc, c_kva = (
                CB_COL_R_NOTE, CB_COL_R_TYPE, CB_COL_R_POLE, CB_COL_R_TRIP, CB_COL_R_DESC, CB_COL_R_KVA
            )

        desc_text = cb_clean_text(ckt.description)
        is_spare = any(x in desc_text for x in ["SPARE", "SPACE", "UNUSED"])

        kva_val = "" if is_spare else cb_calculate_estimated_load(
            ckt.breaker_amps, ckt.description, ckt.load_type
        )
        note_val = "" if is_spare else "1"
        type_val = "" if is_spare else cb_clean_text(ckt.load_type)

        target[f"{c_desc}{row_idx}"] = desc_text
        target[f"{c_pole}{row_idx}"] = ckt.poles
        target[f"{c_trip}{row_idx}"] = cb_clean_text(ckt.breaker_amps)
        target[f"{c_kva}{row_idx}"] = kva_val
        target[f"{c_note}{row_idx}"] = note_val
        target[f"{c_type}{row_idx}"] = type_val

        occupied_slots.add((side, row_idx))

        if ckt.poles and ckt.poles > 1:
            for i in range(1, ckt.poles):
                ext_row = row_idx + i
                if ext_row <= max_row:
                    target[f"{c_desc}{ext_row}"] = "---"
                    target[f"{c_pole}{ext_row}"] = "-"
                    target[f"{c_trip}{ext_row}"] = "-"
                    target[f"{c_kva}{ext_row}"] = kva_val
                    target[f"{c_type}{ext_row}"] = type_val
                    target[f"{c_note}{ext_row}"] = note_val
                    occupied_slots.add((side, ext_row))

    wb.save(workbook_path)
    return safe_name

# --- API Class ---


class Api:
    # --- FIX: Removed self.window from __init__ to prevent circular reference ---
    def __init__(self):
        # --- Configuration for Bundle Management ---
        self.github_repo = "jacobhusband/ElectricalCommands"
        # AutoCAD plugin releases are tracked in a separate repo; resolved dynamically.
        self.release_tag = None
        self.app_version = APP_VERSION
        self.app_update_repo = APP_UPDATE_REPO
        self.app_installer_name = APP_INSTALLER_NAME

        appdata_path = os.getenv('APPDATA')
        if not appdata_path:
            raise EnvironmentError(
                "CRITICAL: Could not determine the APPDATA directory.")
        self.app_plugins_folder = os.path.join(
            appdata_path, 'Autodesk', 'ApplicationPlugins')
        os.makedirs(self.app_plugins_folder, exist_ok=True)

        test_mode_raw = str(os.getenv('ACIES_TEST_MODE', '')).strip().lower()
        self.test_mode = test_mode_raw in ('1', 'true', 'yes', 'on')
        self._test_mode_records = []
        self._workroom_cad_file_cache = {}
        if self.test_mode:
            logging.info(
                "Api initialized in ACIES_TEST_MODE. CAD tools run in deterministic validation mode.")

    # --- Application update helpers ---
    def _fetch_latest_release(self):
        """Fetch latest release metadata for this application.

        GitHub returns 404 for /releases/latest when no published release exists,
        so we fall back to the first item from /releases.
        """
        endpoints = [
            f"{GITHUB_API_BASE}/repos/{self.app_update_repo}/releases/latest",
            f"{GITHUB_API_BASE}/repos/{self.app_update_repo}/releases?per_page=1"
        ]

        data = None
        last_error = None

        for url in endpoints:
            try:
                response = requests.get(url, timeout=10)
                if response.status_code == 404:
                    continue
                response.raise_for_status()
                payload = response.json()
                data = payload[0] if isinstance(
                    payload, list) and payload else payload
                if data:
                    break
            except Exception as e:
                last_error = e

        if not data:
            raise Exception(
                "No published releases found for the updater repo. "
                "Publish a release on GitHub so the updater can find it."
            ) from last_error

        tag_name = data.get('tag_name') or ''
        latest_version = _normalize_version(tag_name)

        asset = None
        for a in data.get('assets', []):
            if a.get('name', '').lower() == self.app_installer_name.lower():
                asset = a
                break

        download_url = asset.get('browser_download_url') if asset else None

        return {
            'tag': tag_name,
            'latest_version': latest_version,
            'download_url': download_url,
            'release_notes': data.get('body') or '',
            'html_url': data.get('html_url') or ''
        }

    def get_app_update_status(self):
        """Check GitHub for a newer installer."""
        try:
            release = self._fetch_latest_release()
            latest_version = release['latest_version']
            download_url = release['download_url']
            update_available = bool(download_url) and _is_remote_newer(
                latest_version, self.app_version)
            return {
                'status': 'success',
                'current_version': self.app_version,
                'latest_version': latest_version,
                'update_available': update_available,
                'download_url': download_url,
                'release_notes': release['release_notes'],
                'release_page': release['html_url'],
            }
        except Exception as e:
            logging.error(f"Update check failed: {e}")
            return {
                'status': 'error',
                'message': str(e),
                'current_version': self.app_version
            }

    def _fetch_latest_bundle_release(self):
        """Fetch latest AutoCAD plugin release (or latest tag as fallback)."""
        endpoints = [
            f"{GITHUB_API_BASE}/repos/{self.github_repo}/releases/latest",
            f"{GITHUB_API_BASE}/repos/{self.github_repo}/releases?per_page=1"
        ]

        data = None
        last_error = None

        for url in endpoints:
            try:
                response = requests.get(url, timeout=10)
                if response.status_code == 404:
                    continue
                response.raise_for_status()
                payload = response.json()
                data = payload[0] if isinstance(
                    payload, list) and payload else payload
                if data:
                    break
            except Exception as e:
                last_error = e

        if data:
            return {
                'tag': data.get('tag_name') or '',
                'assets': data.get('assets', []),
                'release_notes': data.get('body') or '',
                'html_url': data.get('html_url') or ''
            }

        # Fallback: look at tags if no releases are published yet.
        try:
            tag_resp = requests.get(
                f"{GITHUB_API_BASE}/repos/{self.github_repo}/tags?per_page=1",
                timeout=10
            )
            tag_resp.raise_for_status()
            tags = tag_resp.json()
            if isinstance(tags, list) and tags:
                tag_name = tags[0].get('name') or ''
                release_data = {}
                try:
                    rel_resp = requests.get(
                        f"{GITHUB_API_BASE}/repos/{self.github_repo}/releases/tags/{tag_name}",
                        timeout=10
                    )
                    if rel_resp.status_code != 404:
                        rel_resp.raise_for_status()
                        release_data = rel_resp.json() or {}
                except Exception as inner:
                    logging.warning(
                        f"Could not fetch release data for tag {tag_name}: {inner}")

                return {
                    'tag': tag_name,
                    'assets': release_data.get('assets', []),
                    'release_notes': release_data.get('body', ''),
                    'html_url': release_data.get('html_url', '')
                }
        except Exception as e:
            last_error = e

        raise Exception(
            "No published AutoCAD plugin releases or tags found in the plugin repository."
        ) from last_error

    def download_and_install_app_update(self, download_url=None):
        """Download the newest installer and launch it silently."""
        if not download_url:
            return {'status': 'error', 'message': 'No download URL provided (publish a release asset named acies-scheduler-setup.exe).'}

        target = Path(tempfile.gettempdir()) / self.app_installer_name
        app_path = sys.executable if getattr(sys, "frozen", False) else None

        try:
            with requests.get(download_url, stream=True, timeout=60) as r:
                r.raise_for_status()
                with open(target, 'wb') as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)

            # Prepare the restart command separately to avoid f-string backslash syntax error
            restart_cmd = f'Start-Process -FilePath "{app_path}"' if app_path else ""

            # Use a small helper process to show a message, run installer (visible), and relaunch app after completion
            ps_script = (
                '$msg = "Updating ACIES Scheduler...' + "`n" +
                'Installer will run and the app will reopen automatically." ; '
                'Add-Type -AssemblyName PresentationFramework; '
                '[System.Windows.MessageBox]::Show($msg, "Updating", "OK", "Information") | Out-Null; '
                f'Start-Process -FilePath "{target}" '
                f'-ArgumentList \'/SILENT /NORESTART /CLOSEAPPLICATIONS /RESTARTAPPLICATIONS\' -Wait; '
                'Start-Sleep -Seconds 2; '
                f'{restart_cmd}'
            )
            subprocess.Popen(
                ["powershell", "-NoProfile", "-WindowStyle",
                    "Hidden", "-Command", ps_script],
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            # Exit current app shortly to free files for installer
            if app_path:
                threading.Thread(target=lambda: (
                    time.sleep(1), os._exit(0)), daemon=True).start()

            return {
                'status': 'success',
                'message': 'Installer launched. Follow prompts to finish.',
                'installer_path': str(target)
            }
        except Exception as e:
            logging.error(f"Failed to download/install update: {e}")
            return {'status': 'error', 'message': str(e)}

    def check_bundle_installed(self, bundle_name):
        """
        Checks if a specific bundle directory exists in the ApplicationPlugins folder.
        Returns a simple boolean status.
        """
        if not bundle_name or not bundle_name.endswith('.bundle'):
            return {'status': 'error', 'message': 'Invalid bundle name provided.'}

        bundle_path = os.path.join(self.app_plugins_folder, bundle_name)
        is_installed = os.path.isdir(bundle_path)

        logging.info(
            f"Checking for bundle '{bundle_name}': {'Installed' if is_installed else 'Not Installed'}")

        return {'status': 'success', 'is_installed': is_installed}

    def _is_autocad_running(self):
        """Checks if the AutoCAD process (acad.exe) is currently running."""
        try:
            # Use tasklist to check for the process name.
            # This is standard on Windows and doesn't require installing psutil.
            output = subprocess.check_output(
                'tasklist /FI "IMAGENAME eq acad.exe" /FO CSV',
                shell=True, creationflags=subprocess.CREATE_NO_WINDOW
            ).decode()
            # If "acad.exe" is in the output, it's running.
            return "acad.exe" in output.lower()
        except Exception as e:
            logging.error(f"Error checking for AutoCAD process: {e}")
            return False  # Assume not running if check fails to avoid blocking user

    def get_bundle_statuses(self):
        """
        Fetches remote assets and compares against local bundles + version.txt.
        """
        logging.info("Fetching bundle statuses...")
        try:
            local_bundles = {
                name for name in os.listdir(self.app_plugins_folder)
                if str(name).endswith('.bundle')
            }

            release_tag = BUNDLE_RELEASE_TAG
            assets = []
            try:
                release_info = self._fetch_latest_bundle_release()
                self.release_tag = release_info.get('tag') or BUNDLE_RELEASE_TAG
                release_tag = self.release_tag
                assets = release_info.get('assets', []) or []
                if not assets and release_tag:
                    api_url = f"{GITHUB_API_BASE}/repos/{self.github_repo}/releases/tags/{release_tag}"
                    tag_response = requests.get(api_url, timeout=10)
                    if tag_response.status_code != 404:
                        tag_response.raise_for_status()
                        assets = tag_response.json().get('assets', [])
            except Exception as e:
                self.release_tag = release_tag
                logging.warning(
                    f"Could not refresh plugin release assets; falling back to known bundle catalog: {e}"
                )

            asset_by_bundle = {}
            for asset in assets:
                asset_name = asset.get('name')
                if not asset_name or 'Source code' in asset_name or not asset_name.endswith('.zip'):
                    continue
                bundle_name = _bundle_name_from_asset_name(asset_name, release_tag)
                if bundle_name:
                    asset_by_bundle[bundle_name] = asset

            bundle_names = list(KNOWN_PLUGIN_BUNDLES)
            for bundle_name in sorted(asset_by_bundle):
                if bundle_name not in bundle_names:
                    bundle_names.append(bundle_name)
            for bundle_name in sorted(local_bundles):
                if bundle_name not in bundle_names:
                    bundle_names.append(bundle_name)
            bundle_names = [
                bundle_name for bundle_name in bundle_names
                if bundle_name not in HIDDEN_PLUGIN_BUNDLES
            ]

            statuses = []
            for bundle_name in bundle_names:
                bundle_path = os.path.join(self.app_plugins_folder, bundle_name)
                is_installed = bundle_name in local_bundles
                local_version = None
                asset = asset_by_bundle.get(bundle_name)

                if is_installed:
                    version_file = os.path.join(bundle_path, 'version.txt')
                    if os.path.exists(version_file):
                        with open(version_file, 'r') as f:
                            local_version = f.read().strip()

                if asset:
                    if not is_installed:
                        state = 'not_installed'
                    elif local_version != release_tag:
                        state = 'update_available'
                    else:
                        state = 'installed'
                else:
                    state = 'installed' if is_installed else 'not_published'

                status = {
                    'name': bundle_name.replace('.bundle', ''),
                    'bundle_name': bundle_name,
                    'state': state,
                    'local_version': local_version or 'unknown',
                    'remote_version': release_tag,
                    'asset': asset,
                }
                statuses.append(status)

            return {'status': 'success', 'data': statuses}

        except Exception as e:
            logging.error(f"Error getting bundle statuses: {e}")
            return {'status': 'error', 'message': str(e)}

    def install_single_bundle(self, asset):
        """Downloads, extracts bundle, and writes version.txt."""

        # 1. Safety Check: AutoCAD locks files (DLLs). It must be closed to Update/Install.
        if self._is_autocad_running():
            return {
                'status': 'error',
                'message': 'AutoCAD is currently running. Please close AutoCAD and try again to prevent file locking errors.'
            }

        asset_name = asset.get('name')
        download_url = asset.get('browser_download_url')
        release_tag = self.release_tag or BUNDLE_RELEASE_TAG
        bundle_name = asset_name.replace(f"-{release_tag}.zip", ".bundle")
        if bundle_name == asset_name:
            bundle_name = asset_name.replace(".zip", ".bundle")
        extract_path = os.path.join(self.app_plugins_folder, bundle_name)

        logging.info(f"Installing bundle: {bundle_name}")
        try:
            response = requests.get(download_url, stream=True, timeout=30)
            response.raise_for_status()

            os.makedirs(extract_path, exist_ok=True)

            with zipfile.ZipFile(io.BytesIO(response.content)) as z:
                z.extractall(extract_path)

            # --- NEW: Write version.txt ---
            with open(os.path.join(extract_path, 'version.txt'), 'w') as f:
                f.write(release_tag)

            logging.info(
                f"Successfully installed {bundle_name} version {release_tag}")
            return {'status': 'success', 'bundle_name': bundle_name}
        except Exception as e:
            logging.error(f"Failed to install {bundle_name}: {e}")
            return {'status': 'error', 'message': f"Failed to install {bundle_name}: {e}"}

    def uninstall_bundle(self, bundle_name):
        """Removes a bundle directory."""

        # 1. Safety Check: AutoCAD locks files.
        if self._is_autocad_running():
            return {
                'status': 'error',
                'message': 'AutoCAD is currently running. Please close AutoCAD before uninstalling to ensure all files can be removed.'
            }

        if not bundle_name or not bundle_name.endswith('.bundle'):
            return {'status': 'error', 'message': 'Invalid bundle name.'}

        bundle_path = os.path.join(self.app_plugins_folder, bundle_name)

        try:
            if os.path.isdir(bundle_path):
                shutil.rmtree(bundle_path)
                return {'status': 'success', 'bundle_name': bundle_name}
            else:
                return {'status': 'error', 'message': 'Bundle not found.'}
        except Exception as e:
            logging.error(f"Failed to uninstall {bundle_name}: {e}")
            return {'status': 'error', 'message': f"Failed to uninstall {bundle_name}: {e}"}

    def _run_script_with_progress(self, command, tool_id):
        """
        Runs a script in a separate thread, captures its stdout, and sends
        progress updates to the frontend.
        """
        def script_runner():
            window = webview.windows[0]
            try:
                print(f"DEBUG THREAD: Starting script for {tool_id}")
                startupinfo = None
                if sys.platform == "win32":
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    startupinfo.wShowWindow = subprocess.SW_HIDE

                # FIX: Handle both string and list commands
                if isinstance(command, list):
                    process = subprocess.Popen(
                        command,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.STDOUT,
                        text=True,
                        encoding='utf-8',
                        errors='replace',
                        startupinfo=startupinfo
                    )
                    self._trace_cad_auto_select(
                        'script_subprocess_spawn',
                        tool_id=tool_id,
                        command_type='argv',
                        command=command,
                    )
                else:
                    process = subprocess.Popen(
                        command,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.STDOUT,
                        text=True,
                        encoding='utf-8',
                        errors='replace',
                        shell=True,
                        startupinfo=startupinfo
                    )
                    self._trace_cad_auto_select(
                        'script_subprocess_spawn',
                        tool_id=tool_id,
                        command_type='shell_string',
                        command=command,
                    )

                print(f"DEBUG THREAD: Process started, reading output...")
                for line in iter(process.stdout.readline, ''):
                    line = line.strip()
                    print(f"DEBUG THREAD OUTPUT: {line}")
                    if line.startswith("PROGRESS:"):
                        message = line[len("PROGRESS:"):].strip()
                        if message.startswith("TRACE"):
                            self._trace_cad_auto_select(
                                'script_trace',
                                tool_id=tool_id,
                                message=message,
                            )
                            continue
                        js_message = json.dumps(message)
                        window.evaluate_js(
                            f'window.updateToolStatus("{tool_id}", {js_message})')

                process.stdout.close()
                return_code = process.wait()
                print(f"DEBUG THREAD: Process finished with return code {return_code}")
                self._trace_cad_auto_select(
                    'script_subprocess_exit',
                    tool_id=tool_id,
                    return_code=return_code,
                )

                if return_code == 0:
                    window.evaluate_js(
                        f'window.updateToolStatus("{tool_id}", "DONE")')
                else:
                    error_message = f"Script finished with error code {return_code}."
                    js_error = json.dumps(error_message)
                    window.evaluate_js(
                        f'window.updateToolStatus("{tool_id}", "ERROR: " + {js_error})')

            except Exception as e:
                print(f"DEBUG THREAD ERROR: {e}")
                import traceback
                traceback.print_exc()
                logging.error(f"Failed to execute script for {tool_id}: {e}")
                js_error = json.dumps(str(e))
                window.evaluate_js(
                    f'window.updateToolStatus("{tool_id}", "ERROR: " + {js_error})')

        thread = threading.Thread(target=script_runner)
        thread.start()

    def _build_powershell_script_command(self, script_path, *args):
        command = [
            'powershell.exe',
            '-ExecutionPolicy',
            'Bypass',
            '-File',
            str(script_path),
        ]
        for arg in args:
            if arg is None:
                continue
            command.append(str(arg))
        return command

    def get_user_settings(self):
        """Reads and returns user settings from settings.json."""
        try:
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                settings = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            settings = build_default_user_settings()

        settings, sanitized = _sanitize_user_settings_payload(settings)
        repaired_settings, repaired = _merge_missing_user_settings_from_legacy(
            settings,
            _load_legacy_user_settings(),
        )
        repaired_settings, repaired_sanitized = _sanitize_user_settings_payload(
            repaired_settings
        )
        if sanitized or repaired or repaired_sanitized:
            if repaired:
                logging.info("Recovered missing user settings from the legacy AppData settings file.")
            self.save_user_settings(repaired_settings)
            return repaired_settings
        return repaired_settings

    def save_user_settings(self, data):
        """Saves user settings to settings.json."""
        try:
            normalized_data, _ = _sanitize_user_settings_payload(data)
            if os.path.exists(SETTINGS_FILE):
                shutil.copy2(SETTINGS_FILE, SETTINGS_FILE + '.bak')
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(normalized_data, f, ensure_ascii=False, indent=2)
            return {'status': 'success'}
        except Exception as e:
            logging.error(f"Error saving user settings: {e}")
            return {'status': 'error', 'message': str(e)}

    def get_cloud_sync_config(self):
        config = {
            "apiKey": str(os.getenv(FIREBASE_API_KEY_ENV) or "").strip(),
            "authDomain": str(os.getenv(FIREBASE_AUTH_DOMAIN_ENV) or "").strip(),
            "projectId": str(os.getenv(FIREBASE_PROJECT_ID_ENV) or "").strip(),
            "appId": str(os.getenv(FIREBASE_APP_ID_ENV) or "").strip(),
            "storageBucket": str(os.getenv(FIREBASE_STORAGE_BUCKET_ENV) or "").strip(),
            "messagingSenderId": str(os.getenv(FIREBASE_MESSAGING_SENDER_ID_ENV) or "").strip(),
        }
        enabled = all(
            config.get(key)
            for key in ("apiKey", "authDomain", "projectId", "appId")
        )
        return {
            "status": "success",
            "enabled": enabled,
            "config": config,
        }

    def get_local_sync_metadata(self):
        try:
            return {
                "status": "success",
                "files": _get_local_sync_metadata(),
            }
        except Exception as e:
            logging.error(f"Error loading local sync metadata: {e}")
            return {"status": "error", "message": str(e), "files": {}}

    def create_cloud_sync_backup(self, reason="", metadata=None):
        try:
            result = _create_cloud_sync_backup(reason=reason, metadata=metadata)
            return {
                "status": "success",
                **result,
            }
        except Exception as e:
            logging.error(f"Error creating cloud sync backup: {e}")
            return {"status": "error", "message": str(e)}

    def _get_google_oauth_client_id(self):
        return (os.getenv(GOOGLE_OAUTH_CLIENT_ID_ENV) or "").strip()

    def _get_google_oauth_client_secret(self):
        return (os.getenv(GOOGLE_OAUTH_CLIENT_SECRET_ENV) or "").strip()

    def _base64url_encode(self, raw_bytes):
        return base64.urlsafe_b64encode(raw_bytes).rstrip(b"=").decode("ascii")

    def _base64url_decode(self, raw_text):
        text = str(raw_text or "").strip()
        if not text:
            return b""
        padding = "=" * ((4 - (len(text) % 4)) % 4)
        return base64.urlsafe_b64decode(f"{text}{padding}")

    def _decode_jwt_payload(self, token):
        raw = str(token or "").strip()
        if not raw:
            return {}
        parts = raw.split(".")
        if len(parts) < 2:
            return {}
        try:
            payload = json.loads(self._base64url_decode(parts[1]).decode("utf-8"))
        except Exception:
            return {}
        return payload if isinstance(payload, dict) else {}

    def _generate_google_pkce_pair(self):
        verifier = secrets.token_urlsafe(64)
        challenge = self._base64url_encode(
            hashlib.sha256(verifier.encode("ascii")).digest()
        )
        return verifier, challenge

    def _extract_google_error_message(self, response):
        try:
            payload = response.json() or {}
        except Exception:
            payload = {}
        message = str(
            payload.get("error_description")
            or payload.get("error")
            or response.text
            or "Google sign-in failed."
        ).strip()
        if "client_secret is missing" in message.lower():
            if not self._get_google_oauth_client_secret():
                return (
                    "Google sign-in may require a client secret for this OAuth client. "
                    f"Add {GOOGLE_OAUTH_CLIENT_SECRET_ENV} to .env and restart the app, "
                    "or verify the Google OAuth client settings."
                )
            return (
                "Google sign-in was rejected by Google OAuth. Verify that "
                f"{GOOGLE_OAUTH_CLIENT_ID_ENV} and {GOOGLE_OAUTH_CLIENT_SECRET_ENV} "
                "match the same Google OAuth client."
            )
        return message

    def _build_google_auth_record(self, token_payload, profile_payload, existing_auth=None):
        existing_auth = existing_auth if isinstance(existing_auth, dict) else {}
        issued_at = datetime.datetime.now(datetime.timezone.utc)
        expires_in = 0
        try:
            expires_in = int(token_payload.get("expires_in") or 0)
        except Exception:
            expires_in = 0
        expires_at = (
            issued_at + datetime.timedelta(seconds=max(expires_in, 0))
        ).replace(microsecond=0).isoformat().replace("+00:00", "Z")

        refresh_token = str(token_payload.get("refresh_token") or "").strip()
        if not refresh_token:
            refresh_token = str(existing_auth.get("refreshToken") or "").strip()
        id_token = str(token_payload.get("id_token") or "").strip()
        if not id_token:
            id_token = str(existing_auth.get("idToken") or "").strip()

        return {
            "provider": "google",
            "subject": str(profile_payload.get("sub") or existing_auth.get("subject") or "").strip(),
            "email": str(profile_payload.get("email") or existing_auth.get("email") or "").strip(),
            "displayName": str(profile_payload.get("name") or existing_auth.get("displayName") or "").strip(),
            "avatarUrl": str(profile_payload.get("picture") or existing_auth.get("avatarUrl") or "").strip(),
            "signedInAt": str(existing_auth.get("signedInAt") or utc_now_iso()),
            "tokenIssuedAt": utc_now_iso(),
            "expiresAt": expires_at,
            "idToken": id_token,
            "accessToken": str(token_payload.get("access_token") or "").strip(),
            "refreshToken": refresh_token,
            "tokenType": str(token_payload.get("token_type") or "Bearer").strip() or "Bearer",
            "scope": str(token_payload.get("scope") or " ".join(GOOGLE_OAUTH_SCOPES)).strip(),
        }

    def _sanitize_google_auth_record(self, auth_record):
        if not isinstance(auth_record, dict):
            return {
                "signedIn": False,
                "provider": "google",
                "email": "",
                "displayName": "",
                "avatarUrl": "",
                "signedInAt": "",
                "expiresAt": "",
                "hasRefreshToken": False,
            }
        return {
            "signedIn": True,
            "provider": str(auth_record.get("provider") or "google").strip() or "google",
            "email": str(auth_record.get("email") or "").strip(),
            "displayName": str(auth_record.get("displayName") or "").strip(),
            "avatarUrl": str(auth_record.get("avatarUrl") or "").strip(),
            "signedInAt": str(auth_record.get("signedInAt") or "").strip(),
            "expiresAt": str(auth_record.get("expiresAt") or "").strip(),
            "hasRefreshToken": bool(str(auth_record.get("refreshToken") or "").strip()),
        }

    def _build_google_sync_session(self, auth_record):
        if not isinstance(auth_record, dict):
            return {
                "signedIn": False,
                "idToken": "",
                "accessToken": "",
                "firebaseReady": False,
                "auth": self._sanitize_google_auth_record(None),
            }
        id_token = str(auth_record.get("idToken") or "").strip()
        access_token = str(auth_record.get("accessToken") or "").strip()
        return {
            "signedIn": True,
            "idToken": id_token,
            "accessToken": access_token,
            "firebaseReady": bool(id_token or access_token),
            "auth": self._sanitize_google_auth_record(auth_record),
        }

    def _load_google_auth_record(self):
        settings = self.get_user_settings()
        auth_record = settings.get("googleAuth")
        return auth_record if isinstance(auth_record, dict) else None

    def _persist_google_auth_record(self, auth_record):
        settings = self.get_user_settings()
        settings["googleAuth"] = auth_record if isinstance(auth_record, dict) and auth_record else None
        result = self.save_user_settings(settings)
        if result.get("status") != "success":
            raise RuntimeError(result.get("message") or "Failed to save Google sign-in state.")
        return settings

    def _build_google_oauth_status_html(self, title, body, accent="#2563eb"):
        safe_title = html.escape(str(title or "").strip() or "Google Sign-In")
        safe_body = html.escape(str(body or "").strip() or "Return to the app.")
        return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{safe_title}</title>
  <style>
    body {{
      margin: 0;
      min-height: 100vh;
      display: grid;
      place-items: center;
      padding: 24px;
      background: #0f172a;
      color: #e2e8f0;
      font-family: Inter, Segoe UI, sans-serif;
    }}
    .card {{
      width: min(440px, 100%);
      border-radius: 20px;
      padding: 28px;
      background: rgba(15, 23, 42, 0.92);
      border: 1px solid rgba(148, 163, 184, 0.22);
      box-shadow: 0 24px 60px rgba(2, 6, 23, 0.45);
    }}
    h1 {{
      margin: 0 0 12px;
      font-size: 1.45rem;
      color: {accent};
    }}
    p {{
      margin: 0;
      line-height: 1.6;
      color: #cbd5e1;
    }}
  </style>
</head>
<body>
  <div class="card">
    <h1>{safe_title}</h1>
    <p>{safe_body}</p>
  </div>
</body>
</html>"""

    def _start_google_oauth_callback_server(self):
        class GoogleOAuthCallbackServer(http.server.ThreadingHTTPServer):
            allow_reuse_address = True

        class GoogleOAuthCallbackHandler(http.server.BaseHTTPRequestHandler):
            def log_message(self, fmt, *args):
                return

            def do_GET(self):
                parsed = urlparse(self.path)
                if parsed.path != GOOGLE_OAUTH_CALLBACK_PATH:
                    self.send_response(404)
                    self.send_header("Content-Type", "text/plain; charset=utf-8")
                    self.end_headers()
                    self.wfile.write(b"Not found.")
                    return

                params = parse_qs(parsed.query or "", keep_blank_values=True)
                self.server.oauth_result = {
                    "code": (params.get("code") or [""])[0],
                    "state": (params.get("state") or [""])[0],
                    "error": (params.get("error") or [""])[0],
                    "error_description": (params.get("error_description") or [""])[0],
                }

                error_code = str(self.server.oauth_result.get("error") or "").strip().lower()
                if error_code:
                    title = "Google Sign-In Cancelled" if error_code == "access_denied" else "Google Sign-In Failed"
                    body = (
                        "You can close this window and return to the app."
                        if error_code == "access_denied"
                        else "There was a problem completing Google sign-in. Return to the app for details."
                    )
                    page = self.server.owner._build_google_oauth_status_html(
                        title,
                        body,
                        "#f59e0b" if error_code == "access_denied" else "#ef4444",
                    )
                else:
                    page = self.server.owner._build_google_oauth_status_html(
                        "Google Sign-In Complete",
                        "Return to the app. Your sign-in will finish automatically.",
                    )

                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                self.wfile.write(page.encode("utf-8"))
                self.server.oauth_event.set()
                threading.Thread(target=self.server.shutdown, daemon=True).start()

        server = GoogleOAuthCallbackServer(("127.0.0.1", 0), GoogleOAuthCallbackHandler)
        server.owner = self
        server.oauth_event = threading.Event()
        server.oauth_result = {}
        server_thread = threading.Thread(target=server.serve_forever, daemon=True)
        server_thread.start()
        return server

    def _exchange_google_auth_code(self, client_id, code, code_verifier, redirect_uri):
        data = {
            "client_id": client_id,
            "code": code,
            "code_verifier": code_verifier,
            "grant_type": "authorization_code",
            "redirect_uri": redirect_uri,
        }
        client_secret = self._get_google_oauth_client_secret()
        if client_secret:
            data["client_secret"] = client_secret
        response = requests.post(
            GOOGLE_OAUTH_TOKEN_ENDPOINT,
            data=data,
            timeout=20,
        )
        if response.status_code != 200:
            raise RuntimeError(self._extract_google_error_message(response))
        payload = response.json() or {}
        if not str(payload.get("access_token") or "").strip():
            raise RuntimeError("Google sign-in completed without an access token.")
        return payload

    def _refresh_google_auth_record_if_needed(self, auth_record):
        if not isinstance(auth_record, dict):
            return None

        refresh_token = str(auth_record.get("refreshToken") or "").strip()
        client_id = self._get_google_oauth_client_id()
        expires_at = parse_utc_iso(auth_record.get("expiresAt"))
        if not refresh_token or not client_id or not expires_at:
            return auth_record

        now = datetime.datetime.now(datetime.timezone.utc)
        if expires_at > now + datetime.timedelta(minutes=5):
            return auth_record

        try:
            data = {
                "client_id": client_id,
                "grant_type": "refresh_token",
                "refresh_token": refresh_token,
            }
            client_secret = self._get_google_oauth_client_secret()
            if client_secret:
                data["client_secret"] = client_secret
            response = requests.post(
                GOOGLE_OAUTH_TOKEN_ENDPOINT,
                data=data,
                timeout=20,
            )
            if response.status_code != 200:
                message = self._extract_google_error_message(response)
                try:
                    payload = response.json() or {}
                except Exception:
                    payload = {}
                if str(payload.get("error") or "").strip().lower() == "invalid_grant":
                    self._persist_google_auth_record(None)
                    return None
                logging.warning(f"Could not refresh Google token: {message}")
                return auth_record

            token_payload = response.json() or {}
            refreshed_auth = self._build_google_auth_record(
                token_payload,
                {
                    "sub": auth_record.get("subject"),
                    "email": auth_record.get("email"),
                    "name": auth_record.get("displayName"),
                    "picture": auth_record.get("avatarUrl"),
                },
                existing_auth=auth_record,
            )
            refreshed_auth["signedInAt"] = str(auth_record.get("signedInAt") or utc_now_iso())
            self._persist_google_auth_record(refreshed_auth)
            return refreshed_auth
        except Exception as exc:
            logging.warning(f"Could not refresh Google auth state: {exc}")
            return auth_record

    def _fetch_google_user_profile(self, access_token):
        response = requests.get(
            GOOGLE_OAUTH_USERINFO_ENDPOINT,
            headers={"Authorization": f"Bearer {access_token}"},
            timeout=20,
        )
        if response.status_code != 200:
            raise RuntimeError(self._extract_google_error_message(response))
        payload = response.json() or {}
        if not str(payload.get("email") or "").strip():
            raise RuntimeError("Google sign-in completed without an email address.")
        return payload

    def get_google_auth_state(self):
        try:
            auth_record = self._refresh_google_auth_record_if_needed(
                self._load_google_auth_record()
            )
            return {
                "status": "success",
                "auth": self._sanitize_google_auth_record(auth_record),
            }
        except Exception as e:
            logging.error(f"Error loading Google auth state: {e}")
            return {
                "status": "error",
                "message": str(e),
                "auth": self._sanitize_google_auth_record(None),
            }

    def get_google_sync_session(self):
        try:
            auth_record = self._refresh_google_auth_record_if_needed(
                self._load_google_auth_record()
            )
            return {
                "status": "success",
                **self._build_google_sync_session(auth_record),
            }
        except Exception as e:
            logging.error(f"Error loading Google sync session: {e}")
            return {
                "status": "error",
                "message": str(e),
                **self._build_google_sync_session(None),
            }

    def sign_in_with_google(self):
        client_id = self._get_google_oauth_client_id()
        if not client_id:
            return {
                "status": "error",
                "message": (
                    "Google sign-in is not configured. Add "
                    f"{GOOGLE_OAUTH_CLIENT_ID_ENV} to .env and restart the app."
                ),
            }

        existing_auth = self._load_google_auth_record()
        callback_server = self._start_google_oauth_callback_server()
        callback_port = callback_server.server_address[1]
        redirect_uri = f"http://127.0.0.1:{callback_port}{GOOGLE_OAUTH_CALLBACK_PATH}"
        code_verifier, code_challenge = self._generate_google_pkce_pair()
        state = secrets.token_urlsafe(32)
        auth_url = (
            f"{GOOGLE_OAUTH_AUTH_ENDPOINT}?"
            + urlencode(
                {
                    "client_id": client_id,
                    "redirect_uri": redirect_uri,
                    "response_type": "code",
                    "scope": " ".join(GOOGLE_OAUTH_SCOPES),
                    "state": state,
                    "access_type": "offline",
                    "prompt": "consent select_account",
                    "code_challenge": code_challenge,
                    "code_challenge_method": "S256",
                }
            )
        )

        try:
            open_result = self.open_url(auth_url)
            if open_result.get("status") != "success":
                raise RuntimeError(open_result.get("message") or "Could not open the browser for Google sign-in.")

            if not callback_server.oauth_event.wait(GOOGLE_OAUTH_TIMEOUT_SECONDS):
                return {
                    "status": "error",
                    "message": "Google sign-in timed out. Please try again.",
                }

            result = callback_server.oauth_result or {}
            error_code = str(result.get("error") or "").strip().lower()
            if error_code:
                if error_code == "access_denied":
                    return {
                        "status": "cancelled",
                        "message": "Google sign-in was cancelled.",
                    }
                error_detail = str(result.get("error_description") or result.get("error") or "").strip()
                return {
                    "status": "error",
                    "message": error_detail or "Google sign-in did not complete.",
                }

            if str(result.get("state") or "").strip() != state:
                return {
                    "status": "error",
                    "message": "Google sign-in returned an invalid state token.",
                }

            code = str(result.get("code") or "").strip()
            if not code:
                return {
                    "status": "error",
                    "message": "Google sign-in finished without an authorization code.",
                }

            token_payload = self._exchange_google_auth_code(
                client_id,
                code,
                code_verifier,
                redirect_uri,
            )
            profile_payload = self._fetch_google_user_profile(
                str(token_payload.get("access_token") or "").strip()
            )
            auth_record = self._build_google_auth_record(
                token_payload,
                profile_payload,
                existing_auth=existing_auth,
            )
            self._persist_google_auth_record(auth_record)
            return {
                "status": "success",
                "auth": self._sanitize_google_auth_record(auth_record),
                "syncSession": self._build_google_sync_session(auth_record),
            }
        except Exception as e:
            logging.error(f"Error signing in with Google: {e}")
            return {'status': 'error', 'message': str(e)}
        finally:
            try:
                callback_server.shutdown()
            except Exception:
                pass
            try:
                callback_server.server_close()
            except Exception:
                pass

    def sign_out_google(self):
        try:
            auth_record = self._load_google_auth_record()
            revoke_token = ""
            if isinstance(auth_record, dict):
                revoke_token = str(
                    auth_record.get("refreshToken") or auth_record.get("accessToken") or ""
                ).strip()

            if revoke_token:
                try:
                    requests.post(
                        GOOGLE_OAUTH_REVOKE_ENDPOINT,
                        data={"token": revoke_token},
                        headers={"Content-Type": "application/x-www-form-urlencoded"},
                        timeout=10,
                    )
                except Exception as revoke_exc:
                    logging.warning(f"Google token revoke failed during sign-out: {revoke_exc}")

            self._persist_google_auth_record(None)
            return {
                "status": "success",
                "auth": self._sanitize_google_auth_record(None),
            }
        except Exception as e:
            logging.error(f"Error signing out of Google: {e}")
            return {'status': 'error', 'message': str(e)}

    def _normalize_outlook_scan_date(self, scan_date):
        raw = str(scan_date or "").strip()
        if not raw:
            return ""
        try:
            selected = datetime.date.fromisoformat(raw)
        except ValueError:
            return ""
        today = datetime.datetime.now().astimezone().date()
        if selected > today:
            selected = today
        return selected.isoformat()

    def _build_outlook_scan_range(self, timeframe="week", scan_date=""):
        now = datetime.datetime.now().astimezone()
        tzinfo = now.tzinfo
        normalized_scan_date = self._normalize_outlook_scan_date(scan_date)
        if normalized_scan_date:
            selected = datetime.date.fromisoformat(normalized_scan_date)
            start_local = datetime.datetime.combine(
                selected,
                datetime.time.min,
                tzinfo=tzinfo,
            )
            end_local = datetime.datetime.combine(
                selected,
                datetime.time.max,
                tzinfo=tzinfo,
            )
            return {
                "scanDate": normalized_scan_date,
                "timeframe": "",
                "startUtc": start_local.astimezone(datetime.timezone.utc),
                "endUtc": end_local.astimezone(datetime.timezone.utc),
            }

        today = now.date()
        normalized_timeframe = str(timeframe or "week").strip().lower()
        if normalized_timeframe not in {"week", "month"}:
            normalized_timeframe = "week"
        if normalized_timeframe == "month":
            start_local = datetime.datetime.combine(
                today.replace(day=1),
                datetime.time.min,
                tzinfo=tzinfo,
            )
        else:
            weekday = today.weekday()
            days_since_sunday = (weekday + 1) % 7
            start_date = today - datetime.timedelta(days=days_since_sunday)
            start_local = datetime.datetime.combine(
                start_date,
                datetime.time.min,
                tzinfo=tzinfo,
            )
        end_local = datetime.datetime.combine(
            today,
            datetime.time.max,
            tzinfo=tzinfo,
        )
        return {
            "scanDate": "",
            "timeframe": normalized_timeframe,
            "startUtc": start_local.astimezone(datetime.timezone.utc),
            "endUtc": end_local.astimezone(datetime.timezone.utc),
        }

    def _format_outlook_scan_period_label(self, timeframe="week", scan_date=""):
        normalized_scan_date = self._normalize_outlook_scan_date(scan_date)
        if normalized_scan_date:
            try:
                selected = datetime.date.fromisoformat(normalized_scan_date)
                return selected.strftime("%m/%d/%Y")
            except ValueError:
                pass
        return "this month" if str(timeframe or "").strip().lower() == "month" else "this week"

    def _build_outlook_scan_error_result(self, message, timeframe="week", source="", scan_date=""):
        return {
            "status": "error",
            "message": str(message or "").strip() or "Outlook scan failed.",
            "candidateCount": 0,
            "scannedCount": 0,
            "emailsIncludedCount": 0,
            "deliverablesIncludedCount": 0,
            "threadsDetected": 0,
            "dedupedEmailCount": 0,
            "dedupeSkippedEmailCount": 0,
            "analysisMode": "batch",
            "suggestions": [],
            "skippedMessages": [],
            "promptTruncated": False,
            "truncated": False,
            "timeframe": timeframe,
            "scanDate": self._normalize_outlook_scan_date(scan_date),
            "source": str(source or "").strip(),
        }

    def _count_outlook_scan_project_context_deliverables(self, project_context):
        total = 0
        if not isinstance(project_context, list):
            return total
        for project in project_context:
            if not isinstance(project, dict):
                continue
            deliverables = project.get("deliverables")
            if isinstance(deliverables, list):
                total += len(deliverables)
        return total

    def _notify_outlook_scan_progress(self, payload):
        try:
            if not webview.windows:
                return
            payload_json = json.dumps(payload or {}, ensure_ascii=True)
            webview.windows[0].evaluate_js(
                f"window.updateOutlookScanProgress({payload_json})"
            )
        except Exception as exc:
            logging.debug(f"_notify_outlook_scan_progress failed: {exc}")

    def _send_outlook_scan_progress(self, stage, message="", **kwargs):
        payload = {
            "stage": str(stage or "").strip(),
            "message": str(message or "").strip(),
        }
        allowed_keys = {
            "active",
            "source",
            "timeframe",
            "scanDate",
            "totalEmails",
            "processedEmails",
            "includedEmails",
            "skippedEmails",
            "deliverablesInPeriod",
            "relevantEmails",
            "threadsDetected",
            "dedupedEmailCount",
            "dedupeSkippedEmailCount",
        }
        for key, value in (kwargs or {}).items():
            if key not in allowed_keys or value is None:
                continue
            if key in {
                "totalEmails",
                "processedEmails",
                "includedEmails",
                "skippedEmails",
                "deliverablesInPeriod",
                "relevantEmails",
                "threadsDetected",
                "dedupedEmailCount",
                "dedupeSkippedEmailCount",
            }:
                try:
                    payload[key] = max(int(value), 0)
                except Exception:
                    continue
            else:
                payload[key] = value
        if "active" not in payload:
            payload["active"] = payload["stage"] not in {"done", "error"}
        self._notify_outlook_scan_progress(payload)

    def _get_outlook_scan_capability(self):
        desktop_available, desktop_reason = self._get_desktop_outlook_availability()
        return {
            "desktopAvailable": desktop_available,
            "desktopReason": desktop_reason,
        }

    def get_outlook_scan_capability(self):
        try:
            capability = self._get_outlook_scan_capability()
            return {
                "status": "success",
                "desktopAvailable": capability.get("desktopAvailable", False),
                "desktopReason": capability.get("desktopReason", ""),
            }
        except Exception as e:
            logging.error(f"Error loading Outlook scan capability: {e}")
            return {
                "status": "error",
                "message": str(e),
                "desktopAvailable": False,
                "desktopReason": str(e),
            }

    def _get_desktop_outlook_availability(self):
        if sys.platform != "win32":
            return False, "Desktop Outlook inbox scan is only available on Windows."
        install_path = self._find_desktop_outlook_install_path()
        if not install_path:
            return False, "Classic Outlook is not installed on this machine."
        try:
            import win32com.client  # noqa: F401
        except Exception:
            return False, "Desktop Outlook support requires pywin32."
        return True, ""

    def _find_desktop_outlook_install_path(self):
        if sys.platform != "win32":
            return ""

        candidate_paths = []
        try:
            import winreg

            registry_locations = [
                (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"),
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"),
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"),
            ]
            for hive, subkey in registry_locations:
                try:
                    with winreg.OpenKey(hive, subkey) as key:
                        value = ""
                        try:
                            value = winreg.QueryValue(key, None)
                        except Exception:
                            try:
                                value = winreg.QueryValueEx(key, "")[0]
                            except Exception:
                                value = ""
                        value = os.path.abspath(str(value or "").strip()) if value else ""
                        if value:
                            candidate_paths.append(value)
                except Exception:
                    continue
        except Exception:
            pass

        program_files_roots = [
            os.environ.get("ProgramFiles", r"C:\Program Files"),
            os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
        ]
        office_variants = [
            ("Microsoft Office", "root", "Office16"),
            ("Microsoft Office", "root", "Office15"),
            ("Microsoft Office", "Office16"),
            ("Microsoft Office", "Office15"),
            ("Microsoft Office", "Office14"),
        ]
        for root in program_files_roots:
            root = str(root or "").strip()
            if not root:
                continue
            for variant in office_variants:
                candidate_paths.append(os.path.join(root, *variant, "OUTLOOK.EXE"))

        for candidate in candidate_paths:
            full_path = os.path.abspath(str(candidate or "").strip()) if candidate else ""
            if full_path and os.path.exists(full_path):
                return full_path
        return ""

    def _run_with_outlook_com(self, func):
        pythoncom = None
        try:
            import pythoncom as _pythoncom

            pythoncom = _pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pythoncom = None

        try:
            return func()
        finally:
            if pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def _get_desktop_outlook_namespace(self):
        available, reason = self._get_desktop_outlook_availability()
        if not available:
            raise RuntimeError(reason or "Desktop Outlook is unavailable.")
        try:
            import win32com.client
        except Exception as exc:
            raise RuntimeError("Desktop Outlook support requires pywin32.") from exc

        try:
            application = win32com.client.Dispatch("Outlook.Application")
            namespace = application.GetNamespace("MAPI")
            return application, namespace
        except Exception as exc:
            raise RuntimeError(f"Could not access Desktop Outlook: {exc}") from exc

    def _coerce_outlook_datetime(self, raw_value):
        if isinstance(raw_value, datetime.datetime):
            dt_value = raw_value
        else:
            try:
                dt_value = datetime.datetime.fromtimestamp(float(raw_value))
            except Exception:
                return None
        if dt_value.tzinfo is None:
            dt_value = dt_value.replace(tzinfo=datetime.datetime.now().astimezone().tzinfo)
        try:
            return dt_value.astimezone(datetime.timezone.utc)
        except Exception:
            return None

    def _format_outlook_datetime(self, raw_value):
        dt_value = self._coerce_outlook_datetime(raw_value)
        if dt_value is None:
            return ""
        return dt_value.isoformat().replace("+00:00", "Z")

    def _get_desktop_outlook_internet_message_id(self, mail_item):
        accessor = getattr(mail_item, "PropertyAccessor", None)
        if accessor is None:
            return ""
        try:
            return str(accessor.GetProperty(OUTLOOK_MAPI_PR_INTERNET_MESSAGE_ID) or "").strip()
        except Exception:
            return ""

    def _build_desktop_outlook_message_summary(self, mail_item):
        body_text = str(getattr(mail_item, "Body", "") or "").strip()
        preview = re.sub(r"\s+", " ", body_text).strip()
        if len(preview) > 500:
            preview = preview[:497].rstrip() + "..."
        return {
            "id": str(getattr(mail_item, "EntryID", "") or "").strip(),
            "subject": str(getattr(mail_item, "Subject", "") or "").strip(),
            "bodyPreview": preview,
            "receivedDateTime": self._format_outlook_datetime(getattr(mail_item, "ReceivedTime", None)),
            "webLink": "",
            "internetMessageId": self._get_desktop_outlook_internet_message_id(mail_item),
            "conversationId": str(getattr(mail_item, "ConversationID", "") or "").strip(),
            "hasAttachments": bool(int(getattr(getattr(mail_item, "Attachments", None), "Count", 0) or 0)),
            "source": OUTLOOK_SCAN_SOURCE_DESKTOP,
            "from": {
                "name": str(getattr(mail_item, "SenderName", "") or "").strip(),
                "address": str(getattr(mail_item, "SenderEmailAddress", "") or "").strip(),
            },
        }

    def _list_desktop_outlook_inbox_messages(
        self,
        timeframe="week",
        limit=OUTLOOK_SCAN_FETCH_LIMIT,
        scan_date="",
    ):
        scan_range = self._build_outlook_scan_range(timeframe=timeframe, scan_date=scan_date)
        start_utc = scan_range["startUtc"]
        end_utc = scan_range["endUtc"]
        max_items = max(int(limit or 0), 1)

        def _read_messages():
            _, namespace = self._get_desktop_outlook_namespace()
            inbox = namespace.GetDefaultFolder(OUTLOOK_MAPI_INBOX_FOLDER_ID)
            items = inbox.Items
            try:
                items.Sort("[ReceivedTime]", True)
            except Exception:
                pass

            messages = []
            has_more = False
            current = None
            try:
                current = items.GetFirst()
            except Exception:
                current = None

            while current is not None:
                message_class = str(getattr(current, "MessageClass", "") or "").strip().lower()
                if message_class and not message_class.startswith("ipm.note"):
                    try:
                        current = items.GetNext()
                    except Exception:
                        current = None
                    continue

                received_dt = self._coerce_outlook_datetime(getattr(current, "ReceivedTime", None))
                if received_dt is not None and end_utc is not None and received_dt > end_utc:
                    try:
                        current = items.GetNext()
                    except Exception:
                        current = None
                    continue
                if received_dt is not None and received_dt < start_utc:
                    break

                summary = self._build_desktop_outlook_message_summary(current)
                if summary.get("id"):
                    messages.append(summary)
                    if len(messages) >= max_items:
                        has_more = True
                        break

                try:
                    current = items.GetNext()
                except Exception:
                    current = None

            return messages, has_more

        return self._run_with_outlook_com(_read_messages)

    def _get_desktop_outlook_message_body_text(self, message_id):
        target_id = str(message_id or "").strip()
        if not target_id:
            raise RuntimeError("Desktop Outlook message is missing an identifier.")

        def _read_message():
            _, namespace = self._get_desktop_outlook_namespace()
            mail_item = namespace.GetItemFromID(target_id)
            summary = self._build_desktop_outlook_message_summary(mail_item)
            body_text = str(getattr(mail_item, "Body", "") or "").strip()
            if not body_text:
                html_body = str(getattr(mail_item, "HTMLBody", "") or "").strip()
                if html_body:
                    body_text = self._html_to_plain_text(html_body)
            if not body_text:
                body_text = str(summary.get("bodyPreview") or "").strip()
            return summary, body_text

        return self._run_with_outlook_com(_read_message)

    def open_outlook_desktop_message(self, message_id):
        payload = message_id
        if isinstance(payload, str):
            payload = {"messageId": payload}
        if not isinstance(payload, dict):
            payload = {}
        target_id = str(payload.get("messageId") or payload.get("id") or "").strip()
        if not target_id:
            return {"status": "error", "message": "Missing Outlook desktop message identifier."}

        try:
            def _open_message():
                _, namespace = self._get_desktop_outlook_namespace()
                mail_item = namespace.GetItemFromID(target_id)
                mail_item.Display()

            self._run_with_outlook_com(_open_message)
            return {"status": "success"}
        except Exception as exc:
            logging.error(f"Error opening Desktop Outlook message: {exc}")
            return {"status": "error", "message": str(exc)}

    def _outlook_message_candidate_score(self, message_summary):
        summary = message_summary if isinstance(message_summary, dict) else {}
        subject = str(summary.get("subject") or "")
        preview = str(summary.get("bodyPreview") or "")
        sender = str((summary.get("from") or {}).get("address") or "")
        combined = f"{subject}\n{preview}".lower()
        if not combined.strip():
            return 0

        score = 1
        if re.search(r"\b\d{5,}\b", combined):
            score += 2
        if re.search(r"[A-Z]:\\\\", f"{subject}\n{preview}"):
            score += 2
        if re.search(
            r"\b(rfi|submittal|record drawings|record set|coordination|meeting|bulletin|revision|dd\d+|cd\d+|ifp|ifc|survey|due|deliverable)\b",
            combined,
        ):
            score += 2
        if re.search(r"\b(review|comments|deadline|schedule|submit|issuance|issue)\b", combined):
            score += 1
        if summary.get("hasAttachments"):
            score += 1
        if "noreply" in sender.lower() and score < 3:
            return 0
        return score

    def _html_to_plain_text(self, value):
        raw = str(value or "")
        if not raw:
            return ""
        stripped = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", raw)
        stripped = re.sub(r"(?i)<br\s*/?>", "\n", stripped)
        stripped = re.sub(r"(?i)</(p|div|li|tr|h[1-6])>", "\n", stripped)
        stripped = re.sub(r"(?s)<[^>]+>", " ", stripped)
        text = html.unescape(stripped)
        text = re.sub(r"[ \t\f\v]+", " ", text)
        text = re.sub(r"\r?\n\s*", "\n", text)
        return text.strip()

    def _trim_outlook_scan_prompt_text(self, value, limit=OUTLOOK_SCAN_EMAIL_BODY_PROMPT_CHARS):
        text = str(value or "").strip()
        if not text:
            return ""
        text = re.sub(r"\r\n?", "\n", text)
        text = re.sub(r"[ \t\f\v]+", " ", text)
        text = re.sub(r"\n{3,}", "\n\n", text)
        max_chars = max(int(limit or 0), 0)
        if not max_chars or len(text) <= max_chars:
            return text
        if max_chars <= 3:
            return text[:max_chars]
        return text[: max_chars - 3].rstrip() + "..."

    def _normalize_outlook_scan_dedupe_text(self, value):
        text = str(value or "").replace("\r\n", "\n").replace("\r", "\n")
        text = text.replace("\xa0", " ")
        text = re.sub(r"[ \t\f\v]+", " ", text)
        text = re.sub(r"[ \t]*\n[ \t]*", "\n", text)
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip()

    def _looks_like_outlook_reply_header_block(self, lines, start_idx):
        labels = []
        for idx in range(start_idx, min(start_idx + 6, len(lines))):
            raw_line = str(lines[idx] or "").strip()
            if not raw_line:
                if labels:
                    break
                continue
            match = re.match(r"^(from|sent|to|cc|subject):", raw_line, re.IGNORECASE)
            if not match:
                break
            label = match.group(1).lower()
            if label not in labels:
                labels.append(label)
        return "subject" in labels and ("from" in labels or "sent" in labels)

    def _strip_outlook_reply_history(self, value):
        text = str(value or "").replace("\r\n", "\n").replace("\r", "\n")
        if not text.strip():
            return ""
        lines = text.split("\n")
        kept_lines = []
        for idx, raw_line in enumerate(lines):
            line = str(raw_line or "")
            stripped = line.strip()
            if re.match(
                r"^-{2,}\s*(original message|forwarded message)\s*-{0,}\s*$",
                stripped,
                re.IGNORECASE,
            ):
                break
            if re.match(r"^begin forwarded message:?\s*$", stripped, re.IGNORECASE):
                break
            if re.match(r"^on .+ wrote:\s*$", stripped, re.IGNORECASE):
                break
            if self._looks_like_outlook_reply_header_block(lines, idx):
                break
            if stripped.startswith(">"):
                continue
            kept_lines.append(line)
        cleaned = "\n".join(kept_lines)
        return self._normalize_outlook_scan_dedupe_text(cleaned)

    def _extract_outlook_scan_unique_body(self, value):
        text = str(value or "")
        stripped = self._strip_outlook_reply_history(text)
        if stripped:
            text = stripped
        return self._normalize_outlook_scan_dedupe_text(text)

    def _looks_like_outlook_banner_block(self, value):
        block = self._normalize_outlook_scan_dedupe_text(value)
        if not block:
            return False
        lower = block.lower()
        if (
            "external email" in lower
            or "this email originated from outside" in lower
            or "use caution with links" in lower
        ):
            return len(block.split("\n")) <= 4
        return False

    def _looks_like_outlook_device_footer_line(self, value):
        line = str(value or "").strip().lower()
        if not line:
            return False
        return bool(
            re.match(r"^sent from my (iphone|ipad|android|mobile device).*$", line)
            or re.match(r"^get outlook for (ios|android)$", line)
            or line == "sent from outlook for ios"
        )

    def _looks_like_outlook_signature_intro_line(self, value):
        line = str(value or "").strip()
        lower = line.lower()
        if not lower:
            return False
        if line in {"--", "__", "___"}:
            return True
        if self._looks_like_outlook_device_footer_line(lower):
            return True
        return bool(
            re.match(
                r"^(thanks|thanks again|many thanks|thank you|best|best regards|kind regards|warm regards|regards|sincerely|respectfully|cheers|all the best)[,!.]?$",
                lower,
            )
        )

    def _looks_like_outlook_disclaimer_block(self, value):
        block = self._normalize_outlook_scan_dedupe_text(value)
        if not block:
            return False
        if self._looks_like_outlook_banner_block(block):
            return True
        lines = [str(line or "").strip() for line in block.split("\n") if str(line or "").strip()]
        if lines and all(self._looks_like_outlook_device_footer_line(line) for line in lines):
            return True
        lower = block.lower()
        phrases = (
            "this email and any attachments",
            "this message and any attachments",
            "intended only for the named recipient",
            "intended only for the addressee",
            "intended only for the person",
            "if you are not the intended recipient",
            "please delete this email",
            "please notify the sender",
            "confidential",
            "privileged",
            "disclosure",
            "do not disseminate",
            "do not copy",
            "do not distribute",
            "views or opinions expressed",
            "environment before printing",
        )
        matches = sum(1 for phrase in phrases if phrase in lower)
        return matches >= 2 or (
            matches >= 1
            and len(block) >= 80
            and ("recipient" in lower or "confidential" in lower or "privileged" in lower)
        )

    def _looks_like_outlook_signature_block(self, value):
        block = self._normalize_outlook_scan_dedupe_text(value)
        if not block:
            return False
        lines = [str(line or "").strip() for line in block.split("\n") if str(line or "").strip()]
        if not lines or len(lines) > 8:
            return False
        if self._looks_like_outlook_signature_intro_line(lines[0]):
            return True
        lower = block.lower()
        short_lines = sum(1 for line in lines if len(line) <= 60)
        has_contact = bool(
            re.search(r"@|https?://|www\.|\b(?:cell|mobile|office|direct|phone|tel|fax)\b", lower)
        )
        title_hits = sum(
            1
            for keyword_pattern in (
                r"\bproject manager\b",
                r"\bmanager\b",
                r"\bengineer\b",
                r"\bdesigner\b",
                r"\bprincipal\b",
                r"\bassociate\b",
                r"\bcoordinator\b",
                r"\bdirector\b",
                r"\bconsultant\b",
                r"\bengineering\b",
                r"\bsolutions\b",
                r"\binc\.?\b",
                r"\bllc\b",
            )
            if re.search(keyword_pattern, lower)
        )
        if has_contact and short_lines >= max(2, len(lines) - 1):
            return True
        return title_hits >= 2 and short_lines >= max(2, len(lines) - 1)

    def _strip_outlook_noise_blocks(self, value):
        blocks = [
            self._normalize_outlook_scan_dedupe_text(raw_block)
            for raw_block in re.split(r"\n\s*\n+", str(value or ""))
        ]
        blocks = [block for block in blocks if block]
        while blocks and self._looks_like_outlook_banner_block(blocks[0]):
            blocks.pop(0)
        while len(blocks) > 1 and (
            self._looks_like_outlook_disclaimer_block(blocks[-1])
            or self._looks_like_outlook_signature_block(blocks[-1])
        ):
            blocks.pop()
        return "\n\n".join(blocks).strip()

    def _strip_outlook_signature_lines(self, value):
        normalized = self._normalize_outlook_scan_dedupe_text(value)
        if not normalized:
            return ""
        lines = normalized.split("\n")
        non_empty_indexes = [idx for idx, line in enumerate(lines) if str(line or "").strip()]
        if len(non_empty_indexes) < 2:
            return normalized
        search_start = max(0, non_empty_indexes[max(len(non_empty_indexes) - 8, 0)])
        for idx in range(search_start, len(lines)):
            if not self._looks_like_outlook_signature_intro_line(lines[idx]):
                continue
            candidate = self._normalize_outlook_scan_dedupe_text("\n".join(lines[:idx]))
            if candidate:
                return candidate
        return normalized

    def _build_outlook_scan_analysis_body(self, value):
        unique_body = self._extract_outlook_scan_unique_body(value)
        if not unique_body:
            return ""
        cleaned = self._strip_outlook_noise_blocks(unique_body)
        cleaned = self._strip_outlook_signature_lines(cleaned)
        cleaned = self._strip_outlook_noise_blocks(cleaned)
        cleaned = self._normalize_outlook_scan_dedupe_text(cleaned)
        return cleaned or unique_body

    def _build_outlook_scan_prompt_payload(
        self,
        entry,
        body_limit=OUTLOOK_SCAN_EMAIL_BODY_PROMPT_CHARS,
    ):
        entry = entry if isinstance(entry, dict) else {}
        message = entry.get("message") if isinstance(entry.get("message"), dict) else {}
        existing_payload = (
            entry.get("promptPayload") if isinstance(entry.get("promptPayload"), dict) else {}
        )
        prompt_ref = str(entry.get("promptRef") or existing_payload.get("messageRef") or "").strip()
        body_source = str(
            entry.get("analysisBody") or existing_payload.get("body") or ""
        ).strip()
        message_from = message.get("from") if isinstance(message.get("from"), dict) else {}
        payload_from = (
            existing_payload.get("from") if isinstance(existing_payload.get("from"), dict) else {}
        )
        return {
            "messageRef": prompt_ref,
            "subject": str(message.get("subject") or existing_payload.get("subject") or "").strip(),
            "from": {
                "name": str(message_from.get("name") or payload_from.get("name") or "").strip(),
            },
            "receivedDateTime": str(
                message.get("receivedDateTime") or existing_payload.get("receivedDateTime") or ""
            ).strip(),
            "body": self._trim_outlook_scan_prompt_text(body_source, body_limit),
        }

    def _build_outlook_scan_prompt_config(self, reduced=False):
        if reduced:
            return {
                "bodyLimit": OUTLOOK_SCAN_RETRY_EMAIL_BODY_PROMPT_CHARS,
                "emailBudgetChars": OUTLOOK_SCAN_RETRY_PROMPT_EMAIL_BUDGET_CHARS,
                "maxEmails": OUTLOOK_SCAN_RETRY_AI_LIMIT,
                "skipReason": OUTLOOK_SCAN_RETRY_SKIP_REASON,
            }
        return {
            "bodyLimit": OUTLOOK_SCAN_EMAIL_BODY_PROMPT_CHARS,
            "emailBudgetChars": OUTLOOK_SCAN_PROMPT_EMAIL_BUDGET_CHARS,
            "maxEmails": OUTLOOK_SCAN_AI_LIMIT,
            "skipReason": OUTLOOK_SCAN_PROMPT_SKIP_REASON,
        }

    def _is_outlook_scan_timeout_error(self, error):
        message = str(error or "").strip().lower()
        return (
            "deadline_exceeded" in message
            or "deadline exceeded" in message
            or "deadline expired" in message
            or "timed out" in message
            or "timeout" in message
            or ("504" in message and "deadline" in message)
        )

    def _dedupe_outlook_scan_skipped_messages(self, skipped_messages):
        if not isinstance(skipped_messages, list):
            return []
        deduped = []
        seen = set()
        for skipped in skipped_messages:
            if not isinstance(skipped, dict):
                continue
            message = skipped.get("message") if isinstance(skipped.get("message"), dict) else {}
            reason = str(skipped.get("reason") or "").strip()
            if reason == OUTLOOK_SCAN_RETRY_SKIP_REASON:
                reason = OUTLOOK_SCAN_PROMPT_SKIP_REASON
            message_id = str(message.get("id") or "").strip().lower()
            subject = str(message.get("subject") or "").strip().lower()
            message_key = message_id or subject or json.dumps(message, ensure_ascii=True, sort_keys=True)
            key = (message_key, reason)
            if key in seen:
                continue
            seen.add(key)
            deduped.append({
                "message": message,
                "reason": reason,
            })
        return deduped

    def _split_outlook_scan_dedupe_blocks(self, value):
        text = self._normalize_outlook_scan_dedupe_text(value)
        if not text:
            return []
        blocks = []
        seen_keys = set()
        for raw_block in re.split(r"\n\s*\n+", text):
            block_text = self._normalize_outlook_scan_dedupe_text(raw_block)
            if not block_text:
                continue
            block_key = re.sub(r"\s+", " ", block_text).strip().lower()
            if not block_key or block_key in seen_keys:
                continue
            seen_keys.add(block_key)
            blocks.append({
                "text": block_text,
                "key": block_key,
            })
        return blocks

    def _dedupe_outlook_scan_entries(self, hydrated_entries):
        if not isinstance(hydrated_entries, list):
            return [], [], {
                "threadsDetected": 0,
                "dedupedEmailCount": 0,
                "dedupeSkippedEmailCount": 0,
            }

        entries_by_thread = {}
        threadless_entries = []
        for entry in hydrated_entries:
            if not isinstance(entry, dict):
                continue
            message = entry.get("message") if isinstance(entry.get("message"), dict) else {}
            prompt_payload = (
                entry.get("promptPayload")
                if isinstance(entry.get("promptPayload"), dict)
                else {}
            )
            conversation_id = str(
                message.get("conversationId") or prompt_payload.get("conversationId") or ""
            ).strip()
            if conversation_id:
                entries_by_thread.setdefault(conversation_id.lower(), []).append(entry)
            else:
                threadless_entries.append(entry)

        deduped_entries = {}
        skipped_messages = []
        deduped_email_count = 0
        dedupe_skipped_email_count = 0

        for thread_entries in entries_by_thread.values():
            seen_blocks = set()
            ordered_entries = sorted(
                thread_entries,
                key=lambda entry: (
                    -(entry.get("receivedSort") or 0),
                    str(entry.get("promptRef") or ""),
                ),
            )
            for entry in ordered_entries:
                prompt_payload = (
                    entry.get("promptPayload")
                    if isinstance(entry.get("promptPayload"), dict)
                    else {}
                )
                raw_body = str(entry.get("analysisBody") or prompt_payload.get("body") or "").strip()
                normalized_original = self._normalize_outlook_scan_dedupe_text(raw_body)
                unique_body = self._extract_outlook_scan_unique_body(raw_body)
                deduped_blocks = []
                for block in self._split_outlook_scan_dedupe_blocks(unique_body):
                    if block["key"] in seen_blocks:
                        continue
                    seen_blocks.add(block["key"])
                    deduped_blocks.append(block["text"])
                deduped_body = "\n\n".join(deduped_blocks).strip()
                if not deduped_body:
                    dedupe_skipped_email_count += 1
                    skipped_messages.append({
                        "message": entry.get("message") or {},
                        "reason": OUTLOOK_SCAN_DEDUPE_SKIP_REASON_THREAD,
                    })
                    continue
                normalized_deduped = self._normalize_outlook_scan_dedupe_text(deduped_body)
                if normalized_deduped and normalized_deduped != normalized_original:
                    deduped_email_count += 1
                deduped_entries[str(entry.get("promptRef") or "")] = {
                    **entry,
                    "analysisBody": deduped_body,
                }

        seen_threadless_bodies = set()
        for entry in threadless_entries:
            prompt_payload = (
                entry.get("promptPayload")
                if isinstance(entry.get("promptPayload"), dict)
                else {}
            )
            raw_body = str(entry.get("analysisBody") or prompt_payload.get("body") or "").strip()
            normalized_original = self._normalize_outlook_scan_dedupe_text(raw_body)
            deduped_body = self._extract_outlook_scan_unique_body(raw_body)
            normalized_deduped = self._normalize_outlook_scan_dedupe_text(deduped_body)
            if not normalized_deduped:
                dedupe_skipped_email_count += 1
                skipped_messages.append({
                    "message": entry.get("message") or {},
                    "reason": OUTLOOK_SCAN_DEDUPE_SKIP_REASON_THREAD,
                })
                continue
            content_key = re.sub(r"\s+", " ", normalized_deduped).strip().lower()
            if content_key in seen_threadless_bodies:
                dedupe_skipped_email_count += 1
                skipped_messages.append({
                    "message": entry.get("message") or {},
                    "reason": OUTLOOK_SCAN_DEDUPE_SKIP_REASON_DUPLICATE,
                })
                continue
            seen_threadless_bodies.add(content_key)
            if normalized_deduped != normalized_original:
                deduped_email_count += 1
            deduped_entries[str(entry.get("promptRef") or "")] = {
                **entry,
                "analysisBody": deduped_body,
            }

        ordered_entries = []
        for entry in hydrated_entries:
            prompt_ref = str(entry.get("promptRef") or "")
            if prompt_ref in deduped_entries:
                ordered_entries.append(deduped_entries[prompt_ref])

        return ordered_entries, skipped_messages, {
            "threadsDetected": len(entries_by_thread),
            "dedupedEmailCount": deduped_email_count,
            "dedupeSkippedEmailCount": dedupe_skipped_email_count,
        }

    def _normalize_outlook_scan_task_list(self, raw_tasks):
        normalized_tasks = []
        if not isinstance(raw_tasks, list):
            return normalized_tasks
        for raw_task in raw_tasks:
            if isinstance(raw_task, dict):
                text = str(raw_task.get("text") or "").strip()
                done = bool(raw_task.get("done"))
                links = raw_task.get("links") if isinstance(raw_task.get("links"), list) else []
            else:
                text = str(raw_task or "").strip()
                done = False
                links = []
            if not text:
                continue
            normalized_tasks.append({
                "text": text,
                "done": done,
                "links": links,
            })
        return normalized_tasks

    def _normalize_outlook_scan_project_context(self, raw_context):
        if not isinstance(raw_context, list):
            return []

        normalized = []
        for raw_project in raw_context:
            if not isinstance(raw_project, dict):
                continue
            deliverables = []
            raw_deliverables = raw_project.get("deliverables")
            raw_deliverables = raw_deliverables if isinstance(raw_deliverables, list) else []
            for raw_deliverable in raw_deliverables:
                if not isinstance(raw_deliverable, dict):
                    continue
                name = str(raw_deliverable.get("name") or "").strip()
                due = str(raw_deliverable.get("due") or "").strip()
                status = str(raw_deliverable.get("status") or "").strip()
                if not (name or due or status):
                    continue
                deliverables.append({
                    "name": name,
                    "due": due,
                    "status": status,
                })
            deliverables.sort(
                key=lambda deliverable: (
                    1 if parse_due_str(deliverable.get("due")) is None else 0,
                    parse_due_str(deliverable.get("due")) or datetime.datetime.max,
                    str(deliverable.get("name") or "").lower(),
                )
            )
            project = {
                "id": str(raw_project.get("id") or "").strip(),
                "name": str(raw_project.get("name") or "").strip(),
                "nick": str(raw_project.get("nick") or "").strip(),
                "path": str(raw_project.get("path") or "").strip(),
                "deliverables": deliverables,
            }
            if any(project.get(key) for key in ("id", "name", "nick", "path")) or deliverables:
                normalized.append(project)

        normalized.sort(
            key=lambda project: (
                min(
                    (
                        parse_due_str(deliverable.get("due"))
                        for deliverable in project.get("deliverables") or []
                        if parse_due_str(deliverable.get("due")) is not None
                    ),
                    default=datetime.datetime.max,
                ),
                str(project.get("id") or "").lower(),
                str(project.get("name") or "").lower(),
                str(project.get("path") or "").lower(),
            )
        )
        return normalized

    def _select_outlook_scan_prompt_context(
        self,
        hydrated_entries,
        project_context,
        timeframe="week",
        scan_date="",
        body_limit=OUTLOOK_SCAN_EMAIL_BODY_PROMPT_CHARS,
        email_budget_chars=OUTLOOK_SCAN_PROMPT_EMAIL_BUDGET_CHARS,
        max_emails=OUTLOOK_SCAN_AI_LIMIT,
        skip_reason=OUTLOOK_SCAN_PROMPT_SKIP_REASON,
    ):
        prioritized_emails = sorted(
            hydrated_entries,
            key=lambda entry: (
                -int(entry.get("score") or 0),
                -(entry.get("receivedSort") or 0),
                str(entry.get("promptRef") or ""),
            ),
        )

        flat_deliverables = []
        for project in project_context:
            for deliverable in project.get("deliverables") or []:
                flat_deliverables.append({
                    "projectId": str(project.get("id") or "").strip(),
                    "projectName": str(project.get("name") or "").strip(),
                    "projectNick": str(project.get("nick") or "").strip(),
                    "projectPath": str(project.get("path") or "").strip(),
                    "deliverable": str(deliverable.get("name") or "").strip(),
                    "due": str(deliverable.get("due") or "").strip(),
                    "status": str(deliverable.get("status") or "").strip(),
                })
        flat_deliverables.sort(
            key=lambda deliverable: (
                1 if parse_due_str(deliverable.get("due")) is None else 0,
                parse_due_str(deliverable.get("due")) or datetime.datetime.max,
                str(deliverable.get("projectId") or "").lower(),
                str(deliverable.get("projectName") or "").lower(),
                str(deliverable.get("deliverable") or "").lower(),
            )
        )

        included_emails = []
        email_chars = 0
        skipped_messages = []
        for entry in prioritized_emails:
            max_email_count = max(int(max_emails or 0), 0)
            if max_email_count and len(included_emails) >= max_email_count:
                skipped_messages.append({
                    "message": entry.get("message") or {},
                    "reason": skip_reason,
                })
                continue
            prompt_payload = self._build_outlook_scan_prompt_payload(
                entry,
                body_limit=body_limit,
            )
            serialized = json.dumps(prompt_payload, ensure_ascii=True, separators=(",", ":"))
            if included_emails and email_chars + len(serialized) > email_budget_chars:
                skipped_messages.append({
                    "message": entry.get("message") or {},
                    "reason": skip_reason,
                })
                continue
            included_emails.append({
                **entry,
                "promptPayload": prompt_payload,
            })
            email_chars += len(serialized)

        included_deliverables = []
        deliverable_chars = 0
        for deliverable in flat_deliverables:
            serialized = json.dumps(
                deliverable,
                ensure_ascii=True,
                separators=(",", ":"),
            )
            if included_deliverables and deliverable_chars + len(serialized) > OUTLOOK_SCAN_PROMPT_DELIVERABLE_BUDGET_CHARS:
                continue
            included_deliverables.append(deliverable)
            deliverable_chars += len(serialized)

        prompt_truncated = (
            len(included_emails) < len(prioritized_emails)
            or len(included_deliverables) < len(flat_deliverables)
        )

        while True:
            prompt = self._build_outlook_scan_batch_prompt(
                included_emails,
                included_deliverables,
                "",
                [],
                timeframe,
                scan_date=scan_date,
                prompt_truncated=prompt_truncated,
            )
            if len(prompt) <= OUTLOOK_SCAN_PROMPT_MAX_CHARS:
                break
            if len(included_emails) > 1:
                removed = included_emails.pop()
                skipped_messages.append({
                    "message": removed.get("message") or {},
                    "reason": skip_reason,
                })
                prompt_truncated = True
                continue
            if included_deliverables:
                included_deliverables.pop()
                prompt_truncated = True
                continue
            break

        included_email_refs = {str(entry.get("promptRef") or "") for entry in included_emails}
        trimmed_messages = [
            skipped
            for skipped in skipped_messages
            if str(((skipped.get("message") or {}).get("id")) or "").strip()
        ]

        seen_message_ids = {
            str(((skipped.get("message") or {}).get("id")) or "").strip().lower()
            for skipped in trimmed_messages
        }
        for entry in prioritized_emails:
            prompt_ref = str(entry.get("promptRef") or "")
            message_id = str(((entry.get("message") or {}).get("id")) or "").strip().lower()
            if prompt_ref in included_email_refs or (message_id and message_id in seen_message_ids):
                continue
            trimmed_messages.append({
                "message": entry.get("message") or {},
                "reason": skip_reason,
            })
            if message_id:
                seen_message_ids.add(message_id)

        return included_emails, included_deliverables, prompt_truncated, trimmed_messages

    def _build_outlook_scan_batch_prompt(
        self,
        included_emails,
        included_deliverables,
        user_name,
        discipline,
        timeframe,
        scan_date="",
        prompt_truncated=False,
    ):
        current_date = datetime.date.today().strftime("%m/%d/%Y")
        disciplines_str = ", ".join(discipline) if isinstance(discipline, list) else (discipline or "Engineering")
        task_notes_contract = self._build_deliverable_task_notes_prompt_contract(
            disciplines_str,
        )
        period_label = self._format_outlook_scan_period_label(
            timeframe=timeframe,
            scan_date=scan_date,
        )
        if self._normalize_outlook_scan_date(scan_date):
            period_instruction = f"received on {period_label}"
            deliverable_instruction = "due on that same day"
        else:
            period_instruction = f"for {period_label}"
            deliverable_instruction = "due in that same period"
        email_payload = [
            entry.get("promptPayload") or {}
            for entry in included_emails
            if isinstance(entry, dict)
        ]
        deliverable_payload = [
            deliverable
            for deliverable in included_deliverables
            if isinstance(deliverable, dict)
        ]
        email_json = json.dumps(email_payload, ensure_ascii=True, separators=(",", ":"))
        deliverable_json = json.dumps(deliverable_payload, ensure_ascii=True, separators=(",", ":"))
        truncation_note = (
            "The provided emails or deliverables were trimmed to fit one AI request. "
            "Only reason over the included data."
            if prompt_truncated
            else "The included emails and deliverables are the full context provided for this scan."
        )
        return f"""
You are an intelligent assistant for {user_name}, a(n) {disciplines_str} engineering project manager.
Review the batched Outlook emails {period_instruction} and compare them against the current deliverables {deliverable_instruction}.
Suggest only NEW deliverables that should be added, if any. Do not suggest a deliverable if the same or an equivalent deliverable already appears in CURRENT_DELIVERABLES_IN_PERIOD.

Use only evidence from the included emails. Every suggestion must cite one or more supportingMessageRefs from INCLUDED_EMAILS.
Each included email body is a reduced plain-text thread view. Quoted history, signatures, device footers, and disclaimer boilerplate may already be removed. Treat the body as the unique useful text for that reply when provided.
Prefer concise, standardized deliverable names when possible (for example: DD60, DD90, CD60, CD90, CD100, CDF, RFI, RFI #2, Submittal, Lighting Submittal, Controls Submittal, Record Set, Record Drawings, IFP, Site Survey, Survey Report, ASR, ASR #2, PCC, PCC #3, Bulletin #2, Coordination, Meeting, Revision).
Use the provided project identifiers when available. If a project must be inferred from the emails, return the best available id/name/path fields.
Focus only on the primary {disciplines_str} engineering tasks mentioned. Ignore work for other disciplines.
Use the current date {current_date} when interpreting due dates. If a due date is not clear, leave it empty.
{task_notes_contract}
{truncation_note}

Return ONLY valid JSON with this exact shape:
{{"suggestions":[{{"projectId":"","projectName":"","projectNick":"","projectPath":"","deliverable":"","due":"","tasks":[""],"notes":"","supportingMessageRefs":["E001"]}}]}}

If no additions are warranted, return:
{{"suggestions":[]}}

INCLUDED_EMAILS:
{email_json}

CURRENT_DELIVERABLES_IN_PERIOD:
{deliverable_json}
""".strip()

    def _normalize_outlook_scan_batch_suggestions(self, raw_payload):
        if isinstance(raw_payload, list):
            raw_suggestions = raw_payload
        elif isinstance(raw_payload, dict):
            if isinstance(raw_payload.get("suggestions"), list):
                raw_suggestions = raw_payload.get("suggestions") or []
            elif any(
                key in raw_payload
                for key in (
                    "deliverable",
                    "deliverableName",
                    "project",
                    "projectId",
                    "projectName",
                )
            ):
                raw_suggestions = [raw_payload]
            else:
                raise RuntimeError("AI returned invalid Outlook scan suggestions.")
        else:
            raise RuntimeError("AI returned invalid Outlook scan suggestions.")

        if not isinstance(raw_suggestions, list):
            raise RuntimeError("AI returned invalid Outlook scan suggestions.")
        if not raw_suggestions:
            return []

        suggestions = []
        for raw_suggestion in raw_suggestions:
            if not isinstance(raw_suggestion, dict):
                continue
            project = raw_suggestion.get("project")
            project = project if isinstance(project, dict) else {}
            deliverable_payload = raw_suggestion.get("deliverable")
            deliverable_payload = deliverable_payload if isinstance(deliverable_payload, dict) else {}

            project_info = {
                "id": str(project.get("id") or raw_suggestion.get("projectId") or "").strip(),
                "name": str(project.get("name") or raw_suggestion.get("projectName") or "").strip(),
                "nick": str(project.get("nick") or raw_suggestion.get("projectNick") or "").strip(),
                "path": str(project.get("path") or raw_suggestion.get("projectPath") or "").strip(),
            }
            deliverable_name = str(
                deliverable_payload.get("name")
                or raw_suggestion.get("deliverable")
                or raw_suggestion.get("deliverableName")
                or raw_suggestion.get("name")
                or ""
            ).strip()
            supporting_refs = (
                raw_suggestion.get("supportingMessageRefs")
                or raw_suggestion.get("supportingMessageIds")
                or raw_suggestion.get("messageRefs")
                or []
            )
            if isinstance(supporting_refs, str):
                supporting_refs = [supporting_refs]
            if not isinstance(supporting_refs, list):
                supporting_refs = []
            normalized_refs = []
            seen_refs = set()
            for raw_ref in supporting_refs:
                ref = str(raw_ref or "").strip()
                if not ref or ref in seen_refs:
                    continue
                seen_refs.add(ref)
                normalized_refs.append(ref)

            if not deliverable_name:
                continue
            if not any(project_info.values()):
                continue

            suggestions.append({
                "project": project_info,
                "deliverable": {
                    "name": deliverable_name,
                    "due": str(deliverable_payload.get("due") or raw_suggestion.get("due") or "").strip(),
                    "notes": str(deliverable_payload.get("notes") or raw_suggestion.get("notes") or "").strip(),
                    "tasks": self._normalize_outlook_scan_task_list(
                        deliverable_payload.get("tasks")
                        if isinstance(deliverable_payload.get("tasks"), list)
                        else raw_suggestion.get("tasks")
                    ),
                },
                "supportingMessageRefs": normalized_refs,
            })

        if suggestions:
            return suggestions
        raise RuntimeError("AI returned invalid Outlook scan suggestions.")

    def _request_outlook_scan_batch_suggestions(self, prompt, api_key):
        final_api_key = self._resolve_google_ai_api_key(api_key)
        try:
            self._ensure_aiohttp()
            client = genai.Client(
                api_key=final_api_key,
                http_options=types.HttpOptions(timeout=120000),
            )
            response = client.models.generate_content(
                model="gemini-3-flash-preview",
                contents=[
                    types.Content(
                        role="user",
                        parts=[types.Part.from_text(text=prompt)],
                    ),
                ],
                config=types.GenerateContentConfig(
                    temperature=0,
                    response_mime_type="application/json",
                ),
            )

            raw_text = response.text
            if raw_text is None:
                if hasattr(response, "parts") and response.parts:
                    raw_text = "".join(
                        part.text for part in response.parts
                        if hasattr(part, "text") and part.text
                    )
                if not raw_text:
                    logging.error(f"AI returned empty Outlook scan response. Full response: {response}")
                    raise RuntimeError("AI returned an empty Outlook scan response. Please try again.")

            cleaned = raw_text.strip()
            if not cleaned:
                raise RuntimeError("AI returned an empty Outlook scan response. Please try again.")

            logging.debug(f"Outlook batch AI response text: {cleaned[:500]}...")
            return self._normalize_outlook_scan_batch_suggestions(json.loads(cleaned))
        except json.JSONDecodeError as exc:
            logging.error(f"Failed to parse Outlook scan AI response as JSON: {exc}")
            raise RuntimeError("AI returned invalid Outlook scan suggestions JSON. Please try again.")
        except Exception as exc:
            msg = str(exc)
            lower = msg.lower()
            logging.error(f"Error processing Outlook scan batch with AI: {type(exc).__name__}: {exc}")
            if (
                "api key expired" in lower
                or "api_key_invalid" in lower
                or "invalid api key" in lower
            ):
                raise RuntimeError(
                    "Your Google API key is expired/invalid. "
                    "Create a new key in Google AI Studio, update your settings, then try again."
                )
            if "model" in lower and ("not found" in lower or "does not exist" in lower):
                raise RuntimeError(
                    "AI model not available. The Gemini 3 Flash model may not be accessible with your API key."
                )
            if "quota" in lower or "rate limit" in lower:
                raise RuntimeError("API rate limit exceeded. Please wait a moment and try again.")
            raise RuntimeError(f"AI error: {msg}")

    def _analyze_outlook_scan_batch(
        self,
        message_summaries,
        body_loader,
        project_context,
        api_key,
        user_name,
        discipline,
        timeframe,
        source,
        has_more=False,
        scan_date="",
    ):
        hydrated_entries = []
        skipped_messages = []
        relevant_email_count = 0
        total_emails = len(message_summaries) if isinstance(message_summaries, list) else 0
        deliverables_in_period = self._count_outlook_scan_project_context_deliverables(
            project_context
        )

        self._send_outlook_scan_progress(
            "hydrating",
            "Reading Outlook emails...",
            source=source,
            timeframe=timeframe,
            scanDate=scan_date,
            totalEmails=total_emails,
            processedEmails=0,
            skippedEmails=0,
            deliverablesInPeriod=deliverables_in_period,
        )

        for idx, summary in enumerate(message_summaries):
            score = self._outlook_message_candidate_score(summary)
            if score > 0:
                relevant_email_count += 1
            try:
                hydrated_summary, body_text = body_loader(summary)
                body_content = self._build_outlook_scan_analysis_body(
                    body_text or (hydrated_summary or {}).get("bodyPreview") or ""
                )
                if not body_content:
                    skipped_messages.append({
                        "message": hydrated_summary or summary,
                        "reason": "Message body was empty.",
                    })
                    continue
                received_dt = parse_utc_iso((hydrated_summary or {}).get("receivedDateTime"))
                received_sort = received_dt.timestamp() if received_dt else 0
                prompt_ref = f"E{idx + 1:03d}"
                hydrated_entries.append({
                    "promptRef": prompt_ref,
                    "score": score,
                    "receivedSort": received_sort,
                    "message": hydrated_summary or summary,
                    "analysisBody": body_content,
                })
            except Exception as exc:
                skipped_messages.append({
                    "message": summary,
                    "reason": str(exc),
                })
            processed_emails = idx + 1
            if processed_emails == 1 or processed_emails % 10 == 0 or processed_emails == total_emails:
                self._send_outlook_scan_progress(
                    "hydrating",
                    f"Processed {processed_emails} of {total_emails} emails.",
                    source=source,
                    timeframe=timeframe,
                    scanDate=scan_date,
                    totalEmails=total_emails,
                    processedEmails=processed_emails,
                    skippedEmails=len(skipped_messages),
                    deliverablesInPeriod=deliverables_in_period,
                    relevantEmails=relevant_email_count,
                )

        hydrated_entries, dedupe_skipped_messages, dedupe_stats = self._dedupe_outlook_scan_entries(
            hydrated_entries
        )
        skipped_messages.extend(dedupe_skipped_messages)

        prompt_config = self._build_outlook_scan_prompt_config()
        included_emails, included_deliverables, prompt_truncated, trimmed_messages = self._select_outlook_scan_prompt_context(
            hydrated_entries,
            project_context,
            timeframe=timeframe,
            scan_date=scan_date,
            body_limit=prompt_config["bodyLimit"],
            email_budget_chars=prompt_config["emailBudgetChars"],
            max_emails=prompt_config["maxEmails"],
            skip_reason=prompt_config["skipReason"],
        )
        skipped_messages.extend(trimmed_messages)
        skipped_messages = self._dedupe_outlook_scan_skipped_messages(skipped_messages)

        self._send_outlook_scan_progress(
            "preparing_ai",
            "Preparing the batched AI review...",
            source=source,
            timeframe=timeframe,
            scanDate=scan_date,
            totalEmails=total_emails,
            processedEmails=total_emails,
            includedEmails=len(included_emails),
            skippedEmails=len(skipped_messages),
            deliverablesInPeriod=deliverables_in_period,
            relevantEmails=relevant_email_count,
            threadsDetected=dedupe_stats["threadsDetected"],
            dedupedEmailCount=dedupe_stats["dedupedEmailCount"],
            dedupeSkippedEmailCount=dedupe_stats["dedupeSkippedEmailCount"],
        )

        batch_suggestions = []
        if included_emails:
            self._send_outlook_scan_progress(
                "reviewing_ai",
                "Reviewing the included emails with AI...",
                source=source,
                timeframe=timeframe,
                scanDate=scan_date,
                totalEmails=total_emails,
                processedEmails=total_emails,
                includedEmails=len(included_emails),
                skippedEmails=len(skipped_messages),
                deliverablesInPeriod=deliverables_in_period,
                relevantEmails=relevant_email_count,
                threadsDetected=dedupe_stats["threadsDetected"],
                dedupedEmailCount=dedupe_stats["dedupedEmailCount"],
                dedupeSkippedEmailCount=dedupe_stats["dedupeSkippedEmailCount"],
            )
            prompt = self._build_outlook_scan_batch_prompt(
                included_emails,
                included_deliverables,
                user_name,
                discipline,
                timeframe,
                scan_date=scan_date,
                prompt_truncated=prompt_truncated,
            )
            try:
                batch_suggestions = self._request_outlook_scan_batch_suggestions(prompt, api_key)
            except RuntimeError as exc:
                if not self._is_outlook_scan_timeout_error(exc):
                    raise
                retry_prompt_config = self._build_outlook_scan_prompt_config(reduced=True)
                included_emails, included_deliverables, retry_prompt_truncated, retry_trimmed_messages = self._select_outlook_scan_prompt_context(
                    hydrated_entries,
                    project_context,
                    timeframe=timeframe,
                    scan_date=scan_date,
                    body_limit=retry_prompt_config["bodyLimit"],
                    email_budget_chars=retry_prompt_config["emailBudgetChars"],
                    max_emails=retry_prompt_config["maxEmails"],
                    skip_reason=retry_prompt_config["skipReason"],
                )
                skipped_messages.extend(retry_trimmed_messages)
                skipped_messages = self._dedupe_outlook_scan_skipped_messages(skipped_messages)
                prompt_truncated = prompt_truncated or retry_prompt_truncated
                self._send_outlook_scan_progress(
                    "reviewing_ai",
                    "AI request timed out. Retrying with a reduced text-only batch...",
                    source=source,
                    timeframe=timeframe,
                    scanDate=scan_date,
                    totalEmails=total_emails,
                    processedEmails=total_emails,
                    includedEmails=len(included_emails),
                    skippedEmails=len(skipped_messages),
                    deliverablesInPeriod=deliverables_in_period,
                    relevantEmails=relevant_email_count,
                    threadsDetected=dedupe_stats["threadsDetected"],
                    dedupedEmailCount=dedupe_stats["dedupedEmailCount"],
                    dedupeSkippedEmailCount=dedupe_stats["dedupeSkippedEmailCount"],
                )
                retry_prompt = self._build_outlook_scan_batch_prompt(
                    included_emails,
                    included_deliverables,
                    user_name,
                    discipline,
                    timeframe,
                    scan_date=scan_date,
                    prompt_truncated=prompt_truncated,
                )
                try:
                    batch_suggestions = self._request_outlook_scan_batch_suggestions(
                        retry_prompt,
                        api_key,
                    )
                except RuntimeError as retry_exc:
                    if self._is_outlook_scan_timeout_error(retry_exc):
                        raise RuntimeError(
                            "AI timed out while reviewing Outlook emails, even after retrying with a reduced text-only batch. Try scanning a quieter day or narrowing the email set."
                        ) from retry_exc
                    raise

        self._send_outlook_scan_progress(
            "matching",
            "Matching scan results to your current projects...",
            source=source,
            timeframe=timeframe,
            scanDate=scan_date,
            totalEmails=total_emails,
            processedEmails=total_emails,
            includedEmails=len(included_emails),
            skippedEmails=len(skipped_messages),
            deliverablesInPeriod=deliverables_in_period,
            relevantEmails=relevant_email_count,
            threadsDetected=dedupe_stats["threadsDetected"],
            dedupedEmailCount=dedupe_stats["dedupedEmailCount"],
            dedupeSkippedEmailCount=dedupe_stats["dedupeSkippedEmailCount"],
        )

        messages_by_ref = {
            str(entry.get("promptRef") or ""): entry.get("message") or {}
            for entry in included_emails
        }
        enriched_suggestions = []
        for suggestion in batch_suggestions:
            related_messages = []
            supporting_message_ids = []
            seen_ids = set()
            for ref in suggestion.get("supportingMessageRefs") or []:
                message = messages_by_ref.get(str(ref or "").strip())
                if not isinstance(message, dict) or not message:
                    continue
                message_id = str(message.get("id") or "").strip()
                dedupe_key = message_id.lower() if message_id else str(ref or "").strip().lower()
                if dedupe_key in seen_ids:
                    continue
                seen_ids.add(dedupe_key)
                related_messages.append(message)
                if message_id:
                    supporting_message_ids.append(message_id)
            enriched_suggestions.append({
                **suggestion,
                "supportingMessageIds": supporting_message_ids,
                "relatedMessages": related_messages,
            })

        result = {
            "status": "success",
            "analysisMode": "batch",
            "candidateCount": len(included_emails),
            "relevantEmailCount": relevant_email_count,
            "scannedCount": total_emails,
            "emailsIncludedCount": len(included_emails),
            "deliverablesIncludedCount": len(included_deliverables),
            "suggestions": enriched_suggestions,
            "skippedMessages": skipped_messages,
            "promptTruncated": prompt_truncated,
            "truncated": bool(has_more) or prompt_truncated,
            "threadsDetected": dedupe_stats["threadsDetected"],
            "dedupedEmailCount": dedupe_stats["dedupedEmailCount"],
            "dedupeSkippedEmailCount": dedupe_stats["dedupeSkippedEmailCount"],
            "timeframe": timeframe,
            "scanDate": self._normalize_outlook_scan_date(scan_date),
            "source": source,
        }
        selected_label = self._format_outlook_scan_period_label(
            timeframe=timeframe,
            scan_date=scan_date,
        )
        if total_emails <= 0:
            completion_message = (
                "Scan complete. No emails were found for the selected day."
                if self._normalize_outlook_scan_date(scan_date)
                else "Scan complete. No emails were found for the selected timeframe."
            )
        elif enriched_suggestions:
            completion_message = (
                f"Scan complete. Found {len(enriched_suggestions)} deliverable "
                f"suggestion{'s' if len(enriched_suggestions) != 1 else ''}."
            )
        else:
            completion_message = "Scan complete. No new deliverables were suggested."
        if self._normalize_outlook_scan_date(scan_date):
            completion_message += f" Day: {selected_label}."
        if prompt_truncated:
            completion_message += " Prompt trimmed to keep the scan bounded."
        self._send_outlook_scan_progress(
            "done",
            completion_message,
            active=False,
            source=source,
            timeframe=timeframe,
            scanDate=scan_date,
            totalEmails=total_emails,
            processedEmails=total_emails,
            includedEmails=len(included_emails),
            skippedEmails=len(skipped_messages),
            deliverablesInPeriod=deliverables_in_period,
            relevantEmails=relevant_email_count,
            threadsDetected=dedupe_stats["threadsDetected"],
            dedupedEmailCount=dedupe_stats["dedupedEmailCount"],
            dedupeSkippedEmailCount=dedupe_stats["dedupeSkippedEmailCount"],
        )
        return result

    def scan_outlook_inbox(self, payload, api_key, user_name, discipline):
        timeframe = "week"
        scan_date = ""
        source = ""
        try:
            request_payload = payload or {}
            if isinstance(request_payload, str):
                request_payload = json.loads(request_payload)
            if not isinstance(request_payload, dict):
                request_payload = {}
            raw_scan_date = request_payload.get("scanDate")
            scan_date = self._normalize_outlook_scan_date(raw_scan_date)
            timeframe = str(request_payload.get("timeframe") or "week").strip().lower()
            if timeframe not in ("week", "month"):
                timeframe = "week"
            if not scan_date and (
                raw_scan_date is not None or "timeframe" not in request_payload
            ):
                scan_date = self._normalize_outlook_scan_date(
                    datetime.datetime.now().astimezone().date().isoformat()
                )
            project_context = self._normalize_outlook_scan_project_context(
                request_payload.get("projectContext")
            )
            deliverables_in_period = self._count_outlook_scan_project_context_deliverables(
                project_context
            )
            self._send_outlook_scan_progress(
                "starting",
                "Starting Outlook inbox scan...",
                timeframe=timeframe,
                scanDate=scan_date,
                deliverablesInPeriod=deliverables_in_period,
            )

            desktop_available, desktop_reason = self._get_desktop_outlook_availability()
            source = OUTLOOK_SCAN_SOURCE_DESKTOP
            if not desktop_available:
                message = desktop_reason or "Desktop Outlook is unavailable on this machine."
                self._send_outlook_scan_progress(
                    "error",
                    message,
                    active=False,
                    source=source,
                    timeframe=timeframe,
                    scanDate=scan_date,
                    deliverablesInPeriod=deliverables_in_period,
                )
                return self._build_outlook_scan_error_result(
                    message,
                    timeframe=timeframe,
                    source=source,
                    scan_date=scan_date,
                )

            self._send_outlook_scan_progress(
                "listing",
                "Loading emails from Desktop Outlook...",
                source=source,
                timeframe=timeframe,
                scanDate=scan_date,
                deliverablesInPeriod=deliverables_in_period,
            )
            try:
                message_summaries, has_more = self._list_desktop_outlook_inbox_messages(
                    timeframe=timeframe,
                    limit=OUTLOOK_SCAN_FETCH_LIMIT,
                    scan_date=scan_date,
                )
            except Exception as exc:
                logging.error(f"Desktop Outlook inbox scan failed: {exc}")
                message = f"Desktop Outlook could not be read: {exc}"
                self._send_outlook_scan_progress(
                    "error",
                    message,
                    active=False,
                    source=source,
                    timeframe=timeframe,
                    scanDate=scan_date,
                    deliverablesInPeriod=deliverables_in_period,
                )
                return self._build_outlook_scan_error_result(
                    message,
                    timeframe=timeframe,
                    source=source,
                    scan_date=scan_date,
                )

            return self._analyze_outlook_scan_batch(
                message_summaries,
                lambda summary: self._get_desktop_outlook_message_body_text(summary.get("id")),
                project_context,
                api_key,
                user_name,
                discipline,
                timeframe,
                source,
                has_more=has_more,
                scan_date=scan_date,
            )
        except json.JSONDecodeError:
            message = "Invalid Outlook scan payload."
            self._send_outlook_scan_progress(
                "error",
                message,
                active=False,
                source=source,
                timeframe=timeframe,
                scanDate=scan_date,
            )
            return self._build_outlook_scan_error_result(
                message,
                timeframe=timeframe,
                scan_date=scan_date,
            )
        except Exception as e:
            logging.error(f"Error scanning Outlook inbox: {e}")
            message = str(e)
            self._send_outlook_scan_progress(
                "error",
                message,
                active=False,
                source=source,
                timeframe=timeframe,
                scanDate=scan_date,
            )
            return self._build_outlook_scan_error_result(
                message,
                timeframe=timeframe,
                source=source,
                scan_date=scan_date,
            )

    def get_installed_autocad_versions(self):
        """Scans for installed AutoCAD versions in the typical directory."""
        versions = []
        base_dir = r"C:\Program Files\Autodesk"
        if not os.path.exists(base_dir):
            return {'status': 'success', 'versions': []}

        for year in range(2022, 2027):  # 2022 to 2026
            folder_name = f"AutoCAD {year}"
            folder_path = os.path.join(base_dir, folder_name)
            exe_path = os.path.join(folder_path, "accoreconsole.exe")
            if os.path.exists(exe_path):
                versions.append({'year': year, 'path': exe_path})

        return {'status': 'success', 'versions': versions}

    def _ensure_aiohttp(self):
        """
        Validate that aiohttp is available with ClientSession. This prevents
        cryptic AttributeErrors during GenAI calls when the dependency is
        missing or only partially installed.
        """
        try:
            import aiohttp  # type: ignore
        except Exception as exc:
            raise RuntimeError(
                "Missing dependency: aiohttp. Install/upgrade it (e.g. "
                "pip install \"aiohttp>=3.13\") and rebuild the app."
            ) from exc

        if not hasattr(aiohttp, "ClientSession"):
            raise RuntimeError(
                "aiohttp is installed but ClientSession is unavailable. "
                "Reinstall it with `pip install --force-reinstall aiohttp>=3.13`."
            )

    def get_chat_response(self, chat_history):
        settings = self.get_user_settings()
        api_key = settings.get('apiKey', '').strip()

        if not api_key:
            api_key = (os.getenv("GOOGLE_API_KEY") or "").strip()

        if not api_key:
            return {
                'status': 'error',
                'response': 'API key not found. Please configure it in the AI settings.'
            }

        try:
            # Validate input
            if not isinstance(chat_history, list):
                return {
                    'status': 'error',
                    'response': 'Invalid chat history format.'
                }

            if len(chat_history) == 0:
                return {
                    'status': 'error',
                'response': 'Chat history is empty.'
            }

            self._ensure_aiohttp()
            client = genai.Client(api_key=api_key)
            model = "gemini-3-flash-preview"

            # Convert JavaScript chat history to Gemini API format
            contents = []
            for msg in chat_history:
                role = msg.get('role', 'user')
                # Map 'model' role to 'assistant' if needed
                if role == 'model':
                    role = 'model'  # Gemini uses 'model' for assistant responses

                parts = msg.get('parts', [])
                if not parts:
                    continue

                # Extract text from parts
                text_parts = []
                for part in parts:
                    if isinstance(part, dict) and 'text' in part:
                        text_parts.append(
                            types.Part.from_text(text=part['text']))
                    elif isinstance(part, str):
                        text_parts.append(types.Part.from_text(text=part))

                if text_parts:
                    contents.append(types.Content(role=role, parts=text_parts))

            if not contents:
                return {
                    'status': 'error',
                    'response': 'No valid messages found in chat history.'
                }

            tools = [types.Tool(google_search=types.GoogleSearch())]

            generate_content_config = types.GenerateContentConfig(
                temperature=0,
                top_p=0.7,
                thinking_config=types.ThinkingConfig(
                    thinking_budget=-1,
                ),
                tools=tools,
            )

            response = client.models.generate_content(
                model=model,
                contents=contents,
                config=generate_content_config,
            )

            return {'status': 'success', 'response': response.text}

        except Exception as e:
            msg = str(e)
            lower = msg.lower()
            if ("api key expired" in lower or
                "api_key_invalid" in lower or
                    "invalid api key" in lower):
                return {'status': 'error',
                        'response': ('Your Google API key is expired/invalid. '
                                     'Create a new key in Google AI Studio, '
                                     'update your settings, then try again.')}
            logging.error(f"Error getting chat response from AI: {e}")
            return {'status': 'error', 'response': f"An AI error occurred: {msg}"}

    def _process_email_with_ai_legacy(self, email_text, api_key, user_name, discipline):
        """
        Processes email text using Google GenAI to extract project details.
        """
        final_api_key = (api_key or "").strip()
        if not final_api_key:
            final_api_key = (os.getenv("GOOGLE_API_KEY") or "").strip()

        if not final_api_key:
            return {
                'status': 'error',
                'message': 'AI API key is not configured. Please provide it in the app settings or set GOOGLE_API_KEY in your .env file.'
            }

        current_date = datetime.date.today().strftime("%m/%d/%Y")
        disciplines_str = ', '.join(discipline) if isinstance(
            discipline, list) else (discipline or 'Engineering')
        task_notes_contract = self._build_deliverable_task_notes_prompt_contract(
            disciplines_str,
            include_json_field_names=False,
        )
        prompt = f"""
You are an intelligent assistant for {user_name}, a(n) {disciplines_str} engineering project manager. Your task is to analyze an email and extract specific project details. Focus ONLY on the primary {disciplines_str} engineering tasks mentioned. Ignore tasks for other disciplines.
Analyze the following email text and extract the information into a valid JSON object with the following keys: "id", "name", "due", "path", "deliverable", "tasks", "notes".
- "id": Find a project number or project ID (e.g., "250597", "P-12345", "Job #1042"). Look in the subject line, headers, and body. This could be called a job number, project number, project ID, or similar. If none, leave it empty.
- "name": Determine the project name, typically including the client and address or building name (e.g., "BofA, 22004 Sherman Way, Canoga Park, CA"). Include enough detail to uniquely identify the project. If no formal name is found, compose one from the client name and location mentioned.
- "due": Find the due date and format it as "MM/DD/YY". The current date is {current_date}. If the year is not specified in the email, assume the current year or the next year if the date would be in the past. Ensure the due date is on or after today. If multiple dates, choose the most relevant upcoming one.
- "path": Find the main project file path (e.g., "M:\\\\Gensler\\\\...").
- "deliverable": Infer the deliverable name from the email if possible. Prefer concise, standardized names when present (for example: DD60, DD90, CD60, CD90, CD100, CDF, RFI, RFI #2, Submittal, Lighting Submittal, Controls Submittal, Record Set, Record Drawings, IFP, Site Survey, Survey Report, ASR, ASR #2, PCC, PCC #3, Bulletin #2, Coordination, Meeting, Revision). If no deliverable is clear, leave it empty.
{task_notes_contract}
If a piece of information is not found, the value should be an empty string "" for strings, or an empty array [] for tasks.
Here is the email:
---
{email_text}
---
Return ONLY the JSON object.
""".strip()
        try:
            self._ensure_aiohttp()
            client = genai.Client(
                api_key=final_api_key,
                http_options=types.HttpOptions(timeout=120000),
            )
            model = "gemini-3-flash-preview"

            contents = [
                types.Content(
                    role="user",
                    parts=[types.Part.from_text(text=prompt)],
                ),
            ]

            generate_content_config = types.GenerateContentConfig(
                temperature=0,
                response_mime_type="application/json",
            )

            response = client.models.generate_content(
                model=model,
                contents=contents,
                config=generate_content_config,
            )

            # Log response for debugging
            logging.debug(f"AI response object: {response}")

            # Handle case where response.text might be None or empty
            raw_text = response.text
            if raw_text is None:
                # Try accessing via parts if .text is None
                if hasattr(response, 'parts') and response.parts:
                    raw_text = ''.join(
                        part.text for part in response.parts
                        if hasattr(part, 'text') and part.text
                    )
                if not raw_text:
                    logging.error(f"AI returned empty response. Full response: {response}")
                    return {'status': 'error', 'message': 'AI returned an empty response. Please try again.'}

            cleaned = raw_text.strip()
            if not cleaned:
                return {'status': 'error', 'message': 'AI returned an empty response. Please try again.'}

            logging.debug(f"AI response text: {cleaned[:500]}...")
            project_data = json.loads(cleaned)

            # Handle case where AI returns a list instead of a dict
            if isinstance(project_data, list):
                project_data = project_data[0] if project_data else {}

            project_data.setdefault("id", "")
            project_data.setdefault("name", "")
            project_data.setdefault("due", "")
            project_data.setdefault("path", "")
            project_data.setdefault("deliverable", "")
            project_data.setdefault("tasks", [])
            project_data.setdefault("notes", "")

            if 'tasks' in project_data and isinstance(project_data['tasks'], list):
                project_data['tasks'] = [{'text': str(task), 'done': False, 'links': [
                ]} for task in project_data['tasks']]

            return {'status': 'success', 'data': project_data}

        except json.JSONDecodeError as e:
            logging.error(f"Failed to parse AI response as JSON: {e}")
            return {'status': 'error', 'message': 'AI returned invalid JSON. Please try again.'}
        except Exception as e:
            msg = str(e)
            lower = msg.lower()
            logging.error(f"Error processing email with AI: {type(e).__name__}: {e}")
            if ("api key expired" in lower or
                "api_key_invalid" in lower or
                    "invalid api key" in lower):
                return {'status': 'error',
                        'message': ('Your Google API key is expired/invalid. '
                                    'Create a new key in Google AI Studio → API keys, '
                                    'update your settings, then try again.')}
            if "model" in lower and ("not found" in lower or "does not exist" in lower):
                return {'status': 'error',
                        'message': 'AI model not available. The Gemini 3 Flash model may not be accessible with your API key.'}
            if "quota" in lower or "rate limit" in lower:
                return {'status': 'error',
                        'message': 'API rate limit exceeded. Please wait a moment and try again.'}
            return {'status': 'error', 'message': f"AI error: {msg}"}

    def _resolve_google_ai_api_key(self, api_key):
        final_api_key = (api_key or "").strip()
        if not final_api_key:
            final_api_key = (os.getenv("GOOGLE_API_KEY") or "").strip()
        if not final_api_key:
            raise RuntimeError(
                "AI API key is not configured. Please provide it in the app settings or set GOOGLE_API_KEY in your .env file."
            )
        return final_api_key

    def _build_email_analysis_prompt(self, email_text, user_name, discipline):
        current_date = datetime.date.today().strftime("%m/%d/%Y")
        disciplines_str = ', '.join(discipline) if isinstance(
            discipline, list) else (discipline or 'Engineering')
        task_notes_contract = self._build_deliverable_task_notes_prompt_contract(
            disciplines_str,
            include_json_field_names=False,
        )
        return f"""
You are an intelligent assistant for {user_name}, a(n) {disciplines_str} engineering project manager. Your task is to analyze an email and extract specific project details. Focus ONLY on the primary {disciplines_str} engineering tasks mentioned. Ignore tasks for other disciplines.
Analyze the following email text and extract the information into a valid JSON object with the following keys: "id", "name", "due", "path", "deliverable", "tasks", "notes".
- "id": Find a project number or project ID (e.g., "250597", "P-12345", "Job #1042"). Look in the subject line, headers, and body. This could be called a job number, project number, project ID, or similar. If none, leave it empty.
- "name": Determine the project name, typically including the client and address or building name (e.g., "BofA, 22004 Sherman Way, Canoga Park, CA"). Include enough detail to uniquely identify the project. If no formal name is found, compose one from the client name and location mentioned.
- "due": Find the due date and format it as "MM/DD/YY". The current date is {current_date}. If the year is not specified in the email, assume the current year or the next year if the date would be in the past. Ensure the due date is on or after today. If multiple dates, choose the most relevant upcoming one.
- "path": Find the main project file path (e.g., "M:\\\\Gensler\\\\...").
- "deliverable": Infer the deliverable name from the email if possible. Prefer concise, standardized names when present (for example: DD60, DD90, CD60, CD90, CD100, CDF, RFI, RFI #2, Submittal, Lighting Submittal, Controls Submittal, Record Set, Record Drawings, IFP, Site Survey, Survey Report, ASR, ASR #2, PCC, PCC #3, Bulletin #2, Coordination, Meeting, Revision). If no deliverable is clear, leave it empty.
{task_notes_contract}
If a piece of information is not found, the value should be an empty string "" for strings, or an empty array [] for tasks.
Here is the email:
---
{email_text}
---
Return ONLY the JSON object.
""".strip()

    def _normalize_email_project_data(self, project_data):
        if isinstance(project_data, list):
            project_data = project_data[0] if project_data else {}
        if not isinstance(project_data, dict):
            project_data = {}
        project_data.setdefault("id", "")
        project_data.setdefault("name", "")
        project_data.setdefault("due", "")
        project_data.setdefault("path", "")
        project_data.setdefault("deliverable", "")
        project_data.setdefault("tasks", [])
        project_data.setdefault("notes", "")
        normalized_tasks = []
        if isinstance(project_data.get("tasks"), list):
            for task in project_data["tasks"]:
                if isinstance(task, dict):
                    text = str(task.get("text") or "").strip()
                    done = bool(task.get("done"))
                    links = task.get("links") if isinstance(task.get("links"), list) else []
                else:
                    text = str(task).strip()
                    done = False
                    links = []
                if not text:
                    continue
                normalized_tasks.append({
                    "text": text,
                    "done": done,
                    "links": links,
                })
        project_data["tasks"] = normalized_tasks
        return project_data

    def _build_deliverable_task_notes_prompt_contract(
        self,
        disciplines_str,
        include_json_field_names=True,
    ):
        tasks_label = '"tasks"' if include_json_field_names else 'tasks'
        notes_label = '"notes"' if include_json_field_names else 'notes'
        return "\n".join([
            f'- {tasks_label}: Create a JSON array of strings listing every actionable {disciplines_str} engineering step needed to complete the deliverable. Include revisions, submissions, coordination, and follow-up work. Be concise. Example: ["Revise lighting plans per architect comments", "Update panel schedules", "Issue updated PDF set for resubmittal"].',
            f'- {notes_label}: Provide only non-actionable context the user should note, such as constraints, approvals, dependencies, reminders, delivery method, or special instructions. Do NOT put action items, next steps, or completion work in {notes_label}. Example: "Awaiting architect approval before final issue." If there is no notable non-actionable context, return "".',
        ])

    def _extract_project_data_from_email_text(self, email_text, api_key, user_name, discipline):
        final_api_key = self._resolve_google_ai_api_key(api_key)
        prompt = self._build_email_analysis_prompt(email_text, user_name, discipline)
        try:
            self._ensure_aiohttp()
            client = genai.Client(
                api_key=final_api_key,
                http_options=types.HttpOptions(timeout=120000),
            )
            response = client.models.generate_content(
                model="gemini-3-flash-preview",
                contents=[
                    types.Content(
                        role="user",
                        parts=[types.Part.from_text(text=prompt)],
                    ),
                ],
                config=types.GenerateContentConfig(
                    temperature=0,
                    response_mime_type="application/json",
                ),
            )

            logging.debug(f"AI response object: {response}")
            raw_text = response.text
            if raw_text is None:
                if hasattr(response, 'parts') and response.parts:
                    raw_text = ''.join(
                        part.text for part in response.parts
                        if hasattr(part, 'text') and part.text
                    )
                if not raw_text:
                    logging.error(f"AI returned empty response. Full response: {response}")
                    raise RuntimeError('AI returned an empty response. Please try again.')

            cleaned = raw_text.strip()
            if not cleaned:
                raise RuntimeError('AI returned an empty response. Please try again.')

            logging.debug(f"AI response text: {cleaned[:500]}...")
            return self._normalize_email_project_data(json.loads(cleaned))
        except json.JSONDecodeError as e:
            logging.error(f"Failed to parse AI response as JSON: {e}")
            raise RuntimeError('AI returned invalid JSON. Please try again.')
        except Exception as e:
            msg = str(e)
            lower = msg.lower()
            logging.error(f"Error processing email with AI: {type(e).__name__}: {e}")
            if ("api key expired" in lower or
                "api_key_invalid" in lower or
                    "invalid api key" in lower):
                raise RuntimeError(
                    'Your Google API key is expired/invalid. '
                    'Create a new key in Google AI Studio, '
                    'update your settings, then try again.'
                )
            if "model" in lower and ("not found" in lower or "does not exist" in lower):
                raise RuntimeError(
                    'AI model not available. The Gemini 3 Flash model may not be accessible with your API key.'
                )
            if "quota" in lower or "rate limit" in lower:
                raise RuntimeError('API rate limit exceeded. Please wait a moment and try again.')
            raise RuntimeError(f"AI error: {msg}")

    def process_email_with_ai(self, email_text, api_key, user_name, discipline):
        """
        Processes email text using Google GenAI to extract project details.
        """
        try:
            project_data = self._extract_project_data_from_email_text(
                email_text,
                api_key,
                user_name,
                discipline,
            )
            return {'status': 'success', 'data': project_data}
        except Exception as e:
            return {'status': 'error', 'message': str(e)}

    def get_tasks(self):
        """Reads and returns the content of tasks.json."""
        try:
            with open(TASKS_FILE, 'r', encoding='utf-8') as f:
                payload = json.load(f)
                return _overlay_projects_with_lighting_schedule_records(payload)
        except (FileNotFoundError, json.JSONDecodeError):
            return []

    def save_tasks(self, data):
        """Saves data to tasks.json and creates a backup."""
        try:
            if os.path.exists(TASKS_FILE):
                shutil.copy2(TASKS_FILE, TASKS_FILE + '.bak')
            with open(TASKS_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return {'status': 'success'}
        except Exception as e:
            logging.error(f"Error saving tasks: {e}")
            return {'status': 'error', 'message': str(e)}

    def get_lighting_schedule_sync(self, dwg_path):
        """Read project-specific light fixture schedule sync JSON near the DWG."""
        try:
            _, sync_path = _resolve_lighting_schedule_sync_path(dwg_path)
            exists = os.path.exists(sync_path)
            modified = None
            payload = None
            if exists:
                payload = _read_json_file_strict(sync_path)
                mtime = os.path.getmtime(sync_path)
                modified = datetime.datetime.fromtimestamp(mtime).isoformat()

            result = {
                'status': 'success',
                'exists': exists,
                'path': sync_path,
                'modified': modified,
            }
            if payload is not None:
                result['data'] = payload
            return result
        except ValueError as e:
            logging.warning(f"Lighting schedule sync read validation failed: {e}")
            return {'status': 'error', 'message': str(e)}
        except FileNotFoundError:
            return {'status': 'success', 'exists': False, 'path': '', 'modified': None}
        except Exception as e:
            logging.error(f"Error reading lighting schedule sync: {e}")
            return {'status': 'error', 'message': str(e)}

    def save_lighting_schedule_sync(self, dwg_path, payload):
        """Write project-specific light fixture schedule sync JSON near the DWG."""
        try:
            if not isinstance(payload, dict):
                raise ValueError("Sync payload must be a JSON object.")

            _, sync_path = _resolve_lighting_schedule_sync_path(dwg_path)
            sync_dir = os.path.dirname(sync_path)
            if not os.path.isdir(sync_dir):
                raise ValueError(f"DWG folder does not exist: {sync_dir}")

            try:
                json.dumps(payload)
            except (TypeError, ValueError) as exc:
                raise ValueError(f"Sync payload is not valid JSON: {exc}")

            _atomic_write_json_file(sync_path, payload)
            mtime = os.path.getmtime(sync_path)
            modified = datetime.datetime.fromtimestamp(mtime).isoformat()
            return {
                'status': 'success',
                'path': sync_path,
                'modified': modified,
            }
        except ValueError as e:
            logging.warning(f"Lighting schedule sync save validation failed: {e}")
            return {'status': 'error', 'message': str(e)}
        except Exception as e:
            logging.error(f"Error saving lighting schedule sync: {e}")
            return {'status': 'error', 'message': str(e)}

    def get_lighting_schedule_record(self, project_id):
        """Load the canonical lighting schedule record from SQLite."""
        try:
            record = _get_lighting_schedule_record(project_id)
            return {
                'status': 'success',
                'exists': bool(record),
                'projectId': _resolve_lighting_schedule_project_id(project_id),
                'data': record,
            }
        except Exception as e:
            logging.error(f"Error loading lighting schedule record: {e}")
            return {'status': 'error', 'message': str(e)}

    def save_lighting_schedule_record(self, project_id, payload):
        """Persist the canonical lighting schedule record to SQLite."""
        try:
            safe_payload = payload if isinstance(payload, dict) else {}
            record = _save_lighting_schedule_record(
                project_id,
                safe_payload,
                updated_by=safe_payload.get('updatedBy', 'desktop'),
            )
            return {
                'status': 'success',
                'projectId': record['projectId'],
                'data': record,
            }
        except Exception as e:
            logging.error(f"Error saving lighting schedule record: {e}")
            return {'status': 'error', 'message': str(e)}

    def get_lighting_schedule_version(self, project_id):
        """Return only the current version metadata for polling."""
        try:
            return {
                'status': 'success',
                'data': _get_lighting_schedule_version(project_id),
            }
        except Exception as e:
            logging.error(f"Error loading lighting schedule version: {e}")
            return {'status': 'error', 'message': str(e)}

    def get_lighting_schedule_links(self, project_id):
        """Return linked DWG/table records for a project."""
        try:
            return {
                'status': 'success',
                'projectId': _resolve_lighting_schedule_project_id(project_id),
                'data': _get_lighting_schedule_links(project_id),
            }
        except Exception as e:
            logging.error(f"Error loading lighting schedule links: {e}")
            return {'status': 'error', 'message': str(e)}

    def read_t24_output_json(self, json_path):
        """Read and validate TXTSUMEXPORT output (T24Output.json)."""
        try:
            normalized_path = _normalize_t24_output_json_path(json_path)
            if not os.path.isfile(normalized_path):
                raise FileNotFoundError(
                    f"T24Output.json was not found at: {normalized_path}")

            payload = _read_json_file_strict(normalized_path)
            if not isinstance(payload, list):
                raise ValueError(
                    "T24Output.json must contain a JSON array of room entries.")

            rows = []
            for idx, item in enumerate(payload, start=1):
                if not isinstance(item, dict):
                    raise ValueError(
                        f"Entry #{idx} must be a JSON object with RoomType and SquareFeet.")

                room_type = str(item.get("RoomType", "")).strip()
                if not room_type:
                    raise ValueError(f"Entry #{idx} is missing RoomType.")

                square_feet_raw = item.get("SquareFeet", None)
                try:
                    square_feet = float(square_feet_raw)
                except (TypeError, ValueError):
                    raise ValueError(
                        f"Entry #{idx} has invalid SquareFeet value: {square_feet_raw!r}.")

                if not math.isfinite(square_feet) or square_feet < 0:
                    raise ValueError(
                        f"Entry #{idx} has non-finite or negative SquareFeet: {square_feet_raw!r}.")

                rows.append({
                    "roomType": room_type,
                    "squareFeet": round(square_feet, 4)
                })

            total_square_feet = round(
                sum(row["squareFeet"] for row in rows), 4)
            mtime = os.path.getmtime(normalized_path)
            modified = datetime.datetime.fromtimestamp(mtime).isoformat()

            return {
                "status": "success",
                "path": normalized_path,
                "modified": modified,
                "data": {
                    "rows": rows,
                    "totalSquareFeet": total_square_feet
                }
            }
        except FileNotFoundError as e:
            return {"status": "error", "message": str(e)}
        except ValueError as e:
            logging.warning(f"T24Output.json validation failed: {e}")
            return {"status": "error", "message": str(e)}
        except Exception as e:
            logging.error(f"Error reading T24Output.json: {e}")
            return {"status": "error", "message": str(e)}

    def get_notes(self):
        """Reads and returns the content of notes.json."""
        try:
            with open(NOTES_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            return {}

    def save_notes(self, data):
        """Saves notes data to notes.json and creates a backup."""
        try:
            if os.path.exists(NOTES_FILE):
                shutil.copy2(NOTES_FILE, NOTES_FILE + '.bak')
            with open(NOTES_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return {'status': 'success'}
        except Exception as e:
            logging.error(f"Error saving notes: {e}")
            return {'status': 'error', 'message': str(e)}

    def get_timesheets(self):
        """Reads and returns the content of timesheets.json."""
        try:
            with open(TIMESHEETS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            return {'weeks': {}, 'lastModified': None}

    def save_timesheets(self, data):
        """Saves timesheets data to timesheets.json and creates a backup."""
        try:
            if os.path.exists(TIMESHEETS_FILE):
                shutil.copy2(TIMESHEETS_FILE, TIMESHEETS_FILE + '.bak')
            with open(TIMESHEETS_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return {'status': 'success'}
        except Exception as e:
            logging.error(f"Error saving timesheets: {e}")
            return {'status': 'error', 'message': str(e)}

    # ===================== TEMPLATES API =====================

    def get_templates(self):
        """Reads and returns templates data, installing defaults on first run."""
        try:
            with open(TEMPLATES_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)

            # Check if defaults need installation
            if not data.get('defaultTemplatesInstalled'):
                data = self._install_default_templates(data)

            data = self._ensure_default_templates(data)
            return data
        except (FileNotFoundError, json.JSONDecodeError):
            # First run - create initial structure with defaults
            data = {'templates': [], 'defaultTemplatesInstalled': False, 'lastModified': None}
            data = self._install_default_templates(data)
            data = self._ensure_default_templates(data)
            return data

    def _install_default_templates(self, data):
        """Install default templates on first run."""
        for tpl_def in DEFAULT_TEMPLATES:
            if os.path.exists(tpl_def['sourcePath']):
                template = {
                    'id': f"tpl_{os.urandom(6).hex()}",
                    'name': tpl_def['name'],
                    'discipline': tpl_def['discipline'],
                    'fileType': tpl_def['fileType'],
                    'sourcePath': tpl_def['sourcePath'],
                    'isDefault': True,
                    'dateAdded': datetime.datetime.now().isoformat(),
                    'description': tpl_def.get('description', '')
                }
                data['templates'].append(template)

        data['defaultTemplatesInstalled'] = True
        data['lastModified'] = datetime.datetime.now().isoformat()
        self.save_templates(data)
        return data

    def _ensure_default_templates(self, data):
        """Ensure bundled default templates exist and update paths if needed."""
        templates = data.get('templates') or []
        updated = False
        for tpl_def in DEFAULT_TEMPLATES:
            name_key = str(tpl_def.get('name', '')).strip().lower()
            expected_path = tpl_def.get('sourcePath')
            existing = next(
                (
                    t for t in templates
                    if t.get('isDefault') and str(t.get('name', '')).strip().lower() == name_key
                ),
                None
            )

            if existing:
                existing_path = str(existing.get('sourcePath') or '')
                expected_abs = os.path.abspath(expected_path) if expected_path else ''
                existing_abs = os.path.abspath(existing_path) if existing_path else ''
                if expected_abs and (not os.path.exists(existing_path) or existing_abs != expected_abs):
                    existing['sourcePath'] = expected_path
                    existing['fileType'] = tpl_def.get('fileType')
                    existing['discipline'] = tpl_def.get('discipline')
                    existing['description'] = tpl_def.get('description', '')
                    updated = True
                continue

            if expected_path and os.path.exists(expected_path):
                template = {
                    'id': f"tpl_{os.urandom(6).hex()}",
                    'name': tpl_def['name'],
                    'discipline': tpl_def['discipline'],
                    'fileType': tpl_def['fileType'],
                    'sourcePath': expected_path,
                    'isDefault': True,
                    'dateAdded': datetime.datetime.now().isoformat(),
                    'description': tpl_def.get('description', '')
                }
                templates.append(template)
                updated = True

        if updated:
            data['templates'] = templates
            data['defaultTemplatesInstalled'] = True
            data['lastModified'] = datetime.datetime.now().isoformat()
            self.save_templates(data)
        return data

    def save_templates(self, data):
        """Saves templates data with backup."""
        try:
            if os.path.exists(TEMPLATES_FILE):
                shutil.copy2(TEMPLATES_FILE, TEMPLATES_FILE + '.bak')
            with open(TEMPLATES_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return {'status': 'success'}
        except Exception as e:
            logging.error(f"Error saving templates: {e}")
            return {'status': 'error', 'message': str(e)}

    def _normalize_template_key(self, template_key):
        raw = str(template_key or '').strip().lower()
        compact = re.sub(r'[^a-z0-9]+', '', raw)
        alias_map = {
            'narrative': 'narrative',
            'narrativeofchanges': 'narrative',
            'plancheck': 'planCheck',
            'plancheckcomments': 'planCheck',
            'plancheckresponseletter': 'planCheck',
            'pcc': 'planCheck',
        }
        if raw in TEMPLATE_KEY_BY_NAME:
            return TEMPLATE_KEY_BY_NAME[raw]
        if raw in alias_map:
            return alias_map[raw]
        return alias_map.get(compact, '')

    def _infer_template_key(self, template):
        if not isinstance(template, dict):
            return ''
        name_key = str(template.get('name') or '').strip().lower()
        if name_key in TEMPLATE_KEY_BY_NAME:
            return TEMPLATE_KEY_BY_NAME[name_key]
        source_name = os.path.basename(str(template.get('sourcePath') or '')).lower()
        if 'narrative of changes' in source_name:
            return 'narrative'
        if 'plan check' in source_name or 'pcc' in source_name:
            return 'planCheck'
        return ''

    def _find_template_by_key(self, template_key):
        normalized_key = self._normalize_template_key(template_key)
        if not normalized_key:
            return None, ''
        templates_data = self.get_templates()
        templates = templates_data.get('templates', []) if isinstance(
            templates_data, dict) else []
        for template in templates:
            if self._infer_template_key(template) == normalized_key:
                return template, normalized_key
        return None, normalized_key

    def _sanitize_template_filename_stem(self, value, fallback):
        text = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(value or '').strip())
        text = re.sub(r'\s+', ' ', text).strip().strip('.')
        text = text[:120]
        return text or fallback

    def _get_template_output_name_parts(self, template, template_key='', context=None, new_name=None):
        normalized_key = self._normalize_template_key(
            template_key or self._infer_template_key(template)
        )
        context_data = context if isinstance(context, dict) else {}
        source_path = str(template.get('sourcePath') or '').strip()
        extension = os.path.splitext(source_path)[1]
        if not extension:
            ext_hint = str(template.get('fileType') or '').strip().lstrip('.')
            extension = f'.{ext_hint}' if ext_hint else ''

        if new_name is not None and str(new_name).strip():
            base_name = str(new_name).strip()
            if extension and base_name.lower().endswith(extension.lower()):
                base_name = base_name[:-len(extension)]
            return base_name, extension, normalized_key

        if normalized_key in TEMPLATE_DEFAULT_FILENAME_BY_KEY:
            return TEMPLATE_DEFAULT_FILENAME_BY_KEY[normalized_key], extension, normalized_key

        fallback_name = (
            str(template.get('name') or '').strip()
            or Path(source_path).stem
            or 'Template'
        )
        return fallback_name, extension, normalized_key

    def _resolve_template_destination_path(self, folder_path, base_name, extension, conflict_policy='timestamp'):
        policy = str(conflict_policy or 'timestamp').strip().lower()
        safe_base = self._sanitize_template_filename_stem(base_name, 'Template')
        ext = str(extension or '').strip()
        if ext and not ext.startswith('.'):
            ext = f'.{ext}'
        candidate_path = os.path.join(folder_path, f'{safe_base}{ext}')
        if not os.path.exists(candidate_path):
            return candidate_path, os.path.basename(candidate_path), False
        if policy == 'overwrite':
            return candidate_path, os.path.basename(candidate_path), False
        if policy != 'timestamp':
            return candidate_path, os.path.basename(candidate_path), True

        timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H%M')
        suffix_index = 0
        while True:
            suffix = f' - {timestamp}' if suffix_index == 0 else f' - {timestamp} ({suffix_index + 1})'
            candidate_name = self._sanitize_template_filename_stem(
                f'{safe_base}{suffix}', safe_base)
            candidate_path = os.path.join(folder_path, f'{candidate_name}{ext}')
            if not os.path.exists(candidate_path):
                return candidate_path, os.path.basename(candidate_path), False
            suffix_index += 1

    def add_template(self, name, discipline, source_path, description=''):
        """Adds a new template to the collection."""
        try:
            if not os.path.exists(source_path):
                return {'status': 'error', 'message': 'Source file does not exist'}

            # Detect file type
            ext = os.path.splitext(source_path)[1].lower().lstrip('.')
            if ext not in ['doc', 'docx', 'dwg', 'xlsx', 'xls']:
                return {'status': 'error', 'message': f'Unsupported file type: .{ext}'}

            data = self.get_templates()
            template = {
                'id': f"tpl_{os.urandom(6).hex()}",
                'name': name,
                'discipline': discipline,
                'fileType': ext,
                'sourcePath': source_path,
                'isDefault': False,
                'dateAdded': datetime.datetime.now().isoformat(),
                'description': description
            }
            data['templates'].append(template)
            data['lastModified'] = datetime.datetime.now().isoformat()
            self.save_templates(data)
            return {'status': 'success', 'template': template}
        except Exception as e:
            logging.error(f"Error adding template: {e}")
            return {'status': 'error', 'message': str(e)}

    def remove_template(self, template_id):
        """Removes a template by ID."""
        try:
            data = self.get_templates()
            original_count = len(data['templates'])
            data['templates'] = [t for t in data['templates'] if t['id'] != template_id]

            if len(data['templates']) == original_count:
                return {'status': 'error', 'message': 'Template not found'}

            data['lastModified'] = datetime.datetime.now().isoformat()
            self.save_templates(data)
            return {'status': 'success'}
        except Exception as e:
            logging.error(f"Error removing template: {e}")
            return {'status': 'error', 'message': str(e)}

    def copy_template_to_folder(
        self,
        template_id,
        destination_folder,
        new_name=None,
        context=None,
        options=None,
        conflict_policy='timestamp',
    ):
        """Copies a template file to a folder, with optional open-after-save behavior."""
        try:
            data = self.get_templates()
            template = next((t for t in data['templates'] if t['id'] == template_id), None)

            if not template:
                return {'status': 'error', 'message': 'Template not found'}

            source = template['sourcePath']
            if not os.path.exists(source):
                return {'status': 'error', 'message': f'Template file not found: {source}'}

            if isinstance(destination_folder, (list, tuple)):
                destination_folder = destination_folder[0] if destination_folder else ''

            destination_folder = os.path.normpath(str(destination_folder or '').strip())
            if not destination_folder:
                return {'status': 'error', 'message': 'Destination folder is required'}

            options_payload = dict(options or {}) if isinstance(options, dict) else {}
            context_payload = dict(context or {}) if isinstance(context, dict) else {}

            if not os.path.isdir(destination_folder):
                create_destination = False
                if options_payload:
                    create_destination = bool(options_payload.get('createDestination'))
                if create_destination:
                    os.makedirs(destination_folder, exist_ok=True)
                else:
                    return {'status': 'error', 'message': 'Destination folder does not exist'}

            base_name, extension, normalized_key = self._get_template_output_name_parts(
                template,
                template_key=options_payload.get('templateKey'),
                context=context_payload,
                new_name=new_name,
            )
            if not extension:
                return {'status': 'error', 'message': 'Template file extension could not be determined.'}

            dest_path, filename, has_conflict = self._resolve_template_destination_path(
                destination_folder,
                base_name,
                extension,
                conflict_policy=conflict_policy,
            )
            if has_conflict:
                return {'status': 'error', 'message': 'File already exists at destination'}

            shutil.copy2(source, dest_path)
            if context_payload or options_payload:
                template_options = dict(options_payload)
                if normalized_key and 'templateKey' not in template_options:
                    template_options['templateKey'] = normalized_key
                _apply_template_context(
                    dest_path,
                    context=context_payload,
                    options=template_options,
                )

            result = {
                'status': 'success',
                'path': dest_path,
                'filename': filename,
            }

            if options_payload.get('openOutputs'):
                opened_folder_path = self._find_project_root_by_id(destination_folder) or destination_folder
                result['openedFolderPath'] = opened_folder_path
                warnings = []

                folder_result = self.open_directory_strict(opened_folder_path)
                if folder_result.get('status') != 'success':
                    folder_error = folder_result.get('message') or 'Failed to open folder.'
                    result['openFolderError'] = folder_error
                    warnings.append(f'Could not open folder: {folder_error}')

                file_result = self.open_path(dest_path)
                if file_result.get('status') != 'success':
                    file_error = file_result.get('message') or 'Failed to open file.'
                    result['openFileError'] = file_error
                    warnings.append(f'Could not open file: {file_error}')

                if warnings:
                    result['warnings'] = warnings

            return result
        except Exception as e:
            logging.error(f"Error copying template: {e}")
            return {'status': 'error', 'message': str(e)}

    def copy_template_to_path(self, template_id, destination_path, context=None, options=None):
        """Copies a template file to the specified file path."""
        try:
            data = self.get_templates()
            template = next((t for t in data['templates'] if t['id'] == template_id), None)

            if not template:
                return {'status': 'error', 'message': 'Template not found'}

            source = template['sourcePath']
            if not os.path.exists(source):
                return {'status': 'error', 'message': f'Template file not found: {source}'}

            if isinstance(destination_path, (list, tuple)):
                destination_path = destination_path[0] if destination_path else ''

            dest_path = str(destination_path or '').strip()
            if not dest_path:
                return {'status': 'error', 'message': 'Destination path is required'}

            ext = os.path.splitext(source)[1]
            if ext and not dest_path.lower().endswith(ext.lower()):
                dest_path = dest_path + ext

            dest_dir = os.path.dirname(dest_path)
            if not dest_dir:
                return {'status': 'error', 'message': 'Destination path is invalid'}
            if not os.path.isdir(dest_dir):
                os.makedirs(dest_dir, exist_ok=True)

            if os.path.exists(dest_path):
                return {'status': 'error', 'message': 'File already exists at destination'}

            shutil.copy2(source, dest_path)
            if context or options:
                _apply_template_context(dest_path, context=context, options=options)
            return {'status': 'success', 'path': dest_path}
        except Exception as e:
            logging.error(f"Error copying template to path: {e}")
            return {'status': 'error', 'message': str(e)}

    def create_template_for_workroom(self, template_key, launch_context=None, context=None, conflict_policy='timestamp'):
        """Auto-create a template in the active Workroom project's saved path."""
        try:
            settings = self.get_user_settings()
            workroom_context = self._resolve_workroom_context(settings, launch_context)
            source = str(workroom_context.get('source') or '').strip().lower()
            project_path = str(workroom_context.get('project_path') or '').strip()
            if source != 'workroom':
                logging.info(
                    "create_template_for_workroom: fallback to manual save dialog "
                    f"(source={source or 'none'})."
                )
                return {
                    'status': 'fallback',
                    'autoCreated': False,
                    'fallbackUsed': True,
                    'message': 'Template auto-create is only enabled for Workroom launches.',
                }
            if not project_path:
                logging.info(
                    "create_template_for_workroom: fallback to manual save dialog (missing project path)."
                )
                return {
                    'status': 'fallback',
                    'autoCreated': False,
                    'fallbackUsed': True,
                    'message': 'Saved project path is missing.',
                }
            if not os.path.isdir(project_path):
                logging.info(
                    "create_template_for_workroom: fallback to manual save dialog "
                    f"(project path is not a directory: {project_path})."
                )
                return {
                    'status': 'fallback',
                    'autoCreated': False,
                    'fallbackUsed': True,
                    'message': 'Saved project path is invalid.',
                }

            template, normalized_key = self._find_template_by_key(template_key)
            if not normalized_key:
                return {
                    'status': 'error',
                    'autoCreated': False,
                    'fallbackUsed': False,
                    'message': f'Unsupported template key: {template_key}',
                }
            if not template:
                return {
                    'status': 'error',
                    'autoCreated': False,
                    'fallbackUsed': False,
                    'message': f'Template not found for key: {normalized_key}',
                }

            source_path = str(template.get('sourcePath') or '').strip()
            if not source_path or not os.path.exists(source_path):
                return {
                    'status': 'error',
                    'autoCreated': False,
                    'fallbackUsed': False,
                    'message': f'Template file not found: {source_path or "<empty>"}',
                }

            extension = os.path.splitext(source_path)[1]
            if not extension:
                ext_hint = str(template.get('fileType') or '').strip().lstrip('.')
                extension = f'.{ext_hint}' if ext_hint else ''
            if not extension:
                return {
                    'status': 'error',
                    'autoCreated': False,
                    'fallbackUsed': False,
                    'message': 'Template file extension could not be determined.',
                }

            default_name = TEMPLATE_DEFAULT_FILENAME_BY_KEY.get(
                normalized_key,
                str(template.get('name') or 'Template').strip() or 'Template',
            )
            destination_path, filename, has_conflict = self._resolve_template_destination_path(
                project_path,
                default_name,
                extension,
                conflict_policy=conflict_policy,
            )
            if has_conflict:
                return {
                    'status': 'error',
                    'autoCreated': False,
                    'fallbackUsed': False,
                    'message': 'File already exists at destination.',
                }

            context_payload = dict(context or {}) if isinstance(context, dict) else {}
            if 'projectName' not in context_payload:
                context_payload['projectName'] = os.path.basename(
                    project_path.rstrip("\\/"))
            options_payload = {'templateKey': normalized_key}

            shutil.copy2(source_path, destination_path)
            _apply_template_context(
                destination_path, context=context_payload, options=options_payload)

            logging.info(
                "create_template_for_workroom: auto-created template "
                f"(key={normalized_key}, path={destination_path})."
            )
            return {
                'status': 'success',
                'path': destination_path,
                'filename': filename,
                'autoCreated': True,
                'fallbackUsed': False,
                'message': 'Template created.',
            }
        except Exception as e:
            logging.error(f"create_template_for_workroom failed: {e}")
            return {
                'status': 'error',
                'autoCreated': False,
                'fallbackUsed': False,
                'message': str(e),
            }

    def select_folder(self):
        """Shows a folder selection dialog."""
        try:
            window = webview.windows[0]
            folder_path = window.create_file_dialog(
                webview.FOLDER_DIALOG
            )
            if not folder_path:
                return {'status': 'cancelled', 'path': None}
            # folder_path is a tuple, get first element
            return {'status': 'success', 'path': folder_path[0] if folder_path else None}
        except Exception as e:
            logging.error(f"Error in folder dialog: {e}")
            return {'status': 'error', 'message': str(e)}

    def select_template_output_folder(self, default_dir=None):
        """Shows a folder dialog for template-tool output."""
        try:
            window = webview.windows[0]
            directory = str(default_dir or '').strip() or get_default_documents_dir()
            if os.path.isfile(directory):
                directory = os.path.dirname(directory)
            if not os.path.isdir(directory):
                directory = get_default_documents_dir()

            folder_path = window.create_file_dialog(
                webview.FOLDER_DIALOG,
                directory=directory,
            )
            if not folder_path:
                return {'status': 'cancelled', 'path': None}
            if isinstance(folder_path, (list, tuple)):
                folder_path = folder_path[0] if folder_path else ''
            return {'status': 'success', 'path': folder_path}
        except TypeError:
            return self.select_folder()
        except Exception as e:
            logging.error(f"Error in template output folder dialog: {e}")
            return {'status': 'error', 'message': str(e)}

    def select_template_save_location(self, default_dir=None, default_name=None, file_type=None):
        """Shows a save dialog for template output."""
        try:
            window = webview.windows[0]
            directory = str(default_dir or '').strip() or get_default_documents_dir()
            save_filename = str(default_name or 'Template').strip() or 'Template'

            ext = str(file_type or '').lower().lstrip('.')
            if ext and not save_filename.lower().endswith(f".{ext}"):
                save_filename = f"{save_filename}.{ext}"

            file_type_map = {
                'doc': 'Word Document (*.doc)',
                'docx': 'Word Document (*.docx)',
                'dwg': 'AutoCAD Drawing (*.dwg)',
                'xlsx': 'Excel Spreadsheet (*.xlsx)',
                'xls': 'Excel Spreadsheet (*.xls)',
            }
            file_types = None
            if ext and ext in file_type_map:
                file_types = (file_type_map[ext],)

            file_path = window.create_file_dialog(
                webview.FileDialog.SAVE,
                directory=directory,
                save_filename=save_filename,
                file_types=file_types
            )
            if not file_path:
                return {'status': 'cancelled', 'path': None}
            if isinstance(file_path, (list, tuple)):
                file_path = file_path[0] if file_path else ''
            return {'status': 'success', 'path': file_path}
        except Exception as e:
            logging.error(f"Error in save dialog: {e}")
            return {'status': 'error', 'message': str(e)}

    def select_template_file(self):
        """Shows file dialog for selecting template files."""
        try:
            window = webview.windows[0]
            file_types = (
                'Document Files (*.doc;*.docx)',
                'AutoCAD Files (*.dwg)',
                'Excel Files (*.xlsx;*.xls)',
                'All Supported (*.doc;*.docx;*.dwg;*.xlsx;*.xls)'
            )
            file_paths = window.create_file_dialog(
                webview.OPEN_DIALOG,
                allow_multiple=False,
                file_types=file_types
            )
            if not file_paths:
                return {'status': 'cancelled', 'path': None}
            return {'status': 'success', 'path': file_paths[0]}
        except Exception as e:
            logging.error(f"Error in file dialog: {e}")
            return {'status': 'error', 'message': str(e)}

    def verify_template_exists(self, template_id):
        """Verifies if a template's source file still exists."""
        try:
            data = self.get_templates()
            template = next((t for t in data['templates'] if t['id'] == template_id), None)
            if not template:
                return {'status': 'error', 'message': 'Template not found'}
            exists = os.path.exists(template['sourcePath'])
            return {'status': 'success', 'exists': exists, 'path': template['sourcePath']}
        except Exception as e:
            return {'status': 'error', 'message': str(e)}

    # ===================== CHECKLISTS API =====================

    def get_checklists(self):
        """Reads and returns checklists data."""
        try:
            with open(CHECKLISTS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            # First run - return empty structure
            return {'checklists': [], 'lastModified': None}

    def save_checklists(self, data):
        """Saves checklists data with backup."""
        try:
            if os.path.exists(CHECKLISTS_FILE):
                shutil.copy2(CHECKLISTS_FILE, CHECKLISTS_FILE + '.bak')
            with open(CHECKLISTS_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return {'status': 'success'}
        except Exception as e:
            logging.error(f"Error saving checklists: {e}")
            return {'status': 'error', 'message': str(e)}

    def _find_expense_signature_row(self, worksheet):
        for row_idx in range(1, worksheet.max_row + 1):
            value = worksheet.cell(row=row_idx, column=1).value
            if isinstance(value, str) and value.strip().upper() == "EMPLOYEE:":
                return row_idx
        raise ValueError("Could not locate the expense-sheet signature block.")

    def _get_expense_section_starts(self, worksheet, signature_row):
        section_starts = []
        for row_idx in range(1, signature_row):
            value = worksheet.cell(row=row_idx, column=1).value
            if isinstance(value, str) and value.strip().upper().startswith("PROJECT:"):
                section_starts.append(row_idx)
        return section_starts

    def _capture_expense_section_template(self, worksheet, start_row):
        section_end_row = start_row + EXPENSE_SECTION_ROWS - 1
        row_merge_map = {}

        for merged_range in worksheet.merged_cells.ranges:
            if merged_range.min_row < start_row or merged_range.max_row > section_end_row:
                continue
            row_offset = merged_range.min_row - start_row
            row_merge_map.setdefault(row_offset, []).append(
                (merged_range.min_col, merged_range.max_col)
            )

        rows = []
        for row_offset in range(EXPENSE_SECTION_ROWS):
            source_row = start_row + row_offset
            row_dimension = worksheet.row_dimensions[source_row]
            rows.append({
                "height": row_dimension.height,
                "hidden": row_dimension.hidden,
                "outlineLevel": row_dimension.outlineLevel,
                "collapsed": row_dimension.collapsed,
                "merges": list(row_merge_map.get(row_offset, [])),
                "cells": [
                    {
                        "value": worksheet.cell(row=source_row, column=col_idx).value,
                        "style": copy(worksheet.cell(row=source_row, column=col_idx)._style),
                    }
                    for col_idx in range(1, worksheet.max_column + 1)
                ],
            })

        return {
            "max_column": worksheet.max_column,
            "rows": rows,
        }

    def _apply_expense_template_row(self, worksheet, template, row_offset, target_row, copy_values=True):
        row_template = template["rows"][row_offset]

        for col_idx, cell_template in enumerate(row_template["cells"], start=1):
            target_cell = worksheet.cell(row=target_row, column=col_idx)
            target_cell._style = copy(cell_template["style"])
            target_cell.value = cell_template["value"] if copy_values else None

        target_dimension = worksheet.row_dimensions[target_row]
        target_dimension.height = row_template["height"]
        target_dimension.hidden = row_template["hidden"]
        target_dimension.outlineLevel = row_template["outlineLevel"]
        target_dimension.collapsed = row_template["collapsed"]

        for min_col, max_col in row_template["merges"]:
            worksheet.merge_cells(
                start_row=target_row,
                start_column=min_col,
                end_row=target_row,
                end_column=max_col,
            )

    def _apply_expense_section_template(self, worksheet, template, target_start_row):
        for row_offset in range(EXPENSE_SECTION_ROWS):
            self._apply_expense_template_row(
                worksheet,
                template,
                row_offset,
                target_start_row + row_offset,
                copy_values=True,
            )

    def _reset_expense_section_merges(self, worksheet, section_start, actual_data_rows):
        section_end = section_start + EXPENSE_HEADER_ROWS + actual_data_rows + EXPENSE_FOOTER_ROWS - 1
        merges_to_remove = []
        for merged_range in worksheet.merged_cells.ranges:
            if merged_range.min_row < section_start or merged_range.max_row > section_end:
                continue
            if merged_range.max_col > 5:
                continue
            merges_to_remove.append(str(merged_range))

        for merged_range in merges_to_remove:
            worksheet.unmerge_cells(merged_range)

        header_row = section_start + 2
        worksheet.merge_cells(
            start_row=header_row,
            start_column=2,
            end_row=header_row,
            end_column=3,
        )

        data_start_row = section_start + EXPENSE_HEADER_ROWS
        subtotal_row = data_start_row + actual_data_rows
        for row_idx in range(data_start_row, subtotal_row):
            worksheet.merge_cells(
                start_row=row_idx,
                start_column=2,
                end_row=row_idx,
                end_column=3,
            )

        total_row = subtotal_row + 2
        worksheet.merge_cells(
            start_row=total_row,
            start_column=4,
            end_row=total_row,
            end_column=5,
        )

    def _capture_worksheet_row_template(self, worksheet, source_row):
        row_dimension = worksheet.row_dimensions[source_row]
        return {
            "height": row_dimension.height,
            "hidden": row_dimension.hidden,
            "outlineLevel": row_dimension.outlineLevel,
            "collapsed": row_dimension.collapsed,
            "cells": [
                {
                    "value": worksheet.cell(row=source_row, column=col_idx).value,
                    "style": copy(worksheet.cell(row=source_row, column=col_idx)._style),
                }
                for col_idx in range(1, worksheet.max_column + 1)
            ],
        }

    def _apply_worksheet_row_template(self, worksheet, template, target_row, copy_values=True):
        for col_idx, cell_template in enumerate(template["cells"], start=1):
            target_cell = worksheet.cell(row=target_row, column=col_idx)
            target_cell._style = copy(cell_template["style"])
            target_cell.value = cell_template["value"] if copy_values else None

        target_dimension = worksheet.row_dimensions[target_row]
        target_dimension.height = template["height"]
        target_dimension.hidden = template["hidden"]
        target_dimension.outlineLevel = template["outlineLevel"]
        target_dimension.collapsed = template["collapsed"]

    def _render_project_expense_sheet(self, worksheet, projects, mileage_rate):
        if not projects:
            return {"image_paths": [], "image_start_row": None}

        signature_row = self._find_expense_signature_row(worksheet)
        section_starts = self._get_expense_section_starts(worksheet, signature_row)
        if not section_starts:
            raise ValueError("No formatted expense sections were found in the template.")

        section_template = self._capture_expense_section_template(worksheet, section_starts[0])
        first_section_start = section_starts[0]
        expense_detail_number_format = worksheet.cell(
            row=first_section_start + EXPENSE_HEADER_ROWS,
            column=5,
        ).number_format
        expense_subtotal_number_format = worksheet.cell(
            row=first_section_start + EXPENSE_HEADER_ROWS + EXPENSE_DEFAULT_DATA_ROWS,
            column=5,
        ).number_format
        expense_rate_number_format = worksheet.cell(
            row=first_section_start + EXPENSE_HEADER_ROWS + EXPENSE_DEFAULT_DATA_ROWS + 1,
            column=4,
        ).number_format
        expense_total_number_format = worksheet.cell(
            row=first_section_start + EXPENSE_HEADER_ROWS + EXPENSE_DEFAULT_DATA_ROWS + 2,
            column=4,
        ).number_format

        while len(section_starts) < len(projects):
            worksheet.insert_rows(signature_row, EXPENSE_SECTION_ROWS)
            self._apply_expense_section_template(worksheet, section_template, signature_row)
            section_starts.append(signature_row)
            signature_row += EXPENSE_SECTION_ROWS

        section_layouts = []

        for project_index, project in enumerate(projects):
            section_start = section_starts[project_index]
            entries = project.get('entries', []) or []
            num_entries = max(len(entries), 1)
            extra_rows = max(0, num_entries - EXPENSE_DEFAULT_DATA_ROWS)

            if extra_rows:
                insert_at = section_start + EXPENSE_HEADER_ROWS + EXPENSE_DEFAULT_DATA_ROWS
                worksheet.insert_rows(insert_at, extra_rows)
                for row_offset in range(extra_rows):
                    self._apply_expense_template_row(
                        worksheet,
                        section_template,
                        EXPENSE_INSERTED_DATA_TEMPLATE_OFFSET,
                        insert_at + row_offset,
                        copy_values=False,
                    )
                for later_index in range(project_index + 1, len(section_starts)):
                    section_starts[later_index] += extra_rows
                signature_row += extra_rows

            data_start_row = section_start + EXPENSE_HEADER_ROWS
            actual_data_rows = max(num_entries, EXPENSE_DEFAULT_DATA_ROWS)
            subtotal_row = data_start_row + actual_data_rows
            rate_row = subtotal_row + 1
            total_row = rate_row + 1

            section_layouts.append({
                "project": project,
                "section_start": section_start,
                "data_start_row": data_start_row,
                "actual_data_rows": actual_data_rows,
                "subtotal_row": subtotal_row,
                "rate_row": rate_row,
                "total_row": total_row,
            })

        for section_layout in section_layouts:
            self._reset_expense_section_merges(
                worksheet,
                section_layout["section_start"],
                section_layout["actual_data_rows"],
            )

        image_paths = []
        for section_layout in section_layouts:
            project = section_layout["project"]
            entries = project.get('entries', []) or []
            section_start = section_layout["section_start"]
            data_start_row = section_layout["data_start_row"]
            subtotal_row = section_layout["subtotal_row"]
            rate_row = section_layout["rate_row"]
            total_row = section_layout["total_row"]

            worksheet.cell(
                row=section_start,
                column=1,
                value=f"PROJECT: {project.get('projectName', '')}",
            )
            worksheet.cell(
                row=section_start + 1,
                column=1,
                value=f"JOB #: {project.get('projectId', '')}",
            )

            for row_idx in range(data_start_row, subtotal_row):
                worksheet.cell(row=row_idx, column=1, value=None)
                worksheet.cell(row=row_idx, column=2, value=None)
                worksheet.cell(row=row_idx, column=4, value=None)
                worksheet.cell(row=row_idx, column=5, value=None)
                worksheet.cell(row=row_idx, column=5).number_format = expense_detail_number_format

            for entry_index, entry in enumerate(entries):
                row_idx = data_start_row + entry_index
                worksheet.cell(row=row_idx, column=1, value=entry.get('date', ''))
                worksheet.cell(row=row_idx, column=2, value=entry.get('description', ''))

                mileage = entry.get('mileage', 0) or 0
                expense = entry.get('expense', 0) or 0

                worksheet.cell(row=row_idx, column=4, value=mileage if mileage else '')
                expense_cell = worksheet.cell(
                    row=row_idx,
                    column=5,
                    value=expense if expense else '',
                )
                expense_cell.number_format = expense_detail_number_format

            worksheet.cell(row=subtotal_row, column=3, value="SUBTOTAL")
            worksheet.cell(
                row=subtotal_row,
                column=4,
                value=f'=SUM(D{data_start_row}:D{subtotal_row - 1})',
            )
            worksheet.cell(
                row=subtotal_row,
                column=5,
                value=f'=SUM(E{data_start_row}:E{subtotal_row - 1})',
            )
            worksheet.cell(row=subtotal_row, column=5).number_format = expense_subtotal_number_format

            worksheet.cell(
                row=rate_row,
                column=3,
                value=f"{mileage_rate:.2f} CENTS PER MILE",
            )
            worksheet.cell(
                row=rate_row,
                column=4,
                value=f'=D{subtotal_row}*{mileage_rate}',
            )
            worksheet.cell(row=rate_row, column=4).number_format = expense_rate_number_format
            worksheet.cell(row=rate_row, column=5, value=None)

            worksheet.cell(row=total_row, column=3, value="TOTAL")
            worksheet.cell(
                row=total_row,
                column=4,
                value=f'=D{rate_row}+E{subtotal_row}',
            )
            worksheet.cell(row=total_row, column=4).number_format = expense_total_number_format

            for image in project.get('images', []):
                image_path = image.get('path', '')
                if image_path:
                    resolved_image_path = self._resolve_expense_image_path(image_path)
                    if resolved_image_path:
                        image_paths.append(resolved_image_path)

        return {
            "image_paths": image_paths,
            "image_start_row": signature_row + EXPENSE_IMAGE_ROW_OFFSET,
        }

    def _add_expense_sheet_images(self, worksheet, image_paths, image_start_row):
        from openpyxl.drawing.image import Image as XLImage
        from openpyxl.utils.units import points_to_pixels

        temp_files_to_cleanup = []
        if not image_paths or image_start_row is None:
            return temp_files_to_cleanup

        target_width = self._get_expense_image_target_width_pixels(worksheet)

        for image_path in image_paths:
            if not os.path.isfile(image_path):
                continue

            try:
                ext = os.path.splitext(image_path)[1].lower()
                if ext == '.pdf':
                    continue

                with PILImage.open(image_path) as source_img:
                    if hasattr(source_img, 'n_frames') and source_img.n_frames > 1:
                        source_img.seek(0)
                    pil_img = ImageOps.exif_transpose(source_img).copy()
                    if pil_img.mode in ('RGBA', 'P', 'LA'):
                        pil_img = pil_img.convert('RGB')
                    elif pil_img.mode != 'RGB':
                        pil_img = pil_img.convert('RGB')

                source_width, source_height = pil_img.size
                if source_width <= 0 or source_height <= 0:
                    continue

                temp_img_path = os.path.join(
                    tempfile.gettempdir(),
                    f"expense_img_{os.urandom(4).hex()}.png",
                )
                pil_img.save(temp_img_path, 'PNG')
                temp_files_to_cleanup.append(temp_img_path)

                image = XLImage(temp_img_path)
                image.width = max(int(round(target_width)), 1)
                image.height = max(
                    int(round(target_width * (source_height / source_width))),
                    1,
                )
                worksheet.add_image(image, f"A{image_start_row}")

                rows_for_image = 0
                covered_height = 0
                next_row = image_start_row
                while covered_height < image.height:
                    row_height_points = worksheet.row_dimensions[next_row].height
                    if row_height_points is None:
                        row_height_points = worksheet.sheet_format.defaultRowHeight or 15
                    covered_height += max(points_to_pixels(row_height_points), 1)
                    rows_for_image += 1
                    next_row += 1
                image_start_row += rows_for_image + 1
            except UnidentifiedImageError:
                continue
            except Exception as img_err:
                logging.warning(f"Could not add image {image_path}: {img_err}")

        return temp_files_to_cleanup

    def _get_expense_image_target_width_pixels(self, worksheet, start_column=1, end_column=5):
        total_width = 0
        default_width = worksheet.sheet_format.defaultColWidth or 8.43
        for column_index in range(start_column, end_column + 1):
            column_letter = openpyxl.utils.get_column_letter(column_index)
            column_width = worksheet.column_dimensions[column_letter].width
            total_width += self._excel_column_width_to_pixels(
                default_width if column_width is None else column_width
            )
        return max(total_width, 1)

    def _excel_column_width_to_pixels(self, column_width):
        width = float(column_width or 8.43)
        return int(math.floor(((256 * width + math.floor(128 / 7)) / 256) * 7))

    def _find_nearest_existing_directory(self, path):
        current = os.path.normpath(str(path or "").strip())
        while current:
            if os.path.isdir(current):
                return current
            parent = os.path.dirname(current)
            if parent == current:
                break
            current = parent
        return ""

    def _resolve_expense_image_path(self, raw_path):
        normalized_path = self._coerce_local_file_path(raw_path)
        if not normalized_path:
            return ""
        if os.path.isfile(normalized_path):
            return normalized_path

        filename = os.path.basename(normalized_path)
        if not filename:
            return ""

        search_root = self._find_nearest_existing_directory(os.path.dirname(normalized_path))
        if not search_root:
            return ""

        original_parent = os.path.basename(os.path.dirname(normalized_path).rstrip("\\/")).strip().lower()
        exact_parent_matches = []
        any_matches = []

        try:
            for candidate in Path(search_root).rglob(filename):
                candidate_path = os.path.normpath(str(candidate))
                if not os.path.isfile(candidate_path):
                    continue
                any_matches.append(candidate_path)
                candidate_parent = os.path.basename(
                    os.path.dirname(candidate_path).rstrip("\\/")
                ).strip().lower()
                if original_parent and candidate_parent == original_parent:
                    exact_parent_matches.append(candidate_path)
        except Exception:
            return ""

        if len(exact_parent_matches) == 1:
            return exact_parent_matches[0]
        if len(any_matches) == 1:
            return any_matches[0]
        if exact_parent_matches:
            return sorted(exact_parent_matches)[0]
        return ""

    def export_timesheet_excel(self, data):
        """Exports timesheet data to an Excel file using a template."""
        try:
            import openpyxl

            template_path = get_bundled_template_path("Template_Timesheet.xlsx")

            if not template_path.exists():
                return {'status': 'error', 'message': f'Template file not found: {template_path}'}

            # Determine output path
            week_key = data.get('weekKey', '')
            user_name = data.get('userName', '')
            file_path = data.get('filePath')
            if not file_path:
                default_dir = get_default_documents_dir()
                default_name = build_timesheet_filename(user_name, week_key)
                window = webview.windows[0]
                file_path = window.create_file_dialog(
                    webview.FileDialog.SAVE,
                    directory=default_dir,
                    save_filename=default_name,
                    file_types=('Excel Files (*.xlsx)',)
                )
                if not file_path:
                    return {'status': 'cancelled'}
            if isinstance(file_path, (list, tuple)):
                file_path = file_path[0] if file_path else ''
            if not file_path:
                return {'status': 'cancelled'}
            if not str(file_path).lower().endswith('.xlsx'):
                file_path = f"{file_path}.xlsx"

            # Copy template to output location
            shutil.copy2(str(template_path), file_path)

            # Open the copied file and modify the "time log" sheet
            wb = openpyxl.load_workbook(file_path)

            # Get the "time log" sheet
            if "time log" not in wb.sheetnames:
                return {'status': 'error', 'message': 'Sheet "time log" not found in template'}

            ws = wb["time log"]
            expense_data = data.get('expenses', {})
            expense_projects = expense_data.get('projects', [])

            # Row 1: Employee name (column M = 13)
            ws.cell(row=1, column=13, value=data.get('userName', 'Employee'))

            # Row 2: Week of (column M = 13)
            ws.cell(row=2, column=13, value=data.get('weekDisplay', ''))

            # Data rows start at row 5
            entries = data.get('entries', [])
            data_start_row = 5
            insert_after_row = 22
            insert_at_row = insert_after_row + 1
            base_totals_row = 26
            max_template_project_rows_before_insert = insert_after_row - data_start_row + 1
            overflow_rows = max(0, len(entries) - max_template_project_rows_before_insert)

            if overflow_rows:
                inserted_row_template = self._capture_worksheet_row_template(ws, insert_after_row)
                ws.insert_rows(insert_at_row, overflow_rows)
                for row_offset in range(overflow_rows):
                    self._apply_worksheet_row_template(
                        ws,
                        inserted_row_template,
                        insert_at_row + row_offset,
                        copy_values=False,
                    )

            day_order = TIMESHEET_DAY_ORDER
            export_mileage_by_row = _build_export_mileage_by_row(entries, expense_projects)

            for row_idx, entry in enumerate(entries, data_start_row):
                # Column A (1): PROJECT #
                ws.cell(row=row_idx, column=1, value=entry.get('projectId', ''))
                # Column B (2): Task #
                ws.cell(row=row_idx, column=2, value=entry.get('taskNumber', ''))
                # Column C (3): PROJECT NAME
                ws.cell(row=row_idx, column=3, value=entry.get('projectName', ''))
                # Column D (4): FUNCTION (M E P F)
                ws.cell(row=row_idx, column=4, value=entry.get('function', ''))
                # Column E (5): PM
                ws.cell(row=row_idx, column=5, value=entry.get('pmInitials', ''))
                # Columns F-I (6-9): DESIGN/ENG, DRAFTING, TRAVEL, OTHERS - leave empty
                # Column J (10): SERVICE & WORK DESCRIPTION
                ws.cell(row=row_idx, column=10, value=entry.get('serviceDescription', ''))

                # Columns K-Q (11-17): MON through SUN
                hours = entry.get('hours', {})
                for day_idx, day in enumerate(day_order):
                    h = hours.get(day, 0)
                    ws.cell(row=row_idx, column=11 + day_idx, value=h if h else '')

                # Column R (18): MILEAGE derived from project expense dates.
                mileage = export_mileage_by_row[row_idx - data_start_row] if row_idx - data_start_row < len(export_mileage_by_row) else 0
                ws.cell(row=row_idx, column=18, value=mileage if mileage else '')

            # Totals row (row 26 in template based on CSV)
            totals_row = base_totals_row + overflow_rows
            data_end_row = totals_row - 1

            for day_idx, _day in enumerate(day_order):
                col_idx = 11 + day_idx
                col_letter = openpyxl.utils.get_column_letter(col_idx)
                ws.cell(
                    row=totals_row,
                    column=col_idx,
                    value=f'=SUM({col_letter}{data_start_row}:{col_letter}{data_end_row})',
                )

            mileage_col_letter = openpyxl.utils.get_column_letter(18)
            ws.cell(
                row=totals_row,
                column=18,
                value=f'=SUM({mileage_col_letter}{data_start_row}:{mileage_col_letter}{data_end_row})',
            )

            # =============== EXPENSE SHEET EXPORT ===============
            # If expense data is provided, also populate the expense sheet
            temp_files_to_cleanup = []

            if expense_projects:
                if EXPENSE_SHEET_NAME in wb.sheetnames:
                    ws_exp = wb[EXPENSE_SHEET_NAME]
                    render_result = self._render_project_expense_sheet(
                        ws_exp,
                        expense_projects,
                        expense_data.get('mileageRate', 0.70),
                    )
                    temp_files_to_cleanup.extend(
                        self._add_expense_sheet_images(
                            ws_exp,
                            render_result["image_paths"],
                            render_result["image_start_row"],
                        )
                    )

            # Save the modified file
            wb.save(file_path)

            # Clean up temp files after save
            for temp_file in temp_files_to_cleanup:
                try:
                    os.remove(temp_file)
                except:
                    pass

            # Open file on Windows
            if sys.platform == "win32":
                os.startfile(file_path)

            return {'status': 'success', 'path': file_path}

        except ImportError:
            return {'status': 'error', 'message': 'openpyxl not installed. Run: pip install openpyxl'}
        except Exception as e:
            logging.error(f"Error exporting timesheet: {e}")
            return {'status': 'error', 'message': str(e)}

    def select_expense_images(self):
        """Shows file dialog for selecting expense receipt images."""
        try:
            window = webview.windows[0]
            file_types = (
                'Image Files (*.jpg;*.jpeg;*.png;*.gif;*.bmp)',
                'PDF Files (*.pdf)',
                'All Files (*.*)'
            )
            file_paths = window.create_file_dialog(
                webview.OPEN_DIALOG,
                allow_multiple=True,
                file_types=file_types
            )
            if not file_paths:
                return {'status': 'cancelled', 'paths': []}
            return {'status': 'success', 'paths': list(file_paths)}
        except Exception as e:
            logging.error(f"Error in expense image selection dialog: {e}")
            return {'status': 'error', 'message': str(e)}

    def resolve_expense_attachment_path(self, path):
        """Resolves a stored expense attachment path to an existing local file."""
        try:
            resolved_path = self._resolve_expense_image_path(path)
            if not resolved_path or not os.path.isfile(resolved_path):
                return {'status': 'error', 'message': 'Attachment file not found.'}
            return {
                'status': 'success',
                'path': resolved_path,
                'filename': os.path.basename(resolved_path),
            }
        except Exception as e:
            logging.error(f"Error resolving expense attachment path: {e}")
            return {'status': 'error', 'message': str(e)}

    def get_expense_image_preview(self, path, max_size=1600):
        """Returns a browser-safe preview data URL for a local expense image."""
        try:
            resolved_path = self._resolve_expense_image_path(path)
            if not resolved_path or not os.path.isfile(resolved_path):
                return {'status': 'error', 'message': 'Image file not found.'}

            try:
                preview_max_size = max(1, int(max_size or 1600))
            except Exception:
                preview_max_size = 1600

            with PILImage.open(resolved_path) as source_img:
                if hasattr(source_img, 'n_frames') and source_img.n_frames > 1:
                    source_img.seek(0)
                preview = ImageOps.exif_transpose(source_img).copy()
                if preview.mode in ('P', 'LA'):
                    preview = preview.convert('RGBA')
                elif preview.mode not in ('RGB', 'RGBA'):
                    preview = preview.convert('RGB')

            resampling = getattr(
                getattr(PILImage, 'Resampling', PILImage),
                'LANCZOS',
                getattr(PILImage, 'LANCZOS', getattr(PILImage, 'BICUBIC', 3))
            )
            preview.thumbnail((preview_max_size, preview_max_size), resampling)

            buffer = io.BytesIO()
            preview.save(buffer, format='PNG')
            encoded = base64.b64encode(buffer.getvalue()).decode('ascii')

            return {
                'status': 'success',
                'dataUrl': f'data:image/png;base64,{encoded}',
                'width': preview.width,
                'height': preview.height,
                'path': resolved_path,
                'filename': os.path.basename(resolved_path),
            }
        except UnidentifiedImageError:
            resolved_path = self._resolve_expense_image_path(path)
            return {
                'status': 'unsupported',
                'message': 'Preview not available for this file type.',
                'path': resolved_path or '',
                'filename': os.path.basename(resolved_path) if resolved_path else '',
            }
        except Exception as e:
            logging.error(f"Error generating expense image preview: {e}")
            return {'status': 'error', 'message': str(e)}

    def export_expense_sheet_excel(self, data):
        """Exports expense sheet data to an Excel file with images."""
        try:
            import openpyxl

            template_path = get_bundled_template_path("Template_Timesheet.xlsx")

            if not template_path.exists():
                return {'status': 'error', 'message': f'Template file not found: {template_path}'}

            # Determine output path
            week_key = data.get('weekKey', 'expense')
            user_docs = os.path.join(os.path.expanduser('~'), 'Documents')
            file_path = os.path.join(user_docs, f"Expense_Sheet_{week_key}.xlsx")

            # Copy template to output location
            shutil.copy2(str(template_path), file_path)

            # Open the copied file
            wb = openpyxl.load_workbook(file_path)

            # Get the expense sheet
            if EXPENSE_SHEET_NAME not in wb.sheetnames:
                return {'status': 'error', 'message': f'Sheet "{EXPENSE_SHEET_NAME}" not found in template'}

            ws = wb[EXPENSE_SHEET_NAME]
            render_result = self._render_project_expense_sheet(
                ws,
                data.get('projects', []),
                data.get('mileageRate', 0.70),
            )
            temp_files_to_cleanup = self._add_expense_sheet_images(
                ws,
                render_result["image_paths"],
                render_result["image_start_row"],
            )

            # Save the modified file
            wb.save(file_path)

            # Clean up temp files after save
            for temp_file in temp_files_to_cleanup:
                try:
                    os.remove(temp_file)
                except:
                    pass

            # Open file on Windows
            if sys.platform == "win32":
                os.startfile(file_path)

            return {'status': 'success', 'path': file_path}

        except ImportError:
            return {'status': 'error', 'message': 'openpyxl not installed. Run: pip install openpyxl'}
        except Exception as e:
            logging.error(f"Error exporting expense sheet: {e}")
            return {'status': 'error', 'message': str(e)}

    def mark_overdue_projects_complete(self):
        """Marks all deliverables with due dates before today as complete."""
        try:
            tasks = self.get_tasks()
            today = datetime.date.today()
            count = 0
            for task in tasks:
                deliverables = task.get('deliverables')
                if isinstance(deliverables, list):
                    for deliverable in deliverables:
                        due_str = deliverable.get('due', '')
                        if due_str:
                            due_date = parse_due_str(due_str)
                            if due_date and due_date.date() < today:
                                deliverable['statuses'] = ['Complete']
                                deliverable['status'] = 'Complete'
                                sync_status_arrays(deliverable)
                                if isinstance(deliverable.get('tasks'), list):
                                    for t in deliverable['tasks']:
                                        t['done'] = True
                                count += 1
                else:
                    due_str = task.get('due', '')
                    if due_str:
                        due_date = parse_due_str(due_str)
                        if due_date and due_date.date() < today:
                            task['statuses'] = ['Complete']
                            task['status'] = 'Complete'
                            sync_status_arrays(task)
                            if isinstance(task.get('tasks'), list):
                                for t in task['tasks']:
                                    t['done'] = True
                            count += 1
            if count > 0:
                self.save_tasks(tasks)
            return {'status': 'success', 'count': count}
        except Exception as e:
            logging.error(f"Error marking overdue projects: {e}")
            return {'status': 'error', 'message': str(e)}

    def mark_overdue_projects_delivered(self):
        """Marks all deliverables with due dates before today as delivered."""
        try:
            tasks = self.get_tasks()
            today = datetime.date.today()
            count = 0
            for task in tasks:
                deliverables = task.get('deliverables')
                if isinstance(deliverables, list):
                    for deliverable in deliverables:
                        due_str = deliverable.get('due', '')
                        if due_str:
                            due_date = parse_due_str(due_str)
                            if due_date and due_date.date() < today:
                                deliverable['statuses'] = ['Delivered']
                                deliverable['status'] = 'Delivered'
                                sync_status_arrays(deliverable)
                                if isinstance(deliverable.get('tasks'), list):
                                    for t in deliverable['tasks']:
                                        t['done'] = True
                                count += 1
                else:
                    due_str = task.get('due', '')
                    if due_str:
                        due_date = parse_due_str(due_str)
                        if due_date and due_date.date() < today:
                            task['statuses'] = ['Delivered']
                            task['status'] = 'Delivered'
                            sync_status_arrays(task)
                            if isinstance(task.get('tasks'), list):
                                for t in task['tasks']:
                                    t['done'] = True
                            count += 1
            if count > 0:
                self.save_tasks(tasks)
            return {'status': 'success', 'count': count}
        except Exception as e:
            logging.error(f"Error marking overdue projects delivered: {e}")
            return {'status': 'error', 'message': str(e)}

    def delete_all_notes(self):
        """Deletes all notes data."""
        try:
            if os.path.exists(NOTES_FILE):
                os.remove(NOTES_FILE)
            return {'status': 'success'}
        except Exception as e:
            logging.error(f"Error deleting notes: {e}")
            return {'status': 'error', 'message': str(e)}

    def check_autocad_running(self):
        """Checks if AutoCAD is running."""
        is_running = self._is_autocad_running()
        return {'is_running': is_running}

    def uninstall_all_plugins(self):
        """Uninstalls all installed plugins."""
        try:
            # Check AutoCAD again just in case
            if self._is_autocad_running():
                return {'status': 'error', 'message': 'AutoCAD is currently running. Please close AutoCAD before uninstalling.'}

            bundles = os.listdir(self.app_plugins_folder)
            count = 0
            for bundle in bundles:
                if bundle.endswith('.bundle'):
                    bundle_path = os.path.join(self.app_plugins_folder, bundle)
                    try:
                        shutil.rmtree(bundle_path)
                        count += 1
                        logging.info(f"Uninstalled bundle: {bundle}")
                    except Exception as e:
                        logging.error(f"Failed to uninstall {bundle}: {e}")
            return {'status': 'success', 'count': count}
        except Exception as e:
            logging.error(f"Error uninstalling plugins: {e}")
            return {'status': 'error', 'message': str(e)}

    def open_url(self, url, browser=None):
        """Opens a URL or protocol link using the OS handler or a preferred browser."""
        try:
            target = str(url or "").strip()
            preferred_browser = str(browser or "").strip().lower()
            if not target:
                return {'status': 'error', 'message': 'URL is required.'}

            if target.lower().startswith("file://"):
                local_path = self._coerce_local_email_path(target)
                if local_path:
                    return self.open_path(local_path)

            if sys.platform == "win32":
                if preferred_browser == "edge" and target.lower().startswith(("http://", "https://")):
                    try:
                        os.startfile(f"microsoft-edge:{target}")
                    except OSError:
                        logging.warning("Could not launch Microsoft Edge directly; falling back to the default browser.")
                        os.startfile(target)
                else:
                    os.startfile(target)
            else:
                subprocess.run(
                    ['open', target] if sys.platform == "darwin" else ['xdg-open', target],
                    check=False
                )
            return {'status': 'success'}
        except Exception as e:
            logging.error(f"Error opening URL: {e}")
            return {'status': 'error', 'message': str(e)}

    def open_notepad_with_text(self, text):
        """Writes text to a temp file and opens it in Notepad on Windows."""
        try:
            content = str(text or "")
            if not content.strip():
                return {'status': 'error', 'message': 'Text is required.'}

            temp_dir = os.path.join(get_app_data_dir(), 'temp')
            os.makedirs(temp_dir, exist_ok=True)

            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            unique = uuid.uuid4().hex[:8]
            output_path = os.path.join(
                temp_dir, f"deliverables_export_{timestamp}_{unique}.txt")
            with open(output_path, 'w', encoding='utf-8', newline='') as f:
                f.write(content)

            if sys.platform == "win32":
                subprocess.Popen(["notepad.exe", output_path])
            else:
                self.open_path(output_path)

            return {'status': 'success', 'path': output_path}
        except Exception as e:
            logging.error(f"Error opening Notepad text export: {e}")
            return {'status': 'error', 'message': str(e)}

    def export_deliverables_excel(self, data):
        """Exports selected deliverables to a basic Excel workbook."""
        try:
            payload = data or {}
            if isinstance(payload, str):
                payload = json.loads(payload)
            if not isinstance(payload, dict):
                return {'status': 'error', 'message': 'Invalid export payload.'}

            raw_entries = payload.get('entries', [])
            if not isinstance(raw_entries, list):
                return {'status': 'error', 'message': 'Deliverable entries are required.'}

            cleaned_entries = []
            for raw_entry in raw_entries:
                if not isinstance(raw_entry, dict):
                    continue

                due_raw = str(raw_entry.get('due') or '').strip()
                due_value = parse_due_str(due_raw)
                deliverable_name = str(raw_entry.get('deliverableName') or '').strip()
                cleaned_entries.append({
                    'projectId': str(raw_entry.get('projectId') or '').strip(),
                    'projectName': str(raw_entry.get('projectName') or '').strip(),
                    'deliverableName': deliverable_name or 'Untitled Deliverable',
                    'due': due_raw,
                    'dueValue': due_value.date() if due_value else None,
                    'statusText': str(raw_entry.get('statusText') or '').strip() or 'None',
                })

            if not cleaned_entries:
                return {'status': 'error', 'message': 'Select at least one deliverable to export.'}

            cleaned_entries.sort(key=lambda entry: (
                0 if entry['dueValue'] else 1,
                -(entry['dueValue'].toordinal()) if entry['dueValue'] else 0,
                entry['projectId'].lower(),
                entry['projectName'].lower(),
                entry['deliverableName'].lower(),
            ))

            file_path = payload.get('filePath')
            if isinstance(file_path, (list, tuple)):
                file_path = file_path[0] if file_path else ''
            file_path = str(file_path or '').strip()

            if not file_path:
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                selection = self.select_template_save_location(
                    default_dir=get_default_documents_dir(),
                    default_name=f"Deliverables_{timestamp}",
                    file_type='xlsx',
                )
                if selection.get('status') != 'success':
                    return selection
                file_path = str(selection.get('path') or '').strip()

            if not file_path:
                return {'status': 'cancelled'}
            if not file_path.lower().endswith('.xlsx'):
                file_path = f"{file_path}.xlsx"

            workbook = openpyxl.Workbook()
            worksheet = workbook.active
            worksheet.title = "Deliverables"

            headers = ("Project ID", "Project Name", "Deliverable", "Due Date", "Status")
            worksheet.append(headers)
            for header_cell in worksheet[1]:
                header_cell.font = openpyxl.styles.Font(bold=True)

            for row_index, entry in enumerate(cleaned_entries, start=2):
                worksheet.cell(row=row_index, column=1, value=entry['projectId'])
                worksheet.cell(row=row_index, column=2, value=entry['projectName'])
                worksheet.cell(row=row_index, column=3, value=entry['deliverableName'])
                due_cell = worksheet.cell(
                    row=row_index,
                    column=4,
                    value=entry['dueValue'] or entry['due'],
                )
                if entry['dueValue']:
                    due_cell.number_format = "mm/dd/yyyy"
                worksheet.cell(row=row_index, column=5, value=entry['statusText'])

            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions

            for column_index, width in {
                1: 14,
                2: 28,
                3: 32,
                4: 14,
                5: 18,
            }.items():
                column_letter = openpyxl.utils.get_column_letter(column_index)
                worksheet.column_dimensions[column_letter].width = width

            workbook.save(file_path)
            workbook.close()

            if sys.platform == "win32":
                os.startfile(file_path)
            else:
                self.open_path(file_path)

            return {'status': 'success', 'path': file_path, 'count': len(cleaned_entries)}
        except ImportError:
            return {'status': 'error', 'message': 'openpyxl not installed. Run: pip install openpyxl'}
        except Exception as e:
            logging.error(f"Error exporting deliverables to Excel: {e}")
            return {'status': 'error', 'message': str(e)}

    def save_dropped_email(self, upload, context=None):
        """Persists a dropped .msg/.eml payload and returns a normalized email reference."""
        try:
            payload = upload or {}
            if isinstance(payload, str):
                payload = json.loads(payload)
            if not isinstance(payload, dict):
                return {'status': 'error', 'message': 'Invalid upload payload.'}

            data_url = payload.get('dataUrl') or payload.get('data_url')
            if not data_url:
                return {'status': 'error', 'message': 'Missing email payload data.'}

            file_name = str(payload.get('name') or 'email.msg').strip() or 'email.msg'
            data, mime = self._decode_data_url(data_url)
            max_size_bytes = 20 * 1024 * 1024
            if len(data) > max_size_bytes:
                return {'status': 'error', 'message': 'Email attachment exceeds the 20 MB limit.'}

            ext = os.path.splitext(file_name)[1].lower()
            if ext not in ('.msg', '.eml'):
                mime_lower = (mime or '').lower()
                if 'rfc822' in mime_lower:
                    ext = '.eml'
                elif 'ms-outlook' in mime_lower:
                    ext = '.msg'
                else:
                    return {'status': 'error', 'message': 'Only .msg or .eml files are supported.'}

            context_data = context or {}
            if isinstance(context_data, str):
                try:
                    context_data = json.loads(context_data)
                except Exception:
                    context_data = {}

            project_hint = self._sanitize_email_path_component(
                context_data.get('projectId') or context_data.get('projectName'),
                'project'
            )
            deliverable_hint = self._sanitize_email_path_component(
                context_data.get('deliverableId') or context_data.get('deliverableName'),
                'deliverable'
            )
            file_stem = self._sanitize_email_path_component(
                os.path.splitext(os.path.basename(file_name))[0],
                'email'
            )

            email_root = self._get_email_links_root()
            target_dir = os.path.abspath(os.path.join(email_root, project_hint, deliverable_hint))
            if not self._is_within_directory(email_root, target_dir):
                return {'status': 'error', 'message': 'Invalid destination path.'}
            os.makedirs(target_dir, exist_ok=True)

            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            unique = uuid.uuid4().hex[:8]
            final_name = f"{timestamp}_{unique}_{file_stem}{ext}"
            output_path = os.path.abspath(os.path.join(target_dir, final_name))
            if not self._is_within_directory(email_root, output_path):
                return {'status': 'error', 'message': 'Invalid output path.'}

            with open(output_path, 'wb') as f:
                f.write(data)

            label = f"{file_stem}{ext}"
            saved_at = datetime.datetime.now().isoformat()
            return {
                'status': 'success',
                'emailRef': {
                    'raw': output_path,
                    'url': Path(output_path).as_uri(),
                    'label': label,
                    'source': 'saved-file',
                    'savedAt': saved_at,
                }
            }
        except Exception as e:
            logging.error(f"Error saving dropped email: {e}")
            return {'status': 'error', 'message': str(e)}

    def delete_saved_email(self, path):
        """Deletes a previously saved dropped email if it lives inside the managed email-links directory."""
        try:
            raw = str(path or '').strip()
            if not raw:
                return {'status': 'error', 'message': 'Path is required.'}

            target_path = os.path.abspath(self._coerce_local_email_path(raw))
            email_root = self._get_email_links_root()
            if not self._is_within_directory(email_root, target_path):
                return {'status': 'error', 'message': 'Path is outside managed email storage.'}

            if os.path.isdir(target_path):
                return {'status': 'error', 'message': 'Expected a file path, not a directory.'}

            deleted = False
            if os.path.exists(target_path):
                os.remove(target_path)
                deleted = True

                current = os.path.dirname(target_path)
                while current and current != email_root and self._is_within_directory(email_root, current):
                    if os.listdir(current):
                        break
                    os.rmdir(current)
                    current = os.path.dirname(current)

            return {'status': 'success', 'deleted': deleted}
        except Exception as e:
            logging.error(f"Error deleting saved email: {e}")
            return {'status': 'error', 'message': str(e)}

    def open_path(self, path):
        """Opens a path in the file explorer."""
        try:
            p = os.path.normpath(path)
            if sys.platform == "win32":
                if os.path.exists(p):
                    os.startfile(p)
                else:
                    parent = os.path.dirname(p)
                    if os.path.exists(parent):
                        os.startfile(parent)
                    else:
                        return {'status': 'error', 'message': 'Path and parent do not exist.'}
            else:
                subprocess.run(['open', p] if sys.platform ==
                               "darwin" else ['xdg-open', p])
            return {'status': 'success'}
        except Exception as e:
            logging.error(f"Error opening path: {e}")
            return {'status': 'error', 'message': str(e)}

    def reveal_path(self, path):
        """Reveals a file or folder in the system file explorer."""
        try:
            raw_path = str(path or '').strip()
            if not raw_path:
                return {'status': 'error', 'message': 'Path is required.'}

            normalized_path = os.path.normpath(raw_path)
            target_exists = os.path.exists(normalized_path)
            parent_path = os.path.dirname(normalized_path)
            parent_exists = bool(parent_path) and os.path.exists(parent_path)

            if sys.platform == "win32":
                if target_exists and os.path.isfile(normalized_path):
                    subprocess.run(
                        ["explorer.exe", "/select,", normalized_path],
                        check=False,
                    )
                    return {
                        'status': 'success',
                        'mode': 'select-file',
                        'path': normalized_path,
                    }
                if target_exists and os.path.isdir(normalized_path):
                    os.startfile(normalized_path)
                    return {
                        'status': 'success',
                        'mode': 'open-directory',
                        'path': normalized_path,
                    }
                if parent_exists:
                    os.startfile(parent_path)
                    return {
                        'status': 'success',
                        'mode': 'open-parent',
                        'path': parent_path,
                    }
                return {
                    'status': 'error',
                    'message': 'Path and parent do not exist.',
                }

            fallback_path = normalized_path if target_exists else parent_path
            if not fallback_path or not os.path.exists(fallback_path):
                return {'status': 'error', 'message': 'Path and parent do not exist.'}

            subprocess.run(
                ['open', fallback_path] if sys.platform == "darwin" else ['xdg-open', fallback_path],
                check=False,
            )
            return {'status': 'success', 'mode': 'open-path', 'path': fallback_path}
        except Exception as e:
            logging.error(f"Error revealing path: {e}")
            return {'status': 'error', 'message': str(e)}

    def open_directory_strict(self, path):
        """Opens an existing directory without falling back to parent paths."""
        try:
            raw_path = str(path or '').strip()
            if not raw_path:
                return {'status': 'error', 'message': 'Path is required.'}

            directory_path = os.path.normpath(raw_path)
            if not os.path.exists(directory_path):
                return {'status': 'error', 'message': 'Directory does not exist.'}
            if not os.path.isdir(directory_path):
                return {'status': 'error', 'message': 'Expected a directory path.'}

            if sys.platform == "win32":
                os.startfile(directory_path)
            else:
                subprocess.run(
                    ['open', directory_path] if sys.platform == "darwin" else ['xdg-open', directory_path]
                )
            return {'status': 'success'}
        except Exception as e:
            logging.error(f"Error opening directory: {e}")
            return {'status': 'error', 'message': str(e)}

    def get_local_project_copy_info(self, server_project_path):
        """Returns the expected local copy path for a server project and whether it exists."""
        try:
            normalized_server_path = os.path.normpath(str(server_project_path or '').strip())
            if not normalized_server_path:
                return {'status': 'error', 'message': 'Server project path is required.'}

            project_name = os.path.basename(normalized_server_path.rstrip('\\/'))
            if not project_name:
                return {'status': 'error', 'message': 'Invalid server project path.'}

            local_root = os.path.join(_get_windows_documents_dir(), 'Local Projects')
            local_project_path = os.path.normpath(os.path.join(local_root, project_name))
            exists = os.path.isdir(self._to_windows_extended_path(local_project_path))
            return {
                'status': 'success',
                'serverProjectPath': normalized_server_path,
                'projectName': project_name,
                'path': local_project_path,
                'exists': exists,
            }
        except Exception as e:
            logging.error(f"Error reading local project copy info: {e}")
            return {'status': 'error', 'message': str(e)}

    def open_timesheets_folder(self):
        """Opens the folder where timesheets.json is stored."""
        try:
            folder = os.path.dirname(TIMESHEETS_FILE)
            return self.open_path(folder)
        except Exception as e:
            logging.error(f"Error opening timesheets folder: {e}")
            return {'status': 'error', 'message': str(e)}

    def get_timesheets_info(self):
        """Returns details about the timesheets file path and state."""
        try:
            path = os.path.normpath(TIMESHEETS_FILE)
            exists = os.path.exists(path)
            size = os.path.getsize(path) if exists else 0
            mtime = os.path.getmtime(path) if exists else None
            mtime_iso = None
            if mtime is not None:
                mtime_iso = datetime.datetime.fromtimestamp(mtime).isoformat()
            return {
                'status': 'success',
                'path': path,
                'exists': exists,
                'size': size,
                'modified': mtime_iso
            }
        except Exception as e:
            logging.error(f"Error reading timesheets info: {e}")
            return {'status': 'error', 'message': str(e)}

    def write_timesheets_test_file(self):
        """Writes a small test file next to timesheets.json to verify writes."""
        try:
            folder = os.path.dirname(TIMESHEETS_FILE)
            os.makedirs(folder, exist_ok=True)
            test_path = os.path.join(folder, "timesheets_write_test.txt")
            stamp = datetime.datetime.now().isoformat()
            with open(test_path, "w", encoding="utf-8") as f:
                f.write(f"write_test {stamp}\n")
            exists = os.path.exists(test_path)
            if not exists:
                return {
                    'status': 'error',
                    'message': 'Test file write did not persist.',
                    'path': test_path
                }
            return {'status': 'success', 'path': test_path, 'exists': True}
        except Exception as e:
            logging.error(f"Error writing timesheets test file: {e}")
            return {'status': 'error', 'message': str(e)}

    def create_folder(self, path):
        """Creates a directory."""
        if not path:
            return {'status': 'error', 'message': 'Path cannot be empty.'}
        try:
            p = os.path.normpath(path)
            os.makedirs(p, exist_ok=True)
            return {'status': 'success'}
        except Exception as e:
            logging.error(f"Error creating folder: {e}")
            return {'status': 'error', 'message': str(e)}

    def _get_copy_project_disciplines(self, settings):
        """Normalize settings discipline value into a non-empty list."""
        if not isinstance(settings, dict):
            return ['Electrical']

        raw_value = settings.get('discipline', ['Electrical'])
        if isinstance(raw_value, str):
            candidates = [
                part.strip() for part in re.split(r'[,/;]+', raw_value) if part and part.strip()
            ]
        elif isinstance(raw_value, (list, tuple, set)):
            candidates = [str(item or '').strip() for item in raw_value if str(item or '').strip()]
        else:
            candidates = []

        if not candidates:
            return ['Electrical']

        canonical_map = {
            'electrical': 'Electrical',
            'mechanical': 'Mechanical',
            'plumbing': 'Plumbing',
            'arch': 'Arch',
            'architecture': 'Arch',
            'general': 'Electrical',
        }
        result = []
        seen = set()
        for item in candidates:
            canonical = canonical_map.get(item.lower(), item)
            key = canonical.lower()
            if key in seen:
                continue
            seen.add(key)
            result.append(canonical)

        return result or ['Electrical']

    def _to_windows_extended_path(self, path):
        """Prefix Windows absolute paths so copy operations can exceed MAX_PATH."""
        normalized = os.path.normpath(str(path or ''))
        if sys.platform != "win32" or not normalized:
            return normalized
        if normalized.startswith('\\\\?\\'):
            return normalized
        if normalized.startswith('\\\\'):
            unc_path = normalized.lstrip('\\')
            return f"\\\\?\\UNC\\{unc_path}"
        if re.match(r'^[A-Za-z]:\\', normalized):
            return f"\\\\?\\{normalized}"
        return normalized

    def _copy_folder_contents(self, source_folder, destination_folder):
        """Recursively copy a folder with per-file failure tracking."""
        source_display_root = os.path.normpath(source_folder)
        destination_display_root = os.path.normpath(destination_folder)
        source_copy_root = self._to_windows_extended_path(source_display_root)
        destination_copy_root = self._to_windows_extended_path(destination_display_root)

        copied_file_count = 0
        failed_files = []

        os.makedirs(destination_copy_root, exist_ok=True)

        for current_root, child_dirs, child_files in os.walk(source_copy_root):
            relative_root = os.path.relpath(current_root, source_copy_root)
            if relative_root in ('.', os.curdir):
                relative_root = ''

            destination_root = destination_copy_root if not relative_root else os.path.join(
                destination_copy_root, relative_root
            )
            source_display = source_display_root if not relative_root else os.path.join(
                source_display_root, relative_root
            )
            destination_display = destination_display_root if not relative_root else os.path.join(
                destination_display_root, relative_root
            )

            try:
                os.makedirs(destination_root, exist_ok=True)
            except Exception as e:
                failed_files.append({
                    'source': source_display,
                    'destination': destination_display,
                    'error': str(e),
                })
                continue

            for child_dir in child_dirs:
                source_dir = os.path.join(source_display, child_dir)
                destination_dir = os.path.join(destination_display, child_dir)
                try:
                    os.makedirs(os.path.join(destination_root, child_dir), exist_ok=True)
                except Exception as e:
                    failed_files.append({
                        'source': source_dir,
                        'destination': destination_dir,
                        'error': str(e),
                    })

            for child_file in child_files:
                source_path = os.path.join(current_root, child_file)
                destination_path = os.path.join(destination_root, child_file)
                source_display_path = os.path.join(source_display, child_file)
                destination_display_path = os.path.join(destination_display, child_file)
                try:
                    shutil.copy2(source_path, destination_path)
                    copied_file_count += 1
                except Exception as e:
                    failed_files.append({
                        'source': source_display_path,
                        'destination': destination_display_path,
                        'error': str(e),
                    })

        return {
            'copiedFileCount': copied_file_count,
            'failedFiles': failed_files,
        }

    def _resolve_copy_project_source_path(self, server_project_path, launch_context, settings):
        raw_server_path = str(server_project_path or '').strip()
        if raw_server_path:
            return {
                'status': 'success',
                'path': os.path.normpath(raw_server_path),
                'resolvedFromWorkroom': False,
                'resolutionMode': 'manual_selection',
                'workroomProjectPath': '',
            }

        context = self._resolve_workroom_context(settings, launch_context)
        source = str(context.get('source') or '').strip().lower()
        workroom_project_path = str(context.get('project_path') or '').strip()

        if source != 'workroom':
            return {
                'status': 'error',
                'code': 'server_project_path_required',
                'message': 'Server project path is required.',
                'resolvedFromWorkroom': False,
                'resolutionMode': 'missing_server_path',
                'workroomProjectPath': '',
            }

        if not workroom_project_path:
            return {
                'status': 'error',
                'code': 'manual_selection_required',
                'message': 'Could not auto-resolve project folder from Project Workroom. Please select it manually.',
                'resolvedFromWorkroom': True,
                'resolutionMode': 'workroom_missing_project_path',
                'workroomProjectPath': '',
            }

        resolved_root = self._find_workroom_project_root_by_id(workroom_project_path)
        if not resolved_root:
            return {
                'status': 'error',
                'code': 'manual_selection_required',
                'message': 'Could not auto-resolve project folder from Project Workroom. Please select it manually.',
                'resolvedFromWorkroom': True,
                'resolutionMode': 'project_id_ancestor_not_found',
                'workroomProjectPath': workroom_project_path,
            }

        return {
            'status': 'success',
            'path': os.path.normpath(resolved_root),
            'resolvedFromWorkroom': True,
            'resolutionMode': 'project_id_ancestor',
            'workroomProjectPath': workroom_project_path,
        }

    def _get_backup_drawings_folder_names(self, settings):
        required_folders = []
        seen_required = set()
        for folder_name in [*self._get_copy_project_disciplines(settings), 'Xrefs']:
            clean_name = str(folder_name or '').strip()
            if not clean_name:
                continue
            key = clean_name.lower()
            if key in seen_required:
                continue
            seen_required.add(key)
            required_folders.append(clean_name)
        return required_folders

    def _build_backup_drawings_timestamp(self, now=None):
        return (now or datetime.datetime.now()).strftime("%Y%m%d_%H%M%S")

    def _reserve_unique_archive_folder(self, archive_root, timestamp=None):
        archive_root = os.path.normpath(str(archive_root or '').strip())
        if not archive_root:
            raise ValueError("Archive root path is required.")

        archive_root_copy_path = self._to_windows_extended_path(archive_root)
        os.makedirs(archive_root_copy_path, exist_ok=True)

        base_name = str(timestamp or self._build_backup_drawings_timestamp()).strip()
        if not base_name:
            raise ValueError("Archive folder name is required.")

        candidate_index = 1
        while True:
            suffix = '' if candidate_index == 1 else f" ({candidate_index})"
            candidate_path = os.path.join(archive_root, f"{base_name}{suffix}")
            candidate_copy_path = self._to_windows_extended_path(candidate_path)
            try:
                os.makedirs(candidate_copy_path, exist_ok=False)
                return os.path.normpath(candidate_path)
            except FileExistsError:
                candidate_index += 1

    def backup_project_drawings(self, project_root_path=None, launch_context=None):
        """Copy configured discipline folders and Xrefs into Archive\\<timestamp>."""
        try:
            settings = self.get_user_settings()
            source_resolution = self._resolve_copy_project_source_path(
                project_root_path, launch_context, settings
            )
            if source_resolution.get('status') != 'success':
                return source_resolution

            normalized_project_root = source_resolution.get('path') or ''
            resolved_from_workroom = bool(source_resolution.get('resolvedFromWorkroom'))
            resolution_mode = str(source_resolution.get('resolutionMode') or '').strip()
            workroom_project_path = str(source_resolution.get('workroomProjectPath') or '').strip()

            if not os.path.isdir(self._to_windows_extended_path(normalized_project_root)):
                if resolved_from_workroom:
                    return {
                        'status': 'error',
                        'code': 'manual_selection_required',
                        'message': 'Could not auto-resolve project folder from Project Workroom. Please select it manually.',
                        'resolvedFromWorkroom': True,
                        'resolvedProjectRootPath': normalized_project_root,
                        'resolutionMode': 'project_id_ancestor_not_accessible',
                        'workroomProjectPath': workroom_project_path,
                    }
                return {'status': 'error', 'message': 'Project root path does not exist.'}

            archive_root_path = os.path.normpath(os.path.join(normalized_project_root, 'Archive'))
            archive_path = self._reserve_unique_archive_folder(archive_root_path)
            required_folders = self._get_backup_drawings_folder_names(settings)

            copied_folders = []
            missing_source_folders = []
            copied_file_count = 0
            failed_files = []

            for folder_name in required_folders:
                source_folder = os.path.join(normalized_project_root, folder_name)
                destination_folder = os.path.join(archive_path, folder_name)
                if os.path.isdir(self._to_windows_extended_path(source_folder)):
                    copy_result = self._copy_folder_contents(source_folder, destination_folder)
                    copied_folders.append(folder_name)
                    copied_file_count += int(copy_result.get('copiedFileCount', 0) or 0)
                    failed_files.extend(copy_result.get('failedFiles', []))
                else:
                    missing_source_folders.append(folder_name)

            if not copied_folders:
                shutil.rmtree(self._to_windows_extended_path(archive_path), ignore_errors=True)
                return {
                    'status': 'error',
                    'code': 'nothing_to_backup',
                    'message': 'No configured discipline folders or Xrefs folder were found to back up.',
                    'projectRootPath': normalized_project_root,
                    'resolvedProjectRootPath': normalized_project_root,
                    'resolvedFromWorkroom': resolved_from_workroom,
                    'resolutionMode': resolution_mode or 'manual_selection',
                    'workroomProjectPath': workroom_project_path,
                    'archiveRootPath': archive_root_path,
                    'archivePath': archive_path,
                    'disciplines': self._get_copy_project_disciplines(settings),
                    'copiedFolders': [],
                    'missingSourceFolders': missing_source_folders,
                    'copiedFileCount': 0,
                    'failedFileCount': 0,
                    'failedFiles': [],
                }

            failed_file_count = len(failed_files)
            has_warnings = bool(missing_source_folders or failed_file_count)
            message = 'Drawing backup created.'
            if has_warnings:
                message = 'Drawing backup created with warnings.'

            return {
                'status': 'success',
                'message': message,
                'projectRootPath': normalized_project_root,
                'resolvedProjectRootPath': normalized_project_root,
                'resolvedFromWorkroom': resolved_from_workroom,
                'resolutionMode': resolution_mode or 'manual_selection',
                'workroomProjectPath': workroom_project_path,
                'archiveRootPath': archive_root_path,
                'archivePath': archive_path,
                'disciplines': self._get_copy_project_disciplines(settings),
                'copiedFolders': copied_folders,
                'missingSourceFolders': missing_source_folders,
                'copiedFileCount': copied_file_count,
                'failedFileCount': failed_file_count,
                'failedFiles': failed_files,
            }
        except Exception as e:
            logging.error(f"Error backing up project drawings: {e}")
            return {'status': 'error', 'message': str(e)}

    def copy_project_locally(self, server_project_path=None, launch_context=None):
        """Copy key project folders from server to local Documents\\Local Projects."""
        try:
            settings = self.get_user_settings()
            source_resolution = self._resolve_copy_project_source_path(
                server_project_path, launch_context, settings
            )
            if source_resolution.get('status') != 'success':
                return source_resolution

            normalized_server_path = source_resolution.get('path') or ''
            resolved_from_workroom = bool(source_resolution.get('resolvedFromWorkroom'))
            resolution_mode = str(source_resolution.get('resolutionMode') or '').strip()
            workroom_project_path = str(source_resolution.get('workroomProjectPath') or '').strip()

            if not os.path.isdir(self._to_windows_extended_path(normalized_server_path)):
                if resolved_from_workroom:
                    return {
                        'status': 'error',
                        'code': 'manual_selection_required',
                        'message': 'Could not auto-resolve project folder from Project Workroom. Please select it manually.',
                        'resolvedFromWorkroom': True,
                        'resolvedServerProjectPath': normalized_server_path,
                        'resolutionMode': 'project_id_ancestor_not_accessible',
                        'workroomProjectPath': workroom_project_path,
                    }
                return {'status': 'error', 'message': 'Server project path does not exist.'}

            project_name = os.path.basename(normalized_server_path.rstrip('\\/'))
            if not project_name:
                return {'status': 'error', 'message': 'Invalid server project path.'}

            local_root = os.path.join(_get_windows_documents_dir(), 'Local Projects')
            os.makedirs(self._to_windows_extended_path(local_root), exist_ok=True)
            local_project_path = os.path.normpath(os.path.join(local_root, project_name))
            local_project_copy_path = self._to_windows_extended_path(local_project_path)

            if os.path.exists(local_project_copy_path):
                if os.path.isdir(local_project_copy_path):
                    return {
                        'status': 'error',
                        'code': 'local_project_exists',
                        'message': f'Local project already exists: {local_project_path}',
                        'serverProjectPath': normalized_server_path,
                        'resolvedServerProjectPath': normalized_server_path,
                        'resolvedFromWorkroom': resolved_from_workroom,
                        'resolutionMode': resolution_mode or 'manual_selection',
                        'workroomProjectPath': workroom_project_path,
                        'localProjectPath': local_project_path,
                        'projectName': project_name,
                    }
                return {
                    'status': 'error',
                    'message': f'Local project path is unavailable: {local_project_path}'
                }

            disciplines = self._get_copy_project_disciplines(settings)

            required_folders = []
            seen_required = set()
            for folder_name in ['Arch', *disciplines, 'Xrefs', 'Documents', 'RFI']:
                clean_name = str(folder_name or '').strip()
                if not clean_name:
                    continue
                key = clean_name.lower()
                if key in seen_required:
                    continue
                seen_required.add(key)
                required_folders.append(clean_name)

            os.makedirs(local_project_copy_path, exist_ok=False)
            for folder_name in required_folders:
                folder_path = os.path.join(local_project_path, folder_name)
                os.makedirs(self._to_windows_extended_path(folder_path), exist_ok=True)

            copied_folders = []
            missing_server_folders = []
            copied_file_count = 0
            failed_files = []
            for folder_name in required_folders:
                source_folder = os.path.join(normalized_server_path, folder_name)
                destination_folder = os.path.join(local_project_path, folder_name)
                if os.path.isdir(self._to_windows_extended_path(source_folder)):
                    copy_result = self._copy_folder_contents(source_folder, destination_folder)
                    copied_folders.append(folder_name)
                    copied_file_count += int(copy_result.get('copiedFileCount', 0) or 0)
                    failed_files.extend(copy_result.get('failedFiles', []))
                else:
                    missing_server_folders.append(folder_name)

            failed_file_count = len(failed_files)
            copy_warnings = []
            message = 'Project copied locally.'
            if failed_file_count:
                message = f'Project copied locally with warnings: {failed_file_count} file(s) failed.'
                copy_warnings.append(message)

            return {
                'status': 'success',
                'message': message,
                'serverProjectPath': normalized_server_path,
                'resolvedServerProjectPath': normalized_server_path,
                'resolvedFromWorkroom': resolved_from_workroom,
                'resolutionMode': resolution_mode or 'manual_selection',
                'workroomProjectPath': workroom_project_path,
                'localProjectPath': local_project_path,
                'projectName': project_name,
                'disciplines': disciplines,
                'copiedFolders': copied_folders,
                'missingServerFolders': missing_server_folders,
                'copyWarnings': copy_warnings,
                'copiedFileCount': copied_file_count,
                'failedFileCount': failed_file_count,
                'failedFiles': failed_files,
            }
        except Exception as e:
            logging.error(f"Error copying project locally: {e}")
            return {'status': 'error', 'message': str(e)}

    def get_wire_sizer_url(self):
        """Return a URL or relative path to the Wire Sizer build output."""
        candidate = BASE_DIR / "WireSizerApplication" / "dist" / "index.html"
        try:
            if candidate.exists():
                try:
                    rel_path = candidate.relative_to(BASE_DIR)
                    return {
                        'status': 'success',
                        'url': rel_path.as_posix()
                    }
                except ValueError:
                    return {
                        'status': 'success',
                        'url': candidate.resolve().as_uri()
                    }
        except Exception as e:
            logging.error(f"Error resolving Wire Sizer path: {e}")
            return {
                'status': 'error',
                'message': 'Failed to resolve Wire Sizer path.'
            }
        return {
            'status': 'error',
            'message': 'Wire Sizer build not found. Run npm install and npm run build in WireSizerApplication.'
        }

    # --- Circuit Breaker AI (Panel Schedule) ---

    def _find_free_port(self):
        """Find an available TCP port."""
        import socket
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind(('127.0.0.1', 0))
            return s.getsockname()[1]

    def start_circuit_breaker_server(self):
        """Start the CircuitBreakerAI FastAPI server as a subprocess."""
        # If already running, return the URL
        if hasattr(self, '_cb_process') and self._cb_process and self._cb_process.poll() is None:
            return {'status': 'success', 'url': f'http://127.0.0.1:{self._cb_port}'}

        server_script = BASE_DIR / "CircuitBreakerAI" / "server.py"
        if not server_script.exists():
            return {'status': 'error', 'message': 'CircuitBreakerAI server.py not found.'}

        port = self._find_free_port()
        self._cb_port = port

        python_exe = sys.executable

        # Build environment with API key - prefer user settings, then env vars
        env = os.environ.copy()
        settings = self.get_user_settings()
        api_key = settings.get('apiKey', '').strip()
        if not api_key:
            api_key = os.environ.get('GEMINI_API_KEY', os.environ.get('GOOGLE_API_KEY', ''))
        if api_key:
            env['GEMINI_API_KEY'] = api_key

        try:
            startupinfo = None
            creationflags = 0
            if sys.platform == "win32":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = 0  # SW_HIDE
                creationflags = subprocess.CREATE_NO_WINDOW

            self._cb_process = subprocess.Popen(
                [python_exe, str(server_script), "--port", str(port)],
                cwd=str(server_script.parent),
                startupinfo=startupinfo,
                creationflags=creationflags,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                env=env,
            )

            # Wait for server to be ready (poll the port)
            import socket
            for _ in range(100):  # 10 seconds max
                try:
                    with socket.create_connection(('127.0.0.1', port), timeout=0.1):
                        return {'status': 'success', 'url': f'http://127.0.0.1:{port}'}
                except (ConnectionRefusedError, OSError):
                    time.sleep(0.1)

            return {'status': 'error', 'message': 'CircuitBreakerAI server failed to start.'}
        except Exception as e:
            logging.error(f"Error starting CircuitBreakerAI server: {e}")
            return {'status': 'error', 'message': str(e)}

    def stop_circuit_breaker_server(self):
        """Stop the CircuitBreakerAI server subprocess."""
        if hasattr(self, '_cb_process') and self._cb_process:
            try:
                self._cb_process.terminate()
                self._cb_process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                self._cb_process.kill()
            except Exception:
                pass
            self._cb_process = None
        return {'status': 'success'}

    def save_circuit_breaker_schedule(self, suggested_name=None):
        """Save the CircuitBreakerAI schedule to a user-selected path."""
        try:
            if isinstance(suggested_name, dict):
                suggested_name = suggested_name.get('suggestedName') or suggested_name.get('suggested_name')

            output_path = BASE_DIR / "CircuitBreakerAI" / "ElectricalPanels" / "Filled_Panel_Schedules.xlsx"
            if not output_path.exists():
                return {'status': 'error', 'message': 'No schedule file found. Please generate a panel first.'}

            default_name = str(suggested_name or 'Panel_Schedule').strip() or 'Panel_Schedule'
            selection = self.select_template_save_location(default_name=default_name, file_type='xlsx')
            if selection.get('status') != 'success':
                return selection

            dest_path = selection.get('path')
            if not dest_path:
                return {'status': 'cancelled', 'path': None}

            if not dest_path.lower().endswith('.xlsx'):
                dest_path = f"{dest_path}.xlsx"

            dest_dir = os.path.dirname(dest_path)
            if dest_dir and not os.path.isdir(dest_dir):
                os.makedirs(dest_dir, exist_ok=True)

            shutil.copy2(str(output_path), dest_path)

            try:
                os.remove(output_path)
            except Exception as cleanup_error:
                logging.warning(f"Could not remove CircuitBreakerAI temp schedule: {cleanup_error}")

            return {'status': 'success', 'path': dest_path}
        except Exception as e:
            logging.error(f"Error saving CircuitBreakerAI schedule: {e}")
            return {'status': 'error', 'message': str(e)}

    def run_panel_schedule_background(self, payload):
        """Runs Panel Schedule AI in a background thread and notifies the UI."""
        if getattr(self, "_panel_schedule_running", False):
            return {'status': 'error', 'message': 'Panel Schedule AI is already running.'}

        self._panel_schedule_running = True
        thread = threading.Thread(
            target=self._panel_schedule_worker,
            args=(payload,),
            daemon=True
        )
        self._panel_schedule_thread = thread
        thread.start()
        return {'status': 'started'}

    def _panel_schedule_worker(self, payload):
        window = webview.windows[0]
        result = {'status': 'error', 'message': 'Panel Schedule AI failed.'}
        try:
            result = self._process_panel_schedule_payload(payload)
        except Exception as e:
            logging.error(f"Panel Schedule AI error: {e}")
            result = {'status': 'error', 'message': str(e)}
        finally:
            self._panel_schedule_running = False

        try:
            js_payload = json.dumps(result)
            window.evaluate_js(
                f"window.handlePanelScheduleResult({js_payload})"
            )
        except Exception as e:
            logging.error(f"Failed to notify Panel Schedule AI result: {e}")

    def _normalize_panel_schedule_paths(self, value):
        values = value if isinstance(value, (list, tuple)) else [value]
        normalized = []
        for raw in values:
            path = str(raw or '').strip()
            if path:
                normalized.append(path)
        return normalized

    def _get_panel_schedule_path_value(self, data, plural_keys, singular_keys):
        if not isinstance(data, dict):
            return None
        for key in plural_keys:
            if key in data:
                return data.get(key)
        for key in singular_keys:
            if key in data:
                return data.get(key)
        return []

    def _decode_data_url(self, data_url):
        if not data_url or "base64," not in data_url:
            raise ValueError("Invalid image data.")
        header, encoded = data_url.split("base64,", 1)
        mime = ""
        if header.startswith("data:"):
            mime = header.split(";", 1)[0].replace("data:", "").strip()
        return base64.b64decode(encoded), mime

    def _get_email_links_root(self):
        root = os.path.join(os.path.dirname(TASKS_FILE), "email-links")
        os.makedirs(root, exist_ok=True)
        return os.path.abspath(root)

    def _is_within_directory(self, base_path, target_path):
        try:
            base = os.path.abspath(base_path)
            target = os.path.abspath(target_path)
            return os.path.commonpath([base, target]) == base
        except Exception:
            return False

    def _sanitize_email_path_component(self, value, fallback):
        text = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(value or "").strip())
        text = re.sub(r"\s+", " ", text).strip().strip(".")
        text = text[:80]
        return text or fallback

    def _coerce_local_file_path(self, raw):
        value = str(raw or "").strip()
        if not value:
            return ""
        if value.lower().startswith("file://"):
            parsed = urlparse(value)
            path = unquote(parsed.path or "")
            if parsed.netloc:
                unc_path = path.replace("/", "\\")
                if unc_path and not unc_path.startswith("\\"):
                    unc_path = "\\" + unc_path
                return os.path.normpath(f"\\\\{parsed.netloc}{unc_path}")
            if sys.platform == "win32" and re.match(r"^/[A-Za-z]:", path):
                path = path[1:]
            return os.path.normpath(path)
        return os.path.normpath(value)

    def _coerce_local_email_path(self, raw):
        return self._coerce_local_file_path(raw)

    def _panel_schedule_save_uploads(self, uploads, label):
        temp_paths = []
        resolved_paths = []
        for upload in uploads or []:
            if not isinstance(upload, dict):
                continue
            data_url = upload.get("dataUrl") or upload.get("data_url")
            if not data_url:
                continue
            file_name = str(upload.get("name") or f"{label}.jpg")
            try:
                raw, mime = self._decode_data_url(data_url)
            except Exception:
                continue
            suffix = os.path.splitext(file_name)[1] or ""
            if not suffix:
                if mime.endswith("png"):
                    suffix = ".png"
                elif mime.endswith("jpeg") or mime.endswith("jpg"):
                    suffix = ".jpg"
                elif mime.endswith("heic"):
                    suffix = ".heic"
                else:
                    suffix = ".img"
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
            tmp.write(raw)
            tmp.close()
            temp_paths.append(tmp.name)
            resolved_paths.append(tmp.name)
        return resolved_paths, temp_paths

    def _resolve_panel_schedule_api_key(self):
        settings = self.get_user_settings()
        api_key = settings.get('apiKey', '').strip()
        if not api_key:
            api_key = os.environ.get('GEMINI_API_KEY') or os.environ.get('GOOGLE_API_KEY') or ''
        return (api_key or '').strip()

    def _format_panel_schedule_ai_error(self, exc):
        msg = str(exc)
        lower = msg.lower()
        if ("api key expired" in lower or
                "api_key_invalid" in lower or
                "invalid api key" in lower):
            return ('Your Google API key is expired/invalid. '
                    'Create a new key in Google AI Studio, update your settings, then try again.')
        if "model" in lower and ("not found" in lower or "does not exist" in lower):
            return 'AI model not available. The Gemini 3 Flash model may not be accessible with your API key.'
        if "quota" in lower or "rate limit" in lower:
            return 'API rate limit exceeded. Please wait a moment and try again.'
        return msg

    def _normalize_panel_schedule_extension(self, ext_value):
        ext = str(ext_value or "").strip().lower()
        if not ext:
            return ""
        if not ext.startswith("."):
            ext = f".{ext}"
        return ext

    def _panel_schedule_excel_requirement_message(self):
        return "Editing .xls requires Microsoft Excel installed on this machine."

    def _convert_excel_workbook(self, source_path, dest_path, target_extension):
        target_extension = self._normalize_panel_schedule_extension(target_extension)
        if target_extension not in (".xlsx", ".xls"):
            raise ValueError("Excel conversion target must be .xlsx or .xls.")

        source_abs = os.path.abspath(str(source_path))
        dest_abs = os.path.abspath(str(dest_path))
        if not os.path.exists(source_abs):
            raise FileNotFoundError(f"Workbook not found: {source_abs}")

        try:
            import win32com.client
        except Exception as exc:
            raise RuntimeError(self._panel_schedule_excel_requirement_message()) from exc

        excel = None
        workbook = None
        file_format = 56 if target_extension == ".xls" else 51

        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            workbook = excel.Workbooks.Open(source_abs)
            workbook.SaveAs(dest_abs, FileFormat=file_format)
        except Exception as exc:
            raise RuntimeError(
                f"{self._panel_schedule_excel_requirement_message()} "
                f"Excel conversion failed: {exc}"
            ) from exc
        finally:
            if workbook is not None:
                try:
                    workbook.Close(False)
                except Exception:
                    pass
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass

    def _update_panel_schedule_workbook(self, panel_data, output_path):
        output_extension = self._normalize_panel_schedule_extension(
            os.path.splitext(str(output_path))[1]
        )
        if output_extension == ".xlsx":
            return cb_update_excel_workbook(panel_data, output_path)
        if output_extension != ".xls":
            raise ValueError("Panel schedule must be an .xlsx or .xls file.")

        temp_xlsx = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        temp_xlsx_path = temp_xlsx.name
        temp_xlsx.close()

        try:
            self._convert_excel_workbook(output_path, temp_xlsx_path, ".xlsx")
            sheet_name = cb_update_excel_workbook(panel_data, temp_xlsx_path)
            self._convert_excel_workbook(temp_xlsx_path, output_path, ".xls")
            return sheet_name
        finally:
            try:
                os.remove(temp_xlsx_path)
            except Exception:
                pass

    def _build_panel_schedule_prompt(self, panel_name, num_breaker_imgs, num_dir_imgs):
        return f"""
Analyze these electrical panel photos for Panel: {panel_name}.

You are provided with {num_breaker_imgs} images of the CIRCUIT BREAKERS (first {num_breaker_imgs} images)
and {num_dir_imgs} images of the CIRCUIT DIRECTORY (last {num_dir_imgs} images).

IMAGE INTERPRETATION RULES
- All breaker images belong to the SAME panel.
- All directory images belong to the SAME panel.
- Breaker coverage may be split across multiple photos of the same panel, such as upper half, middle, and bottom half images.
- Directory coverage may be split across multiple photos, such as circuits 1-42 on one image and 43-84 on another.
- Merge partial and overlapping views into one complete understanding of the panel.
- Do not treat split photos as separate panels.
- Reconcile overlapping information by using visible circuit numbers, breaker positions, and the clearest label text.
- Breaker images come first. Directory images come last.
- Do not assume each image shows the entire panel or a complete directory by itself.

TASK 1: HEADER
- Extract Voltage, Bus Rating, Wire, Phase, Mounting, Enclosure.
- Look at the directory images or labels on the panel across all images.

TASK 2: CIRCUITS & POLES
- Identify every breaker visible across the Breaker Images.
- Use visible breaker positions and circuit numbering to merge split sections from multiple photos.
- CRITICAL: Determine the 'poles' (1, 2, or 3).
    - A 1-pole breaker takes up 1 circuit space.
    - A 2-pole breaker has a tied handle and takes up 2 vertical circuit spaces (e.g. 1 & 3).
    - A 3-pole breaker has a tied handle and takes up 3 vertical circuit spaces (e.g. 1, 3, & 5).
- Provide the circuit_number as the TOP-most circuit number the breaker occupies.
- Extract Amperage and Load Description from the labels (Breaker Images) or circuit directory (Directory Images).
- Resolve ditto marks (") if seen in the directory.
- If the same circuit appears in overlapping photos, combine the evidence and keep one final record.

TASK 3: LOAD TYPES
- LIGHTING -> 'C', RECEPTACLES -> 'G', MOTORS/HVAC -> 'M', KITCHEN -> 'K', DEDICATED -> 'D'
""".strip()

    def _analyze_panel_schedule_images(self, panel_name, breaker_paths, directory_paths):
        api_key = self._resolve_panel_schedule_api_key()
        if not api_key:
            raise RuntimeError(
                'AI API key is not configured. Please add it in Settings or set GOOGLE_API_KEY in your .env.'
            )

        self._ensure_aiohttp()
        client = genai.Client(api_key=api_key)

        num_breaker_imgs = len(breaker_paths)
        num_dir_imgs = len(directory_paths)
        prompt = self._build_panel_schedule_prompt(
            panel_name, num_breaker_imgs, num_dir_imgs
        )

        gemini_images = []
        try:
            for path in breaker_paths + directory_paths:
                if not os.path.exists(path):
                    raise ValueError(f"Image not found: {path}")
                gemini_images.append(PILImage.open(path))

            cb_enforce_rate_limit()

            response = client.models.generate_content(
                model="gemini-3-flash-preview",
                contents=[prompt, *gemini_images],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=PanelData
                ),
            )
        finally:
            for img in gemini_images:
                try:
                    img.close()
                except Exception:
                    pass

        panel_data = response.parsed
        if panel_data is None:
            raw_text = response.text or ""
            if not raw_text.strip():
                raise RuntimeError("AI returned an empty response. Please try again.")
            data = json.loads(raw_text)
            if hasattr(PanelData, "model_validate"):
                panel_data = PanelData.model_validate(data)
            else:
                panel_data = PanelData.parse_obj(data)

        panel_data.panel_name = panel_name
        return panel_data

    def _process_panel_schedule_payload(self, payload):
        data = payload or {}
        if isinstance(data, str):
            try:
                data = json.loads(data)
            except Exception:
                data = {}

        output_mode = str(
            data.get('outputMode') or data.get('output_mode') or 'new'
        ).strip().lower()
        if output_mode not in ("new", "existing"):
            output_mode = "new"

        output_path = (
            data.get('outputPath')
            or data.get('output_path')
            or (data.get('newOutputPath') if output_mode == "new" else data.get('existingOutputPath'))
        )
        output_path = str(output_path or '').strip()
        if not output_path:
            raise ValueError("Panel schedule file is required.")

        output_path = os.path.normpath(output_path)
        output_extension_hint = self._normalize_panel_schedule_extension(
            data.get('outputExtension') or data.get('output_extension')
        )
        output_extension = self._normalize_panel_schedule_extension(
            os.path.splitext(output_path)[1]
        )

        created_workbook = False

        if output_mode == "new":
            if not output_extension:
                output_extension = output_extension_hint if output_extension_hint in (".xlsx", ".xls") else ".xlsx"
                output_path = f"{output_path}{output_extension}"
            if output_extension not in (".xlsx", ".xls"):
                raise ValueError("Panel schedule must be an .xlsx or .xls file.")
            if os.path.exists(output_path):
                raise ValueError("The selected file already exists. Choose a new name or add to the existing schedule.")
            if not CB_TEMPLATE_PATH.exists():
                raise ValueError("Panel schedule template not found.")
            dest_dir = os.path.dirname(output_path)
            if dest_dir:
                os.makedirs(dest_dir, exist_ok=True)
            if output_extension == ".xlsx":
                shutil.copy2(str(CB_TEMPLATE_PATH), output_path)
            else:
                self._convert_excel_workbook(str(CB_TEMPLATE_PATH), output_path, ".xls")
            created_workbook = True
        else:
            if output_extension not in (".xlsx", ".xls"):
                raise ValueError("Panel schedule must be an .xlsx or .xls file.")
            if not os.path.exists(output_path):
                raise ValueError("Selected panel schedule was not found.")

        panels_payload = data.get('panels')
        panel_requests = []
        if isinstance(panels_payload, list):
            for index, raw_panel in enumerate(panels_payload):
                if not isinstance(raw_panel, dict):
                    continue
                panel_name = str(
                    raw_panel.get('panelName') or raw_panel.get('panel_name') or ''
                ).strip()
                if not panel_name:
                    panel_name = f"PANEL {index + 1}"
                panel_id = str(
                    raw_panel.get('panelId') or raw_panel.get('panel_id') or f"panel_{index + 1}"
                ).strip() or f"panel_{index + 1}"
                panel_requests.append({
                    'panel_id': panel_id,
                    'panel_name': panel_name,
                    'breaker_paths': self._normalize_panel_schedule_paths(
                        self._get_panel_schedule_path_value(
                            raw_panel,
                            ('breakerPaths', 'breaker_paths'),
                            ('breakerPath', 'breaker_path'),
                        )
                    ),
                    'directory_paths': self._normalize_panel_schedule_paths(
                        self._get_panel_schedule_path_value(
                            raw_panel,
                            ('directoryPaths', 'directory_paths'),
                            ('directoryPath', 'directory_path'),
                        )
                    ),
                    'breaker_uploads': raw_panel.get('breakerUploads') or raw_panel.get('breaker_uploads') or [],
                    'directory_uploads': raw_panel.get('directoryUploads') or raw_panel.get('directory_uploads') or [],
                })

        if not panel_requests:
            panel_name = str(data.get('panelName') or data.get('panel_name') or '').strip()
            if not panel_name:
                panel_name = "PANEL"
            panel_requests.append({
                'panel_id': 'panel_1',
                'panel_name': panel_name,
                'breaker_paths': self._normalize_panel_schedule_paths(
                    self._get_panel_schedule_path_value(
                        data,
                        ('breakerPaths', 'breaker_paths'),
                        ('breakerPath', 'breaker_path'),
                    )
                ),
                'directory_paths': self._normalize_panel_schedule_paths(
                    self._get_panel_schedule_path_value(
                        data,
                        ('directoryPaths', 'directory_paths'),
                        ('directoryPath', 'directory_path'),
                    )
                ),
                'breaker_uploads': data.get('breakerUploads') or data.get('breaker_uploads') or [],
                'directory_uploads': data.get('directoryUploads') or data.get('directory_uploads') or [],
            })

        results = []
        success_count = 0
        failure_count = 0

        for index, panel_request in enumerate(panel_requests):
            panel_id = str(panel_request.get('panel_id') or f"panel_{index + 1}").strip() or f"panel_{index + 1}"
            panel_name = str(panel_request.get('panel_name') or '').strip() or f"PANEL {index + 1}"
            breaker_paths = list(panel_request.get('breaker_paths') or [])
            directory_paths = list(panel_request.get('directory_paths') or [])
            breaker_uploads = panel_request.get('breaker_uploads') or []
            directory_uploads = panel_request.get('directory_uploads') or []
            temp_paths = []

            try:
                if breaker_uploads:
                    uploaded_paths, created = self._panel_schedule_save_uploads(
                        breaker_uploads, "breaker"
                    )
                    breaker_paths.extend(uploaded_paths)
                    temp_paths.extend(created)
                if directory_uploads:
                    uploaded_paths, created = self._panel_schedule_save_uploads(
                        directory_uploads, "directory"
                    )
                    directory_paths.extend(uploaded_paths)
                    temp_paths.extend(created)

                if not breaker_paths or not directory_paths:
                    raise ValueError(
                        "At least one breaker photo and at least one directory photo are required."
                    )

                try:
                    panel_data = self._analyze_panel_schedule_images(
                        panel_name, breaker_paths, directory_paths
                    )
                except Exception as e:
                    raise RuntimeError(self._format_panel_schedule_ai_error(e)) from e

                sheet_name = self._update_panel_schedule_workbook(panel_data, output_path)
                success_count += 1
                results.append({
                    'panelId': panel_id,
                    'panelName': panel_name,
                    'status': 'success',
                    'sheetName': sheet_name,
                    'message': f"Panel '{sheet_name}' added to schedule."
                })
            except Exception as e:
                failure_count += 1
                results.append({
                    'panelId': panel_id,
                    'panelName': panel_name,
                    'status': 'error',
                    'message': str(e)
                })
            finally:
                for temp_path in temp_paths:
                    try:
                        os.remove(temp_path)
                    except Exception:
                        pass

        output_folder = os.path.dirname(output_path)
        first_success = next(
            (item for item in results if item.get('status') == 'success'),
            None
        )

        if success_count == 0:
            if created_workbook:
                try:
                    os.remove(output_path)
                except Exception:
                    pass
            first_error_message = next(
                (item.get('message') for item in results if item.get('message')),
                "Panel Schedule AI failed."
            )
            return {
                'status': 'error',
                'message': first_error_message,
                'outputPath': output_path,
                'outputFolder': output_folder,
                'successCount': success_count,
                'failureCount': failure_count,
                'results': results
            }

        if failure_count > 0:
            message = (
                f"Added {success_count} panel sheet{'s' if success_count != 1 else ''}; "
                f"{failure_count} panel{'s' if failure_count != 1 else ''} failed."
            )
        else:
            message = (
                f"Added {success_count} panel sheet{'s' if success_count != 1 else ''} to schedule."
            )

        return {
            'status': 'success',
            'message': message,
            'outputPath': output_path,
            'outputFolder': output_folder,
            'sheetName': first_success.get('sheetName') if first_success else '',
            'successCount': success_count,
            'failureCount': failure_count,
            'results': results
        }

    def _normalize_launch_context(self, launch_context):
        if isinstance(launch_context, dict):
            return launch_context
        if isinstance(launch_context, str):
            raw_text = launch_context.strip()
            if not raw_text:
                return {}
            try:
                parsed = json.loads(raw_text)
                if isinstance(parsed, dict):
                    return parsed
                logging.info(
                    "_normalize_launch_context: Ignoring JSON launch context because it is not an object "
                    f"(type={type(parsed).__name__}).")
                return {}
            except Exception as e:
                preview = raw_text[:120]
                logging.warning(
                    "_normalize_launch_context: Failed to parse launch context JSON string "
                    f"(preview={preview!r}): {e}")
                return {}
        if launch_context is not None:
            logging.info(
                "_normalize_launch_context: Ignoring unsupported launch context type "
                f"(type={type(launch_context).__name__}).")
        return {}

    def _primary_discipline_from_settings(self, settings):
        discipline = settings.get('discipline', 'Electrical')
        if isinstance(discipline, list):
            discipline = next((d for d in discipline if d), 'Electrical')
        discipline = str(discipline or 'Electrical').strip()
        discipline_map = {
            'electrical': 'Electrical',
            'mechanical': 'Mechanical',
            'plumbing': 'Plumbing',
            'general': 'Electrical'
        }
        return discipline_map.get(discipline.lower(), 'Electrical')

    def _get_settings_workroom_disciplines(self, settings):
        discipline_map = {
            'electrical': 'Electrical',
            'mechanical': 'Mechanical',
            'plumbing': 'Plumbing',
        }
        raw_value = settings.get('discipline', ['Electrical'])
        if isinstance(raw_value, str):
            candidates = [raw_value]
        elif isinstance(raw_value, list):
            candidates = raw_value
        else:
            candidates = ['Electrical']

        result = []
        seen = set()
        for item in candidates:
            normalized = discipline_map.get(str(item or '').strip().lower())
            if not normalized or normalized in seen:
                continue
            seen.add(normalized)
            result.append(normalized)
        return result or ['Electrical']

    def _resolve_workroom_discipline(self, settings, launch_context):
        settings_disciplines = self._get_settings_workroom_disciplines(settings)
        context = self._normalize_launch_context(launch_context)
        requested = str(context.get('discipline') or '').strip().lower()
        mapped_requested = {
            'electrical': 'Electrical',
            'mechanical': 'Mechanical',
            'plumbing': 'Plumbing',
        }.get(requested)
        # Workroom selection should drive CAD routing when explicitly provided.
        if mapped_requested:
            return mapped_requested, 'launch_context'
        if len(settings_disciplines) == 1:
            return settings_disciplines[0], 'settings_single'
        return settings_disciplines[0], 'settings_fallback'

    def _resolve_workroom_context(self, settings, launch_context):
        context = self._normalize_launch_context(launch_context)
        source = str(context.get('source') or '').strip().lower()
        project_path = str(context.get('projectPath') or '').strip()
        root_project_path = str(context.get('rootProjectPath') or '').strip()
        effective_project_path = root_project_path or project_path
        if source == 'workroom':
            discipline, discipline_source = self._resolve_workroom_discipline(
                settings, launch_context)
        else:
            discipline = self._primary_discipline_from_settings(settings)
            discipline_source = 'settings_primary'
        normalized_project_path = os.path.normpath(
            effective_project_path) if effective_project_path else ''
        logging.info(
            f"_resolve_workroom_context: source={source or 'none'} "
            f"discipline={discipline} discipline_source={discipline_source} "
            f"project_path={normalized_project_path or '<empty>'} "
            f"(projectPath={project_path or '<empty>'}, rootProjectPath={root_project_path or '<empty>'})")
        return {
            'source': source,
            'project_path': normalized_project_path,
            'discipline': discipline,
            'discipline_source': discipline_source,
        }

    def _is_workroom_auto_select_enabled(self, settings, launch_context):
        context = self._resolve_workroom_context(settings, launch_context)
        source = str(context.get('source') or '').strip().lower()
        auto_select_setting = bool(settings.get('workroomAutoSelectCadFiles', True))
        if source != 'workroom':
            logging.info(
                "_is_workroom_auto_select_enabled: disabled because launch source is not workroom "
                f"(source={source or 'none'}, launch_context_type={type(launch_context).__name__}).")
            return False
        if not auto_select_setting:
            logging.info(
                "_is_workroom_auto_select_enabled: disabled because workroomAutoSelectCadFiles is off.")
            return False
        return True

    def _should_force_workroom_clean_xrefs_manual_selection(self, context):
        source = str((context or {}).get('source') or '').strip().lower()
        return source == 'workroom'

    def _list_base_level_dwgs(self, folder_path):
        if not folder_path or not os.path.isdir(folder_path):
            return []
        file_paths = []
        try:
            with os.scandir(folder_path) as entries:
                for entry in entries:
                    if entry.is_file() and entry.name.lower().endswith('.dwg'):
                        file_paths.append(entry.path)
        except Exception as e:
            logging.warning(f"Failed to list DWGs in {folder_path}: {e}")
            return []
        file_paths.sort(key=lambda path: os.path.basename(path).lower())
        return file_paths

    def _write_files_list_temp(self, file_paths):
        temp_file = tempfile.NamedTemporaryFile(
            mode='w', suffix='.txt', delete=False, encoding='utf-8')
        for path in file_paths:
            temp_file.write(path + '\n')
        temp_file.close()
        return temp_file.name

    def _trace_cad_auto_select(self, event, **fields):
        try:
            payload = {
                'timestamp': datetime.datetime.now(
                    datetime.timezone.utc).isoformat().replace('+00:00', 'Z'),
                'event': str(event or '').strip() or 'unknown',
            }
            for key, value in (fields or {}).items():
                payload[str(key)] = value
            with open(CAD_AUTO_SELECT_TRACE_FILE, 'a', encoding='utf-8') as trace_file:
                trace_file.write(
                    json.dumps(payload, ensure_ascii=True, default=str) + '\n')
        except Exception:
            pass

    def trace_cad_auto_select_event(self, event, payload=None):
        payload_data = dict(payload) if isinstance(payload, dict) else {
            'value': payload,
        }
        payload_data.setdefault('trace_source', 'frontend')
        self._trace_cad_auto_select(
            str(event or '').strip() or 'trace_event',
            **payload_data,
        )
        return {'status': 'success'}

    def clear_cad_auto_select_trace(self):
        try:
            with open(CAD_AUTO_SELECT_TRACE_FILE, 'w', encoding='utf-8'):
                pass
            return {
                'status': 'success',
                'path': CAD_AUTO_SELECT_TRACE_FILE,
            }
        except Exception as e:
            logging.error(f"clear_cad_auto_select_trace failed: {e}")
            return {
                'status': 'error',
                'message': str(e),
                'path': CAD_AUTO_SELECT_TRACE_FILE,
            }

    def get_cad_auto_select_trace(self, line_limit=200):
        try:
            try:
                limit = int(line_limit)
            except (TypeError, ValueError):
                limit = 200
            limit = max(1, min(limit, 1000))

            if not os.path.exists(CAD_AUTO_SELECT_TRACE_FILE):
                return {
                    'status': 'success',
                    'path': CAD_AUTO_SELECT_TRACE_FILE,
                    'entries': [],
                    'lines': [],
                    'lineCount': 0,
                }

            with open(CAD_AUTO_SELECT_TRACE_FILE, 'r', encoding='utf-8') as trace_file:
                lines = trace_file.read().splitlines()
            selected_lines = lines[-limit:]
            entries = []
            for line in selected_lines:
                try:
                    entries.append(json.loads(line))
                except json.JSONDecodeError:
                    entries.append({'raw': line})
            return {
                'status': 'success',
                'path': CAD_AUTO_SELECT_TRACE_FILE,
                'entries': entries,
                'lines': selected_lines,
                'lineCount': len(selected_lines),
            }
        except Exception as e:
            logging.error(f"get_cad_auto_select_trace failed: {e}")
            return {
                'status': 'error',
                'message': str(e),
                'path': CAD_AUTO_SELECT_TRACE_FILE,
                'entries': [],
                'lines': [],
                'lineCount': 0,
            }

    def _normalize_dwg_file_paths(self, file_paths, require_exists=True):
        if not isinstance(file_paths, (list, tuple, set)):
            return []

        resolved_paths = []
        seen = set()
        for raw_path in file_paths:
            raw_text = str(raw_path or '').strip()
            if not raw_text:
                continue
            normalized = os.path.normpath(raw_text)
            if os.path.splitext(normalized)[1].lower() != '.dwg':
                continue
            if require_exists and not os.path.isfile(normalized):
                continue
            key = os.path.normcase(normalized)
            if key in seen:
                continue
            seen.add(key)
            resolved_paths.append(normalized)
        return resolved_paths

    def _get_workroom_cad_file_cache_store(self):
        cache = getattr(self, '_workroom_cad_file_cache', None)
        if not isinstance(cache, dict):
            cache = {}
            self._workroom_cad_file_cache = cache
        return cache

    def _get_workroom_cad_file_cache_key(self, project_path, discipline):
        normalized_project_path = os.path.normpath(
            str(project_path or '').strip()) if project_path else ''
        normalized_discipline = str(discipline or '').strip()
        if not normalized_project_path or not normalized_discipline:
            return None
        return (os.path.normcase(normalized_project_path), normalized_discipline.lower())

    def _set_workroom_cad_file_cache_entry(self, project_path, discipline, folder_path, file_paths, resolution_mode=''):
        cache_key = self._get_workroom_cad_file_cache_key(project_path, discipline)
        if cache_key is None:
            return None

        normalized_project_path = os.path.normpath(str(project_path or '').strip())
        normalized_discipline = str(discipline or '').strip()
        normalized_folder_path = os.path.normpath(
            str(folder_path or '').strip()) if folder_path else ''
        normalized_files = self._normalize_dwg_file_paths(
            file_paths, require_exists=False)

        entry = {
            'project_path': normalized_project_path,
            'discipline': normalized_discipline,
            'folder_path': normalized_folder_path,
            'files': normalized_files,
            'resolution_mode': str(resolution_mode or '').strip(),
        }
        self._get_workroom_cad_file_cache_store()[cache_key] = entry
        return entry

    def _get_workroom_cad_file_cache_entry(self, project_path, discipline):
        cache_key = self._get_workroom_cad_file_cache_key(project_path, discipline)
        if cache_key is None:
            return None
        entry = self._get_workroom_cad_file_cache_store().get(cache_key)
        return entry if isinstance(entry, dict) else None

    def _get_launch_context_cad_file_paths(self, launch_context):
        context = self._normalize_launch_context(launch_context)
        return self._normalize_dwg_file_paths(
            context.get('cadFilePaths'), require_exists=True)

    def _get_shared_parent_folder(self, file_paths):
        parent_dirs = []
        seen = set()
        for path in file_paths:
            parent = os.path.dirname(os.path.normpath(str(path or '').strip()))
            if not parent:
                continue
            key = os.path.normcase(parent)
            if key in seen:
                continue
            seen.add(key)
            parent_dirs.append(parent)
        if len(parent_dirs) == 1:
            return parent_dirs[0]
        return ''

    def _find_project_root_by_id(self, path):
        normalized_path = os.path.normpath(str(path or '').strip()) if path else ''
        if not normalized_path:
            return ''

        current = normalized_path.rstrip("\\/")
        if not current:
            return ''

        basename = os.path.basename(current).strip()
        if os.path.isfile(current) or (not os.path.isdir(current) and os.path.splitext(basename)[1]):
            parent = os.path.dirname(current)
            current = parent.rstrip("\\/") or parent

        while current:
            folder_name = os.path.basename(current).strip()
            if re.match(r'^\d{6}(?!\d)', folder_name):
                logging.info(
                    "_find_project_root_by_id: matched %s from %s",
                    current,
                    normalized_path,
                )
                return current

            parent = os.path.dirname(current)
            if not parent:
                break
            if os.path.normcase(parent) == os.path.normcase(current):
                break
            current = parent.rstrip("\\/") or parent

        return ''

    def _find_workroom_project_root_by_id(self, project_path):
        return self._find_project_root_by_id(project_path)

    def _resolve_workroom_discipline_folder(self, project_path, discipline):
        requested_discipline = str(discipline or '').strip() or 'Electrical'
        discipline_map = {
            'electrical': 'Electrical',
            'mechanical': 'Mechanical',
            'plumbing': 'Plumbing',
            'arch': 'Arch',
        }
        resolved_discipline = discipline_map.get(
            requested_discipline.lower(), requested_discipline)

        known_folder_names = {'electrical', 'mechanical', 'plumbing', 'arch'}
        candidate_entries = []
        seen = set()

        def _add_candidate(path_value, mode):
            if not path_value:
                return
            normalized = os.path.normpath(path_value)
            key = os.path.normcase(normalized)
            if key in seen:
                return
            seen.add(key)
            candidate_entries.append({
                'mode': mode,
                'path': normalized,
            })

        normalized_project_path = os.path.normpath(
            project_path) if project_path else ''
        project_base = os.path.basename(
            normalized_project_path.rstrip("\\/")
        ).strip() if normalized_project_path else ''
        project_root_by_id = self._find_workroom_project_root_by_id(
            normalized_project_path)

        if project_root_by_id:
            _add_candidate(
                os.path.join(project_root_by_id, resolved_discipline),
                'project_id_ancestor_child_folder',
            )

        if normalized_project_path:
            if project_base.lower() == resolved_discipline.lower():
                _add_candidate(
                    normalized_project_path, 'project_path_is_discipline_folder')

            _add_candidate(
                os.path.join(normalized_project_path, resolved_discipline),
                'project_path_child_folder',
            )

            if project_base.lower() in known_folder_names:
                parent_path = os.path.dirname(
                    normalized_project_path.rstrip("\\/"))
                if parent_path:
                    _add_candidate(
                        os.path.join(parent_path, resolved_discipline),
                        'discipline_sibling_folder',
                    )

        for entry in candidate_entries:
            folder_path = entry['path']
            if os.path.isdir(folder_path):
                return {
                    'resolved_folder': folder_path,
                    'mode': entry['mode'],
                    'discipline': resolved_discipline,
                    'candidates': [item['path'] for item in candidate_entries],
                }

        return {
            'resolved_folder': '',
            'mode': 'not_found',
            'discipline': resolved_discipline,
            'candidates': [item['path'] for item in candidate_entries],
        }

    def _is_workroom_auto_select_tool_allowed(self, tool_name):
        normalized_name = str(tool_name or '').strip()
        if normalized_name.endswith('_test_mode'):
            normalized_name = normalized_name[:-10]
        return normalized_name in {
            'run_publish_script',
            'run_freeze_layers_script',
            'run_thaw_layers_script',
        }

    def _resolve_workroom_auto_file_selection(self, settings, launch_context, tool_name):
        if not self._is_workroom_auto_select_tool_allowed(tool_name):
            logging.info(
                f"{tool_name}: Workroom DWG auto-select skipped (tool not allowlisted).")
            self._trace_cad_auto_select(
                'auto_select_skipped_tool_not_allowed',
                tool_name=tool_name,
            )
            return None
        if not self._is_workroom_auto_select_enabled(settings, launch_context):
            self._trace_cad_auto_select(
                'auto_select_skipped_disabled',
                tool_name=tool_name,
                launch_context=self._normalize_launch_context(launch_context),
            )
            return None
        context = self._resolve_workroom_context(settings, launch_context)
        launch_payload = self._normalize_launch_context(launch_context)
        project_path = context.get('project_path') or ''
        discipline = context.get('discipline') or 'Electrical'
        self._trace_cad_auto_select(
            'auto_select_request',
            tool_name=tool_name,
            source=context.get('source') or '',
            project_path=project_path,
            discipline=discipline,
            discipline_source=context.get('discipline_source') or '',
            launch_context=launch_payload,
        )
        cached_entry = self._get_workroom_cad_file_cache_entry(
            project_path, discipline)
        if cached_entry:
            cached_dwg_files = self._normalize_dwg_file_paths(
                cached_entry.get('files'), require_exists=True)
            if cached_dwg_files:
                files_list_path = self._write_files_list_temp(cached_dwg_files)
                cached_folder_path = str(cached_entry.get('folder_path') or '').strip()
                folder_path = cached_folder_path or self._get_shared_parent_folder(
                    cached_dwg_files)
                logging.info(
                    f"{tool_name}: Auto-selected {len(cached_dwg_files)} DWG(s) from cached Workroom detection "
                    f"(folder={folder_path or '<multiple>'}; cached_mode={cached_entry.get('resolution_mode') or 'unknown'}).")
                self._trace_cad_auto_select(
                    'auto_select_selected',
                    tool_name=tool_name,
                    selection_source='workroom_cached_detection',
                    project_path=project_path,
                    discipline=str(cached_entry.get('discipline') or discipline),
                    folder_path=folder_path,
                    files_list_path=files_list_path,
                    count=len(cached_dwg_files),
                    file_paths=cached_dwg_files,
                    cached_resolution_mode=cached_entry.get('resolution_mode') or '',
                )
                return {
                    'files_list_path': files_list_path,
                    'project_path': project_path,
                    'discipline': str(cached_entry.get('discipline') or discipline),
                    'folder_path': folder_path,
                    'count': len(cached_dwg_files),
                    'resolution_mode': 'workroom_cached_detection',
                }
            cached_file_count = len(cached_entry.get('files') or [])
            if cached_file_count:
                logging.info(
                    f"{tool_name}: Cached Workroom detection had {cached_file_count} DWG(s), but none still exist. "
                    "Falling back.")
                self._trace_cad_auto_select(
                    'auto_select_cached_stale',
                    tool_name=tool_name,
                    project_path=project_path,
                    discipline=discipline,
                    cached_file_count=cached_file_count,
                    cached_files=cached_entry.get('files') or [],
                )
            else:
                logging.info(
                    f"{tool_name}: Cached Workroom detection is empty for project_path={project_path} "
                    f"discipline={discipline}. Falling back.")
                self._trace_cad_auto_select(
                    'auto_select_cached_empty',
                    tool_name=tool_name,
                    project_path=project_path,
                    discipline=discipline,
                )
        explicit_dwg_files = self._get_launch_context_cad_file_paths(launch_context)
        if explicit_dwg_files:
            files_list_path = self._write_files_list_temp(explicit_dwg_files)
            shared_parent = self._get_shared_parent_folder(explicit_dwg_files)
            logging.info(
                f"{tool_name}: Auto-selected {len(explicit_dwg_files)} explicit DWG(s) from launch context "
                f"(folder={shared_parent or '<multiple>'}).")
            self._trace_cad_auto_select(
                'auto_select_selected',
                tool_name=tool_name,
                selection_source='launch_context_explicit_files',
                project_path=project_path,
                discipline=discipline,
                folder_path=shared_parent,
                files_list_path=files_list_path,
                count=len(explicit_dwg_files),
                file_paths=explicit_dwg_files,
            )
            return {
                'files_list_path': files_list_path,
                'project_path': project_path,
                'discipline': discipline,
                'folder_path': shared_parent,
                'count': len(explicit_dwg_files),
                'resolution_mode': 'launch_context_explicit_files',
            }
        if 'cadFilePaths' in launch_payload:
            logging.info(
                f"{tool_name}: Launch context cadFilePaths were provided but no valid DWG files remained. "
                "Falling back to folder scan.")
            self._trace_cad_auto_select(
                'auto_select_explicit_invalid',
                tool_name=tool_name,
                project_path=project_path,
                discipline=discipline,
                cad_file_paths=launch_payload.get('cadFilePaths'),
            )
        if not project_path:
            logging.info(
                f"{tool_name}: Workroom auto-select fallback to manual file picker (missing project path).")
            self._trace_cad_auto_select(
                'auto_select_fallback_manual',
                tool_name=tool_name,
                reason='missing_project_path',
                discipline=discipline,
            )
            return None
        folder_resolution = self._resolve_workroom_discipline_folder(
            project_path, discipline)
        discipline_folder = folder_resolution.get('resolved_folder') or ''
        if not discipline_folder:
            candidates = folder_resolution.get('candidates') or []
            candidate_text = '; '.join(candidates) if candidates else 'none'
            logging.info(
                f"{tool_name}: Workroom auto-select fallback to manual file picker "
                f"(project_path={project_path}; discipline={discipline}; checked={candidate_text}).")
            self._trace_cad_auto_select(
                'auto_select_fallback_manual',
                tool_name=tool_name,
                reason='discipline_folder_not_found',
                project_path=project_path,
                discipline=discipline,
                checked=candidates,
            )
            return None
        dwg_files = self._list_base_level_dwgs(discipline_folder)
        if not dwg_files:
            logging.info(
                f"{tool_name}: Workroom auto-select fallback to manual file picker (no DWGs in {discipline_folder}).")
            self._trace_cad_auto_select(
                'auto_select_fallback_manual',
                tool_name=tool_name,
                reason='no_dwgs_in_folder',
                project_path=project_path,
                discipline=discipline,
                folder_path=discipline_folder,
            )
            return None
        files_list_path = self._write_files_list_temp(dwg_files)
        logging.info(
            f"{tool_name}: Auto-selected {len(dwg_files)} DWG(s) from {discipline_folder} "
            f"(mode={folder_resolution.get('mode')}).")
        self._trace_cad_auto_select(
            'auto_select_selected',
            tool_name=tool_name,
            selection_source=folder_resolution.get('mode', '') or 'folder_scan',
            project_path=project_path,
            discipline=folder_resolution.get('discipline') or discipline,
            folder_path=discipline_folder,
            files_list_path=files_list_path,
            count=len(dwg_files),
            file_paths=dwg_files,
        )
        return {
            'files_list_path': files_list_path,
            'project_path': project_path,
            'discipline': folder_resolution.get('discipline') or discipline,
            'folder_path': discipline_folder,
            'count': len(dwg_files),
            'resolution_mode': folder_resolution.get('mode', ''),
        }

    def _resolve_workroom_arch_folder(self, settings, launch_context):
        context = self._resolve_workroom_context(settings, launch_context)
        source = str(context.get('source') or '').strip().lower()
        if source != 'workroom':
            return ''
        project_path = context.get('project_path') or ''
        if not project_path:
            logging.info(
                "_resolve_workroom_arch_folder: Missing project path for workroom launch.")
            return ''
        arch_resolution = self._resolve_workroom_discipline_folder(
            project_path, 'Arch')
        arch_folder = arch_resolution.get('resolved_folder') or ''
        if os.path.isdir(arch_folder):
            logging.info(
                f"_resolve_workroom_arch_folder: Resolved Arch folder at {arch_folder} "
                f"(mode={arch_resolution.get('mode')}).")
            return arch_folder
        candidates = arch_resolution.get('candidates') or []
        if candidates:
            logging.info(
                "_resolve_workroom_arch_folder: Arch folder not found. "
                f"Checked: {'; '.join(candidates)}")
        return ''

    def get_workroom_cad_files(self, launch_context=None):
        try:
            settings = self.get_user_settings()
            context = self._resolve_workroom_context(settings, launch_context)
            project_path = context.get('project_path') or ''
            discipline = context.get('discipline') or self._primary_discipline_from_settings(
                settings)
            normalized_launch_context = self._normalize_launch_context(launch_context)
            self._trace_cad_auto_select(
                'get_workroom_cad_files_request',
                project_path=project_path,
                discipline=discipline,
                source=context.get('source') or '',
                discipline_source=context.get('discipline_source') or '',
                launch_context=normalized_launch_context,
            )

            if not project_path:
                self._trace_cad_auto_select(
                    'get_workroom_cad_files_missing_project_path',
                    discipline=discipline,
                    launch_context=normalized_launch_context,
                )
                return {
                    'status': 'success',
                    'projectPath': '',
                    'discipline': discipline,
                    'folderPath': '',
                    'resolutionMode': 'missing_project_path',
                    'files': [],
                    'message': 'Project folder path is missing.',
                }

            folder_resolution = self._resolve_workroom_discipline_folder(
                project_path, discipline)
            resolved_discipline = folder_resolution.get(
                'discipline') or discipline
            folder_path = folder_resolution.get('resolved_folder') or ''
            resolution_mode = folder_resolution.get('mode') or ''

            if not folder_path:
                candidates = folder_resolution.get('candidates') or []
                if candidates:
                    logging.info(
                        "get_workroom_cad_files: Discipline folder not found "
                        f"(project_path={project_path}; discipline={resolved_discipline}; "
                        f"checked={'; '.join(candidates)})."
                    )
                else:
                    logging.info(
                        "get_workroom_cad_files: Discipline folder not found "
                        f"(project_path={project_path}; discipline={resolved_discipline})."
                    )
                self._set_workroom_cad_file_cache_entry(
                    project_path,
                    resolved_discipline,
                    '',
                    [],
                    resolution_mode or 'not_found',
                )
                self._trace_cad_auto_select(
                    'get_workroom_cad_files_result',
                    project_path=project_path,
                    discipline=resolved_discipline,
                    resolution_mode=resolution_mode or 'not_found',
                    folder_path='',
                    candidate_paths=candidates,
                    count=0,
                    file_paths=[],
                )
                return {
                    'status': 'success',
                    'projectPath': project_path,
                    'discipline': resolved_discipline,
                    'folderPath': '',
                    'resolutionMode': resolution_mode or 'not_found',
                    'files': [],
                    'message': f'{resolved_discipline} folder not found.',
                }

            dwg_paths = self._list_base_level_dwgs(folder_path)
            files = [
                {
                    'name': os.path.basename(path),
                    'path': path,
                }
                for path in dwg_paths
            ]
            self._set_workroom_cad_file_cache_entry(
                project_path,
                resolved_discipline,
                folder_path,
                dwg_paths,
                resolution_mode,
            )
            self._trace_cad_auto_select(
                'get_workroom_cad_files_result',
                project_path=project_path,
                discipline=resolved_discipline,
                resolution_mode=resolution_mode,
                folder_path=folder_path,
                count=len(dwg_paths),
                file_paths=dwg_paths,
            )

            return {
                'status': 'success',
                'projectPath': project_path,
                'discipline': resolved_discipline,
                'folderPath': folder_path,
                'resolutionMode': resolution_mode,
                'files': files,
                'message': (
                    f'Found {len(files)} DWG file{"s" if len(files) != 1 else ""}.'
                ),
            }
        except Exception as e:
            logging.error(f"get_workroom_cad_files failed: {e}")
            self._trace_cad_auto_select(
                'get_workroom_cad_files_error',
                error=str(e),
                launch_context=self._normalize_launch_context(launch_context),
            )
            return {'status': 'error', 'message': str(e), 'files': []}

    def _notify_tool_status(self, tool_id, message):
        try:
            if not webview.windows:
                return
            js_message = json.dumps(str(message or '').strip())
            webview.windows[0].evaluate_js(
                f'window.updateToolStatus("{tool_id}", {js_message})')
        except Exception as e:
            logging.debug(
                f"_notify_tool_status failed for {tool_id}: {e}")

    def _record_test_mode_event(self, event):
        if not self.test_mode:
            return
        event_data = dict(event or {})
        event_data['timestamp'] = datetime.datetime.utcnow().isoformat() + 'Z'
        self._test_mode_records.append(event_data)

    def get_test_mode_records(self):
        return {
            'status': 'success',
            'test_mode': self.test_mode,
            'records': list(self._test_mode_records),
        }

    def _run_workroom_cad_tool_in_test_mode(self, settings, launch_context, tool_id, tool_name):
        context = self._resolve_workroom_context(settings, launch_context)
        source = context.get('source') or 'none'
        project_path = context.get('project_path') or ''
        discipline = context.get('discipline') or self._primary_discipline_from_settings(settings)
        discipline_source = context.get('discipline_source') or 'unknown'

        if not self._is_workroom_auto_select_enabled(settings, launch_context):
            message = (
                "TEST MODE: Workroom auto-select is disabled for this run. "
                f"(source={source}, project_path={project_path or '<empty>'}, "
                f"discipline={discipline}, discipline_source={discipline_source})"
            )
            self._record_test_mode_event({
                'tool_id': tool_id,
                'tool_name': tool_name,
                'status': 'error',
                'reason': 'auto_select_disabled',
                'source': source,
                'project_path': project_path,
                'discipline': discipline,
                'discipline_source': discipline_source,
            })
            self._notify_tool_status(tool_id, f"ERROR: {message}")
            return {'status': 'error', 'test_mode': True, 'message': message}

        auto_selection = self._resolve_workroom_auto_file_selection(
            settings, launch_context, f'{tool_name}_test_mode'
        )
        if not auto_selection:
            message = (
                "TEST MODE: Expected auto-selected DWG files but none were resolved. "
                f"(source={source}, project_path={project_path or '<empty>'}, "
                f"discipline={discipline}, discipline_source={discipline_source})"
            )
            self._record_test_mode_event({
                'tool_id': tool_id,
                'tool_name': tool_name,
                'status': 'error',
                'reason': 'no_auto_selected_files',
                'source': source,
                'project_path': project_path,
                'discipline': discipline,
                'discipline_source': discipline_source,
            })
            self._notify_tool_status(tool_id, f"ERROR: {message}")
            return {'status': 'error', 'test_mode': True, 'message': message}

        count = int(auto_selection.get('count') or 0)
        folder_path = auto_selection.get('folder_path') or ''
        self._record_test_mode_event({
            'tool_id': tool_id,
            'tool_name': tool_name,
            'status': 'success',
            'source': source,
            'project_path': project_path,
            'discipline': discipline,
            'discipline_source': discipline_source,
            'folder_path': folder_path,
            'count': count,
            'resolution_mode': auto_selection.get('resolution_mode') or '',
        })
        self._notify_tool_status(
            tool_id,
            f"TEST MODE: Auto-selected {count} DWG(s) from {folder_path or 'unknown folder'}."
        )
        # Match normal tool completion signal so UI state is identical to production runs.
        self._notify_tool_status(tool_id, "DONE")
        return {
            'status': 'success',
            'test_mode': True,
            'tool_id': tool_id,
            'tool_name': tool_name,
            'source': source,
            'project_path': project_path,
            'discipline': discipline,
            'discipline_source': discipline_source,
            'folder_path': folder_path,
            'count': count,
        }

    def _run_workroom_clean_xrefs_tool_in_test_mode(self, settings, launch_context):
        context = self._resolve_workroom_context(settings, launch_context)
        source = context.get('source') or 'none'
        project_path = context.get('project_path') or ''
        discipline = context.get('discipline') or self._primary_discipline_from_settings(settings)
        discipline_source = context.get('discipline_source') or 'unknown'

        if not self._should_force_workroom_clean_xrefs_manual_selection(context):
            message = (
                "TEST MODE: Non-workroom launch. "
                "Clean-xrefs uses manual file selection with the default picker location "
                f"(source={source}, project_path={project_path or '<empty>'}, "
                f"discipline={discipline}, discipline_source={discipline_source})."
            )
            self._record_test_mode_event({
                'tool_id': 'toolCleanXrefs',
                'tool_name': 'run_clean_xrefs_script',
                'status': 'success',
                'reason': 'non_workroom_manual_picker',
                'source': source,
                'project_path': project_path,
                'discipline': discipline,
                'discipline_source': discipline_source,
                'manual_selection_enforced': False,
                'arch_folder': '',
            })
            self._notify_tool_status('toolCleanXrefs', message)
            self._notify_tool_status('toolCleanXrefs', "DONE")
            return {
                'status': 'success',
                'test_mode': True,
                'tool_id': 'toolCleanXrefs',
                'tool_name': 'run_clean_xrefs_script',
                'source': source,
                'project_path': project_path,
                'discipline': discipline,
                'discipline_source': discipline_source,
                'manual_selection_enforced': False,
                'arch_folder': '',
            }

        arch_folder = self._resolve_workroom_arch_folder(settings, launch_context)
        if arch_folder:
            message = (
                "TEST MODE: Workroom manual selection enforced. "
                f"Opening Arch folder picker at {arch_folder}."
            )
        else:
            message = (
                "TEST MODE: Workroom manual selection enforced. "
                "Arch folder not found; opening default file picker."
            )

        self._record_test_mode_event({
            'tool_id': 'toolCleanXrefs',
            'tool_name': 'run_clean_xrefs_script',
            'status': 'success',
            'source': source,
            'project_path': project_path,
            'discipline': discipline,
            'discipline_source': discipline_source,
            'manual_selection_enforced': True,
            'arch_folder': arch_folder,
        })
        self._notify_tool_status('toolCleanXrefs', message)
        # Match normal tool completion signal so UI state is identical to production runs.
        self._notify_tool_status('toolCleanXrefs', "DONE")
        return {
            'status': 'success',
            'test_mode': True,
            'tool_id': 'toolCleanXrefs',
            'tool_name': 'run_clean_xrefs_script',
            'source': source,
            'project_path': project_path,
            'discipline': discipline,
            'discipline_source': discipline_source,
            'manual_selection_enforced': True,
            'arch_folder': arch_folder,
        }

    def run_publish_script(self, launch_context=None):
        """Runs the PlotDWGs.ps1 PowerShell script with progress updates."""
        script_path = os.path.join(BASE_DIR, "scripts", "PlotDWGs.ps1")
        if not os.path.exists(script_path):
            raise Exception("PlotDWGs.ps1 not found in scripts directory.")
        settings = self.get_user_settings()
        if self.test_mode:
            return self._run_workroom_cad_tool_in_test_mode(
                settings,
                launch_context,
                'toolPublishDwgs',
                'run_publish_script',
            )
        acad_path = settings.get('autocadPath', '')
        if not acad_path:
            raise Exception("No AutoCAD version selected in settings.")
        publish_options = settings.get('publishDwgOptions') or {}
        auto_detect = publish_options.get('autoDetectPaperSize', True)
        shrink_percent = publish_options.get('shrinkPercent', 100)

        def _ps_bool(value):
            return "1" if value else "0"

        auto_selection = self._resolve_workroom_auto_file_selection(
            settings, launch_context, 'run_publish_script')
        if self._is_workroom_auto_select_enabled(settings, launch_context) and not auto_selection:
            fallback_context = self._resolve_workroom_context(
                settings, launch_context)
            logging.info(
                "run_publish_script: Workroom auto-select unavailable; opening file picker "
                f"(source={fallback_context.get('source') or 'none'}, "
                f"project_path={fallback_context.get('project_path') or '<empty>'}, "
                f"discipline={fallback_context.get('discipline')}, "
                f"discipline_source={fallback_context.get('discipline_source')}).")
            self._notify_tool_status(
                'toolPublishDwgs',
                "Workroom auto-select unavailable. Opening file picker...",
            )
            self._trace_cad_auto_select(
                'tool_manual_fallback',
                tool_id='toolPublishDwgs',
                tool_name='run_publish_script',
                reason='auto_selection_unavailable',
                source=fallback_context.get('source') or 'none',
                project_path=fallback_context.get('project_path') or '',
                discipline=fallback_context.get('discipline') or '',
            )
        command = self._build_powershell_script_command(
            script_path,
            '-AcadCore',
            acad_path,
            '-AutoDetectPaperSize',
            _ps_bool(auto_detect),
            '-ShrinkPercent',
            shrink_percent,
        )
        if auto_selection:
            command.extend([
                '-FilesListPath',
                auto_selection['files_list_path'],
            ])
            selected_count = auto_selection.get('count')
            if isinstance(selected_count, int) and selected_count > 0:
                self._notify_tool_status(
                    'toolPublishDwgs',
                    f"Using auto-selected DWGs ({selected_count}) via {auto_selection.get('resolution_mode') or 'workroom'}...",
                )
        self._trace_cad_auto_select(
            'tool_command_launch',
            tool_id='toolPublishDwgs',
            tool_name='run_publish_script',
            command=command,
            command_type='argv',
            has_files_list_path='-FilesListPath' in command,
            files_list_path=auto_selection.get('files_list_path') if auto_selection else '',
            auto_selection=auto_selection,
        )
        self._run_script_with_progress(command, 'toolPublishDwgs')
        return {'status': 'success'}

    def run_freeze_layers_script(self, launch_context=None):
        """Runs the FreezeLayersDWGs.ps1 PowerShell script with progress updates."""
        script_path = os.path.join(BASE_DIR, "scripts", "FreezeLayersDWGs.ps1")
        if not os.path.exists(script_path):
            raise Exception(
                "FreezeLayersDWGs.ps1 not found in scripts directory.")
        settings = self.get_user_settings()
        if self.test_mode:
            return self._run_workroom_cad_tool_in_test_mode(
                settings,
                launch_context,
                'toolFreezeLayers',
                'run_freeze_layers_script',
            )
        acad_path = settings.get('autocadPath', '')
        if not acad_path:
            raise Exception("No AutoCAD version selected in settings.")
        freeze_options = settings.get('freezeLayerOptions') or {}
        scan_all = freeze_options.get('scanAllLayers', True)

        def _ps_bool(value):
            return "1" if value else "0"

        auto_selection = self._resolve_workroom_auto_file_selection(
            settings, launch_context, 'run_freeze_layers_script')
        if self._is_workroom_auto_select_enabled(settings, launch_context) and not auto_selection:
            fallback_context = self._resolve_workroom_context(
                settings, launch_context)
            logging.info(
                "run_freeze_layers_script: Workroom auto-select unavailable; opening file picker "
                f"(source={fallback_context.get('source') or 'none'}, "
                f"project_path={fallback_context.get('project_path') or '<empty>'}, "
                f"discipline={fallback_context.get('discipline')}, "
                f"discipline_source={fallback_context.get('discipline_source')}).")
            self._notify_tool_status(
                'toolFreezeLayers',
                "Workroom auto-select unavailable. Opening file picker...",
            )
            self._trace_cad_auto_select(
                'tool_manual_fallback',
                tool_id='toolFreezeLayers',
                tool_name='run_freeze_layers_script',
                reason='auto_selection_unavailable',
                source=fallback_context.get('source') or 'none',
                project_path=fallback_context.get('project_path') or '',
                discipline=fallback_context.get('discipline') or '',
            )
        command = self._build_powershell_script_command(
            script_path,
            '-AcadCore',
            acad_path,
            '-ScanAllLayers',
            _ps_bool(scan_all),
        )
        if auto_selection:
            command.extend([
                '-FilesListPath',
                auto_selection['files_list_path'],
            ])
            selected_count = auto_selection.get('count')
            if isinstance(selected_count, int) and selected_count > 0:
                self._notify_tool_status(
                    'toolFreezeLayers',
                    f"Using auto-selected DWGs ({selected_count}) via {auto_selection.get('resolution_mode') or 'workroom'}...",
                )
        self._trace_cad_auto_select(
            'tool_command_launch',
            tool_id='toolFreezeLayers',
            tool_name='run_freeze_layers_script',
            command=command,
            command_type='argv',
            has_files_list_path='-FilesListPath' in command,
            files_list_path=auto_selection.get('files_list_path') if auto_selection else '',
            auto_selection=auto_selection,
        )
        self._run_script_with_progress(command, 'toolFreezeLayers')
        return {'status': 'success'}

    def run_thaw_layers_script(self, launch_context=None):
        """Runs the ThawLayersDWGs.ps1 PowerShell script with progress updates."""
        script_path = os.path.join(BASE_DIR, "scripts", "ThawLayersDWGs.ps1")
        if not os.path.exists(script_path):
            raise Exception(
                "ThawLayersDWGs.ps1 not found in scripts directory.")
        settings = self.get_user_settings()
        if self.test_mode:
            return self._run_workroom_cad_tool_in_test_mode(
                settings,
                launch_context,
                'toolThawLayers',
                'run_thaw_layers_script',
            )
        acad_path = settings.get('autocadPath', '')
        if not acad_path:
            raise Exception("No AutoCAD version selected in settings.")
        thaw_options = settings.get('thawLayerOptions') or {}
        scan_all = thaw_options.get('scanAllLayers', True)

        def _ps_bool(value):
            return "1" if value else "0"

        auto_selection = self._resolve_workroom_auto_file_selection(
            settings, launch_context, 'run_thaw_layers_script')
        if self._is_workroom_auto_select_enabled(settings, launch_context) and not auto_selection:
            fallback_context = self._resolve_workroom_context(
                settings, launch_context)
            logging.info(
                "run_thaw_layers_script: Workroom auto-select unavailable; opening file picker "
                f"(source={fallback_context.get('source') or 'none'}, "
                f"project_path={fallback_context.get('project_path') or '<empty>'}, "
                f"discipline={fallback_context.get('discipline')}, "
                f"discipline_source={fallback_context.get('discipline_source')}).")
            self._notify_tool_status(
                'toolThawLayers',
                "Workroom auto-select unavailable. Opening file picker...",
            )
            self._trace_cad_auto_select(
                'tool_manual_fallback',
                tool_id='toolThawLayers',
                tool_name='run_thaw_layers_script',
                reason='auto_selection_unavailable',
                source=fallback_context.get('source') or 'none',
                project_path=fallback_context.get('project_path') or '',
                discipline=fallback_context.get('discipline') or '',
            )
        command = self._build_powershell_script_command(
            script_path,
            '-AcadCore',
            acad_path,
            '-ScanAllLayers',
            _ps_bool(scan_all),
        )
        if auto_selection:
            command.extend([
                '-FilesListPath',
                auto_selection['files_list_path'],
            ])
            selected_count = auto_selection.get('count')
            if isinstance(selected_count, int) and selected_count > 0:
                self._notify_tool_status(
                    'toolThawLayers',
                    f"Using auto-selected DWGs ({selected_count}) via {auto_selection.get('resolution_mode') or 'workroom'}...",
                )
        self._trace_cad_auto_select(
            'tool_command_launch',
            tool_id='toolThawLayers',
            tool_name='run_thaw_layers_script',
            command=command,
            command_type='argv',
            has_files_list_path='-FilesListPath' in command,
            files_list_path=auto_selection.get('files_list_path') if auto_selection else '',
            auto_selection=auto_selection,
        )
        self._run_script_with_progress(command, 'toolThawLayers')
        return {'status': 'success'}

    def run_clean_xrefs_script(self, launch_context=None):
        """Runs the removeXREFPaths.ps1 PowerShell script with progress updates."""
        try:
            script_path = os.path.join(BASE_DIR, "scripts", "removeXREFPaths.ps1")
            if not os.path.exists(script_path):
                raise Exception(
                    "removeXREFPaths.ps1 not found in scripts directory.")
            settings = self.get_user_settings()
            if self.test_mode:
                return self._run_workroom_clean_xrefs_tool_in_test_mode(
                    settings,
                    launch_context,
                )
            acad_path = settings.get('autocadPath', '')
            if not acad_path:
                raise Exception("No AutoCAD version selected in settings.")
            context = self._resolve_workroom_context(settings, launch_context)
            discipline = context.get('discipline') or self._primary_discipline_from_settings(settings)
            discipline_short_map = {
                'Electrical': 'E',
                'Mechanical': 'M',
                'Plumbing': 'P',
                'General': 'G'
            }
            discipline_short = discipline_short_map.get(
                discipline, (discipline[:1] or 'E').upper())

            clean_options = settings.get('cleanDwgOptions') or {}
            def _bool_setting(key, default=True):
                value = clean_options.get(key, default)
                return bool(value) if value is not None else default

            strip_xrefs = _bool_setting('stripXrefs', True)
            set_by_layer = _bool_setting('setByLayer', True)
            purge = _bool_setting('purge', True)
            audit = _bool_setting('audit', True)
            hatch_color = _bool_setting('hatchColor', True)

            def _ps_bool(value):
                return "1" if value else "0"

            source = context.get('source') or 'none'
            project_path = context.get('project_path') or ''
            discipline_source = context.get('discipline_source') or 'unknown'
            force_manual_selection = self._should_force_workroom_clean_xrefs_manual_selection(
                context)

            auto_selection = None
            arch_folder = self._resolve_workroom_arch_folder(
                settings, launch_context) if force_manual_selection else ''
            if force_manual_selection:
                if arch_folder:
                    self._notify_tool_status(
                        'toolCleanXrefs',
                        "Opening Arch folder for file selection...",
                    )
                    logging.info(
                        "run_clean_xrefs_script: Workroom manual selection enforced. "
                        f"(source={source}, project_path={project_path or '<empty>'}, "
                        f"discipline={discipline}, discipline_source={discipline_source}, "
                        f"arch_folder={arch_folder})")
                else:
                    self._notify_tool_status(
                        'toolCleanXrefs',
                        "Arch folder not found. Opening file picker...",
                    )
                    logging.info(
                        "run_clean_xrefs_script: Workroom manual selection enforced, Arch folder not found. "
                        f"(source={source}, project_path={project_path or '<empty>'}, "
                        f"discipline={discipline}, discipline_source={discipline_source})")
            else:
                auto_selection = self._resolve_workroom_auto_file_selection(
                    settings, launch_context, 'run_clean_xrefs_script')
                if self._is_workroom_auto_select_enabled(settings, launch_context) and not auto_selection:
                    fallback_context = self._resolve_workroom_context(
                        settings, launch_context)
                    logging.info(
                        "run_clean_xrefs_script: Workroom auto-select unavailable; opening file picker "
                        f"(source={fallback_context.get('source') or 'none'}, "
                        f"project_path={fallback_context.get('project_path') or '<empty>'}, "
                        f"discipline={fallback_context.get('discipline')}, "
                        f"discipline_source={fallback_context.get('discipline_source')}).")
                    self._notify_tool_status(
                        'toolCleanXrefs',
                        "Workroom auto-select unavailable. Opening file picker...",
                    )

            command = (
                f'powershell.exe -ExecutionPolicy Bypass -File "{script_path}" '
                f'-AcadCore "{acad_path}" -DisciplineShort "{discipline_short}" '
                f'-StripXrefs {_ps_bool(strip_xrefs)} '
                f'-SetByLayer {_ps_bool(set_by_layer)} '
                f'-Purge {_ps_bool(purge)} '
                f'-Audit {_ps_bool(audit)} '
                f'-HatchColor {_ps_bool(hatch_color)}'
            )
            if auto_selection:
                command += f' -FilesListPath "{auto_selection["files_list_path"]}"'
            if arch_folder:
                command += f' -DefaultDirectory "{arch_folder}"'
            self._run_script_with_progress(command, 'toolCleanXrefs')
            return {'status': 'success'}
        except Exception as e:
            logging.error(f"run_clean_xrefs_script failed: {e}")
            if webview.windows:
                window = webview.windows[0]
                window.evaluate_js(f'window.updateToolStatus("toolCleanXrefs", "ERROR: {str(e)}")')
            return {'status': 'error', 'message': str(e)}

    def select_files(self, options):
        """Shows a file dialog and returns selected paths."""
        try:
            window = webview.windows[0]
            file_paths = window.create_file_dialog(
                webview.FileDialog.OPEN,  # DEPRECATION FIX
                allow_multiple=options.get('allow_multiple', False),
                file_types=tuple(options.get('file_types', ()))
            )
            if not file_paths:
                return {'status': 'cancelled', 'paths': []}
            return {'status': 'success', 'paths': file_paths}
        except Exception as e:
            logging.error(f"Error in file dialog: {e}")
            return {'status': 'error', 'message': str(e)}







    def import_and_process_csv(self):
        """Handles CSV file import dialog."""
        try:
            window = webview.windows[0]
            file_paths = window.create_file_dialog(
                webview.FileDialog.OPEN,  # DEPRECATION FIX
                allow_multiple=False,
                file_types=('CSV Files (*.csv)',)
            )
            if not file_paths:
                return {'status': 'cancelled', 'data': []}
            with open(file_paths[0], 'r', encoding='utf-8-sig') as f:
                csv_content = f.read()
            projects = self._process_csv_rows(csv_content, file_paths[0])
            return {'status': 'success', 'data': projects}
        except Exception as e:
            logging.error(f"Error during CSV import: {e}")
            return {'status': 'error', 'message': str(e), 'data': []}

    def _process_csv_rows(self, csv_content, csv_path):
        """Helper to process CSV data."""
        new_projects = []
        fallback_paths = {}
        csv_dir = os.path.dirname(csv_path)
        fallback_json_path = os.path.join(csv_dir, 'tasks.json')
        if os.path.exists(fallback_json_path):
            try:
                with open(fallback_json_path, 'r', encoding='utf-8') as f:
                    fallback_data = json.load(f)
                    for p in fallback_data:
                        if p.get('id') and p.get('path'):
                            fallback_paths[str(p['id']).strip()] = p['path']
            except Exception as e:
                logging.warning(f"Could not load fallback tasks.json: {e}")
        f = io.StringIO(csv_content)
        reader = csv.reader(f)
        try:
            header = next(reader)
            if 'project name' not in ''.join(header).lower():
                f.seek(0)
        except StopIteration:
            return []
        for row in reader:
            if not any(field.strip() for field in row):
                continue
            row.extend([''] * 12)
            proj_id, name, nick, notes, due, tasks_str, path, *_ = row
            if not any([proj_id, name, path]):
                continue
            final_path = (path or '').strip()
            if final_path and not os.path.exists(os.path.normpath(final_path)):
                fallback_path = fallback_paths.get((proj_id or '').strip())
                if fallback_path:
                    final_path = fallback_path
            tasks_list = [
                {'text': t.strip(), 'done': False, 'links': []}
                for t in tasks_str.replace('\r', '\n').replace(';', '\n').replace('\u2022', '\n').split('\n')
                if t.strip()
            ]
            project = {
                'id': (proj_id or '').strip(),
                'name': (name or '').strip(),
                'nick': (nick or '').strip(),
                'notes': (notes or '').strip(),
                'due': (due or '').strip(),
                'path': final_path,
                'tasks': tasks_list,
                'refs': [],
                'statuses': []
            }
            new_projects.append(project)
        return new_projects


# --- Main Application Setup ---
if __name__ == '__main__':
    api = Api()
    window = webview.create_window(
        'ACIES Desktop Application',
        'index.html',
        js_api=api,
        width=1400,
        height=900,
        resizable=True,
        min_size=(1024, 768)
    )
    # --- FIX: Removed the line that caused the circular reference ---
    webview.start()
    # Cleanup subprocesses on exit
    api.stop_circuit_breaker_server()
