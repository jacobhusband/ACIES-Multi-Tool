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
from pathlib import Path
from dotenv import load_dotenv
import datetime
import threading
import requests  # Added for GitHub API calls
import zipfile   # Added for extracting bundles
import math
import random
import re
import base64
from typing import List
import importlib
import uuid
from urllib.parse import urlparse, unquote

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
from PIL import Image as PILImage

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

    word = win32com.client.Dispatch("Word.Application")
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


def get_app_data_path(file_name="tasks.json"):
    """
    Determines the correct, cross-platform path for storing user data
    and returns the full path for the given file_name.

    On Windows, we store data in the user's Documents folder to bypass
    Windows Store Python's file system virtualization, which redirects
    writes to AppData to a package-specific virtualized location.
    """
    if sys.platform == "win32":
        _migrate_legacy_app_data()
        # Use Documents folder - it is NOT virtualized by Windows Store Python
        # This ensures both dev (python main.py) and packaged apps use the same location
        base_dir = _get_windows_documents_dir()
    elif sys.platform == "darwin":  # macOS
        base_dir = os.path.join(os.path.expanduser(
            '~'), 'Library', 'Application Support')
    else:  # Linux and other Unix-like systems
        base_dir = os.path.join(os.path.expanduser('~'), '.local', 'share')
    if not base_dir:
        base_dir = os.path.expanduser('~')
    app_data_dir = os.path.join(base_dir, 'ProjectManagementApp')
    os.makedirs(app_data_dir, exist_ok=True)
    return os.path.join(app_data_dir, file_name)


# --- Configuration ---
BASE_DIR = Path(__file__).resolve().parent
loaded = load_dotenv(BASE_DIR / '.env', override=True)  # <— force .env to win
logging.info(f".env loaded: {loaded} (override=True)")
TASKS_FILE = get_app_data_path("tasks.json")
NOTES_FILE = get_app_data_path("notes.json")
SETTINGS_FILE = get_app_data_path("settings.json")
TIMESHEETS_FILE = get_app_data_path("timesheets.json")
TEMPLATES_FILE = get_app_data_path("templates.json")
CHECKLISTS_FILE = get_app_data_path("checklists.json")
LIGHTING_SCHEDULE_SYNC_FILE = "T24LightingFixtureSchedule.sync.json"


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
            # 1. Get local bundles
            local_bundles = os.listdir(self.app_plugins_folder)

            # 2. Resolve the latest AutoCAD plugin release/tag from its repo
            release_info = self._fetch_latest_bundle_release()
            self.release_tag = release_info.get('tag') or BUNDLE_RELEASE_TAG
            release_tag = self.release_tag

            # Fetch assets either from the release payload or via the tag endpoint
            assets = release_info.get('assets', []) or []
            if not assets and release_tag:
                api_url = f"{GITHUB_API_BASE}/repos/{self.github_repo}/releases/tags/{release_tag}"
                tag_response = requests.get(api_url, timeout=10)
                tag_response.raise_for_status()
                assets = tag_response.json().get('assets', [])

            if not release_tag:
                raise Exception("Latest AutoCAD plugin release tag is empty.")
            if not assets:
                raise Exception(
                    f"No assets found for AutoCAD plugin release/tag '{release_tag}'. Publish bundle zip assets.")

            statuses = []
            for asset in assets:
                asset_name = asset.get('name')
                if not asset_name or 'Source code' in asset_name or not asset_name.endswith('.zip'):
                    continue

                bundle_name = asset_name.replace(
                    f"-{release_tag}.zip", ".bundle")
                if bundle_name == asset_name:
                    bundle_name = asset_name.replace(".zip", ".bundle")
                bundle_path = os.path.join(
                    self.app_plugins_folder, bundle_name)

                is_installed = bundle_name in local_bundles
                local_version = None

                # Check for version.txt to determine if update is needed
                if is_installed:
                    version_file = os.path.join(bundle_path, 'version.txt')
                    if os.path.exists(version_file):
                        with open(version_file, 'r') as f:
                            local_version = f.read().strip()

                # Determine Status
                if not is_installed:
                    state = 'not_installed'
                elif local_version != release_tag:
                    # Installed, but version mismatch (or missing version file) -> Update
                    state = 'update_available'
                else:
                    state = 'installed'

                status = {
                    'name': bundle_name.replace('.bundle', ''),
                    'bundle_name': bundle_name,
                    'state': state,  # 'installed', 'not_installed', 'update_available'
                    'local_version': local_version or 'unknown',
                    'remote_version': release_tag,
                    'asset': asset
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

                print(f"DEBUG THREAD: Process started, reading output...")
                for line in iter(process.stdout.readline, ''):
                    line = line.strip()
                    print(f"DEBUG THREAD OUTPUT: {line}")
                    if line.startswith("PROGRESS:"):
                        message = line[len("PROGRESS:"):].strip()
                        js_message = json.dumps(message)
                        window.evaluate_js(
                            f'window.updateToolStatus("{tool_id}", {js_message})')

                process.stdout.close()
                return_code = process.wait()
                print(f"DEBUG THREAD: Process finished with return code {return_code}")

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

    def get_user_settings(self):
        """Reads and returns user settings from settings.json."""
        try:
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
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
                'workroomAutoSelectCadFiles': True
            }

    def save_user_settings(self, data):
        """Saves user settings to settings.json."""
        try:
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return {'status': 'success'}
        except Exception as e:
            logging.error(f"Error saving user settings: {e}")
            return {'status': 'error', 'message': str(e)}

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

    def process_email_with_ai(self, email_text, api_key, user_name, discipline):
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
        prompt = f"""
You are an intelligent assistant for {user_name}, a(n) {disciplines_str} engineering project manager. Your task is to analyze an email and extract specific project details. Focus ONLY on the primary {disciplines_str} engineering tasks mentioned. Ignore tasks for other disciplines.
Analyze the following email text and extract the information into a valid JSON object with the following keys: "id", "name", "due", "path", "deliverable", "tasks", "notes".
- "id": Find a project number or project ID (e.g., "250597", "P-12345", "Job #1042"). Look in the subject line, headers, and body. This could be called a job number, project number, project ID, or similar. If none, leave it empty.
- "name": Determine the project name, typically including the client and address or building name (e.g., "BofA, 22004 Sherman Way, Canoga Park, CA"). Include enough detail to uniquely identify the project. If no formal name is found, compose one from the client name and location mentioned.
- "due": Find the due date and format it as "MM/DD/YY". The current date is {current_date}. If the year is not specified in the email, assume the current year or the next year if the date would be in the past. Ensure the due date is on or after today. If multiple dates, choose the most relevant upcoming one.
- "path": Find the main project file path (e.g., "M:\\\\Gensler\\\\...").
- "deliverable": Infer the deliverable name from the email if possible. Prefer concise, standardized names when present (for example: DD60, DD90, CD60, CD90, CD100, CDF, RFI, RFI #2, Submittal, Lighting Submittal, Controls Submittal, Record Set, Record Drawings, IFP, Site Survey, Survey Report, ASR, ASR #2, PCC, PCC #3, Bulletin #2, Coordination, Meeting, Revision). If no deliverable is clear, leave it empty.
- "tasks": Create a JSON array of strings listing only the key {disciplines_str} engineering action items. Be concise. Examples: ["Update CAD per architect's comments", "Fill out permit forms", "Prepare binded CADs for IFP submission"].
- "notes": Provide a brief, one-sentence summary of the email's main request.
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

    def get_tasks(self):
        """Reads and returns the content of tasks.json."""
        try:
            with open(TASKS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
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

    def copy_template_to_folder(self, template_id, destination_folder, new_name=None, context=None, options=None):
        """Copies a template file to the specified folder with optional rename."""
        try:
            data = self.get_templates()
            template = next((t for t in data['templates'] if t['id'] == template_id), None)

            if not template:
                return {'status': 'error', 'message': 'Template not found'}

            source = template['sourcePath']
            if not os.path.exists(source):
                return {'status': 'error', 'message': f'Template file not found: {source}'}

            if not os.path.isdir(destination_folder):
                create_destination = False
                if isinstance(options, dict):
                    create_destination = bool(options.get('createDestination'))
                if create_destination:
                    os.makedirs(destination_folder, exist_ok=True)
                else:
                    return {'status': 'error', 'message': 'Destination folder does not exist'}

            # Determine output filename
            ext = os.path.splitext(source)[1]
            if new_name:
                # Ensure proper extension
                if not new_name.lower().endswith(ext.lower()):
                    new_name = new_name + ext
                dest_filename = new_name
            else:
                dest_filename = os.path.basename(source)

            dest_path = os.path.join(destination_folder, dest_filename)

            # Check if file exists and handle
            if os.path.exists(dest_path):
                return {'status': 'error', 'message': 'File already exists at destination'}

            shutil.copy2(source, dest_path)
            if context or options:
                _apply_template_context(dest_path, context=context, options=options)
            return {'status': 'success', 'path': dest_path}
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

    def export_timesheet_excel(self, data):
        """Exports timesheet data to an Excel file using a template."""
        try:
            import openpyxl
            from openpyxl.drawing.image import Image as XLImage
            from PIL import Image as PILImage

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

            # Row 1: Employee name (column M = 13)
            ws.cell(row=1, column=13, value=data.get('userName', 'Employee'))

            # Row 2: Week of (column M = 13)
            ws.cell(row=2, column=13, value=data.get('weekDisplay', ''))

            # Data rows start at row 5
            entries = data.get('entries', [])
            day_order = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun']

            for row_idx, entry in enumerate(entries, 5):
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

                # Column R (18): MILEAGE
                mileage = entry.get('mileage', 0)
                ws.cell(row=row_idx, column=18, value=mileage if mileage else '')

            # Totals row (row 26 in template based on CSV)
            totals = data.get('totals', {})
            totals_row = 26

            for day_idx, day in enumerate(day_order):
                ws.cell(row=totals_row, column=11 + day_idx, value=totals.get(day, 0))

            # Total mileage
            ws.cell(row=totals_row, column=18, value=totals.get('mileage', 0) or '')

            # =============== EXPENSE SHEET EXPORT ===============
            # If expense data is provided, also populate the expense sheet
            expense_data = data.get('expenses', {})
            expense_projects = expense_data.get('projects', [])
            temp_files_to_cleanup = []

            if expense_projects:
                sheet_name = "Project Expense Sheet"
                if sheet_name in wb.sheetnames:
                    ws_exp = wb[sheet_name]
                    mileage_rate = expense_data.get('mileageRate', 0.70)

                    HEADER_ROWS = 3
                    DEFAULT_DATA_ROWS = 5

                    current_row = 1
                    all_images = []

                    for proj_idx, project in enumerate(expense_projects):
                        exp_entries = project.get('entries', [])
                        num_entries = max(len(exp_entries), 1)

                        rows_needed = num_entries - DEFAULT_DATA_ROWS

                        if proj_idx > 0:
                            total_section_rows = HEADER_ROWS + max(num_entries, DEFAULT_DATA_ROWS) + 4
                            ws_exp.insert_rows(current_row, total_section_rows)

                        if rows_needed > 0 and proj_idx == 0:
                            ws_exp.insert_rows(current_row + HEADER_ROWS + DEFAULT_DATA_ROWS, rows_needed)

                        ws_exp.cell(row=current_row, column=1, value=f"PROJECT: {project.get('projectName', '')}")
                        ws_exp.cell(row=current_row + 1, column=1, value=f"JOB #: {project.get('projectId', '')}")

                        data_start_row = current_row + HEADER_ROWS
                        total_mileage = 0
                        total_expense = 0

                        for entry_idx, entry in enumerate(exp_entries):
                            row = data_start_row + entry_idx
                            ws_exp.cell(row=row, column=1, value=entry.get('date', ''))
                            ws_exp.cell(row=row, column=2, value=entry.get('description', ''))

                            exp_mileage = entry.get('mileage', 0) or 0
                            expense = entry.get('expense', 0) or 0

                            ws_exp.cell(row=row, column=4, value=exp_mileage if exp_mileage else '')
                            expense_cell = ws_exp.cell(row=row, column=5, value=expense if expense else '')
                            if expense:
                                expense_cell.number_format = '0.00'

                            total_mileage += exp_mileage
                            total_expense += expense

                        actual_data_rows = max(num_entries, DEFAULT_DATA_ROWS)
                        subtotal_row = data_start_row + actual_data_rows
                        rate_row = subtotal_row + 1
                        total_row = rate_row + 1

                        ws_exp.cell(row=subtotal_row, column=3, value="SUBTOTAL")
                        # Use SUM formulas so subtotals update when rows are added
                        last_data_row = subtotal_row - 1
                        ws_exp.cell(row=subtotal_row, column=4, value=f'=SUM(D{data_start_row}:D{last_data_row})')
                        subtotal_expense_cell = ws_exp.cell(row=subtotal_row, column=5, value=f'=SUM(E{data_start_row}:E{last_data_row})')
                        subtotal_expense_cell.number_format = '0.00'

                        ws_exp.cell(row=rate_row, column=3, value=f"{mileage_rate:.2f} CENTS PER MILE")
                        # Formula: mileage subtotal * rate
                        rate_mileage_cell = ws_exp.cell(row=rate_row, column=4, value=f'=D{subtotal_row}*{mileage_rate}')
                        rate_mileage_cell.number_format = '0.00'

                        ws_exp.cell(row=total_row, column=3, value="TOTAL")
                        # Formula: mileage reimbursement + expense subtotal
                        total_cell = ws_exp.cell(row=total_row, column=4, value=f'=D{rate_row}+E{subtotal_row}')
                        total_cell.number_format = '0.00'

                        for img in project.get('images', []):
                            all_images.append(img.get('path', ''))

                        current_row = total_row + 2

                    # Add images starting at row 45, spanning columns A-E
                    if all_images:
                        image_start_row = 45
                        # Width spanning A to E (approximately 5 columns ~500 pixels)
                        TARGET_WIDTH = 500
                        TARGET_HEIGHT = 375

                        for img_path in all_images:
                            if not os.path.exists(img_path):
                                continue

                            try:
                                # Use PIL to convert unsupported formats (like .mpo) to PNG
                                pil_img = PILImage.open(img_path)
                                # Convert to RGB if needed (handles RGBA, P mode, etc.)
                                if pil_img.mode in ('RGBA', 'P', 'LA'):
                                    pil_img = pil_img.convert('RGB')
                                # For MPO files, get the first frame
                                if hasattr(pil_img, 'n_frames') and pil_img.n_frames > 1:
                                    pil_img.seek(0)

                                # Save to temp file as PNG for openpyxl compatibility
                                temp_img_path = os.path.join(tempfile.gettempdir(), f"expense_img_{os.urandom(4).hex()}.png")
                                pil_img.save(temp_img_path, 'PNG')
                                temp_files_to_cleanup.append(temp_img_path)

                                img = XLImage(temp_img_path)

                                # Scale to fit within target dimensions while maintaining aspect ratio
                                scale = min(1.0, TARGET_WIDTH / img.width, TARGET_HEIGHT / img.height)
                                img.width = int(img.width * scale)
                                img.height = int(img.height * scale)

                                # Place image at column A, current row
                                ws_exp.add_image(img, f"A{image_start_row}")

                                # Calculate rows needed for this image (row height ~15 pixels)
                                rows_for_image = max(int(img.height / 15) + 2, 18)
                                image_start_row += rows_for_image

                            except Exception as img_err:
                                logging.warning(f"Could not add image {img_path}: {img_err}")

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

    def export_expense_sheet_excel(self, data):
        """Exports expense sheet data to an Excel file with images."""
        try:
            import openpyxl
            from openpyxl.drawing.image import Image as XLImage
            from PIL import Image as PILImage

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
            sheet_name = "Project Expense Sheet"
            if sheet_name not in wb.sheetnames:
                return {'status': 'error', 'message': f'Sheet "{sheet_name}" not found in template'}

            ws = wb[sheet_name]

            projects = data.get('projects', [])
            mileage_rate = data.get('mileageRate', 0.70)

            # Template structure per project section (11 rows each):
            # Row 1: PROJECT: [name]
            # Row 2: JOB #: [number]
            # Row 3: DATE, DESCRIPTION, MILEAGE, EXPENSE (headers)
            # Rows 4-8: 5 data rows
            # Row 9: SUBTOTAL
            # Row 10: mileage rate
            # Row 11: TOTAL
            # Row 12: blank separator

            ROWS_PER_SECTION = 12
            HEADER_ROWS = 3
            DEFAULT_DATA_ROWS = 5
            FOOTER_ROWS = 4  # subtotal, rate, total, blank

            current_row = 1
            all_images = []

            for proj_idx, project in enumerate(projects):
                entries = project.get('entries', [])
                num_entries = max(len(entries), 1)  # At least 1 row

                # Calculate if we need to insert extra rows
                rows_needed = num_entries - DEFAULT_DATA_ROWS

                if proj_idx > 0:
                    # Insert a new section for additional projects
                    # We need to insert rows for the entire section
                    total_section_rows = HEADER_ROWS + max(num_entries, DEFAULT_DATA_ROWS) + FOOTER_ROWS
                    ws.insert_rows(current_row, total_section_rows)

                # If this project has more entries than default, insert extra data rows
                if rows_needed > 0 and proj_idx == 0:
                    # Insert extra rows after the header rows (row 4 onwards for first section)
                    ws.insert_rows(current_row + HEADER_ROWS + DEFAULT_DATA_ROWS, rows_needed)

                # Row 1: PROJECT name
                ws.cell(row=current_row, column=1, value=f"PROJECT: {project.get('projectName', '')}")

                # Row 2: JOB #
                ws.cell(row=current_row + 1, column=1, value=f"JOB #: {project.get('projectId', '')}")

                # Row 3 is header row (DATE, DESCRIPTION, MILEAGE, EXPENSE) - already in template

                # Data rows start at row 4 of this section
                data_start_row = current_row + HEADER_ROWS
                total_mileage = 0
                total_expense = 0

                for entry_idx, entry in enumerate(entries):
                    row = data_start_row + entry_idx
                    ws.cell(row=row, column=1, value=entry.get('date', ''))
                    ws.cell(row=row, column=2, value=entry.get('description', ''))

                    mileage = entry.get('mileage', 0) or 0
                    expense = entry.get('expense', 0) or 0

                    ws.cell(row=row, column=4, value=mileage if mileage else '')
                    expense_cell = ws.cell(row=row, column=5, value=expense if expense else '')
                    if expense:
                        expense_cell.number_format = '0.00'

                    total_mileage += mileage
                    total_expense += expense

                # Calculate footer row positions
                actual_data_rows = max(num_entries, DEFAULT_DATA_ROWS)
                subtotal_row = data_start_row + actual_data_rows
                rate_row = subtotal_row + 1
                total_row = rate_row + 1

                # SUBTOTAL row - use SUM formulas so subtotals update when rows are added
                ws.cell(row=subtotal_row, column=3, value="SUBTOTAL")
                last_data_row = subtotal_row - 1
                ws.cell(row=subtotal_row, column=4, value=f'=SUM(D{data_start_row}:D{last_data_row})')
                subtotal_expense_cell = ws.cell(row=subtotal_row, column=5, value=f'=SUM(E{data_start_row}:E{last_data_row})')
                subtotal_expense_cell.number_format = '0.00'

                # Mileage rate row - formula: mileage subtotal * rate
                ws.cell(row=rate_row, column=3, value=f"{mileage_rate:.2f} CENTS PER MILE")
                rate_mileage_cell = ws.cell(row=rate_row, column=4, value=f'=D{subtotal_row}*{mileage_rate}')
                rate_mileage_cell.number_format = '0.00'

                # TOTAL row - formula: mileage reimbursement + expense subtotal
                ws.cell(row=total_row, column=3, value="TOTAL")
                total_cell = ws.cell(row=total_row, column=4, value=f'=D{rate_row}+E{subtotal_row}')
                total_cell.number_format = '0.00'

                # Collect images for this project
                for img in project.get('images', []):
                    all_images.append(img.get('path', ''))

                # Move to next section
                current_row = total_row + 2  # +2 for blank row after total

            # Add images starting at row 45, spanning columns A-E
            temp_files_to_cleanup = []
            if all_images:
                image_start_row = 45
                # Width spanning A to E (approximately 5 columns ~500 pixels)
                TARGET_WIDTH = 500
                TARGET_HEIGHT = 375

                for img_path in all_images:
                    if not os.path.exists(img_path):
                        continue

                    try:
                        # Use PIL to convert unsupported formats (like .mpo) to PNG
                        pil_img = PILImage.open(img_path)
                        # Convert to RGB if needed (handles RGBA, P mode, etc.)
                        if pil_img.mode in ('RGBA', 'P', 'LA'):
                            pil_img = pil_img.convert('RGB')
                        # For MPO files, get the first frame
                        if hasattr(pil_img, 'n_frames') and pil_img.n_frames > 1:
                            pil_img.seek(0)

                        # Save to temp file as PNG for openpyxl compatibility
                        temp_img_path = os.path.join(tempfile.gettempdir(), f"expense_img_{os.urandom(4).hex()}.png")
                        pil_img.save(temp_img_path, 'PNG')
                        temp_files_to_cleanup.append(temp_img_path)

                        img = XLImage(temp_img_path)

                        # Scale to fit within target dimensions while maintaining aspect ratio
                        scale = min(1.0, TARGET_WIDTH / img.width, TARGET_HEIGHT / img.height)
                        img.width = int(img.width * scale)
                        img.height = int(img.height * scale)

                        # Place image at column A, current row
                        ws.add_image(img, f"A{image_start_row}")

                        # Calculate rows needed for this image (row height ~15 pixels)
                        rows_for_image = max(int(img.height / 15) + 2, 18)
                        image_start_row += rows_for_image

                    except Exception as img_err:
                        logging.warning(f"Could not add image {img_path}: {img_err}")

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

    def open_url(self, url):
        """Opens a URL or protocol link using the OS default handler."""
        try:
            target = str(url or "").strip()
            if not target:
                return {'status': 'error', 'message': 'URL is required.'}

            if target.lower().startswith("file://"):
                local_path = self._coerce_local_email_path(target)
                if local_path:
                    return self.open_path(local_path)

            if sys.platform == "win32":
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
            return f"\\\\?\\UNC\\{normalized.lstrip('\\')}"
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
                return {
                    'status': 'error',
                    'message': f'Local project already exists: {local_project_path}'
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
        if isinstance(value, (list, tuple)):
            return [str(v) for v in value if v]
        if value:
            return [str(value)]
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

    def _coerce_local_email_path(self, raw):
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

        prompt = f"""
Analyze these electrical panel photos for Panel: {panel_name}.

You are provided with {num_breaker_imgs} images of the CIRCUIT BREAKERS (first {num_breaker_imgs} images)
and {num_dir_imgs} images of the CIRCUIT DIRECTORY (last {num_dir_imgs} images).

TASK 1: HEADER
- Extract Voltage, Bus Rating, Wire, Phase, Mounting, Enclosure.
- Look at the directory images or labels on the panel.

TASK 2: CIRCUITS & POLES
- Identify every breaker visible in the Breaker Images.
- CRITICAL: Determine the 'poles' (1, 2, or 3).
    - A 1-pole breaker takes up 1 circuit space.
    - A 2-pole breaker has a tied handle and takes up 2 vertical circuit spaces (e.g. 1 & 3).
    - A 3-pole breaker has a tied handle and takes up 3 vertical circuit spaces (e.g. 1, 3, & 5).
- Provide the circuit_number as the TOP-most circuit number the breaker occupies.
- Extract Amperage and Load Description from the labels (Breaker Images) or circuit directory (Directory Images).
- Resolve ditto marks (") if seen in the directory.

TASK 3: LOAD TYPES
- LIGHTING -> 'C', RECEPTACLES -> 'G', MOTORS/HVAC -> 'M', KITCHEN -> 'K', DEDICATED -> 'D'
""".strip()

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
                        raw_panel.get('breakerPaths') or raw_panel.get('breakerPath')
                    ),
                    'directory_paths': self._normalize_panel_schedule_paths(
                        raw_panel.get('directoryPaths') or raw_panel.get('directoryPath')
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
                    data.get('breakerPaths') or data.get('breakerPath')
                ),
                'directory_paths': self._normalize_panel_schedule_paths(
                    data.get('directoryPaths') or data.get('directoryPath')
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
                    raise ValueError("Both a breaker photo and a directory photo are required.")

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

    def _find_workroom_project_root_by_id(self, project_path):
        normalized_project_path = os.path.normpath(project_path) if project_path else ''
        if not normalized_project_path:
            return ''

        current = normalized_project_path.rstrip("\\/")
        if not current:
            return ''

        while current:
            folder_name = os.path.basename(current).strip()
            if re.match(r'^\d{6}(?!\d)', folder_name):
                logging.info(
                    "_find_workroom_project_root_by_id: matched %s from %s",
                    current,
                    normalized_project_path,
                )
                return current

            parent = os.path.dirname(current)
            if not parent:
                break
            if os.path.normcase(parent) == os.path.normcase(current):
                break
            current = parent.rstrip("\\/") or parent

        return ''

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
            return None
        if not self._is_workroom_auto_select_enabled(settings, launch_context):
            return None
        context = self._resolve_workroom_context(settings, launch_context)
        project_path = context.get('project_path') or ''
        discipline = context.get('discipline') or 'Electrical'
        if not project_path:
            logging.info(
                f"{tool_name}: Workroom auto-select fallback to manual file picker (missing project path).")
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
            return None
        dwg_files = self._list_base_level_dwgs(discipline_folder)
        if not dwg_files:
            logging.info(
                f"{tool_name}: Workroom auto-select fallback to manual file picker (no DWGs in {discipline_folder}).")
            return None
        files_list_path = self._write_files_list_temp(dwg_files)
        logging.info(
            f"{tool_name}: Auto-selected {len(dwg_files)} DWG(s) from {discipline_folder} "
            f"(mode={folder_resolution.get('mode')}).")
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

            if not project_path:
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
        command = (
            f'powershell.exe -ExecutionPolicy Bypass -File "{script_path}" '
            f'-AcadCore "{acad_path}" '
            f'-AutoDetectPaperSize {_ps_bool(auto_detect)} '
            f'-ShrinkPercent {shrink_percent}'
        )
        if auto_selection:
            command += f' -FilesListPath "{auto_selection["files_list_path"]}"'
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
        command = (
            f'powershell.exe -ExecutionPolicy Bypass -File "{script_path}" '
            f'-AcadCore "{acad_path}" '
            f'-ScanAllLayers {_ps_bool(scan_all)}'
        )
        if auto_selection:
            command += f' -FilesListPath "{auto_selection["files_list_path"]}"'
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
        command = (
            f'powershell.exe -ExecutionPolicy Bypass -File "{script_path}" '
            f'-AcadCore "{acad_path}" '
            f'-ScanAllLayers {_ps_bool(scan_all)}'
        )
        if auto_selection:
            command += f' -FilesListPath "{auto_selection["files_list_path"]}"'
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
