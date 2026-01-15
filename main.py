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
        # Use Documents folder - it is NOT virtualized by Windows Store Python
        # This ensures both dev (python main.py) and packaged apps use the same location
        user_profile = os.getenv('USERPROFILE')
        if user_profile:
            base_dir = os.path.join(user_profile, 'Documents')
        else:
            base_dir = os.path.expanduser('~')
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

DEFAULT_TEMPLATES = [
    {
        "name": "Narrative of Changes",
        "discipline": "General",
        "fileType": "docx",
        "sourcePath": r"C:\Users\JacobH\OneDrive - ACIES Engineering\Documents\Company Files\Narrative of Changes.docx",
        "description": "Standard narrative of changes template with MEP sections"
    },
    {
        "name": "Plan Check Comments",
        "discipline": "General",
        "fileType": "doc",
        "sourcePath": r"C:\Users\JacobH\OneDrive - ACIES Engineering\Documents\Company Files\PCC.doc",
        "description": "Plan check comments response table template"
    }
]

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

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

                for line in iter(process.stdout.readline, ''):
                    line = line.strip()
                    if line.startswith("PROGRESS:"):
                        message = line[len("PROGRESS:"):].strip()
                        js_message = json.dumps(message)
                        window.evaluate_js(
                            f'window.updateToolStatus("{tool_id}", {js_message})')

                process.stdout.close()
                return_code = process.wait()

                if return_code == 0:
                    window.evaluate_js(
                        f'window.updateToolStatus("{tool_id}", "DONE")')
                else:
                    error_message = f"Script finished with error code {return_code}."
                    js_error = json.dumps(error_message)
                    window.evaluate_js(
                        f'window.updateToolStatus("{tool_id}", "ERROR: " + {js_error})')

            except Exception as e:
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
            return {'userName': '', 'discipline': ['Electrical'], 'apiKey': '', 'autocadPath': '', 'showSetupHelp': True}

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

        self._ensure_aiohttp()
        current_date = datetime.date.today().strftime("%m/%d/%Y")
        disciplines_str = ', '.join(discipline) if isinstance(
            discipline, list) else discipline
        prompt = f"""
You are an intelligent assistant for {user_name}, a(n) {disciplines_str} engineering project manager. Your task is to analyze an email and extract specific project details. Focus ONLY on the primary {disciplines_str} engineering tasks mentioned. Ignore tasks for other disciplines.
Analyze the following email text and extract the information into a valid JSON object with the following keys: "id", "name", "due", "path", "tasks", "notes".
- "id": Find a project number (e.g., "250597"). If none, leave it empty.
- "name": Determine the project name, typically including the client and address (e.g., "BofA, 22004 Sherman Way, Canoga Park, CA").
- "due": Find the due date and format it as "MM/DD/YY". The current date is {current_date}. If the year is not specified in the email, assume the current year or the next year if the date would be in the past. Ensure the due date is on or after today. If multiple dates, choose the most relevant upcoming one.
- "path": Find the main project file path (e.g., "M:\\\\Gensler\\\\...").
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
            client = genai.Client(api_key=final_api_key)
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

            cleaned = (response.text or "").strip()
            project_data = json.loads(cleaned)

            project_data.setdefault("id", "")
            project_data.setdefault("name", "")
            project_data.setdefault("due", "")
            project_data.setdefault("path", "")
            project_data.setdefault("tasks", [])
            project_data.setdefault("notes", "")

            if 'tasks' in project_data and isinstance(project_data['tasks'], list):
                project_data['tasks'] = [{'text': str(task), 'done': False, 'links': [
                ]} for task in project_data['tasks']]

            return {'status': 'success', 'data': project_data}

        except Exception as e:
            msg = str(e)
            lower = msg.lower()
            if ("api key expired" in lower or
                "api_key_invalid" in lower or
                    "invalid api key" in lower):
                return {'status': 'error',
                        'message': ('Your Google API key is expired/invalid. '
                                    'Create a new key in Google AI Studio → API keys, '
                                    'update your settings, then try again.')}
            logging.error(f"Error processing email with AI: {e}")
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

            return data
        except (FileNotFoundError, json.JSONDecodeError):
            # First run - create initial structure with defaults
            data = {'templates': [], 'defaultTemplatesInstalled': False, 'lastModified': None}
            data = self._install_default_templates(data)
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

    def copy_template_to_folder(self, template_id, destination_folder, new_name=None):
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
            return {'status': 'success', 'path': dest_path}
        except Exception as e:
            logging.error(f"Error copying template: {e}")
            return {'status': 'error', 'message': str(e)}

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

    def export_timesheet_excel(self, data):
        """Exports timesheet data to an Excel file using a template."""
        try:
            import openpyxl
            from openpyxl.drawing.image import Image as XLImage
            from PIL import Image as PILImage

            # Get the template path (relative to this script)
            script_dir = Path(__file__).resolve().parent
            template_path = script_dir / "templates" / "Template_Timesheet.xlsx"

            if not template_path.exists():
                return {'status': 'error', 'message': f'Template file not found: {template_path}'}

            # Determine output path
            week_key = data.get('weekKey', 'timesheet')
            user_docs = os.path.join(os.path.expanduser('~'), 'Documents')
            file_path = os.path.join(user_docs, f"Timesheet_{week_key}.xlsx")

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

            # Get the template path
            script_dir = Path(__file__).resolve().parent
            template_path = script_dir / "templates" / "Template_Timesheet.xlsx"

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

    def run_publish_script(self):
        """Runs the PlotDWGs.ps1 PowerShell script with progress updates."""
        script_path = os.path.join(BASE_DIR, "PlotDWGs.ps1")
        if not os.path.exists(script_path):
            raise Exception("PlotDWGs.ps1 not found in application directory.")
        settings = self.get_user_settings()
        acad_path = settings.get('autocadPath', '')
        if not acad_path:
            raise Exception("No AutoCAD version selected in settings.")
        command = f'powershell.exe -ExecutionPolicy Bypass -File "{script_path}" -AcadCore "{acad_path}"'
        self._run_script_with_progress(command, 'toolPublishDwgs')
        return {'status': 'success'}

    def run_freeze_layers_script(self):
        """Runs the FreezeLayersDWGs.ps1 PowerShell script with progress updates."""
        script_path = os.path.join(BASE_DIR, "FreezeLayersDWGs.ps1")
        if not os.path.exists(script_path):
            raise Exception(
                "FreezeLayersDWGs.ps1 not found in application directory.")
        settings = self.get_user_settings()
        acad_path = settings.get('autocadPath', '')
        if not acad_path:
            raise Exception("No AutoCAD version selected in settings.")
        command = f'powershell.exe -ExecutionPolicy Bypass -File "{script_path}" -AcadCore "{acad_path}"'
        self._run_script_with_progress(command, 'toolFreezeLayers')
        return {'status': 'success'}

    def run_thaw_layers_script(self):
        """Runs the ThawLayersDWGs.ps1 PowerShell script with progress updates."""
        script_path = os.path.join(BASE_DIR, "ThawLayersDWGs.ps1")
        if not os.path.exists(script_path):
            raise Exception(
                "ThawLayersDWGs.ps1 not found in application directory.")
        settings = self.get_user_settings()
        acad_path = settings.get('autocadPath', '')
        if not acad_path:
            raise Exception("No AutoCAD version selected in settings.")
        command = f'powershell.exe -ExecutionPolicy Bypass -File "{script_path}" -AcadCore "{acad_path}"'
        self._run_script_with_progress(command, 'toolThawLayers')
        return {'status': 'success'}

    def run_clean_xrefs_script(self):
        """Runs the removeXREFPaths.ps1 PowerShell script with progress updates."""
        script_path = os.path.join(BASE_DIR, "removeXREFPaths.ps1")
        if not os.path.exists(script_path):
            raise Exception(
                "removeXREFPaths.ps1 not found in application directory.")
        settings = self.get_user_settings()
        acad_path = settings.get('autocadPath', '')
        if not acad_path:
            raise Exception("No AutoCAD version selected in settings.")
        command = f'powershell.exe -ExecutionPolicy Bypass -File "{script_path}" -AcadCore "{acad_path}"'
        self._run_script_with_progress(command, 'toolCleanXrefs')
        return {'status': 'success'}

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
