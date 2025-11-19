import webview
import json
import os
import subprocess
import sys
import shutil
import csv
import io
import logging
from pathlib import Path
from dotenv import load_dotenv
import datetime
import threading
import requests  # Added for GitHub API calls
import zipfile   # Added for extracting bundles

# --- Google GenAI (new client) ---
# Uses GOOGLE_API_KEY from environment/.env
from google import genai
from google.genai import types

# --- Helper function to get the application data path ---


def get_app_data_path(file_name="tasks.json"):
    """
    Determines the correct, cross-platform path for storing user data
    and returns the full path for the given file_name.
    """
    if sys.platform == "win32":
        base_dir = os.getenv('APPDATA')
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
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# --- API Class ---


class Api:
    # --- FIX: Removed self.window from __init__ to prevent circular reference ---
    def __init__(self):
        # --- Configuration for Bundle Management ---
        self.github_repo = "jacobhusband/ElectricalCommands"
        self.release_tag = "v0.0.0"

        appdata_path = os.getenv('APPDATA')
        if not appdata_path:
            raise EnvironmentError(
                "CRITICAL: Could not determine the APPDATA directory.")
        self.app_plugins_folder = os.path.join(
            appdata_path, 'Autodesk', 'ApplicationPlugins')
        os.makedirs(self.app_plugins_folder, exist_ok=True)

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

            # 2. Get remote assets
            # NOTE: Ensure your repo is public or you have the token logic here
            api_url = f"https://api.github.com/repos/{self.github_repo}/releases/tags/{self.release_tag}"
            response = requests.get(api_url, timeout=10)
            response.raise_for_status()
            assets = response.json().get('assets', [])

            statuses = []
            for asset in assets:
                asset_name = asset.get('name')
                if not asset_name or 'Source code' in asset_name or not asset_name.endswith('.zip'):
                    continue

                bundle_name = asset_name.replace(
                    f"-{self.release_tag}.zip", ".bundle")
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
                elif local_version != self.release_tag:
                    # Installed, but version mismatch (or missing version file) -> Update
                    state = 'update_available'
                else:
                    state = 'installed'

                status = {
                    'name': bundle_name.replace('.bundle', ''),
                    'bundle_name': bundle_name,
                    'state': state,  # 'installed', 'not_installed', 'update_available'
                    'local_version': local_version or 'unknown',
                    'remote_version': self.release_tag,
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
        bundle_name = asset_name.replace(f"-{self.release_tag}.zip", ".bundle")
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
                f.write(self.release_tag)

            logging.info(
                f"Successfully installed {bundle_name} version {self.release_tag}")
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
            return {'userName': '', 'discipline': 'Electrical', 'apiKey': ''}

    def save_user_settings(self, data):
        """Saves user settings to settings.json."""
        try:
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return {'status': 'success'}
        except Exception as e:
            logging.error(f"Error saving user settings: {e}")
            return {'status': 'error', 'message': str(e)}

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

            client = genai.Client(api_key=api_key)
            model = "gemini-2.5-pro"

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
        prompt = f"""
You are an intelligent assistant for {user_name}, a(n) {discipline} engineering project manager. Your task is to analyze an email and extract specific project details. Focus ONLY on the primary {discipline} engineering tasks mentioned. Ignore tasks for other disciplines.
Analyze the following email text and extract the information into a valid JSON object with the following keys: "id", "name", "due", "path", "tasks", "notes".
- "id": Find a project number (e.g., "250597"). If none, leave it empty.
- "name": Determine the project name, typically including the client and address (e.g., "BofA, 22004 Sherman Way, Canoga Park, CA").
- "due": Find the due date and format it as "MM/DD/YY". The current date is {current_date}. If the year is not specified in the email, assume the current year or the next year if the date would be in the past. Ensure the due date is on or after today. If multiple dates, choose the most relevant upcoming one.
- "path": Find the main project file path (e.g., "M:\\\\Gensler\\\\...").
- "tasks": Create a JSON array of strings listing only the key {discipline} engineering action items. Be concise. Examples: ["Update CAD per architect's comments", "Fill out permit forms", "Prepare binded CADs for IFP submission"].
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
            model = "gemini-2.5-pro"

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

    def run_publish_script(self):
        """Runs the PlotDWGs.ps1 PowerShell script with progress updates."""
        script_path = os.path.join(BASE_DIR, "PlotDWGs.ps1")
        if not os.path.exists(script_path):
            raise Exception("PlotDWGs.ps1 not found in application directory.")
        command = f'powershell.exe -ExecutionPolicy Bypass -File "{script_path}"'
        self._run_script_with_progress(command, 'toolPublishDwgs')
        return {'status': 'success'}

    def run_clean_xrefs_script(self):
        """Runs the removeXREFPaths.ps1 PowerShell script with progress updates."""
        script_path = os.path.join(BASE_DIR, "removeXREFPaths.ps1")
        if not os.path.exists(script_path):
            raise Exception(
                "removeXREFPaths.ps1 not found in application directory.")
        command = f'powershell.exe -ExecutionPolicy Bypass -File "{script_path}"'
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

    def abort_clean_dwgs(self):
        """Creates an abort signal file to stop the Clean DWGs process."""
        try:
            abort_file = os.path.join(os.environ.get(
                'TEMP', ''), "abort_cleandwgs.flag")
            with open(abort_file, 'w') as f:
                f.write('abort')
            logging.info("Abort signal created for Clean DWGs process")
            return {'status': 'success'}
        except Exception as e:
            logging.error(f"Error creating abort file: {e}")
            return {'status': 'error', 'message': str(e)}

    def run_clean_dwgs_script(self, data):
        """Prepares files and runs the CleanDWGs.ps1 PowerShell script."""

        required_bundle = "ElectricalCommands.CleanCADCommands.bundle"
        bundle_path = os.path.join(self.app_plugins_folder, required_bundle)

        if not os.path.isdir(bundle_path):
            logging.warning(
                f"Prerequisite not met: '{required_bundle}' is not installed.")
            return {
                'status': 'prerequisite_failed',
                'message': f'The "{required_bundle.replace(".bundle", "")}" bundle is required. Please install it from the manager above.'
            }

        try:
            titleblock_path = data.get('titleblock')
            selected_disciplines = data.get('disciplines', [])

            if not titleblock_path:
                raise ValueError("Missing titleblock path.")

            # Resolve paths
            titleblock_full = Path(titleblock_path).resolve()
            titleblock_parent = titleblock_full.parent
            project_root = titleblock_parent.parent

            if not project_root or not project_root.exists():
                raise ValueError(
                    "Could not determine project root from titleblock path.")

            # Create output directory with timestamp
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            project_name = project_root.name
            documents_folder = os.path.join(
                os.path.expanduser('~'), 'Documents')
            output_root = os.path.join(
                documents_folder, "AutoCAD Clean DWGs", f"{project_name}_{timestamp}")
            os.makedirs(output_root, exist_ok=True)

            # Determine all folders to copy
            folders_to_copy = set()
            folders_to_copy.add(titleblock_parent.name)  # e.g., "Xref"
            # e.g., "Electrical", "Mechanical"
            folders_to_copy.update(selected_disciplines)

            # Copy the entire contents of each required folder
            for folder in folders_to_copy:
                src_folder = project_root / folder
                if src_folder.exists() and src_folder.is_dir():
                    dest_folder = Path(output_root) / folder
                    shutil.copytree(src_folder, dest_folder,
                                    dirs_exist_ok=True)
                    logging.info(
                        f"Copied folder: {src_folder} to {dest_folder}")
                else:
                    logging.warning(
                        f"Folder not found or is not a directory: {src_folder}")

            # Determine the path of the titleblock in the new copied location
            new_titleblock_path = str(
                Path(output_root) / titleblock_full.relative_to(project_root))

            # Build the PowerShell command
            script_path = os.path.join(BASE_DIR, "CleanDWGs.ps1")
            if not os.path.exists(script_path):
                raise FileNotFoundError("CleanDWGs.ps1 not found.")

            # We only need to pass the discipline names, not the Xref folder name
            disciplines_arg = ','.join(sorted(selected_disciplines))

            command = [
                'powershell.exe',
                '-ExecutionPolicy', 'Bypass',
                '-File', script_path,
                '-TitleblockPath', new_titleblock_path,
                '-Disciplines', disciplines_arg,
                '-OutputFolder', output_root,
                # Pass original root for context if needed
                '-ProjectRoot', str(project_root)
            ]

            self._run_script_with_progress(command, 'toolCleanDwgs')
            return {'status': 'success'}

        except Exception as e:
            logging.error(f"Error in run_clean_dwgs_script: {e}")
            window = webview.windows[0]
            js_error = json.dumps(str(e))
            window.evaluate_js(
                f'window.updateToolStatus("toolCleanDwgs", "ERROR: " + {js_error})')
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
