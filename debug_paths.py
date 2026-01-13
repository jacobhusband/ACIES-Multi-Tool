import os
import sys
from pathlib import Path

# Add the project dir to path
project_dir = r"c:\Users\JacobH\Documents\dev\ProjectManagement"
sys.path.append(project_dir)

import main

print(f"Platform: {sys.platform}")
print(f"APPDATA Env: {os.getenv('APPDATA')}")
print(f"TASKS_FILE: {main.TASKS_FILE}")
print(f"NOTES_FILE: {main.NOTES_FILE}")
print(f"SETTINGS_FILE: {main.SETTINGS_FILE}")
print(f"TIMESHEETS_FILE: {main.TIMESHEETS_FILE}")

api = main.Api()
print(f"Api timesheets info: {api.get_timesheets_info()}")
