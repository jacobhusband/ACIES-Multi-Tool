import os
import sys
from pathlib import Path

# Add the project dir to path
project_dir = r"c:\Users\JacobH\Documents\dev\ProjectManagement"
sys.path.append(project_dir)

import main

api = main.Api()
result = api.write_timesheets_test_file()
print(f"Test File Write Result: {result}")
