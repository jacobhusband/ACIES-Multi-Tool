import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
INDEX_HTML_PATH = REPO_ROOT / "index.html"
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ActivityTrayUiTests(unittest.TestCase):
    def test_activity_tray_markup_exists(self):
        text = INDEX_HTML_PATH.read_text(encoding="utf-8")

        self.assertIn('id="activityTray"', text)
        self.assertIn('id="activityTrayToggle"', text)
        self.assertIn('id="activityTrayCounts"', text)
        self.assertIn('id="activityTrayBody"', text)
        self.assertIn('id="activityTrayEmpty"', text)
        self.assertIn('id="activityTrayList"', text)
        self.assertIn("Activity", text)
        self.assertIn("No activity yet.", text)

    def test_activity_tray_styles_exist(self):
        text = STYLES_CSS_PATH.read_text(encoding="utf-8")

        self.assertIn(".activity-tray", text)
        self.assertIn(".activity-tray.is-collapsed .activity-tray-body", text)
        self.assertIn(".activity-card", text)
        self.assertIn(".activity-card-progress-bar", text)
        self.assertIn(".activity-card-action.accept", text)

    def test_activity_tray_script_helpers_exist(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn("const activityTrayState = {", text)
        self.assertIn("function initActivityTray() {", text)
        self.assertIn("function beginActivity({", text)
        self.assertIn("function updateActivity(activityId, patch = {}) {", text)
        self.assertIn("function completeActivity(activityId, patch = {}) {", text)
        self.assertIn("function failActivity(activityId, patch = {}) {", text)
        self.assertIn("function acceptActivity(activityId) {", text)
        self.assertIn("function renderActivityTray() {", text)
        self.assertIn('textContent: "Accept"', text)
        self.assertIn('"data-activity-action": "open"', text)


if __name__ == "__main__":
    unittest.main()
