import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
INDEX_HTML_PATH = REPO_ROOT / "index.html"
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class OutlookScanUiTests(unittest.TestCase):
    def test_outlook_scan_markup_exists(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")

        self.assertIn('id="outlookScanBtn"', html)
        self.assertIn('id="outlookScanDlg"', html)
        self.assertIn('id="outlookScanTimeframe"', html)
        self.assertIn('id="outlookScanConnectBtn"', html)
        self.assertIn('id="outlookScanSignOutBtn"', html)
        self.assertIn('id="outlookScanRunBtn"', html)
        self.assertIn('id="outlookScanSuggestions"', html)
        self.assertIn('id="outlookScanSkipped"', html)
        self.assertIn('id="settings_microsoftAuthStatus"', html)
        self.assertIn('id="settings_microsoftAuthActionBtn"', html)

    def test_outlook_scan_script_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn("const DEFAULT_MICROSOFT_AUTH_STATE = {", script)
        self.assertIn("const DEFAULT_OUTLOOK_SCAN_STATE = {", script)
        self.assertIn("function normalizeMicrosoftAuthState(raw = {}) {", script)
        self.assertIn("function renderOutlookScanUi() {", script)
        self.assertIn("function renderMicrosoftAuthUi() {", script)
        self.assertIn("function buildOutlookEmailRefFromMessage(message = {}) {", script)
        self.assertIn("function buildOutlookScanDerivedState(result = null) {", script)
        self.assertIn("function runOutlookInboxScan() {", script)
        self.assertIn("function acceptOutlookScanSuggestion(suggestionKey) {", script)
        self.assertIn("window.pywebview.api.scan_outlook_inbox(", script)
        self.assertIn("messageId = String(", script)
        self.assertIn("internetMessageId = String(", script)
        self.assertIn("if (messageId || internetMessageId) {", script)
        self.assertIn('source: "outlook-url"', script)

    def test_outlook_scan_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        self.assertIn(".outlook-scan-toolbar {", css)
        self.assertIn(".outlook-scan-results {", css)
        self.assertIn(".outlook-scan-card,", css)
        self.assertIn(".outlook-scan-related-list,", css)
        self.assertIn(".outlook-scan-skipped-item {", css)


if __name__ == "__main__":
    unittest.main()
