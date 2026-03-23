import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
INDEX_HTML_PATH = REPO_ROOT / "index.html"
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class OutlookScanUiTests(unittest.TestCase):
    def test_email_intake_markup_exists_and_legacy_markup_is_removed(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")

        self.assertIn('id="outlookScanBtn"', html)
        self.assertIn('id="outlookScanDlg"', html)
        self.assertIn("Email Intake", html)
        self.assertIn('id="emailIntakeModeGroup"', html)
        self.assertIn('id="emailIntakeModePaste"', html)
        self.assertIn('id="emailIntakeModeScan"', html)
        self.assertIn('id="emailIntakePastePanel"', html)
        self.assertIn('id="emailIntakeScanPanel"', html)
        self.assertIn('id="emailArea"', html)
        self.assertIn('id="btnProcessEmail"', html)
        self.assertIn('id="outlookScanCapabilityStatus"', html)
        self.assertIn('id="outlookScanCapabilityDetails"', html)
        self.assertIn('id="outlookScanDate"', html)
        self.assertIn('id="outlookScanDate" type="date"', html)
        self.assertIn('id="outlookScanRunBtn"', html)
        self.assertIn('id="outlookScanProgress"', html)
        self.assertIn('id="outlookScanReport"', html)
        self.assertIn('id="outlookScanReportSummary"', html)
        self.assertIn('id="outlookScanReportLog"', html)
        self.assertIn('id="outlookScanSuggestions"', html)
        self.assertIn('id="outlookScanSkipped"', html)

        self.assertNotIn('id="aiBtn"', html)
        self.assertNotIn('id="emailDlg"', html)
        self.assertNotIn('id="outlookScanConnectBtn"', html)
        self.assertNotIn('id="outlookScanSignOutBtn"', html)
        self.assertNotIn('id="settings_microsoftAuthStatus"', html)
        self.assertNotIn('id="settings_microsoftAuthActionBtn"', html)

    def test_email_intake_script_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn("const DEFAULT_OUTLOOK_SCAN_CAPABILITY = {", script)
        self.assertIn('mode: "paste"', script)
        self.assertIn('function normalizeEmailIntakeMode(value = "") {', script)
        self.assertIn('function normalizeOutlookScanCapability(raw = {}) {', script)
        self.assertIn("function renderOutlookScanUi() {", script)
        self.assertIn(
            "window.pywebview.api.get_outlook_scan_capability()",
            script,
        )
        self.assertIn("async function openOutlookScanDialog(mode = outlookScanState.mode) {", script)
        self.assertIn("async function processEmailIntakePaste() {", script)
        self.assertIn('closeDlg("outlookScanDlg");', script)
        self.assertIn("handleAiProjectResult(res.data || {});", script)
        self.assertIn('outlookScanBtn.onclick = () => openOutlookScanDialog();', script)
        self.assertIn('btnProcessEmail.onclick = () => processEmailIntakePaste();', script)
        self.assertIn('setEmailIntakeMode("paste")', script)
        self.assertIn('setEmailIntakeMode("scan")', script)
        self.assertIn("window.pywebview.api.scan_outlook_inbox(", script)
        self.assertIn("window.pywebview.api.open_outlook_desktop_message(", script)

        self.assertNotIn("DEFAULT_MICROSOFT_AUTH_STATE", script)
        self.assertNotIn("renderMicrosoftAuthUi", script)
        self.assertNotIn("loadMicrosoftAuthState", script)
        self.assertNotIn("handleMicrosoftSignIn", script)
        self.assertNotIn("handleMicrosoftSignOut", script)
        self.assertNotIn('closeDlg("emailDlg")', script)

    def test_email_intake_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        self.assertIn(".email-intake-mode-switch {", css)
        self.assertIn(".email-intake-mode-option {", css)
        self.assertIn(".email-intake-mode-label {", css)
        self.assertIn(".email-intake-panel[hidden] {", css)
        self.assertIn(".email-intake-panel-head {", css)
        self.assertIn(".email-intake-scan-status {", css)
        self.assertIn(".outlook-scan-toolbar {", css)
        self.assertIn(".outlook-scan-results {", css)
        self.assertIn(".outlook-scan-report summary {", css)
        self.assertIn(".outlook-scan-skipped-item {", css)


if __name__ == "__main__":
    unittest.main()
