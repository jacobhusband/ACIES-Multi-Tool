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
        self.assertIn('id="outlookScanDate"', html)
        self.assertIn('id="outlookScanDate" type="date"', html)
        self.assertIn('id="outlookScanConnectBtn"', html)
        self.assertIn('id="outlookScanSignOutBtn"', html)
        self.assertIn('id="outlookScanRunBtn"', html)
        self.assertIn('id="outlookScanProgress"', html)
        self.assertIn('id="outlookScanReport"', html)
        self.assertIn('id="outlookScanReportSummary"', html)
        self.assertIn('id="outlookScanReportLog"', html)
        self.assertIn('id="outlookScanSuggestions"', html)
        self.assertIn('id="outlookScanSkipped"', html)
        self.assertIn('id="settings_microsoftAuthStatus"', html)
        self.assertIn('id="settings_microsoftAuthActionBtn"', html)

    def test_outlook_scan_script_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn("const DEFAULT_MICROSOFT_AUTH_STATE = {", script)
        self.assertIn("const DEFAULT_OUTLOOK_SCAN_STATE = {", script)
        self.assertIn("function getTodayLocalDateInputValue() {", script)
        self.assertIn("function normalizeOutlookScanDateInput(value = \"\") {", script)
        self.assertIn("function createEmptyOutlookScanProgress() {", script)
        self.assertIn("function normalizeOutlookScanProgress(raw = {}, previous = null) {", script)
        self.assertIn("function normalizeMicrosoftAuthState(raw = {}) {", script)
        self.assertIn("function getOutlookScanSourceLabel(source = \"\") {", script)
        self.assertIn("function isMicrosoftAdminConsentErrorMessage(message = \"\") {", script)
        self.assertIn("function isMicrosoftConsentPolicyErrorMessage(message = \"\") {", script)
        self.assertIn("function handleMicrosoftConsentDiagnostic() {", script)
        self.assertIn("function handleMicrosoftAdminConsent() {", script)
        self.assertIn("window.updateOutlookScanProgress = function (payload = {}) {", script)
        self.assertIn("function renderOutlookScanUi() {", script)
        self.assertIn("function renderMicrosoftAuthUi() {", script)
        self.assertIn("function formatOutlookScanSelectedDayLabel(scanDate = \"\", timeframe = \"\") {", script)
        self.assertIn("function getOutlookScanDayRange(scanDate = \"\") {", script)
        self.assertIn("function buildOutlookScanProjectContext(scanDate = \"\") {", script)
        self.assertIn("scanDateInput.max = getTodayLocalDateInputValue();", script)
        self.assertIn("return due && due >= start && due <= end;", script)
        self.assertIn("function buildOutlookScanReportModel() {", script)
        self.assertIn("function buildOutlookEmailRefFromMessage(message = {}) {", script)
        self.assertIn("function openOutlookScanMessage(message = {}) {", script)
        self.assertIn("function buildOutlookScanDerivedState(result = null) {", script)
        self.assertIn("function runOutlookInboxScan() {", script)
        self.assertIn("function acceptOutlookScanSuggestion(suggestionKey) {", script)
        self.assertIn("toolbarBtn.disabled = false;", script)
        self.assertIn('Microsoft sign-in is in progress', script)
        self.assertIn('microsoftAdminConsentRequired', script)
        self.assertIn('microsoftConsentDiagnosticRecommended', script)
        self.assertIn('window.pywebview.api.diagnose_microsoft_consent_policy()', script)
        self.assertIn('"Retry minimal consent"', script)
        self.assertIn('window.pywebview.api.request_microsoft_admin_consent()', script)
        self.assertIn('"Admin approval"', script)
        self.assertIn("window.pywebview.api.scan_outlook_inbox(", script)
        self.assertIn("const projectContext = buildOutlookScanProjectContext(scanDate);", script)
        self.assertIn("window.pywebview.api.open_outlook_desktop_message(", script)
        self.assertIn("messageId = String(", script)
        self.assertIn("internetMessageId = String(", script)
        self.assertIn("if (messageId || internetMessageId) {", script)
        self.assertIn('source: "outlook-url"', script)
        self.assertIn('source: "outlook-desktop"', script)
        self.assertIn('"Using Desktop Outlook"', script)
        self.assertIn('"Desktop Outlook available"', script)
        self.assertIn("emailsIncludedCount", script)
        self.assertIn("deliverablesIncludedCount", script)
        self.assertIn("threadsDetected", script)
        self.assertIn("dedupedEmailCount", script)
        self.assertIn("dedupeSkippedEmailCount", script)
        self.assertIn('"Threads detected"', script)
        self.assertIn('"Emails shortened"', script)
        self.assertIn('"Duplicate emails skipped"', script)
        self.assertIn('"Deliverables on day"', script)
        self.assertIn('"Choose a day, then run a scan."', script)
        self.assertIn('"Day"', script)
        self.assertIn("promptTruncated", script)
        self.assertIn("progressLog", script)
        self.assertIn("reportOpen", script)
        self.assertIn("reportStats", script)
        self.assertNotIn("result?.items", script)
        self.assertNotIn("analysisStatus", script)

    def test_outlook_scan_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        self.assertIn(".outlook-scan-toolbar {", css)
        self.assertIn(".outlook-scan-results {", css)
        self.assertIn(".outlook-scan-progress,", css)
        self.assertIn(".outlook-scan-report summary {", css)
        self.assertIn(".outlook-scan-report-summary {", css)
        self.assertIn(".outlook-scan-log-item {", css)
        self.assertIn(".outlook-scan-card,", css)
        self.assertIn(".outlook-scan-related-list,", css)
        self.assertIn(".outlook-scan-skipped-item {", css)


if __name__ == "__main__":
    unittest.main()
