import pathlib
import unittest

ROOT = pathlib.Path(__file__).resolve().parents[1]


class EmailCaptureUiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.html = (ROOT / "index.html").read_text(encoding="utf-8")
        cls.script = (ROOT / "script.js").read_text(encoding="utf-8")
        cls.styles = (ROOT / "styles.css").read_text(encoding="utf-8")

    def test_email_tab_exists_in_nav(self):
        self.assertIn('data-tab="email"', self.html)
        self.assertIn('id="emailTabBadge"', self.html)
        self.assertIn('id="email-panel"', self.html)

    def test_email_panel_contains_inbox_and_follow_ups(self):
        self.assertIn('id="emailInboxFilters"', self.html)
        self.assertIn('id="emailCaptureRescanBtn"', self.html)
        self.assertIn('id="emailCaptureProgress"', self.html)
        self.assertIn('id="emailInboxList"', self.html)
        self.assertIn('id="followUpsSection"', self.html)
        self.assertIn('id="followUpsList"', self.html)
        self.assertIn('id="addFollowUpBtn"', self.html)

    def test_follow_up_dialog_exists(self):
        self.assertIn('<dialog id="followUpDlg">', self.html)
        self.assertIn('id="followUpTitleInput"', self.html)
        self.assertIn('id="followUpWaitingOnInput"', self.html)
        self.assertIn('id="followUpDueInput"', self.html)
        self.assertIn('id="followUpSaveBtn"', self.html)

    def test_project_picker_dialog_exists(self):
        self.assertIn('<dialog id="emailProjectPickerDlg">', self.html)
        self.assertIn('id="emailProjectPickerSearch"', self.html)
        self.assertIn('id="emailProjectPickerList"', self.html)

    def test_script_defines_email_capture_module(self):
        self.assertIn("let emailCaptureState = {", self.script)
        self.assertIn("window.updateEmailCaptureProgress = function", self.script)
        self.assertIn("function renderEmailInboxView() {", self.script)
        self.assertIn("function renderFollowUpsView() {", self.script)
        self.assertIn("function initEmailCaptureUi() {", self.script)
        self.assertIn("async function acceptCapturedEmail(", self.script)
        self.assertIn("async function acceptCapturedEmailIntoProject(", self.script)
        self.assertIn("async function dismissCapturedEmail(", self.script)
        self.assertIn("async function snoozeFollowUp(", self.script)
        self.assertIn("async function resolveFollowUp(", self.script)
        self.assertIn("function startEmailCaptureAutoScan() {", self.script)
        self.assertIn("async function updateEmailTabBadge() {", self.script)

    def test_script_wires_email_tab_and_auto_scan(self):
        self.assertIn('tab === "email"', self.script)
        self.assertIn("initEmailCaptureUi();", self.script)
        self.assertIn("startEmailCaptureAutoScan();", self.script)
        self.assertIn("run_email_capture_scan", self.script)
        self.assertIn("list_captured_emails", self.script)
        self.assertIn("update_captured_email_status", self.script)
        self.assertIn("create_follow_up", self.script)
        self.assertIn("update_follow_up", self.script)
        self.assertIn("get_email_capture_badge_counts", self.script)

    def test_styles_cover_new_ui(self):
        self.assertIn(".tab-badge", self.styles)
        self.assertIn(".email-inbox-card", self.styles)
        self.assertIn(".follow-up-card", self.styles)
        self.assertIn(".email-directedness-pill", self.styles)

    def test_legacy_outlook_scan_ids_remain_absent(self):
        for legacy_id in (
            'id="outlookScanRunBtn"',
            'id="outlookScanProgress"',
            'id="outlookScanReport"',
            'id="outlookScanSuggestions"',
            'id="emailIntakeScanPanel"',
            'id="outlookScanDate"',
        ):
            self.assertNotIn(legacy_id, self.html)
        self.assertNotIn("Scan Outlook Day", self.html)


if __name__ == "__main__":
    unittest.main()
