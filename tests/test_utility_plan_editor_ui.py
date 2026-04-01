import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class UtilityPlanEditorUiTests(unittest.TestCase):
    def test_index_html_includes_utility_plan_tool_dialog_and_detection_banner(self):
        html = (ROOT / "index.html").read_text(encoding="utf-8")

        self.assertIn('id="toolUtilityPlanEditor"', html)
        self.assertIn('id="utilityPlanEditorDlg"', html)
        self.assertIn('id="utilityPlanShell"', html)
        self.assertIn('id="utilityPlanDetectionBanner"', html)
        self.assertIn('id="utilityPlanRefreshDetectionBtn"', html)
        self.assertIn('src="utility-plan-editor.js"', html)
        self.assertNotIn('id="checklistModal"', html)

    def test_script_js_preserves_utility_plan_survey_defaults(self):
        script = (ROOT / "script.js").read_text(encoding="utf-8")

        self.assertIn('value: "utility_plan"', script)
        self.assertIn('value: "photos"', script)
        self.assertIn('value: "findings"', script)
        self.assertIn('value: "recommendations"', script)
        self.assertIn('function normalizeWorkroomSurveyPage(value)', script)
        self.assertNotIn("function openChecklistModal(", script)

    def test_utility_plan_editor_script_is_dialog_only(self):
        utility_plan_script = (ROOT / "utility-plan-editor.js").read_text(encoding="utf-8")

        self.assertIn('utilityPlanState.hostMode = "dialog";', utility_plan_script)
        self.assertIn('browseBtn.textContent = "Browse PDF";', utility_plan_script)
        self.assertNotIn("mountUtilityPlanInWorkroom", utility_plan_script)
        self.assertNotIn("ensureWorkroomSurveyDraftReady", utility_plan_script)
        self.assertNotIn('nextMode === "workroom"', utility_plan_script)

    def test_styles_css_keeps_standalone_utility_plan_shell(self):
        styles = (ROOT / "styles.css").read_text(encoding="utf-8")

        self.assertIn("dialog.utility-plan-dialog {", styles)
        self.assertIn(".utility-plan-shell {", styles)
        self.assertIn(".utility-plan-body {", styles)
        self.assertNotIn(".utility-plan-shell.is-embedded {", styles)


if __name__ == "__main__":
    unittest.main()
