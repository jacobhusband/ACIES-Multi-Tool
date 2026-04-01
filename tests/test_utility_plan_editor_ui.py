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
        self.assertIn('id="workroomSurveyTemplateBtn"', html)
        self.assertIn('src="utility-plan-editor.js"', html)

    def test_script_js_adds_paginated_survey_creator_flow(self):
        script = (ROOT / "script.js").read_text(encoding="utf-8")

        self.assertIn('value: "utility_plan"', script)
        self.assertIn('value: "photos"', script)
        self.assertIn('value: "findings"', script)
        self.assertIn('value: "recommendations"', script)
        self.assertIn('function normalizeWorkroomSurveyPage(value)', script)
        self.assertIn('modal.classList.toggle("is-survey-phase"', script)
        self.assertIn('launchBtn.hidden = true;', script)

    def test_styles_css_hides_workroom_rail_in_survey_phase(self):
        styles = (ROOT / "styles.css").read_text(encoding="utf-8")

        self.assertIn(".workroom-modal.is-survey-phase .workroom-body", styles)
        self.assertIn(".workroom-modal.is-survey-phase .workroom-tools-column", styles)
        self.assertIn(".workroom-survey-step-row", styles)
        self.assertIn(".workroom-survey-page-shell", styles)


if __name__ == "__main__":
    unittest.main()
