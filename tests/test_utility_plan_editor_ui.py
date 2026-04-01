import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class UtilityPlanEditorUiTests(unittest.TestCase):
    def test_index_html_includes_utility_plan_tool_and_dialog(self):
        html = (ROOT / "index.html").read_text(encoding="utf-8")

        self.assertIn('id="toolUtilityPlanEditor"', html)
        self.assertIn('id="utilityPlanEditorDlg"', html)
        self.assertIn('id="workroomSurveyTemplateBtn"', html)
        self.assertIn('src="utility-plan-editor.js"', html)

    def test_script_js_maps_survey_phase_to_utility_plan_editor(self):
        script = (ROOT / "script.js").read_text(encoding="utf-8")

        self.assertIn('"toolUtilityPlanEditor"', script)
        self.assertIn('survey_report: ["toolUtilityPlanEditor", "toolCreateElectricalSurveyTemplate"]', script)
        self.assertIn('await triggerWorkroomTool("toolUtilityPlanEditor")', script)


if __name__ == "__main__":
    unittest.main()
