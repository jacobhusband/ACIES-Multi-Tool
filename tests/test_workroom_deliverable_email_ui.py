import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class WorkroomDeliverableEmailUiTests(unittest.TestCase):
    def test_workroom_deliverable_email_script_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn("function renderWorkroomCadRoutingControl() {", script)
        self.assertIn('const { project, deliverable } = getActiveWorkroomContext();', script)
        self.assertIn('textContent: "Attached Emails"', script)
        self.assertIn('className: "workroom-deliverable-email-slots"', script)
        self.assertIn('renderDeliverableEmailSlots(emailSlots, deliverable, {', script)
        self.assertIn("persistNow: true,", script)
        self.assertIn('scope: "workroom",', script)

    def test_workroom_deliverable_email_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        self.assertIn(".workroom-cad-routing-row {", css)
        self.assertIn(".workroom-cad-routing-section {", css)
        self.assertIn(".workroom-cad-routing-email {", css)
        self.assertIn(".workroom-deliverable-email-slots {", css)


if __name__ == "__main__":
    unittest.main()
