import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
INDUSTRY_PROJECTS_JS = REPO_ROOT / "industry-projects.js"


class CommandLineCloseActionsUiTests(unittest.TestCase):
    def setUp(self):
        self.js_code = INDUSTRY_PROJECTS_JS.read_text(encoding="utf-8")

    def test_run_command_dock_item_closes_expanded_actions(self):
        # Verify runCommandDockItem collapses the command dock and blurs input
        self.assertIn("function runCommandDockItem(item) {", self.js_code)
        self.assertIn("if (!industryActivePrompt) {", self.js_code)
        self.assertIn("setCommandDockExpanded(false);", self.js_code)
        self.assertIn("dock.input.blur();", self.js_code)

    def test_complete_add_deliverable_prompt_closes_actions(self):
        # Verify completeAddDeliverablePrompt collapses the command dock
        self.assertIn("function completeAddDeliverablePrompt() {", self.js_code)
        self.assertIn("selectDeliverableForCommands(deliverable, project, { expand: false });", self.js_code)
        self.assertIn("setCommandDockExpanded(false);", self.js_code)

    def test_status_commands_routed_through_run_command_dock_item(self):
        # Verify status items use the standard runCommandDockItem runner
        self.assertIn('node.dataset.commandKey = item.key;', self.js_code)
        self.assertIn('runCommandDockItem(item);', self.js_code)
        self.assertIn('key: `status:${status}`', self.js_code)


if __name__ == "__main__":
    unittest.main()
