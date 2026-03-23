import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectDeliverableTaskCountUiTests(unittest.TestCase):
    def test_projects_deliverable_task_count_script_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn("function createDeliverableTaskCountBadge(deliverable) {", script)
        self.assertIn('className: "deliverable-task-count-badge"', script)
        self.assertIn('className: "deliverable-task-count-value"', script)
        self.assertIn("function updateDeliverableTaskStats(card, deliverable) {", script)
        self.assertIn(
            'const taskCountBadge = card.querySelector(".deliverable-task-count-badge");',
            script,
        )
        self.assertIn(
            'const taskCountBadge = createDeliverableTaskCountBadge(deliverable);',
            script,
        )
        self.assertIn("leftSection.appendChild(taskCountBadge);", script)
        self.assertIn(
            'taskCountBadge.setAttribute("aria-label", taskSummaryLabel);',
            script,
        )
        self.assertIn("updateDeliverableTaskStats(card, deliverable);", script)
        self.assertGreaterEqual(script.count("updateStatsDisplay();"), 3)

    def test_projects_deliverable_task_count_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        self.assertIn(".deliverable-task-count-badge {", css)
        self.assertIn(".deliverable-task-count-icon {", css)
        self.assertIn(".deliverable-task-count-value {", css)
        self.assertIn(".deliverable-task-count-badge.is-empty {", css)
        self.assertIn(".deliverable-task-count-badge.is-complete {", css)
        self.assertIn(
            ".deliverable-card-new.details-collapsed .deliverable-task-count-badge {",
            css,
        )


if __name__ == "__main__":
    unittest.main()
