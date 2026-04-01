import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectDeliverableTaskPreviewUiTests(unittest.TestCase):
    def test_projects_task_preview_always_renders_full_list(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn(
            "function createTasksPreview(deliverable, card, project = null) {",
            script,
        )
        self.assertIn(
            "getPinnedDeliverableTaskEntries(deliverable).forEach(({ item: taskObj, index }) => {",
            script,
        )
        self.assertIn("renderTaskList();", script)
        self.assertNotIn("const maxVisible = 3;", script)
        self.assertNotIn('className: "deliverable-expand-btn"', script)
        self.assertNotIn('textContent: "Show less"', script)
        self.assertNotIn(
            "textContent: `+${deliverable.tasks.length - maxVisible} more`",
            script,
        )

    def test_projects_task_preview_expand_button_styles_removed(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        self.assertNotIn(".deliverable-expand-btn {", css)
        self.assertNotIn(".deliverable-expand-btn:hover {", css)


if __name__ == "__main__":
    unittest.main()
