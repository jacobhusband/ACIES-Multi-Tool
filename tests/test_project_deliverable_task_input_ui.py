import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"


class ProjectDeliverableTaskInputUiTests(unittest.TestCase):
    def test_projects_task_input_blur_submission_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn(
            "const commitPendingTask = async ({ refocus = false } = {}) => {",
            script,
        )
        self.assertIn("if (!taskText || isCommittingTask) return false;", script)
        self.assertIn('taskInput.value = "";', script)
        self.assertIn(
            "deliverable.tasks.push({",
            script,
        )
        self.assertIn("pinned: false,", script)
        self.assertIn("attachments: [],", script)
        self.assertIn("updateStatsDisplay();", script)
        self.assertIn("await save();", script)
        self.assertIn("renderTaskList();", script)
        self.assertIn("await commitPendingTask({ refocus: true });", script)
        self.assertIn('taskInput.addEventListener("blur", () => {', script)
        self.assertIn("if (!taskInput.value.trim()) return;", script)
        self.assertIn("setTimeout(() => {", script)
        self.assertIn("void commitPendingTask({ refocus: false });", script)
        self.assertIn("createWorkItemAttachmentControls({", script)
        self.assertIn("actions.append(attachments, pinBtn, deleteBtn);", script)


if __name__ == "__main__":
    unittest.main()
