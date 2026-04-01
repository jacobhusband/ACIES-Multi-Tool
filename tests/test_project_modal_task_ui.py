import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
INDEX_HTML_PATH = REPO_ROOT / "index.html"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectModalTaskUiTests(unittest.TestCase):
    def test_modal_task_editor_script_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn("flushPendingModalDeliverableTasks();", script)
        self.assertIn("function getModalDeliverableTaskItems(card) {", script)
        self.assertIn("function renderModalDeliverableTaskList(card, container, options = {}) {", script)
        self.assertIn("function commitPendingModalDeliverableTask(card, container, options = {}) {", script)
        self.assertIn("function flushPendingModalDeliverableTasks() {", script)
        self.assertIn('const modalBody = element.closest(".modal-body");', script)
        self.assertIn("card._taskItems = (deliverable.tasks || []).map(normalizeTask);", script)
        self.assertIn("renderModalDeliverableTaskList(card, taskList);", script)
        self.assertIn("getModalDeliverableTaskItems(card)", script)
        self.assertIn("taskItems.push({", script)
        self.assertIn("emailRefs: [],", script)
        self.assertIn('taskInput.addEventListener("blur", () => {', script)
        self.assertIn("setTimeout(() => {", script)
        self.assertIn("focusModalTaskInput(container);", script)
        self.assertIn("links: normalizedTask.links,", script)
        self.assertIn("emailRefs: normalizedTask.emailRefs,", script)
        self.assertIn('titleUnpinned: "Pin task",', script)
        self.assertIn('scope: "edit-modal",', script)
        self.assertIn('kind: "task",', script)
        self.assertIn("createWorkItemAttachmentControls({", script)

    def test_modal_task_editor_markup_removes_old_button_and_template(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")

        self.assertIn('<div class="title section-header">Tasks</div>', html)
        self.assertIn('<div class="deliverable-task-list stack"></div>', html)
        self.assertNotIn('class="btn-accent tiny d-add-task"', html)
        self.assertNotIn('<template id="task-row-template">', html)

    def test_modal_task_editor_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        self.assertIn(".deliverable-task-list {", css)
        self.assertIn(".deliverable-task-list .deliverable-task-item {", css)
        self.assertIn(".deliverable-task-list .task-add-row {", css)
        self.assertIn(".deliverable-task-list .task-add-input {", css)


if __name__ == "__main__":
    unittest.main()
