import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectDeliverableInlineEditUiTests(unittest.TestCase):
    @staticmethod
    def _block(text: str, start_marker: str, end_marker: str) -> str:
        start = text.rindex(start_marker)
        end = text.index(end_marker, start)
        return text[start:end]

    @staticmethod
    def _css_block(css: str, selector: str) -> str:
        block_start = css.index(selector)
        block_end = css.index("}", block_start)
        return css[block_start:block_end]

    def test_inline_edit_helper_script_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        helper_block = self._block(
            script,
            "function createWorkItemDoneCheckbox({",
            "function createNotesSection(deliverable, card, project = null) {",
        )

        for expected in (
            "function createWorkItemDoneCheckbox({",
            'className: "work-item-checkbox",',
            "function createInlineWorkItemTextControl({",
            'className: "work-item-text-wrap"',
            'className: `work-item-text-trigger ${textClassName}`.trim(),',
            'className: "work-item-text-input",',
            'title = "Click to edit",',
            "const nextText = trimmedText || previousText;",
            "await onCommit(nextText, previousText);",
            'void finishEditing("commit");',
            'void finishEditing("cancel");',
            '} else if (event.key === "Escape") {',
        ):
            self.assertIn(expected, helper_block)

    def test_projects_tab_work_item_inline_edit_script_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        notes_block = self._block(
            script,
            "function createNotesSection(deliverable, card, project = null) {",
            "function createTasksPreview(deliverable, card, project = null) {",
        )
        tasks_block = self._block(
            script,
            "function createTasksPreview(deliverable, card, project = null) {",
            "function renderDeliverableCard(deliverable, isPrimary, project) {",
        )

        for expected in (
            'const textControl = createInlineWorkItemTextControl({',
            'title: "Click to edit note",',
            "liveNote.text = nextText;",
            "syncDeliverableNoteFields(deliverable);",
            "await save();",
        ):
            self.assertIn(expected, notes_block)

        for expected in (
            "const checkbox = createWorkItemDoneCheckbox({",
            'const textControl = createInlineWorkItemTextControl({',
            'title: "Click to edit task",',
            "liveTask.done = nextDone;",
            "liveTask.text = nextText;",
            "item.append(checkbox, content, actions);",
            "await save();",
        ):
            self.assertIn(expected, tasks_block)

        self.assertNotIn('item.addEventListener("click"', tasks_block)

    def test_modal_work_item_inline_edit_script_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        modal_task_block = self._block(
            script,
            "function renderModalDeliverableTaskList(card, container, options = {}) {",
            "function commitPendingModalDeliverableTask(card, container, options = {}) {",
        )
        modal_note_block = self._block(
            script,
            "function renderModalDeliverableNoteList(card, container, options = {}) {",
            "function commitPendingModalDeliverableNote(card, container, options = {}) {",
        )

        for expected in (
            "const checkbox = createWorkItemDoneCheckbox({",
            'const textControl = createInlineWorkItemTextControl({',
            'title: "Click to edit task",',
            "liveTask.done = nextDone;",
            "liveTask.text = nextText;",
            "item.append(checkbox, content, actions);",
        ):
            self.assertIn(expected, modal_task_block)

        for expected in (
            'const textControl = createInlineWorkItemTextControl({',
            'title: "Click to edit note",',
            "liveNote.text = nextText;",
        ):
            self.assertIn(expected, modal_note_block)

        self.assertNotIn('item.addEventListener("click"', modal_task_block)
        self.assertNotIn("await save();", modal_task_block)
        self.assertNotIn("await save();", modal_note_block)

    def test_inline_edit_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        for expected in (
            ".work-item-checkbox {",
            ".work-item-checkbox:focus-visible {",
            ".work-item-text-wrap {",
            ".work-item-text-trigger {",
            ".work-item-text-trigger:hover {",
            ".work-item-text-trigger:focus-visible {",
            ".work-item-text-input {",
        ):
            self.assertIn(expected, css)

        item_block = self._css_block(
            css,
            ".deliverable-task-item,\n.deliverable-note-item {",
        )
        self.assertNotIn("cursor: pointer;", item_block)
        self.assertNotIn("user-select: none;", item_block)

        text_trigger_block = self._css_block(css, ".work-item-text-trigger {")
        self.assertIn("cursor: text;", text_trigger_block)
        self.assertIn("background: transparent;", text_trigger_block)

        checkbox_block = self._css_block(css, ".work-item-checkbox {")
        self.assertIn("accent-color: var(--accent);", checkbox_block)
        self.assertIn("cursor: pointer;", checkbox_block)


if __name__ == "__main__":
    unittest.main()
