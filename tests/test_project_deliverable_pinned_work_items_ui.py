import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectDeliverablePinnedWorkItemsUiTests(unittest.TestCase):
    @staticmethod
    def _css_block(css: str, selector: str) -> str:
        block_start = css.index(selector)
        block_end = css.index("}", block_start)
        return css[block_start:block_end]

    @staticmethod
    def _script_block(script: str, start_marker: str, end_marker: str) -> str:
        block_start = script.index(start_marker)
        block_end = script.index(end_marker, block_start)
        return script[block_start:block_end]

    @staticmethod
    def _script_block_within(
        script: str,
        container_start: str,
        start_marker: str,
        end_marker: str,
    ) -> str:
        container_index = script.index(container_start)
        block_start = script.index(start_marker, container_index)
        block_end = script.index(end_marker, block_start)
        return script[block_start:block_end]

    def test_projects_pinned_work_items_script_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertEqual(script.count("function createNotesSection("), 1)
        self.assertEqual(script.count("function createTasksPreview("), 1)
        self.assertEqual(script.count("function renderDeliverableCard("), 1)

        for expected in (
            "function getPinnedFirstEntries(items = [], normalizeItem = (item) => item) {",
            "function getPinnedDeliverableTaskEntries(deliverable) {",
            "function getPinnedDeliverableNoteEntries(deliverable) {",
            "function getPinnedDeliverablePreviewItems(deliverable) {",
            "function createWorkItemPinButton({",
            "function renderDeliverableStatusBadges(container, deliverable) {",
            "function renderDeliverablePinnedPreview(container, deliverable) {",
            "function updateDeliverableWorkItemUi(card, deliverable) {",
            "function createPinnedAttachmentPreviewButton(attachment) {",
            "function createPinnedWorkItemAttachments(previewItem = {}) {",
            "function createNotesSection(deliverable, card, project = null) {",
            "function createTasksPreview(deliverable, card, project = null) {",
            "function renderDeliverableCard(deliverable, isPrimary, project) {",
            'className: "deliverable-pinned-inline-group"',
            'titleUnpinned: "Pin task",',
            'titleUnpinned: "Pin note",',
            "liveNote.pinned = nextPinned;",
            "renderNoteList();",
            "pinned: !!task?.pinned,",
            "pinned: !!noteItem?.pinned,",
            "attachments: normalizeAttachments(item.attachments, {",
            'pillIcon.classList.add("deliverable-pinned-inline-pill-icon");',
            'className: `deliverable-pinned-inline-pill ${',
            'className: "deliverable-pinned-inline-pill-text"',
            'className: "deliverable-pinned-item-attachments"',
            'className: "deliverable-pinned-link"',
            "renderDeliverableStatusBadges(badgesContainer, deliverable);",
            "const notesSection = createNotesSection(deliverable, card, project);",
        ):
            self.assertIn(expected, script)

        self.assertNotIn("function createNotesSection(deliverable, project) {", script)
        self.assertNotIn("function createTasksPreview(deliverable, card) {", script)
        self.assertNotIn("const notesSection = createNotesSection(deliverable, project);", script)
        self.assertNotIn('className: "deliverable-pinned-preview-heading"', script)
        self.assertNotIn('className: "deliverable-pinned-preview"', script)
        self.assertNotIn('className: "deliverable-pinned-item-kind"', script)
        self.assertNotIn('className: "deliverable-pinned-item-kind-label"', script)
        self.assertNotIn("createPinnedWorkItemAttachments(previewItem);", script)

    def test_projects_pinned_work_items_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        for expected in (
            ".work-item-actions {",
            ".work-item-pin-btn {",
            ".work-item-pin-btn.is-pinned {",
            ".deliverable-pinned-inline-group {",
            ".deliverable-pinned-inline-pill {",
            ".deliverable-pinned-inline-pill-icon {",
            ".deliverable-pinned-inline-pill-text {",
            ".deliverable-pinned-inline-pill.is-task {",
            ".deliverable-pinned-inline-pill.is-note {",
            ".deliverable-pinned-item-attachments {",
            ".deliverable-pinned-link,",
        ):
            self.assertIn(expected, css)

        self.assertNotIn(".deliverable-pinned-preview-heading {", css)
        self.assertNotIn(".deliverable-pinned-preview {", css)
        self.assertNotIn(".deliverable-pinned-item-kind {", css)

        pill_block = self._css_block(css, ".deliverable-pinned-inline-pill {")
        self.assertIn("align-items: flex-start;", pill_block)
        self.assertIn("max-width: min(100%, 240px);", pill_block)

        pill_text_block = self._css_block(css, ".deliverable-pinned-inline-pill-text {")
        self.assertIn("white-space: normal;", pill_text_block)
        self.assertIn("overflow-wrap: anywhere;", pill_text_block)
        self.assertNotIn("text-overflow: ellipsis;", pill_text_block)
        self.assertNotIn("white-space: nowrap;", pill_text_block)

    def test_projects_note_toggle_uses_work_item_sync_and_save(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        projects_note_toggle_block = self._script_block_within(
            script,
            "function createNotesSection(deliverable, card, project = null) {",
            'const pinBtn = createWorkItemPinButton({',
            'const deleteBtn = el("button", {',
        )

        for expected in (
            'titlePinned: "Unpin note",',
            'titleUnpinned: "Pin note",',
            "liveNote.pinned = nextPinned;",
            "deliverable.noteItems[index].pinned = nextPinned;",
            "syncDeliverableWorkItemFields(deliverable);",
            "renderNoteList();",
            "updateDeliverableWorkItemUi(card, deliverable);",
            "await save();",
            "return nextPinned;",
        ):
            self.assertIn(expected, projects_note_toggle_block)

        self.assertNotIn(
            "syncDeliverableNoteFields(deliverable);",
            projects_note_toggle_block,
        )

    def test_main_and_modal_note_toggles_share_next_pinned_contract(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        projects_note_toggle_block = self._script_block_within(
            script,
            "function createNotesSection(deliverable, card, project = null) {",
            'const pinBtn = createWorkItemPinButton({',
            'const deleteBtn = el("button", {',
        )
        modal_note_toggle_block = self._script_block_within(
            script,
            "function renderModalDeliverableNoteList(card, container, options = {}) {",
            'const pinBtn = createWorkItemPinButton({',
            'const deleteBtn = el("button", {',
        )

        for expected in (
            'const pinBtn = createWorkItemPinButton({',
            'titlePinned: "Unpin note",',
            'titleUnpinned: "Pin note",',
            "liveNote.pinned = nextPinned;",
            "return nextPinned;",
        ):
            self.assertIn(expected, projects_note_toggle_block)
            self.assertIn(expected, modal_note_toggle_block)

    def test_work_item_pin_button_supports_async_toggle_rollback(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        pin_button_block = self._script_block(
            script,
            "function createWorkItemPinButton({",
            "function getModalDeliverableCardContext(card) {",
        )

        for expected in (
            "let isTogglePending = false;",
            'button.setAttribute("aria-busy", String(isTogglePending));',
            "if (isTogglePending) return;",
            "const previousPinned = currentPinned;",
            "const nextPinned = !currentPinned;",
            "const resolvedPinned = onToggle ? await onToggle(nextPinned) : nextPinned;",
            'if (typeof resolvedPinned === "boolean") {',
            "currentPinned = previousPinned;",
            'console.warn("Failed to toggle work item pin state:", error);',
        ):
            self.assertIn(expected, pin_button_block)


if __name__ == "__main__":
    unittest.main()
