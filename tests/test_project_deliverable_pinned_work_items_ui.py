import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectDeliverablePinnedWorkItemsUiTests(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()
