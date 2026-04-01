import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectDeliverablePinnedWorkItemsUiTests(unittest.TestCase):
    def test_projects_pinned_work_items_script_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "function getPinnedFirstEntries(items = [], normalizeItem = (item) => item) {",
            "function getPinnedDeliverableTaskEntries(deliverable) {",
            "function getPinnedDeliverableNoteEntries(deliverable) {",
            "function getPinnedDeliverablePreviewItems(deliverable) {",
            "function createWorkItemPinButton({",
            "function renderDeliverablePinnedPreview(container, deliverable) {",
            "function updateDeliverableWorkItemUi(card, deliverable) {",
            "function createPinnedWorkItemAttachments(previewItem = {}) {",
            'className: "deliverable-pinned-preview"',
            'titleUnpinned: "Pin task",',
            'titleUnpinned: "Pin note",',
            "pinned: !!task?.pinned,",
            "pinned: !!noteItem?.pinned,",
            "links: normalizeTaskLinks(item.links),",
            "emailRefs: normalizeEmailRefs(item.emailRefs),",
            'className: "deliverable-pinned-item-attachments"',
            'className: "deliverable-pinned-link"',
            'className: "deliverable-pinned-email-btn"',
        ):
            self.assertIn(expected, script)

    def test_projects_pinned_work_items_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        for expected in (
            ".work-item-actions {",
            ".work-item-pin-btn {",
            ".work-item-pin-btn.is-pinned {",
            ".deliverable-pinned-preview {",
            ".deliverable-pinned-preview-heading {",
            ".deliverable-pinned-preview-list {",
            ".deliverable-pinned-item {",
            ".deliverable-pinned-item-body {",
            ".deliverable-pinned-item-kind {",
            ".deliverable-pinned-item-text {",
            ".deliverable-pinned-item-attachments {",
            ".deliverable-pinned-link,",
            ".deliverable-pinned-email-btn {",
        ):
            self.assertIn(expected, css)


if __name__ == "__main__":
    unittest.main()
