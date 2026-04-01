import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class WorkItemAttachmentUiTests(unittest.TestCase):
    def test_generic_attachment_owner_script_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "function syncAttachmentOwnerEmailRefs(owner) {",
            "function syncAttachmentOwnerLinks(owner, legacyLinkPath = \"\") {",
            "function getAttachmentOwnerKindLabel(kind = \"deliverable\") {",
            "function getAttachmentOwnerLabel(descriptor = {}) {",
            "function buildAttachmentOwnerEmailContext(descriptor = {}, scope = \"projects-tab\") {",
            "function applyAttachmentOwnerEmailRefAtIndex(owner, slotIndex, nextRef, options = {}) {",
            "function removeAttachmentOwnerEmailRefAtIndex(owner, slotIndex, options = {}) {",
            "function renderAttachmentOwnerEmailSlots(container, descriptor = {}, options = {}) {",
            "function createAttachmentLinksControl(descriptor = {}, options = {}) {",
            "function createWorkItemAttachmentControls({",
            "attachmentOwnerKind: kind,",
            "attachmentOwnerLabel: getAttachmentOwnerLabel({",
            "allowProjectScope: false,",
            "scope: \"projects-tab\",",
            "scope: \"edit-modal\",",
            "scope: \"workroom\",",
        ):
            self.assertIn(expected, script)

    def test_item_level_note_and_task_payload_fields_exist(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "emailRefs: normalizeEmailRefs(task?.emailRefs),",
            "links: normalizeDeliverableLinks(noteItem?.links),",
            "emailRefs: normalizeEmailRefs(noteItem?.emailRefs),",
            "task.emailRefs = normalizedTask.emailRefs;",
            "noteItem.links = normalizedNoteItem.links;",
            "noteItem.emailRefs = normalizedNoteItem.emailRefs;",
            "emailRefs: normalizedTask.emailRefs,",
            "links: normalizedNoteItem.links,",
            "emailRefs: normalizedNoteItem.emailRefs,",
        ):
            self.assertIn(expected, script)

    def test_work_item_attachment_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        for expected in (
            ".work-item-content {",
            ".work-item-attachments {",
            ".work-item-link-control .deliverable-link-trigger {",
            ".work-item-email-slots {",
            ".work-item-attachments .deliverable-email-btn {",
            ".work-item-attachments .deliverable-email-remove {",
        ):
            self.assertIn(expected, css)


if __name__ == "__main__":
    unittest.main()
