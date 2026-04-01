import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
INDEX_HTML_PATH = REPO_ROOT / "index.html"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectModalNoteUiTests(unittest.TestCase):
    def test_modal_note_editor_script_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn("flushPendingModalDeliverableNotes();", script)
        self.assertIn("const hasExplicitNoteItems = Array.isArray(noteItems);", script)
        self.assertIn("function getModalDeliverableNoteItems(card) {", script)
        self.assertIn(
            "function renderModalDeliverableNoteList(card, container, options = {}) {",
            script,
        )
        self.assertIn(
            "function commitPendingModalDeliverableNote(card, container, options = {}) {",
            script,
        )
        self.assertIn("function flushPendingModalDeliverableNotes() {", script)
        self.assertIn("card._noteItems = normalizeDeliverableNoteItems(", script)
        self.assertIn("renderModalDeliverableNoteList(card, noteList);", script)
        self.assertIn("noteItems.push({ text: noteText, pinned: false });", script)
        self.assertIn("const noteItems = getModalDeliverableNoteItems(card)", script)
        self.assertIn('if (!hasExplicitNoteItems && String(legacyNotes || "").trim()) {', script)

    def test_modal_note_editor_markup_replaces_textarea(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")

        self.assertIn('<div class="title section-header">Notes</div>', html)
        self.assertIn('<div class="deliverable-note-list stack"></div>', html)
        self.assertNotIn('class="d-notes"', html)

    def test_modal_note_editor_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        self.assertIn(".deliverable-note-list {", css)
        self.assertIn(".deliverable-note-item {", css)
        self.assertIn(".deliverable-notes-list-edit {", css)
        self.assertIn(".note-add-bullet {", css)


if __name__ == "__main__":
    unittest.main()
