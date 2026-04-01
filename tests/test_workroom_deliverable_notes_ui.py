import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
INDEX_HTML_PATH = REPO_ROOT / "index.html"
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class WorkroomDeliverableNotesUiTests(unittest.TestCase):
    def test_workroom_notes_markup_uses_list_editor(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")

        self.assertIn('id="workroomAddNoteBtn"', html)
        self.assertIn('id="workroomNotesContainer"', html)
        self.assertNotIn('id="workroomDeliverableNotesTextarea"', html)

    def test_workroom_notes_script_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            'const addNoteBtn = document.getElementById("workroomAddNoteBtn");',
            'const container = document.getElementById("workroomNotesContainer");',
            'deliverable.noteItems.push({ text: "New note", pinned: false });',
            'className: "workroom-note-text-input",',
            'titleUnpinned: "Pin note",',
            'titleUnpinned: "Pin task",',
            'className: "checklist-task-actions"',
        ):
            self.assertIn(expected, script)

    def test_workroom_notes_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        for expected in (
            ".workroom-notes-container {",
            ".workroom-note-row {",
            ".workroom-note-text-input {",
            ".checklist-task-actions {",
            ".checklist-work-item-pin {",
        ):
            self.assertIn(expected, css)


if __name__ == "__main__":
    unittest.main()
