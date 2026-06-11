import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"


class ProjectDeliverableTaskInputUiTests(unittest.TestCase):
    @staticmethod
    def _block(text: str, start_marker: str, end_marker: str) -> str:
        start = text.index(start_marker)
        end = text.index(end_marker, start)
        return text[start:end]

    def test_projects_note_icon_input_commit_and_cancel_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        notes_block = self._block(
            script,
            "function createNotesSection(deliverable, card, project = null) {",
            "function createTasksPreview(deliverable, card, project = null) {",
        )

        for expected in (
            'className: "deliverable-note-add-btn"',
            'addNoteBtn.addEventListener("click", (e) => {',
            "isAddingNote = true;",
            'className: "deliverable-note-add-input"',
            'noteInput.addEventListener("keydown", (e) => {',
            'void finishAddingNote("commit");',
            'void finishAddingNote("cancel");',
            'noteInput.addEventListener("blur", () => {',
            "deliverable.noteItems.push({",
            "notesExpanded = true;",
            "await save();",
        ):
            self.assertIn(expected, notes_block)

        self.assertNotIn("deliverable.tasks.push({", notes_block)
        self.assertNotIn("commitPendingTask", notes_block)


if __name__ == "__main__":
    unittest.main()
