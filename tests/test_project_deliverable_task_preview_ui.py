import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectDeliverableTaskPreviewUiTests(unittest.TestCase):
    @staticmethod
    def _block(text: str, start_marker: str, end_marker: str) -> str:
        start = text.index(start_marker)
        end = text.index(end_marker, start)
        return text[start:end]

    def test_projects_notes_footer_limits_rows_and_toggles_overflow(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        notes_block = self._block(
            script,
            "function createNotesSection(deliverable, card, project = null) {",
            "function createTasksPreview(deliverable, card, project = null) {",
        )

        self.assertIn("const visibleEntries = notesExpanded ? noteEntries : noteEntries.slice(0, 2);", notes_block)
        self.assertIn("const hasOverflow = noteEntries.length > 2;", notes_block)
        self.assertIn('toggleBtn.textContent = notesExpanded ? "Hide notes" : "Show all notes";', notes_block)
        self.assertIn('toggleBtn.setAttribute("aria-expanded", String(notesExpanded));', notes_block)
        self.assertNotIn("getPinnedDeliverableTaskEntries(deliverable).forEach", notes_block)

    def test_projects_notes_footer_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        for expected in (
            ".deliverable-notes-footer {",
            ".deliverable-notes-footer-controls {",
            ".deliverable-note-add-btn {",
            ".deliverable-note-row {",
            ".deliverable-note-row .note-text {",
            ".deliverable-note-delete-btn {",
            ".deliverable-note-add-input {",
            ".deliverable-notes-toggle {",
        ):
            self.assertIn(expected, css)


if __name__ == "__main__":
    unittest.main()
