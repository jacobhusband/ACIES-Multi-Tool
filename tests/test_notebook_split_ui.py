import re
import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
INDEX_HTML_PATH = REPO_ROOT / "index.html"
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


def get_css_block(css: str, selector: str) -> str:
    pattern = re.compile(rf"{re.escape(selector)}\s*\{{(.*?)\}}", re.S)
    match = pattern.search(css)
    if not match:
        raise AssertionError(f"CSS selector {selector!r} not found")
    return match.group(1)


class NotebookSplitUiTests(unittest.TestCase):
    def test_notebook_markup_uses_shared_content_shell(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")

        self.assertIn('class="notebook-layout"', html)
        self.assertIn('class="notebook-sidebar"', html)
        self.assertIn('id="notebookContentShell"', html)
        self.assertIn('id="notesTabsContainer"', html)
        self.assertIn('id="checklistsTabsContainer"', html)
        self.assertIn('id="textsNotesPane"', html)
        self.assertIn('id="textsChecklistsPane"', html)
        self.assertIn('id="checklistSearchResults"', html)
        self.assertNotIn('id="checklistsHelpBtn"', html)

        shell_index = html.index('id="notebookContentShell"')
        notes_index = html.index('id="textsNotesPane"')
        checklists_index = html.index('id="textsChecklistsPane"')
        self.assertLess(shell_index, notes_index)
        self.assertLess(notes_index, checklists_index)

        checklist_section_match = re.search(
            r'<section[^>]*id="textsChecklistsPane"[^>]*>',
            html,
        )
        self.assertIsNotNone(checklist_section_match)
        self.assertIn("hidden", checklist_section_match.group(0))

    def test_notebook_script_uses_shared_active_type_and_mode_aware_search(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn('let activeNotebookType = "note";', script)
        self.assertIn('let notesSearchQuery = "";', script)
        self.assertIn('let checklistSearchQuery = "";', script)
        self.assertIn("function ensureActiveNotebookSelection(", script)
        self.assertIn('const isActive = activeNotebookType === "note" && tabName === activeNoteTab;', script)
        self.assertIn(
            'activeNotebookType === "checklist" && checklist.id === activeChecklistTabId',
            script,
        )
        self.assertIn('notesPane.hidden = isChecklistMode;', script)
        self.assertIn('checklistsPane.hidden = !isChecklistMode;', script)
        self.assertIn('openExternalUrl(isChecklistMode ? HELP_LINKS.checklists : HELP_LINKS.notes);', script)
        self.assertIn('searchInput.placeholder = isChecklistMode', script)
        self.assertIn("function renderChecklistSearchResults() {", script)
        self.assertIn('focusChecklistRowInput(item.id, { scrollIntoView: true });', script)
        self.assertNotIn("checklistsHelpBtn", script)

    def test_notebook_styles_keep_vertical_tabs_and_shared_width_model(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        full_height_block = get_css_block(css, "#texts-panel .full-height")
        self.assertIn(
            "height: clamp(700px, calc(100vh - var(--header-height) - 3.5rem), 920px);",
            full_height_block,
        )
        self.assertIn(
            "max-height: clamp(700px, calc(100vh - var(--header-height) - 3.5rem), 920px);",
            full_height_block,
        )
        self.assertIn("overflow: clip;", full_height_block)

        shared_nav_block = get_css_block(css, ".notes-nav-list,\n.checklists-nav-list")
        self.assertIn("width: 100%;", shared_nav_block)
        self.assertIn("flex-direction: column;", shared_nav_block)
        self.assertIn("flex-wrap: nowrap;", shared_nav_block)
        self.assertIn("overflow-y: auto;", shared_nav_block)

        shared_tab_block = get_css_block(
            css,
            ".notes-nav-list .inner-tab-btn,\n.checklists-nav-list .inner-tab-btn",
        )
        self.assertIn("width: 100%;", shared_tab_block)

        shared_create_block = get_css_block(
            css,
            ".notes-nav-list .add-tab-btn,\n.checklist-create-trigger",
        )
        self.assertIn("width: 100%;", shared_create_block)

        self.assertIn(".notebook-layout {", css)
        self.assertIn(".notebook-content-shell {", css)
        self.assertIn(".checklist-search-results {", css)
        self.assertIn(".checklist-search-result-item {", css)
        self.assertIn("height: clamp(620px, calc(100vh - var(--header-height) - 2.5rem), 780px);", css)
        self.assertIn("max-height: clamp(620px, calc(100vh - var(--header-height) - 2.5rem), 780px);", css)

        checklist_editor_block = get_css_block(css, ".checklist-editor")
        self.assertIn("min-height: 0;", checklist_editor_block)
        self.assertIn("overflow: hidden;", checklist_editor_block)

        checklist_items_block = get_css_block(css, ".checklist-items")
        self.assertIn("min-height: 0;", checklist_items_block)
        self.assertIn("overflow-y: auto;", checklist_items_block)

        self.assertNotRegex(
            css,
            re.compile(
                r"\.checklists-nav-list\s*\{\s*flex-direction:\s*row;\s*flex-wrap:\s*wrap;",
                re.S,
            ),
        )
        self.assertNotRegex(
            css,
            re.compile(
                r"\.notes-nav-list,\s*\.checklists-nav-list\s*\{\s*flex-direction:\s*row;\s*flex-wrap:\s*wrap;",
                re.S,
            ),
        )
        self.assertNotIn("max-height: 60vh;", css)


if __name__ == "__main__":
    unittest.main()
