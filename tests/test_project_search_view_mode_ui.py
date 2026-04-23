import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"


class ProjectSearchViewModeUiTests(unittest.TestCase):
    def test_project_search_from_card_view_switches_to_list_without_saving_preference(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        handler_start = script.index("function handleProjectSearchInput() {")
        handler_end = script.index("function initEventListeners() {", handler_start)
        handler_block = script[handler_start:handler_end]
        setter_start = script.index("function setProjectsViewMode(mode, options = {}) {")
        setter_end = script.index("function shiftProjectCardWeek(days) {", setter_start)
        setter_block = script[setter_start:setter_end]
        listeners_start = script.index("function initEventListeners() {")
        listeners_end = script.index(
            'document.getElementById("notesSearch").addEventListener',
            listeners_start,
        )
        listeners_block = script[listeners_start:listeners_end]

        self.assertIn('projectsViewMode === "card" && val("search")', handler_block)
        self.assertIn('setProjectsViewMode("list", { persist: false });', handler_block)
        self.assertIn("render();", handler_block)
        self.assertIn("debounce(handleProjectSearchInput, 250)", listeners_block)
        self.assertIn("const shouldPersist = options.persist !== false;", setter_block)
        self.assertIn(
            "if (shouldPersist && userSettings.projectsViewMode !== next) {",
            setter_block,
        )


if __name__ == "__main__":
    unittest.main()
