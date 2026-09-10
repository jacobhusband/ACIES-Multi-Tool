import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectCardNoneColumnUiTests(unittest.TestCase):
    def test_default_card_columns_place_in_progress_between_pinned_and_waiting(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        start = script.index("const DEFAULT_PROJECT_CARD_COLUMNS = [")
        end = script.index("];", start)
        block = script[start:end]

        self.assertLess(
            block.index('{ key: "pinned", label: "Pinned", hidden: false }'),
            block.index('{ key: "In progress", label: "In progress", hidden: false }'),
        )
        self.assertLess(
            block.index('{ key: "In progress", label: "In progress", hidden: false }'),
            block.index('{ key: "Waiting", label: "Waiting", hidden: false }'),
        )

    def test_saved_card_column_settings_insert_missing_defaults_by_default_order(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        start = script.index("function normalizeProjectCardColumns(raw) {")
        end = script.index("function syncProjectViewPreferencesFromSettings()", start)
        block = script[start:end]

        for expected in (
            "const findDefaultIndex = (key) =>",
            "const insertMissingColumn = (column) => {",
            "(existing) => findDefaultIndex(existing.key) > defaultIndex",
            "normalized.splice(nextIndex, 0, { ...column });",
            "if (!seen.has(col.key)) insertMissingColumn(col);",
        ):
            self.assertIn(expected, block)

    def test_card_view_routes_blank_deliverables_to_in_progress_column(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        start = script.index("function renderCardView(")
        end = script.index("function setProjectsViewMode(", start)
        block = script[start:end]

        self.assertIn('"In progress": "in-progress",', script)
        self.assertIn('STATUS_PRIORITY.find((s) => hasStatus(deliverable, s)) || "In progress"', block)
        self.assertNotIn('|| "Waiting"', block)

    def test_drop_sets_a_required_status(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        start = script.index('host.addEventListener("drop", async (e) => {')
        end = script.index("function render() {", start)
        block = script[start:end]
        self.assertIn("setSingleStatus(deliverable, targetKey);", block)
        self.assertNotIn('setSingleStatus(deliverable, "");', block)

    def test_in_progress_column_style_exists(self):
        css = (REPO_ROOT / "deliverables-workspace.css").read_text(encoding="utf-8")
        self.assertIn(".kanban-column--in-progress .kanban-column__label", css)


if __name__ == "__main__":
    unittest.main()
