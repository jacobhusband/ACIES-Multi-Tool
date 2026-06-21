import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
INDEX_HTML_PATH = REPO_ROOT / "index.html"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class PageItemsRollupUiTests(unittest.TestCase):
    """The cross-project roll-up of tagged page items replaces the old
    coordination-items feature."""

    @staticmethod
    def _block(text: str, start_marker: str, end_marker: str) -> str:
        start = text.index(start_marker)
        end = text.index(end_marker, start)
        return text[start:end]

    def test_rollup_script_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        self.assertIn("function extractPageTaggedItems(html) {", script)
        self.assertIn("function collectRollupSources() {", script)
        self.assertIn("function renderRollupSourcePanel(source) {", script)
        self.assertIn("function renderPageItemsRollupView() {", script)
        self.assertIn("function setPageItemDoneInHtml(page, itemId, done) {", script)
        self.assertIn("function openPageToItem(source, itemId) {", script)
        self.assertIn("function refreshRollupIfActive() {", script)

        extract_block = self._block(
            script,
            "function extractPageTaggedItems(html) {",
            "function collectRollupSources() {",
        )
        self.assertIn(
            '.page-item[data-tag="action"], .page-item[data-tag="coordination"]',
            extract_block,
        )

        collect_block = self._block(
            script,
            "function collectRollupSources() {",
            "function rollupSourceTitle(source) {",
        )
        # project pages, deliverable pages, and global pages all contribute.
        self.assertIn('kind: "project"', collect_block)
        self.assertIn('kind: "deliverable"', collect_block)
        self.assertIn('kind: "global"', collect_block)

    def test_rollup_view_mode_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        self.assertIn("function normalizeProjectsViewMode(value) {", script)
        self.assertIn('if (value === "card" || value === "rollup") return value;', script)
        # Legacy "coordination" view mode migrates to the roll-up.
        self.assertIn('if (value === "coordination") return "rollup";', script)
        self.assertIn('setProjectsViewMode("rollup")', script)
        self.assertIn(
            'if (projectsViewMode === "rollup") renderPageItemsRollupView();',
            script,
        )

        render_block = self._block(
            script,
            "function render() {",
            "function updateSortHeaders() {",
        )
        self.assertIn("if (isRollupView) {", render_block)
        self.assertIn("renderPageItemsRollupView();", render_block)

        view_mode_block = self._block(
            script,
            "function updateProjectsViewModeUi() {",
            "function updateProjectsLayoutWidthUi() {",
        )
        self.assertIn('document.getElementById("viewToggleRollup")', view_mode_block)
        self.assertIn('document.getElementById("projectsRollupView")', view_mode_block)

    def test_rollup_markup_exists(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")
        self.assertIn('id="viewToggleRollup"', html)
        self.assertIn('data-view-mode="rollup"', html)
        self.assertIn('id="projectsRollupView"', html)

    def test_rollup_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")
        self.assertIn(".projects-rollup-view {", css)
        self.assertIn(".rollup-source-panel {", css)
        self.assertIn(".rollup-source-header {", css)
        self.assertIn(".rollup-item-row {", css)
        self.assertIn('.rollup-item-badge[data-tag="coordination"] {', css)
        self.assertIn(".page-item-flash {", css)

    def test_legacy_coordination_feature_removed(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        # Script: coordination feature functions/consts are gone.
        self.assertNotIn("const COORDINATION_PARTIES = [", script)
        self.assertNotIn("const COORDINATION_ICON_PATH =", script)
        self.assertNotIn("function openCoordinationDialog(", script)
        self.assertNotIn("function renderCoordinationView(", script)
        self.assertNotIn("function renderCoordinationProjectPanel(", script)
        self.assertNotIn("function addCoordinationItemFromDialog(", script)
        self.assertNotIn("function createDeliverableCoordinationAction(", script)
        self.assertNotIn(
            "coordinationItems: normalizeCoordinationItems(project.coordinationItems),",
            script,
        )

        # Markup: coordination dialog and old view toggle are gone.
        self.assertNotIn('id="coordinationDlg"', html)
        self.assertNotIn('id="viewToggleCoordination"', html)
        self.assertNotIn('id="projectsCoordinationView"', html)

        # Styles: coordination-specific blocks are gone.
        self.assertNotIn(".coordination-dialog {", css)
        self.assertNotIn(".projects-coordination-view {", css)
        self.assertNotIn(".coordination-project-panel {", css)


if __name__ == "__main__":
    unittest.main()
