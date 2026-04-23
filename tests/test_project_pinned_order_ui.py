import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
INDEX_HTML_PATH = REPO_ROOT / "index.html"
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectPinnedDeliverablesUiTests(unittest.TestCase):
    def test_list_view_uses_deliverable_pin_state_for_pinned_section(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "isPinnedDeliverable: isDeliverablePinned(deliverable),",
            "const aPinned = !!a?.isPinnedDeliverable;",
            "const bPinned = !!b?.isPinnedDeliverable;",
            "if (row?.isPinnedDeliverable || !shouldSortCompletedProjectsLast()) return 0;",
            "const {\n      project,\n      projectIndex,\n      projectListContext,\n      deliverable,\n      dueDate,\n      isPinnedDeliverable,",
            "if (isPinnedDeliverable) {",
            'appendSectionSeparator("Pinned Deliverables");',
        ):
            self.assertIn(expected, script)

        self.assertNotIn("appendSectionSeparator(\"Pinned Projects\");", script)
        self.assertNotIn("const pinned = getPinnedProjectsInManualOrder(items);", script)
        self.assertNotIn("items = pinned.concat(unpinned);", script)
        self.assertNotIn("isPinnedProject:", script)

    def test_card_view_uses_same_deliverable_rows_as_list_view(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        card_view_start = script.index("function renderCardView(")
        card_view_end = script.index("function setProjectsViewMode(", card_view_start)
        card_view_block = script[card_view_start:card_view_end]
        render_start = script.index("function render() {")
        render_end = script.index("function renderStatusToggles(", render_start)
        render_block = script[render_start:render_end]

        for expected in (
            "function renderCardView(items = db, projectListContextMap = null) {",
            "const deliverableRows = buildProjectDeliverableRowEntries(",
            "sortProjectDeliverableRows(deliverableRows);",
            "for (const { project, deliverable, dueDate, isPinnedDeliverable } of deliverableRows) {",
            "if (pinnedShown && isPinnedDeliverable) {",
            "buckets.get(\"pinned\").push({ project, deliverable });",
        ):
            self.assertIn(expected, card_view_block)

        self.assertIn("renderCardView(items, projectListContextMap);", render_block)

    def test_card_view_drop_from_pinned_column_unpins_deliverable_before_status_update(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        drop_start = script.index('host.addEventListener("drop", async (e) => {')
        drop_end = script.index("await save();", drop_start)
        drop_block = script[drop_start:drop_end]

        self.assertIn(
            'const { deliverableId, sourceColumnKey } = kanbanDragState;',
            drop_block,
        )
        self.assertIn("setDeliverablePinnedState(deliverable, true);", drop_block)
        self.assertIn(
            'if (sourceColumnKey === "pinned" && deliverable.pinned) {',
            drop_block,
        )
        self.assertIn("setDeliverablePinnedState(deliverable, false);", drop_block)
        self.assertLess(
            drop_block.index("setDeliverablePinnedState(deliverable, false);"),
            drop_block.index("deliverable.statuses = [targetKey];"),
        )
        self.assertNotIn("setProjectPinnedState(project", drop_block)

    def test_project_pin_ui_is_removed(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        for removed in (
            'id="pinUrgentDeliverablesBtn"',
            'class="pin-header"',
            'class="pin-btn"',
            'class="cell-select"',
            'aria-label="Pin project"',
        ):
            self.assertNotIn(removed, html)

        for removed in (
            "pinUrgentDeliverablesBtn",
            'tr.classList.toggle("is-pinned-project", isPinned);',
            "if (isPinned) enablePinnedProjectRowDrag(tr, pinBtn, project);",
            "setProjectPinnedState(project, !project?.pinned, db);",
        ):
            self.assertNotIn(removed, script)

        for removed in (
            ".pin-btn {",
            ".pin-header {",
            ".table th.cell-select {",
            ".row td.cell-select {",
            ".project-row.is-pinned-project .pin-btn {",
        ):
            self.assertNotIn(removed, css)


if __name__ == "__main__":
    unittest.main()
