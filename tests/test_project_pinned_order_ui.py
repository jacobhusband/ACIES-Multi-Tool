import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectPinnedOrderUiTests(unittest.TestCase):
    def test_project_pinned_order_script_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "let pinnedProjectDragState = null;",
            "let projectPinHandleSuppressClickUntil = 0;",
            "function normalizePinnedProjectOrder(value) {",
            "function getPinnedProjectsInManualOrder(items = db) {",
            "function syncPinnedProjectOrders(items = db, { seedMissing = false } = {}) {",
            "function getNextPinnedProjectOrder(items = db) {",
            "function getLowestPinnedProjectOrder(projects = []) {",
            "function setProjectPinnedState(project, nextPinned, items = db) {",
            "function movePinnedProjectToTarget(project, targetProject, before = true, items = db) {",
            "pinnedOrder: normalizePinnedProjectOrder(project?.pinnedOrder),",
            "if (!out.pinned) out.pinnedOrder = null;",
            "if (syncPinnedProjectOrders(merged, { seedMissing: true })) {",
            "syncPinnedProjectOrders(db, { seedMissing: true });",
            "pinnedOrder: normalizePinnedProjectOrder(existingProject?.pinnedOrder),",
            "base.pinnedOrder = base.pinned ? getLowestPinnedProjectOrder(projects) : null;",
            "const pinned = getPinnedProjectsInManualOrder(items);",
            "items = pinned.concat(unpinned);",
            "pinned: false,",
            "pinnedOrder: null,",
        ):
            self.assertIn(expected, script)

    def test_project_pinned_order_drag_and_ungrouped_render_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "function clearPinnedProjectDragStyles() {",
            'tbody.querySelectorAll(".project-row").forEach((row) => {',
            'row.classList.remove("project-dragging", "project-drop-before", "project-drop-after");',
            "function enablePinnedProjectRowDrag(row, handle, project) {",
            "handle.draggable = true;",
            'row.classList.add("project-dragging");',
            'row.classList.toggle("project-drop-before", before);',
            'row.classList.toggle("project-drop-after", !before);',
            "const moved = movePinnedProjectToTarget(",
            "projectPinHandleSuppressClickUntil = Date.now() + 250;",
            "await save();",
            "renderProjectsPreservingExpandedDeliverables();",
            'tr.classList.add("project-row");',
            'tr.classList.toggle("is-pinned-project", isPinned);',
            "if (isPinned) enablePinnedProjectRowDrag(tr, pinBtn, project);",
            "if (Date.now() < projectPinHandleSuppressClickUntil) return;",
            "setProjectPinnedState(project, !project?.pinned, db);",
            "const orderedDeliverableRows = [];",
            "const pinnedDeliverableRows = new Map();",
            "const unpinnedDeliverableRows = [];",
            "items.filter((project) => project?.pinned).forEach((project) => {",
            "orderedDeliverableRows.forEach((row) => {",
        ):
            self.assertIn(expected, script)

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

    def test_project_pinned_order_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        for expected in (
            ".project-row.is-pinned-project .pin-btn {",
            ".project-row.is-pinned-project .pin-btn:active {",
            ".project-row.project-dragging {",
            ".project-row.project-drop-before td {",
            ".project-row.project-drop-after td {",
            "cursor: grab;",
            "cursor: grabbing;",
            "opacity: 0.6;",
            "border-top: 2px solid var(--accent);",
            "border-bottom: 2px solid var(--accent);",
        ):
            self.assertIn(expected, css)


if __name__ == "__main__":
    unittest.main()
