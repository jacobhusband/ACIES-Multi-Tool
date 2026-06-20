import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
INDEX_HTML_PATH = REPO_ROOT / "index.html"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectCoordinationUiTests(unittest.TestCase):
    @staticmethod
    def _block(text: str, start_marker: str, end_marker: str) -> str:
        start = text.index(start_marker)
        end = text.index(end_marker, start)
        return text[start:end]

    def test_coordination_script_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn("const COORDINATION_PARTIES = [", script)
        self.assertIn('"Project Manager",', script)
        self.assertIn('const COORDINATION_OTHER_PARTY_VALUE = "__other__";', script)
        self.assertIn("const COORDINATION_ICON_PATH =", script)

        self.assertIn("function normalizeCoordinationItem(item) {", script)
        self.assertIn("function normalizeCoordinationItems(items = []) {", script)
        self.assertIn("function getProjectCoordinationItems(project) {", script)
        self.assertIn("function getProjectOpenCoordinationCount(project) {", script)

        item_block = self._block(
            script,
            "function normalizeCoordinationItem(item) {",
            "function normalizeCoordinationItems(items = []) {",
        )
        self.assertIn(
            "dueDate: normalizeDeliverableNoteDueDate(item?.dueDate),", item_block
        )

        normalize_block = self._block(
            script,
            "function normalizeProject(project) {",
            "function getLatestDueDeliverableId(deliverables = []) {",
        )
        self.assertIn(
            "coordinationItems: normalizeCoordinationItems(project.coordinationItems),",
            normalize_block,
        )

        self.assertIn(
            "function createDeliverableCoordinationAction(deliverable, project, card) {",
            script,
        )
        self.assertIn("function updateCoordinationTriggerState(button, project) {", script)
        self.assertIn("function updateProjectCoordinationUi(project) {", script)
        self.assertIn("function openCoordinationDialog(project) {", script)
        self.assertIn("function renderCoordinationItemList() {", script)
        self.assertIn("async function addCoordinationItemFromDialog() {", script)
        self.assertIn("function populateCoordinationPartySelect() {", script)
        self.assertIn("function resolveCoordinationPartyValue() {", script)
        self.assertIn("function getCoordinationPartyHue(party) {", script)

        render_block = self._block(
            script,
            "function renderCoordinationItemList() {",
            "async function addCoordinationItemFromDialog() {",
        )
        self.assertIn("const dueControl = createNoteDueDateControl({", render_block)
        self.assertIn("--coordination-party-hue", render_block)
        self.assertIn("createInlineWorkItemTextControl({", render_block)
        self.assertIn('textClassName: "coordination-item-note",', render_block)

        add_block = self._block(
            script,
            "async function addCoordinationItemFromDialog() {",
            "function createDeliverableAttachmentAction(deliverable, project, card) {",
        )
        self.assertIn(
            "const dueDate = normalizeDeliverableNoteDueDate(dueInput?.value);",
            add_block,
        )

        actions_block = self._block(
            script,
            "function createDeliverableCardTopActions(deliverable, project, card) {",
            "let openDeliverableActionsDropdown = null;",
        )
        self.assertIn(
            "rightActions.append(coordinationBtn, attachmentBtn);", actions_block
        )

    def test_coordination_view_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn("function normalizeProjectsViewMode(value) {", script)
        self.assertIn(
            'return value === "card" || value === "coordination" ? value : "list";',
            script,
        )
        self.assertIn("function getProjectsWithCoordinationItems() {", script)
        self.assertIn("function getEarliestOpenCoordinationDueMs(project) {", script)
        self.assertIn("function sortCoordinationItemsForDisplay(items) {", script)
        self.assertIn("function renderCoordinationProjectPanel(project) {", script)
        self.assertIn("function renderCoordinationView() {", script)
        self.assertIn('setProjectsViewMode("coordination")', script)
        self.assertIn(
            'if (projectsViewMode === "coordination") renderCoordinationView();',
            script,
        )

        render_block = self._block(
            script,
            "function render() {",
            "function updateSortHeaders() {",
        )
        self.assertIn("if (isCoordinationView) {", render_block)
        self.assertIn("renderCoordinationView();", render_block)

        view_mode_block = self._block(
            script,
            "function updateProjectsViewModeUi() {",
            "function updateProjectsLayoutWidthUi() {",
        )
        self.assertIn(
            'document.getElementById("viewToggleCoordination")', view_mode_block
        )
        self.assertIn(
            'document.getElementById("projectsCoordinationView")', view_mode_block
        )

        panel_block = self._block(
            script,
            "function renderCoordinationProjectPanel(project) {",
            "function renderCoordinationView() {",
        )
        self.assertIn("createInlineWorkItemTextControl({", panel_block)
        self.assertIn('textClassName: "coordination-item-note",', panel_block)
        self.assertIn("coordination-project-item-row", panel_block)

    def test_coordination_dialog_markup_exists(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")

        self.assertIn('<dialog id="coordinationDlg"', html)
        self.assertIn('id="coordinationDlgSummary"', html)
        self.assertIn('id="coordinationPartySelect"', html)
        self.assertIn('id="coordinationCustomPartyField"', html)
        self.assertIn('id="coordinationCustomPartyInput"', html)
        self.assertIn('id="coordinationNoteInput"', html)
        self.assertIn('id="coordinationDueInput"', html)
        self.assertIn('id="coordinationAddBtn"', html)
        self.assertIn('id="coordinationItemList"', html)
        self.assertIn('<section class="coordination-composer"', html)
        self.assertIn("closeDlg('coordinationDlg')", html)
        self.assertIn('id="viewToggleCoordination"', html)
        self.assertIn('id="projectsCoordinationView"', html)

    def test_coordination_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        self.assertIn(
            ".deliverable-card-action-btn.deliverable-card-coordination-action {", css
        )
        self.assertIn(".deliverable-coordination-count {", css)
        self.assertIn(".has-open-coordination", css)
        self.assertIn(".is-coordination-clear", css)
        self.assertIn(".coordination-dialog {", css)
        self.assertIn(".coordination-composer {", css)
        self.assertIn(".coordination-composer-note-row {", css)
        self.assertIn(".coordination-item-list {", css)
        self.assertIn(".coordination-item-row {", css)
        self.assertIn(".coordination-item-party {", css)
        self.assertIn(".coordination-item-due {", css)
        self.assertIn(".coordination-item-delete-btn {", css)
        self.assertIn("--coordination-party-hue", css)
        self.assertIn(".projects-coordination-view {", css)
        self.assertIn(".projects-coordination-view[hidden] {", css)
        self.assertIn(".coordination-project-panel {", css)
        self.assertIn(".coordination-project-item-row {", css)
        self.assertIn(
            ".coordination-project-item-row > :is(.work-item-text-wrap) {", css
        )
        self.assertIn(
            ".coordination-project-item-row .coordination-item-due {", css
        )
        self.assertIn(".coordination-project-count {", css)
        self.assertIn(".coordination-view-empty {", css)

        project_row_block = self._block(
            css, ".coordination-project-item-row {", "}"
        )
        self.assertIn("display: grid;", project_row_block)
        self.assertIn(
            "grid-template-columns: auto minmax(0, 1fr) auto;",
            project_row_block,
        )
        self.assertIn('"check note delete"', project_row_block)
        self.assertIn('". party due"', project_row_block)


if __name__ == "__main__":
    unittest.main()
