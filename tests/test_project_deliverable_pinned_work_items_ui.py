import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectDeliverablePinnedWorkItemsUiTests(unittest.TestCase):
    @staticmethod
    def _css_block(css: str, selector: str) -> str:
        block_start = css.index(selector)
        block_end = css.index("}", block_start)
        return css[block_start:block_end]

    @staticmethod
    def _script_block(script: str, start_marker: str, end_marker: str) -> str:
        block_start = script.index(start_marker)
        block_end = script.index(end_marker, block_start)
        return script[block_start:block_end]

    @staticmethod
    def _script_block_within(
        script: str,
        container_start: str,
        start_marker: str,
        end_marker: str,
    ) -> str:
        container_index = script.index(container_start)
        block_start = script.index(start_marker, container_index)
        block_end = script.index(end_marker, block_start)
        return script[block_start:block_end]

    def test_projects_pinned_work_items_script_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertEqual(script.count("function createNotesSection("), 1)
        self.assertEqual(script.count("function createTasksPreview("), 1)
        self.assertEqual(script.count("function renderDeliverableCard("), 1)

        for expected in (
            "function getPinnedFirstEntries(items = [], normalizeItem = (item) => item) {",
            "function getPinnedDeliverableTaskEntries(deliverable) {",
            "function getPinnedDeliverableNoteEntries(deliverable) {",
            "function getPinnedDeliverablePreviewItems(deliverable) {",
            "function getExpandedProjectDeliverableIds() {",
            "function restoreExpandedProjectDeliverables(expandedIds = []) {",
            "function renderProjectsPreservingExpandedDeliverables() {",
            "function createWorkItemPinButton({",
            "function renderDeliverableStatusBadges(container, deliverable) {",
            "function createDeliverableStatusSection(deliverable, project, card) {",
            "function renderDeliverablePinnedPreview(container, deliverable) {",
            "function updateDeliverableWorkItemUi(card, deliverable) {",
            "function createPinnedAttachmentPreviewButton(attachment) {",
            "function createPinnedWorkItemAttachments(previewItem = {}) {",
            "function createNotesSection(deliverable, card, project = null) {",
            "function createTasksPreview(deliverable, card, project = null) {",
            "function renderDeliverableCard(deliverable, isPrimary, project) {",
            'className: "deliverable-status-inline-group",',
            'className: "deliverable-pinned-inline-group"',
            'titleUnpinned: "Pin task",',
            "renderNoteList();",
            "pinned: !!task?.pinned,",
            "pinned: !!noteItem?.pinned,",
            'pillIcon.classList.add("deliverable-pinned-inline-pill-icon");',
            'className: `deliverable-pinned-inline-pill ${',
            'className: "deliverable-pinned-inline-pill-text"',
            'className: "deliverable-pinned-item-attachments"',
            'className: "deliverable-pinned-link"',
            "renderProjectsPreservingExpandedDeliverables();",
            "const notesSection = createNotesSection(deliverable, card, project);",
            'card.dataset.deliverableId = deliverableId;',
        ):
            self.assertIn(expected, script)

        self.assertNotIn("function createNotesSection(deliverable, project) {", script)
        self.assertNotIn("function createTasksPreview(deliverable, card) {", script)
        self.assertNotIn("const notesSection = createNotesSection(deliverable, project);", script)
        self.assertNotIn('className: "deliverable-pinned-preview-heading"', script)
        self.assertNotIn('className: "deliverable-pinned-preview"', script)
        self.assertNotIn('className: "deliverable-pinned-item-kind"', script)
        self.assertNotIn('className: "deliverable-pinned-item-kind-label"', script)
        self.assertNotIn("createPinnedWorkItemAttachments(previewItem);", script)

    def test_status_row_groups_pinned_preview_and_top_actions(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        status_badges_block = self._script_block(
            script,
            "function renderDeliverableStatusBadges(container, deliverable) {",
            "function setDeliverableStatusDropdownState(dropdown, isOpen) {",
        )
        self.assertNotIn('className: "deliverable-pinned-inline-group"', status_badges_block)
        self.assertNotIn("renderDeliverablePinnedPreview(", status_badges_block)

        status_section_block = self._script_block(
            script,
            "function createDeliverableStatusSection(deliverable, project, card) {",
            "function createNotesSectionLegacy(deliverable, project) {",
        )
        for expected in (
            'className: "deliverable-status-row"',
            'className: "deliverable-status-inline-group"',
            'className: "deliverable-pinned-inline-group"',
            "renderDeliverablePinnedPreview(pinnedHost, deliverable);",
            "statusInlineGroup.append(pinnedHost);",
            "statusSection.append(statusBadges, statusInlineGroup);",
            "updateDeliverableStatusRowVisibility(statusSection);",
        ):
            self.assertIn(expected, status_section_block)
        self.assertNotIn("statusInlineGroup.append(pinnedHost, statusDropdown);", status_section_block)

        top_actions_block = self._script_block(
            script,
            "function createDeliverableCardTopActions(deliverable, project, card) {",
            "let openDeliverableActionsDropdown = null;",
        )
        for expected in (
            'className: "deliverable-card-action-row"',
            'className: "deliverable-card-action-group deliverable-card-action-group-left"',
            'className: "deliverable-card-action-group deliverable-card-action-group-right"',
            "const statusDropdown = createStatusDropdown(deliverable, project, card);",
            'statusDropdown.classList.add("deliverable-card-status-action");',
            "const toolDropdown = createDeliverableToolDropdown(deliverable, project, card);",
            'toolDropdown.classList.add("deliverable-card-tool-action");',
            "leftActions.append(pinBtn, statusDropdown, toolDropdown);",
            "rightActions.append(",
            "rightActions.append(attachmentBtn);",
            "actions.append(leftActions, rightActions);",
        ):
            self.assertIn(expected, top_actions_block)
        self.assertNotIn("createExpandToggle(card)", top_actions_block)

        self.assertNotIn("function createDeliverableCardBottomActions(", script)
        self.assertNotIn("const bottomActions = createDeliverableCardBottomActions(", script)

        update_work_item_ui_block = self._script_block(
            script,
            "function updateDeliverableWorkItemUi(card, deliverable) {",
            "function createWorkItemDoneCheckbox({",
        )
        for expected in (
            "renderDeliverableStatusBadges(",
            'card?.querySelector(".deliverable-status-badges")',
            "renderDeliverablePinnedPreview(",
            'card?.querySelector(".deliverable-pinned-inline-group")',
            "updateDeliverableStatusRowVisibility(",
            'card?.querySelector(".deliverable-status-row")',
        ):
            self.assertIn(expected, update_work_item_ui_block)

        legacy_card_block = self._script_block(
            script,
            "function renderDeliverableCardLegacy(deliverable, isPrimary, project) {",
            "function renderDeliverablePinnedPreview(container, deliverable) {",
        )
        current_card_block = self._script_block(
            script,
            "function renderDeliverableCard(deliverable, isPrimary, project) {",
            "function normalizeProjectMatchValue(value) {",
        )
        self.assertIn(
            "const statusSection = createDeliverableStatusSection(",
            legacy_card_block,
        )
        self.assertIn(
            "const actionRow = createDeliverableCardTopActions(deliverable, project, card);",
            legacy_card_block,
        )
        self.assertIn(
            "card.append(actionRow, header, progress, statusSection, tasksPreview, notesSection);",
            legacy_card_block,
        )
        self.assertIn(
            "const statusSection = createDeliverableStatusSection(",
            current_card_block,
        )
        self.assertIn(
            "const actionRow = createDeliverableCardTopActions(deliverable, project, card);",
            current_card_block,
        )
        self.assertIn(
            "card.append(actionRow, header, statusSection, notesSection);",
            current_card_block,
        )
        self.assertNotIn("bottomActions", legacy_card_block)
        self.assertNotIn("bottomActions", current_card_block)

    def test_status_change_rerenders_projects_while_restoring_expanded_cards(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        helper_block = self._script_block(
            script,
            "function getExpandedProjectDeliverableIds() {",
            "function renderDeliverableCardLegacy(deliverable, isPrimary, project) {",
        )
        for expected in (
            'document.getElementById("tbody")',
            '.querySelectorAll(".deliverable-card-new[data-deliverable-id]")',
            'if (card.classList.contains("details-collapsed")) return;',
            "setDeliverableDetailsCollapsed(card, false);",
            "function renderProjectsPreservingExpandedDeliverables() {",
            "const expandedIds = getExpandedProjectDeliverableIds();",
            "render();",
            "restoreExpandedProjectDeliverables(expandedIds);",
        ):
            self.assertIn(expected, helper_block)

        status_dropdown_block = self._script_block(
            script,
            "function createStatusDropdown(deliverable, project, card) {",
            "function createDeliverableStatusSection(deliverable, project, card) {",
        )
        for expected in (
            "await save();",
            "setDeliverableStatusDropdownState(dropdown, false);",
            "renderProjectsPreservingExpandedDeliverables();",
        ):
            self.assertIn(expected, status_dropdown_block)

        self.assertNotIn(
            'const badgesContainer = card.querySelector(".deliverable-status-badges");',
            status_dropdown_block,
        )
        self.assertNotIn(
            "renderDeliverableStatusBadges(badgesContainer, deliverable);",
            status_dropdown_block,
        )

    def test_projects_pinned_work_items_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        for expected in (
            ".deliverable-status-row {",
            ".deliverable-status-row[hidden] {",
            ".deliverable-status-badges {",
            ".deliverable-status-inline-group {",
            ".deliverable-status-dropdown {",
            ".deliverable-card-action-row {",
            ".deliverable-card-action-group {",
            ".deliverable-notes-footer {",
            ".deliverable-note-add-btn {",
            ".deliverable-note-row {",
            ".deliverable-note-delete-btn {",
            ".deliverable-note-add-input {",
            ".deliverable-notes-toggle {",
            ".work-item-actions {",
            ".work-item-pin-btn {",
            ".work-item-pin-btn.is-pinned {",
            ".deliverable-pinned-inline-group {",
            ".deliverable-pinned-inline-pill {",
            ".deliverable-pinned-inline-pill-icon {",
            ".deliverable-pinned-inline-pill-text {",
            ".deliverable-pinned-inline-pill.is-task {",
            ".deliverable-pinned-item-attachments {",
            ".deliverable-pinned-link,",
        ):
            self.assertIn(expected, css)

        self.assertNotIn(".deliverable-pinned-preview-heading {", css)
        self.assertNotIn(".deliverable-pinned-preview {", css)
        self.assertNotIn(".deliverable-pinned-item-kind {", css)
        self.assertNotIn(".deliverable-pinned-inline-pill.is-note {", css)

        status_row_block = self._css_block(css, ".deliverable-status-row {")
        self.assertIn("display: flex;", status_row_block)
        self.assertIn("align-items: flex-start;", status_row_block)
        self.assertIn("flex-wrap: wrap;", status_row_block)
        self.assertIn("gap: 0.5rem;", status_row_block)
        self.assertNotIn("display: block;", status_row_block)

        status_badges_block = self._css_block(css, ".deliverable-status-badges {")
        self.assertIn("align-items: center;", status_badges_block)
        self.assertNotIn("width: fit-content;", status_badges_block)
        self.assertNotIn("vertical-align: top;", status_badges_block)

        status_inline_group_block = self._css_block(
            css,
            ".deliverable-status-inline-group {",
        )
        self.assertIn("display: inline-flex;", status_inline_group_block)
        self.assertIn("align-items: flex-start;", status_inline_group_block)
        self.assertIn("flex-wrap: wrap;", status_inline_group_block)
        self.assertIn("gap: 0.375rem;", status_inline_group_block)
        self.assertIn("max-width: 100%;", status_inline_group_block)

        status_dropdown_block = self._css_block(css, ".deliverable-status-dropdown {")
        self.assertIn("display: inline-block;", status_dropdown_block)
        self.assertIn("flex: 0 0 auto;", status_dropdown_block)
        self.assertNotIn("margin-inline-start:", status_dropdown_block)
        self.assertNotIn("vertical-align:", status_dropdown_block)

        pinned_group_block = self._css_block(css, ".deliverable-pinned-inline-group {")
        self.assertIn("display: contents;", pinned_group_block)
        self.assertNotIn("display: inline-flex;", pinned_group_block)
        self.assertNotIn("align-items:", pinned_group_block)
        self.assertNotIn("max-width:", pinned_group_block)

        pinned_group_hidden_block = self._css_block(
            css,
            ".deliverable-pinned-inline-group[hidden] {",
        )
        self.assertIn("display: none;", pinned_group_hidden_block)

        pill_block = self._css_block(css, ".deliverable-pinned-inline-pill {")
        self.assertIn("flex: 0 1 auto;", pill_block)
        self.assertIn("align-items: flex-start;", pill_block)
        self.assertIn("max-width: min(100%, 240px);", pill_block)

        pill_text_block = self._css_block(css, ".deliverable-pinned-inline-pill-text {")
        self.assertIn("white-space: normal;", pill_text_block)
        self.assertIn("overflow-wrap: anywhere;", pill_text_block)
        self.assertNotIn("text-overflow: ellipsis;", pill_text_block)
        self.assertNotIn("white-space: nowrap;", pill_text_block)

    def test_projects_notes_footer_supports_add_delete_and_overflow_toggle(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        notes_block = self._script_block(
            script,
            "function createNotesSection(deliverable, card, project = null) {",
            "function createTasksPreview(deliverable, card, project = null) {",
        )

        for expected in (
            'className: "deliverable-notes-footer"',
            'className: "deliverable-note-add-btn"',
            'className: "deliverable-note-row"',
            'className: "deliverable-note-delete-btn"',
            'className: "deliverable-note-add-input"',
            'className: "deliverable-notes-toggle"',
            "noteEntries.slice(0, 2)",
            'toggleBtn.textContent = notesExpanded ? "Hide notes" : "Show all notes";',
            'noteInput.addEventListener("keydown", (e) => {',
            'void finishAddingNote("commit");',
            'void finishAddingNote("cancel");',
            "deliverable.noteItems.push({",
            "deliverable.noteItems.splice(index, 1);",
            "syncDeliverableWorkItemFields(deliverable);",
            "syncDeliverableNoteFields(deliverable);",
            "renderNoteList();",
            "updateDeliverableWorkItemUi(card, deliverable);",
            "await save();",
        ):
            self.assertIn(expected, notes_block)

        self.assertNotIn(
            'const pinBtn = createWorkItemPinButton({',
            notes_block,
        )

    def test_modal_note_rows_do_not_expose_pin_controls(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        modal_note_block = self._script_block(
            script,
            "function renderModalDeliverableNoteList(card, container, options = {}) {",
            "function commitPendingModalDeliverableNote(card, container, options = {}) {",
        )

        self.assertIn("actions.append(dueControl, attachments, deleteBtn);", modal_note_block)
        self.assertNotIn('const pinBtn = createWorkItemPinButton({', modal_note_block)
        self.assertNotIn('titlePinned: "Unpin note",', modal_note_block)
        self.assertNotIn('titleUnpinned: "Pin note",', modal_note_block)
        self.assertNotIn("liveNote.pinned = nextPinned;", modal_note_block)

    def test_work_item_pin_button_supports_async_toggle_rollback(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        pin_button_block = self._script_block(
            script,
            "function createWorkItemPinButton({",
            "function getModalDeliverableCardContext(card) {",
        )

        for expected in (
            "let isTogglePending = false;",
            'button.setAttribute("aria-busy", String(isTogglePending));',
            "if (isTogglePending) return;",
            "const previousPinned = currentPinned;",
            "const nextPinned = !currentPinned;",
            "const resolvedPinned = onToggle ? await onToggle(nextPinned) : nextPinned;",
            'if (typeof resolvedPinned === "boolean") {',
            "currentPinned = previousPinned;",
            'console.warn("Failed to toggle work item pin state:", error);',
        ):
            self.assertIn(expected, pin_button_block)


if __name__ == "__main__":
    unittest.main()
