import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class DeliverableToolDropdownUiTests(unittest.TestCase):
    def test_shared_tool_registry_has_category_field_per_entry(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "const SHARED_TOOL_LAUNCH_REGISTRY = Object.freeze([",
            'id: "toolCopyProjectLocally"',
            'menuLabel: "Local Projects"',
            'id: "toolPublishDwgs"',
            'menuLabel: "Publish"',
            'id: "toolManageLayers"',
            'menuLabel: "Freeze/Thaw"',
            'id: "toolCleanXrefs"',
            'menuLabel: "Prepare XREFs"',
            'id: "toolCreateNarrativeTemplate"',
            'menuLabel: "Create NOC"',
            'id: "toolCreatePlanCheckTemplate"',
            'menuLabel: "Create PCC"',
            'id: "toolWireSizer"',
            'id: "toolCircuitBreaker"',
            'id: "toolBackupDrawings"',
            'menuLabel: "Backup DWGs"',
            'id: "toolLightingSchedule"',
            "isReady: false,",
            'category: "general",',
            'category: "electrical",',
            'category: "templates",',
            "function getReadySharedToolLaunchEntries() {",
            "function getDeliverableToolMenuEntries() {",
            "function getDeliverableToolMenuEntriesByCategory() {",
            "const DELIVERABLE_TOOL_CATEGORIES = Object.freeze([",
            '{ key: "general", label: "General" },',
            '{ key: "electrical", label: "Electrical" },',
            '{ key: "templates", label: "Templates" },',
            "function buildProjectsTabToolLaunchContext(project, deliverable) {",
            'source: "projects-tab",',
            "function launchSharedToolCard(toolId, launchContext = null) {",
        ):
            self.assertIn(expected, text)

    def test_unified_deliverable_actions_dropdown_wiring_exists(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "function createDeliverableActionsDropdown(deliverable, project, card) {",
            'className: "deliverable-actions-trigger",',
            'className: "deliverable-actions-menu"',
            "function buildDeliverableActionsItem({",
            "function buildDeliverableActionsSubmenu(label, renderSubmenu) {",
            "getDeliverableToolMenuEntriesByCategory()",
            '"Unpin deliverable" : "Pin deliverable"',
            '"Expand details" : "Collapse details"',
            '"Expand details"',
            '"Collapse details"',
            'buildDeliverableActionsSubmenu("Status"',
            'buildDeliverableActionsSubmenu("Tools"',
            'label: "Attachments",',
            'label: "Open Project Folder",',
            'label: "Edit Project",',
            'label: "Delete Project",',
            "openEdit(projectIndex)",
            "removeProject(projectIndex)",
            "openAttachmentPanel(attachmentContext)",
            "async function handleDeliverableActionsDrop(context, event) {",
            "addDroppedEmailToAttachmentContext(context, event)",
            "isSupportedEmailFile",
            "const actionsDropdown = createDeliverableActionsDropdown(deliverable, project, card);",
            "actions.append(actionsDropdown);",
            "const parentRect = parent.getBoundingClientRect();\n"
            "    const containingMenuRect =\n"
            "      wrapper.parentElement?.getBoundingClientRect?.() || parentRect;\n"
            "    const menuOverlap = 1;",
            "const topCorrection = targetTop - placedRect.top;",
            "const leftCorrection = targetLeft - placedRect.left;",
            "let targetTop = Math.max(viewportPadding, parentRect.top);",
        ):
            self.assertIn(expected, text)

    def test_deliverable_actions_close_behavior_is_wired(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "function closeOpenDeliverableActionsDropdown({ except = null, focusTrigger = false } = {}) {",
            "function setDeliverableActionsDropdownState(dropdown, isOpen) {",
            "function ensureDeliverableActionsGlobalHandlers() {",
            "function closeDeliverableActionsSiblingSubmenus(wrapper) {",
            'document.addEventListener("keydown", (e) => {',
            'if (e.key !== "Escape" || !openDeliverableActionsDropdown) return;',
            "closeOpenDeliverableActionsDropdown({ focusTrigger: true });",
            "closeOpenDeliverableToolDropdown();",
            "closeOpenDeliverableStatusDropdowns();",
            "if (openAttachmentPanelContext) {",
            "void requestAttachmentPanelClose();",
        ):
            self.assertIn(expected, text)

    def test_deliverable_actions_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        for expected in (
            ".deliverable-actions-dropdown {",
            ".deliverable-actions-trigger {",
            ".deliverable-actions-trigger.is-dragover {",
            ".deliverable-actions-trigger.has-attachments::after {",
            ".deliverable-actions-trigger-icon {",
            ".deliverable-actions-menu {",
            ".deliverable-actions-menu.open {",
            ".deliverable-actions-item {",
            ".deliverable-actions-item.has-submenu::after {",
            ".deliverable-actions-item.danger {",
            ".deliverable-actions-divider {",
            ".deliverable-actions-submenu-wrapper {",
            ".deliverable-actions-submenu {",
            ".deliverable-actions-submenu.open {",
            "left: -8px;",
            "right: -8px;",
        ):
            self.assertIn(expected, css)


if __name__ == "__main__":
    unittest.main()
