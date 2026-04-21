import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class DeliverableToolDropdownUiTests(unittest.TestCase):
    def test_deliverable_tool_dropdown_registry_and_launch_context_wiring_exist(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "const SHARED_TOOL_LAUNCH_REGISTRY = Object.freeze([",
            'id: "toolCopyProjectLocally"',
            'label: "Local Project Manager"',
            'menuLabel: "Local Projects"',
            'id: "toolPublishDwgs"',
            'label: "Publish CAD DWGs in Headless Mode"',
            'menuLabel: "Publish"',
            'id: "toolManageLayers"',
            'menuLabel: "Freeze/Thaw"',
            'id: "toolCleanXrefs"',
            'label: "Prepare CAD DWG for Reference"',
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
            "function getReadySharedToolLaunchEntries() {",
            "function getDeliverableToolMenuEntries() {",
            '    "toolCopyProjectLocally",',
            "function buildProjectsTabToolLaunchContext(project, deliverable) {",
            'source: "projects-tab",',
            "function launchSharedToolCard(toolId, launchContext = null) {",
            "pendingCadLaunchContext = launchContext ? deepCloneJson(launchContext, null) : null;",
            "card.click();",
            "function createDeliverableToolDropdown(deliverable, project, card) {",
            "getDeliverableToolMenuEntries().forEach((entry) => {",
            'className: "deliverable-tool-option",',
            "textContent: entry.menuLabel || entry.label,",
            'const launchContext = buildProjectsTabToolLaunchContext(project, deliverable);',
            "launchSharedToolCard(entry.id, launchContext);",
            'const toolDropdown = createDeliverableToolDropdown(deliverable, project, card);',
            "actions.append(attachmentControl, toolDropdown, expandToggle);",
        ):
            self.assertIn(expected, text)

    def test_deliverable_tool_dropdown_close_behavior_is_wired(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "function closeOpenDeliverableToolDropdown({ except = null, focusTrigger = false } = {}) {",
            "function closeOpenDeliverableStatusDropdowns(exceptDropdown = null) {",
            "function ensureDeliverableToolDropdownGlobalHandlers() {",
            'document.addEventListener("keydown", (e) => {',
            'if (e.key !== "Escape" || !openDeliverableToolDropdown) return;',
            "closeOpenDeliverableToolDropdown({ focusTrigger: true });",
            "closeOpenDeliverableStatusDropdowns();",
            "closeOpenDeliverableToolDropdown();",
            "if (openAttachmentPanelContext) {",
            "void requestAttachmentPanelClose();",
        ):
            self.assertIn(expected, text)

    def test_deliverable_tool_dropdown_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        for expected in (
            ".deliverable-tool-dropdown {",
            ".deliverable-tool-trigger {",
            ".deliverable-tool-trigger-icon {",
            ".deliverable-tool-menu {",
            ".deliverable-tool-menu.open {",
            ".deliverable-tool-option {",
        ):
            self.assertIn(expected, css)


if __name__ == "__main__":
    unittest.main()
