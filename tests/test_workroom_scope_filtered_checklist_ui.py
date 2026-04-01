import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
INDEX_HTML_PATH = REPO_ROOT / "index.html"
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class WorkroomGuidedFlowUiTests(unittest.TestCase):
    def test_workroom_markup_uses_survey_first_shell_and_right_rail_work_items(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")

        for expected in (
            'id="workroomPhaseTabs"',
            'id="workroomSurveySection"',
            'id="workroomSurveyLaunchBtn"',
            'id="workroomPreDesignSection"',
            'id="workroomChecklistSection"',
            'id="workroomPostPermitSection"',
            'id="workroomDesignGeneralNotesSection"',
            'id="workroomUtilityPaneTasks"',
            'id="workroomUtilityPaneNotes"',
            'id="workroomRecommendedToolsList"',
            'id="workroomAnytimeToolsList"',
            'id="workroomToolStatus"',
            'id="workroomNarrativeTemplateBtn"',
            'id="workroomPlanCheckTemplateBtn"',
            'class="workroom-panel workroom-workitems-panel"',
        ):
            self.assertIn(expected, html)

        for removed in (
            'class="workroom-panel workroom-utility-panel"',
            'id="workroomUtilityTabGeneral"',
            'id="workroomUtilityPaneGeneral"',
            'data-workroom-left-tab="scope"',
            'id="workroomLeftPaneScope"',
            'id="workroomScopeContent"',
            'id="workroomCadFilesStatus"',
            'id="workroomCadFilesList"',
        ):
            self.assertNotIn(removed, html)

    def test_script_tracks_phase_state_return_context_and_workroom_state(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "function normalizeWorkroomPhase(value) {",
            "function normalizeWorkroomReturnType(value) {",
            'value: "survey_report"',
            "workroomPhase: normalizeWorkroomPhase(deliverable.workroomPhase),",
            "workroomReturnType: normalizeWorkroomReturnType(deliverable.workroomReturnType),",
            'workroomPhase: seed.workroomPhase || "survey_report",',
            'workroomReturnType: seed.workroomReturnType || "",',
            "function renderWorkroomPhaseTabs() {",
            "function renderWorkroomUtilityTabs() {",
            'return ["tasks", "notes"];',
            "workroomPhase: phaseValue,",
            "workroomReturnType: normalizeWorkroomReturnType(deliverable?.workroomReturnType),",
            "visibleUtilityTabs: getVisibleWorkroomUtilityTabs(phaseValue),",
            'phaseTabCount: document.querySelectorAll("#workroomPhaseTabs .workroom-phase-tab").length,',
        ):
            self.assertIn(expected, script)

    def test_script_uses_survey_phase_and_trimmed_pre_design_walkthrough(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            'key: "buildingAreaSqFt"',
            'key: "storyCount"',
            'key: "surveyStatus"',
            'key: "sitePhotosStatus"',
            'key: "panelPhotosStatus"',
            'section: "scope_jurisdiction"',
            'section: "building_profile"',
            'section: "field_verification"',
            "const WORKROOM_PRE_DESIGN_PAGES = [",
            "function renderWorkroomSurveyPhase() {",
            "Field Package First",
            "function renderWorkroomScopePanel() {",
            "These answers are shared across every deliverable on the project.",
        ):
            self.assertIn(expected, script)

        self.assertNotIn('key: "field_verification"', script)

    def test_script_preserves_scope_filtered_checklist_logic_inside_phase_flow(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "checklistProfiles: normalizeProjectChecklistProfiles(",
            "checklistContext: normalizeChecklistContext(",
            "applicabilityOverrides: normalizeChecklistApplicabilityOverrides(",
            "function buildChecklistScopeAnswerState(",
            "function evaluateChecklistItemApplicability(",
            'visibility: "filtered"',
            'visibility: "unresolved"',
            'className: "workroom-checklist-filtered-group"',
            'textContent: `Filtered for this project (${section.filteredItems.length})`',
            'label = "Show anyway";',
            'label = "Hide for this deliverable";',
            'label = "Clear override";',
            "setChecklistApplicabilityOverride(",
            "viewModel.summary.visibleCompletedTotal",
            "viewModel.summary.visibleTotal",
        ):
            self.assertIn(expected, script)

    def test_script_maps_phase_specific_checklists_and_tools(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "const WORKROOM_PHASES = WORKROOM_PHASE_OPTIONS.reduce(",
            "const WORKROOM_PHASE_CHECKLIST_MAP = {",
            'title: "Electrical TI - General Checklist"',
            'title: "Preflight Electrical Checklist"',
            "const WORKROOM_PHASE_TOOL_MAP = {",
            'survey_report: ["toolCreateElectricalSurveyTemplate"]',
            'pre_design: ["toolCleanXrefs"]',
            'design: ["toolLightingSchedule", "toolCircuitBreaker"]',
            'preflight: ["toolFreezeLayers", "toolPublishDwgs", "toolThawLayers"]',
            'post_permit: ["toolCreateNarrativeTemplate", "toolCreatePlanCheckTemplate"]',
            "const WORKROOM_ALWAYS_AVAILABLE_TOOLS = [",
            '"toolCopyProjectLocally"',
            '"toolBackupDrawings"',
            '"toolWireSizer"',
            '"toolCircuitBreaker"',
            "function renderWorkroomToolsPanel() {",
        ):
            self.assertIn(expected, script)

    def test_styles_define_phase_shell_tool_groups_and_postpermit_layout(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        for expected in (
            ".workroom-main-column {",
            ".workroom-tools-column {",
            ".workroom-phase-shell {",
            ".workroom-phase-section[hidden],",
            ".workroom-utility-pane[hidden],",
            ".workroom-checklist-shell {",
            ".workroom-checklist-progress,",
            ".workroom-workitems-panel {",
            ".workroom-design-notes-section {",
            ".workroom-phase-summary-grid {",
            ".workroom-tool-groups {",
            ".workroom-tool-item {",
            ".workroom-tool-copy {",
            ".workroom-tool-body {",
            ".workroom-tool-meta {",
            ".workroom-survey-launch-btn {",
            ".workroom-survey-checklist {",
            ".workroom-postpermit-grid {",
            ".workroom-postpermit-controls {",
            ".workroom-postpermit-actions {",
            ".workroom-scope-input {",
            ".workroom-checklist-status-badge {",
            ".workroom-checklist-filtered-group {",
        ):
            self.assertIn(expected, css)


if __name__ == "__main__":
    unittest.main()
