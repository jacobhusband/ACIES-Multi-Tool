import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class WorkflowHotkeyUiTests(unittest.TestCase):
    def test_tools_tab_no_longer_exposes_workflow_ui(self):
        html = (ROOT / "index.html").read_text(encoding="utf-8")
        script = (ROOT / "script.js").read_text(encoding="utf-8")
        styles = (ROOT / "styles.css").read_text(encoding="utf-8")

        for removed in (
            'id="workflowsSection"',
            'id="workflowsGrid"',
            'id="settings_workflowCad_manageLayersElectricalTopLevel"',
            'id="settings_workflowCad_cleanXrefsElectricalXrefsToNewestArch"',
            'id="settings_workflowCad_cleanXrefsSearchZipArchives"',
            "Workflow CAD Defaults",
        ):
            self.assertNotIn(removed, html)

        self.assertIn('id="workflowBuilderDlg"', html)
        self.assertIn('id="workflowPreFlightDlg"', html)
        self.assertIn("initWorkflowsUi();", script)
        self.assertNotIn(".tool-card--workflow", styles)
        self.assertIn(".workflow-builder-dialog", styles)

    def test_legacy_workflow_defaults_still_normalize_saved_settings(self):
        script = (ROOT / "script.js").read_text(encoding="utf-8")

        for expected in (
            "const DEFAULT_WORKFLOW_CAD_DEFAULTS = {",
            "function normalizeWorkflowCadDefaults(",
            "userSettings.workflowCadDefaults = normalizeWorkflowCadDefaults(",
        ):
            self.assertIn(expected, script)

    def test_workflow_preflight_preserves_dwg_source_metadata(self):
        script = (ROOT / "script.js").read_text(encoding="utf-8")

        for expected in (
            "function normalizeWorkflowDwgSource(",
            "function buildFileWorkflowDwgSources(",
            "inputState.sources = buildFileWorkflowDwgSources(chosen);",
            "entry.dwgFileSources = sources;",
            "stepDefaults.dwgFileSources",
        ):
            self.assertIn(expected, script)

    def test_workflow_launch_context_accepts_snake_case_project_path(self):
        script = (ROOT / "script.js").read_text(encoding="utf-8")

        for expected in (
            "launchContext?.project_path",
            "function getWorkflowMissingToolIds(",
            "Workflow tools are unavailable.",
        ):
            self.assertIn(expected, script)


if __name__ == "__main__":
    unittest.main()
