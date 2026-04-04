import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
INDEX_HTML_PATH = REPO_ROOT / "index.html"
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectFiltersUiTests(unittest.TestCase):
    def test_project_filter_markup_uses_custom_dropdowns(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")

        self.assertIn('data-filter-dropdown="timeframe"', html)
        self.assertIn('id="timeframeFilterTrigger"', html)
        self.assertIn('id="timeframeFilterMenu"', html)
        self.assertIn('data-filter-value="future"', html)
        self.assertIn('data-filter-dropdown="status"', html)
        self.assertIn('id="statusFilterTrigger"', html)
        self.assertIn('id="statusFilterMenu"', html)
        self.assertIn('data-filter-value="Pending Review"', html)
        self.assertIn('data-filter-dropdown="deliverables"', html)
        self.assertIn('id="deliverablesFilterTrigger"', html)
        self.assertIn('id="deliverablesFilterMenu"', html)
        self.assertIn('data-filter-value="active"', html)
        self.assertIn("Show only active deliverables", html)
        self.assertIn("Show all incomplete deliverables", html)
        self.assertIn('role="menuitemradio"', html)
        self.assertNotIn('id="timeframeFilterSelect"', html)
        self.assertNotIn('id="statusFilterSelect"', html)
        self.assertNotIn('id="toggleNonPrimaryBtn"', html)
        self.assertNotIn('data-filter-value="primary"', html)

    def test_project_filter_script_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn("function syncProjectsFilterDropdowns() {", script)
        self.assertIn("function setProjectsFilterDropdownState(", script)
        self.assertIn("function moveProjectsFilterOptionFocus(", script)
        self.assertIn("function matchesProjectStatusFilter(deliverable, filter) {", script)
        self.assertIn("function matchesProjectDeliverablesFilter(", script)
        self.assertIn("function getProjectListRenderContext(project) {", script)
        self.assertIn(
            "function buildProjectTimeframeNote(",
            script,
        )
        self.assertIn('document.querySelectorAll(".projects-filter-dropdown").forEach((dropdown) => {', script)
        self.assertIn('const filterKey = dropdown.dataset.filterDropdown;', script)
        self.assertIn('if (filterKey === "timeframe") return dueFilter || "all";', script)
        self.assertIn('if (filterKey === "status") return statusFilter || "all";', script)
        self.assertIn('if (filterKey === "deliverables") return deliverablesFilter || "active";', script)
        self.assertIn('dueFilter = value;', script)
        self.assertIn('statusFilter = value;', script)
        self.assertIn('deliverablesFilter = value;', script)
        self.assertIn("const projectListContextMap = new Map();", script)
        self.assertIn("const visibleDeliverables = projectListContext.visibleDeliverables;", script)
        self.assertIn("Active deliverable is outside this timeframe.", script)
        self.assertIn('statusFilter !== "all" || deliverablesFilter !== "all"', script)
        self.assertIn("No matching projects", script)
        self.assertNotIn("toggleNonPrimaryBtn", script)
        self.assertNotIn('return deliverablesFilter || "primary";', script)

    def test_project_filter_styles_use_custom_dropdown_classes(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        self.assertIn(".projects-filter-dropdown {", css)
        self.assertIn('.projects-filter-dropdown[data-filter-dropdown="deliverables"] {', css)
        self.assertIn(".projects-filter-trigger {", css)
        self.assertIn(".projects-filter-menu {", css)
        self.assertIn(".projects-filter-option {", css)
        self.assertIn(".projects-filter-option.is-selected {", css)
        self.assertIn(".projects-filter-dropdown.open .projects-filter-chevron {", css)
        self.assertIn(".project-timeframe-note {", css)
        self.assertNotIn(".projects-filter-select {", css)


if __name__ == "__main__":
    unittest.main()
