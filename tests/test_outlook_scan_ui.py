import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
INDEX_HTML_PATH = REPO_ROOT / "index.html"
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class OutlookScanUiTests(unittest.TestCase):



    def test_existing_project_without_deliverable_routes_to_project_important_page(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn('id="aiNoMatchAddHelp"', html)
        self.assertIn("function hasAiSeparateDeliverable(rawAiData = {}) {", script)
        self.assertIn("function buildAiImportantText(rawAiData = {}) {", script)
        self.assertIn("function appendAiImportantToProject(project, rawAiData = {}) {", script)
        self.assertIn('const block = `<p data-important="true">${escapeHtml(text)}</p>`;', script)
        self.assertIn("function addAiImportantToProject(projectIndex, aiProject, rawAiData) {", script)
        self.assertIn("function addAiResultToProject(projectIndex, aiProject, rawAiData) {", script)
        self.assertIn("if (hasAiSeparateDeliverable(rawAiData)) {", script)
        self.assertIn("return addAiImportantToProject(projectIndex, aiProject, rawAiData);", script)
        self.assertIn("pendingImportantPageScroll = true;", script)
        self.assertIn("openProjectPage(target, appended.subpage);", script)
        self.assertIn("add the email content to its page as /important", script)



if __name__ == "__main__":
    unittest.main()
