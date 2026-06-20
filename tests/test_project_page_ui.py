import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"
INDEX_HTML_PATH = REPO_ROOT / "index.html"
MAIN_PY_PATH = REPO_ROOT / "main.py"


class ProjectPageUiTests(unittest.TestCase):
    def test_page_data_model_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        self.assertIn("function normalizePage(raw) {", script)
        self.assertIn("page: normalizePage(project.page),", script)
        self.assertIn("page: normalizePage(deliverable.page),", script)

    def test_page_editor_controller_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        self.assertIn("function openProjectPage(project) {", script)
        self.assertIn("function openDeliverablePage(project, deliverable) {", script)
        self.assertIn("function serializePageHtml(editor) {", script)
        self.assertIn("function hydratePageImages(editor) {", script)
        self.assertIn("function insertPageTodo() {", script)
        self.assertIn("function createPageTodoElement() {", script)
        self.assertIn("function handlePageTodoEnter(editor) {", script)
        self.assertIn("if (handlePageTodoEnter(editor)) e.preventDefault();", script)
        self.assertIn("function normalizePageLinkHref(url) {", script)
        self.assertIn('openExternalUrl(link.getAttribute("href"));', script)
        self.assertIn("function renderPageDeliverablesGrid(project) {", script)
        self.assertIn("function ensurePageViewReady() {", script)
        self.assertIn("window.pywebview.api.save_page_asset", script)
        self.assertIn("window.pywebview.api.get_page_asset", script)

    def test_page_entry_points_exist(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        self.assertIn("if (deliverable) openDeliverablePage(project, deliverable);", script)
        self.assertIn("openProjectPage(project);", script)

    def test_page_view_markup_exists(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")
        self.assertIn('id="pageView"', html)
        self.assertIn('id="pageEditor"', html)
        self.assertIn('id="pageTitle"', html)
        self.assertIn('id="pageBreadcrumb"', html)
        self.assertIn('id="pageDeliverablesGrid"', html)
        self.assertIn('id="pageInsertImageBtn"', html)
        self.assertIn('data-page-action="todo"', html)
        self.assertIn('id="pageImageInput"', html)

    def test_page_view_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")
        self.assertIn(".page-view {", css)
        self.assertIn(".page-editor {", css)
        self.assertIn(".page-toolbar {", css)
        self.assertIn(".page-todo {", css)
        self.assertIn(".page-deliverable-card {", css)

    def test_backend_page_asset_api_exists(self):
        main_py = MAIN_PY_PATH.read_text(encoding="utf-8")
        self.assertIn("def save_page_asset(self, owner_key, data_url, filename=\"\"):", main_py)
        self.assertIn("def get_page_asset(self, asset_path, max_size=1600):", main_py)
        self.assertIn("def _resolve_page_asset_path(self, asset_path):", main_py)


if __name__ == "__main__":
    unittest.main()
