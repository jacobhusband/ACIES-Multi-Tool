import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
INDEX_HTML_PATH = REPO_ROOT / "index.html"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class PageTaggedItemsUiTests(unittest.TestCase):
    """Lines in a page can be tagged as action or coordination items."""

    @staticmethod
    def _block(text: str, start_marker: str, end_marker: str) -> str:
        start = text.index(start_marker)
        end = text.index(end_marker, start)
        return text[start:end]

    def test_tagged_item_helpers_exist(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        self.assertIn("const PAGE_ITEM_TAGS = {", script)
        self.assertIn("function normalizePageItemTag(tag) {", script)
        self.assertIn('function createPageItemElement(tag = "action", id = "") {', script)
        self.assertIn("function setPageItemTag(item, tag) {", script)
        self.assertIn("function insertPageItem(tag) {", script)
        self.assertIn("function togglePageItem(box) {", script)
        self.assertIn("function handlePageItemEnter(editor) {", script)
        self.assertIn("function clearPageItemTag(editor) {", script)
        self.assertIn("function upgradeLegacyPageTodos(editor) {", script)

    def test_tagged_item_dom_structure(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        create_block = self._block(
            script,
            'function createPageItemElement(tag = "action", id = "") {',
            "function setPageItemTag(item, tag) {",
        )
        self.assertIn('item.className = "page-item";', create_block)
        self.assertIn("item.dataset.tag = itemTag;", create_block)
        self.assertIn('item.dataset.itemId = id || createId("pi");', create_block)
        self.assertIn('item.dataset.checked = "false";', create_block)
        self.assertIn('box.className = "page-item-box";', create_block)
        self.assertIn('badge.className = "page-item-badge";', create_block)
        self.assertIn('text.className = "page-item-text";', create_block)

    def test_serialize_backfills_item_ids(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        serialize_block = self._block(
            script,
            "function serializePageHtml(editor) {",
            "function queuePageSave() {",
        )
        self.assertIn('editor.querySelectorAll(".page-item").forEach', serialize_block)
        self.assertIn('item.dataset.itemId = createId("pi");', serialize_block)

    def test_render_upgrades_legacy_todos(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        self.assertIn("upgradeLegacyPageTodos(editor);", script)

    def test_toolbar_wiring_for_tags(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        self.assertIn(
            'else if (btn.dataset.pageAction === "action") insertPageItem("action");',
            script,
        )
        self.assertIn(
            'else if (btn.dataset.pageAction === "coordination") insertPageItem("coordination");',
            script,
        )
        self.assertIn(
            'else if (btn.dataset.pageAction === "untag") clearPageItemTag(getPageEditorEl());',
            script,
        )
        # The click handler toggles the new item checkbox.
        self.assertIn('e.target?.closest?.(".page-item-box")', script)
        self.assertIn("togglePageItem(box);", script)

    def test_toolbar_markup_for_tags(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")
        self.assertIn('data-page-action="action"', html)
        self.assertIn('data-page-action="coordination"', html)
        self.assertIn('data-page-action="untag"', html)
        # The legacy single To-do button is gone.
        self.assertNotIn('data-page-action="todo"', html)

    def test_tagged_item_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")
        self.assertIn(".page-item {", css)
        self.assertIn(".page-item-box {", css)
        self.assertIn(".page-item-badge {", css)
        self.assertIn('.page-item-badge[data-tag="action"] {', css)
        self.assertIn('.page-item-badge[data-tag="coordination"] {', css)
        self.assertIn(".page-item-text {", css)
        self.assertIn(".page-item.done .page-item-text {", css)
        # The legacy to-do selectors are gone.
        self.assertNotIn(".page-todo {", css)


if __name__ == "__main__":
    unittest.main()
