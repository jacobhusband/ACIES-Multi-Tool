import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectCardScrollUiTests(unittest.TestCase):
    def test_card_view_horizontal_scroll_is_free_scrolling(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")
        view_start = css.index(".projects-card-view {")
        view_end = css.index(".projects-card-view[hidden]", view_start)
        view_block = css[view_start:view_end]
        column_start = css.index(".kanban-column {")
        column_end = css.index(".kanban-card-project-meta", column_start)
        column_block = css[column_start:column_end]

        self.assertIn("overflow-x: auto;", view_block)
        self.assertNotIn("scroll-snap-type", view_block)
        self.assertNotIn("scroll-snap-align", column_block)


if __name__ == "__main__":
    unittest.main()
