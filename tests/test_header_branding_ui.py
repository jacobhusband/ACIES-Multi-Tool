import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
INDEX_HTML_PATH = REPO_ROOT / "index.html"
SPEC_PATH = REPO_ROOT / "ACIES Scheduler.spec"


class HeaderBrandingUiTests(unittest.TestCase):
    def test_ui_uses_supplied_logo_on_both_surfaces(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")

        self.assertEqual(2, html.count('src="assets/acies-modern-logo.png"'))
        self.assertTrue((REPO_ROOT / 'assets' / 'acies-modern-logo.png').is_file())

    def test_header_logo_has_accessible_name(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")

        header_start = html.index('<div class="header-start">')
        header_end = html.index('<div class="header-center">')
        brand = html[header_start:header_end]

        self.assertIn('class="brand-mark supplied-logo"', brand)
        self.assertIn('aria-label="ACIES Engineering"', brand)
        self.assertIn('src="assets/acies-modern-logo.png"', brand)

    def test_loader_mark_is_decorative(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")

        loader_start = html.index('id="appLoader"')
        loader_end = html.index('<header class="app-header">')
        loader = html[loader_start:loader_end]

        # The loader already announces itself as "Loading"; the mark inside it
        # must not add a second announcement.
        self.assertIn('class="brand-mark app-loader-mark supplied-logo"', loader)
        self.assertIn('aria-hidden="true"', loader)
        self.assertNotIn("aria-label=\"ACIES Engineering\"", loader)

    def test_version_chip_is_in_header_actions_before_update_button(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")

        header_start_index = html.index('<div class="header-start">')
        header_actions_index = html.index('<div class="header-actions">')
        version_chip_index = html.index('id="versionChip"')
        app_update_index = html.index('id="appUpdateBtn"')

        header_start_section = html[header_start_index:header_actions_index]
        header_actions_section = html[header_actions_index:app_update_index]

        self.assertGreater(version_chip_index, header_actions_index)
        self.assertLess(version_chip_index, app_update_index)
        self.assertNotIn('id="versionChip"', header_start_section)
        self.assertIn('id="versionChip"', header_actions_section)

    def test_spec_packages_logo_inside_assets_folder(self):
        spec = SPEC_PATH.read_text(encoding="utf-8")

        self.assertIn("('assets\\\\acies.png', 'assets')", spec)


if __name__ == "__main__":
    unittest.main()
