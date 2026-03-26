import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
INDEX_HTML_PATH = REPO_ROOT / "index.html"


class PanelScheduleAiMultiPhotoUiTests(unittest.TestCase):
    def test_panel_schedule_markup_mentions_multi_photo_guidance(self):
        text = INDEX_HTML_PATH.read_text(encoding="utf-8")

        self.assertIn(
            "Upload one or more breaker photos and one or more directory photos",
            text,
        )
        self.assertIn(
            "upper half, middle, and bottom half photos.",
            text,
        )
        self.assertIn(
            "circuits 1-42 on one image and 43-84 on",
            text,
        )
        self.assertIn("Drag &amp; drop photos or click to select", text)
        self.assertIn("No photos", text)

    def test_panel_schedule_script_uses_array_backed_photo_state_and_plural_payloads(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn("breakerPaths: [],", text)
        self.assertIn("directoryPaths: [],", text)
        self.assertIn("breakerFiles: [],", text)
        self.assertIn("directoryFiles: [],", text)
        self.assertIn("function setCircuitBreakerFiles(kind, files) {", text)
        self.assertIn("function setCircuitBreakerPaths(kind, paths) {", text)
        self.assertIn("allow_multiple: true,", text)
        self.assertIn("input.multiple = true;", text)
        self.assertIn("const files = Array.from(e.dataTransfer?.files || []);", text)
        self.assertIn("if (files.length) setCircuitBreakerFiles(kind, files);", text)
        self.assertIn(
            "breakerPaths: [...normalizeCircuitBreakerPaths(panel.breakerPaths)],",
            text,
        )
        self.assertIn(
            "directoryPaths: [...normalizeCircuitBreakerPaths(panel.directoryPaths)],",
            text,
        )
        self.assertIn("breakerPaths: firstPanel.breakerPaths || [],", text)
        self.assertIn("directoryPaths: firstPanel.directoryPaths || [],", text)
        self.assertIn("function filesToUploadPayloads(files) {", text)


if __name__ == "__main__":
    unittest.main()
