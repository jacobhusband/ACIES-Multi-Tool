import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"
INDEX_HTML_PATH = REPO_ROOT / "index.html"


class ExpenseAttachmentUiTests(unittest.TestCase):
    def test_expense_attachment_preview_script_wiring_exists(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        self.assertIn("const EXPENSE_IMAGE_THUMB_MAX_SIZE = 320;", text)
        self.assertIn("const EXPENSE_IMAGE_MODAL_MAX_SIZE = 1800;", text)
        self.assertIn("window.pywebview.api.get_expense_image_preview", text)
        self.assertIn("window.pywebview.api.resolve_expense_attachment_path", text)
        self.assertIn("function hydrateExpenseImageThumb(previewButton, image) {", text)
        self.assertIn("function openExpenseImagePreview(image) {", text)
        self.assertIn("function openExpenseImagePreviewDialog({ dataUrl, filename, width, height }) {", text)
        self.assertIn("function closeExpenseImagePreviewDialog() {", text)
        self.assertIn("function resetExpenseImagePreviewTransform() {", text)
        self.assertIn("function applyExpenseImagePreviewZoom(nextScale, clientX, clientY) {", text)
        self.assertIn("function applyExpenseImagePreviewPan(deltaX, deltaY) {", text)
        self.assertIn("function openExpenseAttachment(image) {", text)
        self.assertIn('event.preventDefault();', text)
        self.assertIn('event.stopPropagation();', text)
        self.assertNotIn("function setExpenseImagePreviewDialogState(", text)
        self.assertNotIn("expenseImagePreviewTitle", text)
        self.assertNotIn("expenseImagePreviewOpenBtn", text)
        self.assertNotIn("expenseImagePreviewStatus", text)

    def test_expense_attachment_preview_dialog_markup_exists(self):
        text = INDEX_HTML_PATH.read_text(encoding="utf-8")
        self.assertIn('<dialog id="expenseImagePreviewDlg"', text)
        self.assertIn('id="expenseImagePreviewCloseBtn"', text)
        self.assertIn('id="expenseImagePreviewStage"', text)
        self.assertIn('id="expenseImagePreviewImg"', text)
        self.assertNotIn('id="expenseImagePreviewTitle"', text)
        self.assertNotIn('id="expenseImagePreviewOpenBtn"', text)
        self.assertNotIn('id="expenseImagePreviewStatus"', text)

    def test_expense_attachment_preview_styles_exist(self):
        text = STYLES_CSS_PATH.read_text(encoding="utf-8")
        self.assertIn(".expense-image-preview-btn {", text)
        self.assertIn(".expense-attachment-placeholder {", text)
        self.assertIn(".expense-image-preview-dialog {", text)
        self.assertIn(".expense-image-preview-close {", text)
        self.assertIn(".expense-image-preview-stage {", text)
        self.assertIn(".expense-image-preview-stage.is-zoomed {", text)
        self.assertIn(".expense-image-preview-stage.is-dragging {", text)
        self.assertIn("@keyframes expense-thumb-sheen {", text)


if __name__ == "__main__":
    unittest.main()
