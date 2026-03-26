import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
INDEX_HTML_PATH = REPO_ROOT / "index.html"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class DeliverableLinkAndProjectPathUiTests(unittest.TestCase):
    def test_project_path_root_normalization_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn(
            'const PROJECT_ROOT_SEGMENT_REGEX =\n  /^(\\d{6})(?!\\d)(?:\\s*(?:[-_]\\s*)?(.*))?$/;',
            script,
        )
        self.assertIn("function normalizeWindowsPath(rawPath) {", script)
        self.assertIn("function findProjectRootPath(rawPath) {", script)
        self.assertIn("function normalizeProjectPath(rawPath) {", script)
        self.assertIn('path: normalizeProjectPath(project.path || ""),', script)
        self.assertIn('path: normalizeProjectPath(val("f_path")),', script)
        self.assertIn(
            'if (project?.path !== normalizeWindowsPath(item?.path || "")) {',
            script,
        )
        self.assertIn(
            "function normalizeProjectPathInput(pathInput, { forceProjectFields = false } = {}) {",
            script,
        )
        self.assertIn(
            "normalizeProjectPathInput(pathInput, { forceProjectFields: true })",
            script,
        )
        self.assertIn("normalizeProjectPathInput(pathInput);", script)
        self.assertIn(
            "return findWorkroomProjectRootById(projectPath) || projectPath;", script
        )

    def test_deliverable_links_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn("function normalizeDeliverableLinkEntry(entry) {", script)
        self.assertIn("function hasMeaningfulDeliverableLinkLabel(link) {", script)
        self.assertIn("function normalizeDeliverableLinks(links, legacyLinkPath = \"\") {", script)
        self.assertIn("links: normalizeDeliverableLinks(project.links),", script)
        self.assertIn(
            'links: normalizeDeliverableLinks(deliverable.links, deliverable.linkPath || ""),',
            script,
        )
        self.assertIn(
            'links: normalizeDeliverableLinks(seed.links, seed.linkPath || ""),',
            script,
        )
        self.assertIn("base.links = normalizeDeliverableLinks([...(base.links || []), ...(incoming.links || [])]);", script)
        self.assertIn("links: [],", script)
        self.assertIn("function getDeliverableCardLinks(card) {", script)
        self.assertIn("function setDeliverableCardLinks(card, links) {", script)
        self.assertIn("function getModalProjectDraft() {", script)
        self.assertIn("function setModalProjectDraft(project) {", script)
        self.assertIn("function getModalProjectLinks() {", script)
        self.assertIn("function setModalProjectLinks(links) {", script)
        self.assertIn("function ensureDeliverableLinksPanel() {", script)
        self.assertIn("let deliverableLinksPanelClosePending = false;", script)
        self.assertIn("let deliverableLinksPanelActiveHost = null;", script)
        self.assertIn("let deliverableLinksPanelDetachOutsideListeners = null;", script)
        self.assertIn("let deliverableLinksPanelFocusLossFrame = 0;", script)
        self.assertIn("function getDeliverableLinksPanelOutsideTargets(trigger) {", script)
        self.assertIn("function isDeliverableLinksInteractionTarget(target) {", script)
        self.assertIn("function detachDeliverableLinksOutsideListeners() {", script)
        self.assertIn("function scheduleDeliverableLinksPanelFocusLossCheck() {", script)
        self.assertIn("function attachDeliverableLinksOutsideListeners(trigger) {", script)
        self.assertIn("function openDeliverableLinksPanel(context) {", script)
        self.assertIn("async function closeDeliverableLinksPanel({ focusTrigger = false } = {}) {", script)
        self.assertIn("async function requestDeliverableLinksPanelClose(options = {}) {", script)
        self.assertIn("function getOpenDeliverableLinksPanelEntries() {", script)
        self.assertIn("async function updateOpenDeliverableLinksScopes({", script)
        self.assertIn("async function moveOpenDeliverableLinkEntryScope(entry, nextScope) {", script)
        self.assertIn("async function removeOpenDeliverableLinkEntry(entry) {", script)
        self.assertIn("function createDeliverableLinksControl(deliverable, options = {}) {", script)
        self.assertIn('textContent: "Links",', script)
        self.assertIn('placeholder: "Name (optional)",', script)
        self.assertIn('placeholder: "Paste a file path or URL...",', script)
        self.assertIn('className: "deliverable-links-add-input deliverable-links-name-input"', script)
        self.assertIn('className: "deliverable-links-add-input deliverable-links-link-input"', script)
        self.assertIn('className: "deliverable-links-scope-checkbox"', script)
        self.assertIn('textContent: "Project-wide"', script)
        self.assertIn("function getPendingDeliverableLinkEntry() {", script)
        self.assertIn('const raw = String(elements.linkInput.value || "")', script)
        self.assertIn('label: String(elements.nameInput.value || "").trim(),', script)
        self.assertIn("function clearPendingDeliverableLinkEntry() {", script)
        self.assertIn("elements.nameInput.value = \"\";", script)
        self.assertIn("elements.linkInput.value = \"\";", script)
        self.assertIn("function focusDeliverableLinksPrimaryInput({ select = true } = {}) {", script)
        self.assertIn("elements.nameInput.focus();", script)
        self.assertIn("const visibleLabel =", script)
        self.assertIn(
            'String(entry.label || entry.raw || entry.url || "Link").trim() || "Link";',
            script,
        )
        self.assertIn('"aria-label": `Open ${visibleLabel}`', script)
        self.assertIn('"aria-label": `Remove ${visibleLabel}`', script)
        self.assertIn('removeBtn.textContent = "Remove";', script)
        self.assertIn('className: "deliverable-link-scope-checkbox"', script)
        self.assertIn('checked: entry.scope === "project"', script)
        self.assertIn(
            'projectLinks.push(nextEntry);',
            script,
        )
        self.assertIn(
            'deliverableLinks.push(nextEntry);',
            script,
        )
        self.assertIn(
            'const links = getOpenDeliverableLinksPanelEntries();',
            script,
        )
        self.assertNotIn("deliverable-link-checkbox", script)
        self.assertNotIn("nextLinks[index].done", script)
        self.assertNotIn("Mark ${visibleLabel} complete", script)
        self.assertIn("function commitOpenDeliverableLinksInput({ refocus = false } = {}) {", script)
        self.assertIn("function openDeliverableLinkEntry(linkEntry) {", script)
        self.assertIn("await window.pywebview.api.reveal_path(localPath);", script)
        self.assertIn("openExternalUrl(url);", script)
        self.assertIn('target.addEventListener("pointerdown", handleOutsidePointerLike, true);', script)
        self.assertIn('target.addEventListener("mousedown", handleOutsidePointerLike, true);', script)
        self.assertIn('ensureDeliverableLinksPanel().panel.addEventListener("focusout", handleFocusOut);', script)
        self.assertIn('trigger.addEventListener("focusout", handleFocusOut);', script)
        self.assertIn('ownerDocument.addEventListener("focusin", handleOutsideFocusIn, true);', script)
        self.assertIn('ownerDocument.addEventListener("scroll", handleViewportScroll, true);', script)
        self.assertIn('ownerDocument.defaultView?.addEventListener("resize", handleViewportResize);', script)
        self.assertIn("const activeElement = ownerDocument?.activeElement || document.activeElement;", script)
        self.assertIn("deliverableLinksPanelActiveHost instanceof HTMLDialogElement &&", script)
        self.assertIn("e.target === deliverableLinksPanelActiveHost", script)
        self.assertIn("if (getPendingDeliverableLinkEntry()) {", script)
        self.assertIn("await commitOpenDeliverableLinksInput();", script)
        self.assertIn("attachDeliverableLinksOutsideListeners(context.trigger);", script)
        self.assertIn("detachDeliverableLinksOutsideListeners();", script)
        self.assertIn("focusDeliverableLinksPrimaryInput();", script)
        self.assertIn("openDeliverableLinksContext.addScope = \"deliverable\";", script)
        self.assertIn("projectDraft: p", script)
        self.assertIn("projectDraft: getModalProjectDraft()", script)
        self.assertIn("links: getModalProjectLinks(),", script)
        self.assertIn("project,", script)
        self.assertIn("const links = getDeliverableCardLinks(card);", script)
        self.assertIn("links,", script)

    def test_deliverable_links_markup_and_styles_exist(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        self.assertIn('<div class="deliverable-link-control"></div>', html)
        self.assertIn(".deliverable-link-control {", css)
        self.assertIn(".deliverable-link-trigger {", css)
        self.assertIn(".deliverable-links-panel {", css)
        self.assertIn(".deliverable-links-panel[hidden] {", css)
        self.assertIn(".deliverable-links-list {", css)
        self.assertIn(".deliverable-link-item {", css)
        self.assertIn("grid-template-columns: minmax(0, 1fr) auto;", css)
        self.assertIn(".deliverable-link-open {", css)
        self.assertIn("font: inherit;", css)
        self.assertIn(".deliverable-link-actions {", css)
        self.assertIn(".deliverable-link-scope-control,", css)
        self.assertIn(".deliverable-link-scope-checkbox,", css)
        self.assertIn(".deliverable-link-remove {", css)
        self.assertIn('text-transform: uppercase;', css)
        self.assertNotIn(".deliverable-link-item:hover .deliverable-link-remove {", css)
        self.assertNotIn(".deliverable-link-checkbox {", css)
        self.assertNotIn(".deliverable-link-item.done .deliverable-link-label,", css)
        self.assertIn(".deliverable-links-add-bullet {", css)
        self.assertIn(".deliverable-links-add-fields {", css)
        self.assertIn(".deliverable-links-scope-control {", css)
        self.assertIn(".deliverable-links-scope-checkbox {", css)
        self.assertIn(".deliverable-links-add-input {", css)
        self.assertIn(".deliverable-links-name-input {", css)
        self.assertIn(".deliverable-links-link-input {", css)


if __name__ == "__main__":
    unittest.main()
