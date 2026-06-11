import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectDeliverableNoteDueDateUiTests(unittest.TestCase):
    @staticmethod
    def _block(text, start_marker, end_marker):
        start = text.index(start_marker)
        end = text.index(end_marker, start)
        return text[start:end]

    def test_note_data_model_includes_due_date(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        for expected in (
            "function normalizeDeliverableNoteDueDate(",
            'dueDate: normalizeDeliverableNoteDueDate(noteItem?.dueDate)',
            "dueDate: normalizeDeliverableNoteDueDate(out.dueDate)",
        ):
            self.assertIn(expected, script)

        normalize_block = self._block(
            script,
            "function normalizeDeliverableNoteDueDate(",
            "function normalizeDeliverableNoteItems(",
        )
        self.assertIn("parseDueStr(raw)", normalize_block)
        self.assertIn('return ""', normalize_block)

    def test_note_sync_preserves_due_date(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        sync_block = self._block(
            script,
            "function syncDeliverableWorkItemFields(",
            "function getPinnedFirstEntries(",
        )
        self.assertIn("noteItem.dueDate = normalizedNoteItem.dueDate;", sync_block)

    def test_new_notes_seed_empty_due_date(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        modal_commit_block = self._block(
            script,
            "function commitPendingModalDeliverableNote(",
            "function flushPendingModalDeliverableNotes(",
        )
        self.assertIn('dueDate: ""', modal_commit_block)

        card_block = self._block(
            script,
            "function createNotesSection(",
            "function createTasksPreview(",
        )
        self.assertIn('dueDate: ""', card_block)

    def test_note_sort_orders_pinned_then_due_date_then_index(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        sort_block = self._block(
            script,
            "function getDeliverableNoteSortEntries(",
            "function getPinnedDeliverablePreviewItems(",
        )
        self.assertIn("normalizeDeliverableNoteItem(item)", sort_block)
        self.assertIn("parseDueStr(left.item?.dueDate)", sort_block)
        self.assertIn("parseDueStr(right.item?.dueDate)", sort_block)
        self.assertIn("return leftDue - rightDue;", sort_block)
        self.assertIn("return left.index - right.index;", sort_block)
        # Pinned grouping wins over due dates.
        self.assertLess(
            sort_block.index("left.item?.pinned"),
            sort_block.index("parseDueStr(left.item?.dueDate)"),
        )
        # Dated notes sort ahead of dateless notes within a pin group.
        self.assertIn("return leftDue ? -1 : 1;", sort_block)

    def test_due_date_control_wiring(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        control_block = self._block(
            script,
            "function createNoteDueDateControl(",
            "function renderModalDeliverableNoteList(",
        )
        for expected in (
            'className: "note-due-wrap"',
            'className: "note-due-input"',
            'placeholder: "MM/DD/YYYY"',
            "autoFormatDateInput(input)",
            "showCalendarForInput(input",
            "dueState(currentValue)",
            "humanDate(currentValue)",
            'note-due-badge ${ds === "overdue" ? "overdue" : ds === "dueSoon" ? "due-soon" : "ok"}',
            'className: "note-due-add-btn"',
            "normalizeDeliverableNoteDueDate(input.value)",
        ):
            self.assertIn(expected, control_block)

    def test_modal_note_list_uses_due_date_sort_and_control(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        modal_block = self._block(
            script,
            "function renderModalDeliverableNoteList(",
            "function commitPendingModalDeliverableNote(",
        )
        self.assertIn("getDeliverableNoteSortEntries(noteItems)", modal_block)
        self.assertNotIn("getPinnedFirstEntries(", modal_block)
        self.assertIn("createNoteDueDateControl({", modal_block)
        self.assertIn("allowAdd: true", modal_block)
        self.assertIn("liveNote.dueDate = nextDueDate;", modal_block)
        self.assertIn("actions.append(dueControl, attachments, deleteBtn);", modal_block)

    def test_card_notes_use_due_date_sort_and_badge(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        card_block = self._block(
            script,
            "function createNotesSection(",
            "function createTasksPreview(",
        )
        self.assertIn("getDeliverableNoteSortEntries(noteItems)", card_block)
        self.assertIn("createNoteDueDateControl({", card_block)
        self.assertIn("allowAdd: true", card_block)
        self.assertIn("liveNote.dueDate = nextDueDate;", card_block)
        self.assertIn("await save();", card_block)
        self.assertIn("row.append(content, dueControl, deleteBtn);", card_block)

    def test_modal_save_preserves_note_due_date(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        read_form_block = self._block(
            script,
            "function readForm(",
            "function addRefRowFrom(",
        )
        self.assertIn("dueDate: normalizedNoteItem.dueDate,", read_form_block)

    def test_note_due_date_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")
        for expected in (
            ".note-due-wrap {",
            ".note-due-badge {",
            ".note-due-badge.overdue {",
            ".note-due-badge.due-soon {",
            ".note-due-badge.ok {",
            ".note-due-input {",
            ".note-due-add-btn {",
        ):
            self.assertIn(expected, css)
        overdue_block = self._block(css, ".note-due-badge.overdue {", "}")
        self.assertIn("var(--danger)", overdue_block)
        due_soon_block = self._block(css, ".note-due-badge.due-soon {", "}")
        self.assertIn("var(--warning)", due_soon_block)


if __name__ == "__main__":
    unittest.main()
