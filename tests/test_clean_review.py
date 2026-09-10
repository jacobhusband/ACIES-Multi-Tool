import tempfile
import unittest
from pathlib import Path

import pymupdf as fitz
from clean_review import compare_sets


class PdfComparisonTests(unittest.TestCase):
    def test_intentional_removal_allowed_but_other_changes_highlighted(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            def fixture(name, stamp, extra=False):
                path = root / name
                with fitz.open() as doc:
                    page = doc.new_page(width=300, height=200)
                    page.insert_text((20, 30), 'KEEP THIS TEXT')
                    if stamp: page.draw_rect(fitz.Rect(200, 130, 260, 180), fill=(0, 0, 1))
                    if extra: page.draw_rect(fitz.Rect(30, 70, 80, 100), fill=(1, 0, 0))
                    page.set_rotation(270)  # AutoCAD-style landscape page encoding.
                    doc.save(path)
                return [{'drawing': 'E01.dwg', 'layout': 'E01', 'file': str(path)}]
            original = fixture('original.pdf', True)
            reference = fixture('reference.pdf', False)
            cleaned = fixture('clean.pdf', False)
            report = compare_sets(original, reference, cleaned, root / 'match')
            self.assertEqual('match_within_tolerance', report['status'])
            self.assertGreater(report['pages'][0]['intentionalRemovalPixels'], 0)
            altered = fixture('changed.pdf', False, True)
            report = compare_sets(original, reference, altered, root / 'review')
            self.assertEqual('review_required', report['status'])
            self.assertGreater(report['pages'][0]['unexpectedPixels'], 0)
            with fitz.open(root / 'review/Comparison.pdf') as document:
                self.assertEqual(1, len(document))
                self.assertIn('review_required', document[0].get_text())
                self.assertTrue(document[0].get_images())
            with self.assertRaisesRegex(RuntimeError, 'layout lists'):
                compare_sets(original, reference, [], root / 'missing')


if __name__ == '__main__':
    unittest.main()
