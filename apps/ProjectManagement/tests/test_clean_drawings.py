import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch

import fitz
import clean_drawings as clean


class CleanDrawingTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name) / 'Project'
        (self.root / 'Electrical').mkdir(parents=True)
        (self.root / 'Xrefs').mkdir()
        self.tb = self.root / 'Xrefs' / 'x-TB.dwg'
        self.sheet = self.root / 'Electrical' / 'E01.dwg'
        self.tb.write_bytes(b'titleblock original')
        self.sheet.write_bytes(b'sheet original')
        self.selection = {'titleblock': 'Xrefs/x-TB.dwg', 'drawings': ['Electrical/E01.dwg'], 'width': 36, 'height': 24}
        self.calls = []
        self.media = []
        self.fail_operation = ''
        self.work = Path(self.temp.name) / 'worker'
        self.work.mkdir()

    def worker(self, acad, dll, drawing, job, workspace):
        op = job['Operation']
        self.calls.append(op)
        if op == self.fail_operation:
            raise RuntimeError('simulated failure')
        if op == 'scan':
            return {'files': [str(self.tb), str(self.sheet)], 'references': [], 'media': self.media}
        if op in ('titleblock', 'sheet'):
            self.assertTrue(clean._inside(drawing, workspace))
            self.assertTrue(clean._inside(job['Output'], workspace))
            Path(job['Output']).write_bytes(b'cleaned')
        return {'modelEntities': 1, 'paperEntities': 1, 'externalReferences': 0}

    def run_job(self, worker=None):
        with patch.object(clean.tempfile, 'mkdtemp', return_value=str(self.work)):
            return clean.run(self.root, self.selection, 'unused', worker=worker or self.worker, dll='test.dll')

    def test_delivery_preserves_originals_and_relative_structure(self):
        result = self.run_job()
        output = Path(result['output'])
        self.assertEqual(b'titleblock original', self.tb.read_bytes())
        self.assertEqual(b'sheet original', self.sheet.read_bytes())
        self.assertEqual(b'cleaned', (output / 'Xrefs' / 'x-TB.dwg').read_bytes())
        self.assertEqual(b'cleaned', (output / 'Electrical' / 'E01.dwg').read_bytes())
        self.assertTrue((output / 'cleanup-report.json').is_file())
        self.assertEqual(2, self.calls.count('verify'))
        self.assertFalse(self.work.exists())

    def test_media_stops_before_prepare_and_delivery(self):
        self.media = ['logo RasterImage']
        with self.assertRaisesRegex(RuntimeError, 'embedding is not supported'):
            self.run_job()
        self.assertEqual(['scan'], self.calls)
        self.assertFalse((self.root / 'Cleaned CAD').exists())
        self.assertTrue(self.work.exists())

    def test_temp_cleanup_failure_does_not_mark_delivered_drawings_failed(self):
        with patch.object(clean.shutil, 'rmtree', side_effect=PermissionError('locked temp file')):
            result = self.run_job()
        self.assertEqual('success', result['status'])
        self.assertTrue(Path(result['output']).is_dir())
        self.assertIn('could not be removed', result['cleanupWarning'])

    def test_sheet_failure_never_delivers_partial_titleblock(self):
        self.fail_operation = 'sheet'
        with self.assertRaisesRegex(RuntimeError, 'simulated failure'):
            self.run_job()
        self.assertFalse((self.root / 'Cleaned CAD').exists())
        self.assertEqual(b'titleblock original', self.tb.read_bytes())

    def test_reopen_count_mismatch_blocks_delivery(self):
        def worker(*args):
            result = self.worker(*args)
            if args[3]['Operation'] == 'verify':
                result['modelEntities'] = 0
            return result
        with self.assertRaisesRegex(RuntimeError, 'counts changed'):
            self.run_job(worker)
        self.assertFalse((self.root / 'Cleaned CAD').exists())

    def test_source_edit_during_processing_blocks_delivery(self):
        def worker(*args):
            result = self.worker(*args)
            if args[3]['Operation'] == 'sheet':
                self.sheet.write_bytes(b'user edit')
            return result
        with self.assertRaisesRegex(RuntimeError, 'Source changed'):
            self.run_job(worker)
        self.assertEqual(b'user edit', self.sheet.read_bytes())
        self.assertFalse((self.root / 'Cleaned CAD').exists())

    def test_path_escape_and_nonfinite_dimensions_rejected(self):
        outside = self.root.parent / 'outside.dwg'
        outside.write_bytes(b'outside')
        self.selection['titleblock'] = '../outside.dwg'
        with self.assertRaisesRegex(RuntimeError, 'inside the project'):
            self.run_job()
        self.selection['titleblock'] = 'Xrefs/x-TB.dwg'
        self.selection['width'] = float('nan')
        with self.assertRaisesRegex(RuntimeError, 'valid sheet dimensions'):
            self.run_job()
        self.assertEqual([], self.calls)

    def test_pdf_orientation_and_mixed_page_sizes(self):
        folder = self.root / 'Electrical' / 'Checkset'
        folder.mkdir()
        with fitz.open() as doc:
            doc.new_page(width=612, height=792)  # Ignore administrative cover.
            doc.new_page(width=36 * 72, height=24 * 72)
            doc.new_page(width=24 * 72, height=36 * 72)
            doc.save(folder / 'set.pdf')
        result = clean.discover(self.root)
        self.assertEqual([(36, 24, 2), (24, 36, 3)],
                         [(s['width'], s['height'], s['page']) for s in result['sizes']])
        self.assertEqual([str(Path('Electrical/E01.dwg'))], result['drawings'])

    def test_worker_reports_progress_while_waiting(self):
        process = Mock()
        process.wait.side_effect = [clean.subprocess.TimeoutExpired('acad', 5), 0]
        def launch(*args, **kwargs):
            request = json.loads(Path(kwargs['env']['ACIES_CLEAN_JOB']).read_text())
            Path(request['ResultPath']).write_text(json.dumps({'success': True, 'details': {}}))
            return process
        messages = []
        with patch.object(clean.subprocess, 'Popen', side_effect=launch):
            clean.run_worker('acad', 'test.dll', self.tb, {'Operation': 'scan'}, self.work, notify=messages.append)
        self.assertTrue(any('reference inspection: running' in message for message in messages))
        process.kill.assert_not_called()

    def test_worker_timeout_still_terminates_process(self):
        process = Mock()
        with patch.object(clean.subprocess, 'Popen', return_value=process), \
                patch.object(clean.time, 'monotonic', side_effect=[0, 301]):
            with self.assertRaisesRegex(RuntimeError, 'timed out'):
                clean.run_worker('acad', 'test.dll', self.tb, {'Operation': 'scan'}, self.work)
        process.kill.assert_called_once()


if __name__ == '__main__':
    unittest.main()
