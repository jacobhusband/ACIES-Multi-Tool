"""Live PDF-to-CAD conversion, using only disposable source fixtures."""
import os
from pathlib import Path
import subprocess
import tempfile
import unittest

import pymupdf
import clean_drawings as clean


@unittest.skipUnless(os.environ.get('ACIES_TEST_CORE') and os.environ.get('ACIES_TEST_CLEAN_DLL'), 'Core Console integration is opt-in')
class HeadlessMediaTests(unittest.TestCase):
    def test_pdf_page_placement_and_source_free_plot(self):
        with tempfile.TemporaryDirectory(prefix='acies-pdf-test-') as temporary:
            root = Path(temporary)
            pdf, drawing, output = root / 'form.pdf', root / 'input.dwg', root / 'embedded.dwg'
            with pymupdf.open() as document:
                for label in ('WRONG PAGE', 'ENERGY FORM PAGE TWO'):
                    page = document.new_page(width=612, height=792)
                    page.insert_text((72, 72), label, fontsize=18)
                    page.draw_rect(pymupdf.Rect(72, 100, 200, 200), color=(1, 0, 0))
                document.save(pdf)
            def core(text, input_dwg=None):
                script = root / 'run.scr'
                script.write_text(text + '\n_.QUIT\n_Y\n')
                args = [os.environ['ACIES_TEST_CORE']]
                if input_dwg: args += ['/i', str(input_dwg)]
                result = subprocess.run(args + ['/s', str(script)], cwd=root, capture_output=True, timeout=90,
                                        env=dict(os.environ, ACIES_EMBED_PROBE_DIR=str(root)),
                                        creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
                self.assertEqual(0, result.returncode, result.stdout.decode('utf-16-le', errors='replace')[-3000:])
            def worker(path, **job):
                return clean.run_worker(os.environ['ACIES_TEST_CORE'], os.environ['ACIES_TEST_CLEAN_DLL'], path, job, root)
            framework = Path(os.environ['ACIES_TEST_CLEAN_DLL']).parent.name
            probe = Path(__file__).resolve().parents[2] / 'ElectricalCommands/tests/HeadlessEmbeddingProbe/bin/Release' / framework / 'Acies.HeadlessEmbeddingProbe.dll'
            self.assertTrue(probe.is_file(), 'Build the HeadlessEmbeddingProbe test project first.')
            core(f'''(setq savedSecure (getvar "SECURELOAD"))
(setvar "SECURELOAD" 0)
(command "_.NETLOAD" "{probe.as_posix()}")
(setvar "SECURELOAD" savedSecure)
ACIESPDFFIXTURE''')
            original = drawing.read_bytes()
            result = worker(drawing, Operation='embed-media', Output=str(output))
            self.assertEqual(0, result['externalMedia'])
            self.assertEqual(original, drawing.read_bytes())
            pdf.unlink()
            self.assertEqual(result, worker(output, Operation='verify-media'))
            plot = root / 'native.pdf'
            core(f'''(setvar "FILEDIA" 0)
(command "_.-PLOT" "_Y" "Model" "DWG To PDF.pc3" "ANSI full bleed A (8.50 x 11.00 Inches)" "_I" "_P" "_N" "_E" "_F" "_C" "_Y" "." "_Y" "_A" "{plot.as_posix()}" "_N" "_Y")''', output)
            with pymupdf.open(plot) as document:
                text = document[0].get_text()
                self.assertIn('ENERGY FORM PAGE TWO', text)
                self.assertNotIn('WRONG PAGE', text)
            with self.assertRaisesRegex(RuntimeError, 'new DWG path'):
                worker(output, Operation='embed-media', Output=str(output))


if __name__ == '__main__':
    unittest.main()
