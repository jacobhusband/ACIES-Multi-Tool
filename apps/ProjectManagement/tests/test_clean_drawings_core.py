"""Opt-in live test: ACIES_TEST_CORE and ACIES_TEST_CLEAN_DLL must point to installed binaries."""
import json
import os
from pathlib import Path
import subprocess
import tempfile
import unittest
from PIL import Image

import clean_drawings as clean


@unittest.skipUnless(os.environ.get('ACIES_TEST_CORE') and os.environ.get('ACIES_TEST_CLEAN_DLL'), 'Core Console integration is opt-in')
class CoreCleanDrawingTests(unittest.TestCase):
    def test_external_raster_image_still_blocks_cleanup(self):
        with tempfile.TemporaryDirectory(prefix='acies-clean-raster-') as temporary:
            root = Path(temporary)
            image = root / 'external.png'
            Image.new('RGB', (4, 4), 'white').save(image)
            drawing = root / 'image.dwg'
            script = root / 'make-image.scr'
            script.write_text(f'''(setvar "FILEDIA" 0)
(command "_.-IMAGE" "_ATTACH" "{image.as_posix()}" "0,0" 1 0)
(command "_.QSAVE" "{drawing.as_posix()}")
_.QUIT
_Y
''', encoding='utf-8')
            result = subprocess.run([os.environ['ACIES_TEST_CORE'], '/s', str(script)],
                                    cwd=root, capture_output=True, timeout=60,
                                    creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
            self.assertEqual(0, result.returncode)
            scan = clean.run_worker(os.environ['ACIES_TEST_CORE'], os.environ['ACIES_TEST_CLEAN_DLL'], drawing,
                                    {'Operation': 'scan', 'Files': [str(drawing)], 'Catalog': []}, root)
            self.assertEqual(1, len(scan['media']))
            self.assertIn('RasterImage', scan['media'][0])
            output = root / 'must-not-exist.dwg'
            with self.assertRaisesRegex(RuntimeError, 'Unsupported media'):
                clean.run_worker(os.environ['ACIES_TEST_CORE'], os.environ['ACIES_TEST_CLEAN_DLL'], drawing,
                                 {'Operation': 'titleblock', 'Width': 36, 'Height': 24, 'Output': str(output)}, root)
            self.assertFalse(output.exists())

    def test_fixed_boundary_sheet_binding_and_reopen(self):
        with tempfile.TemporaryDirectory(prefix='acies-clean-integration-') as temporary:
            root = Path(temporary)
            (root / 'Xrefs').mkdir()
            (root / 'Electrical').mkdir()
            tb = (root / 'Xrefs' / 'x-TB.dwg').as_posix()
            sheet = (root / 'Electrical' / 'E01.dwg').as_posix()
            scripts = [f'''(setvar "FILEDIA" 0)
(command "_.RECTANG" "0,0" "36,24")
(command "_.WIPEOUT" "2,2" "4,2" "4,4" "2,4" "")
(command "_.LINE" "-10,-10" "-5,-5" "")
(command "_.LINE" "100,100" "110,110" "")
(command "_.QSAVE" "{tb}")
_.QUIT
_Y
''', f'''(setvar "FILEDIA" 0)
(command "_.LINE" "1,1" "5,5" "")
(setvar "TILEMODE" 0)
(command "_.-XREF" "_ATTACH" "{tb}" "0,0" 1 1 0)
(command "_.WIPEOUT" "8,8" "9,8" "9,9" "8,9" "")
(command "_.QSAVE" "{sheet}")
_.QUIT
_Y
''']
            for index, text in enumerate(scripts):
                script = root / f'make-{index}.scr'
                script.write_text(text, encoding='utf-8')
                result = subprocess.run([os.environ['ACIES_TEST_CORE'], '/s', str(script)],
                                        cwd=root, capture_output=True, timeout=60,
                                        creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
                self.assertEqual(0, result.returncode)
            original = {p: p.read_bytes() for p in (Path(tb), Path(sheet))}
            scan = clean.run_worker(os.environ['ACIES_TEST_CORE'], os.environ['ACIES_TEST_CLEAN_DLL'], sheet,
                                    {'Operation': 'scan', 'Files': [sheet], 'Catalog': [tb]}, root)
            self.assertEqual([str(Path(tb))], scan['titleblocks'])
            self.assertEqual([], scan['media'], 'Wipeouts must not be classified as external media.')
            result = clean.run(root, {'titleblock': 'Xrefs/x-TB.dwg', 'drawings': ['Electrical/E01.dwg'], 'width': 36, 'height': 24},
                               os.environ['ACIES_TEST_CORE'], dll=Path(os.environ['ACIES_TEST_CLEAN_DLL']))
            report = json.loads((Path(result['output']) / 'cleanup-report.json').read_text())
            self.assertEqual(2, report['titleblockValidation']['modelEntities'])
            self.assertEqual(1, report['titleblockValidation']['wipeouts'])
            self.assertEqual(2, report['drawings'][0]['validation']['wipeouts'])
            self.assertEqual(0, report['drawings'][0]['validation']['externalReferences'])
            for path, data in original.items():
                self.assertEqual(data, path.read_bytes())
            rejected_output = root / 'must-not-exist.dwg'
            with self.assertRaisesRegex(RuntimeError, 'did not complete'):
                clean.run_worker(os.environ['ACIES_TEST_CORE'], os.environ['ACIES_TEST_CLEAN_DLL'], sheet,
                                 {'Operation': 'sheet', 'Titleblock': str(root / 'wrong.dwg'),
                                  'Width': 36, 'Height': 24, 'Output': str(rejected_output)}, root)
            self.assertFalse(rejected_output.exists())


if __name__ == '__main__':
    unittest.main()
