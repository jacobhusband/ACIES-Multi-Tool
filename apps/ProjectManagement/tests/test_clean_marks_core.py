import os
from pathlib import Path
import subprocess
import tempfile
import unittest

import clean_drawings as clean


@unittest.skipUnless(os.environ.get('ACIES_TEST_CORE') and os.environ.get('ACIES_TEST_CLEAN_DLL'), 'Core Console integration is opt-in')
class TitleblockMarkTests(unittest.TestCase):
    def test_named_stamp_removed_logo_preserved_and_original_unchanged(self):
        with tempfile.TemporaryDirectory(prefix='acies-marks-test-') as temporary:
            root = Path(temporary)
            drawing = root / 'titleblock.dwg'
            output = root / 'unmarked.dwg'
            script = root / 'fixture.scr'
            script.write_text(f'''(setvar "FILEDIA" 0)
(command "_.LINE" "1,1" "2,2" "")
(command "_.-BLOCK" "EngineerStamp" "0,0" (entlast) "")
(command "_.-INSERT" "EngineerStamp" "0,0" 1 1 0)
(command "_.LINE" "10,10" "11,11" "")
(command "_.-BLOCK" "CompanyLogo" "0,0" (entlast) "")
(command "_.-INSERT" "CompanyLogo" "0,0" 1 1 0)
(command "_.QSAVE" "{drawing.as_posix()}")
_.QUIT
_Y
''')
            result = subprocess.run([os.environ['ACIES_TEST_CORE'], '/s', str(script)], cwd=root, capture_output=True, timeout=90,
                                    creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
            self.assertEqual(0, result.returncode)
            original = drawing.read_bytes()
            removed = clean.run_worker(os.environ['ACIES_TEST_CORE'], os.environ['ACIES_TEST_CLEAN_DLL'], drawing,
                                       {'Operation': 'remove-titleblock-marks', 'Output': str(output)}, root)
            self.assertEqual(1, len(removed['removed']))
            self.assertIn('EngineerStamp', removed['removed'][0]['source'])
            self.assertEqual(original, drawing.read_bytes())
            counts = clean.run_worker(os.environ['ACIES_TEST_CORE'], os.environ['ACIES_TEST_CLEAN_DLL'], output,
                                      {'Operation': 'verify'}, root)
            self.assertEqual(1, counts['modelEntities'])


if __name__ == '__main__':
    unittest.main()
