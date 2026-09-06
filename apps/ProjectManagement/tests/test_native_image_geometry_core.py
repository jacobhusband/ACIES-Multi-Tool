"""Opt-in real Core Console tests for the experimental native image encoding."""
import json
import os
from pathlib import Path
import subprocess
import tempfile
import unittest

from PIL import Image, ImageDraw
import clean_drawings as clean


@unittest.skipUnless(os.environ.get('ACIES_TEST_CORE') and os.environ.get('ACIES_TEST_CLEAN_DLL'), 'Core Console integration is opt-in')
class NativeImageGeometryTests(unittest.TestCase):
    def core(self, root, text, drawing=None):
        script = root / 'fixture.scr'
        script.write_text(text + '\n_.QUIT\n_Y\n', encoding='utf-8')
        args = [os.environ['ACIES_TEST_CORE']]
        if drawing:
            args += ['/i', str(drawing)]
        result = subprocess.run(args + ['/s', str(script)], cwd=root, capture_output=True, timeout=90,
                                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        (root / 'fixture.log').write_bytes(result.stdout + result.stderr)
        self.assertEqual(0, result.returncode)

    def worker(self, root, drawing, **job):
        return clean.run_worker(os.environ['ACIES_TEST_CORE'], os.environ['ACIES_TEST_CLEAN_DLL'], drawing, job, root)

    def test_pixels_survive_without_image_or_plugin(self):
        # Keep artifacts when requested, for manual visual inspection / plotting.
        with tempfile.TemporaryDirectory(prefix='acies-native-image-') as temporary:
            root = Path(os.environ.get('ACIES_NATIVE_ARTIFACTS', temporary))
            root.mkdir(parents=True, exist_ok=True)
            image = root / 'logo.png'
            pixels = Image.new('RGB', (8, 6), (255, 255, 255))
            draw = ImageDraw.Draw(pixels)
            draw.rectangle((0, 0, 3, 1), fill=(255, 0, 0))
            draw.rectangle((4, 0, 7, 1), fill=(0, 255, 0))
            draw.rectangle((0, 2, 7, 5), fill=(0, 0, 255))
            pixels.save(image)
            drawing, output = root / 'input.dwg', root / 'embedded.dwg'
            self.core(root, f'''(setvar "FILEDIA" 0)
(command "_.-IMAGE" "_ATTACH" "{image.as_posix()}" "10,20" 8 90)
(command "_.QSAVE" "{drawing.as_posix()}")''')
            original = drawing.read_bytes()
            result = self.worker(root, drawing, Operation='embed-prototype', Output=str(output))
            self.assertEqual(1, result['embedding']['images'])
            self.assertEqual(3, result['embedding']['solids'])
            self.assertEqual(original, drawing.read_bytes())
            image.unlink()
            report = root / 'native-entities.txt'
            # No NETLOAD: inspect all block definitions using built-in AutoLISP only.
            self.core(root, f'''
(setq out (open "{report.as_posix()}" "w"))
(setq blk (tblnext "BLOCK" T))
(while blk (setq ent (cdr (assoc -2 blk))) (while ent (setq dat (entget ent)) (if (member (cdr (assoc 0 dat)) '("SOLID" "IMAGE" "ACAD_PROXY_ENTITY")) (prin1 dat out)) (setq ent (entnext ent))) (setq blk (tblnext "BLOCK")))
(close out)''', output)
            entities = report.read_text()
            self.assertNotIn('"IMAGE"', entities)
            self.assertNotIn('ACAD_PROXY_ENTITY', entities)
            self.assertEqual(3, entities.count('"SOLID"'))
            for color in (16711680, 65280, 255):
                self.assertIn(f'(420 . {color})', entities)
            # At 90 degrees the 8x6 image spans x=4..10 and y=20..28.
            self.assertIn('(12 4.0 20.0 0.0)', entities)
            self.assertIn('(11 10.0 28.0 0.0)', entities)
            reopened = self.worker(root, output, Operation='verify')
            self.assertEqual(result['validation'], reopened)
            if os.environ.get('ACIES_NATIVE_PLOT'):
                import pymupdf
                pdf = root / 'embedded.pdf'
                self.core(root, f'''(setvar "FILEDIA" 0)
(command "_.-PLOT" "_Y" "Model" "DWG To PDF.pc3" "ANSI full bleed A (8.50 x 11.00 Inches)" "_I" "_P" "_N" "_E" "_F" "_C" "_Y" "." "_Y" "_A" "{pdf.as_posix()}" "_N" "_Y")''', output)
                with pymupdf.open(pdf) as document:
                    colors = {tuple(item['fill']) for item in document[0].get_drawings() if item['fill']}
                    self.assertEqual({(1, 0, 0), (0, 1, 0), (0, 0, 1)}, colors)

    def test_unsupported_alpha_does_not_save_output(self):
        with tempfile.TemporaryDirectory(prefix='acies-native-alpha-') as temporary:
            root = Path(temporary)
            image = root / 'alpha.png'
            Image.new('RGBA', (4, 4), (255, 0, 0, 128)).save(image)
            drawing, output = root / 'input.dwg', root / 'rejected.dwg'
            self.core(root, f'''(setvar "FILEDIA" 0)
(command "_.-IMAGE" "_ATTACH" "{image.as_posix()}" "0,0" 1 0)
(command "_.IMAGECLIP" (entlast) "_OFF")
(command "_.QSAVE" "{drawing.as_posix()}")''')
            original = drawing.read_bytes()
            with self.assertRaisesRegex(RuntimeError, 'Alpha transparency'):
                self.worker(root, drawing, Operation='embed-prototype', Output=str(output))
            self.assertFalse(output.exists())
            self.assertEqual(original, drawing.read_bytes())

    def test_complexity_and_partial_clip_are_rejected(self):
        with tempfile.TemporaryDirectory(prefix='acies-native-limits-') as temporary:
            root = Path(temporary)
            for name, dimensions, clip in [('complex', (225, 225), False), ('clipped', (4, 4), True)]:
                with self.subTest(name=name):
                    image = root / (name + '.png')
                    pixels = Image.new('RGB', dimensions)
                    pixels.putdata([((i >> 16) & 255, (i >> 8) & 255, i & 255) for i in range(dimensions[0] * dimensions[1])])
                    pixels.save(image)
                    drawing, output = root / (name + '.dwg'), root / (name + '-rejected.dwg')
                    clipping = '(command "_.IMAGECLIP" (entlast) "_NEW" "_RECTANGULAR" "0,0" "0.5,0.5")' if clip else ''
                    self.core(root, f'''(setvar "FILEDIA" 0)
(command "_.-IMAGE" "_ATTACH" "{image.as_posix()}" "0,0" 1 0)
{clipping}
(command "_.QSAVE" "{drawing.as_posix()}")''')
                    with self.assertRaisesRegex(RuntimeError, 'unclipped' if clip else '50,000 solids'):
                        self.worker(root, drawing, Operation='embed-prototype', Output=str(output))
                    self.assertFalse(output.exists())


if __name__ == '__main__':
    unittest.main()
