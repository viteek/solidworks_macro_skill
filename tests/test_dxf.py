import json
import math
from pathlib import Path
import sys
import tempfile
import unittest

import ezdxf

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from validate_dxf import validate_dxf


class PolylineTests(unittest.TestCase):
    def test_bulge_direction_preserves_contour_and_area(self):
        profile = json.loads((Path(__file__).resolve().parents[1] / 'profiles/neutral-mm.json').read_text())
        # A 10 x 10 square with one semicircular segment adds/removes pi * 5^2 / 2.
        for segment in (0, 1, 3):
            for bulge in (-1, 1):
                with self.subTest(segment=segment, bulge=bulge), tempfile.TemporaryDirectory() as directory:
                    document = ezdxf.new()
                    document.units = 4
                    document.layers.new('CUT')
                    vertices = [(0, 0), (10, 0), (10, 10), (0, 10)]
                    document.modelspace().add_lwpolyline(
                        [(x, y, bulge if i == segment else 0) for i, (x, y) in enumerate(vertices)],
                        format='xyb', close=True, dxfattribs={'layer': 'CUT'})
                    path = Path(directory) / 'semicircle.dxf'
                    document.saveas(path)
                    report = validate_dxf(path, profile)
                    self.assertEqual(report['status'], 'PASS', report['issues'])
                    self.assertAlmostEqual(report['metrics']['net_area_mm2'],
                                           100 + bulge * math.pi * 25 / 2, delta=0.02)
                    self.assertEqual(report['metrics']['outer_contours'], 1)
                    self.assertEqual(report['metrics']['hole_count'], 0)


if __name__ == '__main__':
    unittest.main()
