"""Offline validation of independent sheet-metal geometry evidence."""
import copy
from pathlib import Path
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from smoke_sheetmetal import assert_metrics


class SheetMetalSmokeTests(unittest.TestCase):
    def setUp(self):
        self.observation = {'units': 'mm', 'bodies': [
            {'volume_mm3': 10000, 'bounds_mm': [0, 0, -2, 100, 50, 0]}]}

    def test_observed_blank_matches_independent_volume_and_sizes(self):
        self.assertEqual(assert_metrics(self.observation), [])

    def test_matching_volume_does_not_hide_wrong_dimensions(self):
        self.observation['bodies'][0]['bounds_mm'] = [0, 0, 0, 200, 25, 2]
        self.assertIn('Independent dimensions mismatch', assert_metrics(self.observation))

    def test_missing_or_incorrect_unit_contract_cannot_pass(self):
        for units in (None, 'm', 'in'):
            with self.subTest(units=units):
                self.observation['units'] = units
                self.assertTrue(assert_metrics(self.observation))

    def test_dimensions_do_not_hide_incorrect_volume(self):
        self.observation['bodies'][0]['volume_mm3'] = 9990
        self.assertIn('Independent body volume mismatch', assert_metrics(self.observation))

    def test_missing_or_multiple_bodies_cannot_pass(self):
        for value in (None, {}, {'bodies': []}, {'bodies': [None]},
                      {'bodies': [copy.deepcopy(self.observation['bodies'][0])] * 2}):
            with self.subTest(value=value):
                self.assertTrue(assert_metrics(value))

    def test_nonfinite_boolean_or_text_measurements_cannot_pass(self):
        for bad in (float('nan'), float('inf'), True, '10000', None):
            with self.subTest(volume=bad):
                self.observation['bodies'][0]['volume_mm3'] = bad
                self.assertTrue(assert_metrics(self.observation))

    def test_malformed_or_inverted_bounds_cannot_pass(self):
        for bounds in (None, [], [0] * 5, [0, 0, 0, float('nan'), 50, 2],
                       [0, 0, 0, True, 50, 2], [100, 0, 0, 0, 50, 2]):
            with self.subTest(bounds=bounds):
                self.observation['bodies'][0]['bounds_mm'] = bounds
                self.assertTrue(assert_metrics(self.observation))


if __name__ == '__main__':
    unittest.main()
