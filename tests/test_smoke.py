from pathlib import Path
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import smoke_modeling


class SmokeAssertionsTests(unittest.TestCase):
    def test_volume_and_feature_readback_are_required(self):
        observation = {'collection_status': 'OK', 'solid_body_count': 1, 'volume_m3': .000012,
                       'features': [{'name': 'Recipe_Block', 'type': 'Extrusion'}]}
        self.assertEqual(smoke_modeling.assert_result('modify', observation), [])
        observation['volume_m3'] = .000008
        self.assertTrue(smoke_modeling.assert_result('modify', observation))
        del observation['features']
        self.assertTrue(smoke_modeling.assert_result('modify', observation))

    def test_missing_failed_and_nonfinite_observations_cannot_pass(self):
        for observation in ({}, {'collection_status': 'FAIL'},
            {'collection_status': 'OK', 'solid_body_count': 1, 'volume_m3': float('nan'),
             'features': [{'name': 'Recipe_Block'}]}):
            with self.subTest(observation=observation):
                self.assertTrue(smoke_modeling.assert_result('modify', observation))


if __name__ == '__main__':
    unittest.main()
