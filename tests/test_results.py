import copy
import contextlib
import hashlib
import io
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import verify_results as vr


class ResultTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        (self.root / 'log.txt').write_text('independent collector evidence')
        (self.root / 'macro.bas').write_bytes(b'Sub main()\nEnd Sub\n')
        (self.root / 'fixture.json').write_bytes(b'{"fixture":"offline contract only"}')
        self.spec = {'schema_version': 1, 'task_id': 'T01',
                     'fixture': {'id': 'part', 'revision': '1'},
                     'assertions': [{'id': 'length', 'path': '/length', 'op': 'number',
                                     'expected': .1, 'abs_tol': .000001}]}
        self.obs = {'schema_version': 1, 'task_id': 'T01',
                    'fixture': self.spec['fixture'].copy(),
                    'environment': {'solidworks_version': '2024', 'service_pack': '5',
                                    'api_revision': '32.5', 'architecture': 'x64', 'ui_language': 'en-US'},
                    'provenance': {'run_id': 'test', 'collected_at': '2026-01-01T00:00:00Z',
                                   'macro_sha256': hashlib.sha256((self.root / 'macro.bas').read_bytes()).hexdigest(),
                                   'fixture_sha256': hashlib.sha256((self.root / 'fixture.json').read_bytes()).hexdigest(),
                                   'macro_artifact': 'macro.bas', 'fixture_artifact': 'fixture.json'},
                    'values': {'length': .1},
                    'evidence': {p: {'status': 'PASS', 'artifacts': ['log.txt']}
                                 for p in vr.INPUT_PHASES}}

    def check(self):
        return vr.verify(self.spec, self.obs, self.root)

    def test_complete_evidence_and_numeric_tolerance(self):
        self.obs['values']['length'] += .0000005
        self.assertEqual(self.check()['status'], 'PASS')
        self.obs['values']['length'] = .11
        self.assertEqual(self.check()['phases']['result']['status'], 'FAIL')

    def test_missing_and_nonfinite_numbers_never_pass(self):
        del self.obs['values']['length']
        self.assertEqual(self.check()['status'], 'NOT_TESTED')
        self.obs['values']['length'] = float('nan')
        self.assertEqual(self.check()['status'], 'FAIL')

    def test_json_pointer_array_indices_follow_rfc_6901(self):
        values = {'items': [10, 20], 'a/b': 1, 'm~n': 2}
        self.assertEqual(vr.pointer(values, '/items/1'), 20)
        self.assertEqual(vr.pointer(values, '/items/0'), 10)
        self.assertEqual(vr.pointer(values, '/a~1b'), 1)
        self.assertEqual(vr.pointer(values, '/m~0n'), 2)
        for path in ('/items/01', '/items/00', '/items/+1', '/items/-', '/items/١', '/m~2n'):
            with self.subTest(path=path), self.assertRaises(KeyError):
                vr.pointer(values, path)
        self.obs['values'] = {'length': [.1]}
        self.spec['assertions'][0]['path'] = '/length/00'
        self.assertEqual(self.check()['status'], 'NOT_TESTED')

    def test_missing_runtime_and_unbacked_claim_are_not_tested(self):
        del self.obs['evidence']['runtime']
        self.assertEqual(self.check()['phases']['result']['status'], 'NOT_TESTED')
        self.obs['evidence']['runtime'] = {'status': 'PASS', 'artifacts': []}
        self.assertEqual(self.check()['status'], 'NOT_TESTED')

    def test_identity_duplicates_and_path_escape_rejected(self):
        self.obs['fixture']['revision'] = '2'
        with self.assertRaises(ValueError):
            self.check()
        self.obs['fixture']['revision'] = '1'
        self.spec['assertions'] *= 2
        with self.assertRaises(ValueError):
            self.check()
        self.spec['assertions'] = self.spec['assertions'][:1]
        self.obs['evidence']['runtime']['artifacts'] = ['../log.txt']
        with self.assertRaises(ValueError):
            self.check()

    def test_objects_features_mates_properties_and_files(self):
        self.obs['values'] = {'features': ['Boss', 'Cut'], 'mates': 2,
                              'properties': {'PartNo': 'P-01'}, 'object': {'type': 'PartDoc'}}
        (self.root / 'part.step').write_bytes(b'ISO-10303-21;\nEND-ISO-10303-21;')
        self.spec['assertions'] = [
            {'id': 'features', 'path': '/features', 'op': 'unordered_equals', 'expected': ['Cut', 'Boss']},
            {'id': 'mates', 'path': '/mates', 'op': 'equals', 'expected': 2},
            {'id': 'property', 'path': '/properties/PartNo', 'op': 'equals', 'expected': 'P-01'},
            {'id': 'object', 'path': '/object/type', 'op': 'equals', 'expected': 'PartDoc'},
            {'id': 'file', 'op': 'file', 'file': 'part.step', 'min_bytes': 10,
             'contains': ['ISO-10303-21;', 'END-ISO-10303-21;']}]
        self.assertEqual(self.check()['status'], 'PASS')
        (self.root / 'part.step').unlink()
        self.assertNotEqual(self.check()['status'], 'PASS')

    def test_strict_json_rejects_duplicate_keys(self):
        (self.root / 'bad.json').write_text('{"task_id":"A","task_id":"B"}')
        with self.assertRaises(ValueError):
            vr.read_json(self.root / 'bad.json')

    def test_non_object_assertions_rejected_by_validator_and_cli(self):
        spec_path = self.root / 'spec.json'
        obs_path = self.root / 'observations.json'
        output = self.root / 'report.json'
        obs_path.write_text(json.dumps(self.obs), encoding='utf-8')
        for item in (None, 123, 'invalid'):
            with self.subTest(item=item):
                self.spec['assertions'] = [self.spec['assertions'][0], item]
                with self.assertRaisesRegex(ValueError, 'Assertion must be an object'):
                    vr.validate_spec(self.spec)
                spec_path.write_text(json.dumps(self.spec), encoding='utf-8')
                result = subprocess.run(
                    [sys.executable, str(Path(vr.__file__)), '--spec', str(spec_path),
                     '--observations', str(obs_path), '--evidence-root', str(self.root),
                     '--output', str(output)], capture_output=True, text=True, check=False)
                self.assertEqual(result.returncode, 3)
                self.assertIn('Assertion must be an object', result.stderr)
                self.assertNotIn('Traceback', result.stderr)
                self.assertFalse(output.exists())

    def test_file_assertion_rejects_malformed_contains_and_schema_version(self):
        for invalid in ('AB', [None], 12, {'item': 'A'}):
            with self.subTest(contains=invalid):
                self.spec['assertions'] = [{'id': 'file', 'op': 'file', 'file': 'part.step',
                                            'contains': invalid}]
                with self.assertRaisesRegex(ValueError, 'contains must be a list of strings'):
                    vr.validate_spec(self.spec)
        self.spec['assertions'] = [{'id': 'length', 'path': '/length', 'op': 'number', 'expected': .1}]
        self.spec['schema_version'] = True
        with self.assertRaisesRegex(ValueError, 'schema_version'):
            vr.validate_spec(self.spec)

    def test_object_fields_reject_wrong_types_via_api_and_cli(self):
        spec_path = self.root / 'spec.json'
        obs_path = self.root / 'observations.json'
        output = self.root / 'report.json'
        fields = [('spec', ()), ('observations', ()), ('spec', ('fixture',)),
                  ('observations', ('fixture',)), ('observations', ('environment',)),
                  ('observations', ('provenance',)), ('observations', ('evidence',)),
                  ('observations', ('values',))]
        fields += [('observations', ('evidence', phase)) for phase in vr.INPUT_PHASES]
        for document, path in fields:
            for invalid in (None, [], 123, 'invalid'):
                with self.subTest(document=document, path=path, invalid=invalid):
                    spec, obs = copy.deepcopy(self.spec), copy.deepcopy(self.obs)
                    if not path:
                        if document == 'spec':
                            spec = invalid
                        else:
                            obs = invalid
                    else:
                        parent = spec if document == 'spec' else obs
                        for key in path[:-1]:
                            parent = parent[key]
                        parent[path[-1]] = invalid
                    with self.assertRaises(ValueError):
                        vr.verify(spec, obs, self.root)
                    spec_path.write_text(json.dumps(spec), encoding='utf-8')
                    obs_path.write_text(json.dumps(obs), encoding='utf-8')
                    result = subprocess.run(
                        [sys.executable, str(Path(vr.__file__)), '--spec', str(spec_path),
                         '--observations', str(obs_path), '--evidence-root', str(self.root),
                         '--output', str(output)], capture_output=True, text=True, check=False)
                    self.assertEqual(result.returncode, 3, result.stderr)
                    self.assertNotIn('Traceback', result.stderr)
                    self.assertFalse(output.exists())

    def test_stale_macro_missing_version_and_incomplete_safety(self):
        (self.root / 'macro.bas').write_text('changed')
        with self.assertRaises(ValueError):
            self.check()
        (self.root / 'macro.bas').write_bytes(b'Sub main()\nEnd Sub\n')
        del self.obs['environment']['ui_language']
        with self.assertRaises(ValueError):
            self.check()
        self.obs['environment']['ui_language'] = 'en-US'
        del self.obs['evidence']['safety']
        report = self.check()
        self.assertEqual(report['phases']['result']['status'], 'PASS')
        self.assertEqual(report['status'], 'NOT_TESTED')

    def test_unordered_assertion_keeps_duplicate_counts_and_boolean_type(self):
        self.spec['assertions'] = [{'id': 'objects', 'op': 'unordered_equals',
                                   'path': '/objects', 'expected': ['A', 'A', 'B']},
                                  {'id': 'count', 'op': 'equals', 'path': '/count', 'expected': 1}]
        self.obs['values'] = {'objects': ['A', 'B', 'B'], 'count': True}
        report = self.check()
        self.assertEqual([r['status'] for r in report['assertions']], ['FAIL', 'FAIL'])

    def test_cli_reports_explicit_exit_codes_and_serializable_evidence(self):
        spec_path, obs_path, output = [self.root / p for p in ('spec.json', 'observations.json', 'report.json')]
        spec_path.write_text(json.dumps(self.spec))
        argv = ['--spec', str(spec_path), '--observations', str(obs_path),
                '--evidence-root', str(self.root), '--output', str(output)]
        for expected in (0, 1, 2, 3):
            output = self.root / ('report-' + str(expected) + '.json')
            argv[-1] = str(output)
            self.obs['values']['length'] = .1
            self.obs['evidence']['runtime']['status'] = 'PASS'
            if expected == 1:
                self.obs['values']['length'] = 100
            if expected == 2:
                self.obs['evidence']['runtime']['status'] = 'NOT_TESTED'
            if expected == 3:
                self.obs['task_id'] = 'wrong'
            obs_path.write_text(json.dumps(self.obs))
            with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
                self.assertEqual(vr.main(argv), expected)
            if expected < 3:
                self.assertEqual(vr.read_json(output)['status'], {0:'PASS', 1:'FAIL', 2:'NOT_TESTED'}[expected])

    def test_cli_existing_output_preserves_spec_observations_and_evidence(self):
        for target in ('spec.json', 'observations.json', 'log.txt', 'macro.bas', 'fixture.json', 'report.json'):
            with self.subTest(output=target):
                spec_path = self.root / 'spec.json'
                obs_path = self.root / 'observations.json'
                spec_path.write_text(json.dumps(self.spec))
                obs_path.write_text(json.dumps(self.obs))
                (self.root / 'log.txt').write_text('independent collector evidence')
                (self.root / 'macro.bas').write_bytes(b'Sub main()\nEnd Sub\n')
                (self.root / 'fixture.json').write_bytes(b'{"fixture":"offline contract only"}')
                (self.root / 'report.json').write_bytes(b'existing report must survive\x00\xff')
                before = {p: p.read_bytes() for p in self.root.iterdir() if p.is_file()}
                argv = ['--spec', str(spec_path), '--observations', str(obs_path),
                        '--evidence-root', str(self.root), '--output', str(self.root / target)]
                with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
                    status = vr.main(argv)
                for path, original in before.items():
                    self.assertTrue(path.read_bytes() == original, 'Existing file changed: ' + str(path))
                self.assertEqual(status, 3)


if __name__ == '__main__':
    unittest.main()
