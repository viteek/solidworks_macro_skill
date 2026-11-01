import copy
import contextlib
import hashlib
import io
import json
from pathlib import Path
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import benchmark


class BenchmarkTests(unittest.TestCase):
    def setUp(self):
        self.specs = benchmark.load_specs(Path(__file__).resolve().parents[1] / 'benchmarks/tasks')

    def test_suite_has_exactly_six_families_of_four(self):
        self.assertEqual(len(self.specs), 24)
        self.assertEqual(sorted([sum(s['family'] == f for s in self.specs)
                                 for f in {s['family'] for s in self.specs}]), [4] * 6)
        for spec in self.specs:
            self.assertTrue(spec['fixture_instructions'])
            self.assertTrue(spec['assertions'])

    def test_missing_task_spec_cannot_shrink_denominator(self):
        with tempfile.TemporaryDirectory() as tmp:
            (Path(tmp) / 'T01.json').write_text(json.dumps(self.specs[0]))
            with self.assertRaises(ValueError):
                benchmark.load_specs(tmp)

    def records(self, variant):
        return [{'run_id': variant + s['task_id'], 'variant': variant,
                 'task_id': s['task_id'], 'fixture': s['fixture'],
                 'environment': {'solidworks_version': '2024', 'service_pack': '5',
                                 'api_revision': '32.5', 'architecture': 'x64', 'ui_language': 'en-US'},
                 'report': {'schema_version': 1, 'task_id': s['task_id'],
                            'fixture': s['fixture'], 'spec_sha256': benchmark.spec_hash(s),
                            'environment': {'solidworks_version': '2024', 'service_pack': '5',
                                            'api_revision': '32.5', 'architecture': 'x64', 'ui_language': 'en-US'},
                            'provenance': {'run_id': variant + s['task_id'], 'collected_at': '2026-01-01T00:00:00Z',
                                           'macro_sha256': 'a' * 64, 'fixture_sha256': 'b' * 64},
                            'status': 'NOT_TESTED',
                            'phases': {p: {'status': 'NOT_TESTED'} for p in benchmark.PHASES},
                            'assertions': []},
                 'measurements': {'attempts': 1, 'api_errors': 0, 'unintended_saves': 0,
                                  'tokens': 100}, 'measurement_evidence': ['session.json']} for s in self.specs]

    def test_not_tested_is_zero_completion(self):
        result = benchmark.compare(self.specs, self.records('A'), self.records('B'))
        self.assertEqual(result['A']['completed'], 0)
        self.assertEqual(result['B']['not_tested'], 24)
        self.assertEqual(result['delta_B_minus_A']['tokens'], 0)

    def test_missing_duplicate_wrong_fixture_and_version_rejected(self):
        for mutation in ('missing', 'duplicate', 'fixture', 'version', 'measurement', 'fixture_hash'):
            a, b = self.records('A'), self.records('B')
            if mutation == 'missing': b.pop()
            if mutation == 'duplicate': b.append(copy.deepcopy(b[0]))
            if mutation == 'fixture': b[0]['fixture'] = {'id': 'wrong', 'revision': '1'}
            if mutation == 'version': b[0]['environment']['service_pack'] = '4'
            if mutation == 'measurement': del b[0]['measurements']['tokens']
            if mutation == 'fixture_hash': b[0]['report']['provenance']['fixture_sha256'] = 'c' * 64
            with self.subTest(mutation=mutation), self.assertRaises(ValueError):
                benchmark.compare(self.specs, a, b)

    def test_claimed_pass_without_assertions_or_evidence_rejected(self):
        a, b = self.records('A'), self.records('B')
        b[0]['report']['status'] = 'PASS'
        with self.assertRaises(ValueError):
            benchmark.compare(self.specs, a, b)
        b = self.records('B')
        b[0]['measurement_evidence'] = []
        with self.assertRaises(ValueError):
            benchmark.compare(self.specs, a, b)

    def test_complete_reports_count_completion_and_all_retry_costs(self):
        a, b = self.records('A'), self.records('B')
        for record in b:
            spec = next(s for s in self.specs if s['task_id'] == record['task_id'])
            record['report']['status'] = 'PASS'
            record['report']['assertions'] = [{'id': x['id'], 'status': 'PASS'} for x in spec['assertions']]
            for phase in benchmark.PHASES:
                record['report']['phases'][phase] = {'status': 'PASS', 'artifacts': [{'file': 'log.json', 'sha256': 'f' * 64}]}
            record['measurements']['attempts'] = 2
            record['measurements']['tokens'] = 250
        result = benchmark.compare(self.specs, a, b)
        self.assertEqual(result['B']['completed'], 24)
        self.assertEqual(result['B']['attempts'], 48)
        self.assertEqual(result['delta_B_minus_A']['tokens'], 24 * 150)
        b[0]['measurements']['unintended_saves'] = 1
        with self.assertRaises(ValueError):
            benchmark.compare(self.specs, a, b)

    def test_cli_import_binds_measurements_and_detects_evidence_tampering(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            record = self.records('A')[0]
            report = record['report']
            for kind in ('macro', 'fixture'):
                (root / (kind + '.bin')).write_bytes(kind.encode())
                report['provenance'][kind + '_artifact'] = kind + '.bin'
                report['provenance'][kind + '_sha256'] = hashlib.sha256(kind.encode()).hexdigest()
            (root / 'report.json').write_text(json.dumps(report))
            record['report'] = 'report.json'
            record['measurement_evidence'] = ['session.json']
            summary = {'run_id': record['run_id'], 'measurements': record['measurements']}
            (root / 'session.json').write_text(json.dumps(summary))
            (root / 'runs.json').write_text(json.dumps({'schema_version': 1, 'runs': [record]}))
            self.assertEqual(len(benchmark.load_runs(root / 'runs.json')), 1)
            summary['measurements']['tokens'] = 999
            (root / 'session.json').write_text(json.dumps(summary))
            with self.assertRaises(ValueError):
                benchmark.load_runs(root / 'runs.json')
            summary['measurements']['tokens'] = 100
            (root / 'session.json').write_text(json.dumps(summary))
            (root / 'macro.bin').write_bytes(b'modified')
            with self.assertRaises(ValueError):
                benchmark.load_runs(root / 'runs.json')

    def write_cli_manifests(self, root):
        """Valid complete offline manifests, explicitly containing no live PASS claims."""
        for kind in ('macro', 'fixture'):
            (root / (kind + '.bin')).write_bytes(kind.encode())
        (root / 'phase.log').write_bytes(b'offline evidence; no CAD execution claimed')
        for variant in ('A', 'B'):
            records = self.records(variant)
            for record in records:
                report = record['report']
                for kind in ('macro', 'fixture'):
                    report['provenance'][kind + '_artifact'] = kind + '.bin'
                    report['provenance'][kind + '_sha256'] = hashlib.sha256(kind.encode()).hexdigest()
                for phase in benchmark.INPUT_PHASES:
                    report['phases'][phase]['artifacts'] = [
                        {'file': 'phase.log', 'sha256': hashlib.sha256((root / 'phase.log').read_bytes()).hexdigest()}]
                report_name = record['run_id'] + '-report.json'
                (root / report_name).write_text(json.dumps(report))
                record['report'] = report_name
                summary_name = record['run_id'] + '-measurements.json'
                (root / summary_name).write_text(json.dumps({'run_id': record['run_id'], 'measurements': record['measurements']}))
                record['measurement_evidence'] = [summary_name]
            (root / (variant + '-runs.json')).write_text(json.dumps({'schema_version': 1, 'runs': records}))

    def test_cli_existing_output_preserves_manifests_reports_and_measurement_evidence(self):
        for target in ('A-runs.json', 'B-runs.json', 'AT01-report.json', 'AT01-measurements.json',
                       'phase.log', 'macro.bin', 'fixture.bin', 'comparison.json'):
            with self.subTest(output=target), tempfile.TemporaryDirectory() as tmp:
                root = Path(tmp)
                self.write_cli_manifests(root)
                (root / 'comparison.json').write_bytes(b'existing comparison must survive\x00\xff')
                before = {path: path.read_bytes() for path in root.iterdir()}
                argv = ['--a', str(root / 'A-runs.json'), '--b', str(root / 'B-runs.json'),
                        '--output', str(root / target)]
                with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
                    status = benchmark.main(argv)
                for path, original in before.items():
                    self.assertTrue(path.read_bytes() == original, 'Existing file changed: ' + str(path))
                self.assertEqual(status, 3)

    def test_cli_fresh_output_accepts_valid_complete_manifests(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            self.write_cli_manifests(root)
            output = root / 'new-comparison.json'
            argv = ['--a', str(root / 'A-runs.json'), '--b', str(root / 'B-runs.json'), '--output', str(output)]
            with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
                self.assertEqual(benchmark.main(argv), 0)
            self.assertEqual(json.loads(output.read_text())['tasks'], 24)


if __name__ == '__main__':
    unittest.main()
