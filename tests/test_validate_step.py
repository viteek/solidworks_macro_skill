import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import validate_step


class StepResourceTests(unittest.TestCase):
    def test_size_limit_precedes_hash_and_parser(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'model.step'
            path.write_bytes(b'12345')
            with patch.object(validate_step, 'file_hash') as digest, \
                    patch.object(validate_step, 'read_step_metrics') as parser:
                with self.assertRaisesRegex(ValueError, 'exceeds maximum'):
                    validate_step.validate_step(path, {}, {}, max_bytes=4)
                digest.assert_not_called()
                parser.assert_not_called()

    def test_equal_limit_is_accepted(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'model.step'
            path.write_bytes(b'1234')
            with patch.object(validate_step, 'file_hash', return_value='hash'), \
                    patch.object(validate_step, 'read_step_metrics', return_value={}) as parser, \
                    patch.object(validate_step, 'compare_metrics', return_value={'status': 'PASS'}):
                result = validate_step.validate_step(path, {}, {}, max_bytes=4)
            self.assertEqual(result['sha256'], 'hash')
            parser.assert_called_once_with(path)

    def test_input_changed_during_analysis_is_refused(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'model.step'
            path.write_bytes(b'original')
            def parse_while_replaced(parsed):
                parsed.write_bytes(b'replaced')
                return {}
            with patch.object(validate_step, 'read_step_metrics', side_effect=parse_while_replaced), \
                    patch.object(validate_step, 'compare_metrics', return_value={'status': 'PASS'}) as compare:
                with self.assertRaisesRegex(ValueError, 'INPUT_CHANGED'):
                    validate_step.validate_step(path, {}, {})
            compare.assert_not_called()

    def test_timeout_terminates_then_kills_unresponsive_worker(self):
        process = Mock()
        process.communicate.side_effect = [subprocess.TimeoutExpired('worker', 1),
                                           subprocess.TimeoutExpired('worker', 2), ('', None)]
        with patch.object(validate_step.subprocess, 'Popen', return_value=process):
            with self.assertRaisesRegex(TimeoutError, 'exceeded'):
                validate_step.run_isolated('model.step', {}, {}, timeout=1)
        process.terminate.assert_called_once()
        process.kill.assert_called_once()
        self.assertEqual(process.communicate.call_count, 3)

    def test_worker_failure_is_controlled(self):
        process = Mock(returncode=9)
        process.communicate.return_value = ('', None)
        with patch.object(validate_step.subprocess, 'Popen', return_value=process):
            with self.assertRaisesRegex(RuntimeError, 'code 9'):
                validate_step.run_isolated('model.step', {}, {})

    def test_timeout_stops_real_worker_process(self):
        popen = subprocess.Popen
        started = []

        def sleeping_worker(*args, **kwargs):
            process = popen([sys.executable, '-c', 'import time; time.sleep(30)'], **kwargs)
            started.append(process)
            return process

        with patch.object(validate_step.subprocess, 'Popen', side_effect=sleeping_worker):
            with self.assertRaises(TimeoutError):
                validate_step.run_isolated('model.step', {}, {}, timeout=0.2)
        self.assertIsNotNone(started[0].poll())

    def test_cli_timeout_writes_error_report(self):
        with tempfile.TemporaryDirectory() as folder:
            folder = Path(folder)
            model, metrics, tolerances, report = [folder / name for name in
                                                  ('model.step', 'metrics.json', 'tol.json', 'report.json')]
            model.write_bytes(b'123')
            metrics.write_text('{}', encoding='utf-8')
            tolerances.write_text('{}', encoding='utf-8')
            args = ['validate_step', str(model), '--source-metrics', str(metrics),
                    '--tolerances', str(tolerances), '--json', str(report)]
            with patch.object(sys, 'argv', args), \
                    patch.object(validate_step, 'run_isolated', side_effect=TimeoutError('timed out')):
                self.assertEqual(validate_step.main(), 1)
            document = json.loads(report.read_text(encoding='utf-8'))
            self.assertEqual(document['status'], 'ERROR')
            self.assertEqual(document['exception'], 'timed out')

    def test_cli_reports_oversized_input_without_ocp(self):
        with tempfile.TemporaryDirectory() as folder:
            folder = Path(folder)
            model, metrics, tolerances, report = [folder / name for name in
                                                  ('model.step', 'metrics.json', 'tol.json', 'report.json')]
            model.write_bytes(b'12345')
            metrics.write_text('{}', encoding='utf-8')
            tolerances.write_text('{}', encoding='utf-8')
            result = subprocess.run([sys.executable, str(Path(validate_step.__file__)), str(model),
                                     '--source-metrics', str(metrics), '--tolerances', str(tolerances),
                                     '--json', str(report), '--max-bytes', '4'],
                                    capture_output=True, text=True, timeout=20)
            self.assertEqual(result.returncode, 1, result.stderr)
            document = json.loads(report.read_text(encoding='utf-8'))
            self.assertEqual(document['status'], 'ERROR')
            self.assertIn('exceeds maximum', document['exception'])


if __name__ == '__main__':
    unittest.main()
