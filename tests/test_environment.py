import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]


class EnvironmentTests(unittest.TestCase):
    def test_offline_report_does_not_claim_live_testing(self):
        with tempfile.TemporaryDirectory() as directory:
            report = Path(directory) / 'environment.json'
            result = subprocess.run([sys.executable, '-B', str(ROOT / 'scripts/verify_environment.py'),
                '--sw-dir', directory, '--json', str(report)], capture_output=True, text=True)
            self.assertEqual(result.returncode, 0, result.stderr)
            data = json.loads(report.read_text())
            self.assertEqual(data['api_status'], 'UNAVAILABLE')
            self.assertEqual(data['live_status'], 'NOT_PROBED')
            self.assertEqual(data['compilation_status'], 'NOT_TESTED')

    def test_required_api_rejects_missing_type_libraries(self):
        with tempfile.TemporaryDirectory() as directory:
            result = subprocess.run([sys.executable, '-B', str(ROOT / 'scripts/verify_environment.py'),
                '--sw-dir', directory, '--require-api'], capture_output=True, text=True)
            self.assertNotEqual(result.returncode, 0)


if __name__ == '__main__':
    unittest.main()
