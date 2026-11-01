import argparse
import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import validated_export
import run_in_solidworks
from runner_safety import TestProfile


class ExportIsolationTests(unittest.TestCase):
    def test_pipeline_runs_reviewed_generated_code_on_a_private_copy(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / 'part.sldprt'
            source.write_bytes(b'original part')
            args = argparse.Namespace(source=str(source), target=str(root / 'out.step'),
                tolerances=str(root / 'tolerances.json'), profile=None, overwrite=False,
                configuration=None, timeout=10)

            def subprocess_boundary(command, **unused):
                copy = Path(command[command.index('--open') + 1])
                self.assertNotEqual(copy, source)
                self.assertEqual(copy.read_bytes(), b'original part')
                profile_path = Path(command[command.index('--profile') + 1])
                macro = Path(command[3]).read_text(encoding='utf-8')
                defines = [command[index + 1] for index, value in enumerate(command) if value == '--define']
                code = run_in_solidworks.prepare_code(macro, defines)
                profile = TestProfile.load(profile_path, code)
                profile.validate_request(copy, [], [], False)
                self.assertEqual(profile.root, copy.parent)
                # Stop before any geometry or SW operations; the job must remain inspectable.
                report = Path(command[command.index('--json') + 1])
                report.write_text(json.dumps({'status': 'BLOCKED_DIALOG', 'macro_still_running': True}))

            with patch.object(validated_export.subprocess, 'run', side_effect=subprocess_boundary):
                report = validated_export.execute(args)
            self.assertEqual(report['status'], 'BLOCKED_EXPORT')
            self.assertTrue(Path(report['job_folder']).is_dir())
            self.assertEqual(source.read_bytes(), b'original part')
            self.assertFalse((root / 'out.step').exists())


if __name__ == '__main__':
    unittest.main()
