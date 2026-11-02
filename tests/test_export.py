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
    def test_existing_report_is_rejected_before_export_side_effects(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source, target, tolerances, report = [root / name for name in
                                                   ('part.sldprt', 'out.step', 'tolerances.json', 'report.json')]
            source.write_bytes(b'CAD input')
            tolerances.write_text('{}', encoding='utf-8')
            report.write_text('keep existing report', encoding='utf-8')
            with patch.object(sys, 'argv', ['validated_export.py', str(source), str(target),
                                            '--tolerances', str(tolerances), '--json', str(report)]), \
                    patch.object(validated_export, 'execute') as execute:
                with self.assertRaises(SystemExit) as error:
                    validated_export.main()
            self.assertEqual(error.exception.code, 2)
            execute.assert_not_called()
            self.assertEqual(report.read_text(encoding='utf-8'), 'keep existing report')
            self.assertFalse(target.exists())

    def test_missing_report_parent_is_rejected_before_export_side_effects(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source, target, tolerances = [root / name for name in
                                           ('part.sldprt', 'out.step', 'tolerances.json')]
            source.write_bytes(b'CAD input')
            tolerances.write_text('{}', encoding='utf-8')
            report = root / 'missing' / 'report.json'
            with patch.object(sys, 'argv', ['validated_export.py', str(source), str(target),
                                            '--tolerances', str(tolerances), '--json', str(report)]), \
                    patch.object(validated_export, 'execute') as execute:
                with self.assertRaises(SystemExit) as error:
                    validated_export.main()
            self.assertEqual(error.exception.code, 2)
            execute.assert_not_called()
            self.assertFalse(target.exists())
            self.assertFalse(report.parent.exists())

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

    def run_exporter(self, root, log_text, raw_bytes=b'ISO-10303-21;'):
        source = root / 'part.sldprt'
        source.write_bytes(b'original part')
        args = argparse.Namespace(source=str(source), target=str(root / 'out.step'),
            tolerances=str(root / 'tolerances.json'), profile=None, overwrite=False,
            configuration=None, timeout=10)
        commands = []

        def subprocess_boundary(command, **unused):
            commands.append(command)
            job = Path(command[command.index('--open') + 1]).parent
            (job / 'raw.step').write_bytes(raw_bytes)
            (job / 'macro.log').write_text(log_text, encoding='cp1250')
            report = Path(command[command.index('--json') + 1])
            report.write_text(json.dumps({'status': 'PASS'}))

        with patch.object(validated_export.subprocess, 'run', side_effect=subprocess_boundary), \
                patch.object(validated_export, 'publish_validated') as publish:
            report = validated_export.execute(args)
        return commands[0], report, publish

    def test_runner_must_confirm_final_export_status_and_fresh_outputs(self):
        with tempfile.TemporaryDirectory() as directory:
            command, _, _ = self.run_exporter(Path(directory), 'ERROR;not reached\n')
            job = Path(command[command.index('--open') + 1]).parent
            pairs = [command[index + 1:index + 3] for index, value in enumerate(command) if value == '--expect-result']
            self.assertEqual(pairs, [[str(job / 'macro.log'), 'EXPORTED']])
            files = [command[index + 1] for index, value in enumerate(command) if value == '--expect-file']
            self.assertEqual(files, [str(job / 'raw.step'), str(job / 'source-metrics.json')])

    def test_export_line_followed_by_cleanup_error_blocks_validation(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            _, report, publish = self.run_exporter(root, 'EXPORTED;out\nERROR_CLEANUP; selection restore failed\n')
            self.assertEqual(report['status'], 'BLOCKED_EXPORT')
            publish.assert_not_called()
            self.assertFalse((root / 'out.step').exists())


if __name__ == '__main__':
    unittest.main()
