from pathlib import Path
import json
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import smoke_assemblies
from runner_safety import TestProfile, file_hash


class AssemblySmokeSafetyTests(unittest.TestCase):
    def make_case(self, folder):
        macro = folder / 'test.bas'
        macro.write_text('Option Explicit\nPublic Sub main()\nEnd Sub\n', encoding='utf-8')
        part = folder / 'copy.sldprt'
        part.write_bytes(b'disposable fixture')
        dependency = folder / 'dependency.sldprt'
        dependency.write_bytes(b'disposable dependency')
        template = folder / 'copy.asmdot'
        template.write_bytes(b'disposable template')
        return macro, part, dependency, template

    def test_prepare_cli_never_launches_a_process_or_loads_com(self):
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary).resolve()
            inputs = []
            for name in ('blank.sldprt', 'block.sldprt', 'assembly.asmdot', 'drawing.drwdot'):
                path = source / name
                path.write_bytes(name.encode())
                inputs.append(path)
            arguments = ['smoke_assemblies.py', '--blank', str(inputs[0]), '--block', str(inputs[1]),
                         '--assembly-template', str(inputs[2]), '--drawing-template', str(inputs[3]),
                         '--test-root', str(source / 'prepared'), '--prepare']
            with patch.object(sys, 'argv', arguments), patch.object(smoke_assemblies.subprocess, 'run') as dispatch:
                self.assertEqual(smoke_assemblies.main(), 0)
            dispatch.assert_not_called()
            data = json.loads((source / 'prepared/fixture-profile.json').read_text())
            self.assertFalse(data['reviewed'])
            self.assertTrue(data['copies_confirmed'])
            self.assertEqual(len(data['inputs']), 4)
            self.assertTrue(any(path.endswith('fixture.sldasm') for path in data['outputs']))

    def test_preparation_does_not_dispatch_and_hashes_complete_local_inputs(self):
        with tempfile.TemporaryDirectory() as temporary:
            folder = Path(temporary).resolve()
            macro, part, dependency, template = self.make_case(folder)
            log = folder / 'result.log'
            with patch.object(smoke_assemblies.subprocess, 'run') as dispatch:
                result = smoke_assemblies.run_step(folder, 'case', macro, part, [], [log], [(log, 'PASS')], False, 10)
            dispatch.assert_not_called()
            self.assertEqual(result['status'], 'NOT_TESTED')
            data = json.loads((folder / 'case-profile.json').read_text())
            self.assertFalse(data['reviewed'])
            self.assertEqual({item['path'] for item in data['inputs']}, {str(part), str(dependency), str(template)})
            self.assertEqual({item['sha256'] for item in data['inputs']}, {file_hash(part), file_hash(dependency), file_hash(template)})
            data['reviewed'] = True
            profile = TestProfile(data)
            profile.validate_request(str(part), [str(log)], [], True)
            command = json.loads((folder / 'case-command.json').read_text())
            self.assertNotIn('--answer', command)
            self.assertNotIn('--recover-session', command)
            self.assertIn('--close-opened', command)

    def test_controller_timeout_returns_unknown_without_automatic_retry(self):
        with tempfile.TemporaryDirectory() as temporary:
            folder = Path(temporary).resolve()
            macro, part, _, _ = self.make_case(folder)
            checked = subprocess.CompletedProcess([], 0, b'checked', b'')
            timeout = subprocess.TimeoutExpired(['runner'], 40)
            with patch.object(smoke_assemblies.subprocess, 'run', side_effect=[checked, checked, timeout]) as dispatch:
                result = smoke_assemblies.run_step(folder, 'case', macro, part, [], [folder / 'result.log'], [], True, 10)
            self.assertEqual(result['status'], 'UNKNOWN')
            self.assertEqual(result['session_state'], 'UNKNOWN')
            self.assertEqual(dispatch.call_count, 3)  # two offline checks, one live attempt

    def test_pass_requires_zero_controller_exit_and_idle_session(self):
        # The controller may write PASS before failing durable state finalization.
        for returncode, session, expected in ((1, 'IDLE', 'UNKNOWN'),
                                               (0, 'UNKNOWN', 'UNKNOWN'),
                                               (0, None, 'UNKNOWN'),
                                               (0, 'IDLE', 'PASS')):
            with self.subTest(returncode=returncode, session=session), tempfile.TemporaryDirectory() as temporary:
                folder = Path(temporary).resolve()
                macro, part, _, _ = self.make_case(folder)
                calls = []

                def controller(command, **kwargs):
                    calls.append(command)
                    if len(calls) == 3:
                        (folder / 'case-runner.json').write_text(json.dumps({'status': 'PASS', 'session_state': session}))
                        return subprocess.CompletedProcess(command, returncode, b'finished', b'')
                    return subprocess.CompletedProcess(command, 0, b'checked', b'')

                with patch.object(smoke_assemblies.subprocess, 'run', side_effect=controller):
                    result = smoke_assemblies.run_step(folder, 'case', macro, part, [], [folder / 'result.log'], [], True, 10)
                self.assertEqual(result['status'], expected)
                self.assertEqual(result['controller_exit_code'], returncode)
                self.assertEqual(result['runner_status'], 'PASS')
                if expected == 'UNKNOWN':
                    self.assertEqual(result['session_state'], 'UNKNOWN')
                self.assertEqual(len(calls), 3)  # no recovery or second controller invocation

    def test_wrapper_keeps_only_one_public_parameterless_entry(self):
        with tempfile.TemporaryDirectory() as temporary:
            macro = Path(temporary) / 'recipe.bas'
            macro.write_text('Option Explicit\nPublic Sub main()\nEnd Sub\n')
            wrapped = smoke_assemblies.wrapper(macro, '    Debug.Print "fixture"')
            self.assertIn('Private Sub RunRecipe()', wrapped)
            self.assertEqual(wrapped.count('Public Sub main()'), 1)


if __name__ == '__main__':
    unittest.main()
