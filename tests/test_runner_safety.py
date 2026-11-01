import hashlib
import json
from pathlib import Path
import sys
import tempfile
import subprocess
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import runner_safety as safety


class ProfileTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.source = self.root / 'copy.sldprt'
        self.source.write_bytes(b'CAD copy')
        self.code = 'Option Explicit\nSub main()\nEnd Sub\n'
        self.profile = self.root / 'profile.json'
        self.data = {
            'schema_version': 1, 'test_root': str(self.root),
            'reviewed': True, 'copies_confirmed': True, 'allow_model_changes': False,
            'macro_sha256': hashlib.sha256(self.code.encode('utf-8')).hexdigest(),
            'entry_point': {'module': 'TestMacro', 'procedure': 'main'},
            'inputs': [{'path': str(self.source), 'sha256': hashlib.sha256(b'CAD copy').hexdigest()}],
            'outputs': [str(self.root / 'out.step')], 'allowed_dialogs': []}

    def load(self, **changes):
        self.data.update(changes)
        self.profile.write_text(json.dumps(self.data))
        return safety.TestProfile.load(self.profile, self.code)

    def test_accepts_reviewed_copies_and_declared_outputs(self):
        profile = self.load()
        profile.validate_request(self.source, [self.root / 'out.step'], [], False)

    def test_rejects_unreviewed_or_changed_code_and_inputs(self):
        for changes in ({'reviewed': False}, {'copies_confirmed': False}, {'reviewed': 'true'},
                        {'macro_sha256': '0' * 64}):
            with self.subTest(changes=changes):
                saved = dict(self.data)
                with self.assertRaises(ValueError):
                    self.load(**changes)
                self.data = saved
        self.source.write_bytes(b'changed after review')
        with self.assertRaises(ValueError):
            self.load()

    def test_rejects_outside_root_and_undeclared_result(self):
        profile = self.load()
        for target in (self.root.parent / 'escape.step', self.root / 'undeclared.step'):
            with self.subTest(target=target), self.assertRaises(ValueError):
                profile.validate_request(self.source, [target], [], False)
        with self.assertRaises(ValueError):
            self.load(outputs=[str(self.root.parent / 'escape.step')])

    def test_requires_separate_confirmation_for_model_changes(self):
        profile = self.load(allow_model_changes=True)
        with self.assertRaises(ValueError):
            profile.validate_request(self.source, [], [], False)
        profile.validate_request(self.source, [], [], True)

    def test_entry_point_cannot_change_after_review(self):
        profile = self.load()
        profile.validate_entry('TestMacro', 'main')
        with self.assertRaises(ValueError):
            profile.validate_entry('TestMacro', 'DeleteEverything')

    def test_rejects_external_missing_or_undeclared_dependencies(self):
        profile = self.load()
        for dependencies in (['production', str(self.root.parent / 'production.sldprt')],
                             ['missing', str(self.root / 'missing.sldprt')], ['incomplete']):
            with self.subTest(dependencies=dependencies), self.assertRaises(ValueError):
                profile.validate_dependencies(dependencies)
        profile.validate_dependencies(['copy', str(self.source)])

    def test_rejects_symlink_escape(self):
        with tempfile.TemporaryDirectory() as outside:
            link = self.root / 'link'
            try:
                link.symlink_to(outside, target_is_directory=True)
            except OSError:
                self.skipTest('Symlinks require privileges on this Windows host')
            with self.assertRaises(ValueError):
                self.load(outputs=[str(link / 'escape.step')])


class SessionTests(unittest.TestCase):
    def test_preflight_reporting_never_overwrites_unvalidated_paths(self):
        script = Path(__file__).resolve().parents[1] / 'scripts/run_in_solidworks.py'
        with tempfile.TemporaryDirectory() as directory:
            for name in ('macro.bas', 'original.sldprt', 'production.json'):
                target = Path(directory) / name
                target.write_bytes(b'valuable original content')
                result = subprocess.run([sys.executable, '-B', str(script), str(target),
                    '--timeout', 'nan', '--json', str(target)], capture_output=True, text=True)
                self.assertEqual(result.returncode, 2, result.stderr)
                self.assertEqual(target.read_bytes(), b'valuable original content')

    def test_running_and_unknown_survive_new_controller(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'session.json'
            session = safety.SessionState(path)
            session.start('macro.bas')
            with self.assertRaises(RuntimeError):
                safety.SessionState(path).ensure_ready()
            session.unknown('timeout')
            with self.assertRaises(RuntimeError):
                safety.SessionState(path).ensure_ready()
            with self.assertRaises(ValueError):
                session.recover(False, process_alive=True)
            session.recover(True, process_alive=True)
            safety.SessionState(path).ensure_ready()

    def test_corrupt_state_does_not_allow_a_run(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'session.json'
            path.write_text('{broken')
            with self.assertRaises(RuntimeError):
                safety.SessionState(path).ensure_ready()

    def test_exited_solidworks_process_can_be_recovered(self):
        with tempfile.TemporaryDirectory() as directory:
            state = safety.SessionState(Path(directory) / 'state.json')
            state.start('macro.bas')
            state.set_process(123)
            state.recover(False, process_alive=False)
            state.ensure_ready()


if __name__ == '__main__':
    unittest.main()
