"""Run the real CLI flow with COM and Windows UI replaced at their boundaries."""
import contextlib
import hashlib
import io
import json
from pathlib import Path
import sys
import tempfile
import threading
import types
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import run_in_solidworks as runner


class RunnerTests(unittest.TestCase):
    def setUp(self):
        if runner.pythoncom is None:
            # Linux CI: replace only the absent Windows boundaries. Runner/profile,
            # state files, locking, code preparation and report checks remain real.
            boundaries = {
                'pythoncom': types.SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None,
                    VT_BYREF=0x4000, VT_I4=3, com_error=OSError),
                'w32': types.SimpleNamespace(GetActiveObject=lambda _: None, GetObject=lambda _: None,
                    dynamic=types.SimpleNamespace(Dispatch=lambda x: x),
                    VARIANT=lambda kind, value: types.SimpleNamespace(value=value)),
                'win32gui': types.SimpleNamespace(error=OSError, IsWindowVisible=lambda _: True,
                    IsWindowEnabled=lambda _: True, GetClassName=lambda _: '', GetWindowText=lambda _: '',
                    PostMessage=lambda *args: None),
                'win32process': types.SimpleNamespace(GetWindowThreadProcessId=lambda _: (0, 0)),
                'win32con': types.SimpleNamespace(BM_CLICK=0xF5)}
            for name, value in boundaries.items():
                patcher = patch.object(runner, name, value)
                patcher.start()
                self.addCleanup(patcher.stop)
        self.directory = tempfile.TemporaryDirectory(prefix='sw-runner-test-')
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.macro = self.root / 'noop.bas'
        self.macro.write_text('Option Explicit\nSub main()\nEnd Sub\n')
        self.report = self.root / 'report.json'

    def invoke(self, *args, operation=None, dialog=False, buttons=None, allowed=False, with_profile=True, dependencies=None, existing_docs=None, second_dialog=False, reviewed_module='TestMacro', observe_operation=None, attachment_error=None, run_error=0, api_return=None, mutating=False):
        release = threading.Event()
        finished = threading.Event()
        busy = threading.Event()
        worker_finished = threading.Event()
        self.macro_calls = []
        test_case = self
        source = self.root / 'copy.sldprt'
        source.write_bytes(b'test copy')
        profile = self.root / 'profile.json'
        outputs = [str(self.report)]
        for index, value in enumerate(args):
            if value in ('--expect-text', '--expect-file'):
                outputs.append(args[index + 1])
        profile_data = {
            'schema_version': 1, 'test_root': str(self.root), 'reviewed': True, 'copies_confirmed': True,
            'allow_model_changes': mutating,
            'macro_sha256': hashlib.sha256(self.macro.read_bytes() if self.macro.suffix == '.swp' else self.macro.read_text().encode('utf-8')).hexdigest(),
            'entry_point': {'module': reviewed_module, 'procedure': 'main'},
            'inputs': [{'path': str(source), 'sha256': hashlib.sha256(b'test copy').hexdigest()}],
            'outputs': outputs,
            'allowed_dialogs': [{'title': 'SOLIDWORKS', 'text': 'Overwrite existing output?',
                                 'answer': 'yes'}] if allowed else []}
        if '--observe' in args:
            observer = Path(args[args.index('--observe') + 1])
            profile_data['observer_sha256'] = hashlib.sha256(observer.read_text().encode('utf-8')).hexdigest()
        profile.write_text(json.dumps(profile_data))

        class Document:
            GetPathName = str(source)
            GetTitle = source.name

        class SolidWorks:
            RevisionNumber = 'test boundary'
            ActiveDoc = None

            @property
            def GetDocuments(self):
                if busy.is_set():
                    raise AssertionError('COM queried while RunMacro2 is still running')
                return test_case.documents

            def GetDocumentDependencies2(self, path, traverse, search, readonly):
                test_case.assertEqual((path, traverse, readonly), (str(source), True, False))
                if isinstance(dependencies, BaseException):
                    raise dependencies
                return dependencies or []

            def OpenDoc6(self, *unused):
                return Document()

            def ActivateDoc3(self, *unused):
                return Document()

            def CloseDoc(self, title):
                test_case.closed_titles.append(title)
                if not getattr(test_case, 'refuse_close', False):
                    test_case.documents = [doc for doc in test_case.documents if doc.GetTitle != title]

            def RunMacro2(self, swb, *unused):
                try:
                    test_case.macro_calls.append(Path(swb).stem)
                    test_case.generated_bytes = Path(swb).read_bytes()
                    test_case.entry_point = unused[:2]
                    test_case.generated_code = Path(swb).read_text(encoding='cp1250') if Path(swb).suffix != '.swp' else None
                    current = observe_operation if Path(swb).stem == 'ObserveMacro' else operation
                    if current == 'wait':
                        busy.set()
                        release.wait(5)
                    elif current:
                        current()
                    unused[-1].value = run_error
                    return run_error == 0 if api_return is None else api_return
                finally:
                    busy.clear()
                    finished.set()

        self.documents = list(existing_docs or [])
        self.closed_titles = []
        sw = SolidWorks()
        def start_watcher(watcher):
            if dialog:
                watcher._check(100, None)
                if second_dialog:
                    watcher._check(200, None)

        with contextlib.ExitStack() as stack:
            stack.enter_context(patch.object(sys, 'argv', [
                'run_in_solidworks.py', str(self.macro), '--json', str(self.report),
                *(['--profile', str(profile), '--open', str(source)] if with_profile else []), *args]))
            stack.enter_context(patch.object(runner, 'SESSION_PATH', self.root / 'session.json', create=True))
            stack.enter_context(patch.object(runner, 'LOCK_PATH', self.root / 'session.lock', create=True))
            stack.enter_context(patch.object(runner.w32, 'GetActiveObject', return_value=sw, side_effect=attachment_error))
            stack.enter_context(patch.object(runner.w32.dynamic, 'Dispatch', side_effect=lambda x: x))
            stack.enter_context(patch.object(runner.w32, 'GetObject', return_value=types.SimpleNamespace(
                ExecQuery=lambda _: [types.SimpleNamespace(ProcessId=123)])))
            stack.enter_context(patch.object(runner.pythoncom, 'CoInitialize'))
            def uninitialize():
                if threading.current_thread() is not threading.main_thread():
                    worker_finished.set()
            stack.enter_context(patch.object(runner.pythoncom, 'CoUninitialize', side_effect=uninitialize))
            stack.enter_context(patch.object(runner.DialogWatcher, 'start', start_watcher))
            stack.enter_context(patch.object(runner.DialogWatcher, 'join'))
            stack.enter_context(patch.object(runner.win32gui, 'IsWindowVisible', return_value=True))
            stack.enter_context(patch.object(runner.win32gui, 'IsWindowEnabled', return_value=True))
            stack.enter_context(patch.object(runner.win32gui, 'GetClassName', return_value='#32770'))
            stack.enter_context(patch.object(runner.win32gui, 'GetWindowText', return_value='SOLIDWORKS'))
            stack.enter_context(patch.object(runner.win32process, 'GetWindowThreadProcessId', return_value=(1, 123)))
            stack.enter_context(patch.object(runner, 'window_texts', return_value=(
                ['Overwrite existing output?'], buttons if buttons is not None else [(101, 'Yes'), (102, 'No')])))
            stack.enter_context(patch.object(runner, 'has_child_class', return_value=False))
            clicks = stack.enter_context(patch.object(runner.win32gui, 'PostMessage'))
            output = stack.enter_context(contextlib.redirect_stdout(io.StringIO()))
            try:
                status = runner.main()
                return status, json.loads(output.getvalue()), clicks.call_args_list
            finally:
                release.set()
                if self.macro_calls:
                    worker_finished.wait(2)

    def test_undeclared_document_is_not_closed(self):
        external = types.SimpleNamespace(GetPathName=str(self.root / 'unexpected.sldprt'), GetTitle='unexpected')
        status, report, _ = self.invoke('--close-opened', operation=lambda: self.documents.append(external))
        self.assertEqual(status, 1)
        self.assertEqual(report['failure_code'], 'UNDECLARED_DOCUMENT')
        self.assertEqual(report['documents_left_open'], [external.GetPathName.lower()])
        self.assertEqual(self.closed_titles, [])
        self.assertEqual(report['session_state'], 'UNKNOWN')

    def test_close_must_be_verified(self):
        self.refuse_close = True
        document = types.SimpleNamespace(GetPathName=str(self.root / 'copy.sldprt'), GetTitle='copy')
        status, report, _ = self.invoke('--close-opened', operation=lambda: self.documents.append(document))
        self.assertEqual(status, 1)
        self.assertEqual(report['failure_code'], 'CLEANUP_FAILED')
        self.assertEqual(report['documents_left_open'], [document.GetPathName.lower()])
        self.assertEqual(report['closed_test_documents'], [])
        self.assertEqual(report['session_state'], 'UNKNOWN')

    def test_output_parent_replacement_is_rejected(self):
        directory = self.root / 'output'
        directory.mkdir()
        output = directory / 'result.txt'
        def replace_directory():
            directory.rename(self.root / 'old-output')
            directory.mkdir()
            output.write_text('generated')
        status, report, _ = self.invoke('--expect-file', str(output), operation=replace_directory)
        self.assertEqual(status, 1)
        self.assertTrue(any('identity changed' in failure for failure in report['expectations_failed']))

    def test_in_memory_mutation_permission_never_allows_saving_a_hashed_input(self):
        status, report, _ = self.invoke('--confirm-model-changes', mutating=True,
            operation=lambda: (self.root / 'copy.sldprt').write_bytes(b'unexpected Save3'))
        self.assertEqual(status, 1, report)
        self.assertEqual(report['status'], 'FAIL')
        self.assertTrue(report['expectations_failed'])
        self.assertEqual(report['session_state'], 'IDLE')

    def test_timeout_writes_failure_report_without_querying_busy_com(self):
        status, report, _ = self.invoke('--timeout', '0.01', operation='wait')
        self.assertEqual(status, 1)
        self.assertEqual(report['status'], 'FAIL')
        self.assertTrue(report['timed_out'])
        self.assertIsNone(report['documents_left_open'])
        self.assertEqual(json.loads((self.root / 'session.json').read_text())['state'], 'UNKNOWN')
        self.report.unlink()
        status, report, _ = self.invoke()
        self.assertEqual(status, 2)
        self.assertEqual(report['status'], 'BLOCKED_SESSION')

    def test_profile_is_required_before_connecting_to_solidworks(self):
        status, report, _ = self.invoke(with_profile=False)
        self.assertEqual(status, 2)
        self.assertEqual(report['status'], 'PREFLIGHT_FAILED')

    def test_json_collision_with_expectations_is_rejected_before_execution(self):
        for option in ('--expect-file', '--expect-text'):
            with self.subTest(option=option):
                self.report.write_text('preserve')
                alias = str(self.root / '.' / self.report.name)
                args = [option, alias] + (['needle'] if option == '--expect-text' else [])
                status, report, _ = self.invoke(*args)
                self.assertEqual(status, 2)
                self.assertEqual(report['status'], 'PREFLIGHT_FAILED')
                self.assertFalse(report['macro_executed'])
                self.assertEqual(self.report.read_text(), 'preserve')
                self.assertEqual(self.macro_calls, [])

    def test_existing_report_is_never_overwritten(self):
        self.report.write_text('valuable report')
        status, _, _ = self.invoke()
        self.assertEqual(status, 2)
        self.assertEqual(self.report.read_text(), 'valuable report')
        self.assertEqual(self.macro_calls, [])

    def test_unchanged_recent_output_is_not_fresh(self):
        output = self.root / 'output.step'
        output.write_text('old')
        status, report, _ = self.invoke('--expect-file', str(output))
        self.assertEqual(status, 1)
        self.assertIn('unchanged file', report['expectations_failed'][0])

    def test_changed_output_is_fresh_even_with_old_timestamp(self):
        import os
        output = self.root / 'output.step'
        output.write_text('old')
        old_time = output.stat().st_mtime_ns
        def update():
            output.write_text('new')
            os.utime(output, ns=(old_time, old_time))
        status, report, _ = self.invoke('--expect-file', str(output), operation=update)
        self.assertEqual(status, 0, report)

    def test_temp_directory_removed_only_for_resolved_session(self):
        real_mkdtemp = tempfile.mkdtemp
        created = []
        def make_temp(*args, **kwargs):
            path = real_mkdtemp(*args, dir=self.root, **kwargs)
            created.append(Path(path))
            return path
        with patch.object(runner.tempfile, 'mkdtemp', side_effect=make_temp):
            status, _, _ = self.invoke()
            self.assertEqual(status, 0)
            self.assertFalse(created[-1].exists())
            self.report.unlink()
            status, _, _ = self.invoke('--timeout', '0.01', operation='wait')
            self.assertEqual(status, 1)
            self.assertTrue((created[-1] / 'TestMacro.swb').exists())

    def test_native_project_is_run_without_text_conversion(self):
        self.macro = self.root / 'complete-project.swp'
        self.macro.write_bytes(b'\xd0\xcf\x11\xe0\x00\x81binary-project')
        status, report, _ = self.invoke('--module', 'FixtureModule', '--entry', 'main', reviewed_module='FixtureModule')
        self.assertEqual(status, 0, report)
        self.assertEqual(self.generated_bytes, self.macro.read_bytes())
        self.assertEqual(self.entry_point, ('FixtureModule', 'main'))
        self.assertEqual(report['project_format'], 'swp')
        self.assertEqual(report['compilation_status'], 'NOT_TESTED')

    def test_different_procedure_is_rejected_after_profile_review(self):
        status, report, _ = self.invoke('--entry', 'OtherProcedure')
        self.assertEqual(status, 2)
        self.assertFalse(hasattr(self, 'generated_bytes'))

    def test_observer_runs_before_cleanup_and_checks_its_output(self):
        observer = self.root / 'observer.bas'
        observer.write_text('Option Explicit\nSub main()\nEnd Sub\n')
        output = self.root / 'observations.json'
        status, report, _ = self.invoke('--observe', str(observer), '--expect-file', str(output),
            observe_operation=lambda: output.write_text('{"solid_count":1}'))
        self.assertEqual(status, 0, report)
        self.assertTrue(report['observation_ok'])
        self.assertEqual(self.entry_point, ('ObserveMacro', 'main'))

    def test_observer_timeout_keeps_session_blocked(self):
        observer = self.root / 'observer.bas'
        observer.write_text('Option Explicit\nSub main()\nEnd Sub\n')
        status, report, _ = self.invoke('--observe', str(observer), '--timeout', '0.01', observe_operation='wait')
        self.assertEqual(status, 1)
        self.assertTrue(report['timed_out'])
        self.assertIsNone(report['documents_left_open'])
        self.assertEqual(json.loads((self.root / 'session.json').read_text())['state'], 'UNKNOWN')

    def test_external_dependency_is_rejected_before_macro_runs(self):
        status, report, _ = self.invoke(dependencies=['production', str(self.root.parent / 'production.sldprt')])
        self.assertEqual(status, 2)
        self.assertFalse(hasattr(self, 'generated_code'))

    def test_com_exception_during_preflight_leaves_unknown_session(self):
        error = runner.pythoncom.com_error(-2147418111, 'COM call rejected', None, None)
        status, report, _ = self.invoke(dependencies=error)
        self.assertEqual(report['status'], 'EXECUTION_UNKNOWN')
        self.assertEqual(json.loads((self.root / 'session.json').read_text())['state'], 'UNKNOWN')

    def test_com_attachment_exception_is_not_reclassified_as_a_safe_validation_failure(self):
        error = runner.pythoncom.com_error(-2147418111, 'COM call rejected', None, None)
        status, report, _ = self.invoke(attachment_error=error)
        self.assertEqual(report['status'], 'EXECUTION_UNKNOWN')
        self.assertEqual(json.loads((self.root / 'session.json').read_text())['state'], 'UNKNOWN')

    def test_busy_run_error_does_not_query_documents_or_clear_session(self):
        status, report, _ = self.invoke(run_error=17)
        self.assertEqual(status, 1)
        self.assertEqual(report['run_error'], 'Busy')
        self.assertIsNone(report['documents_left_open'])
        self.assertEqual(json.loads((self.root / 'session.json').read_text())['state'], 'UNKNOWN')

    def test_nonzero_run_error_cannot_pass_even_with_true_return(self):
        status, report, _ = self.invoke(run_error=22, api_return=True)
        self.assertEqual(status, 1)
        self.assertEqual(report['run_error'], 'Invalidindex')

    def test_primary_macro_cannot_start_observer_after_timeout(self):
        observer = self.root / 'observer.bas'
        observer.write_text('Option Explicit\nSub main()\nEnd Sub\n')
        status, report, _ = self.invoke('--observe', str(observer), '--timeout', '0.01', operation='wait')
        self.assertEqual(status, 1)
        self.assertTrue(report['timed_out'])
        self.assertEqual(self.macro_calls, ['TestMacro'])

    def test_production_document_in_session_is_rejected(self):
        doc = types.SimpleNamespace(GetPathName=str(self.root.parent / 'production.sldprt'), GetTitle='production')
        status, report, _ = self.invoke(existing_docs=[doc])
        self.assertEqual(status, 2)
        self.assertFalse(hasattr(self, 'generated_code'))

    def test_default_leaves_confirmation_unanswered(self):
        status, report, clicks = self.invoke(dialog=True)
        self.assertEqual(status, 1)
        self.assertEqual(report['status'], 'BLOCKED_DIALOG')
        self.assertEqual(clicks, [])

    def test_answer_without_matching_allowlist_leaves_confirmation_unanswered(self):
        status, report, clicks = self.invoke('--answer', 'yes', dialog=True)
        self.assertEqual(report['status'], 'BLOCKED_DIALOG')
        self.assertEqual(clicks, [])

    def test_blocked_dialog_returns_without_waiting_for_macro_completion(self):
        status, report, _ = self.invoke('--strict-dialogs', '--timeout', '0.05', operation='wait', dialog=True)
        self.assertEqual(status, 1)
        self.assertEqual(report['status'], 'BLOCKED_DIALOG')
        self.assertFalse(report['timed_out'])
        self.assertIsNone(report['documents_left_open'])

    def test_explicit_yes_answers_the_matching_button(self):
        status, report, clicks = self.invoke('--answer', 'yes', dialog=True, allowed=True)
        self.assertEqual(status, 0)
        self.assertEqual(report['status'], 'PASS')
        self.assertEqual(clicks[0].args[:2], (101, runner.win32con.BM_CLICK))

    def test_missing_requested_button_blocks_instead_of_clicking_ok(self):
        status, report, clicks = self.invoke('--answer', 'yes', dialog=True, allowed=True, buttons=[(101, 'OK')])
        self.assertEqual(status, 1)
        self.assertEqual(report['status'], 'BLOCKED_DIALOG')
        self.assertEqual(clicks, [])

    def test_strict_mode_overrides_explicit_answer(self):
        status, report, clicks = self.invoke('--answer', 'yes', '--strict-dialogs', dialog=True)
        self.assertEqual(report['status'], 'BLOCKED_DIALOG')
        self.assertEqual(clicks, [])

    def test_allowlist_answers_at_most_one_dialog_per_run(self):
        status, report, clicks = self.invoke('--answer', 'yes', allowed=True, dialog=True, second_dialog=True)
        self.assertEqual(report['status'], 'BLOCKED_DIALOG')
        self.assertEqual(len(clicks), 1)

    def test_explicit_answer_preserves_native_message_calls_and_string_literals(self):
        code = ('Option Explicit\nSub main()\n'
                'swApp.SendMsgToUser2 "swApp.SendMsgToUser2", 2, 5\nEnd Sub\n')
        self.macro.write_text(code)
        self.invoke('--answer', 'ok')
        self.assertEqual(self.generated_code, code)

    def check_log(self, before, after, expected_status, encoding='utf-8'):
        log = self.root / 'macro.log'
        if before is not None:
            log.write_text(before, encoding=encoding)
        def operation():
            if after is not None:
                log.write_text(after, encoding=encoding)
        status, report, _ = self.invoke('--expect-text', str(log), 'OK', operation=operation)
        self.assertEqual(report['status'], expected_status)
        self.assertEqual(status, 0 if expected_status == 'PASS' else 1)

    def test_old_success_does_not_satisfy_expectation(self):
        self.check_log('OK old run\n', None, 'FAIL')

    def test_append_failure_does_not_reuse_old_success(self):
        self.check_log('OK old run\n', 'OK old run\nERROR new run\n', 'FAIL')

    def test_appended_success_satisfies_expectation(self):
        self.check_log('ERROR old run\n', 'ERROR old run\nOK new run\n', 'PASS')

    def test_replaced_log_is_checked(self):
        self.check_log('ERROR old run\n', 'OK new run\n', 'PASS')

    def test_new_log_is_checked(self):
        self.check_log(None, 'OK new run\n', 'PASS')

    def test_utf16_append_failure_does_not_reuse_old_success(self):
        self.check_log('OK old run\n', 'OK old run\nERROR new run\n', 'FAIL', encoding='utf-16')

    def test_utf16_appended_success_is_checked(self):
        self.check_log('ERROR old run\n', 'ERROR old run\nOK new run\n', 'PASS', encoding='utf-16')


if __name__ == '__main__':
    unittest.main()
