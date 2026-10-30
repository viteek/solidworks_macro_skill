"""Run the real CLI flow with COM and Windows UI replaced at their boundaries."""
import contextlib
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
        self.directory = tempfile.TemporaryDirectory(prefix='sw-runner-test-')
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.macro = self.root / 'noop.bas'
        self.macro.write_text('Option Explicit\nSub main()\nEnd Sub\n')
        self.report = self.root / 'report.json'

    def invoke(self, *args, operation=None, dialog=False, buttons=None):
        release = threading.Event()
        finished = threading.Event()
        busy = threading.Event()
        test_case = self

        class SolidWorks:
            RevisionNumber = 'test boundary'
            ActiveDoc = None

            @property
            def GetDocuments(self):
                if busy.is_set():
                    raise AssertionError('COM queried while RunMacro2 is still running')
                return []

            def RunMacro2(self, swb, *unused):
                try:
                    test_case.generated_code = Path(swb).read_text(encoding='cp1250')
                    if operation == 'wait':
                        busy.set()
                        release.wait(5)
                    elif operation:
                        operation()
                    return True
                finally:
                    busy.clear()
                    finished.set()

        sw = SolidWorks()
        def start_watcher(watcher):
            if dialog:
                watcher._check(100, None)

        with contextlib.ExitStack() as stack:
            stack.enter_context(patch.object(sys, 'argv', [
                'run_in_solidworks.py', str(self.macro), '--json', str(self.report), *args]))
            stack.enter_context(patch.object(runner.w32, 'GetActiveObject', return_value=sw))
            stack.enter_context(patch.object(runner.w32.dynamic, 'Dispatch', side_effect=lambda x: x))
            stack.enter_context(patch.object(runner.w32, 'GetObject', return_value=types.SimpleNamespace(
                ExecQuery=lambda _: [types.SimpleNamespace(ProcessId=123)])))
            stack.enter_context(patch.object(runner.pythoncom, 'CoInitialize'))
            stack.enter_context(patch.object(runner.pythoncom, 'CoUninitialize'))
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
            stack.enter_context(contextlib.redirect_stdout(io.StringIO()))
            try:
                status = runner.main()
                return status, json.loads(self.report.read_text()), clicks.call_args_list
            finally:
                release.set()
                finished.wait(2)

    def test_timeout_writes_failure_report_without_querying_busy_com(self):
        status, report, _ = self.invoke('--timeout', '0.01', operation='wait')
        self.assertEqual(status, 1)
        self.assertEqual(report['status'], 'FAIL')
        self.assertTrue(report['timed_out'])
        self.assertIsNone(report['documents_left_open'])

    def test_default_leaves_confirmation_unanswered(self):
        status, report, clicks = self.invoke(dialog=True)
        self.assertEqual(status, 1)
        self.assertEqual(report['status'], 'BLOCKED_DIALOG')
        self.assertEqual(clicks, [])

    def test_blocked_dialog_returns_without_waiting_for_macro_completion(self):
        status, report, _ = self.invoke('--strict-dialogs', '--timeout', '0.05', operation='wait', dialog=True)
        self.assertEqual(status, 1)
        self.assertEqual(report['status'], 'BLOCKED_DIALOG')
        self.assertFalse(report['timed_out'])
        self.assertIsNone(report['documents_left_open'])

    def test_explicit_yes_answers_the_matching_button(self):
        status, report, clicks = self.invoke('--answer', 'yes', dialog=True)
        self.assertEqual(status, 0)
        self.assertEqual(report['status'], 'PASS')
        self.assertEqual(clicks[0].args[:2], (101, runner.win32con.BM_CLICK))

    def test_missing_requested_button_blocks_instead_of_clicking_ok(self):
        status, report, clicks = self.invoke('--answer', 'no', dialog=True, buttons=[(101, 'Yes')])
        self.assertEqual(status, 1)
        self.assertEqual(report['status'], 'BLOCKED_DIALOG')
        self.assertEqual(clicks, [])

    def test_strict_mode_overrides_explicit_answer(self):
        status, report, clicks = self.invoke('--answer', 'yes', '--strict-dialogs', dialog=True)
        self.assertEqual(report['status'], 'BLOCKED_DIALOG')
        self.assertEqual(clicks, [])

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
