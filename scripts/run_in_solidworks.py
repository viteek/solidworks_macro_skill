#!/usr/bin/env python3
r"""Run a VBA macro (.bas/.swb/.txt) in the user's running SOLIDWORKS and report what happened.

Lets you *test* a macro you wrote instead of handing it over untested. Windows only;
needs Python 3 + pywin32 (pip install pywin32) and a running SOLIDWORKS session.

How it works
- The code is saved as a temporary text macro (.swb, cp1250, CRLF) - SOLIDWORKS runs .swb files
  with the SOLIDWORKS type libraries available, so early-bound code and sw* constants work.
- A watcher records recognized Windows dialogs. All are unanswered by default. --answer
  requires exact title/text/answer in the reviewed profile, with at most one response.
  Unknown dialogs stop automation; error dialogs are classified, not automatically dismissed.
  Use --define or --pick-folder to avoid interactive folder pickers.
  --strict-dialogs leaves every dialog unanswered even with --answer (status BLOCKED_DIALOG).
  SendMsgToUser2 calls are never rewritten. Opaque/custom dialogs that cannot be recognized
  may reach the timeout; inspect SOLIDWORKS manually in that case.
- Requires --open on a hashed test copy and --profile on reviewed code, entry point, inputs,
  dependency closure and outputs. SOLIDWORKS must have no documents open. --close-opened closes
  only new saved documents inside the test root. Profiles are preflight, not a VBA sandbox.
- --expect-file checks output freshness; --expect-text checks only appended text, or the contents
  of a new log or a log whose previous byte prefix changed. Identical content is conservatively
  treated as unchanged, even after replacement. Prefer an appended per-run result or run ID.
  --expect-result requires the final new line to start with an exact status prefix; a log whose
  final line is any status other than a known success never satisfies --expect-text either.
- Timeout/blocked dialog: report without further COM calls or document cleanup. The macro may
  still be running in SOLIDWORKS; the timeout does not cancel it. A durable UNKNOWN/RUNNING
  marker blocks further runs until explicit --recover-session after inspection.

Usage
  python run_in_solidworks.py macro.bas --open C:\SWTest\part.sldprt --profile C:\SWTest\profile.json
         [--define "OUT_DIR=\"C:\\SWTest\\out\""] [--answer yes|no|ok|cancel] [--timeout 300]
         [--close-opened] [--expect-file C:\SWTest\out\part.dxf] [--json report.json]

--define NAME=VALUE rewrites the line `Const NAME [As T] = ...` to the given VBA literal, so a macro
whose folder/option is a Const (with a picker only when the Const is empty) can run unattended.

Exit code 0 = macro ran with no compile/run-time error dialog, 1 = error, 2 = could not start.
Never run this against the user's production files: point macros at copies.
"""
import argparse
import json
import math
import os
from pathlib import Path
import re
import shutil
import sys
import tempfile
import threading
import time

sys.path.insert(0, str(Path(__file__).resolve().parent))  # sibling imports under python -I
from runner_safety import (TestProfile, SessionState, canonical, inside, code_hash, file_hash,
                           output_identity, revalidate_output)

STATE_DIR = Path(os.environ.get('LOCALAPPDATA', tempfile.gettempdir())) / 'solidworks-macro-runner'
SESSION_PATH = STATE_DIR / 'session.json'
LOCK_PATH = STATE_DIR / 'controller.lock'

try:
    import pythoncom
    import win32api
    import win32con
    import win32gui
    import win32process
    import win32com.client as w32
except ImportError:
    pythoncom = win32api = win32con = win32gui = win32process = w32 = None


def process_alive(pid):
    """OS-only probe; never query a potentially busy SOLIDWORKS COM server."""
    if pid is None or sys.platform != 'win32':
        return None
    import ctypes
    from ctypes import wintypes
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.OpenProcess.argtypes = (wintypes.DWORD, wintypes.BOOL, wintypes.DWORD)
    kernel.OpenProcess.restype = wintypes.HANDLE
    kernel.CloseHandle.argtypes = (wintypes.HANDLE,)
    handle = kernel.OpenProcess(0x1000, False, int(pid))  # PROCESS_QUERY_LIMITED_INFORMATION
    if not handle:
        return False if ctypes.get_last_error() == 87 else None  # invalid PID vs access denied
    kernel.CloseHandle(handle)
    return True  # PID reuse is treated conservatively as alive

RUN_ERRORS = {
    0: "OK", 1: "InvalidArg", 2: "MacrosAreDisabled", 3: "NotInDesignMode", 4: "OnlyCodeModules",
    5: "OutOfMemory", 6: "InvalidProcname", 7: "InvalidPropertyType", 8: "SuborfuncExpected",
    9: "BadParmCount", 10: "BadVarType", 11: "UserInterrupt", 12: "Exception",
    13: "Overflow", 14: "TypeMismatch", 15: "ParmNotOptional", 16: "UnknownLcid",
    17: "Busy", 18: "ConnectionTerminated", 19: "CallRejected", 20: "CallFailed",
    21: "Zombied", 22: "Invalidindex", 23: "NoPermission", 24: "Reverted",
    25: "TooManyOpenFiles", 26: "DiskError", 27: "CantSave", 28: "OpenFileFailed",
}
# Confirmed against the installed swRunMacroError_e. A rejected/busy COM call is
# unresolved even if Python's worker thread has returned normally.
UNCERTAIN_RUN_ERRORS = {12, 17, 18, 19, 20, 21}
END_BUTTONS = ("&End", "End", "&Zakończ", "Zakończ")
ANSWERS = {
    "ok": ("OK", "&OK"),
    "yes": ("&Yes", "Yes", "&Tak", "Tak"),
    "no": ("&No", "No", "&Nie", "Nie"),
    "cancel": ("Cancel", "Anuluj"),
}


def window_texts(hwnd):
    texts, buttons = [], []

    def cb(child, _):
        cls = win32gui.GetClassName(child)
        txt = win32gui.GetWindowText(child)
        if cls == "Button":
            buttons.append((child, txt))
        elif txt and cls in ("Static", "Edit", "RICHEDIT50W", "DirectUIHWND"):
            texts.append(txt)
        return True

    try:
        win32gui.EnumChildWindows(hwnd, cb, None)
    except win32gui.error:
        pass
    if not texts:   # SOLIDWORKS-styled message boxes use other control classes
        def any_text(child, _):
            t = win32gui.GetWindowText(child)
            if t and win32gui.GetClassName(child) != "Button":
                texts.append(t)
            return True
        try:
            win32gui.EnumChildWindows(hwnd, any_text, None)
        except win32gui.error:
            pass
    return texts, buttons


def has_child_class(hwnd, name):
    found = []

    def cb(child, _):
        if win32gui.GetClassName(child) == name:
            found.append(child)
        return True

    try:
        win32gui.EnumChildWindows(hwnd, cb, None)
    except win32gui.error:
        pass
    return bool(found)


class DialogWatcher(threading.Thread):
    def __init__(self, pid, answer=None, strict=False, allowed_dialogs=()):
        super().__init__(daemon=True)
        self.pid, self.answer, self.strict = pid, answer, strict or answer is None
        self.blocked = False
        self.stop = threading.Event()
        self.dialogs = []
        self.compile_error = False
        self.runtime_error = False
        self._seen = set()
        self.allowed_dialogs = allowed_dialogs
        self._answered = False

    def run(self):
        while not self.stop.is_set():
            try:
                win32gui.EnumWindows(self._check, None)
            except Exception as error:
                self.blocked = True
                self.dialogs.append({'kind': 'watcher_error', 'text': str(error), 'pressed': None})
                return
            time.sleep(0.3)

    def _click(self, hwnd, buttons, names):
        for name in names:
            for child, txt in buttons:
                if txt.replace("&", "") == name.replace("&", "") and win32gui.IsWindowEnabled(child):
                    win32gui.PostMessage(child, win32con.BM_CLICK, 0, 0)
                    return txt
        self.blocked = True
        return None   # never substitute another button for the requested answer

    def _check(self, hwnd, _):
        try:
            if self.blocked:
                return True
            if not win32gui.IsWindowVisible(hwnd):
                return True
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid != self.pid:
                return True
            cls = win32gui.GetClassName(hwnd)
            title = win32gui.GetWindowText(hwnd)
            if cls != "#32770":
                return True
            texts, buttons = window_texts(hwnd)
            key = (hwnd, title, tuple(texts))
            if key in self._seen:
                return True
            self._seen.add(key)
            body = " | ".join(texts)
            kind = "message"
            if has_child_class(hwnd, "SysTreeView32"):
                kind = "folder_picker"
            elif re.search(r"Run-time error|Błąd czasu wykonania|Laufzeitfehler", body, re.I):
                kind, self.runtime_error = "runtime_error", True
            elif re.search(r"Compile error|Błąd kompilacji|Syntax error|Błąd składni", body, re.I):
                kind, self.compile_error = "compile_error", True
            allowed = any(item == {'title': title, 'text': body, 'answer': self.answer}
                          for item in self.allowed_dialogs)
            if self.strict or self._answered or not allowed:
                kind, self.blocked, pressed = "blocked_" + kind, True, None
            else:
                pressed = self._click(hwnd, buttons, ANSWERS[self.answer])
                self._answered = pressed is not None
            self.dialogs.append({"kind": kind, "title": title, "text": body, "pressed": pressed,
                                 "t": round(time.time(), 1)})
        except win32gui.error:
            pass
        return True


def prepare_code(src_text, defines, pick_folders=()):
    lines = [l for l in src_text.replace("\r\n", "\n").split("\n") if not l.startswith("Attribute VB_")]
    code = "\n".join(lines)
    if pick_folders:
        # Shell.Application.BrowseForFolder(...) -> .NameSpace(<next given folder>) (same Folder object type)
        code, n = re.subn(r"\.BrowseForFolder\s*\((?:[^()]|\([^()]*\))*\)", ".NameSpace(HarnessNextFolder())", code)
        if n == 0:
            print("WARN  --pick-folder given but no .BrowseForFolder( call found")
        items = ", ".join('"' + p.replace('"', '""') + '"' for p in pick_folders)
        code += ("\n\nPrivate Function HarnessNextFolder() As Variant\n"
                 "    Static i As Long\n"
                 f"    Dim v As Variant: v = Array({items})\n"
                 "    If i > UBound(v) Then i = UBound(v)\n"
                 "    HarnessNextFolder = v(i): i = i + 1\n"
                 "End Function\n")
    for d in defines:
        name, _, value = d.partition("=")
        pat = re.compile(r"^(\s*(?:Private\s+|Public\s+)?Const\s+" + re.escape(name.strip()) +
                         r"\b(?:\s+As\s+\w+)?\s*=\s*)([^'\n]*)", re.I | re.M)
        code, n = pat.subn(lambda m: m.group(1) + value, code)
        if n == 0:
            print(f"WARN  --define {name}: no 'Const {name} = ...' line found")
    return code


def find_entry(code):
    subs = re.findall(r"^\s*(?:Public\s+)?Sub\s+(\w+)\s*\(\s*\)", code, re.I | re.M)
    return "main" if any(s.lower() == "main" for s in subs) else (subs[0] if subs else "main")


def doc_type(path):
    return {".sldprt": 1, ".sldasm": 2, ".slddrw": 3}.get(os.path.splitext(path)[1].lower(), 0)


def open_docs(sw):
    docs = sw.GetDocuments
    return {d.GetPathName.lower(): d.GetTitle for d in (docs or [])}


def read_log(path):
    try:
        with open(path, "rb") as stream:
            return stream.read()
    except FileNotFoundError:
        return b""


def new_log_text(previous, current):
    """Read appended bytes only; a truncated/replaced log is a new document."""
    appended = current.startswith(previous)
    raw = current[len(previous):] if appended else current
    if current.startswith((b'\xff\xfe', b'\xfe\xff')):
        encoding = 'utf-16-le' if current.startswith(b'\xff\xfe') else 'utf-16-be'
        return raw.decode(encoding, errors='replace').lstrip('\ufeff')
    try:
        return raw.decode('utf-8-sig')
    except UnicodeDecodeError:
        return raw.decode('cp1250', errors='replace')


# A status-shaped final line ("RESULT name STATUS", "RESULT=STATUS", "STATUS;...") is accepted only
# with a known success status, so new negative states (ERROR_CLEANUP, *_FAILED, START...) fail closed.
SUCCESS_STATUSES = {'OK', 'PASS', 'PASSED', 'SUCCESS', 'DONE', 'EXPORTED', 'OBSERVED', 'SAVED', 'AUDIT', 'CREATED'}
# Result line format: "RESULT=STATUS", "RESULT STATUS" or "RESULT NAME STATUS", where NAME is
# [A-Za-z0-9_.-]+, STATUS is upper case and ends at ';', ':', whitespace or end of line. Anything
# else that starts with RESULT falls through to LEADING_STATUS ("RESULT"), which is not a success.
RESULT_STATUS = re.compile(r'RESULT(?:=(?P<bare>[A-Z][A-Z_]*)|\s+(?:(?P<name>[A-Za-z0-9_.-]+)\s+)?'
                           r'(?P<named>[A-Z][A-Z_]*))(?=[;:\s]|$)')
LEADING_STATUS = re.compile(r'(?P<bare>[A-Z][A-Z_]+)(?=[;:= ]|$)')
NEGATIVE_WORD = re.compile(r'(?i)(?:fail\w*|\w*_failed|error\w*|cancel\w*|abort\w*|timeout|blocked|unknown'
                           r'|invalid\w*|rejected)\b')


def terminal_line(text):
    """The last non-empty line carries the run's final status."""
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    return lines[-1] if lines else ''


def final_status_problem(line):
    match = RESULT_STATUS.match(line) or LEADING_STATUS.match(line)
    if match:
        status = match.group('bare') or match.groupdict().get('named')
        if status not in SUCCESS_STATUSES:
            return f"final status {status!r} is not a recognized success status"
        if NEGATIVE_WORD.fullmatch(match.groupdict().get('name') or ''):
            return f"result name {match.group('name')!r} is a failure word"
    elif NEGATIVE_WORD.match(line):
        return "final line reports failure"
    return None


def check_log_text(content, needle):
    """Substring check that never accepts a log whose final line is not a success status."""
    if needle not in content:
        return f"text {needle!r} not found"
    problem = final_status_problem(terminal_line(content))
    if problem:
        return f"text {needle!r} found, but {problem}"
    return None


def check_log_result(content, prefix):
    """The final line must start with PREFIX followed by end of line, ';' or whitespace."""
    line = terminal_line(content)
    rest = line[len(prefix):]
    if not line.startswith(prefix) or (rest and rest[0] != ';' and not rest[0].isspace()):
        return f"final status line {line!r} does not start with {prefix!r}"
    return None


def file_snapshot(path):
    """Capture content and identity, rather than relying on a wall-clock threshold."""
    try:
        with open(path, 'rb') as stream:
            stat = os.fstat(stream.fileno())
            import hashlib
            digest = hashlib.sha256()
            for chunk in iter(lambda: stream.read(1024 * 1024), b''):
                digest.update(chunk)
            return (stat.st_dev, stat.st_ino, stat.st_size, stat.st_mtime_ns, digest.digest())
    except FileNotFoundError:
        return None


def write_report(path, content):
    with open(path, 'x', encoding='utf-8') as stream:
        stream.write(content)


def cleanup_temp(a):
    if getattr(a, 'temp_dir', None):
        try:
            shutil.rmtree(a.temp_dir)
        except OSError as error:
            print(f'WARN  Could not remove temporary directory {a.temp_dir}: {error}', file=sys.stderr)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("macro", nargs='?')
    ap.add_argument('--profile', help='reviewed JSON test profile (required for execution)')
    ap.add_argument('--create-profile', action='store_true', help='write an UNREVIEWED profile and exit without COM')
    ap.add_argument('--allow-model-changes', action='store_true', help='declare CAD modifications when creating a profile')
    ap.add_argument('--confirm-model-changes', action='store_true', help='explicitly authorize the reviewed changes to test copies')
    ap.add_argument('--recover-session', action='store_true', help='clear unresolved execution state after inspection')
    ap.add_argument('--confirm-macro-stopped', action='store_true', help='confirm manual verification that the old VBA macro has ended')
    ap.add_argument("--open", help="document to open and activate before running")
    ap.add_argument("--define", action="append", default=[], help="NAME=VBA_LITERAL, overrides a Const")
    ap.add_argument("--answer", choices=list(ANSWERS),
                    help="explicitly allow this dialog answer (default: leave message boxes unanswered)")
    ap.add_argument("--pick-folder", action="append", default=[],
                    help="answer Shell BrowseForFolder pickers with these folders, in order")
    ap.add_argument("--entry", help="procedure to run (default main)")
    ap.add_argument('--module', help='module name (required with --entry for a native .swp project)')
    ap.add_argument('--observe', help='reviewed VBA observer run after successful macro, before cleanup')
    ap.add_argument('--observe-define', action='append', default=[], help='NAME=VALUE for the observer only')
    ap.add_argument("--pid", type=int, help="SOLIDWORKS process to use when several are running")
    ap.add_argument("--timeout", type=float, default=600)
    ap.add_argument("--close-opened", action="store_true",
                    help="close documents that are new after the run AND lie inside the test root (unsaved changes "
                         "of those test copies are discarded); documents elsewhere are never touched")
    ap.add_argument("--test-root", help="profile creation root, or exact confirmation of the profile root")
    ap.add_argument("--strict-dialogs", action="store_true",
                    help="never answer any dialog; report BLOCKED_DIALOG instead")
    ap.add_argument("--expect-file", action="append", default=[],
                    help="file that must exist and be written during this run (repeatable)")
    ap.add_argument("--expect-text", nargs=2, action="append", default=[], metavar=("FILE", "TEXT"),
                    help="newly appended/replaced text in FILE must contain TEXT and must not end in a "
                         "non-success status line such as FAIL or ERROR_CLEANUP (repeatable)")
    ap.add_argument("--expect-result", nargs=2, action="append", default=[], metavar=("FILE", "PREFIX"),
                    help="the last non-empty newly written line in FILE must start with PREFIX, followed by "
                         "end of line, ';' or whitespace (repeatable; preferred acceptance check)")
    ap.add_argument("--json", help="write the report here")
    a = ap.parse_args()
    a.report_authorized = False

    def failure(status, error):
        report = {'status': status, 'error': str(error), 'macro_executed': bool(getattr(a, 'worker_started', False))}
        out = json.dumps(report, ensure_ascii=False, indent=2)
        print(out)
        if a.json and a.report_authorized and inside(a.json, a.test_profile.root):
            try:
                revalidate_output(a.json, a.test_profile.data['test_root'], a.output_identities[a.json])
                write_report(a.json, out)
            except (OSError, ValueError) as report_error:
                print(f'WARN  Could not create report: {report_error}', file=sys.stderr)
        return 2

    if not math.isfinite(a.timeout) or a.timeout <= 0:
        return failure('PREFLIGHT_FAILED', '--timeout must be finite and greater than zero')
    state = SessionState(SESSION_PATH)
    try:
        if not a.recover_session:
            if not a.macro or not a.profile or not a.open:
                raise ValueError('Execution requires macro, --profile and --open (test copy)')
            raw = Path(a.macro).read_bytes()
            a.native_project = Path(a.macro).suffix.lower() == '.swp'
            if a.native_project:
                if not a.entry or not a.module or a.define or a.pick_folder:
                    raise ValueError('A native .swp needs --module and --entry; source rewriting is unavailable')
                a.prepared_code = raw
            else:
                try:
                    text = raw.decode('utf-8-sig')
                except UnicodeDecodeError:
                    text = raw.decode('cp1250')
                a.prepared_code = prepare_code(text, a.define, a.pick_folder)
            a.entry = a.entry or find_entry(a.prepared_code)
            a.module = a.module or 'TestMacro'
            a.observer_code = None
            if a.observe:
                if Path(a.observe).suffix.lower() == '.swp':
                    raise ValueError('--observe accepts source modules, not binary projects')
                observer_raw = Path(a.observe).read_bytes()
                try:
                    observer_text = observer_raw.decode('utf-8-sig')
                except UnicodeDecodeError:
                    observer_text = observer_raw.decode('cp1250')
                a.observer_code = prepare_code(observer_text, a.observe_define)
                if find_entry(a.observer_code).lower() != 'main':
                    raise ValueError('Observer must expose Sub main()')
            elif a.observe_define:
                raise ValueError('--observe-define requires --observe')
            expected_outputs = a.expect_file + [p for p, _ in a.expect_text + a.expect_result]
            if a.json:
                for path in expected_outputs:
                    if (canonical(a.json) == canonical(path)
                            or (Path(a.json).exists() and Path(path).exists()
                                and os.path.samefile(a.json, path))):
                        raise ValueError('--json must be distinct from --expect-file and --expect-text')
            outputs = expected_outputs + ([a.json] if a.json else [])
            if a.create_profile:
                root = Path(a.test_root or Path(a.open).parent).resolve(strict=True)
                data = {'schema_version': 1, 'test_root': str(root), 'reviewed': False,
                        'copies_confirmed': False, 'allow_model_changes': a.allow_model_changes,
                        'macro_sha256': code_hash(a.prepared_code),
                        'entry_point': {'module': a.module, 'procedure': a.entry},
                        'inputs': [{'path': str(Path(a.open).resolve(strict=True)), 'sha256': file_hash(a.open)}],
                        'outputs': [str(Path(p).resolve()) for p in outputs], 'allowed_dialogs': []}
                if a.observer_code is not None:
                    data['observer_sha256'] = code_hash(a.observer_code)
                # Exclusive create: do not overwrite an earlier reviewed profile.
                with Path(a.profile).open('x', encoding='utf-8') as stream:
                    json.dump(data, stream, indent=2)
                print('Created unreviewed profile. Review code, paths and dependencies; confirm copies before execution.')
                return 0
            a.test_profile = TestProfile.load(a.profile, a.prepared_code)
            a.test_profile.validate_entry(a.module, a.entry)
            expected_observer = code_hash(a.observer_code) if a.observer_code is not None else None
            if a.test_profile.data.get('observer_sha256') != expected_observer:
                raise ValueError('Observer changed or differs from the reviewed profile')
            a.test_profile.validate_request(a.open, outputs, a.pick_folder, a.confirm_model_changes)
            a.output_identities = {path: output_identity(path, a.test_profile.data['test_root']) for path in outputs}
            if a.json:
                protected = {canonical(path) for path in (a.macro, a.profile, a.open, a.observe) if path}
                if Path(a.json).suffix.lower() != '.json' or canonical(a.json) in protected:
                    raise ValueError('--json must name a distinct declared .json output, never macro/profile/input')
                a.report_authorized = True
                if os.path.lexists(a.json):
                    raise ValueError('--json report already exists; choose a new path')
            if a.test_root and canonical(a.test_root) != canonical(a.test_profile.root):
                raise ValueError('--test-root disagrees with the reviewed profile')
            if pythoncom is None:
                raise ValueError('Windows + pywin32 are required for live execution')
        Path(LOCK_PATH).parent.mkdir(parents=True, exist_ok=True)
        with Path(LOCK_PATH).open('a+b') as lock:
            if lock.tell() == 0:
                lock.write(b'0')
                lock.flush()
            lock.seek(0)  # every controller MUST lock the same byte
            try:
                if sys.platform == 'win32':
                    import msvcrt
                    msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(lock.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            except OSError as error:
                return failure('BLOCKED_SESSION', f'Another controller holds the execution lock: {error}')
            if a.recover_session:
                state.recover(a.confirm_macro_stopped, process_alive(state.read().get('solidworks_pid')))
                print('Session recovered. No SOLIDWORKS COM calls were made.')
                return 0
            try:
                state.ensure_ready()
            except RuntimeError as error:
                return failure('BLOCKED_SESSION', error)
            state.start(a.macro)
            a.worker_started, a.session_resolved = False, False
            try:
                status = execute_macro(a, str(a.test_profile.root), canonical, state)
                if a.session_resolved or (not a.worker_started and not getattr(a, 'inventory_dirty', False)
                                          and not getattr(a, 'blocked_dialog', False)):
                    state.finish()
                    cleanup_temp(a)
                else:
                    state.unknown('Timeout, blocked dialog or unresolved COM execution')
                return status
            except Exception as error:
                uncertain = (a.worker_started or getattr(a, 'inventory_dirty', False)
                             or isinstance(error, getattr(pythoncom, 'com_error', ())))
                if uncertain:
                    state.unknown(str(error))
                else:
                    state.finish()
                    cleanup_temp(a)
                return failure('EXECUTION_UNKNOWN' if uncertain else 'PREFLIGHT_FAILED', error)
            finally:
                pythoncom.CoUninitialize()
    except (ValueError, OSError, RuntimeError, KeyError, TypeError) as error:
        return failure('PREFLIGHT_FAILED', error)


def connect(pid):
    """Attach to the SOLIDWORKS session of this process (Running Object Table 'SolidWorks_PID_<pid>');
    never starts a new SOLIDWORKS."""
    rot = pythoncom.GetRunningObjectTable()
    ctx = pythoncom.CreateBindCtx(0)
    for moniker in rot:
        if moniker.GetDisplayName(ctx, None) == f"SolidWorks_PID_{pid}":
            return w32.dynamic.Dispatch(rot.GetObject(moniker).QueryInterface(pythoncom.IID_IDispatch))
    raise RuntimeError(f"SOLIDWORKS PID {pid} is not registered for automation")


def attach(pid, explicit):
    # Preserve COM exception identity: rejected/busy attachment is unresolved too.
    return connect(pid) if explicit else w32.dynamic.Dispatch(w32.GetActiveObject("SldWorks.Application"))


def execute_macro(a, test_root, norm, state):
    pythoncom.CoInitialize()
    pids = [p.ProcessId for p in w32.GetObject("winmgmts:").ExecQuery(
        "SELECT ProcessId FROM Win32_Process WHERE Name='SLDWORKS.exe'")]
    if a.pid is not None:
        if a.pid not in pids:
            raise ValueError(f"No SLDWORKS.exe with PID {a.pid} (running: {pids})")
        pid = a.pid
    elif len(pids) == 1:
        pid = pids[0]
    else:
        raise ValueError(f"Expected exactly one SLDWORKS.exe process, found {len(pids)} ({pids}); "
                         "choose one with --pid (PowerShell: Get-Process SLDWORKS | Select Id,MainWindowTitle)")
    sw = attach(pid, a.pid is not None)
    state.set_process(pid)

    tmpdir = tempfile.mkdtemp(prefix="swrun_")
    a.temp_dir = tmpdir
    strict = a.strict_dialogs or a.answer is None
    code = a.prepared_code
    entry = a.entry
    swb = os.path.join(tmpdir, 'TestMacro.swp' if a.native_project else 'TestMacro.swb')
    with open(swb, "wb") as f:
        f.write(code if a.native_project else code.replace("\n", "\r\n").encode("cp1250"))
    observer_swb = None
    if a.observer_code is not None:
        observer_swb = os.path.join(tmpdir, 'ObserveMacro.swb')
        Path(observer_swb).write_bytes(a.observer_code.replace('\n', '\r\n').encode('cp1250'))

    report = {"macro": os.path.abspath(a.macro), "entry": entry, "solidworks": sw.RevisionNumber}
    report.update(module=a.module, macro_sha256=code_hash(code),
                  project_format='swp' if a.native_project else 'swb', compilation_status='NOT_TESTED')
    before = open_docs(sw)
    # Already loaded writable/dirty documents bypass OpenDoc6 read-only options.
    if before:
        raise ValueError('Use a SOLIDWORKS session with no documents open before each test')
    # Check both last-saved references and actual SW search-rule resolution BEFORE opening.
    for search in (False, True):
        dependencies = sw.GetDocumentDependencies2(os.path.abspath(a.open), True, search, False)
        a.test_profile.validate_dependencies(dependencies)
    previously_active = sw.ActiveDoc
    previously_active = previously_active.GetTitle if previously_active is not None else None
    if a.open:
        target_name = os.path.basename(a.open).lower()
        for p in before:
            if os.path.basename(p) == target_name and norm(p) != norm(a.open):
                print(f"A document named {os.path.basename(a.open)} is already open from {p}. "
                      "SOLIDWORKS can't hold two same-named files - close it first.")
                return 2
        err = w32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        warn = w32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        if not doc_type(a.open):
            raise ValueError('--open must be a native SOLIDWORKS part, assembly or drawing')
        options = 1 if a.test_profile.data['allow_model_changes'] else 3  # silent + read-only
        a.inventory_dirty = True
        m = sw.OpenDoc6(os.path.abspath(a.open), doc_type(a.open), options, "", err, warn)
        if m is None:
            raise ValueError(f"Cannot open {a.open}: error {err.value}")
        if norm(m.GetPathName) != norm(a.open):
            # SOLIDWORKS keys documents by file name: a same-named file from another folder is already loaded
            print(f"SOLIDWORKS returned {m.GetPathName} instead of {a.open} - a file with the same name is "
                  "already open. Close it first (SOLIDWORKS can't hold two same-named files).")
            return 2
        e2 = w32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        active = sw.ActivateDoc3(m.GetTitle, False, 2, e2)
        if active is None or e2.value != 0 or norm(active.GetPathName) != norm(a.open):
            raise ValueError('SOLIDWORKS did not activate the requested test copy')
        report["opened"] = m.GetPathName

    for path in open_docs(sw):
        if not path or norm(path) not in a.test_profile.inputs:
            raise ValueError(f'Opening the copy loaded an undeclared document: {path!r}')

    for f in a.output_identities:
        revalidate_output(f, a.test_profile.data['test_root'], a.output_identities[f])
    previous_files = {f: file_snapshot(f) for f in a.expect_file}
    previous_logs = {f: read_log(f) for f, _ in a.expect_text + a.expect_result}
    watcher = DialogWatcher(pid, a.answer, strict, a.test_profile.data['allowed_dialogs'])
    watcher.start()
    result = {}
    deadline = time.monotonic() + a.timeout
    allow_followups = threading.Event()
    allow_followups.set()

    def run():
        pythoncom.CoInitialize()
        try:
            s = attach(pid, a.pid is not None)
            e = w32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
            result["ok"] = bool(s.RunMacro2(swb, a.module, entry, 1, e))  # 1 = unload after run
            result["err"] = e.value
            if (observer_swb and result['ok'] and result['err'] == 0 and not watcher.blocked
                    and not watcher.stop.is_set() and allow_followups.is_set() and time.monotonic() < deadline):
                oe = w32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
                result['observation_ok'] = bool(s.RunMacro2(observer_swb, 'ObserveMacro', 'main', 1, oe))
                result['observation_error'] = oe.value
        except Exception as ex:  # noqa: BLE001
            result["exception"] = str(ex)
        finally:
            pythoncom.CoUninitialize()

    t0 = time.time()
    runner = threading.Thread(target=run, daemon=True)
    a.worker_started = not watcher.blocked
    if a.worker_started:
        runner.start()
    while runner.is_alive() and not watcher.blocked:
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            break
        runner.join(min(0.1, remaining))
    macro_still_running = runner.is_alive()
    # A late COM return must never initiate another macro after this controller
    # has exhausted its wait or observed an unknown dialog.
    allow_followups.clear()
    if not macro_still_running:
        time.sleep(0.8)
    watcher.stop.set()
    watcher.join(timeout=1)
    report["seconds"] = round(time.time() - t0, 1)
    report["timed_out"] = macro_still_running and not watcher.blocked
    report["macro_still_running"] = macro_still_running
    report["run_ok"] = result.get("ok")
    report["run_error"] = RUN_ERRORS.get(result.get("err"), result.get("err"))
    if observer_swb:
        report['observation_ok'] = result.get('observation_ok')
        report['observation_error'] = result.get('observation_error')
        report['observer_sha256'] = code_hash(a.observer_code)
    if "exception" in result:
        report["exception"] = result["exception"]
    report["dialogs"] = watcher.dialogs
    report["compile_error"] = watcher.compile_error
    report["runtime_error"] = watcher.runtime_error
    a.blocked_dialog = watcher.blocked

    # A synchronous COM call here can hang forever while the macro owns SW's UI thread.
    # Report from captured Python state only; leave live documents and the macro alone.
    uncertain_error = (result.get('err') in UNCERTAIN_RUN_ERRORS
                       or result.get('observation_error') in UNCERTAIN_RUN_ERRORS)
    safe_to_query = (not macro_still_running and not watcher.blocked
                     and 'exception' not in result and not uncertain_error)
    a.session_resolved = False
    after = open_docs(sw) if safe_to_query else {}
    new_docs = [p for p in after if p not in before]
    undeclared = [p for p in new_docs if not p or norm(p) not in a.test_profile.inputs]
    inside_docs = [p for p in new_docs if p not in undeclared and inside(p, test_root)]
    report['undeclared_documents'] = undeclared
    if undeclared:
        report['failure_code'] = 'UNDECLARED_DOCUMENT'
    report["documents_left_open"] = new_docs if safe_to_query else None
    if a.close_opened and safe_to_query:
        for p in inside_docs:
            sw.CloseDoc(after[p])           # test copies only - never documents outside the test root
        remaining = open_docs(sw)
        report["closed_test_documents"] = [p for p in inside_docs if p not in remaining]
        report["documents_left_open"] = [p for p in remaining if p not in before]
        additional = [p for p in report['documents_left_open']
                      if (not p or norm(p) not in a.test_profile.inputs) and p not in undeclared]
        if additional:
            undeclared.extend(additional)
            report['failure_code'] = 'UNDECLARED_DOCUMENT'
        if any(p in remaining for p in inside_docs):
            report.setdefault('failure_code', 'CLEANUP_FAILED')
    a.session_resolved = bool(safe_to_query and not undeclared and
                              (not a.close_opened or not report['documents_left_open']))
    if previously_active and safe_to_query:
        e3 = w32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        try:
            sw.ActivateDoc3(previously_active, False, 2, e3)   # give the user back their active document
        except Exception:  # noqa: BLE001
            pass

    expect_failures = []
    # Permission to edit an in-memory copy never permits overwriting its hashed
    # input file. Persisted results must use distinct declared output paths.
    if safe_to_query:
        try:
            a.test_profile.check_input_hashes()
        except ValueError as error:
            expect_failures.append(str(error))
    for f in a.expect_file:
        try:
            revalidate_output(f, a.test_profile.data['test_root'], a.output_identities[f])
        except (OSError, ValueError) as error:
            expect_failures.append(str(error))
            continue
        current = file_snapshot(f)
        if current is None:
            expect_failures.append(f"missing: {f}")
        elif current == previous_files[f]:
            expect_failures.append(f"not written by this run (unchanged file): {f}")
    log_checks = ([(f, needle, check_log_text) for f, needle in a.expect_text]
                  + [(f, prefix, check_log_result) for f, prefix in a.expect_result])
    for f, needle, check in log_checks:
        try:
            revalidate_output(f, a.test_profile.data['test_root'], a.output_identities[f])
        except (OSError, ValueError) as error:
            expect_failures.append(str(error))
            continue
        problem = check(new_log_text(previous_logs[f], read_log(f)), needle)
        if problem:
            expect_failures.append(f"{problem} in {f}")
    report["expectations_failed"] = expect_failures

    if report["run_ok"] is False and not watcher.dialogs and report["run_error"] == "OK":
        # .swb macros fail silently on compile errors
        report["hint"] = ("Macro did not start and showed no error - almost always a COMPILE error "
                          "(undeclared variable, unknown member, wrong argument count, syntax). "
                          "Run check_api.py and check_vba.py on it.")
    ok = (report["run_ok"] and result.get('err') == 0 and a.session_resolved
          and not watcher.compile_error and not watcher.runtime_error
          and not report["timed_out"] and not expect_failures and not watcher.blocked
          and (not observer_swb or (result.get('observation_ok') and result.get('observation_error') == 0)))
    report["status"] = "PASS" if ok else ("BLOCKED_DIALOG" if watcher.blocked else "FAIL")
    report['session_state'] = 'IDLE' if a.session_resolved else 'UNKNOWN'
    report['session_marker'] = str(SESSION_PATH)
    if watcher.blocked:
        report["hint"] = ("A message box was left unanswered. Inspect it in SOLIDWORKS and wait for the "
                          "macro to finish before rerunning. No document cleanup was attempted.")
    elif report["timed_out"]:
        report["hint"] = ("The timeout does not stop the SOLIDWORKS macro. Its documents were not queried "
                          "or closed. Inspect SOLIDWORKS and wait for it to finish before rerunning.")
    elif uncertain_error:
        report['hint'] = 'SOLIDWORKS reported an unresolved COM status; inspect it before explicit session recovery.'
    out = json.dumps(report, ensure_ascii=False, indent=2)
    print(out)
    if a.json:
        try:
            revalidate_output(a.json, a.test_profile.data['test_root'], a.output_identities[a.json])
            write_report(a.json, out)
        except (OSError, ValueError) as error:
            print(f'WARN  Could not create report: {error}', file=sys.stderr)
            return 2
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
