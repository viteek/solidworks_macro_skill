#!/usr/bin/env python3
r"""Run a VBA macro (.bas/.swb/.txt) in the user's running SOLIDWORKS and report what happened.

Lets you *test* a macro you wrote instead of handing it over untested. Windows only;
needs Python 3 + pywin32 (pip install pywin32) and a running SOLIDWORKS session.

How it works
- The code is saved as a temporary text macro (.swb, cp1250, CRLF) - SOLIDWORKS runs .swb files
  with the SOLIDWORKS type libraries available, so early-bound code and sw* constants work.
- A watcher records recognized Windows dialogs. Message boxes are left unanswered by default;
  --answer explicitly permits only the matching button, never a fallback button. VBA run-time
  errors -> "End", compile errors -> "OK" (and the VBA editor window it opens is closed).
  Folder pickers are cancelled -
  use --define to give folder constants a value instead, or --pick-folder.
  --strict-dialogs keeps message boxes unanswered even with --answer (status BLOCKED_DIALOG).
  SendMsgToUser2 calls are never rewritten. Opaque/custom dialogs that cannot be recognized
  may reach the timeout; inspect SOLIDWORKS manually in that case.
- Optionally opens a document first (--open; refuses if a same-named file from another folder is
  loaded). --close-opened closes only new documents inside the test root (default: folder of --open),
  never the user's own documents; the previously active document is re-activated afterwards.
- --expect-file checks output freshness; --expect-text checks only appended text, or the contents
  of a new log or a log whose previous byte prefix changed. Identical content is conservatively
  treated as unchanged, even after replacement. Prefer an appended per-run result or run ID.
- Timeout/blocked dialog: report without further COM calls or document cleanup. The macro may
  still be running in SOLIDWORKS; the timeout does not cancel it. Finish it before another run.

Usage
  python run_in_solidworks.py macro.bas [--open C:\SWTest\part.sldprt] [--define SAVE_CHANGES=True]
         [--define "OUT_DIR=\"C:\\SWTest\\out\""] [--answer yes|no|ok|cancel] [--timeout 300]
         [--close-opened] [--expect-file C:\SWTest\out\part.dxf] [--json report.json]

--define NAME=VALUE rewrites the line `Const NAME [As T] = ...` to the given VBA literal, so a macro
whose folder/option is a Const (with a picker only when the Const is empty) can run unattended.

Exit code 0 = macro ran with no compile/run-time error dialog, 1 = error, 2 = could not start.
Never run this against the user's production files: point macros at copies.
"""
import argparse
import json
import msvcrt
import os
import re
import sys
import tempfile
import threading
import time

try:
    import pythoncom
    import win32api
    import win32con
    import win32gui
    import win32process
    import win32com.client as w32
except ImportError:
    print("pywin32 is required: pip install pywin32")
    sys.exit(2)

RUN_ERRORS = {
    0: "OK", 1: "InvalidArg", 2: "MacrosAreDisabled", 3: "NotInDesignMode", 4: "OnlyCodeModules",
    5: "OutOfMemory", 6: "InvalidProcname", 7: "InvalidPropertyType", 8: "SuppressedForever",
    9: "CannotOpenFile",
}
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
    def __init__(self, pid, answer=None, strict=False):
        super().__init__(daemon=True)
        self.pid, self.answer, self.strict = pid, answer, strict or answer is None
        self.blocked = False
        self.stop = threading.Event()
        self.dialogs = []
        self.compile_error = False
        self.runtime_error = False
        self._seen = set()

    def run(self):
        while not self.stop.is_set():
            win32gui.EnumWindows(self._check, None)
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
            if not win32gui.IsWindowVisible(hwnd):
                return True
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid != self.pid:
                return True
            cls = win32gui.GetClassName(hwnd)
            title = win32gui.GetWindowText(hwnd)
            if cls == "wndclass_desked_gsk" and self.compile_error:
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)   # VBA editor opened on a compile error
                return True
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
                pressed = self._click(hwnd, buttons, ("Cancel", "Anuluj"))
            elif re.search(r"Run-time error|Błąd czasu wykonania|Laufzeitfehler", body, re.I):
                kind, self.runtime_error = "runtime_error", True
                pressed = self._click(hwnd, buttons, END_BUTTONS)
            elif re.search(r"Compile error|Błąd kompilacji|Syntax error|Błąd składni", body, re.I):
                kind, self.compile_error = "compile_error", True
                pressed = self._click(hwnd, buttons, ("OK",))
            elif self.strict:
                kind, self.blocked, pressed = "blocked_message", True, None   # left for the user
            else:
                pressed = self._click(hwnd, buttons, ANSWERS[self.answer])
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


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("macro")
    ap.add_argument("--open", help="document to open and activate before running")
    ap.add_argument("--define", action="append", default=[], help="NAME=VBA_LITERAL, overrides a Const")
    ap.add_argument("--answer", choices=list(ANSWERS),
                    help="explicitly allow this dialog answer (default: leave message boxes unanswered)")
    ap.add_argument("--pick-folder", action="append", default=[],
                    help="answer Shell BrowseForFolder pickers with these folders, in order")
    ap.add_argument("--entry", help="procedure to run (default main)")
    ap.add_argument("--timeout", type=float, default=600)
    ap.add_argument("--close-opened", action="store_true",
                    help="close documents that are new after the run AND lie inside the test root (unsaved changes "
                         "of those test copies are discarded); documents elsewhere are never touched")
    ap.add_argument("--test-root", help="folder holding the test copies (default: folder of --open)")
    ap.add_argument("--strict-dialogs", action="store_true",
                    help="never answer message boxes: report BLOCKED_DIALOG instead (VBA run-time/compile "
                         "errors are still ended). Use when a prompt could confirm something destructive")
    ap.add_argument("--expect-file", action="append", default=[],
                    help="file that must exist and be written during this run (repeatable)")
    ap.add_argument("--expect-text", nargs=2, action="append", default=[], metavar=("FILE", "TEXT"),
                    help="newly appended/replaced text in FILE must contain TEXT (repeatable)")
    ap.add_argument("--json", help="write the report here")
    a = ap.parse_args()

    def norm(x):
        try:
            x = win32api.GetLongPathName(os.path.abspath(x))
        except Exception:  # noqa: BLE001
            pass
        return os.path.normcase(os.path.abspath(x))

    test_root = norm(a.test_root) if a.test_root else (norm(os.path.dirname(a.open)) if a.open else None)
    if a.open and test_root and not norm(a.open).startswith(test_root + os.sep):
        print(f"--open {a.open} is outside --test-root {a.test_root}")
        return 2
    if a.close_opened and not test_root:
        print("--close-opened needs --open or --test-root (only test copies are ever closed)")
        return 2

    # one macro at a time per machine: SOLIDWORKS runs macros on its UI thread
    with open(os.path.join(tempfile.gettempdir(), "run_in_solidworks.lock"), "a+") as lock:
        waited = time.monotonic()
        while True:
            try:
                msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
                break
            except OSError:
                if time.monotonic() - waited > 1800:
                    print("Another macro test has been running for 30 min - giving up.")
                    return 2
                time.sleep(2)
        return execute_macro(a, test_root, norm)


def execute_macro(a, test_root, norm):
    pythoncom.CoInitialize()
    try:
        sw = w32.dynamic.Dispatch(w32.GetActiveObject("SldWorks.Application"))
    except Exception as e:  # noqa: BLE001
        print(f"Cannot connect to a running SOLIDWORKS: {e}")
        return 2
    pids = [p.ProcessId for p in w32.GetObject("winmgmts:").ExecQuery(
        "SELECT ProcessId FROM Win32_Process WHERE Name='SLDWORKS.exe'")]
    if len(pids) != 1:
        print(f"Expected exactly one SLDWORKS.exe process, found {len(pids)}")
        return 2
    pid = pids[0]

    with open(a.macro, "rb") as source:
        raw = source.read()
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = raw.decode("cp1250")
    tmpdir = tempfile.mkdtemp(prefix="swrun_")
    strict = a.strict_dialogs or a.answer is None
    code = prepare_code(text, a.define, a.pick_folder)
    entry = a.entry or find_entry(code)
    swb = os.path.join(tmpdir, "TestMacro.swb")
    with open(swb, "wb") as f:
        f.write(code.replace("\n", "\r\n").encode("cp1250", errors="replace"))

    report = {"macro": os.path.abspath(a.macro), "entry": entry, "solidworks": sw.RevisionNumber}
    before = open_docs(sw)
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
        m = sw.OpenDoc6(os.path.abspath(a.open), doc_type(a.open), 1, "", err, warn)
        if m is None:
            print(f"Cannot open {a.open}: error {err.value}")
            return 2
        if norm(m.GetPathName) != norm(a.open):
            # SOLIDWORKS keys documents by file name: a same-named file from another folder is already loaded
            print(f"SOLIDWORKS returned {m.GetPathName} instead of {a.open} - a file with the same name is "
                  "already open. Close it first (SOLIDWORKS can't hold two same-named files).")
            return 2
        e2 = w32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        sw.ActivateDoc3(m.GetTitle, False, 2, e2)
        report["opened"] = m.GetPathName

    previous_logs = {f: read_log(f) for f, _ in a.expect_text}
    watcher = DialogWatcher(pid, a.answer, strict)
    watcher.start()
    result = {}

    def run():
        pythoncom.CoInitialize()
        try:
            s = w32.dynamic.Dispatch(w32.GetActiveObject("SldWorks.Application"))
            e = w32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
            result["ok"] = bool(s.RunMacro2(swb, "TestMacro", entry, 1, e))  # 1 = unload after run
            result["err"] = e.value
        except Exception as ex:  # noqa: BLE001
            result["exception"] = str(ex)
        finally:
            pythoncom.CoUninitialize()

    t0 = time.time()
    runner = threading.Thread(target=run, daemon=True)
    runner.start()
    deadline = time.monotonic() + a.timeout
    while runner.is_alive() and not watcher.blocked:
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            break
        runner.join(min(0.1, remaining))
    macro_still_running = runner.is_alive()
    if not macro_still_running:
        time.sleep(0.8)
    watcher.stop.set()
    watcher.join(timeout=1)
    report["seconds"] = round(time.time() - t0, 1)
    report["timed_out"] = macro_still_running and not watcher.blocked
    report["macro_still_running"] = macro_still_running
    report["run_ok"] = result.get("ok")
    report["run_error"] = RUN_ERRORS.get(result.get("err"), result.get("err"))
    if "exception" in result:
        report["exception"] = result["exception"]
    report["dialogs"] = watcher.dialogs
    report["compile_error"] = watcher.compile_error
    report["runtime_error"] = watcher.runtime_error

    # A synchronous COM call here can hang forever while the macro owns SW's UI thread.
    # Report from captured Python state only; leave live documents and the macro alone.
    safe_to_query = not macro_still_running and not watcher.blocked
    after = open_docs(sw) if safe_to_query else {}
    new_docs = [p for p in after if p not in before]
    inside = [p for p in new_docs if test_root and norm(p).startswith(test_root + os.sep)]
    report["documents_left_open"] = new_docs if safe_to_query else None
    if a.close_opened and safe_to_query:
        for p in inside:
            sw.CloseDoc(after[p])           # test copies only - never documents outside the test root
        report["closed_test_documents"] = inside
        report["documents_left_open"] = [p for p in new_docs if p not in inside]
    if previously_active and safe_to_query:
        e3 = w32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        try:
            sw.ActivateDoc3(previously_active, False, 2, e3)   # give the user back their active document
        except Exception:  # noqa: BLE001
            pass

    started = t0
    expect_failures = []
    for f in a.expect_file:
        if not os.path.exists(f):
            expect_failures.append(f"missing: {f}")
        elif os.path.getmtime(f) < started - 2:
            expect_failures.append(f"not written by this run (old file): {f}")
    for f, needle in a.expect_text:
        content = new_log_text(previous_logs[f], read_log(f))
        if needle not in content:
            expect_failures.append(f"text {needle!r} not found in {f}")
    report["expectations_failed"] = expect_failures

    if report["run_ok"] is False and not watcher.dialogs and report["run_error"] == "OK":
        # .swb macros fail silently on compile errors
        report["hint"] = ("Macro did not start and showed no error - almost always a COMPILE error "
                          "(undeclared variable, unknown member, wrong argument count, syntax). "
                          "Run check_api.py and check_vba.py on it.")
    ok = (report["run_ok"] and not watcher.compile_error and not watcher.runtime_error
          and not report["timed_out"] and not expect_failures and not watcher.blocked)
    report["status"] = "PASS" if ok else ("BLOCKED_DIALOG" if watcher.blocked else "FAIL")
    if watcher.blocked:
        report["hint"] = ("A message box was left unanswered. Inspect it in SOLIDWORKS and wait for the "
                          "macro to finish before rerunning. No document cleanup was attempted.")
    elif report["timed_out"]:
        report["hint"] = ("The timeout does not stop the SOLIDWORKS macro. Its documents were not queried "
                          "or closed. Inspect SOLIDWORKS and wait for it to finish before rerunning.")
    out = json.dumps(report, ensure_ascii=False, indent=2)
    print(out)
    if a.json:
        with open(a.json, "w", encoding="utf-8") as f:
            f.write(out)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
