#!/usr/bin/env python3
"""Static checks for SOLIDWORKS VBA macros (.bas / .txt / .vb).

Not a compiler - it catches the mistakes that most often break SOLIDWORKS macros
or make them dangerous on production files.

Usage:
    python check_vba.py macro.bas [more.bas ...] [--cp1250]

--cp1250  also writes <name>_cp1250.bas, re-encoded for import into the VBA editor
          on Polish Windows (fixes garbled Polish characters after File > Import).

Exit code 1 if any ERROR was found.
"""
import re
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from check_api import DECL_RE, PROC_RE, parse_decl_list, split_args, join_continuations

# Only known Boolean members are inferred in direct conditions. A declared Boolean
# assignment additionally provides the expected result type for other SW members.
BOOLEAN_MEMBERS = {'save3', 'saveas', 'exporttodwg2', 'getsaveflag', 'issuppressed',
                   'ishidden', 'select2', 'select4', 'selectbyid2', 'showconfiguration2',
                   'activateSheet', 'setsheets', 'modifydefinition', 'accessselections',
                   'excludefrombom', 'forcerebuild3', 'editrebuild3', 'issheetmetal', 'getuserpreferencetoggle'}
BOOLEAN_MEMBERS = {name.lower() for name in BOOLEAN_MEMBERS}
SW_TYPES = {'modeldoc2', 'sldworks', 'assemblydoc', 'partdoc', 'drawingdoc', 'feature',
            'component2', 'modeldocextension', 'selectionmgr', 'entity', 'exportpdfdata'}
NUMERIC_MEMBERS = {'gettype', 'getselectedobjectcount2', 'getselectedobjecttype3',
                   'getselectedobjectmark', 'getuserpreferenceintegervalue', 'get6',
                   'add3', 'setsystemvalue3', 'setsuppression2', 'resolvealllightweightcomponents'}


def boolean_issues(lines):
    """Bounded, procedure-scoped provenance analysis; not a VBA type checker."""
    module_types, types, origins, module_origins = {}, {}, set(), set()
    in_proc = False
    with_roots = []
    inferred_objects = set()

    def sw_root(name):
        kind = types.get(name.lower(), ('', False))[0].lower()
        return kind.startswith('sldworks.') or kind in SW_TYPES

    for number, logical in join_continuations(lines):
        proc = PROC_RE.match(logical)
        if proc:
            in_proc = True
            types, origins, with_roots = dict(module_types), set(module_origins), []
            inferred_objects = set()
            for parameter in split_args(proc.group(3)):
                parameter = re.sub(r'^(?:(?:ByVal|ByRef|Optional|ParamArray)\s+)+', '', parameter.strip(), flags=re.I)
                types.update(parse_decl_list(parameter))
            continue
        if re.match(r'\s*End\s+(?:Sub|Function|Property)\b', logical, re.I):
            in_proc = False
            types, origins, with_roots = dict(module_types), set(module_origins), []
            continue
        declaration = DECL_RE.match(logical)
        if declaration:
            declared = parse_decl_list(declaration.group(1))
            types.update(declared)
            if not in_proc:
                module_types.update(declared)
            continue
        wm = re.match(r'\s*With\s+(\w+)', logical, re.I)
        if wm:
            with_roots.append(wm.group(1))
            continue
        if re.match(r'\s*End\s+With', logical, re.I):
            if with_roots:
                with_roots.pop()
            continue
        # Strings have been blanked, so colons/Then/Else inside messages cannot split code.
        for statement in re.split(r':|\b(?:Then|Else)\b', logical, flags=re.I):
            if with_roots:
                statement = re.sub(r'(?<![\w.])\.(?=\w)', with_roots[-1] + '.', statement)
            assignment = re.match(r'\s*(Set\s+)?(\w+)\s*=\s*(.+)', statement, re.I)
            if assignment:
                is_object, name, expression = assignment.groups()
                name = name.lower()
                root = re.match(r'\s*(\w+)(?:\.\w+)*', expression)
                if is_object and root:
                    target_kind = types.get(name, ('',))[0].lower()
                    if sw_root(root.group(1)) and (name in inferred_objects or target_kind in ('', 'object', 'variant')):
                        types[name] = ('SldWorks.InferredObject', False)
                        inferred_objects.add(name)
                    elif not sw_root(root.group(1)) and (name in inferred_objects or target_kind in ('object', 'variant')):
                        types[name] = ('Object', False)
                        inferred_objects.discard(name)
                elif types.get(name, ('',))[0].lower() == 'boolean':
                    # Comparisons normalize COM truth values to a VBA Boolean.
                    depth, comparison = 0, False
                    exp = expression.strip()
                    while exp.startswith('(') and exp.endswith(')'):
                        balance, encloses = 0, True
                        for index, char in enumerate(exp):
                            balance += (char == '(') - (char == ')')
                            if balance == 0 and index < len(exp) - 1:
                                encloses = False
                                break
                        if not encloses:
                            break
                        exp = exp[1:-1].strip()
                    for char in exp:
                        depth += (char == '(') - (char == ')')
                        if depth == 0 and char in '=<>':
                            comparison = True
                    source = re.match(r'(?:Not\s+)?\(*(\w+)(?:\.\w+)*\.(\w+)', exp, re.I)
                    contaminated = not comparison and ((source and sw_root(source.group(1))
                        and source.group(2).lower() not in NUMERIC_MEMBERS) or exp.lower() in origins)
                    if contaminated:
                        origins.add(name)
                        if re.match(r'Not\b', exp, re.I):
                            yield number, 'Unsafe Not on a SOLIDWORKS COM Boolean assignment; compare the call with False.'
                    else:
                        origins.discard(name)
                if not re.search(r'\bNot\b', expression, re.I):
                    continue
            else:
                condition = re.search(r'\b(?:If|ElseIf|While|Until)\b(.*)', statement, re.I)
                if not condition:
                    continue
                expression = condition.group(1)
            # Parenthesized zero literals have the same semantics as bare literals.
            while re.search(r'\(\s*(?:False|0)\s*\)', expression, re.I):
                expression = re.sub(r'\(\s*(False|0)\s*\)', r'\1', expression, flags=re.I)
            # Explicit zero comparisons are safe, including Not (ok = False).
            expression = re.sub(r'\b\w+\b\s*(?:<>|=)\s*(?:False|0)\b', '', expression, flags=re.I)
            expression = re.sub(r'\b(?:False|0)\s*(?:<>|=)\s*\w+\b', '', expression, flags=re.I)
            for name in origins:
                if re.search(r'\b' + re.escape(name) + r'\b', expression, re.I):
                    yield number, f"'{name}' holds a SOLIDWORKS COM Boolean; compare explicitly with False (= False / <> False)."
                    break
            for call in re.finditer(r'\b(\w+)(?:\.\w+)*\.(\w+)\b', expression):
                if not sw_root(call.group(1)) or call.group(2).lower() not in BOOLEAN_MEMBERS:
                    continue
                tail = expression[call.end():].lstrip()
                if tail.startswith('('):
                    depth = 0
                    for index, char in enumerate(tail):
                        depth += (char == '(') - (char == ')')
                        if depth == 0:
                            tail = tail[index + 1:]
                            break
                if not re.match(r'\s*\)*\s*(?:<>|=)\s*(?:False|0)\b', tail, re.I):
                    yield number, f"'{call.group(0)}' returns a COM Boolean; compare explicitly with False."

OBSOLETE = {
    r"\.OpenDoc[2-5]?\s*\(": "OpenDoc6",
    r"\.OpenDoc\s+[^6(]": "OpenDoc6",
    r"\.SaveAs[24]\b": "Extension.SaveAs / SaveAs3",
    r"\.Save2\b": "Save3",
    r"\.AddCustomInfo[0-9]?\b": "CustomPropertyManager.Add3",
    r"\.CustomInfo2?\b": "CustomPropertyManager.Get6",
    r"\.Get[45]\s*[ (]": "CustomPropertyManager.Get6 (2018+)",
    r"\.SelectByID\s*[ (]": "Extension.SelectByID2",
    r"\.GetSelectedObject5\b": "GetSelectedObject6",
    r"\.GetSelectedObjectType2\b": "GetSelectedObjectType3",
    r"\.GetRootComponent2?\b": "GetRootComponent3(True)",
    r"\.GetModelDoc\b(?!2)": "GetModelDoc2",
    r"\.ForceRebuild2\b": "ForceRebuild3",
    r"\.EditRebuild\b(?!3)": "EditRebuild3",
    r"(?<![Bb]ody)\.GetMassProperties\b": "Extension.CreateMassProperty (Body2.GetMassProperties is fine)",
}


VBA_RESERVED = {
    "and", "as", "boolean", "byref", "byte", "byval", "call", "case", "close", "const", "currency", "date",
    "declare", "dim", "do", "double", "each", "else", "elseif", "empty", "end", "enum", "eqv", "erase",
    "error", "event", "exit", "false", "for", "friend", "function", "get", "global", "gosub", "goto", "if",
    "imp", "implements", "in", "input", "integer", "is", "let", "lib", "like", "lock", "long", "loop",
    "lset", "me", "mod", "new", "next", "not", "nothing", "null", "on", "open", "option", "optional", "or",
    "paramarray", "preserve", "print", "private", "property", "public", "put", "raiseevent", "redim", "rem",
    "resume", "return", "rset", "seek", "select", "set", "single", "static", "step", "stop", "string",
    "sub", "then", "to", "true", "type", "typeof", "unlock", "until", "variant", "wend", "while", "with",
    "withevents", "write", "xor",
    # VBA language spec 3.3.5.2 special forms; Circle caused silent .swb compile failures.
    "array", "circle", "inputb", "lbound", "scale", "ubound",
}


def strip_comment(line: str) -> str:
    """Remove a trailing ' comment, ignoring apostrophes inside string literals."""
    in_str = False
    for i, ch in enumerate(line):
        if ch == '"':
            in_str = not in_str
        elif ch == "'" and not in_str:
            return line[:i]
    if re.match(r"\s*Rem\s", line, re.I):
        return ""
    return line


def check(path: Path, snippet=False) -> int:
    raw = path.read_bytes()
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError:
        text = raw.decode("cp1250")
    lines = text.splitlines()
    code = [strip_comment(l) for l in lines]
    joined = "\n".join(code)
    part_receivers = {m.group(1).lower() for m in re.finditer(
        r'\b(\w+)\s+As\s+(?:SldWorks\.)?PartDoc\b', joined, re.I)}
    issues = []

    def add(level, lineno, msg):
        issues.append((level, lineno, msg))

    if not snippet and not re.search(r"^\s*Option\s+Explicit\b", joined, re.I | re.M):
        add("ERROR", 0, "Missing 'Option Explicit' - typos become silent new variables.")

    if not snippet and not re.search(r"^\s*(Public\s+|Private\s+)?Sub\s+main\s*\(", joined, re.I | re.M):
        add("WARN", 0, "No 'Sub main()' - SOLIDWORKS runs 'main' by default.")

    # Every public parameterless Sub is a possible entry point for Tools > Macro > Run
    entry = re.findall(r"^\s*(?:Public\s+)?Sub\s+(\w+)\s*\(\s*\)", joined, re.I | re.M)
    if len(entry) > 1:
        add("ERROR", 0, f"Several public parameterless Subs ({', '.join(entry)}) - SW may run the wrong one. "
                        "Make helpers Private or give them a parameter.")

    # restored = any DocumentVisible call whose first argument isn't the literal False (True, a saved variable...)
    if re.search(r"DocumentVisible\s*\(?\s*False", joined, re.I) and not re.search(r"DocumentVisible\s*\(?\s*(?!False\b)\w", joined, re.I):
        add("ERROR", 0, "DocumentVisible False is never restored to True - all later documents would open invisible.")

    uses_sw = re.search(r"\bswApp\b|SldWorks", joined)
    if uses_sw and not re.search(r"Application\.SldWorks|GetObject|CreateObject\(\s*\"SldWorks", joined, re.I):
        add("WARN", 0, "swApp is used but never obtained (Set swApp = Application.SldWorks).")

    has_handler = re.search(r"On\s+Error\s+GoTo\s+\w+", joined, re.I)
    if not has_handler:
        add("WARN", 0, "No 'On Error GoTo <label>' handler - errors leave SW state (graphics, tree) unrestored.")

    for number, message in boolean_issues(lines):
        add('ERROR', number, message)

    for i, l in enumerate(code, 1):
        s = l.strip()
        if not s:
            continue

        for dm in re.finditer(r"(?:\b(?:Dim|Static|Private|Public|ByVal|ByRef|Optional|Const)\s+|,\s*)(\w+)(?=\s*(?:\([^)]*\))?\s+As\b)", s, re.I):
            if dm.group(1).lower() in VBA_RESERVED:
                add("ERROR", i, f"'{dm.group(1)}' is a VBA keyword - can't be a variable name (compile error; a .swb run just won't start).")
        pm = re.match(r"(?:Public\s+|Private\s+)?(?:Sub|Function)\s+(\w+)", s, re.I)
        if pm and pm.group(1).lower() in VBA_RESERVED:
            add("ERROR", i, f"'{pm.group(1)}' is a VBA keyword - rename the procedure.")

        if s.count('"') % 2 == 1:
            add("ERROR", i, "Odd number of quotes - unterminated string literal (syntax error).")

        # Dim a, b As Long  -> a is Variant
        m = re.match(r"(Dim|Private|Public|Static)\s+(.+)", s, re.I)
        if m and not re.match(r"(Private|Public)\s+(Sub|Function|Const|Declare|Type|Enum|Property)\b", s, re.I):
            decl = m.group(2)
            if not re.match(r"(Sub|Function|Const)\b", decl, re.I):
                for part in decl.split(","):
                    part = part.strip()
                    if part and " as " not in part.lower():
                        add("WARN", i, f"'{part}' has no 'As <Type>' (becomes Variant).")

        if re.search(r"On\s+Error\s+Resume\s+Next", s, re.I):
            # Fine only if Err is checked shortly after
            window = "\n".join(code[i:i + 6])
            if not re.search(r"Err\.Number|Err\s*<>|On\s+Error\s+GoTo", window, re.I):
                add("WARN", i, "'On Error Resume Next' without checking Err nearby - hides failures.")

        for pat, repl in OBSOLETE.items():
            if re.search(pat, l):
                add("WARN", i, f"Obsolete API call - prefer {repl}.")
                break
        # IBodyFolder.GetBodies is current; only IPartDoc.GetBodies is superseded.
        obsolete_bodies = re.search(r'\b(\w+)\.GetBodies\b(?!2)', l, re.I)
        if obsolete_bodies and obsolete_bodies.group(1).lower() in part_receivers:
            add('WARN', i, 'Obsolete IPartDoc.GetBodies call - prefer GetBodies2.')

        # Unit sanity: assigning large literals to SystemValue (meters)
        m = re.search(r"\.SystemValue\s*=\s*([0-9.]+)", l)
        if m and float(m.group(1)) >= 5:
            add("WARN", i, f"SystemValue = {m.group(1)} is in METERS - did you mean mm/1000?")
        m = re.search(r"SetSystemValue3\s*\(?\s*([0-9.]+)", l)
        if m and float(m.group(1)) >= 5:
            add("WARN", i, f"SetSystemValue3 {m.group(1)} is in METERS - did you mean mm/1000?")

        if len(lines[i - 1]) > 1023:
            add("ERROR", i, "Line longer than 1023 characters - VBA can't compile it; split with '& _'.")

        if re.match(r"(Public\s+|Private\s+)?Declare\s+(?!PtrSafe)", s, re.I):
            add("ERROR", i, "Declare without PtrSafe - fails in 64-bit VBA (use Declare PtrSafe + LongPtr).")

        m = re.search(r"SelectByID2\s*\(?\s*\"[^\"]*\"\s*,\s*\"[^\"]*\"\s*,\s*([^,]+),\s*([^,]+),\s*([^,]+),", l, re.I)
        if m and any(re.sub(r"[#\s]", "", v) not in ("0", "0.0", "0!") for v in m.groups()):
            add("WARN", i, "SelectByID2 with coordinates (recorded macro?) - fragile; select via object methods.")

        if re.search(r"InStr(Rev)?\s*\(.*\)\s*(<>|=)\s*-1\b", l, re.I):
            add("WARN", i, "InStr returns 0 when not found, never -1.")

        if re.search(r"\.Add2\s*\(?\s*\"", l) and re.search(r"CustomProperty|cpm|PropMgr", joined, re.I):
            add("WARN", i, "CustomPropertyManager.Add2 doesn't overwrite existing values - use Add3 with swCustomPropertyReplaceValue.")

        if re.search(r"\.GetPartBox\b|\.GetBodyBox\b", l):
            add("INFO", i, "GetPartBox/GetBodyBox are approximate (>10% possible) - use GetExtremePoint for exact sizes.")

        if re.search(r"\bCloseAllDocuments\b", l):
            add("WARN", i, "CloseAllDocuments closes the user's open work - make sure this is intended.")

        if re.search(r"\.GetModelDoc2\b", l):
            window = "\n".join(code[i - 1:i + 4])
            if not re.search(r"Is\s+Nothing", window, re.I):
                add("WARN", i, "GetModelDoc2 result not checked for Nothing (lightweight/suppressed components).")

        if re.search(r"\.(OpenDoc6|OpenDoc7)\b", l):
            window = "\n".join(code[i - 1:i + 4])
            if not re.search(r"Is\s+Nothing", window, re.I):
                add("WARN", i, "OpenDoc result not checked for Nothing.")

    non_ascii = any(ord(c) > 127 for c in text)
    for level, lineno, msg in issues:
        where = f"{path.name}:{lineno}" if lineno else path.name
        print(f"{level:5} {where}: {msg}")
    if non_ascii:
        print(f"INFO  {path.name}: contains non-ASCII characters - import a cp1250 copy (--cp1250) or paste via clipboard.")
    if not issues:
        print(f"OK    {path.name}: no issues found.")
    return sum(1 for lvl, _, _ in issues if lvl == "ERROR")


def to_cp1250(path: Path) -> None:
    raw = path.read_bytes()
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        print(f"INFO  {path.name}: not UTF-8, assuming it is already cp1250.")
        return
    out = path.with_name(path.stem + "_cp1250" + path.suffix)
    # VBA editor expects CRLF line endings
    try:
        data = text.replace("\r\n", "\n").replace("\n", "\r\n").encode("cp1250")
    except UnicodeEncodeError:
        bad = sorted({c for c in text if c.encode("cp1250", errors="ignore") == b""})
        print(f"ERROR {path.name}: characters not representable in cp1250: {''.join(bad)!r} - "
              "replace them (or use ChrW) instead of letting them turn into '?'. No file written.")
        return
    out.write_bytes(data)
    print(f"INFO  wrote {out}")


def main(argv):
    args = [a for a in argv if not a.startswith("--")]
    if not args:
        print(__doc__)
        return 2
    errors = 0
    for a in args:
        p = Path(a)
        errors += check(p)
        if "--cp1250" in argv:
            to_cp1250(p)
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
