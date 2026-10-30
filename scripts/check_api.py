#!/usr/bin/env python3
"""Verify the SOLIDWORKS API calls in a VBA macro against the installed type libraries.

Catches what a plain linter can't: invented methods/properties, wrong argument counts and
misspelled sw* constants - the errors VBA reports as "Method or data member not found",
"Wrong number of arguments" or "Variable not defined", which a .swb macro hides completely.
Also flags undeclared variables that are assigned to (Option Explicit compile errors).

Windows + Python 3 + pywin32; reads sldworks.tlb / swconst.tlb from the SOLIDWORKS install folder
(SOLIDWORKS doesn't need to be running).

Usage:
    python check_api.py macro.bas [more.bas ...] [--sw-dir "C:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS"]

Checks variables declared with SOLIDWORKS types (Dim swModel As SldWorks.ModelDoc2 / As ModelDoc2)
and follow member chains through typed return values (swModel.Extension.SaveAs ...). Members
reached through Variant/Object values can't be checked - prefer early binding.
Exit code 1 if any ERROR was found.
"""
import argparse
import glob
import os
import re
import sys

try:
    import pythoncom
except ImportError:
    print("pywin32 is required: pip install pywin32")
    sys.exit(2)

TKIND_ENUM, TKIND_INTERFACE, TKIND_DISPATCH, TKIND_COCLASS = 0, 3, 4, 5
VT_PTR, VT_USERDEFINED = 26, 29
INVOKE_PROPERTYPUT, INVOKE_PROPERTYPUTREF = 4, 8
PARAMFLAG_FOPT, PARAMFLAG_FRETVAL = 16, 8

VBA_KEYWORDS = {
    "if", "then", "else", "elseif", "end", "for", "each", "in", "next", "do", "loop", "while", "wend",
    "until", "select", "case", "with", "set", "let", "call", "exit", "sub", "function", "property",
    "dim", "redim", "preserve", "private", "public", "static", "const", "as", "new", "and", "or",
    "not", "xor", "mod", "is", "like", "to", "step", "byval", "byref", "optional", "paramarray",
    "true", "false", "nothing", "empty", "null", "me", "on", "error", "goto", "resume", "option",
    "explicit", "debug", "erase", "open", "close", "print", "output", "input", "append", "binary",
    "lock", "get", "put", "type", "enum", "attribute", "mid", "rem",
}


class TypeLibIndex:
    def __init__(self, sw_dir):
        self.types = {}      # lower name -> (display name, {lower member: [funcdesc-like dicts]})
        self.consts = {}     # lower const -> (enum name, display name)
        self.enums = {}      # lower enum name -> {lower member}
        for name in ("sldworks.tlb", "swconst.tlb", "swcommands.tlb"):
            path = os.path.join(sw_dir, name)
            if os.path.exists(path):
                self._load(pythoncom.LoadTypeLib(path))
        if not self.types:
            raise SystemExit(f"No sldworks.tlb in {sw_dir}")
        self.owners = {}     # lower member -> interface names that have it
        for _, (tname, members) in self.types.items():
            for k in members:
                self.owners.setdefault(k, set()).add(tname)

    def _members(self, ti, seen=None):
        seen = seen or set()
        ta = ti.GetTypeAttr()
        out = {}
        for i in range(ta.cFuncs):
            fd = ti.GetFuncDesc(i)
            nm = ti.GetNames(fd.memid)[0]
            args = [a for a in fd.args if not (a[1] & PARAMFLAG_FRETVAL)]
            nopt = sum(1 for a in args if a[1] & PARAMFLAG_FOPT) or fd.cParamsOpt
            ret = None
            rt = fd.rettype[0]
            if isinstance(rt, tuple) and rt[0] == VT_PTR and isinstance(rt[1], tuple) and rt[1][0] == VT_USERDEFINED:
                try:
                    ret = ti.GetRefTypeInfo(rt[1][1]).GetDocumentation(-1)[0]
                except pythoncom.com_error:
                    ret = None
            out.setdefault(nm.lower(), []).append({
                "name": nm, "nargs": len(args), "nopt": nopt, "kind": fd.invkind, "ret": ret})
        for j in range(ta.cImplTypes):
            try:
                base = ti.GetRefTypeInfo(ti.GetRefTypeOfImplType(j))
            except pythoncom.com_error:
                continue
            bname = base.GetDocumentation(-1)[0]
            if bname in ("IDispatch", "IUnknown") or bname in seen:
                continue
            seen.add(bname)
            for k, v in self._members(base, seen).items():
                out.setdefault(k, v)
        return out

    def _load(self, tlb):
        for i in range(tlb.GetTypeInfoCount()):
            ti = tlb.GetTypeInfo(i)
            ta = ti.GetTypeAttr()
            name = tlb.GetDocumentation(i)[0]
            if ta.typekind == TKIND_ENUM:
                members = set()
                for v in range(ta.cVars):
                    vn = ti.GetNames(ti.GetVarDesc(v).memid)[0]
                    members.add(vn.lower())
                    self.consts[vn.lower()] = (name, vn)
                self.enums[name.lower()] = members
            elif ta.typekind in (TKIND_DISPATCH, TKIND_INTERFACE):
                self.types[name.lower()] = (name, self._members(ti))
        # coclass names (ModelDoc2, SldWorks...) -> their default interface (IModelDoc2...)
        for i in range(tlb.GetTypeInfoCount()):
            ti = tlb.GetTypeInfo(i)
            ta = ti.GetTypeAttr()
            if ta.typekind != TKIND_COCLASS:
                continue
            name = tlb.GetDocumentation(i)[0]
            for j in range(ta.cImplTypes):
                if ti.GetImplTypeFlags(j) & 2:   # source (events) interface
                    continue
                iname = ti.GetRefTypeInfo(ti.GetRefTypeOfImplType(j)).GetDocumentation(-1)[0]
                if iname.lower() in self.types:
                    self.types.setdefault(name.lower(), self.types[iname.lower()])
                    break

    def resolve(self, typename):
        t = typename.lower()
        for prefix in ("sldworks.", "swconst."):
            if t.startswith(prefix):
                t = t[len(prefix):]
        return self.types.get(t)


def blank_strings(line):
    """Replace string literal contents with spaces and drop the trailing comment."""
    out, in_str = [], False
    for ch in line:
        if ch == '"':
            in_str = not in_str
            out.append(ch)
        elif in_str:
            out.append(" ")
        elif ch == "'":
            break
        else:
            out.append(ch)
    s = "".join(out)
    return "" if re.match(r"\s*Rem\s", s, re.I) else s


def join_continuations(raw_lines):
    """Yield (first_line_number, logical_line) with ' _' continuations joined."""
    buf, start = "", None
    for i, l in enumerate(raw_lines, 1):
        code = blank_strings(l)
        if start is None:
            start = i
        if re.search(r"\s_\s*$", code):
            buf += code.rstrip()[:-1] + " "
            continue
        buf += code
        yield start, buf
        buf, start = "", None
    if buf:
        yield start, buf


def split_args(s):
    """Split a top-level argument list. Returns list of arg strings (empty strings for skipped)."""
    args, depth, cur = [], 0, ""
    for ch in s:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == "," and depth == 0:
            args.append(cur.strip())
            cur = ""
        else:
            cur += ch
    if cur.strip() or args:
        args.append(cur.strip())
    return args


DECL_RE = re.compile(r"^\s*(?:Dim|Private|Public|Static|Global)\s+(?!Sub\b|Function\b|Const\b|Declare\b|Type\b|Enum\b|Property\b)(.+)$", re.I)
PROC_RE = re.compile(r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)\s*\((.*?)\)\s*(?:As\s+([\w.]+))?", re.I)
CONST_RE = re.compile(r"^\s*(?:Private\s+|Public\s+|Global\s+)?Const\s+(.+)$", re.I)
VAR_ITEM_RE = re.compile(r"^(?:WithEvents\s+)?(\w+)[$%&!#@]?\s*(\([^)]*\))?\s*(?:As\s+(?:New\s+)?([\w.]+))?", re.I)


def parse_decl_list(text):
    out = {}
    for part in split_args(text):
        m = VAR_ITEM_RE.match(part.strip())
        if m:
            out[m.group(1).lower()] = (m.group(3) or "Variant", bool(m.group(2)))
    return out


def check(path, idx):
    with open(path, "rb") as source:
        raw = source.read()
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = raw.decode("cp1250")
    raw_lines = text.splitlines()
    logical = list(join_continuations(raw_lines))
    explicit = any(re.match(r"\s*Option\s+Explicit\b", l, re.I) for _, l in logical)

    # pass 1: module-level names, procedures, consts, enums/types
    module_vars, procs, consts, user_types = {}, set(), set(), set()
    in_proc, in_block = False, False
    for _, l in logical:
        if re.match(r"\s*(?:Public\s+|Private\s+)?(Enum|Type)\s+(\w+)", l, re.I):
            user_types.add(re.match(r"\s*(?:Public\s+|Private\s+)?(?:Enum|Type)\s+(\w+)", l, re.I).group(1).lower())
            in_block = True
            continue
        if in_block:
            if re.match(r"\s*End\s+(Enum|Type)", l, re.I):
                in_block = False
            else:
                m = re.match(r"\s*(\w+)", l)
                if m:
                    consts.add(m.group(1).lower())   # enum members are usable bare
            continue
        pm = PROC_RE.match(l)
        if pm:
            procs.add(pm.group(2).lower())
            in_proc = True
            continue
        if re.match(r"\s*End\s+(Sub|Function|Property)\b", l, re.I):
            in_proc = False
            continue
        cm = CONST_RE.match(l)
        if cm:
            for part in split_args(cm.group(1)):
                n = re.match(r"(\w+)", part.strip())
                if n:
                    consts.add(n.group(1).lower())
            continue
        dm = re.match(r"^\s*Declare\s|^\s*(?:Public|Private)\s+Declare\s", l, re.I)
        if dm:
            n = re.search(r"Declare\s+(?:PtrSafe\s+)?(?:Function|Sub)\s+(\w+)", l, re.I)
            if n:
                procs.add(n.group(1).lower())
            continue
        if not in_proc:
            m = DECL_RE.match(l)
            if m:
                module_vars.update(parse_decl_list(m.group(1)))

    issues = []

    def add(level, ln, msg):
        issues.append((level, ln, msg))

    def type_of(name, local):
        t = local.get(name.lower()) or module_vars.get(name.lower())
        return t[0] if t else None

    def check_chain(ln, line, start, root_type, local, statement_start):
        """Walk .member chain starting at index `start` (pointing at '.') with root_type."""
        info = idx.resolve(root_type) if root_type else None
        pos = start
        while info and pos < len(line) and line[pos] == ".":
            m = re.match(r"\.(\w+)", line[pos:])
            if not m:
                return
            member = m.group(1)
            pos += m.end()
            tname, members = info
            cands = members.get(member.lower())
            if not cands:
                owners = idx.owners.get(member.lower(), [])
                if owners:
                    add("WARN", ln, f"{tname} has no member '{member}' - it belongs to {', '.join(sorted(owners)[:4])}. "
                                    f"VBA resolves it at run time (error 438 if the object is another type); "
                                    f"prefer a variable of that type (e.g. Set swAssy = swModel).")
                else:
                    add("ERROR", ln, f"'{member}' does not exist on {tname} or any other SOLIDWORKS interface - "
                                     "invented or misspelled API (run-time error 438).")
                return
            getters = [c for c in cands if c["kind"] not in (INVOKE_PROPERTYPUT, INVOKE_PROPERTYPUTREF)] or cands
            f = getters[0]
            rest = line[pos:]
            # A member without parentheses or trailing arguments is a zero-argument
            # call/property access, not an unchecked call (e.g. ok = ext.SaveAs).
            argtext = ""
            pm = re.match(r"\s*\(", rest)
            if pm:
                depth, j = 0, pos + pm.end() - 1
                for k in range(j, len(line)):
                    if line[k] == "(":
                        depth += 1
                    elif line[k] == ")":
                        depth -= 1
                        if depth == 0:
                            argtext = line[j + 1:k]
                            pos = k + 1
                            break
                else:
                    return
            elif statement_start and not re.match(r"\s*(\.|=|$)", rest) and not re.match(r"\s*(Then|To|And|Or|Is|<|>|&|\+|-|\*|/)", rest, re.I):
                argtext = rest.strip()
                pos = len(line)
            if argtext is not None and f["kind"] not in (INVOKE_PROPERTYPUT, INVOKE_PROPERTYPUTREF):
                n = len(split_args(argtext)) if argtext.strip() else 0
                req = f["nargs"] - f["nopt"]
                # a property returning an array/collection may be indexed: allow extra args then
                if n > f["nargs"] and not (f["nargs"] == 0 and n == 1 and f["ret"] is None):
                    add("ERROR", ln, f"{tname}.{f['name']} takes {f['nargs']} argument(s), {n} given.")
                elif n < req:
                    add("ERROR", ln, f"{tname}.{f['name']} needs at least {req} argument(s), {n} given.")
            info = idx.resolve(f["ret"]) if f["ret"] else None

    local, with_stack, proc_name = {}, [], None
    for ln, l in logical:
        s = l.strip()
        if not s:
            continue
        pm = PROC_RE.match(l)
        if pm:
            local = parse_decl_list(pm.group(3)) if pm.group(3).strip() else {}
            local = {k: v for k, v in local.items()}
            # ByVal/ByRef/Optional prefixes
            local = {}
            for part in split_args(pm.group(3)):
                part = re.sub(r"^(Optional\s+|ByVal\s+|ByRef\s+|ParamArray\s+)+", "", part.strip(), flags=re.I)
                local.update(parse_decl_list(part))
            proc_name = pm.group(2).lower()
            continue
        if re.match(r"\s*End\s+(Sub|Function|Property)\b", l, re.I):
            local, proc_name, with_stack = {}, None, []
            continue
        m = DECL_RE.match(l)
        if m and proc_name:
            local.update(parse_decl_list(m.group(1)))
            continue
        if re.match(r"\s*(Dim|Const|Private|Public|Declare|Attribute|Option)\b", l, re.I):
            continue

        # With blocks
        wm = re.match(r"\s*With\s+(\w+)\s*$", l, re.I)
        if wm:
            with_stack.append(type_of(wm.group(1), local))
            continue
        if re.match(r"\s*End\s+With\b", l, re.I):
            if with_stack:
                with_stack.pop()
            continue

        # Strings are already blanked. Treat inline branches as statement boundaries,
        # so each bare call owns only its own arguments (also for implicit With receivers).
        for stmt in re.split(r":(?!=)|\b(?:Then|Else)\b", l, flags=re.I):
            st = stmt.strip()
            if not st or re.match(r"^\w+$", st) and st.lower() not in VBA_KEYWORDS and stmt is l:
                pass
            # undeclared assignment targets
            am = re.match(r"^(?:Set\s+|Let\s+)?(\w+)\s*(?:\([^=]*\))?\s*=(?!=)", st, re.I)
            fm = re.match(r"^For\s+(?:Each\s+)?(\w+)\b", st, re.I)
            target = (am or fm).group(1) if (am or fm) else None
            if target and explicit:
                tl = target.lower()
                if not (tl in local or tl in module_vars or tl in consts or tl in procs
                        or tl in VBA_KEYWORDS or tl == proc_name):
                    add("ERROR", ln, f"'{target}' is not declared (Option Explicit -> compile error 'Variable not defined').")
            # chains on typed variables
            stmt_head = re.match(r"^(?:Call\s+)?(\w+)\.", st, re.I)
            for cm in re.finditer(r"(?<![\w.])(\w+)(?=\.\w)", st):
                t = type_of(cm.group(1), local)
                if t:
                    check_chain(ln, st, cm.end(), t, local,
                                statement_start=bool(stmt_head) and cm.start() == st.find(stmt_head.group(1)))
            if with_stack and with_stack[-1]:
                for cm in re.finditer(r"(?:^|[\s(,=])(?=\.\w)", st):
                    check_chain(ln, st, cm.end(), with_stack[-1], local,
                                statement_start=cm.start() == 0 and st.startswith("."))
            # sw* constants
            for km in re.finditer(r"(?<![\w.])(sw[A-Z]\w*)(?:\.(\w+))?", st):
                name, member = km.group(1), km.group(2)
                nl = name.lower()
                if nl in local or nl in module_vars or nl in consts or nl in procs:
                    continue
                if member and nl in idx.enums:
                    if member.lower() not in idx.enums[nl]:
                        add("ERROR", ln, f"Enum {name} has no member '{member}'.")
                    continue
                if nl not in idx.consts and nl not in idx.enums:
                    add("ERROR", ln, f"'{name}' is not a SOLIDWORKS constant (swconst) nor a declared name.")

    for level, ln, msg in issues:
        print(f"{level:5} {os.path.basename(path)}:{ln}: {msg}")
    if not issues:
        print(f"OK    {os.path.basename(path)}: all checked API members, argument counts and constants exist.")
    return sum(1 for lv, _, _ in issues if lv == "ERROR")


def default_sw_dir():
    for c in glob.glob(r"C:\Program Files\SOLIDWORKS Corp*\SOLIDWORKS"):
        if os.path.exists(os.path.join(c, "sldworks.tlb")):
            return c
    return r"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="+")
    ap.add_argument("--sw-dir", default=default_sw_dir())
    a = ap.parse_args()
    idx = TypeLibIndex(a.sw_dir)
    errors = sum(check(f, idx) for f in a.files)
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
