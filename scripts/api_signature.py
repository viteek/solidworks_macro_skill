#!/usr/bin/env python3
"""Print the exact signature of SOLIDWORKS API members from the installed type library.

Feature-creation methods take 15-30 positional arguments (FeatureExtrusion3 has 23) - writing them
from memory is the most common source of wrong-argument bugs. Look them up instead:

    python api_signature.py FeatureManager.FeatureExtrusion3
    python api_signature.py SketchManager.CreateCenterRectangle ModelDocExtension.SaveAs3
    python api_signature.py FeatureManager.FeatureExtrusion*      # wildcard: all revisions
    python api_signature.py --find SheetMetalBaseFlange           # which interfaces have it

Interface names may be given with or without the leading "I" (FeatureManager / IFeatureManager).
Prints parameter names, VBA types and optional/ByRef flags, and the return type.
Windows + pywin32; reads sldworks.tlb / swconst.tlb (SOLIDWORKS need not be running).
"""
import argparse
import fnmatch
import glob
import os
import sys

try:
    import pythoncom
except ImportError:
    print("pywin32 is required: pip install pywin32")
    sys.exit(2)

VT = {2: "Integer", 3: "Long", 4: "Single", 5: "Double", 6: "Currency", 7: "Date", 8: "String",
      9: "Object", 11: "Boolean", 12: "Variant", 13: "Unknown", 16: "Byte", 17: "Byte", 18: "Integer",
      19: "Long", 20: "LongLong", 21: "LongLong", 22: "Long", 23: "Long", 24: "(none)", 25: "HRESULT"}
INVOKE = {1: "method", 2: "property get", 4: "property let", 8: "property set"}
PARAMFLAG_FOUT, PARAMFLAG_FRETVAL, PARAMFLAG_FOPT = 2, 8, 16


def type_name(ti, tdesc):
    vt = tdesc[0] if isinstance(tdesc, tuple) else tdesc
    if vt == 26 and isinstance(tdesc, tuple):            # VT_PTR
        return type_name(ti, tdesc[1])
    if vt == 29 and isinstance(tdesc, tuple):            # VT_USERDEFINED
        try:
            return ti.GetRefTypeInfo(tdesc[1]).GetDocumentation(-1)[0]
        except pythoncom.com_error:
            return "?"
    if vt == 27 and isinstance(tdesc, tuple):            # VT_SAFEARRAY
        return type_name(ti, tdesc[1]) + "()"
    return VT.get(vt, f"VT{vt}")


def load(sw_dir):
    libs = []
    for name in ("sldworks.tlb", "swconst.tlb"):
        p = os.path.join(sw_dir, name)
        if os.path.exists(p):
            libs.append(pythoncom.LoadTypeLib(p))
    if not libs:
        raise SystemExit(f"No sldworks.tlb in {sw_dir}")
    interfaces = {}
    for tlb in libs:
        for i in range(tlb.GetTypeInfoCount()):
            ti = tlb.GetTypeInfo(i)
            if ti.GetTypeAttr().typekind in (3, 4):        # interface / dispatch
                interfaces[tlb.GetDocumentation(i)[0]] = ti
    return interfaces


def describe(iname, ti, pattern):
    out = []
    ta = ti.GetTypeAttr()
    for k in range(ta.cFuncs):
        fd = ti.GetFuncDesc(k)
        names = ti.GetNames(fd.memid)
        if not fnmatch.fnmatchcase(names[0].lower(), pattern.lower()):
            continue
        params, ret = [], type_name(ti, fd.rettype[0])
        for idx, arg in enumerate(fd.args):
            flags = arg[1]
            pname = names[idx + 1] if idx + 1 < len(names) else f"arg{idx}"
            tname = type_name(ti, arg[0])
            if flags & PARAMFLAG_FRETVAL:
                ret = tname
                continue
            mark = ("ByRef " if (flags & PARAMFLAG_FOUT) or (isinstance(arg[0], tuple) and arg[0][0] == 26
                                                             and type_name(ti, arg[0]) in VT.values()) else "")
            opt = "Optional " if flags & PARAMFLAG_FOPT else ""
            params.append(f"{opt}{mark}{pname} As {tname}")
        kind = INVOKE.get(fd.invkind, "?")
        lines = [f"{iname}.{names[0]}  [{kind}]  -> {ret}   ({len(params)} args)"]
        lines += [f"    {i + 1:>2}. {p}" for i, p in enumerate(params)]
        out.append("\n".join(lines))
    return out


def default_sw_dir():
    for c in glob.glob(r"C:\Program Files\SOLIDWORKS Corp*\SOLIDWORKS"):
        if os.path.exists(os.path.join(c, "sldworks.tlb")):
            return c
    return r"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("members", nargs="*", help="Interface.Member (wildcards allowed in Member)")
    ap.add_argument("--find", help="list interfaces that have a member matching this text")
    ap.add_argument("--sw-dir", default=default_sw_dir())
    a = ap.parse_args()
    interfaces = load(a.sw_dir)
    if a.find:
        hits = set()
        for iname, ti in interfaces.items():
            ta = ti.GetTypeAttr()
            for k in range(ta.cFuncs):
                n = ti.GetNames(ti.GetFuncDesc(k).memid)[0]
                if a.find.lower() in n.lower():
                    hits.add(f"{iname}.{n}")
        print("\n".join(sorted(hits)) or "no match")
        return 0
    status = 0
    for spec in a.members:
        if "." not in spec:
            print(f"use Interface.Member: {spec}")
            status = 1
            continue
        iface, member = spec.split(".", 1)
        cands = [n for n in interfaces if n.lower() in (iface.lower(), "i" + iface.lower())]
        if not cands:
            print(f"unknown interface {iface}")
            status = 1
            continue
        found = []
        for n in cands:
            found += describe(n, interfaces[n], member)
        print("\n\n".join(found) if found else f"{spec}: no such member")
        print()
    return status


if __name__ == "__main__":
    sys.exit(main())
