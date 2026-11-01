# Plane selection by type

`plane_selection.bas` is a complete original VBA module. **LIVE_SMOKE_TESTED** for the controlled positive case only.
Target: SOLIDWORKS 2018+ desktop VBA; checked against the installed 2026 type library.
Import into a macro with SOLIDWORKS and SOLIDWORKS Constant type-library references.
Run `main`; read the VBA Immediate window (Ctrl+G).

Input: an active disposable part outside sketch edit, containing exactly three reference
plane features. No preselection is required. It enumerates `FirstFeature/GetNextFeature`,
compares the invariant `GetTypeName2` identifier `RefPlane`, and selects the first matching
**object** with mark 0. It does not search for localized names such as Front Plane.
The selected plane is deliberately an ordinal in this restricted fixture; this is not a
general promise that the first plane of any imported or reorganized model is the XY plane.

Expected: exactly one selected `swSelDATUMPLANES` object and
`RESULT plane_selection OK; selected=<actual localized name>; mark=0`.
It changes the selection, not geometry, and never saves/closes documents. Units: none.
It does not restore the previous selection because the selected plane is its result.

Negative cases: no document; assembly/drawing; sketch edit mode; add a fourth reference
plane. These must produce `FAIL`, without an OK line. The fourth-plane case is rejected
before clearing the prior selection. A selection failure after clearing may leave it empty.
Live acceptance requires running the same fixture in English and Polish, checking type/count,
and recording version, service pack and UI language. Static checks cannot establish that.

Official API contracts (consulted 2026-11-01; original implementation, no external code copied):

- [IFeature.GetTypeName2](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IFeature~GetTypeName2.html)
- [IFeature.Select2(Append, Mark)](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IFeature~Select2.html)
- [ISelectionMgr.GetSelectedObjectType3(Index, Mark)](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISelectionMgr~GetSelectedObjectType3.html)
- [IModelDoc2.FirstFeature](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDoc2~FirstFeature.html)

Evidence: `check_vba.py` and `check_api.py --sw-dir "C:\Program Files\Dassault Systemes\SOLIDWORKS 3DEXPERIENCE R2026x\SOLIDWORKS"` pass.
Compilation, execution and result checks remain separate evidence levels.

## Optional result log and cleanup

`RESULT_PATH=""` disables file output. Set it to a fresh result file in the reviewed
private run directory when using runner result assertions. The macro refuses an existing
file and never creates directories. It writes its final `RESULT ... OK` only after the
model assertions and protected cleanup; caught model/assertion/cleanup failures write
`FAIL` when the log was opened. Existing/unwritable log paths must fail, without overwriting
prior output. An absent expected result file is a failed test even if RunMacro2 returns True.
No CAD save or close is added by this log option.

The Failed label captures the original error and uses `Resume Cleanup` before activating
`On Error Resume Next` for cleanup. Each cleanup call is followed by an Err check; cleanup
failures suppress overall OK. Log errors use their own handler and `Resume LogCleanup`.
This matters because VBA cannot handle a second exception inside an already active handler.
Negative acceptance should deliberately trigger an input failure with a writable fresh log
and verify that it contains FAIL, then exercise existing/unwritable result paths separately.

Bounded live evidence (2026-11-01): SOLIDWORKS 2026 SP4.0, API 34.4.0, Polish UI; reviewed CLI/COM execution on a disposable copy, no dialogs, no input save, final session IDLE. See [case report](../../../validation/2026-11-01/modeling-final/plane/runner.json) and [scope/limits](../../../validation/2026-11-01/README.md). Whole-project compilation, other releases and unexecuted negative cases remain NOT_TESTED. Source headers predate this recorded run.
