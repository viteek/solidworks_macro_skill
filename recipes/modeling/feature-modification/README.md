# Extrusion FeatureData modification

`modify_extrude.bas` is a complete original VBA module. **LIVE_SMOKE_TESTED** for the controlled positive case only.
Target SOLIDWORKS 2018+ desktop VBA; L1 and installed 2026 TLB checks pass.
Import with SOLIDWORKS and SOLIDWORKS Constant references. Run `main`, inspect Ctrl+G.
Input: a fresh disposable result of [sketch-extrude](../sketch-extrude/README.md), one
configuration, untouched 10 mm blind `Recipe_Block`, no active sketch. Saved private
copies are accepted. It validates initial depth and one-body volume before modification.

`NEW_DEPTH_M=0.015` means 15 mm; accepted range >0 to 0.1 m. It gets the feature's
`ExtrudeFeatureData2`, checks the forward blind end condition, calls
`AccessSelections(model, Nothing)` in part context, sets depth, commits through
`ModifyDefinition`, rebuilds and reacquires a fresh definition for readback.
On a failed commit/exception after acquiring access it calls `ReleaseSelectionAccess`.
A successful ModifyDefinition releases access itself. No cached definition is trusted
after the commit.

Expected: forward depth **0.015 m**, one solid, volume **0.000012 m³**, and
`RESULT modify_extrude OK...`. It never saves/closes. Selection is cleared; previous selection
is not restored. Failure after commit leaves the partial change for inspection, without OK;
there is no automatic rollback or retry. Do not infer geometry health solely from the Boolean.

Negative cases: missing/renamed Recipe_Block; untouched fixture with a through-all end;
multiple configurations; active sketch; non-extrude using the same name; depth=0 or >0.1 m;
previously modified block or added holes causing initial volume mismatch. Each must FAIL.
A failed AccessSelections/ModifyDefinition requires a controlled integration fixture; these
branches are present but not live exercised. Initial depth/volume are fixture validation,
not a proof that arbitrary production parts are safe to edit.

Official contracts (consulted 2026-11-01; original implementation):

- [AccessSelections(TopDoc, Component)](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IExtrudeFeatureData2~AccessSelections.html)
- [GetEndCondition(Forward)](https://help.solidworks.com/2023/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IExtrudeFeatureData2~GetEndCondition.html)
- [SetDepth(Forward, Depth)](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IExtrudeFeatureData2~SetDepth.html)
- [ModifyDefinition(Data, TopDoc, Component)](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IFeature~ModifyDefinition.html)
- [Official explanation of access/commit/release lifecycle](https://help.solidworks.com/2016/English/api/sldworksapi/Access_Selections_Example_VB.htm)

No external macro code was copied. Compilation, solver execution, result assertions and
negative live cases remain unverified; static TLB evidence is not a runtime certification.

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

Bounded live evidence (2026-11-01): SOLIDWORKS 2026 SP4.0, API 34.4.0, Polish UI; reviewed CLI/COM execution on a disposable copy, no dialogs, no input save, final session IDLE. See [case report](../../../validation/2026-11-01/modeling-final/modify/runner.json) and [scope/limits](../../../validation/2026-11-01/README.md). Whole-project compilation, other releases and unexecuted negative cases remain NOT_TESTED. Source headers predate this recorded run.
