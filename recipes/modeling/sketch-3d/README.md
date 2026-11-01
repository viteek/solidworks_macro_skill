# Three-dimensional sketch path

`sketch_3d.bas` is a complete original VBA module. **LIVE_SMOKE_TESTED** for the controlled positive case only.
Target SOLIDWORKS 2018+ desktop VBA; L1 and installed 2026 TLB checks pass. Import with
SOLIDWORKS and SOLIDWORKS Constant references; run `main`, inspect Ctrl+G.
Input: a disposable part without solid bodies, outside sketch edit, without Recipe_3DPath.
Use a fresh blank fixture; the recipe does not need a plane or preselection.

It opens `Insert3DSketch(True)`, checks ActiveSketch and Is3D, creates three lines with
`AddToDB=True`, closes using Insert3DSketch, rebuilds and reacquires the specific sketch.
Expected `Recipe_3DPath`, Is3D nonzero, three straight segments and length **0.06 m**,
followed by `RESULT sketch_3d OK...`. Model-space endpoints, in meters, are:
`(0,0,0) -> (0.01,0,0) -> (0.01,0.02,0) -> (0.01,0.02,0.03)`.
These are absolute model coordinates, not coordinates transformed from a 2D sketch.
The sketch is intentionally underdefined. Equal endpoint coordinates are a geometric path;
this recipe does not claim merged-point constraints or a sweep-ready relation graph.

No save/close; no solid is added. AddToDB is restored; selection cleared on success.
On error, partial sketch/edit state remains for inspection and FAIL is printed.
Negative cases: wrong document/no document; active 2D/3D sketch; existing named path; existing
solid; force a zero-length segment in a private source copy. Require failure or assertion
mismatch, no OK. Inspect endpoint coordinates/continuity separately in a live acceptance run:
the summed-length assertion does not prove their positions or coincident constraints.

Official contracts (consulted 2026-11-01; no example code copied):

- [Insert3DSketch(UpdateEditRebuild)](https://help.solidworks.com/2016/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.ISketchManager~Insert3DSketch.html)
- [CreateLine(X1,Y1,Z1,X2,Y2,Z2)](https://help.solidworks.com/2019/English/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.ISketchManager~CreateLine.html)
- [ISketch.Is3D](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketch~Is3D.html)
- [ISketchSegment.GetLength](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchSegment~GetLength.html)

Static checks do not establish VBA compilation, execution, topology or output correctness.

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

Bounded live evidence (2026-11-01): SOLIDWORKS 2026 SP4.0, API 34.4.0, Polish UI; reviewed CLI/COM execution on a disposable copy, no dialogs, no input save, final session IDLE. See [case report](../../../validation/2026-11-01/modeling-final/sketch3d/runner.json) and [scope/limits](../../../validation/2026-11-01/README.md). Whole-project compilation, other releases and unexecuted negative cases remain NOT_TESTED. Source headers predate this recorded run.
