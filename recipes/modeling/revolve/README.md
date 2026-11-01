# Full-turn annular revolve

`revolve_ring.bas` is a complete original VBA module. **LIVE_SMOKE_TESTED** for the controlled positive case only.
Target SOLIDWORKS 2018+ desktop VBA; L1 and installed 2026 TLB checks pass. Import with
SOLIDWORKS and SOLIDWORKS Constant references; run `main`, inspect the Immediate window.
Input: a fresh disposable blank part outside sketch edit, no solids. It selects the first
reference plane by invariant type, creates a rectangular profile on one side of its Y axis,
and a construction centerline as the revolve axis. Use a new copy for every attempt.

Profile in sketch meters: x=0.010..0.020, y=0..0.010; axis is x=0. The sketch is deliberately
underdefined. After leaving edit mode it reacquires the construction line before selection,
selects sketch mark 0 and axis mark 4, and calls the 20-argument FeatureRevolve2 for one
solid, full turn. `8*Atn(1)` is **2*pi radians**, independent of degrees display units.

Expected: `Recipe_RevolveProfile`, `Recipe_Ring`, one solid and volume
`pi*(0.020^2 - 0.010^2)*0.010` m³ = about **9424.777961 mm³**, then
`RESULT revolve_ring OK...`. The volume is checked after rebuild with absolute tolerance
1e-11 m³. Ring axis follows local sketch Y, not a promised global axis on every template.
Inspect inner/outer radii and height in the live result in addition to aggregate volume.

No saves/closes; AddToDB restored; selection cleared on success. Failure leaves partial
geometry for inspection and emits FAIL. Negative cases: wrong/no document; active sketch;
pre-existing solid; remove axis selection in a private source copy; move profile across
the axis or use invalid dimensions. Require explicit API/assertion failure and no OK.

Official contracts (consulted 2026-11-01; original code):

- [CreateCenterLine(X1,Y1,Z1,X2,Y2,Z2)](https://help.solidworks.com/2024/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isketchmanager~createcenterline.html)
- [FeatureRevolve2: 20 parameters and selection marks](https://help.solidworks.com/2024/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ifeaturemanager~featurerevolve2.html)
- [ISketchSegment.ConstructionGeometry](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchSegment~ConstructionGeometry.html)
- [IBody2.GetMassProperties](https://help.solidworks.com/2020/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ibody2~getmassproperties.html)

No external examples copied. VBA compilation, execution, axis/geometry checks and negative
integration cases remain unverified.

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

Bounded live evidence (2026-11-01): SOLIDWORKS 2026 SP4.0, API 34.4.0, Polish UI; reviewed CLI/COM execution on a disposable copy, no dialogs, no input save, final session IDLE. See [case report](../../../validation/2026-11-01/modeling-final/revolve/runner.json) and [scope/limits](../../../validation/2026-11-01/README.md). Whole-project compilation, other releases and unexecuted negative cases remain NOT_TESTED. Source headers predate this recorded run.
