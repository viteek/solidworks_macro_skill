# Four-hole circular feature pattern

`circular_pattern.bas` is a complete original VBA module. **LIVE_SMOKE_TESTED** for the controlled positive case only.
Target SOLIDWORKS 2018+ desktop VBA; L1 and installed 2026 TLB checks pass. Import with
SOLIDWORKS and SOLIDWORKS Constant references; run `main`, inspect Ctrl+G.
Input: fresh disposable blank part, exactly three reference planes, one configuration,
no solid/sheet bodies or user sketches, outside sketch edit. No preselection required.

Creates a diameter **40 mm**, thickness **10 mm** disk on the first reference plane,
then a diameter **4 mm** axial through hole at local sketch (10,0) mm. It reacquires the
post-cut body's edges and finds a circular curve of radius 0.020 m. That actual outer
edge defines the rotation center/axis; mark 1 selects direction, mark 4 the cut seed.
CreateDefinition(swFmCirPattern) and CreateFeature create four instances including seed,
with **pi/2 radians between instances**, EqualSpacing=False, Direction2=False.
All sketch lengths are meters; volume is m³; the two circle sketches are underdefined.

Expected: one solid, Recipe_Disk, Recipe_DiskHole, Recipe_FourDiskHoles, four holes around
a radius **10 mm** bolt circle, total instances=4, pitch=pi/2, volume
`pi*(0.020^2 - 4*0.002^2)*0.010` m³ = about **12063.715790 mm³**, then
`RESULT circular_pattern OK...`. It asserts disk and seed volumes, then rebuilt final
volume and fresh feature count/angle. Inspect centers and radii live: aggregate volume
and instance count do not independently prove all positions or topology.

No saves/closes; restores AddToDB; clears selection on success. Failure leaves partial
geometry for inspection. Negative cases: nonblank part; wrong/no document; active sketch;
four reference planes; overlapping holes after changing pitch/count in a private test copy;
missing direction selection. Require failure/mismatch and no OK, never a non-Nothing-only pass.

Official contracts (consulted 2026-11-01; original implementation):

- [ICircularPatternFeatureData: creation lifecycle/marks](https://help.solidworks.com/2020/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ICircularPatternFeatureData.html)
- [Official circular feature example: edge direction selection](https://help.solidworks.com/2026/English/api/sldworksapi/Create_Bidirectional_Circular_Pattern_Feature_Example_VB.htm)
- [TotalInstances: includes total in direction 1](https://help.solidworks.com/2023/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ICircularPatternFeatureData~TotalInstances.html)
- [Spacing: angle in radians](https://help.solidworks.com/2021/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ICircularPatternFeatureData~Spacing.html)
- [ICurve.CircleParams: radius at index 6](https://help.solidworks.com/2021/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ICurve~CircleParams.html)
- [FeatureCut4](https://help.solidworks.com/2024/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ifeaturemanager~featurecut4.html)

No copied example code. Project compilation, live execution, geometry positions and
negative cases remain separate, unverified acceptance levels.

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

Bounded live evidence (2026-11-01): SOLIDWORKS 2026 SP4.0, API 34.4.0, Polish UI; reviewed CLI/COM execution on a disposable copy, no dialogs, no input save, final session IDLE. See [case report](../../../validation/2026-11-01/modeling-final/circular/runner.json) and [scope/limits](../../../validation/2026-11-01/README.md). Whole-project compilation, other releases and unexecuted negative cases remain NOT_TESTED. Source headers predate this recorded run.
