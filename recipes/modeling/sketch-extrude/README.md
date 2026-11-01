# Parametric rectangle and blind extrusion

`sketch_extrude.bas` is a complete original VBA module. **LIVE_SMOKE_TESTED** for the controlled positive case only.
Target: SOLIDWORKS 2018+ desktop VBA; installed 2026 type-library check passed.
Import with SOLIDWORKS and SOLIDWORKS Constant library references, run `main`, inspect the
Immediate window. Use a disposable blank part (saved private copy or new unsaved document),
one configuration, exactly three reference planes, no solid/sheet bodies or user sketches,
outside sketch edit. The runner profile must explicitly allow model changes.

It selects the first `RefPlane` object by invariant feature type, creates a corner rectangle
in that plane's **sketch coordinates**, fixes one corner and asks FullyDefineSketch to add
horizontal/vertical relations and driving dimensions. It verifies `swFullyConstrained`
instead of treating FullyDefineSketch's undefined return value as success. It then creates
the blind boss and reacquires body mass properties after a rebuild.

Inputs: `WIDTH_M=0.04`, `HEIGHT_M=0.02`, `DEPTH_M=0.01`; all must be positive, in meters.
Lengths in the API are meters, angles radians, volume m³, independent of document units.
Expected: `Recipe_Profile`, `Recipe_Block`, one solid, volume **0.000008 m³ = 8000 mm³**,
fully constrained profile and `RESULT sketch_extrude OK...`. Volume absolute tolerance is
1e-11 m³; dimension/constraint geometry still requires live acceptance.

Optional `RESULT_PATH=""` disables file output. Set it to a fresh file in the reviewed
private test directory for runner `--expect-text <path> "RESULT sketch_extrude OK"`.
The macro refuses an existing result file, creates a new log, and writes OK only after
its assertions and protected cleanup; exceptions write FAIL when the log was opened. No directories are created.
It never saves/closes the CAD document. `AddToDB` is restored; errors may leave a partial
sketch/feature in the disposable document for inspection. Selection is cleared on success.

Negative tests, each on a fresh copy: no document/wrong type; pre-existing body or sketch;
four reference planes; multiple configurations; active sketch; zero/negative size; unwritable
or existing RESULT_PATH. Every rejected input must produce FAIL and no OK. Do not rerun
against a partially completed fixture. Change units to inches and repeat: SI volume and
geometry must match. English/Polish runs must create identical dimensions without plane names.

Official contracts (consulted 2026-11-01; original code, no copied external macros):

- [CreateCornerRectangle(X1,Y1,Z1,X2,Y2,Z2)](https://help.solidworks.com/2018/English/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.ISketchManager~CreateCornerRectangle.html)
- [SketchAddConstraints(IdStr)](https://help.solidworks.com/2022/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.imodeldoc2~sketchaddconstraints.html)
- [FullyDefineSketch: ten parameters and undefined return status](https://help.solidworks.com/2015/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.ISketchManager~FullyDefineSketch.html)
- [Sketch constraint result enum](https://help.solidworks.com/2024/english/api/swconst/SolidWorks.Interop.swconst~SolidWorks.Interop.swconst.swConstrainedStatus_e.html)
- [FeatureExtrusion3: 23 parameters](https://help.solidworks.com/2024/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ifeaturemanager~featureextrusion3.html)
- [IBody2.GetMassProperties(Density): solid volume at index 3](https://help.solidworks.com/2020/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ibody2~getmassproperties.html)
- [ISketchSegment.GetSketch](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchSegment~GetSketch.html)

L1 lint and installed member/arity/constant checks pass. Those checks do not compile VBA,
prove the solver's constraints or certify a SOLIDWORKS release/service pack.

Bounded live evidence (2026-11-01): SOLIDWORKS 2026 SP4.0, API 34.4.0, Polish UI; reviewed CLI/COM execution on a disposable copy, no dialogs, no input save, final session IDLE. See [case report](../../../validation/2026-11-01/modeling-final/sketch/runner.json) and [scope/limits](../../../validation/2026-11-01/README.md). Whole-project compilation, other releases and unexecuted negative cases remain NOT_TESTED. Source headers predate this recorded run.
