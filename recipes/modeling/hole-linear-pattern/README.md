# Cylindrical through hole and feature linear pattern

`hole_linear_pattern.bas` is a complete original VBA module. **LIVE_SMOKE_TESTED** for the controlled positive case only.
Target SOLIDWORKS 2018+ desktop VBA; L1 and installed 2026 TLB checks pass. Import with
SOLIDWORKS and SOLIDWORKS Constant references; run `main`, read the Immediate window.
Input: a fresh disposable [sketch-extrude](../sketch-extrude/README.md) result, one
configuration, initial volume 8e-6 m³, no cylindrical faces, outside sketch edit.
Do not run the modification/configuration recipes first. No preselection is needed.

Creates a diameter 4 mm circle at sketch (5,10) mm on the same base reference plane,
cuts through the 10 mm block, then uses `CreateDefinition(swFmLPattern)`/`CreateFeature`
to create three total instances with 10 mm pitch. Direction is an actual horizontal
segment of `Recipe_Profile`, found using endpoint coordinates; its sign determines the
reverse flag. Selection marks: direction=1, seed feature=4. No localized edge names,
screen coordinates, or feature indices are used to find that direction.

Units: all lengths in meters; coordinates are local to the first reference plane's sketch.
The cut reverses its default direction because cuts point opposite the sketch normal and
the generated boss points along it. Expected holes: centers (5,10), (15,10), (25,10) mm.
Expected readback: one solid, three cylindrical faces, pattern instances=3, volume
`8e-6 - 3*pi*(0.002)^2*0.01` m³ (about 7623.008882 mm³), and
`RESULT hole_linear_pattern OK...`. The macro checks seed volume/count before patterning
and checks final volume/count plus fresh pattern FeatureData after rebuilding.

This is an ordinary cylindrical cut, with no Hole Wizard standard, fit, counterbore,
thread or manufacturing certification. For a standard hole, select a named standard/size
from the installed Hole Wizard database and verify its definition; do not infer those
metadata from a diameter or substitute guessed enum values.

No saves/closes; `AddToDB` restored; selection cleared on success. Errors leave partial
geometry in the disposable model and emit FAIL. Hole sketch is intentionally underdefined:
the fully constrained parametric base profile is in the preceding recipe.

Negative cases: missing recipe profile; modified block/holes already present; active sketch;
multiple configurations; direction segment absent. Each must FAIL before success. For cut
direction, pitch/overlap and missing-seed failures, edit constants/calls in a private test
copy and require volume/count mismatch or an explicit API failure; never accept non-Nothing
alone. Inspect actual hole centers on the live fixture; aggregate volume is not a position proof.

Official contracts (consulted 2026-11-01; original implementation):

- [CreateCircleByRadius(XC,YC,Zc,Radius)](https://help.solidworks.com/2024/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isketchmanager~createcirclebyradius.html)
- [FeatureCut4: exact 27-argument order and cut normal](https://help.solidworks.com/2024/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ifeaturemanager~featurecut4.html)
- [ILinearPatternFeatureData: creation lifecycle and marks](https://help.solidworks.com/2023/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ILinearPatternFeatureData.html)
- [D1Axis: a sketch line is an accepted direction](https://help.solidworks.com/2022/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ILinearPatternFeatureData~D1Axis.html)
- [IBody2.GetMassProperties](https://help.solidworks.com/2020/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ibody2~getmassproperties.html)

No external macro code copied. Compilation, execution, positions, negative cases and UI
language independence remain live acceptance work; installed API presence is not certification.

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

Bounded live evidence (2026-11-01): SOLIDWORKS 2026 SP4.0, API 34.4.0, Polish UI; reviewed CLI/COM execution on a disposable copy, no dialogs, no input save, final session IDLE. See [case report](../../../validation/2026-11-01/modeling-final/hole_linear/runner.json) and [scope/limits](../../../validation/2026-11-01/README.md). Whole-project compilation, other releases and unexecuted negative cases remain NOT_TESTED. Source headers predate this recorded run.
