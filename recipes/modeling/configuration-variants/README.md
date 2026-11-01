# Configuration-specific dimension variants

`configuration_variants.bas` is a complete original VBA module. **LIVE_SMOKE_TESTED** for the controlled positive case only.
Target SOLIDWORKS 2018+ desktop VBA; L1 and installed 2026 TLB checks pass. Import with
SOLIDWORKS and SOLIDWORKS Constant references. Run `main`; inspect the Immediate window.
Input: a fresh disposable [sketch-extrude](../sketch-extrude/README.md) result, one original
configuration and untouched 10 mm block, outside sketch edit. Do not run modification or
hole-pattern first. The active configuration name is captured from the API, not `Default`.

Creates `Recipe_10mm` and `Recipe_20mm`, sets the generated extrusion's `D1` only in the
specified named configuration using `SetSystemValue3`, checks its **enum status**, then
activates/rebuilds each variant and checks fresh FeatureData plus solid volume. D1 is accepted
only after its current value matches the original extrusion depth. That fixture-specific
mapping is not a general rule that every feature's first dimension is its depth.

Units: depths 0.010/0.020 meters, volumes 8e-6/16e-6 m³. Expected: three configurations,
those two values/volumes, original configuration active again at 8e-6 m³, and
`RESULT configuration_variants OK...`. It never saves/closes or deletes configurations.
On failure it attempts to restore the original configuration, reports restore failures,
and leaves any partially created variants in the disposable document. No success is logged
unless original restoration and its geometry assertion also pass.

Negative cases: two existing configurations; missing/renamed block; wrong D1 mapping;
non-blind extrusion; holes or modified original; active sketch. Must FAIL with no OK.
Exercise an invalid variant name and a suppressed/absent dimension in a separate fixture
to test creation/setter failures. Restart on a fresh copy after any partial attempt.

Official contracts (consulted 2026-11-01; no copied example code):

- [IModelDoc2.AddConfiguration3(Name, Comment, AlternateName, Options)](https://help.solidworks.com/2016/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDoc2~AddConfiguration3.html)
- [IFeature.Parameter: short D1 name is valid in feature context](https://help.solidworks.com/2016/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IFeature~Parameter.html)
- [IDimension.SetSystemValue3(NewValue, WhichConfigurations, Config_names)](https://help.solidworks.com/2025/English/api/sldworksapi/SolidWorks.interop.sldworks~SolidWorks.interop.sldworks.IDimension~SetSystemValue3.html)
- [IDimension.GetSystemValue3: getter uses swInConfigurationOpts_e](https://help.solidworks.com/2017/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IDimension~GetSystemValue3.html)
- [IModelDoc2.ShowConfiguration2](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDoc2~ShowConfiguration2.html)

Setter options and getter options are different enums. ForceRebuild3 acts on the active
configuration, so merely iterating names without activation cannot verify each geometry.
VBA compilation and live per-configuration geometry/restoration remain unverified.

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

Bounded live evidence (2026-11-01): SOLIDWORKS 2026 SP4.0, API 34.4.0, Polish UI; reviewed CLI/COM execution on a disposable copy, no dialogs, no input save, final session IDLE. See [case report](../../../validation/2026-11-01/modeling-final/configurations/runner.json) and [scope/limits](../../../validation/2026-11-01/README.md). Whole-project compilation, other releases and unexecuted negative cases remain NOT_TESTED. Source headers predate this recorded run.
