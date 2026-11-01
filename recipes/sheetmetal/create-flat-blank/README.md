# Create an unbent sheet-metal blank

`create_flat_blank.bas` is an original complete VBA recipe for a **100 x 50 x 2 mm**
base flange. It uses the current `CreateDefinition(swFmBaseFlange)`/`CreateFeature`
interface instead of the older InsertSheetMetalBaseFlange2 call. It does not bend,
flatten, export, save or close CAD documents. Import with SOLIDWORKS and SOLIDWORKS
Constant type-library references and run `main` on a disposable blank part.

## Inputs and behavior

The active part must have one configuration, exactly three reference planes, no solid or
sheet bodies, no user sketches and no active sketch edit. A saved private copy is accepted.
The recipe selects the first reference-plane **object** by invariant `RefPlane` type,
creates a 100 x 50 mm rectangle in its local sketch coordinates, and creates a base flange
with thickness 0.002 m, bend radius 0.002 m, custom K-factor 0.5, rectangular relief and
relief ratio 0.5. No bend is present, so the allowance/relief does not remove geometry.
The profile is intentionally underdefined. All API length values use meters, volume m³.

After rebuilding it asserts one body, `IBody2.IsSheetMetal <> False`, volume 1e-5 m³,
base-flange and global sheet-metal thickness 0.002 m, and at least one SheetMetal and
FlatPattern feature. It recursively inspects/deduplicates feature folders; each CutListFolder
that exists must contain exactly one body. A cut-list count is reported rather than promising
automatic cut-list creation on every workstation. The tested fixture produced one folder.
FlatPattern presence does not prove its suppression state or bent-part unfold correctness.

`RESULT_PATH=""` disables file output. Set it to a fresh file in the reviewed private run
directory to receive `RESULT create_flat_blank OK...` or `FAIL...`. Existing logs are refused.
Failures `Resume Cleanup` before protected AddToDB restoration; cleanup failures suppress OK.
Partial geometry may remain in the disposable document on failure. No original is saved.

Negative cases still requiring controlled live tests: wrong document, nonblank part, active
sketch, multiple configurations, additional reference planes, nonpositive dimensions and
existing/unwritable log paths. The new helper's offline tests reject missing/multiple bodies,
incorrect units, nonfinite values, incorrect volume, invalid bounds and matching volume with
wrong dimensions. They are evidence-parser tests, not SOLIDWORKS solver tests.

## Observed local integration evidence

**LIVE_SMOKE_PASS**, SOLIDWORKS revision **34.4.0**, identified by the user as 2026 SP4,
2026-11-01. UI language was not independently collected; the template's configuration name
was `Domyslna` with Polish diacritic. This does not certify other releases, service packs,
UI languages or bent sheet metal. Whole native VBA project compilation remains NOT_TESTED;
the runner executed temporary `.swb` sources.

The reviewed runner opened a hashed copy of the blank fixture with explicit model-change
permission and no allowed dialogs. The macro returned OK with one sheet-metal body,
thickness 0.002 m, FlatPattern count 1 and CutListFolder count 1. Independent source metrics:

| Measurement | Observed |
|---|---:|
| Volume | 10000 mm³ |
| Area | 10600 mm² |
| Bounds | [0, 0, -2, 100, 50, 0] mm |
| Centroid | [50, 25, -1] mm |

The existing `inspect_sheetmetal_cutlist.bas` ran successfully on that positive model.
A reviewed observer saved only a **new** native fixture copy. A second read-only runner
session reopened a fresh copy of that saved model, ran the unchanged inspector again and
independently reproduced the same metrics. Both runs ended IDLE with no dialogs, runtime
errors or documents left open; source macro/template/helper and input-copy hashes were unchanged.

See [evidence.json](evidence.json) for exact source/prepared-code hashes and original report
paths. The fixture binary stays outside the repository in the private run directory.
Native fixture SHA256: `ba1d79cc35a12a714ab680543fc2884a48053f8762e6d1451ff8fc1f55ee94f7`.

Preparation and explicitly authorized rerun:

```powershell
python scripts/smoke_sheetmetal.py --blank <private-blank.sldprt> --test-root <NEW-private-run-directory>
python scripts/smoke_sheetmetal.py --blank <private-blank.sldprt> --test-root <ANOTHER-NEW-private-run-directory> --run
```

The helper copies its input, binds code/input/observer hashes in a runner profile, drives
all CAD changes through the reviewed runner, independently checks volume and sizes, and
stops on UNKNOWN without further COM queries, retries or cleanup. `--run` requires an
already-running SOLIDWORKS session appropriate for isolated tests. Reuse neither an old
run directory nor the user's production CAD file as the opened document.

## Official API sources

Consulted 2026-11-01; API contracts informed original code. No external macro code was copied.

- [IBaseFlangeFeatureData: CreateDefinition/CreateFeature workflow](https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IBaseFlangeFeatureData.html)
- [Initialize: exact nine parameters and override semantics](https://help.solidworks.com/2022/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IBaseFlangeFeatureData~Initialize.html)
- [ICustomBendAllowance.KFactor: setting this property chooses K-factor allowance](https://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.ICustomBendAllowance~KFactor.html)
- [IBody2.IsSheetMetal](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IBody2~IsSheetMetal.html)
- [IBody2.GetMassProperties: solid volume at index 3](https://help.solidworks.com/2020/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ibody2~getmassproperties.html)

L1 lint and local 2026 type-library checks pass. DXF export, bend lines, relief geometry,
manufacturing allowances, other materials and production cut-list freshness are not tested.
