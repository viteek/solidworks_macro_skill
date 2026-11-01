# Local audit validation — 2026-11-01

This is a bounded validation record for `3.0.0-dev`, not certification of the full benchmark
or every documented API branch. Live runs used PowerShell/Python and COM through reviewed,
hash-bound runner profiles. No screen automation was used for the final smoke tests.

Environment: Windows x64, SOLIDWORKS Design Professional for Makers 2026 SP4.0,
API revision `34.4.0`, Polish UI; installed in the Dassault Systemes 3DEXPERIENCE R2026x directory.
See `environment.json` for the actual Python/platform/type-library metadata and `summary.json`
for machine-readable check results. Those files describe the workstation, not a support promise.

## Retained checks

| Scope | Evidence | Meaning |
| --- | --- | --- |
| Python regression suite | `quality/unittest.log`, `quality/checks.json` | Offline and installed-library tests; no implicit live execution |
| Every bundled VBA asset and reference fragment | `quality/examples-api.log` | L1 rules and real installed-type-library signatures; not project compilation |
| Nine modeling/selection cases | `modeling-final/smoke-report.json` and per-case logs/profiles/observations | Sketch/extrusion, FeatureData modification, configurations, linear/circular patterns, revolve, 3D sketch, plane selection, persistent-reference round trip |
| Six assembly/drawing/inspection recipes | `assembly-smoke/live-evidence.json` and per-case runner/log files | Three-slot MateReference readback, coincident assembly mate, inserted component transform, three drawing views, annotation-free drawing inspection, ordinary-solid fabrication inspection |
| Positive sheet-metal creation | `sheetmetal-positive/evidence.json`, `sheetmetal-positive/metrics.json` | 100×50×2 mm base flange; volume 10,000 mm³, thickness 2 mm, sheet-metal body, FlatPattern and one cut-list body |
| Persisted sheet-metal inspection | `sheetmetal-reopen/runner.json` and logs | Fresh copy of saved generated part reopened and inspected with independent metrics |
| COM truth values and failed log output | `boolean-final-summary.json` and `boolean-final-*` | Numeric 0/1/−1 plus actual API raw true value 1; missing output parent fails expectations, returns promptly and leaves IDLE |
| STEP export with geometry readback | `step-validation.json` | Private-copy export, OpenCascade volume/area/centroid/bounds comparison and checked publication |
| Native `.swp` invocation | `native-noop-runner.json`, `native-noop-observed.json` | Actual binary project entry invoked, unchanged block geometry independently read back; no-op only |
| External dependency refusal | `external-dependency-rejected.json` | Vendor part referencing an external assembly was rejected before macro execution |
| Saved-input detection | `saved-input-rejected/assertions.json`, `saved-input-rejected/runner.json` | A probe deliberately saved a disposable input; the runner rejected the changed hash despite in-memory edit permission |

Live inputs were copied vendor files/templates or locally generated fixtures. Original vendor and
part-template hashes are in `source-safety.json`. Harness reports also check their immediate input
hashes. Successful runs recorded no dialogs, an IDLE durable session and no remaining documents.
Only new generated fixture files were saved; dirty test copies were discarded by the reviewed runner.
No production documents were used or closed.

## Failures and limits

- `modeling-smoke/` retains the first two pattern failures. VBA rejected `circle` as a variable
  name. The recipes were corrected, a reserved-word regression was added, and both corrected
  cases passed in `pattern-smoke-fixed/`; `modeling-final/` contains the current full rerun.
- `native-missing-module-rejected.json` retains an invalid module/entry experiment. A previous
  in-memory compile observation cannot certify the final binary: its persisted module inventory
  did not match the fifteen imported source modules. Whole-project compilation is therefore
  **NOT_TESTED / not accepted**, including classes/forms and current source-to-binary binding.
  `RunMacro2` is not a complete-project compile-only command.
- The full paired 24-task Codex A/B experiment has **NOT_TESTED** status. Specifications, fixture
  instructions, provenance checks and scoring exist; no improvement percentage or token savings
  is claimed. Local smoke cases are not silently substituted for those benchmark tasks.
- Other SOLIDWORKS releases/service packs are **NOT_TESTED**. The local results do not certify them.
- Drawing fixtures contain no GTol, weld symbols or BOM; their populated branches remain untested.
  Sheet-metal tests cover an unbent base flange and cut-list readback, not bends, DXF unfolding,
  multibody weldments or fabrication readiness. Most per-recipe negative cases remain manual.
- Runner safety has file/COM boundary regression tests. Live timeout/dialog fault injection is not
  claimed here; copied external-dependency rejection and deliberate disposable-input save detection
  are the retained live refusal examples.

Binary CAD fixtures, installed templates and the experimental native project are deliberately not
redistributed. Logs retain their original absolute paths and hashes; archived profiles are historical
records, not ready-to-run approval files. Regenerate profiles and review fresh local copies before reuse.
`artifact-manifest.json` binds the retained evidence bytes. The normal runner profiles are a preflight
contract; they are not an operating-system sandbox for arbitrary VBA.

Some source-module headers retain their original conservative `NOT_LIVE_TESTED` label so the
executed bytes and recorded hashes remain identical. The dated recipe READMEs and this record
provide the later, scoped status. `sheetmetal-final/` records a successful creation/inspection
rerun after the runner's input-hash policy was tightened.
