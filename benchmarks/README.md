# Reproducible A/B benchmark contract

There are exactly 24 tasks in `tasks/`: six families with four tasks each. They are task specifications, not passing results. No binary SOLIDWORKS fixtures or live benchmark measurements are distributed here. Every task currently records `live_evidence_status: NOT_TESTED`.

| Tasks | Family | Work covered |
| --- | --- | --- |
| T01–T04 | docs_properties | Read, resolve and write properties; document traversal |
| T05–T08 | selection_modeling | Selection, extrusion, cut and dimension edit |
| T09–T12 | assemblies | Instances, transform, mate and lightweight state |
| T13–T16 | drawings | Sheets/views, notes, PDF and sheet rename |
| T17–T20 | sheetmetal_export | Flat DXF, bend lines, multibody DXF and STEP |
| T21–T24 | errors_resilience | No document, wrong type, invalid path and no selection |

## Fixture preparation

1. Read the exact `fixture_instructions`, request and assertions in each task JSON. Generate CAD fixtures in the target installed SOLIDWORKS version using recorded local templates. The recipes describe dimensions, configurations, names and starting state; `build_specs.py` only materializes these specifications and does not create CAD files.
2. Set explicit feature/entity names and verify fixture preconditions manually. Record template paths/hashes, units, document options, export options, UI language, version, service pack, API revision and architecture. Do not silently replace an unavailable template or a lightweight precondition: record `NOT_TESTED`.
3. Create a fixture manifest containing fixture ID/revision, every CAD/reference/template file with SHA256, document inventory, selected entities/marks, active configuration/sheet, dirty state, relevant preferences and precondition measurements. For T21 record the empty session inventory. Hash the exact manifest bytes as `fixture_sha256`. Independently verify referenced file hashes before and after every run; the verifier hashes the manifest artifact itself and does not read arbitrary files named inside it.
4. Reuse byte-identical fixture manifests and initial CAD content for A and B. Copy to disposable directories; reset state before each variant. Each variant gets an initially empty output directory. Neither variant may save original documents. Work requiring a final saved model must save only an explicitly named disposable result copy; these tasks require in-memory edits plus independent post-run inspection and unchanged source files.
5. Keep all inputs and evidence under that run's evidence root. Preserve the complete VBA project as `macro_artifact` (`.swp`, or an archive containing all `.bas`, `.cls`, `.frm`, `.frx` and dependencies), then hash it. A single exported module is insufficient evidence for compiling a multi-module project.

All numeric CAD observations use SI meters/m³ unless the JSON pointer explicitly ends in `_mm`. Numeric assertions use the specified absolute tolerance and optional `rel_tol` with `math.isclose`; there is no implicit tolerance. Object assertions use explicit type/name/identity/count fields, not non-null COM dispatch wrappers. Lists use multiset comparison, preserving duplicate counts. Feature and mate assertions must come from traversal after rebuild. Property assertions come from the specified document/configuration scope. PDF page counts/sheet IDs, DXF dimensions/contours/bend lines and STEP reimport geometry must be independently inspected; recognizing a header alone does not satisfy those tasks. For T09, component names are normalized to slash-separated hierarchy paths and instance counts include the subassembly and suppressed node. For T10, rotation rows are normalized from the transform's rotation matrix, separate from scale and translation.

## Evidence levels and statuses

`verify_results.py` is an offline evaluator of independently collected observations. It contains no live COM collector. Macro-produced success messages alone do not establish result correctness, compilation or safety. `PASS`/`FAIL` claims require nonempty real evidence artifacts. Missing evidence becomes `NOT_TESTED`. A failing observed assertion is `FAIL`; missing observations/artifacts remain `NOT_TESTED`. Overall `PASS` requires all six phases to pass.

| Report key | Required evidence for a meaningful PASS |
| --- | --- |
| `static` | Complete source inspection report for syntax/style hazards and forbidden operations; list inspected modules |
| `api` | Actual installed type-library/member/signature/enum checks for used calls, library file hashes and version; unsupported dynamic cases documented |
| `project_compile` | Actual VBE **Debug → Compile** of the complete project, dependency/reference list, compile log or screenshot and exact macro hash; source parsing and loading an `.swp` do not prove compilation |
| `runtime` | Actual invocation log with entry/module, start/end, return/error codes, run ID, environment and macro hash; an expected controlled rejection counts as runtime success only when observed |
| `result` | Derived by this evaluator from every task assertion and `runtime: PASS`; never supplied as an input phase |
| `safety` | Before/after source/reference hashes, save/event audit, preferences/configuration/selection restoration where required, created-file inventory and zero unintended saves |

The evidence consumer must review provenance and collection correctness. Hashes detect changes after collection; they cannot authenticate an operator, establish that a screenshot proves compilation, or prove that a supplied CAD observation is true. Record missing collection as `NOT_TESTED`; do not convert the offline tests below into live certification.

## Observation JSON and verification

The required top-level object is:

```json
{
  "schema_version": 1,
  "task_id": "T01",
  "fixture": {"id": "properties-part", "revision": "1"},
  "environment": {
    "solidworks_version": "2024", "service_pack": "5",
    "api_revision": "32.5", "architecture": "x64", "ui_language": "en-US"
  },
  "provenance": {
    "run_id": "A-T01-unique-session",
    "collected_at": "2026-11-01T10:00:00Z",
    "macro_artifact": "macro.swp", "macro_sha256": "REPLACE_WITH_ACTUAL_64_HEX_DIGEST",
    "fixture_artifact": "fixture-manifest.json", "fixture_sha256": "REPLACE_WITH_ACTUAL_64_HEX_DIGEST"
  },
  "values": {},
  "evidence": {
    "static": {"status": "NOT_TESTED", "artifacts": []},
    "api": {"status": "NOT_TESTED", "artifacts": []},
    "project_compile": {"status": "NOT_TESTED", "artifacts": []},
    "runtime": {"status": "NOT_TESTED", "artifacts": []},
    "safety": {"status": "NOT_TESTED", "artifacts": []}
  }
}
```

This is a template, not collected evidence. Replace environment, timestamp and hashes with actual values. Populate `values` from independent post-run inspection: e.g. T01 needs `document_type`, `properties.document.PartNo`, `properties.document.Description`, `dirty`, and `safety.unintended_saves` / `safety.source_hashes_unchanged`. Each assertion's `path` is an RFC 6901-style JSON pointer relative to `values`; `~1` denotes a slash in a property name. For non-NOT_TESTED phases, list evidence paths relative to `--evidence-root`. Evidence must identify the run, inspected macro and fixture in its own content. Parent-directory traversal, absolute paths, duplicate JSON keys and nonfinite JSON numbers are rejected.

```powershell
python scripts/verify_results.py --spec benchmarks/tasks/T01.json --observations evidence/A/T01/observations.json --evidence-root evidence/A/T01 --output evidence/A/T01/report.json
```

Exit codes: `0 PASS`, `1 FAIL`, `2 NOT_TESTED`, `3 invalid input`. The machine-readable report preserves task/fixture/environment/provenance, specification hash, independent phase statuses, per-assertion results and evidence hashes. A report based on all available assertions but missing runtime evidence cannot report result PASS.

## Paired experiment and measurements

A is the recorded baseline workflow without this skill; B uses the audited skill. Freeze model/version/settings, exact prompts, task order, retry policy and token accounting before running. Use the same task requests, fixture hashes and SOLIDWORKS environment. Record all attempts, including failed generations and retries. Alternate or randomize A/B execution order and record it with timestamps; retain raw generation/API/save logs. Do not remove hard tasks from the denominator or substitute easier fixtures.

Each manifest has `schema_version: 1` and `runs`, exactly one record per T01–T24 (one record aggregates all attempts for that task):

```json
{
  "run_id": "A-T01-unique-session", "variant": "A", "task_id": "T01",
  "fixture": {"id": "properties-part", "revision": "1"},
  "environment": {"solidworks_version": "2024", "service_pack": "5", "api_revision": "32.5", "architecture": "x64", "ui_language": "en-US"},
  "report": "T01/report.json",
  "measurements": {"attempts": 1, "api_errors": 0, "unintended_saves": 0, "tokens": 1234},
  "measurement_evidence": ["T01/measurements.json"]
}
```

The sample numbers describe the format and are not benchmark results. The referenced measurement JSON must contain exactly the same `run_id` and `measurements` object. The evidence root for imported reports is the manifest's directory: preserve artifact paths in reports relative to this common root when generating them. An easy layout is one manifest directory per variant, with `T01/observations.json` referencing `T01/macro.swp`, `T01/fixture-manifest.json`, `T01/runtime.json` etc.; invoke the verifier with the variant directory as `--evidence-root`.

Counter definitions: `attempts` counts every submitted candidate, minimum one; `api_errors` counts all non-success API HRESULTs/return codes and unhandled runtime API errors across attempts (expected rejection before an API call is not an API error); `unintended_saves` counts actual writes to any non-authorized document/output; `tokens` is total input plus output tokens across all generation/retry/tool model calls for that task, using provider usage records, with cached input counted once as input. If any counter cannot be measured, do not invent zero: the manifest is incomplete and comparison must reject it. Include the provider usage records and event logs in the evidence archive; the CLI validates the summary binding, not the truth of upstream instrumentation.

```powershell
python scripts/benchmark.py --a evidence/A/runs.json --b evidence/B/runs.json --output evidence/comparison.json
python -m unittest discover -s tests -p 'test_results.py'
python -m unittest discover -s tests -p 'test_benchmark.py'
```

The comparison checks exact task membership, fixture ID/revision/hash, specification hash, unique run IDs, metadata equality across all tasks and variants, complete counters, phase consistency, report assertion IDs and unchanged source/evidence hashes. Completion requires six-phase PASS and zero unintended saves. It reports per-task status, completion counts/rates, attempts, API errors, unintended saves, tokens and B-minus-A deltas; NOT_TESTED is always in the denominator and never counts as completed. A valid comparison can show 0/24 completions; its process exit code indicates input validity, not successful CAD execution. No speedup or quality improvement is currently claimed.
