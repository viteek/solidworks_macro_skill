# Changelog

## Unreleased

- Runner: new `--expect-result FILE PREFIX` requires the last newly written log line to start with an
  exact status (followed by `;`, whitespace or end of line). `--expect-text` now fails whenever the
  final new line is a FAIL/CANCELLED status, so `FAIL ...; prior=... OK` never passes with `OK`.
  Smoke harnesses and SKILL.md use `--expect-result`.
- Recipe cleanup-failure lines no longer repeat `RESULT <recipe> OK` or `<recipe> OK` in `prior=`.
- DXF and STEP validators re-hash the input after analysis and refuse a file changed meanwhile.
- JSON Pointer array indices follow RFC 6901 (no leading zeros, ASCII digits only, valid `~` escapes).
- `scripts/check_skill.py` validates SKILL.md frontmatter, version and routed paths; it runs in CI.
- Retained live evidence in `validation/2026-11-01` predates these changes.

## 3.0.0-dev - 2026-11-01

- Audit stage I: mandatory reviewed/hash-bound test profiles, copied input/dependency closure checks,
  explicit mutation confirmation, empty-session preflight, durable RUNNING/UNKNOWN states and
  OS-only explicit recovery. Timeouts and uncertain COM statuses do not permit retries or late observers.
- Dialog automation requires an exact title/text/answer allowlist and permits one response per run.
  Reports cannot overwrite inputs, macros or profiles during failed preflight. Native `.swp` bytes
  and reviewed entry points are preserved; optional hash-bound observers run before cleanup.
- Boolean analysis now tracks scoped COM provenance, normalization, aliases and numeric results;
  fixed standard FileExists false positives, unsafe RHS negation, Circle reserved identifiers,
  valid IBodyFolder.GetBodies, wrong API owners and negative/Or-prefixed bare-call arguments.
  PDF sheet enumeration now uses DrawingDoc. RunMacro error enums match the installed type library.
- Audit stage II: sixteen original complete CAD recipes, selection/model-state/Boolean diagnostics,
  model/configuration/FeatureData/assembly/MateReference/drawing/fabrication references and per-recipe
  expected/negative/verification instructions. Existing template paths remain available.
- Audit stages III-IV: portable offline tests and CI, manual self-hosted integration workflow,
  environment discovery (including 3DEXPERIENCE), six-phase evidence evaluation, 24 paired A/B task
  specifications, strict provenance/measurement validation, source policy/index and version matrix.
- Local verification records are in validation/2026-11-01. Scope is explicit: modeling smoke,
  source metrics, validated STEP, native invocation and COM truth-value probes do not certify
  every release, complete VBA projects or the full 24-task A/B experiment.
- Final review: edit permission allows in-memory changes only; saved input changes now fail hash
  checks for every resolved run. Evidence CLIs require fresh output files, and assembly smoke
  requires both a successful process exit and confirmed PASS/IDLE status.

## 2.2.1 - 2026-10-30

- Runner: message boxes are unanswered by default; explicit answers require a matching button.
  Removed fabricated `SendMsgToUser2` results and regex rewriting of message calls/string literals.
- Timeout/blocked prompt reports avoid further COM calls and document cleanup; the macro may still
  be running, so document state is reported as unknown. Blocked prompts return without waiting for
  the timeout. The macro must finish before another test starts.
- `--expect-text` examines newly appended text or a new/changed log, including UTF-16 logs;
  byte-identical replacement is conservatively treated as unchanged.
- DXF: preserve polyline traversal direction for negative bulges, including first/closing arcs.
- API checker: validate zero-argument member access against required parameters; recognize bare
  calls after inline `Then`/`Else` without flagging valid template calls.
- Flat-pattern recipe now uses the safe exporter instead of duplicating overwrite/backup logic.
- Added offline regression tests: `python -B -m unittest discover -s tests -v` (Windows, installed
  SOLIDWORKS type libraries, and the geometry dependencies; does not drive a SOLIDWORKS session).

## 2.2.0 - 2026-10-30 (merge of the local 2.0 line with Codex 2.1.0, after audit)

From 2.1.0 (reviewed and fixed):
- `safe_export_template.bas`: active-document exporter with stage -> check -> publish and backup.
  Fixed: bend lines use mask 1+4+8 (4 alone gives none on SW 2024 SP5 - measured), cleanup via
  `Resume` (On Error Resume Next is ignored inside an active error handler).
- `batch_folder_template.bas`: audit mode, read-only open, same-name / returned-path guards, CSV log per
  file. Fixed: `Resume` in three handlers, 8.3 short-path comparison, folder picker when interactive.
- `source_metrics_template.bas` (visible bodies only - hidden ones aren't exported to STEP).
- Geometry validation: `validate_step.py` (OpenCascade), `validate_dxf.py`, `validated_export.py`,
  `validation_reports.py`, `profiles/`. Fixed: DXF roles by LINETYPE then layer (SW puts bend lines on
  layer 0 as CENTERX2), ignored POINT entities, realistic tolerances, outside-material = warning,
  STEP validator no longer needs the DXF stack, sibling imports under `python -I`, pipeline uses this
  package's runner and removes its job folder after publication.
- Rules: no deleting outputs before success, read before replacing properties (linked values kept),
  no implicit PDM operations, static checks are not a compiler, corrected ExportToDWG2 parameter names,
  no CloseAllDocuments as recovery, cp1250 conversion refuses unrepresentable characters.

Kept from 2.0 (removed or contradicted in 2.1): trigger description, output format, debugging section,
measured SW 2024 pitfalls list, export recipes (DXF from assembly components, PDF per sheet, STEP AP...),
runner dialog answering / `--pick-folder` / compile-error hint.

Runner: `--close-opened` closes only new documents inside the test root and re-activates the user's
document; new `--test-root`, `--strict-dialogs`, `--expect-file`, `--expect-text`; refuses when a
same-named file from another folder is loaded.

Other: Polish feature name corrected (`Dodanie-wyciągnięcie1`); `SetProp` recipe updates values but
keeps linked expressions; flat-pattern recipe stages instead of `Kill`.

## 2.0 (local) - 2026-10-30
Measured SW 2024 pitfalls, type-library API checker, live runner, templates; evaluated on 6 tasks.
