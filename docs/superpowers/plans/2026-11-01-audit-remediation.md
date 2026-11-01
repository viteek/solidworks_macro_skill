# Audit remediation implementation plan

**Goal:** implement the four audit stages without inventing evidence from unavailable SOLIDWORKS installations.

**Spec:** user's technical audit dated 2026-11-01 and explicit selection of stages I–IV.

**Architecture:** preserve existing paths. Keep executable recipes separate from API guidance;
use offline regression checks plus explicit integration evidence. A reviewed test manifest is a
preflight contract, not a filesystem sandbox for arbitrary VBA.

**Tech stack:** Python 3.10+, unittest, VBA, official SOLIDWORKS API documentation, optional pywin32/geometry stack.

## Global constraints

- Never modify user CAD originals. Live execution uses only the reviewed runner on copied vendor/generated fixtures,
  after confirming an empty SOLIDWORKS session; do not close existing user documents.
- Preserve existing templates and entry points; document the deliberate mandatory runner-profile migration.
- No copied external macro code; source links carry provenance and reuse conditions.
- Static checks, project compilation, execution and result assertions are separate evidence levels.
- No version is certified without a report naming release, SP, UI language and fixtures.
- Work is local; do not commit, push or publish without a user request.

## Review focus

1. Timeout or COM exception must leave durable unresolved state and prevent retries.
2. Profiles must reject changed code, external dependencies and escaped/symlinked paths.
3. Dialog allowlists must bind title, text and answer and allow only one automatic response.
4. Recipes must report missing/incorrect selection, invalid units and negative tests without a false PASS.
5. Benchmark scores must reject missing evidence and mismatched fixtures; A/B comparisons use the same tasks.

## Task 1 — runner and analyzers (primary agent)

- [x] Reproduce Boolean false positives and missing checks with unittest.
- [x] Add profile validation and durable session state; test actual files and mocked OS/COM boundaries.
- [x] Restrict dialogs and adapt validated export to private copies.
- [x] Complete native-project support, environment diagnostics and integration instructions.
- [x] Run the full offline suite, all example checks and final independent review.

## Task 2 — modeling and selection recipes (independent files)

- [x] Add complete executable recipes for language-independent plane selection, sketch/extrude,
  FeatureData modification, configuration variants and persistent-reference round trip.
- [x] Add focused references and per-recipe input/expected/negative/verification descriptions.
- [x] Verify new APIs against official Help; run the bundled linter on each artifact.

## Task 3 — assembly, drawing and fabrication recipes (independent files)

- [x] Add assembly mates/MateReference guidance with exact signatures and version evidence.
- [x] Supply complete recipes for assembly inspection, mate creation/reference readback,
  drawing view creation and sheetmetal/cut-list diagnostics.
- [x] Document annotations and PDM boundaries; each artifact has expected and negative cases.
- [x] Verify APIs against official Help; run bundled lint checks.

## Task 4 — result verification, benchmark and sources (independent files)

- [x] Implement verify_results.py with machine-readable compile/API/runtime/result/safety statuses.
- [x] Add 24 exact task specifications in six families, generated fixture instructions and expected assertions.
- [x] Add A/B evaluation that validates evidence identity and reports completion/attempt/API errors/safety/tokens.
- [x] Add source index, version matrix and truthful verification evidence; test invalid/missing/duplicate cases.

## Task 5 — integration (primary agent)

- [x] Route SKILL.md to the new knowledge modules and executable recipes; synchronize COM guidance.
- [x] Add README installation/runner/quality instructions, source policy and changelog.
- [x] Add offline CI and opt-in local integration workflow; do not enable live SW execution on PRs.
- [x] Record executed checks, integration gaps and residual limitations in the final report.

## Execution notes

The independent recipe/verification tasks own separate files and may run concurrently under
the dispatching-parallel-agents skill. The primary agent owns integration and the safety-critical
runner. The source audit's live acceptance criteria require a licensed SOLIDWORKS workstation;
absence of that environment is reported as NOT_LIVE_TESTED, never treated as a pass.
Discovery update: a running 3DEXPERIENCE R2026x installation was found outside the classic SOLIDWORKS Corp directory.
Use its actual type libraries; preflight live state read-only before deciding whether integration can proceed.

## Recorded outcome and remaining acceptance gates

Implementation tasks above are complete. The retained [validation record](../../../validation/2026-11-01/README.md) contains 107 passing Python tests, 58 VBA assets/fragments checked against installed type libraries with zero errors, nine modeling/selection cases, six assembly/drawing/inspection recipes and positive sheet-metal creation/persisted readback. Final independent review findings were fixed and regression tested. All final live tests use CLI/COM only, per the user instruction.

The audit's complete maturity acceptance is still open:

- [ ] Accept matching whole-project compilation evidence for the persisted final VBA project, including classes/forms if used. The mismatched compile experiment was not accepted.
- [ ] Execute the complete paired 24-task Codex A/B experiment with real usage measurements; the implemented scorer has no live improvement result yet.
- [ ] Obtain additional target-release/service-pack evidence; only the installed 2026 SP4 environment was exercised.
- [ ] Select repository LICENSE; the author's answer is pending, so no rights grant was invented.

These gaps remain explicit NOT_TESTED/pending states; local smoke evidence does not silently satisfy them.
