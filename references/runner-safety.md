# Running reviewed macros on disposable copies

`scripts/run_in_solidworks.py` executes arbitrary VBA with the desktop user's privileges. Its profile checks and dialog policy are safeguards, not a filesystem, process, network or COM sandbox. A profile cannot stop a macro from opening undeclared files, changing preferences, saving production documents or launching another program. Review the complete macro/project and observer first; use isolated test copies and a dedicated SOLIDWORKS session.

## Required preflight and profile

Execution requires a macro path, `--profile` and `--open`. Use absolute paths. Start exactly one SOLIDWORKS process with **no documents open**; the runner refuses an already populated session. Every dependency resolved by both last-saved references and SOLIDWORKS search rules must be an explicitly declared input copy inside `test_root` with a matching SHA256. Opening a copy must not load undeclared documents. A copied assembly/drawing that still references production components is not an isolated fixture.

Create a draft without connecting to SOLIDWORKS:

```powershell
python scripts/run_in_solidworks.py C:\SWTest\macro.bas --open C:\SWTest\part.sldprt --profile C:\SWTest\review.json --create-profile --json C:\SWTest\runner.json --expect-file C:\SWTest\out\result.json
```

The draft is exclusively created; an existing profile is not overwritten. It starts with `reviewed: false` and `copies_confirmed: false`. Review before changing those flags. Add all referenced CAD copies to `inputs` with their actual hashes; list every expected output and report in `outputs`. All input/output paths must be absolute and inside the existing `test_root`; outputs cannot alias input CAD files. `--json` must be a distinct declared `.json` output, never the macro, observer, profile or opened input. The profile also binds the prepared macro hash and exact `entry_point` module/procedure.

For intended model edits, create the draft with `--allow-model-changes`, review `allow_model_changes: true`, then execute with `--confirm-model-changes`. This permits in-memory edits to copies; persisted results require distinct declared output paths. It does not grant a sandbox. Without model changes the opened copy uses silent/read-only options. After every resolved execution, including edit-enabled runs, the runner rechecks input hashes and fails if an input was saved. Independently audit hashes as well.

```powershell
python scripts/run_in_solidworks.py C:\SWTest\macro.bas --open C:\SWTest\part.sldprt --profile C:\SWTest\review.json --json C:\SWTest\runner.json --expect-file C:\SWTest\out\result.json --close-opened
```

`--define NAME=VBA_LITERAL` changes a source constant; `--pick-folder` supplies explicit source folder substitutions. These prepared bytes are hash-bound: profile creation and execution must use identical overrides. Any source, override, observer or entry change requires new review. Native `.swp` projects require explicit `--module` and `--entry`; the whole binary is hashed and source rewriting (`--define`, `--pick-folder`) is unavailable. A successful `.swp` invocation does not establish that every module compiled; see [project testing](vba-projects-and-testing.md).

## Observers and output checks

`--observe C:\SWTest\observer.bas` runs a reviewed source observer's `Sub main()` after successful macro invocation, before document cleanup. `--observe-define NAME=VBA_LITERAL` only changes observer constants and requires `--observe`. Both the observer and its prepared overrides are bound by `observer_sha256` in the profile. Binary `.swp` observers are unsupported. List all observer outputs as declared expectations when creating the profile; observer code has the same privileges as the tested macro.

The observer must measure post-run objects independently, not repeat a macro's “success” string. It can capture feature/body counts, dimensions, properties, mates, references and document state before `--close-opened` discards unsaved copy edits. Observer execution success only establishes invocation; its values still need [result assertions](../benchmarks/README.md). The runner's JSON is not the same schema as the six-phase result-verifier observation JSON; retain the raw report and deliberately adapt independently measured values and evidence.

`--expect-file` checks existence and modification time; it does not validate file structure or geometry. `--expect-text FILE TEXT` searches newly appended text, or replacement content when its previous prefix changed; identical pre-existing text is not fresh evidence. Prefer unique run IDs and independently parsed exports. Run the format validators and task assertions separately.

## Dialogs, cleanup and uncertain execution

Dialogs are left unanswered by default. `--answer` permits only a matching requested button for an exact profile allowlist entry (`title`, `text`, `answer`), at most once per run; it has no fallback answer. `--strict-dialogs` leaves every dialog unanswered even when `--answer` is supplied. VBA error dialogs are classified and fail the run; they are not dismissed automatically. Opaque/custom dialogs may reach the timeout. `SendMsgToUser2` calls are not rewritten.

`--close-opened` closes only documents opened during a resolved run inside the test root, discarding unsaved test-copy edits. It does not clean up an uncertain run. Runtime timeout, a blocked dialog, or unresolved COM exception can leave VBA executing. The runner avoids further live document queries/cleanup in that state and records durable `UNKNOWN` (a killed controller can leave `RUNNING`). Stopping Python or cancelling a workflow does **not** cancel VBA. Output freshness checks and JSON writing use filesystem state and can still occur; do not interpret them as proof VBA ended.

The controller lock and session marker live under `%LOCALAPPDATA%\solidworks-macro-runner` (temporary-directory fallback). New runs fail closed while a previous session is unresolved. Inspect the desktop and finish/cancel the actual VBA operation manually, then recover:

```powershell
python scripts/run_in_solidworks.py --recover-session --confirm-macro-stopped
```

Recovery performs no SOLIDWORKS COM calls; it records the operator's confirmation. If the recorded process is provably gone, explicit stopped confirmation is unnecessary; PID reuse/unknown process state is handled conservatively. Do not delete the marker or kill processes merely to bypass a block. A runner `PASS` reports its invocation/dialog/output checks; `compilation_status` remains `NOT_TESTED`, and result/safety acceptance requires independent evidence.
