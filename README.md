# SOLIDWORKS macro skill

Codex guidance, original VBA recipes and validation tools for SOLIDWORKS automation.
This working version (`3.0.0-dev`) implements the audit's correctness, CAD recipe,
verification and benchmark infrastructure. It does not certify all recipes across SOLIDWORKS versions.

## Install in Codex

Clone the repository into your personal skills directory:

```powershell
git clone https://github.com/viteek/solidworks_macro_skill.git "$env:USERPROFILE\.codex\skills\solidworks-macros"
```

Start a new Codex chat and invoke `$solidworks-macros`, or ask for SOLIDWORKS API automation.
Keep the entire folder: `SKILL.md` routes to `references/`, `recipes/`, `assets/` and `scripts/`.
Updating an existing installation is a normal Git update after preserving any local edits.

## What's included

- Existing safe export, batch, assembly traversal and source-metrics templates keep their paths.
- Sixteen executable recipes cover selection/identity, sketches, extrusion/revolve, holes/patterns,
  FeatureData/configurations, components/transforms, mates/MateReference, drawings, sheet-metal creation and fabrication diagnostics.
- Reviewed test profiles, durable execution state, exact dialog allowlists and hash-bound observers.
- VBA lint, installed-library API checks, native `.swp` execution, DXF/STEP geometry validation,
  result assertions and a 24-task A/B benchmark contract.

Read each recipe's README before adapting it. It lists preconditions, expected results, negative
cases, source links and evidence limits. [Recorded validation](validation/2026-11-01/README.md)
describes the checks actually performed; [the compatibility matrix](references/version-compatibility.md)
keeps untested versions and whole-project compilation separate.

## Run the quality checks

Python 3.10+ is required. The offline suite runs on Windows/Linux without SOLIDWORKS:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-test.txt
.\.venv\Scripts\python.exe -B -m unittest discover -s tests -v
.\.venv\Scripts\python.exe -B scripts/check_examples.py
```

For real installed-library checks, install `pywin32` on Windows and run:

```powershell
python scripts/verify_environment.py --require-api
python scripts/check_examples.py --api
```

`--sw-dir` selects an installation explicitly; `SOLIDWORKS_API_DIR` overrides discovery.
Classic SOLIDWORKS Corp and installed Dassault Systèmes 3DEXPERIENCE directories are discovered.
The offline fixture tests verify the checker itself; only actual `.tlb` checks establish installed API evidence.

## Execute on test copies

[Runner setup and recovery](references/runner-safety.md) is required reading for live tests.
Prepare copies with their dependencies, an initially empty SOLIDWORKS session, and a reviewed profile:

```powershell
python scripts/run_in_solidworks.py macro.bas --open C:\SWTest\copy.sldprt --profile C:\SWTest\profile.json --create-profile --expect-file C:\SWTest\result.step --json C:\SWTest\run.json
```

That command only writes an **unreviewed** profile. Review the prepared macro, entry point,
computed paths, input hashes, dependency closure, outputs and allowed dialogs; confirm copies
and mark the profile reviewed. Repeat with identical arguments and omit `--create-profile` to execute.
Model modifications must be declared in the profile and confirmed with `--confirm-model-changes`.
An external reference or already open document blocks execution. Do not close user work for a test.
Profiles validate known inputs and outputs; they do not sandbox arbitrary VBA filesystem access.

After a timeout or uncertain COM error, the durable session blocks another run.
Inspect SOLIDWORKS, finish the pending operation, then follow the explicit recovery command in the guide.
For complete projects and classes/forms, see [VBA project testing](references/vba-projects-and-testing.md).

## Verify outcomes and measure improvement

`--observe` runs a separately reviewed observer before documents close. Use actual readback and
fresh output assertions. [Geometry validation](references/geometry-validation.md) checks exported DXF/STEP;
[the benchmark protocol](benchmarks/README.md) explains independent observations, evidence artifacts,
six phase statuses and paired A/B scoring. No baseline scores are invented.

External code is not vendored. [SOURCES.md](SOURCES.md) records reuse conditions and the source index.
Repository licensing remains an author decision until a LICENSE is selected.
