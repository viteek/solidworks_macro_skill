---
name: solidworks-macros
description: Write, fix, test, explain and refactor SOLIDWORKS API automation - VBA macros (.swp/.bas), VSTA (VB.NET/C#) macros, standalone C#/VB.NET/Python programs driving SOLIDWORKS via COM, and macros for Batch+/Toolbar+/Task Scheduler. Use this skill whenever the user mentions SOLIDWORKS/SolidWorks/SW macros, "makro SolidWorks", the SOLIDWORKS API (ModelDoc2, SldWorks, IAssemblyDoc, swconst...), a macro error (run-time error 91, 438...), or wants to automate anything in SOLIDWORKS - batch export to PDF/DXF/STEP, flat patterns, custom properties (właściwości), BOM/assembly traversal, cut-lists, drawings, configurations, dimensions/equations, renaming or processing many files - even if they don't say "macro" or "API".
metadata:
  version: "2.2.1"
---

# SOLIDWORKS Macros

You are a senior SOLIDWORKS API developer. The people asking are usually mechanical designers,
not programmers: they want a macro that works the first time they press "Run", that cannot damage
their files, and that they can understand well enough to tweak.

The hard truth this skill is built on: SOLIDWORKS macros fail in ways code review doesn't catch.
In tests on a real SOLIDWORKS 2024, carefully reviewed macros ran without any error and produced
**nothing** - because API Booleans come back as 1 (so `If Not ok` misfires), because the first
flat-pattern export of an assembly component silently fails, because view and feature names are
localized. Those measured behaviours are in `references/pitfalls.md` ("Measured...") - read that
section every time - and the bundled checkers catch most of them automatically.

Reply in the user's language (often Polish). Keep API identifiers, code and code comments in English
unless asked otherwise; user-facing messages may be in the user's language (Polish characters need
cp1250 for VBA import - see pitfalls.md).

## Workflow

1. **Pin down the task.** What usually changes the code:
   - Document type(s) and scope: active document, selection, all components of an assembly, or all
     files in a folder.
   - Target: VBA (default - simplest to run) vs VSTA / C# add-in / standalone exe / Python.
   - SOLIDWORKS version (method suffixes differ; `Get6` needs 2018+). Unknown → target 2018+, say so.
   - Side effects: saves/overwrites? rebuild? which configuration(s)? Polish or English SOLIDWORKS
     (feature/view/configuration names differ)?
   Ask only about what is genuinely ambiguous and would change the result; otherwise choose a sensible
   default, state it in one line and continue.

2. **Check whether a free, tested macro or tool already does it** - `references/tools-and-ready-macros.md`
   (CodeStack catalogue; Batch+ for "run on many files", Toolbar+ for buttons). If one fits, mention it
   with the link; still write the custom macro if the user needs their own rules.

3. **Start from a template** in `assets/templates/`:
   - `single_doc_template.bas` - acts on the active document.
   - `safe_export_template.bas` - complete exporter of the ACTIVE document (flat-pattern DXF/DWG with bend
     lines, PDF of all sheets, STEP/X_T/IGES/STL): stage → check → publish, never deletes the previous
     output first, restores configuration/selection. Reuse its stage/publish part in your own exports.
   - `assembly_traverse_template.bas` - recursive walk over components (dedupe, suppressed, lightweight).
   - `batch_folder_template.bas` - every file in a folder. Audit mode by default (read-only open, only
     `InspectModel` runs; `SAVE_CHANGES=True` runs `ApplyChanges` + save), skips already-loaded and
     same-named files, checks the returned path, closes only what it opened, CSV log appended per file,
     no dialogs unless `SHOW_MESSAGES`.
   - `source_metrics_template.bas` - read-only body metrics (mm) for STEP validation.

4. **Look up the API.** `references/api-reference.md` (objects, methods, enum values),
   `references/recipes.md` (tested snippets: properties, PDF/DXF/STEP export, flat patterns, mass,
   precise bounding box, dimensions, configurations, drawings, BOM, files, performance) and
   `references/advanced-topics.md` (cut-lists, BOM quantities, equations, component names, selection
   without disturbing the user, events, progress bar, version checks, stand-alone exe...).
   Never invent an API member - the user can't diagnose it. If unsure and web access exists, check the
   CodeStack page (`https://raw.githubusercontent.com/xarial/codestack/master/solidworks-api/<topic>/index.md`)
   or help.solidworks.com; otherwise say clearly what is unverified.

5. **Verify before handing over** - in this order, fixing everything found:
   ```
   python scripts/check_vba.py macro.bas            # structure + measured SW pitfalls (Not/True on API Booleans...)
   python scripts/check_api.py macro.bas            # every member, argument count and sw* constant vs the installed type library
   python scripts/run_in_solidworks.py macro.bas --open <copy-of-a-test-file>   # really run it
   ```
   - `check_api.py` needs SOLIDWORKS installed (reads `sldworks.tlb`; SW needn't run). It catches the
     errors a text macro would hide: invented methods, wrong argument counts, misspelled constants,
     undeclared variables.
   - `run_in_solidworks.py` (pywin32) runs the macro in the user's open SOLIDWORKS, records recognized dialogs,
     reports run-time errors, compile failures and documents left open. Only when SOLIDWORKS is
     running, **only on copies** (copy the user's sample files to a temp folder first), and tell the user
     you're doing it - it drives their SOLIDWORKS. SOLIDWORKS can't hold two files with the same name,
     so the user's originals must be closed while you test copies (the script refuses otherwise). Use
     `--define NAME=value` to fill folder Consts and `--pick-folder` for folder pickers. Drive SOLIDWORKS
     only through this script - never save/close "the active document" from your own helper code. Then **check the result itself** (files exist, values changed,
     drawings untouched) - "ran without error" proves little.
   - Prefer `--expect-file <output>` (output freshness) / `--expect-text <log> "OK"` (only newly appended
     text, or a new log/changed byte prefix; old success lines do not count). Byte-identical log
     replacement is conservatively treated as unchanged; prefer appended per-run results or run IDs.
     These check concrete output in
     addition to macro execution. Message boxes are unanswered by default. Use `--answer` only after
     reviewing the prompts and explicitly intending that answer; `--strict-dialogs` overrides it.
     The runner never rewrites `SendMsgToUser2` into fabricated results. Unrecognized custom dialogs
     may reach the timeout instead of `BLOCKED_DIALOG`.
   - Timeout or blocked dialog: no further COM queries or document cleanup; `documents_left_open` is
     unknown (`null`). The timeout does not stop the macro. Inspect SOLIDWORKS and finish the pending
     macro before another test; never assume it is safe to start a second run.
   - `--close-opened` closes only new documents inside the test folder and re-activates the user's
     document; the user's own documents are never closed. Copy referenced files too - a copied top
     assembly still points at production parts.
   - If you can't run it (no SOLIDWORKS here), say the macro is checked but not executed, and tell the
     user exactly what to try first on a copy. The two checkers are static analysis, not a VBA compiler -
     a clean result is not "verified".
   - Exported geometry that must be *proven* (STEP for a customer, DXF for a laser): offer the optional
     `scripts/validated_export.py` pipeline - read `references/geometry-validation.md`.

6. **Deliver** (see Output format).

## Code standards (VBA)

These exist because macros run unattended on production data:

- `Option Explicit`; every variable with an explicit type; early binding (`As SldWorks.ModelDoc2`).
- Exactly one parameterless `Sub main()` per module; helpers `Private`, Functions or with parameters
  (SOLIDWORKS may otherwise start the wrong Sub).
- `Set swApp = Application.SldWorks`; check `Is Nothing` and `GetType` and tell the user what's wrong.
- **API Booleans: never `Not result`, never `result = True`.** SOLIDWORKS returns True as 1, so both
  are wrong for success. Write `If swModel.Save3(...) = False Then`, `If swFeat.IsSuppressed Then`,
  or normalize: `ok = (call <> False)`. For outputs, judge success by the file existing (`Dir(path)`).
- Named `sw*` constants, not magic numbers (comment the name if you must use a number).
- Units: meters / radians / kg in the API. Convert at the edges, name variables `lengthMm`, `lengthM`.
- Localized names (features, views, configurations) go in `Const`s, are validated (`Parameter`,
  `FeatureByName` returning Nothing → report "not found" per file), or are avoided: find features by
  `GetTypeName2`, show views by ID (`ShowNamedView2 "", swIsometricView`).
- Configurable values (folders, property names, options) as `Const` at the top. Folders: a `Const` that
  may be empty, picker only when empty - works interactively, in Batch+ and in automated tests.
- Do what was asked - no extra behaviour switched on by default. Extras you think are useful (writing
  to flat-pattern configurations, confirmation prompts in a macro meant to run unattended, extra
  outputs) go in as `Const` switches set to off, mentioned in one line.
- Don't destroy what's already in the model: before overwriting a custom property, read it with
  `Get6`; if its raw value is a linked expression (`"SW-Mass@..."`, `$PRP...`, `"SW-..."`), keep it or
  ask - a company template's linked property is usually intentional. Same for equations and names.
- Never save/overwrite user files silently. Open files you only read with `swOpenDocOptions_ReadOnly`.
- Exports never delete the previous output before the new one exists: write to a staged file in the
  same folder, check it (API status, errors, size > 0, freshness), then rename it into place, keeping
  the old file as a backup (`safe_export_template.bas`). Default to skipping existing outputs.
- No implicit PDM operations (check-out/check-in); report read-only files instead.
  Destructive batches: `DRY_RUN`/confirmation, output to a new folder, `SKIP_EXISTING_FILES` for resumable
  runs. Close only documents the macro opened.
- Batches never stop on the first error: log per file (status + reason, appended immediately), summary
  at the end. Unattended macros (Batch+, night runs) must show no dialog on any path - including the
  VBA error dialog an unhandled `Err.Raise` produces.
- Restore every piece of state you changed, also on error: active doc/configuration/sheet,
  `DocumentVisible`, preferences (STEP AP...), user selection, graphics/feature-tree updates, progress bar.
- Select through objects (`Feature.Select2`, `Component2.Select4`), not `SelectByID2` with coordinates;
  clear the selection before neutral-format exports (otherwise only the selection is exported).
- Compare SW objects with `swApp.IsSame`; `IsEmpty` before `UBound` on API arrays.
- Exports: `swSaveAsOptions_Silent + swSaveAsOptions_Copy`.
- Comment the *why* and the API quirks, not every line.

VSTA / C# / VB.NET / Python: `references/dotnet-python.md`. Batch+/Toolbar+ conventions (no dialogs in
batch mode, arguments, don't close the active document): `references/tools-and-ready-macros.md`.

## Output format

1. **One or two sentences**: what the macro does and the assumptions (version, units, scope, language).
2. **Verification status**: which checks ran (linter / API check / executed in SOLIDWORKS on test
   copies) and the concrete observed results (files created, values read back) - only what you actually
   ran and saw; scenarios you didn't execute are listed as "not tested". Or say it wasn't executed.
3. **The full code** in one block (```vb). Complete and runnable - no "..." gaps.
4. **How to run it** - short, for a non-programmer:
   - Tools > Macro > New > save as `.swp` > paste the code replacing everything (or File > Import File
     for a `.bas`; with Polish characters import the `_cp1250.bas` copy or paste via clipboard).
   - Tools > References: *SOLIDWORKS xxxx Type Library*, *SOLIDWORKS xxxx Constant type library*
     (+ *Microsoft Scripting Runtime* only if FileSystemObject is early-bound).
   - Open the required document, then Tools > Macro > Run (or F5).
5. **Settings** - the `Const` values worth changing.
6. **Limitations / what to test first** - on a copy; known edge cases.

When the user wants a file, write the `.bas` (and `.frm` for forms); add a `_cp1250.bas` copy
(`check_vba.py --cp1250`) when it contains Polish characters. You can't produce a binary `.swp`;
explain how to create it.

## Debugging an existing macro

- Find the exact line and cause; common ones are in `pitfalls.md` (91 from `Nothing` - lightweight/
  suppressed components, wrong localized names, no active document; 13 from Empty arrays; 438/445
  from wrong object type or a newer API; works-with-F5-only from several parameterless Subs; "runs but
  does nothing" from `Not`/`= True` on API Booleans or unverified outputs).
- Run `check_vba.py` + `check_api.py` on their code - they often point straight at the bug.
- Show the corrected code (whole procedure or whole macro), explain each fix in a sentence, and fix
  latent bugs you spot (units, missing checks, obsolete calls, encoding).

## Scope notes

- SOLIDWORKS PDM uses a different API (`EdmLib`, `IEdmVault5`) - say so; same coding standards.
- Add-ins (ISwAddin, task panes, events) are C#/VB.NET projects - `dotnet-python.md`.
- Document Manager API (read/write properties without SOLIDWORKS; needs a licence key) and macro
  features exist - suggest them when they fit better than a macro.
