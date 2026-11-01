# Pitfalls & troubleshooting

Check every macro against this list before handing it over.

## Measured on a real SOLIDWORKS (2024 SP5, Polish UI) - not in the docs

These were found by running macros in SOLIDWORKS, after code that looked correct failed. Each
was reproduced on that version; on other versions reproduce before relying on a workaround - the
safe coding forms below (`= False`, verify files, activate before export, mask 1+4+8 with a check of
the resulting DXF) are harmless everywhere. They are
the most common reasons a "reviewed" macro silently does nothing.

1. **API Booleans arrive as 1, not -1.** `SaveAs`, `Save3`, `ExportToDWG2`, `GetSaveFlag`,
   `IsSuppressed`, `SelectByID2`, `Select4`, `ForceRebuild3`, `EditRebuild3`, `ActivateSheet`,
   `ExportPdfData.SetSheets`, `GetUserPreferenceToggle`... may return nonstandard nonzero truth values.
   Numeric status methods such as `Component2.SetSuppression2` must be checked against their enum,
   not treated as COM Booleans.
   In VBA `Not 1` is -2, so **`Not result` is True for success too**, and `result = True` is False.
   `CBool()` does not normalize it. Exceptions seen: `Body2.IsSheetMetal`, `Feature.Select2` (-1).
   - Required: `If swModel.Save3(...) = False Then`, `If swFeat.IsSuppressed <> False Then`, `ok = (call <> False)`.
   - Wrong: `If Not swModel.Save3(...) Then`, `ok = call: If Not ok Then`, `If call = True Then`,
     `If Not swFeat.IsSuppressed Then` (true for every feature!).
   Also avoid implicit `If result Then` for an unnormalized COM result. `check_vba.py` tracks
   bounded procedure-local provenance; it is not a full VBA compiler or interprocedural type analyzer.
   See [the common COM rule](vba-com-semantics.md).
2. **`ExportToDWG2` on a part loaded as an assembly component fails on the first call** (returns
   False, no file), whatever the arguments. Such documents are invisible and not active. Fix that
   works: `swApp.ActivateDoc3 partTitle, False, swDontRebuildActiveDoc, err` before exporting (then
   re-activate the assembly), *and* verify the file exists - retry once if it doesn't. A part opened
   on its own exports fine on the first call.
3. **Bend lines in flat-pattern DXF need bit 8.** With SheetMetalOptions 1+4 (the documented "bend
   lines" bit) the DXF has no bend lines; with 1+8 or 1+4+8 they appear (linetype CENTERX2). Bit 2048
   adds the bounding box. Use `1 + 4 + 8` (+16 merge faces) and tell the user to check one DXF.
   Some parts export no bend lines with any option (seen on a SOLIDWORKS sample part whose flat
   pattern still lists a Bend-Lines sketch) - don't burn time on it: report which files have none and
   suggest checking that part's flat pattern manually or exporting from a drawing flat-pattern view.
   To test a DXF for bend lines, count entities with a dashed linetype (CENTERX2/CENTER/PHANTOM:
   group code 6) - the word "BEND" doesn't appear in the file.
4. **`ExportPdfData.SetSheets(swExportData_ExportAllSheets, names)` returns False**, yet the PDF
   then contains all sheets. Don't treat that return as an error. `SetSheets(swExportData_ExportSpecifiedSheets,
   swDraw.GetSheetNames)` returns True and also exports all sheets - prefer it.
5. **Named views are localized.** `CreateDrawViewFromModelView3(path, "*Front", ...)` and
   `ShowNamedView2 "*Isometric", -1` silently fail in a Polish SOLIDWORKS ("*Przód", "*Izometryczny",
   "*Góra"). Use view IDs: `swModel.ShowNamedView2 "", swIsometricView` (7; swFrontView 1, swTopView 5),
   or take names from `swModel.GetModelViewNames` (index 1 front, 5 top, 7 isometric).
6. **Feature and configuration names are localized** and differ from what users remember: a Polish
   boss-extrude is `Dodanie-wyciągnięcie1` (not "Wyciągnięcie-Dodanie1"), sketch `Szkic1`, default
   configuration `Domyślna` (English templates/samples: `Default`). Make such names Consts, check
   that `Parameter()` / `FeatureByName()` returned something, and report "not found" per file - or
   find features by type (`GetTypeName2 = "Extrusion"`).
7. **Two files with the same name can't be open at once.** SOLIDWORKS identifies documents by file
   name: while `C:\A\PART1.SLDPRT` is loaded (even invisibly, as a component of an open assembly),
   `OpenDoc6("D:\B\PART1.SLDPRT")` returns the *already open* `C:\A` document - edits and saves land in
   the wrong file. Batch macros over several folders/copies must check
   `LCase(swModel.GetPathName) = LCase(requestedPath)` after opening and skip/log on mismatch; close what
   you open before moving on. When testing on copies, close the user's same-named files first.
8. **Names that are VBA keywords break compilation silently.** A variable called `eNum`, `type`, `name`
   of a Sub like `Open`/`Close`/`Get`/`Put`... is a syntax error; as a `.swb`/RunMacro2 run it just
   doesn't start. Use prefixed names (`enumVal`, `docType`). `check_vba.py` flags declared names that
   are keywords.
9. **VBA file I/O fails on paths longer than ~255 characters** (run-time 52/53/76) - deep PDM folders,
   long temp paths. Keep logs/output in short paths or warn when `Len(path) > 250`.
10. **A macro can "run fine" and produce nothing.** Every test above returned no error dialog. Always
   check the *result* (files exist, values changed), not just that the macro finished.

## Correctness traps

| Trap | What happens | Fix |
|---|---|---|
| Units | API uses meters / radians / kg. `SystemValue = 25` makes a 25 m dimension. | Convert: mm/1000, deg*PI/180. Name variables with units. |
| `Dim a, b As Long` | `a` is Variant, only `b` is Long. Out-params then mismatch. | One `As` per variable. |
| Out-parameters type | `OpenDoc6(..., errors, warnings)` needs Long variables; passing literals or Variants → *ByRef argument type mismatch*. | Declare `Dim errors As Long, warnings As Long`. |
| `GetModelDoc2` returns Nothing | Component suppressed, lightweight, or not loaded. | `ResolveAllLightWeightComponents False`, skip suppressed, or `OpenDoc6` the path. |
| Active vs referenced configuration | Reading a part's properties in the active config instead of the one used in the assembly. | Use `swComp.ReferencedConfiguration`. |
| Same part counted per instance | Writing/exporting the same file N times. | Dedupe with a Dictionary on path (+config). |
| Arrays from API | Return `Empty` when nothing found → `UBound` raises error 13. | `If IsEmpty(v) Then ...` / `If Not IsArray(v)`. |
| Selection indices | 1-based in SelectionManager; arrays are 0-based. | |
| `GetFirstView` | Returns the sheet, not the first view. | Call `GetNextView` once. |
| Drawing sheet activation | View/annotation calls act on the active sheet. | `ActivateSheet` and restore. |
| `Dir()` inside loop with `Dir()` | Dir is not reentrant - nested calls break the outer loop. | Collect file list first. |
| Title vs path | `GetTitle` may omit extension (Windows "hide extensions"). | Use `GetPathName` for logic. |
| Unsaved new doc | `GetPathName = ""`. | Check and ask to save first. |
| Rebuild | Changing dimensions/properties without rebuild → stale mass/geometry/exports. | `EditRebuild3` / `ForceRebuild3 False` before reading results. |
| Localized names | Polish SW: `Szkic1`, `Płaszczyzna przednia`, `Domyślna`; English: `Sketch1`, `Front Plane`, `Default`. | Don't hard-code names - find by type, or make them Consts. |
| Decimal separator | `CDbl("2.5")` fails/misreads on PL locale. | `Val(Replace(s, ",", "."))`; `Str()` for output with dot. |
| Read-only files (PDM / network) | `Save3` fails silently-ish (errors out-param set). | Check return & `errors`; check out in PDM first. |
| Files already open | Macro closes a doc the user had open with unsaved changes. | Remember what was open before; close only what you opened. |
| Several parameterless Subs | SW treats each as an entry point; Tools > Macro > Run may start the wrong one (works with F5, error 91 from the menu). | Exactly one parameterless `Sub main()`; helpers `Private`, Functions or with parameters. |
| `DocumentVisible False` not restored | Session-wide: every document the user opens afterwards is invisible. | Restore `DocumentVisible True` in the cleanup path, also on error. |
| Selection left active before SaveAs | Neutral-format export (STEP, x_t, IGES) contains only the selected bodies/entities. | `ClearSelection2 True` before a full export. |
| `SelectByID2` with coordinates / recorded names | Fails when the entity isn't visible, or selects a neighbour; "Line1" names aren't persistent. | Object methods: `Feature.Select2`, `Component2.Select4`, `Entity.Select4`; API-only selection via `SuspendSelectionList`. |
| Cached property values | `Get6(name, True, ...)` on a non-active configuration returns stale mass/volume. | Pass `UseCached = False`. |
| `Add2` for properties | Doesn't overwrite an existing property. | `Add3` + `swCustomPropertyReplaceValue`. |
| Stale cut-lists | Cut-list bodies/properties out of date or ungrouped. | `UpdateCutList` on the `SolidBodyFolder` before reading. |
| `Is` on SW objects | Two pointers to the same body/feature compare as different. | `swApp.IsSame(a, b) = swObjectSame`. |
| `Sheet.GetViews` | Also returns View Palette views. | `DrawingDoc.GetViews` (index 0 of each sheet = sheet view). |
| `GetPartBox` / `GetBodyBox` | Approximate, can be >10% too big. | `Body2.GetExtremePoint` in ±X/±Y/±Z (recipes.md §4) or the Bounding Box feature. |
| `ForceRebuild3` | Rebuilds the active configuration only. | `Extension.ForceRebuildAll` (SW2017+) for all configurations. |
| Deleting equations / reordering properties | Model not marked dirty → `Save3` saves nothing. | `swModel.SetSaveFlag`. |
| Sketch entities return Nothing | `CreateLine` etc. fail silently when the area is off-screen. | `SketchManager.AddToDB = True` before, restore after. |
| Renaming components | `Name2` get returns full path ("Sub-1/Part-1"), set takes the base name only; fails when "update component names when documents are replaced" is on. | See `advanced-topics.md` §8. |
| Exports renaming the document | `SaveAs` without Copy can re-point the open doc to the new file. | `swSaveAsOptions_Silent + swSaveAsOptions_Copy`. |
| Toolbox / library parts | Writing properties to shared library parts. | Skip them: `swModel.Extension.ToolboxPartType <> swNotAToolboxPart` (0), or path contains the Toolbox folder. |

## Polish characters & encoding

The VBA editor stores code in the system ANSI codepage (cp1250 on Polish Windows). A UTF-8 `.bas`
file imported via File > Import shows `Ä…` instead of `ą` in strings. Options:
- Run `python scripts/check_vba.py macro.bas --cp1250` and import the `_cp1250.bas` file, or
- Paste the code from the clipboard into the editor (pasting converts correctly), or
- Avoid non-ASCII in code; use `ChrW(261)` for `ą` when it must be exact regardless of locale.
Property names with Polish letters (e.g. "Materiał") must match exactly, including diacritics.

## Obsolete methods (prefer the newer one)

| Obsolete | Use |
|---|---|
| `OpenDoc`, `OpenDoc2..5` | `OpenDoc6` (or `OpenDoc7` with DocumentSpecification) |
| `SaveAs2/SaveAs4`, `ModelDoc2.SaveAs` | `Extension.SaveAs` / `SaveAs3` |
| `Save2` | `Save3` |
| `AddCustomInfo3`, `CustomInfo2` | `CustomPropertyManager.Add3` / `Get6` |
| `Get4`, `Get5` | `Get6` (2018+) |
| `SelectByID` | `Extension.SelectByID2` |
| `GetSelectedObject5`, `GetSelectedObjectType2` | `GetSelectedObject6`, `GetSelectedObjectType3` |
| `GetRootComponent`, `GetRootComponent2` | `GetRootComponent3(True)` |
| `GetModelDoc` | `GetModelDoc2` |
| `GetBodies` | `GetBodies2` |
| `ForceRebuild2`/`EditRebuild` | `ForceRebuild3`/`EditRebuild3` |
| `GetMassProperties` | `Extension.CreateMassProperty(2)` |

Old macros using these still often run, but new code should not introduce them.

## Common runtime errors

| Error | Usual cause |
|---|---|
| 91 Object variable or With block variable not set | Something returned Nothing (no active doc, failed open, lightweight component, wrong name in Parameter/FeatureByName). Add null checks. |
| 13 Type mismatch | Empty array from API; assigning ModelDoc2 of wrong type to PartDoc/AssemblyDoc; Variant vs typed out-param. |
| ByRef argument type mismatch | Out-param variable not declared with the exact type. |
| User-defined type not defined | Missing reference (Tools > References > SOLIDWORKS Type Library / Constant type library / Microsoft Scripting Runtime). Or a type name typo. |
| Variable not defined (with Option Explicit) | Typo or undeclared variable - or an `sw...` constant from a library that isn't referenced. |
| 438 Object doesn't support this property or method | Method doesn't exist in this SW version, or wrong object (e.g. calling AssemblyDoc method on a part). |
| 5 Invalid procedure call | Bad argument, e.g. `Left` with negative length when a path has no "\" or ".". |
| -2147417851 (automation error) | SOLIDWORKS crashed / COM call failed - often too many documents open in a long batch. Close every document the macro opened as soon as it's processed (never `CloseAllDocuments` - it closes the user's work); split very long runs (Batch+ restarts SW). |
| Macro "does nothing" | `main` not selected as entry point, or code is in a different module than the one run; or `On Error Resume Next` hiding failures. |
| 91 only from Tools > Macro > Run (works with F5) | Several parameterless Subs - the wrong one was started. |
| 445 Object doesn't support this action | API newer than the installed SW version - check `RevisionNumber` and fall back. |
| 429 ActiveX component can't create object | COM component / DB driver not installed as 64-bit (VBA in SW is 64-bit). |
| Can't find project or library (even on `Left`, `Mid`) | A reference is marked MISSING. |
| Compile error: must be updated for use on 64-bit systems | `Declare` without `PtrSafe` / `LongPtr`. |
| Compile error: Invalid outside procedure (on a long string) | Physical line > 1023 characters - split with `& _`. |
| User-defined type not defined in a `.swb` | `.swb` files can't hold references - save as `.swp` and tick the SOLIDWORKS libraries. |
| Cannot Open (macro) | Macro path > 255 chars, or a VSTA dll shipped without its `bin` folder / blocked by Windows. |
| 424 Object required after merging two macros | Init code (`Set swApp = ...`) or variable names not carried over. |

## References broken after SW upgrade

Macro created in an older version references e.g. "SOLIDWORKS 2022 Type Library" - newer SW
usually remaps automatically. If "MISSING:" appears in Tools > References, untick it and tick the
current version.

## On Error Resume Next

Only around a single call whose failure you then check (`If Err.Number <> 0`), followed by
`On Error GoTo 0` or back to the handler. Never as a blanket at the top of `main` - it turns bugs
into silently wrong output on production files.
