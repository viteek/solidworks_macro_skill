# Advanced topics (distilled from CodeStack)

Expert-level patterns and gotchas, distilled from the CodeStack SOLIDWORKS API library
(https://www.codestack.net/solidworks-api/, source on GitHub `xarial/codestack`). Each section ends
with the source page(s) - when web access is available and you need a full worked example, fetch
`https://raw.githubusercontent.com/xarial/codestack/master/solidworks-api/<path>/index.md` (code files sit
in the same folder). Write your own code; don't paste CodeStack macros wholesale.

## Contents
1. Macro file types, entry points, VBA limits
2. Version checks & compatibility
3. Selection done right
4. Suspending UI, invisible documents, progress bar
5. Events in VBA macros
6. Custom properties - deeper
7. Cut-lists (weldments & sheet metal)
8. Components: names, lookup, visibility, states
9. BOM quantities
10. Dimensions & equations
11. Sheet-metal flat patterns - advanced
12. Export - selection scope, options, single bodies
13. Drawings: views, sheets, property view, BOM tables, notes
14. Tables
15. Geometry: precise bounding box, body copies, transforms
16. Save as previous version, file version without opening
17. Stand-alone (exe) connection details
18. No-API workarounds (WM_COMMAND)

---

## 1. Macro file types, entry points, VBA limits

- `.swp` = binary VBA project: modules, class modules, UserForms, references. `.swb` = plain-text legacy
  macro: one module, **no classes, no forms, no references at all** (not even SOLIDWORKS type libs), so
  early-bound `As SldWorks.ModelDoc2` gives "User-defined type not defined" in the editor. Anything with
  events, forms or early-bound Excel → deliver for a `.swp`. (Run through `RunMacro2`, SOLIDWORKS converts a
  `.swb` to a temporary `.swp` *with* its own type libraries - that's how `scripts/run_in_solidworks.py`
  tests plain-text code. A `.swb` with a compile error then fails silently: `RunMacro2` returns False, no
  dialog - use `check_api.py` to find the cause.)
- **Exactly one parameterless public Sub** (`main`). SOLIDWORKS treats every parameterless Sub as an
  entry point; with several, Tools > Macro > Run may start the wrong one (classic symptom: works with F5,
  fails from the menu with error 91 because init in `main` never ran). Helpers: `Private`, Functions, or
  take a parameter.
- VBA cannot receive arguments (`RunMacro2(path, module, proc, swRunMacroUnloadAfterRun, err)` has no
  args). Workarounds: a settings file next to the macro (`swApp.GetCurrentMacroPathName`), environment
  variable, clipboard, or generating a temp macro from a template.
- Auto-run at startup: `SLDWORKS.exe /m "C:\Macros\x.swp"`.
- Limits: 1023 chars per physical line ("Invalid outside procedure" - split long strings with `& _`);
  ~24 line continuations per statement; macro path ≤ 255 chars ("Cannot Open" - copy to a short path).
- 64-bit VBA (SW2013+): `Declare PtrSafe Function ...` with `LongPtr` handles, else compile error.
  `CreateObject` error 429 = COM component/driver not registered as 64-bit (e.g. ACE OLEDB for Excel).
- "Can't find project or library" (even on `Left`/`Mid`) = a reference marked MISSING. Untick it, tick the
  current SOLIDWORKS libs.
- `InStr` returns 0 when not found (never -1).
- VSTA: ship the whole `bin\Release` folder; keep the project's own namespace and only replace the body of
  `public partial class SolidWorksMacro`; for event/PMP macros untick "Stop VSTA debugger on macro exit".

Src: getting-started/macros/types, troubleshooting/macros/macro-multiple-entry-points,
troubleshooting/macros/swb-macro-error, troubleshooting/macros/too-long-vba-macro-line,
troubleshooting/macros/too-long-macro-path, troubleshooting/macros/32-windows-api-functions-incorrect-use,
getting-started/inter-process-communication/pass-parameters-to-vba-macro

## 2. Version checks & compatibility

- SOLIDWORKS is backward- but not forward-compatible: an API newer than the installed version fails with
  438 / 445. Older revisions (`...4` vs `...5`) often have *different parameter lists*, not just names.
- Runtime check: `major = CInt(Split(swApp.RevisionNumber(), ".")(0))`.
  2015=23, 2016=24, 2017=25, 2018=26, 2019=27, 2020=28, 2021=29, 2022=30, 2023=31, 2024=32, 2025=33.
- File format numbers (`VersionHistory`, `SaveAsPreviousVersion`): n ≥ 5000 → year = 2012 + (n−5000)/1000
  (15000 = 2022).

Src: troubleshooting/macros/future-version-apis, document/rebuild, document/versions-report

## 3. Selection done right

- Avoid `SelectByID2` with coordinates or recorded sketch-entity names ("Line1") - fails when the entity
  isn't visible or picks the wrong one. Rewrite recorded macros to use object methods:
  `Feature.Select2(append, mark)`, `Component2.Select4(append, selData, False)`, `Entity.Select4(append, selData)`,
  `Body2.Select2(append, selData)`.
- `SelectByID2` is fine for persistent names (features, planes, components), drawing **sheets** (no
  `Sheet.Select`), and while a PropertyManager page / command is running.
  Component path: `"Sub-1@TopAssy/Part-2@Sub"`, type `"COMPONENT"`.
- Marks: multi-input features read inputs by `SelectData.Mark` (e.g. FeatureExtrusion2: profile 0,
  up-to surface 1, direction 16). `Set sd = selMgr.CreateSelectData: sd.Mark = 1`.
- **API-only selection that preserves the user's selection:**
  ```vb
  selMgr.SuspendSelectionList
  selMgr.AddSelectionListObject obj, sd      ' or AddSelectionListObjects arr, sd
  ' ... call the API that consumes the selection ...
  selMgr.ResumeSelectionList
  ```
- Many objects at once: `swModel.Extension.MultiSelect2(objArray, append, selData)` → count selected.
  In C# wrap each element in `DispatchWrapper` for mixed types.
- Owning component of a selected face/edge: `selMgr.GetSelectedObjectsComponent4(i, -1)` (dedupe).
- Wait for user selection: class with `WithEvents swPart As SldWorks.PartDoc`, handle `NewSelectionNotify`,
  or a `DoEvents` polling loop.

Src: document/selection/select-by-id, document/selection/api-only-selection, document/selection/select-objects,
document/selection/select-component-by-name, troubleshooting/macros/selection-inconsistency

## 4. Suspending UI, invisible documents, progress bar

```vb
swModel.ActiveView.EnableGraphicsUpdate = False      ' ActiveView may be Nothing for invisible docs
swModel.FeatureManager.EnableFeatureTree = False
swModel.FeatureManager.EnableFeatureTreeWindow = False
' ... work ... ; restore all three in the cleanup path
```
- `swApp.DocumentVisible False, swDocPART` opens subsequent documents of that type invisibly (fast batch
  loading, loading lightweight components' models). It is **session-wide** - if not reset to True (also on
  error), every document the user opens later is invisible.
- Sketch creation: `CreateLine` etc. may silently return Nothing when the area is off-screen. Set
  `swModel.SketchManager.AddToDB = True` (and optionally `DisplayWhenAdded = False`) before, restore after.
- Rebuild: `ForceRebuild3(False)` = active configuration only; `swModel.Extension.ForceRebuildAll` (SW2017+)
  = all configurations; `EditRebuild3` = only dirty features.
- Progress bar for long loops:
  ```vb
  Dim pb As SldWorks.UserProgressBar
  swApp.GetUserProgressBar pb
  pb.Start 0, total, "Exporting"
  If pb.UpdateProgress(i) = swUpdateProgressError_UserCancel Then GoTo Cleanup   ' Esc pressed
  pb.UpdateTitle "File " & i & "/" & total
  pb.End                                                 ' also in cleanup path
  ```
  Use `Long` counters. Status-bar text: `swApp.Frame.SetStatusBarText "..."`.

Src: document/suspend-graphics-update, document/rebuild, application/frame/user-progress-bar,
troubleshooting/macros/create-sketch-segments-error

## 5. Events in VBA macros

- Events need a **class module** (or UserForm) with `Dim WithEvents swApp As SldWorks.SldWorks` /
  `WithEvents swPart As SldWorks.PartDoc` (non-"I" types). Handlers are Functions returning Long (0 = OK).
- Keep the listener alive: module-level variable + modeless UserForm or `While True: DoEvents: Wend`.
  `End` kills all listeners.
- Document events are per type: separate WithEvents for PartDoc / AssemblyDoc / DrawingDoc.
- Save of any kind: `swApp_CommandCloseNotify(Command, reason)` with Save=2, SaveAll=19, SaveAs=620.
  Cancel a UI command: `CommandOpenPreNotify` return 1 (e.g. Edit Sheet Format = 1501).
  Cancel rebuild: `RegenNotify` return 1.
- Since SW2018 custom-property change events fire only for API edits, not UI edits.

Src: application/documents/handle-document-save, document/suspend-rebuild,
document/drawing/lock-sheet-format, data-storage/custom-properties/handle-events

## 6. Custom properties - deeper

- Three sources: file (`CustomPropertyManager("")`), configuration (`CustomPropertyManager(cfg)`), cut-list
  (`cutListFeature.CustomPropertyManager`). Drawings: file-level only.
- BOM-style lookup = **configuration first, then file** (use `comp.ReferencedConfiguration`):
  ```vb
  Function GetPropBom(m As SldWorks.ModelDoc2, cfg As String, nm As String) As String
      Dim v As String, r As String, wr As Boolean, lk As Boolean
      m.Extension.CustomPropertyManager(cfg).Get6 nm, False, v, r, wr, lk
      If r = "" Then m.Extension.CustomPropertyManager("").Get6 nm, False, v, r, wr, lk
      GetPropBom = r
  End Function
  ```
- `UseCached` (2nd arg of Get3..Get6): pass **False** for configuration-dependent expressions (mass, volume)
  or a non-active configuration may return a stale value. `Get6` returns `swCustomInfoGetResult_NotPresent`
  if missing.
- `GetAll3(names, types, vals, resolved, linked)` returns a count; `GetAll2` returns *resolved* values in both
  outputs (expressions lost); `GetNames` returns Empty when none.
- Legacy `Add2` does **not** overwrite an existing property (old code follows it with `Set`). Use `Add3` +
  `swCustomPropertyReplaceValue` and check `= swCustomInfoAddResult_AddedOrChanged`.
- Cut-list properties store any type as Text.
- Linked expression needs doubled quotes: `"""SW-Mass@" & fileName & """"`; config-specific
  `"SW-Mass@@Default@Part1.SLDPRT"`.
- No reorder API: read all (Get with UseCached False), delete, re-add in order, `SetSaveFlag`.
- Summary info (author, keywords, title...) is not a custom property: `swModel.SummaryInfo(swSumInfoAuthor)`
  (read/write).
- SW2022+: `swModel.Extension.UpgradeLegacyCustomProperties(True)` then `ForceRebuildAll`.
- Self-updating values (dates, links to CSV) → a macro feature (`FeatureManager.InsertMacroFeature3`) that
  runs on rebuild - advanced, mention when relevant.

Src: data-storage/custom-properties (+ read-all-properties, write-all-properties, sort, link-sheet-metal,
read-summary-information, update-legacy, auto-date)

## 7. Cut-lists (weldments & sheet metal)

- Cut-list = feature with `GetTypeName2 = "CutListFolder"`; its parent folder is `"SolidBodyFolder"`.
  Both are **sub-features** → traverse `GetFirstSubFeature/GetNextSubFeature` too; skip `"HistoryFolder"`;
  dedupe with `swApp.IsSame(a, b) = swObjectSame`.
- **Call `UpdateCutList` before reading:**
  `If f.GetTypeName2 = "SolidBodyFolder" Then f.GetSpecificFeature2.UpdateCutList`.
- `Set bf = cutFeat.GetSpecificFeature2` → BodyFolder: `GetBodyCount` (skip 0 - empty folders exist),
  `GetBodies`, `SetAutomaticCutList True`, `SetAutomaticUpdate True`.
- Properties per configuration: `ShowConfiguration2`, re-traverse to get fresh feature objects, read, restore.
  For an assembly component, `comp.FirstFeature`… returns cut-lists in its referenced configuration
  without opening the part.
- Typical sheet-metal cut-list properties: "Bounding Box Length", "Bounding Box Width",
  "Sheet Metal Thickness", "Bends", "Cutting Length-Outer", "Material", "Mass", "QUANTITY".
  Detect sheet-metal bodies with `Body2.IsSheetMetal`.
- Body ↔ cut-list: compare bodies with `swApp.IsSame`, never `Is`.
- Exclude: `cutFeat.ExcludeFromCutList = True`. Rename: check
  `FeatureManager.IsNameUsed(swFeatureName, nm)` first, then `feat.Name = nm`.

Src: document/cut-lists (+ rename-cut-list-items, copy-custom-properties, exclude-selected-bodies),
data-storage/custom-properties/read-component-cutlist, import-export/export-individual-cut-lists

## 8. Components: names, lookup, visibility, states

- `Name2` **get** = full path (`Sub-1/Part1-1`, virtual: `Sub-1/Part2^Sub-1`); **set** = short base name only
  (no `-N`, no parents, no `^ctx`). Renaming fails while System Option "Update component names when
  documents are replaced" (`swExtRefUpdateCompNames`) is on.
- Get by full name: split on "/", `assy.FeatureByName(level0)` then `parentComp.FeatureByName(levelN)`,
  `.GetSpecificFeature2` → Component2.
- `comp.ComponentReference = "..."` - shows in BOM, unlike the name.
- Replace: `assy.ReplaceComponents2(newPath, cfg, allInstances, swReplaceComponentsConfiguration_MatchName, reattachMates)`
  acts on the selection → wrap in Suspend/AddSelectionListObject/ResumeSelectionList.
- Visible = not suppressed AND `comp.Visible <> swComponentHidden` AND the same for **every parent**
  (walk `GetParent`).
- Lightweight/suppressed component model without a window: `DocumentVisible False` → `OpenDoc6` Silent →
  `DocumentVisible True` (in cleanup!). Or `ResolveAllLightWeightComponents True`.
- Change suppression in a configuration without activating it:
  `ConfigurationManager.SetConfigurationParams cfg, (names), (values)` with names `"$STATE@Part1<2>"`
  (angle brackets!) and values "S"/"R".

Src: document/assembly/components (rename, get-by-name, replace, get-visible, lightweight-get-model-doc,
batch-modify-parameters, name-to-component-reference)

## 9. BOM quantities

- Parts-only BOM: skip `GetSuppression = swComponentSuppressed` and `ExcludeFromBOM`; group by
  `LCase(path) & "|" & LCase(ReferencedConfiguration)` (path only to merge configurations).
- Respect the subassembly's referenced configuration `ChildComponentDisplayInBOM`:
  `swChildComponent_Promote` → don't count the subassembly, recurse; `_Hide` → count it, don't recurse;
  `_Show` → count and recurse.
- `UNIT_OF_MEASURE` property names *another* property holding the per-instance quantity (else 1).
- Sheet-metal parts: write quantity also to derived flat-pattern configurations
  (`Configuration.Type = swConfiguration_SheetMetal` among `GetChildren`).
- Configuration BOM options: `cfg.BOMPartNoSource` (ConfigurationName / DocumentName / ParentName),
  `cfg.ChildComponentDisplayInBOM`.

Src: document/assembly/compose-flat-bom, document/assembly/components/write-quantities,
document/configuration-bom-options

## 10. Dimensions & equations

- Walk dimensions per feature (recurse sub-features; dedupe - weldments repeat them):
  ```vb
  Set dd = feat.GetFirstDisplayDimension
  Do While Not dd Is Nothing
      Set d = dd.GetDimension2(0)
      Debug.Print d.FullName, d.GetSystemValue3(swThisConfiguration, Empty)(0)
      Set dd = feat.GetNextDisplayDimension(dd)          ' takes the previous one
  Loop
  ```
- `d.Name = "Length"` (short name) also works on drawing dimensions; `d.ReadOnly = True`.
- EquationMgr: no lookup by name - loop `0 To GetCount-1`, parse `Equation(i)` (`"""Name"" = value"`),
  write back the whole line, rebuild. Append: `eqMgr.Add2(-1, """D1@Sketch1"" = 25mm", True)`.
  Delete from last to first; deleting does not set the dirty flag → `SetSaveFlag`. Broken equation:
  read `Value(i)`, then `Status = -1`.

Src: document/dimensions (traverse-all, rename, add-equation), document/change-global-variable-value,
document/delete-model-equations

## 11. Sheet-metal flat patterns - advanced

- Full `SheetMetalOptions` bitmask (declare your own Enum/Consts - not in the type library):
  1 geometry (**always required**), 2 hidden edges, 4 bend lines, 8 sketches, 16 merge coplanar faces,
  32 library features, 64 forming tools, 2048 bounding box. **Measured on SW 2024: bend lines only
  appear when 8 is set** - use 1+4+8(+16). Parts loaded as assembly components must be activated
  (`ActivateDoc3`) or the first export fails; judge success by the file existing (recipes.md §3).
- Call: `swPart.ExportToDWG2(outPath, swModel.GetPathName, swExportToDWG_ExportSheetMetal, True, Empty, False, False, opts, Empty)`
  - Alignment and Views = `Empty` for sheet metal. The part must be saved (ModelName = its path).
  - Output extension decides DXF vs DWG.
- **Multibody:** the selected flat pattern is exported → for each `"FlatPattern"` feature:
  `feat.Select2 False, -1`, export, next. Flat pattern → body: `feat.GetDefinition.FixedFace2` (Face, Edge
  *or* Vertex) `.GetBody`.
- From an assembly: CodeStack makes the component's document visible and active
  (`swApp.ActivateDoc3 path, False, swDontRebuildActiveDoc, err`), activates the referenced configuration,
  freezes tree/graphics, exports, restores, closes what it opened. Offer "skip existing files" so a crashed
  long run can resume.
- `<Cfg>SM-FLAT-PATTERN` configurations (`swConfiguration_SheetMetal`) are created by flat-pattern drawing
  views; corrupt ones can give wrong flats - delete with `DeleteConfiguration2` to regenerate.
- DXF via the built-in Cleanup dialog has no API (only a WM_COMMAND hack - see §18).

Src: document/sheet-metal (export-dxf-dwg, export-all-flat-patterns, get-body-flat-pattern,
remove-flat-pattern-configuration)

## 12. Export - selection scope, options, single bodies

- **If anything is selected, SaveAs to neutral formats exports only the selection** → `ClearSelection2 True`
  before a full export.
- Use `swSaveAsOptions_Silent + swSaveAsOptions_Copy` for exports so the open document isn't renamed /
  re-pointed.
- `Extension.SaveAs2(path, 0, opts, Nothing, "", False, err, warn)`; `SaveAs3(path, 0, opts, exportData, advOpts, err, warn)`.
- Single body to STEP/IGES/x_t: Suspend selection → `AddSelectionListObjects(bodyArr, Nothing)` → SaveAs →
  `ResumeSelectionList2 False`. 3MF/AMF/3DXML/XAML ignore selection → copy bodies (`Body2.Copy`) into a
  temp part (`CreateFeatureFromBody3(copy, False, swCreateFeatureBodySimplify)`) and export that.
- Export options are **application preferences** - save old value, set, export, restore:
  `swStepAP` (203/214), `swSaveIFCFormat`, `swExportOutputCoordinateSystem` (string = coord. system name),
  `swStepExportAppearances`, `swIGESExportAsWireframe`.
- Import: `swApp.LoadFile4(path, "", Nothing, err)`. For batch STEP import disable "prompt for template" and
  3D Interconnect.
- PDF: get a **fresh** `ExportPdfData` per export; `ExportAs3D = True` for 3D PDF of a model.
- Nested output folders: `FileSystemObject.CreateFolder` doesn't create parents - recurse.
- STL from a lightweight component: `Component2.GetTessTriangles(True)` / `GetTessNorms` (meters).

Src: import-export (export-multi-formats, export-individual-bodies, export-bodies, import-step, export-ifc,
export-coordinate-system, export-sheets-pdf, export-stl)

## 13. Drawings: views, sheets, property view, BOM tables, notes

- `Sheet.GetViews` **also returns View Palette views**. Use `swDraw.GetViews` → array per sheet; element 0 of
  each inner array is the sheet view, 1..n are real views.
- Which model drives the sheet's properties: `sheet.CustomPropertyView` (view name; no match → first view).
  Then `view.ReferencedDocument` (Nothing if lightweight/not loaded) + `view.ReferencedConfiguration`.
- Open a model as shown in a view: `ShowConfiguration2 view.ReferencedConfiguration`,
  `GetConfigurationByName(...).ApplyDisplayState view.DisplayState`.
- Drawings of a model without opening them: `swApp.GetDocumentDependencies2(drwPath, False, True, False)`
  → pairs (name, path), step 2.
- `OpenDoc7(spec)` with `swApp.GetOpenDocSpec(path)` (`spec.ReadOnly`, `spec.Silent`, `spec.DetailingMode`),
  error in `spec.Error`.
- Sheet properties: `Sheet.GetProperties2` (0 paperSize, 1 templateType, 2 scale1, 3 scale2, 4 firstAngle,
  5 width, 6 height); new sheet `NewSheet4(...)`, reorder `ReorderSheets(namesArr)`, rename `Sheet.SetName`.
- Replace sheet format: activate sheet → `SetupSheet5(...)` → `sheet.ReloadTemplate(keepMods)` =
  `swReloadTemplate_Success`. Paper sizes: A4=6, A4Vert=7, A3=8, A2=9, A1=10, A0=11.
- BOM tables are **not** linked to a view's configuration: `BomFeature.GetConfigurations(False, vis)` /
  `SetConfigurations(False, vis, names)`.
- `View.GetVisibleComponents` returns only **part** components, drawing context (`GetParent` = Nothing).
- `View.ScaleRatio` = (num, den) array; `View.GetOutline` in sheet meters incl. a ~2% border;
  `View.IsFlatPatternView`.
- Notes: `Note.GetText` (evaluated) vs `Note.PropertyLinkedText` (raw `$PRPSHEET:"..."`); drawing-view notes
  aren't in `Extension.GetAnnotations` - loop views `GetNotes`. Deleting annotations: activate each sheet first.
- Insert BOM (on a model view, sheet active):
  `view.InsertBomTable4(True, x, y, swBOMConfigurationAnchor_TopLeft, swBomType_PartsOnly, "", template, False, swNumberingType_Flat, False)`.

Src: document/drawing (open-referenced-model, open-associated-drawing, insert-sheet, replace-sheet-format,
rename-sheets-custom-properties-values, bom-tables-update-referenced-configuration, get-all-visible-components,
views-size-based-scale), document/notes, document/tables/insert-bom-table, document/set-new-file-name

## 14. Tables

- `TableAnnotation.RowCount`, `ColumnCount`, `Text(row, col)` - 0-based, row 0 = header.
- Find tables: `swDraw.GetViews` → sheet view (index 0) `.GetTableAnnotations` → filter `.Type`
  (`swTableAnnotation_BillOfMaterials`, `_General`, `_RevisionBlock`, `_WeldmentCutList`, `_HoleChart`...).
- Selected tables report `swSelANNOTATIONTABLES`, **revision tables report `swSelREVISIONTABLE`**.
- CSV export: quote cells with `,` `"` or newline; `Print #` writes ANSI - use `ADODB.Stream` for UTF-8.

Src: document/tables (read-table-content-into-array, export-table-csv, find-tables-by-type)

## 15. Geometry: precise bounding box, body copies, transforms

- `GetPartBox` / `Body2.GetBodyBox` are **approximate (can be >10% off)**. Precise: for each solid body
  `body.GetExtremePoint dx, dy, dz, x, y, z` in ±X, ±Y, ±Z → min/max (slower; equals the SW2018+ Bounding Box
  feature). Or insert the Bounding Box feature and read its properties.
- Body pointers die with their document → `Body2.Copy` before closing; never pass the original body to
  `CreateFeatureFromBody3`.
- Compare COM objects with `swApp.IsSame(a, b) = swObjectSame`, not `Is`.
- MathTransform array (16 doubles): [0..8] rotation, [9..11] translation (m), [12] scale, [13..15] unused.

Src: geometry/precise-bounding-box, import-export/export-bodies, import-export/scale-imported-part

## 16. Save as previous version, file version without opening

```vb
Set o = swModel.Extension.GetAdvancedSaveAsOptions(swSaveWithReferencesOptions_None)
o.SaveAsPreviousVersion = 15000          ' SW2022; +1000 per year (target must be supported by SW)
swModel.Extension.SaveAs3 newPath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, o, err, warn
```
`swApp.VersionHistory(path)` → array of `"15000[...]"`: first = created, last = last saved.

Src: document/save-as-previous-version, document/versions-report

## 17. Stand-alone (exe) connection details

- Reference `api\redist\SolidWorks.Interop.sldworks.dll` + `swconst`; **Embed Interop Types = False**.
- `Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"))` may attach or start a new
  session; a new one is invisible, started with `/embed`, **no add-ins** → set `Visible = True`.
  `"SldWorks.Application.31"` targets a version (31 = 2023).
- `Marshal.GetActiveObject` only attaches (and is unavailable on .NET 5+). Specific session: start the
  exe, find `"SolidWorks_PID_" & pid` in the Running Object Table (same elevation required), then wait for
  `OnIdleNotify` before using it.
- `/r` hides the splash; `/b` background mode (Professional/Premium).

Src: getting-started/stand-alone (connect-csharp, connect-vbnet, start-background)

## 18. No-API workarounds (WM_COMMAND)

Some UI commands have no API; CodeStack sends `WM_COMMAND` (&H111) to `swApp.Frame.GetHWnd` (needs a
`PtrSafe` `SendMessage` declare). IDs seen: 51482 Properties dialog of selection, 59526 Edit Material,
54409 select all suppressed components, 24378 combine identical BOM rows, 54244 DXF/DWG export with
Cleanup. Fragile across versions - use only when there is no API and tell the user.

Src: document/cut-lists/browse-custom-properties, document/materials/show-edit-material-dialog,
document/assembly/components/select-suppressed, document/tables/combine-identical-components,
import-export/dwg-export-cleanup
