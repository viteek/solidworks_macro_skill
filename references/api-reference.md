# SOLIDWORKS API - core reference

Compact map of the objects and calls used in 90% of macros. Signatures are VBA-style.
Numeric enum values are given for late binding / Python / C# without the constants assembly.
If a method you need isn't here, prefer the highest-numbered variant you are sure of and flag
uncertainty to the user.

## Contents
1. Object model overview
2. SldWorks (application)
3. ModelDoc2 / ModelDocExtension
4. PartDoc
5. AssemblyDoc / Component2
6. DrawingDoc / Sheet / View
7. Features
8. Selection
9. Custom properties
10. Configurations
11. Dimensions & equations
12. Common enums

---

## 1. Object model overview

```
SldWorks (swApp)
 └─ ModelDoc2 (swModel)  ── cast to PartDoc / AssemblyDoc / DrawingDoc depending on GetType()
     ├─ Extension            (ModelDocExtension: SaveAs, SelectByID2, CustomPropertyManager, CreateMassProperty)
     ├─ SelectionManager     (selected objects)
     ├─ FeatureManager       (create features, EnableFeatureTree)
     ├─ ConfigurationManager (ActiveConfiguration, AddConfiguration2...)
     ├─ SketchManager        (sketch entities)
     ├─ ActiveView           (ModelView: EnableGraphicsUpdate)
     └─ FirstFeature → Feature → GetNextFeature ...
```
In VBA a ModelDoc2 variable can be assigned directly to PartDoc/AssemblyDoc/DrawingDoc variables
(`Set swAssy = swModel`). In C# cast: `(AssemblyDoc)swModel`.

## 2. SldWorks

| Call | Notes |
|---|---|
| `Set swApp = Application.SldWorks` | VBA macro entry. |
| `swApp.ActiveDoc` | ModelDoc2 or Nothing. |
| `swApp.OpenDoc6(path, type, options, config, errors, warnings)` | Returns ModelDoc2. `options` = `swOpenDocOptions_Silent` (1) [+ `ReadOnly` 2, `LoadLightweight` 256...]. `config` = "" for last saved. |
| `swApp.ActivateDoc3(name, useUserPreferences, option, errors)` | Bring an open doc to front. |
| `swApp.CloseDoc(title)` | Pass `swModel.GetTitle` (the document's title, as CodeStack does). |
| `swApp.DocumentVisible(visible, docType)` | False = following OpenDoc6 calls open without windows. Session-wide - always restore True. |
| `swApp.GetUserProgressBar(pb)` | UserProgressBar: `Start`, `UpdateProgress` (returns UserCancel on Esc), `UpdateTitle`, `End`. |
| `swApp.GetOpenDocSpec(path)` + `OpenDoc7(spec)` | Spec: `ReadOnly`, `Silent`, `DetailingMode`, `ConfigurationName`; error in `spec.Error`. |
| `swApp.VersionHistory(path)` | File's SW versions without opening it (first = created, last = last saved). |
| `swApp.IsSame(a, b)` | Compare COM objects (`swObjectSame`) - don't use `Is`. |
| `swApp.CloseAllDocuments(includeUnsaved)` | Closes the user's work too - don't use in macros or as recovery; close only documents the macro opened. |
| `swApp.GetDocuments` | Array of open ModelDoc2. |
| `swApp.GetExportFileData(swExportPdfData)` | For PDF export options (1 = PDF). |
| `swApp.GetUserPreferenceToggle / SetUserPreferenceToggle` | Global options (`swUserPreferenceToggle_e`). |
| `swApp.CommandInProgress = True` | Speeds up long API runs (set False after). |
| `swApp.RevisionNumber` | e.g. "31.x.x" (31 = 2023, 32 = 2024, 33 = 2025). Major = year - 1992. Check with `CInt(Split(swApp.RevisionNumber, ".")(0)) >= 26` before calling newer APIs. |
| `swApp.SendMsgToUser2(msg, icon, buttons)` | SW-styled message box. |
| `swApp.GetCurrentMacroPathName` | Path of the running .swp (useful for relative files). |
| `swApp.RunMacro2(path, module, proc, options, err)` | Run another macro. |

Detect file type from extension: `.sldprt` → `swDocPART` (1), `.sldasm` → `swDocASSEMBLY` (2),
`.slddrw` → `swDocDRAWING` (3).

## 3. ModelDoc2 / ModelDocExtension

| Call | Notes |
|---|---|
| `swModel.GetType` | 1 part, 2 assembly, 3 drawing. |
| `swModel.GetPathName` | "" if never saved. |
| `swModel.GetTitle` | Window title (may lack extension depending on Windows settings). |
| `swModel.Save3(options, errors, warnings)` | Boolean. options `swSaveAsOptions_Silent` (1). |
| `swModel.Extension.SaveAs(path, version, options, exportData, errors, warnings)` | Export by extension (.pdf .step .dxf .x_t .stl .igs .jpg). version `swSaveAsCurrentVersion` (0). `exportData` Nothing or ExportPdfData. 2019+ also has `SaveAs3(..., exportData, advancedSaveAsOptions, errors, warnings)`. |
| `swModel.EditRebuild3` | Rebuild changed features. |
| `swModel.ForceRebuild3(topOnly)` | Full rebuild (Ctrl+Q) of the **active** configuration. |
| `swModel.Extension.ForceRebuildAll` | All configurations (SW2017+). |
| `swModel.SetSaveFlag` | Mark dirty (needed after some API edits, e.g. deleting equations). |
| `swModel.SummaryInfo(swSumInfoAuthor)` | File summary info (author, title, keywords...) - not custom properties. |
| `swModel.ShowConfiguration2(name)` | Activate configuration. |
| `swModel.GetConfigurationNames` | Variant array of strings. |
| `swModel.Parameter("D1@Sketch1")` | Returns Dimension. |
| `swModel.GetEquationMgr` | EquationMgr. |
| `swModel.ClearSelection2(True)` | |
| `swModel.ViewZoomtofit2` | |
| `swModel.ShowNamedView2("", swIsometricView)` | By ID (swFrontView 1, swTopView 5, swIsometricView 7). View **names are localized** ("*Przód", "*Izometryczny" in Polish) - "*Isometric" silently fails there. `GetModelViewNames` lists the local names. |
| `swModel.SetReadOnlyState(False)` | |
| `swModel.GetSaveFlag` | True if dirty. |
| `swModel.Extension.CustomPropertyManager(config)` | "" = file-level. |
| `swModel.Extension.CreateMassProperty` | MassProperty (Mass kg, Volume m³, SurfaceArea m², CenterOfMass). 2020+: `CreateMassProperty2` → MassProperty2. |
| `swModel.Extension.SelectByID2(name, type, x, y, z, append, mark, callout, option)` | See Selection. |
| `swModel.Extension.SetUserPreferenceInteger/Double/Toggle(pref, option, value)` | Document options. |
| `swModel.Extension.Rebuild(swRebuildAll)` | |
| `swModel.ActiveView.EnableGraphicsUpdate = False` | Speed. |
| `swModel.FeatureManager.EnableFeatureTree = False` | Speed (restore!). |

## 4. PartDoc

| Call | Notes |
|---|---|
| `swPart.GetBodies2(bodyType, visibleOnly)` | bodyType `swSolidBody` (0), `swSheetBody` (1). Variant array of Body2 or Empty. |
| `swPart.GetMaterialPropertyName2(config, database)` | Returns material name; `database` out. |
| `swPart.SetMaterialPropertyName2(config, database, name)` | database e.g. "SOLIDWORKS Materials" or "" for custom; config "" = active... use explicit names. |
| `swPart.GetPartBox(True)` | Variant(0..5) xmin,ymin,zmin,xmax,ymax,zmax in m (axis-aligned, approximate). |
| `swPart.ExportToDWG2(file, modelPath, action, single, alignment, isXDirFlipped, isYDirFlipped, sheetMetalOptions, views)` | Flat pattern DXF/DWG; action `swExportToDWG_ExportSheetMetal` (1). Bend lines need options 1+4+8 on SW 2024 (measured). See recipes. |
| `swPart.FeatureByName(name)` | Feature. |
| `swPart.IsWeldment` | |

## 5. AssemblyDoc / Component2

| Call | Notes |
|---|---|
| `swAssy.GetComponents(topLevelOnly)` | Variant array of Component2. False = all levels (flat). |
| `swAssy.ResolveAllLightWeightComponents(warnUser)` | Needed before GetModelDoc2 on lightweight comps. |
| `swAssy.GetComponentCount(topLevelOnly)` | |
| `swAssy.AddComponent5(path, configOption, config, useConfigForPartRefs, newConfig, x, y, z)` | Component must be opened (OpenDoc6) first. |
| `swAssy.ReplaceComponents2(file, config, all, configChoice, reattachMates)` | Acts on selection. |
| `swModel.ConfigurationManager.ActiveConfiguration.GetRootComponent3(True)` | Root Component2 for recursive traversal via `GetChildren`. |
| `swComp.GetChildren` | Variant array. |
| `swComp.GetModelDoc2` | ModelDoc2 - **Nothing** if suppressed / lightweight / virtual-not-loaded. |
| `swComp.GetPathName` | Full path of referenced file. |
| `swComp.Name2` | "Part1-1" or "SubAsm-1/Part1-1". |
| `swComp.ReferencedConfiguration` | Config name used in the assembly - use it for config-specific properties! |
| `swComp.IsSuppressed` / `GetSuppression2` / `SetSuppression2` | `swComponentSuppressed` 0, `Lightweight` 1, `FullyResolved` 2, `Resolved` 3. |
| `swComp.ExcludeFromBOM` | Boolean property. |
| `swComp.IsVirtual`, `IsEnvelope`, `IsPatternInstance` | |
| `swComp.Transform2` | MathTransform. |
| `swComp.Select4(append, selData, showPopup)` | |

## 6. DrawingDoc / Sheet / View

| Call | Notes |
|---|---|
| `swDraw.GetSheetNames` | Variant array. |
| `swDraw.ActivateSheet(name)` / `swDraw.GetCurrentSheet` | Sheet. |
| `swSheet.GetProperties2` | [paperSize, template, scale1, scale2, firstAngle, width, height, ...]. |
| `swDraw.GetViews` | Array per sheet; element 0 = sheet view, 1..n = views. Prefer over `Sheet.GetViews`, which also returns View Palette views. |
| `swSheet.CustomPropertyView` | Name of the view whose model feeds the sheet's `$PRP` properties. |
| `swDraw.GetFirstView` | **The sheet itself** (as a View); `GetNextView` gives the real views of the active sheet. |
| `swView.GetName2`, `Type` | |
| `swView.ReferencedDocument` | ModelDoc2 of the view's model. |
| `swView.GetReferencedModelName`, `ReferencedConfiguration` | |
| `swView.GetFirstTableAnnotation` / `GetTableAnnotations` | BOM tables etc. |
| `swView.GetFirstNote` → `GetNext` | Notes. |
| `swDraw.NewSheet4(...)`, `swDraw.CreateDrawViewFromModelView3(model, viewName, x, y, z)` | Create views (m coordinates on sheet). `viewName` must be the **localized** name - take it from the model's `GetModelViewNames` (index 1 front, 5 top, 7 iso); a wrong name returns Nothing without error. |
| `swDraw.InsertModelAnnotations3(...)` | Import model dimensions. |

## 7. Features

```vb
Set swFeat = swModel.FirstFeature
Do While Not swFeat Is Nothing
    Debug.Print swFeat.Name, swFeat.GetTypeName2
    Set swSub = swFeat.GetFirstSubFeature   ' sketches under a feature etc.
    Set swFeat = swFeat.GetNextFeature
Loop
```
Useful: `swFeat.IsSuppressed2(swThisConfiguration, Nothing)`, `swFeat.SetSuppression2(state, opt, Nothing)`,
`swFeat.GetDefinition` (feature data object, then `ModifyDefinition`), `swFeat.Name = "New"`.
Type names (GetTypeName2) e.g. `ProfileFeature` (sketch), `Extrusion`, `Cut`, `ICE` (boss-extrude in
some versions), `SheetMetal`, `FlatPattern`, `CutListFolder`, `SolidBodyFolder`, `HoleWzd`,
`MateGroup`, `Reference` (component in tree), `RefPlane`. Verify with a Debug.Print run if unsure.

Cut lists (weldments / sheet metal): find features with `GetTypeName2 = "CutListFolder"`,
`swFeat.GetSpecificFeature2` → BodyFolder (`GetBodyCount`, `GetBodies`), and properties via
`swFeat.CustomPropertyManager`.

## 8. Selection

```vb
Set swSelMgr = swModel.SelectionManager
n = swSelMgr.GetSelectedObjectCount2(-1)          ' -1 = all marks
For i = 1 To n                                    ' 1-based!
    t = swSelMgr.GetSelectedObjectType3(i, -1)    ' swSelectType_e
    Set obj = swSelMgr.GetSelectedObject6(i, -1)
    Set swComp = swSelMgr.GetSelectedObjectsComponent4(i, -1)
Next
```
`SelectByID2(name, type, x, y, z, append, mark, callout, option)`, type strings: `"PLANE"`,
`"SKETCH"`, `"BODYFEATURE"`, `"COMPONENT"`, `"DIMENSION"`, `"EDGE"`, `"FACE"`, `"DRAWINGVIEW"`,
`"SHEET"`. Component names in SelectByID2: `"Part1-1@Assem1"`; nested `"Sub-1@Assem1/Part1-1@Sub"`.
Pass `Nothing` for callout and `0` for option.

## 9. Custom properties (CustomPropertyManager)

```vb
Set cpm = swModel.Extension.CustomPropertyManager("")          ' file-level
Set cpm = swModel.Extension.CustomPropertyManager("Default")   ' configuration-specific
```
| Call | Notes |
|---|---|
| `cpm.Add3(name, type, value, option)` | type `swCustomInfoText` (30), `swCustomInfoNumber` (3), `swCustomInfoDate` (64), `swCustomInfoYesOrNo` (11), `swCustomInfoDouble` (5). option `swCustomPropertyDeleteAndAdd` (1), `swCustomPropertyOnlyIfNew` (0), `swCustomPropertyReplaceValue` (2). Returns `swCustomInfoAddResult_e` (0 = added). |
| `cpm.Get6(name, useCached, valOut, resolvedOut, wasResolved, linkToProp)` | 2018+. Returns `swCustomInfoGetResult_e` (`_NotPresent` if missing). Pass `useCached = False` for config-dependent values (mass). Older: `Get5(name, useCached, valOut, resolvedOut, wasResolved)`, `Get4`. |
| `cpm.Set2(name, value)` | Change value only. |
| `cpm.Delete2(name)` | |
| `cpm.GetNames` | Variant array or Empty. |
| `cpm.GetAll3(names, types, values, resolved, links)` | Returns count. |
| `cpm.Count` | |

`valOut` = raw expression (e.g. `"SW-Mass@Part1.SLDPRT"`), `resolvedOut` = evaluated value. Linked
expressions: `"""SW-Mass@" & fileName & """"` - quoting matters.

## 10. Configurations

| Call | Notes |
|---|---|
| `swModel.GetConfigurationNames` | |
| `swModel.ConfigurationManager.ActiveConfiguration` | Configuration; `.Name`. |
| `swModel.ShowConfiguration2(name)` | Activates (triggers rebuild). |
| `swModel.GetConfigurationByName(name)` | Configuration. |
| `swModel.ConfigurationManager.AddConfiguration2(name, comment, altName, options, parent, description, rebuild)` | Newer API; older code uses `swModel.AddConfiguration3(...)`. Signatures changed between versions - verify for the user's version and flag it. |
| `swModel.DeleteConfiguration2(name)` | Not the active one. |
| `swConf.GetParent`, `GetChildren`, `IsDerived` | |

## 11. Dimensions & equations

```vb
Set swDim = swModel.Parameter("D1@Boss-Extrude1")
swDim.SetSystemValue3 0.025, swSetValue_InThisConfiguration, Nothing   ' meters
v = swDim.GetSystemValue3(swThisConfiguration, Nothing)               ' returns Variant array
swModel.EditRebuild3
```
`swSetValue_InThisConfiguration` 1, `swSetValue_InAllConfigurations` 2, `swSetValue_InSpecificConfigurations` 3.
Simpler single-config: `swDim.SystemValue = 0.025` (meters) / `swDim.Value = 25` (document units).
`GetSystemValue3` returns an array - take `(0)`.

EquationMgr: `.GetCount`, `.Equation(i)` get/set the text (`"""Length"" = 250mm"`), `.Value(i)`,
`.GlobalVariable(i)`, `.Add3(index, text, solve, config, configNames)`, `.EvaluateAll`.

## 12. Common enums (numeric values)

| Enum | Value |
|---|---|
| swDocPART / ASSEMBLY / DRAWING | 1 / 2 / 3 |
| swOpenDocOptions_Silent | 1 |
| swOpenDocOptions_ReadOnly | 2 |
| swOpenDocOptions_ViewOnly | 4 |
| swOpenDocOptions_LoadLightweight | 256 |
| swSaveAsCurrentVersion | 0 |
| swSaveAsOptions_Silent | 1 |
| swSaveAsOptions_Copy | 2 |
| swSaveAsOptions_SaveReferenced | 4 |
| swExportPdfData | 1 |
| swExportData_ExportAllSheets / Current / Specified | 1 / 2 / 3 |
| swCustomInfoText / Number / Double / YesOrNo / Date | 30 / 3 / 5 / 11 / 64 |
| swCustomPropertyOnlyIfNew / DeleteAndAdd / ReplaceValue | 0 / 1 / 2 |
| swSolidBody / swSheetBody | 0 / 1 |
| swExportToDWG_ExportSheetMetal | 1 |
| swThisConfiguration / AllConfiguration / SpecifyConfiguration | 1 / 2 / 3 |
| swComponentSuppressed / Lightweight / FullyResolved / Resolved | 0 / 1 / 2 / 3 |
| swMbWarning / Information / Question / Stop | 1 / 2 / 3 / 4 |
| swMbOk / OkCancel / YesNo / YesNoCancel | 2 / 3 / 5 / 6 (verify) |

Values marked "verify" - prefer the named constant with early binding.
