# Recipes - tested patterns (VBA)

Snippets assume early binding and these declarations where relevant:
```vb
Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim errors As Long, warnings As Long
```

## Contents
1. Custom properties (read / write / config-specific / linked)
2. Export: PDF (drawing), STEP / Parasolid / STL, image
3. Sheet-metal flat pattern to DXF
4. Mass properties & bounding box
5. Material
6. Dimensions & configurations
7. Drawing: iterate sheets & views, open referenced model
8. Assembly: unique parts list with quantities (BOM-like)
9. Files & folders (Dir, FileSystemObject, CSV log)
10. Simple user input (InputBox, folder picker)
11. Performance & state restore
12. Same-name drawing for a model

---

## 1. Custom properties

```vb
Function GetProp(m As SldWorks.ModelDoc2, cfg As String, propName As String) As String
    Dim cpm As SldWorks.CustomPropertyManager
    Dim valOut As String, resolved As String, wasResolved As Boolean, linked As Boolean
    Set cpm = m.Extension.CustomPropertyManager(cfg)
    cpm.Get6 propName, False, valOut, resolved, wasResolved, linked
    GetProp = resolved
End Function

Function SetProp(m As SldWorks.ModelDoc2, cfg As String, propName As String, value As String, _
                 Optional replaceLinked As Boolean = False) As Boolean
    ' Creates or updates a text property. A LINKED value ("SW-Mass@...", $PRP...) - often from the
    ' company template - is kept unless replaceLinked = True: overwriting it with text breaks the link.
    Dim cpm As SldWorks.CustomPropertyManager, raw As String, res As String, wasRes As Boolean, linked As Boolean
    Set cpm = m.Extension.CustomPropertyManager(cfg)
    If cpm.Get6(propName, False, raw, res, wasRes, linked) <> swCustomInfoGetResult_NotPresent Then
        If raw = value Then SetProp = True: Exit Function                       ' already current
        If (linked Or InStr(raw, """") > 0 Or InStr(1, raw, "$PRP", vbTextCompare) > 0) And Not replaceLinked Then
            Exit Function                                                         ' keep the link; report it
        End If
    End If
    SetProp = (cpm.Add3(propName, swCustomInfoText, value, swCustomPropertyReplaceValue) = swCustomInfoAddResult_AddedOrChanged)
End Function
```
- Config-specific value wins over file-level in BOMs. For an assembly component, read with
  `swComp.ReferencedConfiguration`, not the active configuration of the part.
- Linked mass property: `SetProp swModel, "", "Masa", """SW-Mass@" & fileNameWithExt & """"`.
  Using `@@config@` form for config-specific: `"SW-Mass@@Default@Part1.SLDPRT"`.
- After changing properties the document must be saved (`Save3`) to persist.
- `swCustomInfoAddResult_AddedOrChanged` = 0. If unsure of the constant, compare with 0.

## 2. Export

### Drawing → PDF (all sheets, one file)
```vb
Dim swExportPdf As SldWorks.ExportPdfData
Dim sheetNames As Variant
Set swExportPdf = swApp.GetExportFileData(swExportPdfData)
sheetNames = swModel.GetSheetNames                 ' swModel is a drawing
' ExportAllSheets makes SetSheets return False (SW 2024) although it works -> use Specified + all names
swExportPdf.SetSheets swExportData_ExportSpecifiedSheets, sheetNames
swExportPdf.ViewPdfAfterSaving = False
swModel.Extension.SaveAs pdfPath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, swExportPdf, errors, warnings
If Dir(pdfPath) = "" Then  ' judge by the file, not by "Not ok" (API True = 1 -> Not ok is True too)
    ' log failure with errors / warnings
End If
```
Works on drawings opened `swOpenDocOptions_Silent + swOpenDocOptions_ReadOnly` (tested) - use read-only
so the drawings can't be modified by the export macro.
One PDF per sheet: loop sheets, `swExportData_ExportSpecifiedSheets` with a single-element array, and get a
**fresh** `ExportPdfData` from `GetExportFileData` for every export (don't reuse one object).
Without `SetSheets` only the active sheet is exported.

### Drawing → DWG/DXF (sheets)
`swModel.Extension.SaveAs(path & ".dxf", 0, swSaveAsOptions_Silent, Nothing, errors, warnings)`
- multi-sheet behavior is controlled by system option `swDxfMultiSheetOption`
(`swApp.SetUserPreferenceIntegerValue swDxfMultiSheetOption, swDxfActiveSheetOnly / swDxfExportAllSheetsToSeparateFiles / swDxfExportAllSheetsToOneFile`).

### Part/Assembly → STEP, Parasolid, IGES, STL
```vb
swModel.ClearSelection2 True     ' with a selection, neutral formats export ONLY the selected entities/bodies
ok = swModel.Extension.SaveAs(outPath, swSaveAsCurrentVersion, swSaveAsOptions_Silent + swSaveAsOptions_Copy, _
                              Nothing, errors, warnings)   ' Copy: the open document is not renamed
```
Extension (`.step`, `.x_t`, `.igs`, `.stl`, `.3mf`) picks the format. Exports the active configuration -
`ShowConfiguration2` first to export another one, restore afterwards.
Format options are **application preferences** - save, set, export, restore:
```vb
oldAp = swApp.GetUserPreferenceIntegerValue(swStepAP)
swApp.SetUserPreferenceIntegerValue swStepAP, 214           ' 203 / 214
' ... export ...
swApp.SetUserPreferenceIntegerValue swStepAP, oldAp
```
Single body export, coordinate-system output and formats that ignore selection: `advanced-topics.md` §12.

### Image
`swModel.ShowNamedView2 "", swIsometricView : swModel.ViewZoomtofit2` then SaveAs with `.png`/`.jpg`.
Use the view ID, not the name - names are localized ("*Isometric" fails in a Polish SOLIDWORKS).

## 3. Sheet-metal flat pattern → DXF

Use the complete [safe export template](../assets/templates/safe_export_template.bas) as the source of
export/publication logic, rather than a second implementation here. Set `OUTPUT_PATH` to the `.dxf`
destination and `CONFIGURATION_NAME` if needed. Keep `ALLOW_OVERWRITE=False` by default; enable it only
when replacing that output is intended. The template stages the export, checks API status and nonzero
size, retains a uniquely named backup, and attempts to restore the original target if publication fails.
An existing file by itself is not evidence that the current export succeeded.

For assembly components, activate the verified component document with `ActivateDoc3` and use its
`ReferencedConfiguration` before exporting. Reuse the template's `ExportToStage`, `PublishStage`, and
cleanup logic together in the adapted module. A retry, if needed on SW2024 SP5, belongs only around
the staged export, at most once; never retry publication or delete an existing final/backup file.
Restore the assembly/configuration afterwards and close only documents opened by the macro.

SheetMetalOptions (measured on SW2024 SP5; see pitfalls.md): 1 = geometry, 2 = hidden edges,
4 = bend lines, 8 = sketches, 16 = merge coplanar faces, 32 = library features, 64 = forming tools,
2048 = bounding box. The template uses 1+4+8 for bend lines; inspect the actual DXF result.
- The part must be saved (`GetPathName <> ""`). The output extension (.dxf/.dwg) picks the format.
- Detect sheet metal: any solid body with `Body2.IsSheetMetal` (works also via `Component2.GetBodies3`), or a
  `"SheetMetal"` / `"FlatPattern"` feature (note: `FlatPattern` and cut-lists are **sub-features**).
- **Multibody** sheet metal: `ExportToDWG2` exports the *selected* flat pattern → for each `"FlatPattern"`
  feature: `feat.Select2 False, -1`, export to its own file. Unselected multibody = unpredictable.
- Thickness: from the SheetMetal feature definition (`GetDefinition` → `SheetMetalFeatureData.Thickness`, m)
  or the cut-list property "Sheet Metal Thickness" - good for file naming like `PART_t2.dxf`.
- From an assembly, CodeStack activates the component's document (`ActivateDoc3`) and its referenced
  configuration before exporting - see `advanced-topics.md` §11. Offer "skip existing files" for long runs
  (SW can crash on big batches; a rerun then resumes).

## 4. Mass properties & bounding box

```vb
Dim swMass As SldWorks.MassProperty
Set swMass = swModel.Extension.CreateMassProperty
If Not swMass Is Nothing Then
    massKg = swMass.Mass
    volumeM3 = swMass.Volume
End If
```
- For assemblies, mass uses resolved components; lightweight may be inaccurate → resolve first.
- Overridden mass ("Override mass properties") is respected by CreateMassProperty2 (2020+) - prefer it if available.

Bounding box - `GetPartBox` / `Body2.GetBodyBox` are **approximate (can be >10% oversized)**; fine for a
rough check, wrong for cutting stock. Precise, axis-aligned, per solid body:
```vb
Function PreciseBox(swPart As SldWorks.PartDoc) As Variant   ' returns minX,minY,minZ,maxX,maxY,maxZ in m
    Dim vBodies As Variant, i As Long, b As SldWorks.Body2
    Dim x As Double, y As Double, z As Double, r(5) As Double
    r(0) = 1E+30: r(1) = 1E+30: r(2) = 1E+30: r(3) = -1E+30: r(4) = -1E+30: r(5) = -1E+30
    vBodies = swPart.GetBodies2(swSolidBody, True)
    If IsEmpty(vBodies) Then Exit Function
    For i = 0 To UBound(vBodies)
        Set b = vBodies(i)
        b.GetExtremePoint -1, 0, 0, x, y, z: If x < r(0) Then r(0) = x
        b.GetExtremePoint 0, -1, 0, x, y, z: If y < r(1) Then r(1) = y
        b.GetExtremePoint 0, 0, -1, x, y, z: If z < r(2) Then r(2) = z
        b.GetExtremePoint 1, 0, 0, x, y, z: If x > r(3) Then r(3) = x
        b.GetExtremePoint 0, 1, 0, x, y, z: If y > r(4) Then r(4) = y
        b.GetExtremePoint 0, 0, 1, x, y, z: If z > r(5) Then r(5) = z
    Next i
    PreciseBox = r
End Function
```
Alternative: the Bounding Box feature (SW2018+) whose properties appear as
`Total Bounding Box Length/Width/Thickness`; sheet-metal cut-lists have "Bounding Box Length/Width".

## 5. Material

```vb
Dim db As String
matName = swPart.GetMaterialPropertyName2(configName, db)
swPart.SetMaterialPropertyName2 configName, "SOLIDWORKS Materials", "AISI 304"
```
Material name must exist in the given database (exact spelling, language of the SW installation!).

## 6. Dimensions & configurations

```vb
Dim swDim As SldWorks.Dimension
Set swDim = swModel.Parameter("D1@Szkic1")   ' names are localized in Polish SW: Szkic1, Dodanie-wyciągnięcie1 (SW 2024 PL)
If swDim Is Nothing Then MsgBox "Dimension not found": Exit Sub
swDim.SetSystemValue3 lengthMm / 1000#, swSetValue_InThisConfiguration, Nothing
swModel.EditRebuild3
```
Loop all configurations:
```vb
Dim cfgNames As Variant, i As Long
cfgNames = swModel.GetConfigurationNames
For i = 0 To UBound(cfgNames)
    swModel.ShowConfiguration2 CStr(cfgNames(i))
    ' ... work ...
Next
```
Restore the originally active configuration at the end.

## 7. Drawing: sheets & views

```vb
Dim swDraw As SldWorks.DrawingDoc, swView As SldWorks.View
Dim sheetNames As Variant, s As Long
Set swDraw = swModel
originalSheet = swDraw.GetCurrentSheet.GetName
sheetNames = swDraw.GetSheetNames
For s = 0 To UBound(sheetNames)
    swDraw.ActivateSheet sheetNames(s)
    Set swView = swDraw.GetFirstView          ' sheet "view"
    Set swView = swView.GetNextView           ' first real view
    Do While Not swView Is Nothing
        Debug.Print sheetNames(s), swView.GetName2, swView.GetReferencedModelName, swView.ReferencedConfiguration
        Set swView = swView.GetNextView
    Loop
Next
swDraw.ActivateSheet originalSheet
```
All sheets without activating each one: `vSheets = swDraw.GetViews` → one array per sheet; element 0 is
the sheet view, 1..n are real views. (Don't use `Sheet.GetViews` for this - it also returns View Palette views.)

Model of the drawing (for property-based PDF names): use the view the sheet takes its properties from -
`swSheet.CustomPropertyView` (a view name; if it doesn't match any view, fall back to the first real view).
Then `view.ReferencedDocument` + `view.ReferencedConfiguration`, reading properties configuration-first,
then file (`advanced-topics.md` §6). `ReferencedDocument` is Nothing if the view is lightweight / the model
isn't loaded → open `view.GetReferencedModelName` silently (and close it afterwards).

## 8. Assembly: unique parts with quantities

Use a `Scripting.Dictionary` keyed by `LCase(path) & "|" & config`:
```vb
Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
' inside traversal, for each non-suppressed, non-excluded component:
key = LCase(swComp.GetPathName) & "|" & swComp.ReferencedConfiguration
If dict.Exists(key) Then dict(key) = dict(key) + 1 Else dict.Add key, 1
```
- Skip `swComp.IsSuppressed`, `swComp.ExcludeFromBOM` (if BOM semantics), envelopes.
- Decide whether subassemblies count as items or are dissolved - ask the user, BOM types differ
  ("top level only" vs "parts only" vs "indented"). To mimic SOLIDWORKS, honour the subassembly's
  `ChildComponentDisplayInBOM` (Promote / Hide / Show) - rules in `advanced-topics.md` §9.
- Virtual components have paths inside the assembly; `GetModelDoc2` works when resolved.
- Toolbox parts: you may want to skip writing to them (read-only / shared library).

## 9. Files & folders

```vb
' Non-recursive file listing (Dir is not reentrant - collect names first, then process)
Dim files As New Collection, f As String
f = Dir(folder & "\*.sldprt")
Do While f <> "": files.Add folder & "\" & f: f = Dir(): Loop
```
Recursive: use `CreateObject("Scripting.FileSystemObject")` and a recursive Sub over `.SubFolders`.
Skip temporary/lock files starting with `~$`.

CSV log (Excel-friendly in PL locale uses `;`):
```vb
Dim ff As Integer: ff = FreeFile
Open logPath For Output As #ff
Print #ff, "File;Status;Message"
Print #ff, path & ";OK;"
Close #ff
```
For UTF-8 with Polish characters use `ADODB.Stream` (Charset "utf-8") instead of `Print #`.

Paths: `Left(p, InStrRev(p, "\"))` = folder, `Mid(p, InStrRev(p, "\") + 1)` = file name,
strip extension with `Left(name, InStrRev(name, ".") - 1)`. Make output dir with `MkDir` (check `Dir(path, vbDirectory)` first).

## 10. User input

```vb
answer = InputBox("Thickness [mm]:", "Macro", "2")
If answer = "" Then Exit Sub      ' Cancel
```
Folders: put them in a `Const` that may be empty, and show the picker only when it is empty. The user
can then hard-code a folder for repeated runs, the macro works under Batch+/Task Scheduler, and
`run_in_solidworks.py --define` can test it unattended:
```vb
Const SOURCE_FOLDER As String = ""   ' empty = ask
...
folder = SOURCE_FOLDER
If folder = "" Then folder = PickFolder("Select the drawings folder")
If folder = "" Then Exit Sub         ' cancelled
```
Folder picker without references (Shell):
```vb
Function PickFolder(title As String) As String
    Dim sh As Object, fld As Object
    Set sh = CreateObject("Shell.Application")
    Set fld = sh.BrowseForFolder(0, title, 0)
    If Not fld Is Nothing Then PickFolder = fld.Self.Path
End Function
```
For more than 2-3 inputs build a UserForm (export `.frm` + `.frx` - the `.frx` is binary, so describe
the controls and let the user draw the form, or create controls at runtime).

Decimal separator: `CDbl("2,5")` depends on Windows locale (PL uses comma). Parse with
`Val(Replace(s, ",", "."))` for locale-independent input.

## 11. Performance & state restore

```vb
Dim swModelView As SldWorks.ModelView
Set swModelView = swModel.ActiveView
swApp.CommandInProgress = True
If Not swModelView Is Nothing Then swModelView.EnableGraphicsUpdate = False
swModel.FeatureManager.EnableFeatureTree = False
On Error GoTo Cleanup
' ... work ...
Cleanup:
swModel.FeatureManager.EnableFeatureTree = True
If Not swModelView Is Nothing Then swModelView.EnableGraphicsUpdate = True
swApp.CommandInProgress = False
If Err.Number <> 0 Then MsgBox "Error " & Err.Number & ": " & Err.Description
```
Also `swModel.FeatureManager.EnableFeatureTreeWindow = False` (restore). `ActiveView` can be Nothing for
invisible documents - null-check.

Batch: open documents with `swOpenDocOptions_Silent`. For speed, `swApp.DocumentVisible False, swDocPART`
opens them without windows - but it is a **session-wide** setting: restore `DocumentVisible True` in the
cleanup path, or every document the user opens later stays invisible. `swApp.Visible = False` only for a
standalone exe. Close every document you opened with `swApp.CloseDoc swModel.GetTitle` (but never ones that
were already open before - check `swApp.GetOpenDocumentByName(path)` first). Long loops: show progress
with `UserProgressBar` (Esc = cancel) - `advanced-topics.md` §4.

## 12. Drawing with the same name as the model

```vb
drwPath = Left(modelPath, InStrRev(modelPath, ".")) & "SLDDRW"
If Dir(drwPath) <> "" Then Set swDrw = swApp.OpenDoc6(drwPath, swDocDRAWING, swOpenDocOptions_Silent, "", errors, warnings)
```
