# Creating parts with the API (sketches, features, materials)

Read this for macros that BUILD geometry: new parts, sketches, extrudes/cuts/revolves, holes,
fillets, patterns, sheet metal, parametric families. Start from `assets/templates/create_part_template.bas`
(a verified parametric plate: sketch → extrude → holes → fillets → material → save → self-check).
Facts marked *measured* were observed on SOLIDWORKS 2024 SP5, Polish UI.

## Contents
1. The build loop
2. New document, units, planes
3. Sketching
4. Dimensions and parameters
5. Features (look up the signature first)
6. Selecting faces and edges
7. Materials
8. Sheet metal
9. Configurations / families
10. Checking the result

---

## 1. The build loop

```
new part from template  →  select plane/face  →  InsertSketch  →  AddToDB=True, create entities
→ (dimensions)  →  InsertSketch (close)  →  select the sketch feature  →  FeatureXxx(...)
→ If feat Is Nothing → stop with a clear message  →  next feature ... → ForceRebuild3 → check → SaveAs
```
- Every `FeatureManager.Feature*` call works on the **current selection** and returns the new
  `Feature` or **Nothing** on failure (no error is raised). Check after every feature - otherwise the
  next step silently builds on missing geometry.
- Close the sketch (`InsertSketch True` again) before creating the feature from it, then select it:
  `Set sk = model.FeatureByPositionReverse(0)` (newest feature) → `sk.Select2 False, 0`.
- Feature names are localized: *measured* `Szkic1`, `Dodanie-wyciągnięcie1`, `Wytnij-wyciągnięcie1`
  (English: `Sketch1`, `Boss-Extrude1`, `Cut-Extrude1`). Keep the returned `Feature` object or rename
  it (`feat.Name = "PLYTA"`) instead of finding it by name later.
- Don't build half-finished parts into existing files: create, check, then `SaveAs` to a new path.
- Build exactly the geometry that was specified. Realistic additions (tube corner radii per EN 10219,
  chamfers, reliefs) change mass and section values the user may rely on - put them in `Const`s that
  default to 0/off and mention them in one line. Also flag specification problems you notice (e.g. a
  hole close to a corner fillet leaving ~2 mm of material) instead of silently changing the design.
- When the user asks for the output folder "in a constant", give it a clear example path
  (`Const OUTPUT_FOLDER As String = "C:\CAD\Czesci"`, "change this") - a picker only when it's empty.
- Self-checks should repair what they can: e.g. a sheet-metal base flange with the wrong thickening
  side (size off by 2·t) → delete/undo the feature and rebuild it with the opposite `ThickenDir`,
  then check again; abort with a clear message only if that fails too.

## 2. New document, units, planes

- Template: `swApp.GetUserPreferenceStringValue(swDefaultTemplatePart)` (*measured*: the company PDM
  template `D:\PDM\...\Część.prtdot`). Check `Dir(tpl) <> ""` - a PDM template may not be local.
  `Set model = swApp.NewDocument(tpl, 0, 0, 0)` → Nothing on failure. The template decides units,
  default material, custom properties and configuration name (`Domyślna`).
- API values are always metres/radians whatever the document units; document units only matter for
  `Dimension.Value` and displayed text. Check: `model.Extension.GetUserPreferenceInteger(swUnitsLinear, 0)` (0 = mm).
- **Standard planes: select by order, not by name.** The first three `RefPlane` features are Front,
  Top, Right (*measured* names `Płaszczyzna przednia / górna / prawa`; `SelectByID2 "Front Plane"` fails
  in a Polish SW). Helper `StandardPlane(model, 0|1|2)` in the template.
- Front plane: sketch X = model X, sketch Y = model Y, extrusion along +Z. Top plane: sketch X = model X,
  sketch Y = model **-Z**. Right plane: sketch X = model **-Z**, sketch Y = model Y (SOLIDWORKS convention,
  not measured here). On faces the sketch axes follow the face orientation - verify with a test build
  (measured: on the +Z top face of a plate sketched on Front, sketch X/Y = model X/Y).

## 3. Sketching

- `model.SketchManager.AddToDB = True` while creating entities (else entities off-screen or near
  existing geometry can snap/fail silently); set it back to False before dimensioning and at the end.
  `DisplayWhenAdded = False` speeds up large sketches.
- Coordinates in metres, relative to the sketch plane.
- `CreateCornerRectangle(x1,y1,0,x2,y2,0)` → 4 lines. `CreateCenterRectangle(cx,cy,0,cornerX,cornerY,0)`
  → *measured* **6** segments (4 sides + 2 construction diagonals) - don't assume 4.
  `CreateCircleByRadius(x,y,0,r)`, `CreateLine(x1,y1,0,x2,y2,0)`, `CreateArc(...)`, `CreateCenterLine(...)`
  (for revolve axes), `CreateSketchSlot(...)`.
- A sketch for an extrude must be closed loops; open profiles only for thin features / sheet metal.
- Relations: `model.SketchAddConstraints "sgHORIZONTAL2D"` on selected entities; usually unnecessary
  when dimensions drive the geometry.

## 4. Dimensions and parameters

- `model.AddDimension2(x, y, 0)` dimensions the selected entity, text placed at (x,y). **Measured:
  with the user option "Input dimension value" on (`swInputDimValOnCreate`, it was on), it opens a modal
  "Modyfikuj" box and the macro stops.** Save the toggle, set False, restore in cleanup (template does).
- Returns a `DisplayDimension`; `dd.GetDimension2(0)` → `Dimension`: `.SystemValue = 0.07` (metres),
  `.FullName` (*measured* `D1@Szkic1@Część11.Part`), `.Name = "Dlugosc"` to give it a stable name.
- Parametric parts: create global variables and link dimensions:
  `model.GetEquationMgr.Add2(-1, """L"" = 250mm", True)` and
  `eq.Add2(-1, """D1@Szkic1"" = ""L""", True)`; change later via the equation text + `EditRebuild3`.
  Equations use the localized feature names - take them from `Dimension.FullName`.
- If geometry is created at exact size and nobody edits it, dimensions are optional; for a part the
  designer will modify, add driving dimensions with readable names.

## 5. Features (look up the signature first)

Feature methods take 15-30 positional arguments; revisions differ in count and order. Before writing
one, print it from the installed type library:
```
python scripts/api_signature.py FeatureManager.FeatureExtrusion3 FeatureManager.FeatureCut4
python scripts/api_signature.py "FeatureManager.HoleWizard*" --find SheetMetalBaseFlange
```
Verified calls (SW 2024):
```vb
' Boss-extrude blind 10 mm along the sketch normal (23 args)
Set feat = fm.FeatureExtrusion3(True, False, False, swEndCondBlind, swEndCondBlind, 0.01, 0, _
    False, False, False, False, 0, 0, False, False, False, False, True, True, True, swStartSketchPlane, 0, False)
' Cut through all (27 args)
Set feat = fm.FeatureCut4(True, False, False, swEndCondThroughAll, swEndCondBlind, 0, 0, _
    False, False, False, False, 0, 0, False, False, False, False, False, True, True, False, False, False, _
    swStartSketchPlane, 0, False, False)
' Constant-radius fillet on the selected edges (14 args)
Set feat = fm.FeatureFillet3(swFeatureFilletUniformRadius, 0.005, 0, 0, swFeatureFilletType_Simple, 0, 0, _
    Empty, Empty, Empty, Empty, Empty, Empty, Empty)
```
- End conditions: `swEndCondBlind 0, ThroughAll 1, UpToNext 2, MidPlane 6, ThroughAllBoth 9`.
  Direction: `Dir` (3rd arg) flips; `Sd` True = single direction.
- Other current revisions (check with `api_signature.py`): `FeatureRevolve2` (20 args, needs a
  centerline in the sketch), `FeatureLinearPattern5` (22), `HoleWizard5` (27), `SimpleHole2` (23),
  `FeatureExtrusionThin2` (30), `InsertSheetMetalBaseFlange2` (19), `InsertSheetMetalEdgeFlange2` (13).
- Simple through holes: a sketch of circles + `FeatureCut4` is the most robust. Hole Wizard gives
  standard holes (threads, counterbores) but its argument list is long and size strings are standard-
  and language-specific - test it on the target installation.

## 6. Selecting faces and edges

- Find geometry instead of picking by coordinates: `bodies = part.GetBodies2(swSolidBody, True)` →
  `bodies(0).GetFaces` / `.GetEdges`.
  - Planar face by normal: `face.Normal` → (nx,ny,nz) (zero vector for non-planar).
  - Straight edge by direction: `edge.GetCurve.IsLine` and `.LineParams` (root xyz, direction xyz).
  - Cylindrical faces: `face.GetSurface.IsCylinder`, `.CylinderParams` (origin, axis, radius).
- Select through `Entity`: `Dim ent As SldWorks.Entity: Set ent = face: ent.Select4 append, Nothing`.
  (*measured*: calling `Select4` on a `Face2`/`Edge` variable also works at run time, but `check_api.py`
  flags it - Entity is the declared owner.)
- Don't index a method result inline: `part.GetBodies2(...)(0)` raised error 13 in VBA (*measured*) -
  assign the array to a Variant first.
- Edge/face pointers become invalid after the next feature - re-query them after each feature.

## 7. Materials

- `part.SetMaterialPropertyName2 configName, database, materialName` - *measured*:
  - database is the .sldmat file name, case-insensitive, with or without `.sldmat` (`"SOLIDWORKS Materials"` or a company `.sldmat`);
  - **an unknown name is silently ignored** (the previous material stays) - always read back with
    `part.GetMaterialPropertyName2(configName, dbOut)` and stop if it differs;
  - names are those inside the database: SOLIDWORKS DB uses English names (`AISI 304`, `AISI 1020`)
    even in the Polish UI; a company database may use local grades instead (e.g. `S235JR`, `S355`,
    `1.4301 (AISI 304)`) - list them before hard-coding a name;
  - **mass in a new, unsaved part: use `Extension.CreateMassProperty2`.** *Measured* (3 separate runs):
    `CreateMassProperty` kept density 1 kg/m³ (mass ≈ 0) after the material change, while
    `CreateMassProperty2` returned 7800 immediately; after saving both agree. Still sanity-check
    `Density > 100`. For a "Masa" property prefer a link (`"SW-Mass@..."`) over a fixed number.
- List databases: `swApp.GetMaterialDatabases`; .sldmat files are **UTF-16** XML
  (`<material name="...">` inside `<classification name="...">`) - decode accordingly when reading them.
- Use the active configuration name (`model.ConfigurationManager.ActiveConfiguration.Name`) or `""`.

## 8. Sheet metal

- Base flange from a sketch: closed profile → flat plate; open profile (lines) → bent part.
  `fm.InsertSheetMetalBaseFlange2(thickness, thickenReverse, bendRadius, extrudeDist1, extrudeDist2,
  flip, endCond1, endCond2, dirToUse, Nothing, useDefaultRelief, reliefType, reliefWidth, reliefDepth,
  reliefRatio, useReliefRatio, merge, useFeatScope, useAutoSelect)` - check with `api_signature.py`.
  For a closed profile the extrude distances are ignored; for an open profile they are the flange length.
- *measured*: an open U profile drawn on the OUTER faces (100 x 40) with `ThickenDir = False` gave a
  104 x 42 part (material added outside); `ThickenDir = True` gave 100 x 40. The side depends on the
  drawing direction of the lines - check the overall size after the feature and flip if needed; never
  trust it blindly. The bend radius argument is the inner radius; the K-factor comes from the template
  (0.5 here: 2 mm, R2, 100/40/40 → flat 173.42 x 80 mm).
- It creates `SheetMetal`, `BaseFlange`, `FlatPattern` features and a cut-list; thickness/radius are then
  in the SheetMetal feature definition.
- Edge flanges need selected edges and a flange profile - more fragile; for simple U/Z profiles prefer an
  open-profile base flange.
- Verify: `Body2.IsSheetMetal`, the bends (sub-features `UiBend`/`SketchBend` of `FlatPattern`), the
  thickness (`SheetMetal` feature `GetDefinition.Thickness`) and the overall size.
- *measured*: for this macro-built part NO DXF export wrote the bend lines (see pitfalls.md #3) - if the
  user needs bend lines in DXF, check the exported file and say so.

## 9. Configurations / families

- *measured*: a `Dimension` object obtained BEFORE `ShowConfiguration2` writes `SetSystemValue3(...,
  swSetValue_InThisConfiguration ...)` into the configuration that was active when it was obtained -
  re-get it (`model.Parameter("DLUGOSC@RURA")`) after every configuration switch.
- *measured*: after a per-configuration value change `EditRebuild3` once left the old geometry;
  `ForceRebuild3 False` was reliable - use it, then check the volume/size of that configuration.

```vb
Set cfg = model.ConfigurationManager.AddConfiguration2(name, "", "", swConfigOption_DontActivate, "", "", False)
model.ShowConfiguration2 name
dimension.SetSystemValue3 0.45, swSetValue_InThisConfiguration, Empty
model.EditRebuild3
```
- *measured*: `AddConfiguration2(name, "", "", 0, "", "", True)` then `ShowConfiguration2` +
  `SetSystemValue3(value, 1 = this configuration, Nothing)` gave 100/200/350 mm configurations.
- Per-configuration properties: `model.Extension.CustomPropertyManager(cfgName).Add3 ...`.
- Restore the originally active configuration at the end.
- Many variants of a simple part are often better as one part with configurations (or a design table)
  than as many files - ask what the user's PDM/BOM process expects.

## 10. Checking the result

After `ForceRebuild3 False`:
- `model.Extension.CreateMassProperty` → `Volume` (m³), `Mass` (kg), `Density`; compare the volume with
  the value expected from the inputs (the template does this) - catches wrong units and failed features.
- Exact size: `Body2.GetExtremePoint` in ±X/±Y/±Z (`GetPartBox` is approximate).
- Feature errors: walk features, `feat.GetErrorCode2(isWarning)` <> 0 → report.
- Then `SaveAs` (`swSaveAsOptions_Silent`), check the return value with `= False` and the error code.
- Live testing: the 3.0 runner (`references/runner-safety.md`) reports any document the macro creates
  as `UNDECLARED_DOCUMENT`, so `create_part_template.bas` (which calls `NewDocument`) cannot PASS under it
  yet. Its live evidence is the 2.3.0 run on SW 2024 SP5; for runner-based checks, build geometry in the
  opened blank fixture copy as the `recipes/modeling/` recipes do, and assert with `--expect-result`.
  A part that fails before SaveAs stays open UNSAVED (no path) - the runner can't close it; close it by
  its title, never other unsaved documents (they may be the user's).
