# Tools around macros & ready-made macros

Read this when the user wants to run a macro on many files, put macros on buttons, pass arguments
to a macro, or when the task is a common one that may already exist as a free, tested macro.

## Contents
1. Check for a ready-made macro first
2. CAD+ toolset (free, open source)
3. Writing macros that work in Batch+ / Toolbar+ / Task Scheduler
4. Catalogue of CodeStack macros by task

---

## 1. Check for a ready-made macro first

CodeStack (https://www.codestack.net/solidworks-tools/) publishes ~120 MIT-licensed, maintained macros
for everyday tasks. If one matches the request closely, say so and link it - the user gets a tested
tool, and you can still adapt it or write a custom one if they need something different (e.g. Polish
property names, company naming rules, a combined workflow). Don't pretend a ready macro doesn't exist.

## 2. CAD+ toolset (free, open source - https://cadplus.xarial.com/)

| Tool | What it does | Suggest it when |
|---|---|---|
| **Batch+** (cadplus.xarial.com/batch/) | Opens files/folders in SOLIDWORKS and runs one or more macros on each file, optional save, background mode, restarts SW on hangs, job log. | "Run this on all files in a folder", hundreds of files, overnight jobs. Then the macro only handles the **active document** - no folder loop, no open/close code. |
| **Toolbar+** (cadplus.xarial.com/toolbar/) | Macros as buttons in custom toolbars/menus, shortcuts, icons, per-button arguments, toggle buttons, shared config for a team. Successor of MyToolbar. | Deploying a macro library to a team, one macro behind several buttons. |
| **Macro+** (cadplus.xarial.com/macro-plus/) | Typed macro operations: input model, arguments, result files, issue reporting. | Macros that must report per-file results back to Batch+. |
| **xPort / eXport+** | Export to PDF / eDrawings / images via eDrawings - no SOLIDWORKS licence needed. | Mass export where SOLIDWORKS isn't required. |
| SOLIDWORKS **Task Scheduler** | Native batch tool (with the SW licence): update files, print, export, run custom tasks. | Company forbids extra add-ins. |
| Native **macro buttons** | Tools > Customize > Commands > Macro - a button with a 16x16/20x20 bmp icon. | One macro, no add-in. |

## 3. Writing macros that work in Batch+ / Toolbar+ / Task Scheduler

- Entry point: one parameterless `Sub main()` in a standard module. Launchers call it via `RunMacro2`.
- Batch+ makes each file the active document: use `swApp.ActiveDoc`, check `Is Nothing` and `GetType`.
  **Don't close the document** and preferably don't save it - let Batch+ do it (its "save" option).
- No modal dialogs on the batch path. Use a `Const SHOW_MESSAGES As Boolean` (or detect arguments mode)
  and report through a log file (appended per file) / status bar (`swApp.Frame.SetStatusBarText`).
  **Re-raising errors is not dialog-free:** a macro that ends with an unhandled `Err.Raise` shows the VBA
  run-time error dialog when started normally (measured), which would stop an unattended night run.
  CodeStack re-raises only in arguments mode, when Batch+ / Toolbar+ started the macro through its
  runner and catches the error. Default to: handle the error, write it to the log, exit cleanly; offer a
  `RAISE_ERRORS` switch (off by default) for users who verified that their Batch+ setup reports raised
  errors without a dialog.
- Settings as `Const` at the top. Optional arguments from Toolbar+/Batch+ behind a compile switch:
  ```vb
  #Const ARGS = False   ' True: read arguments passed by Toolbar+ / Batch+
  ...
  #If ARGS Then
      Dim runner As Object, param As Object, vArgs As Variant
      Set runner = CreateObject("CadPlus.MacroRunner.Sw")
      Set param = runner.PopParameter(swApp)
      vArgs = param.Get("Args")          ' array of strings, e.g. "-pdf", "-out", "D:\Export"
  #End If
  ```
  Newer Macro+ API: `CreateObject("CadPlus.MacroOperationManager").PopOperation(swApp)` → `.Arguments`
  (each item `.GetValue()`), `.Model`, `.SetResultFiles`, `.ReportIssue msg, type`. Wrap in error handling
  and fall back to the constants when the macro is started manually.
- Batch-friendly output: create output folders automatically, append to logs (don't overwrite),
  `SKIP_EXISTING_FILES` option so a crashed run can resume, resolve relative output paths against the
  model's folder, name templates such as `<_FileName_>`, `<_ConfName_>`, `<Property>`; fail clearly on
  unsaved models (`GetPathName = ""`).
- Remove settings that pop dialogs in batch (template prompt, 3D Interconnect for imports).
- Toolbar+ toggle buttons evaluate a VB.NET expression for the pressed state, e.g.
  `Return CType(Application, Object).Sw.GetUserPreferenceToggle(64)`.

## 4. Catalogue of CodeStack macros by task (base https://www.codestack.net/)

**Export**
- Multiple formats, name templates, all configs/sheets: solidworks-api/import-export/export-multi-formats/
- Flat patterns of a part / all assembly components to DXF/DWG: solidworks-api/document/sheet-metal/export-all-flat-patterns/
- Each sheet to its own PDF: solidworks-api/import-export/export-sheets-pdf/
- Flat-pattern drawing views to DXF: solidworks-api/document/drawing/export-sheet-metal-views/
- STL of a component without opening it: solidworks-api/import-export/export-stl/
- Bodies / cut-lists to separate files: solidworks-api/import-export/export-individual-bodies/ , .../export-individual-cut-lists/
- Import STEP files (Batch+ ready): solidworks-api/import-export/import-step/

**Properties, cut-lists, BOM**
- Write BOM quantities into a property: solidworks-api/document/assembly/components/write-quantities/
- Auto date property: solidworks-api/data-storage/custom-properties/auto-date/
- Copy file properties to configurations: solidworks-api/data-storage/custom-properties/copy-file-specific-to-configuration/
- Link properties to a CSV file: solidworks-api/data-storage/custom-properties/link-to-file/
- Sort properties: solidworks-api/data-storage/custom-properties/sort/
- Cut-list properties to the model / rename cut-list items / colorize: solidworks-api/document/cut-lists/
- Material database property to model: solidworks-api/document/materials/copy-custom-property/

**Assembly**
- Where used, replace preserving selection, rename, component reference: solidworks-api/document/assembly/components/
- Export component positions to CSV; open components in position; purge configurations.
- Break/lock external references (good example of ARGS handling): solidworks-api/document/manage-external-references/

**Drawings & tables**
- Tables (BOM, revision, cut-list) to CSV: solidworks-api/document/tables/export-table-csv/
- Replace sheet formats by a map: solidworks-api/document/drawing/replace-sheet-format/
- Rename sheets from properties; open referenced model / associated drawing; scale views by size:
  solidworks-api/document/drawing/
- Export dimensions to CSV; find/delete notes; increment revision note: solidworks-api/document/notes/

**Model & productivity**
- Rebuild + zoom + hide types (args): solidworks-api/document/rebuild/
- Delete broken equations: solidworks-api/document/delete-model-equations/
- Rename features sequentially / by type: solidworks-api/document/features-manager/
- Save all silently; close all except active; run a group of macros: solidworks-api/application/
- Run code on document open/save/new (events): solidworks-api/application/documents/
- Set image quality (Batch+): solidworks-api/options/document/set-image-quality/

Full list: https://www.codestack.net/solidworks-tools/
