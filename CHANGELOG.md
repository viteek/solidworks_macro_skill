# Changelog

## 2.2.0 - 2026-10-30 (merge of the local 2.0 line with Codex 2.1.0, after audit)

From 2.1.0 (reviewed and fixed):
- `safe_export_template.bas`: active-document exporter with stage -> check -> publish and backup.
  Fixed: bend lines use mask 1+4+8 (4 alone gives none on SW 2024 SP5 - measured), cleanup via
  `Resume` (On Error Resume Next is ignored inside an active error handler).
- `batch_folder_template.bas`: audit mode, read-only open, same-name / returned-path guards, CSV log per
  file. Fixed: `Resume` in three handlers, 8.3 short-path comparison, folder picker when interactive.
- `source_metrics_template.bas` (visible bodies only - hidden ones aren't exported to STEP).
- Geometry validation: `validate_step.py` (OpenCascade), `validate_dxf.py`, `validated_export.py`,
  `validation_reports.py`, `profiles/`. Fixed: DXF roles by LINETYPE then layer (SW puts bend lines on
  layer 0 as CENTERX2), ignored POINT entities, realistic tolerances, outside-material = warning,
  STEP validator no longer needs the DXF stack, sibling imports under `python -I`, pipeline uses this
  package's runner and removes its job folder after publication.
- Rules: no deleting outputs before success, read before replacing properties (linked values kept),
  no implicit PDM operations, static checks are not a compiler, corrected ExportToDWG2 parameter names,
  no CloseAllDocuments as recovery, cp1250 conversion refuses unrepresentable characters.

Kept from 2.0 (removed or contradicted in 2.1): trigger description, output format, debugging section,
measured SW 2024 pitfalls list, export recipes (DXF from assembly components, PDF per sheet, STEP AP...),
runner dialog answering / `--pick-folder` / compile-error hint.

Runner: `--close-opened` closes only new documents inside the test root and re-activates the user's
document; new `--test-root`, `--strict-dialogs`, `--expect-file`, `--expect-text`; refuses when a
same-named file from another folder is loaded.

Other: Polish feature name corrected (`Dodanie-wyciągnięcie1`); `SetProp` recipe updates values but
keeps linked expressions; flat-pattern recipe stages instead of `Kill`.

## 2.0 (local) - 2026-10-30
Measured SW 2024 pitfalls, type-library API checker, live runner, templates; evaluated on 6 tasks.
