"""Materialize the reviewed benchmark contract. Does not generate CAD binaries."""
import json
from pathlib import Path


def eq(key, expected):
    return {'id': key.replace('/', '_'), 'path': '/' + key, 'op': 'equals', 'expected': expected}


def num(key, expected, tol=1e-8):
    return {'id': key.replace('/', '_'), 'path': '/' + key, 'op': 'number', 'expected': expected, 'abs_tol': tol}


def items(key, expected):
    return {'id': key, 'path': '/' + key, 'op': 'unordered_equals', 'expected': expected}


def file(name, contains):
    return {'id': name, 'op': 'file', 'file': 'outputs/' + name, 'min_bytes': 20, 'contains': contains}


TASKS = [
 ('docs_properties', 'Read document properties', 'properties-part',
  'Create a 100 x 50 x 10 mm rectangular boss named Base in Default configuration; set document text PartNo=P-001 and Description=Plate. Save baseline part.sldprt, close, reopen saved copy.',
  'Read document PartNo and Description without editing or saving.',
  'ICustomPropertyManager.Get6; IModelDoc2.GetType; IModelDoc2.GetSaveFlag',
  [eq('document_type', 'PartDoc'), eq('properties/document/PartNo', 'P-001'), eq('properties/document/Description', 'Plate'), eq('dirty', False)]),
 ('docs_properties', 'Configuration property precedence', 'config-part',
  'Use properties-part recipe; add configurations Default and Long; document PartNo=P-001; Long configuration PartNo=P-002; Default has no configuration PartNo. Activate Default and save baseline.',
  'Resolve PartNo with configuration precedence and document fallback for both configurations; restore Default.',
  'IModelDocExtension.CustomPropertyManager; ICustomPropertyManager.Get6; IModelDoc2.ShowConfiguration2',
  [eq('resolved/Default', 'P-001'), eq('resolved/Long', 'P-002'), eq('active_configuration', 'Default')]),
 ('docs_properties', 'Write one scoped property', 'config-part',
  'Use config-part recipe above with Long PartNo=P-002 and no Revision property in any scope; duplicate baseline for this run.',
  'Set Long configuration text Revision=B; leave document and Default scopes without Revision. Do not save source.',
  'ICustomPropertyManager.Add3; ICustomPropertyManager.Get6',
  [eq('properties/Long/Revision', 'B'), eq('property_exists/document/Revision', False), eq('property_exists/Default/Revision', False), eq('properties/Long/PartNo', 'P-002')]),
 ('docs_properties', 'Enumerate open document types', 'three-documents',
  'Create and save a Base part from properties-part, an empty assembly and an A4 landscape drawing from configured templates. Open all three and no other documents; activate part.',
  'Enumerate exactly these three documents and retain active part.',
  'ISldWorks.GetFirstDocument; IModelDoc2.GetNext; IModelDoc2.GetType',
  [eq('open_document_count', 3), items('open_document_types', ['PartDoc', 'AssemblyDoc', 'DrawingDoc']), eq('active_document', 'part.sldprt')]),
 ('selection_modeling', 'Select a named feature', 'base-part',
  'Create 100 x 50 x 10 mm centered rectangular boss named Base on Front Plane; save baseline. Clear selections and ensure no sketch is active.',
  'Select only Base feature with mark 0 and report selected object.',
  'IModelDocExtension.SelectByID2; ISelectionMgr.GetSelectedObject6; IFeature.Name',
  [eq('selection/count', 1), eq('selection/mark', 0), eq('selection/name', 'Base'), eq('selection/object_type', 'Feature')]),
 ('selection_modeling', 'Create a dimensioned boss', 'blank-part',
  'Create empty part from recorded metric template, named blank.sldprt; no user sketches or bodies. Front Plane is standard XY plane; note localized plane name in fixture manifest.',
  'Create centered 100 x 50 mm rectangle and 10 mm blind boss in positive Z; name feature Base.',
  'ISketchManager.CreateCenterRectangle; IFeatureManager.FeatureExtrusion2; IPartDoc.GetBodies2',
  [eq('solid_body_count', 1), eq('feature_exists/Base', True), num('volume_m3', .00005, 1e-10), num('bbox/width_m', .1, 1e-6), num('bbox/height_m', .05, 1e-6), num('bbox/depth_m', .01, 1e-6)]),
 ('selection_modeling', 'Create centered through hole', 'base-part',
  'Use base-part recipe with planar top face at Z=10 mm; no other features. Save baseline and clear selection.',
  'Cut a centered diameter 10 mm through hole normal to top face; name cut CenterHole.',
  'ISketchManager.CreateCircleByRadius; IFeatureManager.FeatureCut4; IFeature.GetTypeName2',
  [eq('solid_body_count', 1), eq('feature_exists/CenterHole', True), num('hole_diameter_m', .01), num('volume_m3', .00004921460183660255, 1e-10)]),
 ('selection_modeling', 'Edit existing dimension and rebuild', 'dimension-part',
  'Use base-part with boss depth dimension explicitly named Depth in feature Base (Depth@Base=10 mm). Save baseline; record persistent feature ID.',
  'Change Depth@Base to 20 mm, rebuild, preserve Base feature identity and one body.',
  'IModelDoc2.Parameter; IDimension.SystemValue; IModelDoc2.EditRebuild3',
  [num('dimensions/Depth@Base', .02), num('volume_m3', .0001, 1e-10), eq('solid_body_count', 1), eq('base_identity_preserved', True), eq('rebuild_errors', 0)]),
 ('assemblies', 'Traverse resolved component instances', 'assembly-tree',
  'Create Base part from base-part recipe and distinct Bolt.sldprt as a diameter 10 mm, length 20 mm cylindrical boss in Default. Build sub.sldasm containing Bolt-1 and Bolt-2 from Bolt.sldprt. Top.sldasm contains fixed Base-1, sub-1 and suppressed Spare-1 referencing Base part. Fully resolve unsuppressed components. Save all.',
  'Traverse recursively and count instances, suppressed components and distinct unsuppressed part file/config pairs; include subassembly node.',
  'IAssemblyDoc.GetComponents; IComponent2.GetChildren; IComponent2.IsSuppressed; IComponent2.ReferencedConfiguration',
  [eq('instance_count', 5), eq('suppressed_count', 1), eq('distinct_unsuppressed_parts', 2), items('component_names', ['Base-1', 'sub-1', 'sub-1/Bolt-1', 'sub-1/Bolt-2', 'Spare-1'])]),
 ('assemblies', 'Insert component with transform', 'empty-assembly',
  'Save empty top.sldasm and 100 x 50 x 10 mm part insert.sldprt. No components in top. Record absolute fixture paths; set identity part transform.',
  'Insert one resolved insert.sldprt instance at translation (0.1,0.2,0.3) m with identity rotation.',
  'IAssemblyDoc.AddComponent5; IComponent2.Transform2; IMathTransform.ArrayData',
  [eq('instance_count', 1), num('transform/tx_m', .1), num('transform/ty_m', .2), num('transform/tz_m', .3), eq('rotation_rows', [[1,0,0],[0,1,0],[0,0,1]])]),
 ('assemblies', 'Add one coincident mate', 'two-blocks',
  'Create assembly with fixed BlockA-1 at origin and floating BlockB-1 translated +30 mm Z; same centered base-part geometry. No mates. Name A top planar face MateA and B bottom face MateB via entity naming. Save baseline.',
  'Add aligned coincident mate between named faces, name it Contact, rebuild; report mate objects and face separation.',
  'IAssemblyDoc.AddMate5; IMate2.Type; IMate2.MateEntity; IModelDoc2.EditRebuild3',
  [eq('mate_count', 1), eq('mates/Contact/type', 'coincident'), eq('mates/Contact/entity_count', 2), num('contact_separation_m', 0, 1e-7), eq('rebuild_errors', 0)]),
 ('assemblies', 'Resolve lightweight without unsuppressing', 'lightweight-assembly',
  'Use assembly-tree recipe. Reopen Top with lightweight loading; record Bolt-1 lightweight=True, Spare-1 suppressed=True. If installed settings cannot produce lightweight state, mark fixture NOT_TESTED.',
  'Resolve all lightweight components needed for traversal; leave Spare-1 suppressed and active assembly unchanged.',
  'IAssemblyDoc.ResolveAllLightWeightComponents; IComponent2.GetSuppression; IComponent2.GetModelDoc2',
  [eq('lightweight_before', True), eq('unresolved_after', 0), eq('spare_suppressed', True), eq('active_document', 'Top.sldasm')]),
 ('drawings', 'Enumerate sheets and actual views', 'two-sheet-drawing',
  'Create A4 landscape drawing: Sheet1 has Front and Top model views of base-part, Sheet2 has Right model view. Rename views Front, Top, Right. No sheet-format model views; activate Sheet1 and save.',
  'Enumerate both sheets and three model views, excluding per-sheet container view; restore Sheet1.',
  'IDrawingDoc.GetSheetNames; IDrawingDoc.GetFirstView; IView.GetNextView',
  [items('sheets', ['Sheet1', 'Sheet2']), items('model_views', ['Sheet1/Front','Sheet1/Top','Sheet2/Right']), eq('model_view_count', 3), eq('active_sheet', 'Sheet1')]),
 ('drawings', 'Change one note text', 'noted-drawing',
  'Use two-sheet-drawing recipe. Add free note named ReleaseNote with text DRAFT on Sheet1 and KeepNote with text KEEP on Sheet2. Save baseline.',
  'Change ReleaseNote to RELEASED while keeping KeepNote and all three model views.',
  'INote.GetText; INote.SetText; IAnnotation.GetName',
  [eq('notes/ReleaseNote', 'RELEASED'), eq('notes/KeepNote', 'KEEP'), eq('model_view_count', 3)]),
 ('drawings', 'Export all sheets to PDF', 'two-sheet-drawing',
  'Use two-sheet-drawing recipe; output directory initially empty and writable. Hash source drawing and referenced part before run. Pin PDF export settings in manifest.',
  'Export Sheet1 and Sheet2 to outputs/all-sheets.pdf; inspect exported PDF independently for two pages and sheet identifiers.',
  'IExportPdfData.SetSheets; IModelDocExtension.SaveAs',
  [file('all-sheets.pdf', ['%PDF-']), eq('pdf/page_count', 2), items('pdf/sheet_ids', ['Sheet1','Sheet2']), eq('export/error_code', 0)]),
 ('drawings', 'Rename a sheet and restore active sheet', 'two-sheet-drawing',
  'Use two-sheet-drawing recipe with Sheet1 active and Front/Top views on Sheet1, Right view on Sheet2. Save baseline.',
  'Rename Sheet2 to Detail; retain Sheet1 active and preserve all model view references.',
  'IDrawingDoc.ActivateSheet; ISheet.SetName; IView.ReferencedDocument',
  [items('sheets', ['Sheet1', 'Detail']), eq('active_sheet', 'Sheet1'), items('model_views', ['Sheet1/Front','Sheet1/Top','Detail/Right']), eq('broken_view_references', 0)]),
 ('sheetmetal_export', 'Export rectangular flat blank', 'flat-sheet',
  'Create centered 100 x 50 mm sheet metal base flange, thickness 2 mm, bend radius 2 mm, K-factor 0.5, no bends. Name Sheet-Metal and Flat-Pattern deterministically; save folded Default. Output folder empty.',
  'Export geometry-only flat DXF outputs/blank.dxf; independently parse model-space geometry dimensions in mm, units and closed outer contour.',
  'IPartDoc.ExportToDWG2; IBody2.IsSheetMetal',
  [file('blank.dxf', ['SECTION','ENTITIES','EOF']), num('dxf/width_mm', 100, .01), num('dxf/height_mm', 50, .01), eq('dxf/closed_outer_contours', 1), eq('dxf/bend_line_count', 0), eq('dxf/units', 'mm')]),
 ('sheetmetal_export', 'Export one bend with bend lines', 'bent-sheet',
  'Create same flat-sheet blank then 90 degree edge flange along the 100 mm edge, flange length 20 mm with outer virtual sharp reference, radius 2 mm, thickness 2 mm, K=.5. Save Default folded; record one bend baseline.',
  'Export geometry+bend lines to outputs/bent.dxf; verify exactly one straight bend line and unchanged folded configuration.',
  'IPartDoc.ExportToDWG2; IFeature.IsSuppressed2',
  [file('bent.dxf', ['SECTION','ENTITIES','EOF']), eq('dxf/bend_line_count', 1), eq('dxf/closed_outer_contours', 1), eq('active_configuration', 'Default'), eq('flat_pattern_suppressed', True)]),
 ('sheetmetal_export', 'Export distinct bodies without collision', 'two-sheet-bodies',
  'Create multibody sheet metal part with two unmerged base flanges: 100x50x2 mm at origin and 60x40x2 mm translated 200 mm X. Separate cut-list items tagged CutId=A and CutId=B; update cut list, save Default. Output empty.',
  'Export outputs/A.dxf and outputs/B.dxf exactly once each, preserving two solid bodies and original configuration.',
  'IFeature.GetTypeName2; IBodyFolder.GetBodies; IPartDoc.ExportToDWG2',
  [file('A.dxf', ['ENTITIES','EOF']), file('B.dxf', ['ENTITIES','EOF']), items('exported_files', ['A.dxf','B.dxf']), num('A/width_mm', 100, .01), num('B/width_mm', 60, .01), eq('solid_body_count', 2), eq('active_configuration', 'Default')]),
 ('sheetmetal_export', 'Export STEP and inspect reimport', 'base-part',
  'Use base-part recipe. Output empty; pin AP214 export settings in manifest and record original source hash. Configure importer to SI units.',
  'Export outputs/base.step and independently reimport into temporary document without saving source; measure one solid and volume.',
  'IModelDocExtension.SaveAs; ISldWorks.LoadFile4; IPartDoc.GetBodies2',
  [file('base.step', ['ISO-10303-21;','END-ISO-10303-21;']), eq('reimport/solid_body_count', 1), num('reimport/volume_m3', .00005, 1e-9), eq('export/error_code', 0)]),
 ('errors_resilience', 'Handle missing active document', 'no-active-document',
  'Start isolated SOLIDWORKS session with zero open documents. Capture session inventory; use a JSON fixture manifest recording empty document set instead of a CAD binary.',
  'Attempt active-part operation; return controlled NO_ACTIVE_DOCUMENT diagnostic and make no files or document changes.',
  'ISldWorks.ActiveDoc; ISldWorks.GetDocumentCount',
  [eq('outcome', 'NO_ACTIVE_DOCUMENT'), eq('handled', True), eq('open_document_count_after', 0), eq('files_created', []), eq('unhandled_errors', 0)]),
 ('errors_resilience', 'Reject wrong document type', 'drawing-for-part-operation',
  'Open only two-sheet-drawing as active document. Record before hashes, sheet/view inventory and dirty flag=False.',
  'Request part-only through-hole operation; return WRONG_DOCUMENT_TYPE without creating features or files.',
  'IModelDoc2.GetType',
  [eq('outcome', 'WRONG_DOCUMENT_TYPE'), eq('handled', True), eq('document_type', 'DrawingDoc'), eq('model_view_count', 3), eq('files_created', []), eq('dirty', False)]),
 ('errors_resilience', 'Reject invalid output path and restore state', 'base-part-invalid-output',
  'Use saved base-part and configure output path containing forbidden filename character bad:name.step beneath sandbox. Preselect Base with mark 0. Snapshot preferences and source hashes.',
  'Reject invalid filename before SaveAs with INVALID_OUTPUT_PATH; preserve Base selection and export preferences, create no files.',
  'IModelDocExtension.SaveAs; ISelectionMgr.GetSelectedObject6; ISldWorks.GetUserPreferenceIntegerValue',
  [eq('outcome', 'INVALID_OUTPUT_PATH'), eq('handled', True), eq('files_created', []), eq('selection/name', 'Base'), eq('preferences_restored', True), eq('save_calls', 0)]),
 ('errors_resilience', 'Recover from missing selection', 'base-part',
  'Use base-part; clear every selection; record zero selection count and Base persistent identity. Set export preferences to recorded baseline and save source before run.',
  'Request face-only operation; return SELECTION_REQUIRED, restore global preferences and exit without model edits or save.',
  'ISelectionMgr.GetSelectedObjectCount2; ISelectionMgr.GetSelectedObject6',
  [eq('outcome', 'SELECTION_REQUIRED'), eq('handled', True), eq('selection/count', 0), eq('base_identity_preserved', True), eq('preferences_restored', True), eq('dirty', False)])
]


def main():
    target = Path(__file__).parent / 'tasks'
    target.mkdir(exist_ok=True)
    for index, (family, title, fixture, recipe, request, api, assertions) in enumerate(TASKS, 1):
        task = {'schema_version': 1, 'task_id': f'T{index:02}', 'family': family, 'title': title,
                'fixture': {'id': fixture, 'revision': '1'}, 'fixture_instructions': recipe,
                'request': request, 'api_candidates': api.split('; '),
                'evaluation': 'Independent post-run collector supplies each JSON pointer from reopened/traversed CAD objects or independently parsed exports. File assertions read actual artifacts. See ../README.md for six-phase evidence and safety contract.',
                'assertions': assertions + [eq('safety/unintended_saves', 0), eq('safety/source_hashes_unchanged', True)],
                'live_evidence_status': 'NOT_TESTED'}
        (target / (task['task_id'] + '.json')).write_text(json.dumps(task, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
