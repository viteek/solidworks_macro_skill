#!/usr/bin/env python3
"""Prepare/run isolated assembly and drawing smoke cases through the reviewed runner.

Requires private saved blank/block part fixtures. No UI automation or dialog answers.
Every live mutation, including fixture generation and observation saves, runs as
reviewed VBA through run_in_solidworks.py. Defaults to preparation only.
"""
import argparse
import json
from pathlib import Path
import shutil
import subprocess
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))
from runner_safety import code_hash, file_hash
from run_in_solidworks import prepare_code

ROOT = Path(__file__).resolve().parents[1]


def literal(path):
    return '"' + str(path).replace('"', '""') + '"'


def wrapper(recipe, setup):
    source = recipe.read_text(encoding='utf-8')
    source = source.replace('Public Sub main()', 'Private Sub RunRecipe()', 1)
    return source + '\nPublic Sub main()\n' + setup + '\n    RunRecipe\nEnd Sub\n'


def selection_setup(assembly=False):
    if not assembly:
        return '''    Dim fixtureModel As SldWorks.ModelDoc2
    Dim fixturePart As SldWorks.PartDoc
    Dim fixtureBodies As Variant, fixtureFaces As Variant
    Dim fixtureBody As SldWorks.Body2
    Dim fixtureEntity As SldWorks.Entity
    Dim n As Long
    Set fixtureModel = Application.SldWorks.ActiveDoc
    Set fixturePart = fixtureModel
    fixtureBodies = fixturePart.GetBodies2(swSolidBody, False)
    Set fixtureBody = fixtureBodies(0)
    fixtureFaces = fixtureBody.GetFaces
    fixtureModel.ClearSelection2 True
    For n = 0 To 2
        Set fixtureEntity = fixtureFaces(n)
        If fixtureEntity.Select4(True, Nothing) = False Then Err.Raise vbObjectError + 500, , "Fixture selection failed"
    Next n'''
    return '''    Dim fixtureModel As SldWorks.ModelDoc2
    Dim fixtureAssembly As SldWorks.AssemblyDoc
    Dim fixtureComponents As Variant, fixtureBodies As Variant, fixtureFaces As Variant
    Dim fixtureComponent As SldWorks.Component2
    Dim fixturePart As SldWorks.PartDoc
    Dim fixtureBody As SldWorks.Body2
    Dim fixtureEntity As SldWorks.Entity
    Dim n As Long
    Set fixtureModel = Application.SldWorks.ActiveDoc
    Set fixtureAssembly = fixtureModel
    fixtureComponents = fixtureAssembly.GetComponents(True)
    fixtureModel.ClearSelection2 True
    For n = 0 To 1
        Set fixtureComponent = fixtureComponents(n)
        Set fixturePart = fixtureComponent.GetModelDoc2
        fixtureBodies = fixturePart.GetBodies2(swSolidBody, False)
        Set fixtureBody = fixtureBodies(0)
        fixtureFaces = fixtureBody.GetFaces
        Set fixtureEntity = fixtureComponent.GetCorrespondingEntity(fixtureFaces(0))
        If fixtureEntity.Select4(True, Nothing) = False Then Err.Raise vbObjectError + 501, , "Assembly fixture selection failed"
    Next n'''


def fixture_macro(folder):
    return f'''Option Explicit
Public Sub main()
    On Error GoTo Failed
    Dim app As SldWorks.SldWorks
    Dim original As SldWorks.ModelDoc2, source As SldWorks.ModelDoc2, created As SldWorks.ModelDoc2
    Dim assembly As SldWorks.AssemblyDoc
    Dim component As SldWorks.Component2
    Dim math As SldWorks.MathUtility, transform As SldWorks.MathTransform
    Dim values(0 To 15) As Double, matrixValues As Variant
    Dim errors As Long, warnings As Long, handle As Integer
    Dim result As String, originalName As String
    Set app = Application.SldWorks
    Set original = app.ActiveDoc
    originalName = original.GetTitle
    If Len(Dir$({literal(folder / 'fixture.sldasm')})) > 0 Then Err.Raise vbObjectError + 1, , "Assembly output exists"
    Set source = app.OpenDoc6({literal(folder / 'block.sldprt')}, swDocPART, swOpenDocOptions_Silent, "", errors, warnings)
    If source Is Nothing Then Err.Raise vbObjectError + 2, , "Source open failed"
    If errors <> 0 Then Err.Raise vbObjectError + 3, , "Source open errors"
    Set created = app.NewDocument({literal(folder / 'assembly.asmdot')}, 0, 0#, 0#)
    If created Is Nothing Then Err.Raise vbObjectError + 4, , "Assembly creation failed"
    Set assembly = created
    Set component = assembly.AddComponent5({literal(folder / 'block.sldprt')}, swAddComponentConfigOptions_CurrentSelectedConfig, "", False, "", 0#, 0#, 0#)
    If component Is Nothing Then Err.Raise vbObjectError + 5, , "First component failed"
    Set component = assembly.AddComponent5({literal(folder / 'block.sldprt')}, swAddComponentConfigOptions_CurrentSelectedConfig, "", False, "", 0#, 0#, 0#)
    If component Is Nothing Then Err.Raise vbObjectError + 6, , "Second component failed"
    Set math = app.GetMathUtility
    values(0) = 1#: values(4) = 1#: values(8) = 1#: values(9) = 0.1: values(10) = 0.05: values(12) = 1#
    matrixValues = values
    Set transform = math.CreateTransform(matrixValues)
    If component.SetTransformAndSolve3(transform, True) = False Then Err.Raise vbObjectError + 7, , "Fixture transform failed"
    If created.EditRebuild3 = False Then Err.Raise vbObjectError + 8, , "Fixture rebuild failed"
    If created.Extension.SaveAs({literal(folder / 'fixture.sldasm')}, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, errors, warnings) = False Then Err.Raise vbObjectError + 9, , "Fixture save failed"
    If errors <> 0 Then Err.Raise vbObjectError + 10, , "Fixture save errors"
    result = "RESULT fixture PASS; components=" & CStr(assembly.GetComponentCount(True))
    GoTo Cleanup
Failed:
    result = "RESULT fixture FAIL " & CStr(Err.Number) & ": " & Err.Description
    Resume Cleanup
Cleanup:
    On Error Resume Next
    If Not created Is Nothing Then app.CloseDoc created.GetTitle
    If Not source Is Nothing Then app.CloseDoc source.GetTitle
    Set original = app.ActivateDoc3(originalName, False, swDontRebuildActiveDoc, errors)
    If Err.Number <> 0 Then result = "RESULT fixture FAIL cleanup " & Err.Description
    handle = FreeFile
    Open {literal(folder / 'fixture.log')} For Output As #handle
    Print #handle, result
    Close #handle
End Sub
'''


def drawing_observer(folder):
    return f'''Option Explicit
Public Sub main()
    On Error GoTo Failed
    Dim app As SldWorks.SldWorks
    Dim original As SldWorks.ModelDoc2, model As SldWorks.ModelDoc2
    Dim drawing As SldWorks.DrawingDoc, view As SldWorks.View
    Dim documents As Variant, sheets As Variant, views As Variant
    Dim i As Long, j As Long, k As Long, count As Long, handle As Integer
    Dim errors As Long, warnings As Long, result As String
    Set app = Application.SldWorks
    Set original = app.ActiveDoc
    If StrComp(original.GetPathName, {literal(folder / 'block.sldprt')}, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 1, , "Input activation not restored"
    documents = app.GetDocuments
    For i = LBound(documents) To UBound(documents)
        Set model = documents(i)
        If model.GetType = swDocDRAWING Then
            If Len(model.GetPathName) <> 0 Then Err.Raise vbObjectError + 2, , "Drawing should be unsaved"
            Set drawing = model
            sheets = drawing.GetViews
            For j = LBound(sheets) To UBound(sheets)
                views = sheets(j)
                For k = LBound(views) + 1 To UBound(views)
                    Set view = views(k)
                    If StrComp(view.GetReferencedModelName, {literal(folder / 'block.sldprt')}, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 3, , "Observer reference mismatch"
                    count = count + 1
                Next k
            Next j
            If count < 3 Then Err.Raise vbObjectError + 4, , "Observer found fewer than 3 views"
            If Len(Dir$({literal(folder / 'fixture.slddrw')})) > 0 Then Err.Raise vbObjectError + 5, , "Drawing output exists"
            If model.Extension.SaveAs({literal(folder / 'fixture.slddrw')}, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, errors, warnings) = False Then Err.Raise vbObjectError + 6, , "Drawing fixture save failed"
            If errors <> 0 Then Err.Raise vbObjectError + 7, , "Drawing save errors"
            app.CloseDoc model.GetTitle
        End If
    Next i
    If count < 3 Then Err.Raise vbObjectError + 8, , "No created drawing found"
    result = "RESULT drawing_observer PASS; model_views=" & CStr(count) & "; input_restored=True"
    GoTo Cleanup
Failed:
    result = "RESULT drawing_observer FAIL " & CStr(Err.Number) & ": " & Err.Description
    Resume Cleanup
Cleanup:
    On Error Resume Next
    handle = FreeFile
    Open {literal(folder / 'drawing-observed.log')} For Output As #handle
    Print #handle, result
    Close #handle
    If Err.Number <> 0 Then Debug.Print "OBSERVER_LOG_FAILED: "; Err.Description
End Sub
'''


def run_step(folder, name, macro, opened, defines, outputs, expected, live, timeout, observer=None):
    report_path = folder / (name + '-runner.json')
    outputs = [*outputs, report_path]
    input_files = [path for path in folder.iterdir() if path.suffix.lower() in ('.sldprt', '.sldasm', '.slddrw', '.asmdot', '.drwdot')]
    profile = {'schema_version': 1, 'test_root': str(folder), 'reviewed': live,
               'copies_confirmed': True, 'allow_model_changes': True,
               'macro_sha256': code_hash(prepare_code(macro.read_text(encoding='utf-8'), defines)),
               'entry_point': {'module': 'TestMacro', 'procedure': 'main'},
               'inputs': [{'path': str(path), 'sha256': file_hash(path)} for path in input_files],
               'outputs': [str(path) for path in outputs], 'allowed_dialogs': []}
    if observer:
        profile['observer_sha256'] = code_hash(prepare_code(observer.read_text(encoding='utf-8'), []))
    profile_path = folder / (name + '-profile.json')
    profile_path.write_text(json.dumps(profile, indent=2), encoding='utf-8')
    command = [sys.executable, '-B', str(ROOT / 'scripts/run_in_solidworks.py'), str(macro),
               '--profile', str(profile_path), '--open', str(opened), '--confirm-model-changes',
               '--close-opened', '--timeout', str(timeout), '--json', str(report_path)]
    for define in defines:
        command += ['--define', define]
    for path, text in expected:
        command += ['--expect-text', str(path), text]
    if observer:
        command += ['--observe', str(observer)]
    (folder / (name + '-command.json')).write_text(json.dumps(command, indent=2), encoding='utf-8')
    if not live:
        return {'case': name, 'status': 'NOT_TESTED', 'profile': str(profile_path)}
    checked = [str(macro)] + ([str(observer)] if observer else [])
    checks = []
    for tool, extra in (('check_vba.py', []), ('check_api.py', [])):
        check = subprocess.run([sys.executable, str(ROOT / 'scripts' / tool), *checked, *extra], capture_output=True)
        checks.append(check.stdout + check.stderr)
        if check.returncode:
            (folder / (name + '-static.log')).write_bytes(b'\n'.join(checks))
            return {'case': name, 'status': 'STATIC_FAIL', 'session_state': 'IDLE'}
    (folder / (name + '-static.log')).write_bytes(b'\n'.join(checks))
    try:
        completed = subprocess.run(command, capture_output=True, timeout=timeout + 30)
    except subprocess.TimeoutExpired:
        return {'case': name, 'status': 'UNKNOWN', 'session_state': 'UNKNOWN', 'reason': 'Controller timeout; no automatic retry'}
    (folder / (name + '-controller.log')).write_bytes(completed.stdout + completed.stderr)
    report = json.loads(report_path.read_text(encoding='utf-8')) if report_path.exists() else {}
    unchanged = all(file_hash(item['path']) == item['sha256'] for item in profile['inputs'])
    result = {'case': name, 'status': report.get('status', 'UNKNOWN'), 'session_state': report.get('session_state'),
              'input_files_unchanged': unchanged, 'runner': str(report_path), 'macro_sha256': profile['macro_sha256'],
              'controller_exit_code': completed.returncode, 'runner_status': report.get('status')}
    result['solidworks_revision'] = report.get('solidworks')
    if report.get('session_state') != 'IDLE':
        result.update(status='UNKNOWN', session_state='UNKNOWN', reason='Runner did not confirm an IDLE session; no automatic retry')
    elif completed.returncode != 0:
        # PASS JSON can precede a failure while finalizing the durable session.
        if report.get('status') == 'PASS':
            result.update(status='UNKNOWN', session_state='UNKNOWN', reason='PASS report contradicts controller failure; inspect session state')
        else:
            result.update(status='FAIL', reason='Controller exited unsuccessfully')
    if not unchanged:
        if result['status'] != 'UNKNOWN':
            result['status'] = 'FAIL'
        result['input_hash_failure'] = True
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--blank', type=Path, required=True)
    parser.add_argument('--block', type=Path, required=True)
    parser.add_argument('--assembly-template', type=Path, required=True)
    parser.add_argument('--drawing-template', type=Path, required=True)
    parser.add_argument('--test-root', type=Path, required=True)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument('--run', action='store_true')
    mode.add_argument('--prepare', action='store_true', help='prepare the first fixture profile only; never attach to SOLIDWORKS')
    parser.add_argument('--timeout', type=float, default=45)
    args = parser.parse_args()
    folder = args.test_root.resolve()
    folder.mkdir(parents=True, exist_ok=False)
    for original, name in ((args.blank, 'blank.sldprt'), (args.block, 'block.sldprt'),
                           (args.assembly_template, 'assembly.asmdot'), (args.drawing_template, 'drawing.drwdot')):
        shutil.copy2(original.resolve(strict=True), folder / name)
    fixture = folder / 'generate-fixture.bas'
    fixture.write_text(fixture_macro(folder), encoding='utf-8')
    observer = folder / 'drawing-observer.bas'
    observer.write_text(drawing_observer(folder), encoding='utf-8')
    results = []

    def step(name, macro, opened, defines, outputs, expected, observe=None):
        result = run_step(folder, name, macro, opened, defines, outputs, expected, args.run, args.timeout, observe)
        results.append(result)
        (folder / 'smoke-report.json').write_text(json.dumps({'results': results}, indent=2), encoding='utf-8')
        print(name + ': ' + result['status'], flush=True)
        if result['status'] not in ('PASS', 'NOT_TESTED'):
            raise RuntimeError('Stopped after ' + name + ': ' + result['status'] + '; inspect runner/session evidence')

    step('fixture', fixture, folder / 'blank.sldprt', [], [folder / 'fixture.sldasm', folder / 'fixture.log'],
         [(folder / 'fixture.log', 'RESULT fixture PASS; components=2')])
    if not args.run:
        return 0  # later profiles require generated CAD/dependency closure
    recipes = ROOT / 'recipes'
    for name, recipe, setup, opened, extra in (
        ('mate_reference', recipes / 'assemblies/mate_reference.bas', selection_setup(), folder / 'block.sldprt', ['REFERENCE_NAME="SmokeReference"']),
        ('assembly_context_mate', recipes / 'assemblies/assembly_context_mate.bas', selection_setup(True), folder / 'fixture.sldasm', []),
    ):
        macro = folder / (name + '.bas')
        macro.write_text(wrapper(recipe, setup), encoding='utf-8')
        log = folder / (name + '.log')
        step(name, macro, opened, ['SHOW_MESSAGES=False', 'RESULT_PATH=' + literal(log), *extra], [log], [(log, 'RESULT ' + name + ' PASS')])
    insert = folder / 'insert_component_transform.bas'
    setup = f'''    Dim smokeApp As SldWorks.SldWorks, smokeModel As SldWorks.ModelDoc2
    Dim smokeAssemblyName As String, smokeErrors As Long, smokeWarnings As Long
    Set smokeApp = Application.SldWorks
    smokeAssemblyName = smokeApp.ActiveDoc.GetTitle
    Set smokeModel = smokeApp.OpenDoc6({literal(folder / 'block.sldprt')}, swDocPART, swOpenDocOptions_Silent, "", smokeErrors, smokeWarnings)
    If smokeModel Is Nothing Then Err.Raise vbObjectError + 510, , "Source part open failed"
    Set smokeModel = smokeApp.ActivateDoc3(smokeAssemblyName, False, swDontRebuildActiveDoc, smokeErrors)'''
    insert.write_text(wrapper(recipes / 'assemblies/insert_component_transform.bas', setup), encoding='utf-8')
    log = folder / 'insert_component_transform.log'
    step('insert_component_transform', insert, folder / 'fixture.sldasm', ['SHOW_MESSAGES=False', 'SOURCE_PART_PATH=' + literal(folder / 'block.sldprt'),
         'RESULT_PATH=' + literal(log)], [log], [(log, 'RESULT insert_component_transform PASS')])
    step('create_three_views', recipes / 'drawings/create_three_views.bas', folder / 'block.sldprt',
         ['SHOW_MESSAGES=False', 'DRAWING_TEMPLATE=' + literal(folder / 'drawing.drwdot')],
         [folder / 'fixture.slddrw', folder / 'drawing-observed.log'],
         [(folder / 'drawing-observed.log', 'RESULT drawing_observer PASS; model_views=3')], observer)
    for name, recipe, opened in (
        ('inspect_drawing_semantics', recipes / 'drawings/inspect_drawing_semantics.bas', folder / 'fixture.slddrw'),
        ('inspect_sheetmetal_cutlist', recipes / 'sheetmetal/inspect_sheetmetal_cutlist.bas', folder / 'block.sldprt'),
    ):
        log = folder / (name + '.log')
        step(name, recipe, opened, ['SHOW_MESSAGES=False', 'RESULT_PATH=' + literal(log)], [log], [(log, 'RESULT ' + name + ' PASS')])
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
