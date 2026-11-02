"""Compare a STEP imported by OpenCascade with source geometry measured in SOLIDWORKS.

Checks solid validity/count and per-body volume, area, centroid and world-axis bounds.
This metric comparison is NOT a Boolean shape-equivalence proof or a CAM toolpath simulation.
The CLI limits input size and analysis time, but does not impose an OS memory limit.
"""
import argparse
import importlib.metadata
import json
import math
import os
import subprocess
from pathlib import Path
from types import SimpleNamespace
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))  # sibling imports also under python -I
from validation_reports import file_hash, positive, report_path, write_report  # noqa: E402

_dll_handles = []
DEFAULT_MAX_BYTES = 50 * 1024 * 1024
DEFAULT_TIMEOUT_SECONDS = 120.0
TERMINATE_GRACE_SECONDS = 2.0


def load_ocp():
    # Some Windows wheels need process-local DLL search paths. Never alter system/user PATH.
    if os.name == 'nt' and not _dll_handles:
        for distribution, directories in [('cadquery-ocp', ['cadquery_ocp.libs']), ('vtk', ['vtk.libs', 'vtkmodules'])]:
            package = importlib.metadata.distribution(distribution)
            for name in directories:
                folder = Path(package.locate_file(name))
                if folder.is_dir():
                    _dll_handles.append(os.add_dll_directory(str(folder)))
    # Import only the geometry bindings: full CadQuery import crashed this host on interpreter exit.
    from OCP.STEPControl import STEPControl_Reader
    from OCP.IFSelect import IFSelect_RetDone
    from OCP.BRepCheck import BRepCheck_Analyzer
    from OCP.BRepGProp import BRepGProp
    from OCP.GProp import GProp_GProps
    from OCP.BRepBndLib import BRepBndLib
    from OCP.Bnd import Bnd_Box
    from OCP.TopExp import TopExp_Explorer
    from OCP.TopAbs import TopAbs_SOLID, TopAbs_FACE, TopAbs_EDGE, TopAbs_VERTEX
    from OCP.StepRepr import StepRepr_RepresentationContext
    return SimpleNamespace(**locals())


def validate_metric_document(document):
    if not isinstance(document, dict) or document.get('units') != 'mm':
        raise ValueError('Source and imported metrics require explicit millimetres')
    bodies = document.get('bodies')
    if not isinstance(bodies, list) or not bodies or len(bodies) > 500:
        raise ValueError('Expected between 1 and 500 source solid bodies')
    for body in bodies:
        if not isinstance(body, dict) or set(body) != {'volume_mm3', 'area_mm2', 'centroid_mm', 'bounds_mm'}:
            raise ValueError('Body metrics require volume, area, centroid and bounds')
        positive(body['volume_mm3'], 'volume_mm3')
        positive(body['area_mm2'], 'area_mm2')
        for name, count in [('centroid_mm', 3), ('bounds_mm', 6)]:
            values = body[name]
            if (not isinstance(values, list) or len(values) != count
                    or any(isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) for value in values)):
                raise ValueError('Invalid ' + name)
        if any(body['bounds_mm'][i] >= body['bounds_mm'][i+3] for i in range(3)):
            raise ValueError('Solid bounds must have positive extent on every axis')


def metric_deltas(expected, actual, tolerances):
    checks = {}
    for name, absolute in [('volume_mm3', tolerances['volume_mm3']), ('area_mm2', tolerances['area_mm2'])]:
        error = abs(actual[name] - expected[name])
        limit = absolute + tolerances['relative'] * abs(expected[name])
        checks[name] = {'absolute_error': error, 'limit': limit, 'ok': error <= limit}
    for name in ('centroid_mm', 'bounds_mm'):
        error = max(abs(a-b) for a, b in zip(actual[name], expected[name]))
        checks[name] = {'max_axis_error_mm': error, 'limit_mm': tolerances['linear_mm'], 'ok': error <= tolerances['linear_mm']}
    return checks


def compare_metrics(expected, actual, tolerances):
    if not isinstance(tolerances, dict) or set(tolerances) != {'linear_mm', 'relative', 'area_mm2', 'volume_mm3'}:
        raise ValueError('Provide all four exact tolerance keys')
    for name, value in tolerances.items():
        positive(value, name)
    validate_metric_document(expected)
    validate_metric_document(actual)
    source, imported = expected['bodies'], actual['bodies']
    result = {'status': 'FAIL', 'cam_status': 'NOT_CONFIGURED', 'source_solid_count': len(source),
              'imported_solid_count': len(imported), 'tolerances': tolerances, 'matches': [], 'issues': []}
    if len(source) != len(imported):
        result['issues'].append({'code': 'SOLID_COUNT_MISMATCH'})
        return result
    comparisons = [[metric_deltas(a, b, tolerances) for b in imported] for a in source]
    candidates = [[j for j, checks in enumerate(row) if all(check['ok'] for check in checks.values())]
                  for row in comparisons]
    # Maximum bipartite matching: STEP body order is not a source-identity contract.
    assignments = {}
    def assign(i, seen):
        for j in candidates[i]:
            if j in seen:
                continue
            seen.add(j)
            if j not in assignments or assign(assignments[j], seen):
                assignments[j] = i
                return True
        return False
    for i in range(len(source)):
        if not assign(i, set()):
            nearest = min(range(len(imported)), key=lambda j: sum(not item['ok'] for item in comparisons[i][j].values()))
            result['issues'].append({'code': 'BODY_METRICS_MISMATCH', 'source_body': i,
                                     'best_candidate_deltas': comparisons[i][nearest]})
    for j, i in sorted(assignments.items()):
        result['matches'].append({'source_body': i, 'imported_body': j, 'deltas': comparisons[i][j]})
    result['status'] = 'PASS' if len(assignments) == len(source) and not result['issues'] else 'FAIL'
    return result


def read_step_metrics(path):
    ocp = load_ocp()
    reader = ocp.STEPControl_Reader()
    if reader.ReadFile(str(path)) != ocp.IFSelect_RetDone:
        raise ValueError('STEP parser rejected the file')
    # Choosing output mm is not evidence that the STEP declares its SOURCE length unit.
    # Inspect parsed geometric contexts, including complex STEP entities, before transfer.
    model = reader.StepModel()
    contexts = []
    supported = {'StepBasic_SiUnitAndLengthUnit', 'StepBasic_ConversionBasedUnitAndLengthUnit'}
    for i in range(1, model.NbEntities() + 1):
        entity = model.Value(i)
        if not isinstance(entity, ocp.StepRepr_RepresentationContext) or not hasattr(entity, 'CoordinateSpaceDimension'):
            continue
        # UV parameter domains do not describe physical length and legitimately have no unit assignment.
        if ('ParametricRepresentationContext' in entity.DynamicType().Name()
                and entity.CoordinateSpaceDimension() == 2):
            continue
        if not hasattr(entity, 'NbUnits'):
            raise ValueError('STEP geometric context has no explicit assigned units')
        length_types = [entity.UnitsValue(j).DynamicType().Name() for j in range(1, entity.NbUnits() + 1)
                        if 'LengthUnit' in entity.UnitsValue(j).DynamicType().Name()]
        if len(length_types) != 1 or length_types[0] not in supported:
            raise ValueError('STEP geometric context has missing, ambiguous or unsupported length units')
        contexts.append({'dimension': entity.CoordinateSpaceDimension(), 'length_unit_type': length_types[0]})
    if not contexts:
        raise ValueError('STEP contains no explicit supported geometric unit context')
    reader.SetSystemLengthUnit(1.0)  # Explicit millimetres, set AFTER reading a non-null model.
    if not reader.TransferRoots():
        raise ValueError('STEP contains no transferable roots')
    imported = reader.OneShape()
    def shapes(shape, kind):
        explorer, found = ocp.TopExp_Explorer(shape, kind), []
        while explorer.More():
            found.append(explorer.Current())
            explorer.Next()
        return found
    solids = shapes(imported, ocp.TopAbs_SOLID)
    if (not ocp.BRepCheck_Analyzer(imported).IsValid() or not solids
            or any(not ocp.BRepCheck_Analyzer(solid).IsValid() for solid in solids)):
        raise ValueError('STEP does not contain valid closed solid geometry')
    for kind in (ocp.TopAbs_FACE, ocp.TopAbs_EDGE, ocp.TopAbs_VERTEX):
        all_subshapes = {hash(shape) for shape in shapes(imported, kind)}
        solid_subshapes = {hash(shape) for solid in solids for shape in shapes(solid, kind)}
        if all_subshapes != solid_subshapes:
            raise ValueError('STEP contains extra faces, edges or vertices outside solid bodies')
    bodies = []
    for solid in solids:
        volume, area, bounds = ocp.GProp_GProps(), ocp.GProp_GProps(), ocp.Bnd_Box()
        ocp.BRepGProp.VolumeProperties_s(solid, volume, True, False, False)
        ocp.BRepGProp.SurfaceProperties_s(solid, area)
        ocp.BRepBndLib.AddOptimal_s(solid, bounds, False, False)
        centroid = volume.CentreOfMass()
        bodies.append({'volume_mm3': volume.Mass(), 'area_mm2': area.Mass(),
                       'centroid_mm': [centroid.X(), centroid.Y(), centroid.Z()],
                       'bounds_mm': list(bounds.Get())})
    return {'units': 'mm', 'bodies': bodies, 'source_length_contexts': contexts,
            'kernel': 'OpenCascade / cadquery-ocp ' + importlib.metadata.version('cadquery-ocp')}


def validate_step(path, source_metrics, tolerances, max_bytes=DEFAULT_MAX_BYTES):
    if isinstance(max_bytes, bool) or not isinstance(max_bytes, int) or max_bytes <= 0:
        raise ValueError('Maximum input size must be a positive integer')
    if Path(path).stat().st_size > max_bytes:
        raise ValueError('STEP input exceeds maximum size of %d bytes' % max_bytes)
    digest = file_hash(path)
    actual = read_step_metrics(path)
    result = compare_metrics(source_metrics, actual, tolerances)
    result.update(format='STEP', sha256=digest, imported_metrics=actual,
                  source_configuration=source_metrics.get('configuration'))
    return result


def run_isolated(path, source_metrics, tolerances, max_bytes=DEFAULT_MAX_BYTES,
                 timeout=DEFAULT_TIMEOUT_SECONDS):
    if isinstance(timeout, bool) or not isinstance(timeout, (int, float)) or not math.isfinite(timeout) or timeout <= 0:
        raise ValueError('timeout must be finite and greater than zero')
    request = json.dumps({'path': str(path), 'source_metrics': source_metrics,
                          'tolerances': tolerances, 'max_bytes': max_bytes})
    process = subprocess.Popen([sys.executable, '-I', str(Path(__file__).resolve()), '--worker'],
                               stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                               stderr=subprocess.DEVNULL, text=True, encoding='utf-8')
    try:
        stdout, _ = process.communicate(request, timeout=timeout)
    except subprocess.TimeoutExpired:
        process.terminate()
        try:
            process.communicate(timeout=TERMINATE_GRACE_SECONDS)
        except subprocess.TimeoutExpired:
            process.kill()
            process.communicate()
        raise TimeoutError('STEP analysis exceeded %.3g seconds' % timeout)
    if process.returncode != 0:
        raise RuntimeError('STEP analysis worker exited with code %s' % process.returncode)
    try:
        response = json.loads(stdout)
        if not isinstance(response, dict) or 'status' not in response:
            raise ValueError('missing status')
    except (ValueError, TypeError) as error:
        raise RuntimeError('Invalid response from STEP analysis worker') from error
    return response


def worker_main():
    # Native parser output must not corrupt the JSON protocol on stdout.
    protocol = os.fdopen(os.dup(sys.stdout.fileno()), 'w', encoding='utf-8')
    os.dup2(sys.stderr.fileno(), sys.stdout.fileno())
    try:
        request = json.load(sys.stdin)
        result = validate_step(request['path'], request['source_metrics'], request['tolerances'],
                               request['max_bytes'])
    except Exception as error:
        result = {'status': 'ERROR', 'cam_status': 'NOT_CONFIGURED', 'exception': str(error)}
    with protocol:
        json.dump(result, protocol, allow_nan=False)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source')
    parser.add_argument('--source-metrics', required=True)
    parser.add_argument('--tolerances', required=True)
    parser.add_argument('--json', required=True)
    parser.add_argument('--max-bytes', type=int, default=DEFAULT_MAX_BYTES,
                        help='Maximum STEP input bytes (default: 50 MiB)')
    parser.add_argument('--timeout', type=float, default=DEFAULT_TIMEOUT_SECONDS,
                        help='Analysis timeout in seconds (default: 120); no memory limit is imposed')
    args = parser.parse_args()
    try:
        output = report_path(args.json, [args.source, args.source_metrics, args.tolerances])
    except ValueError as error:
        parser.error(str(error))
    try:
        metrics = json.loads(Path(args.source_metrics).read_text(encoding='utf-8'))
        tolerances = json.loads(Path(args.tolerances).read_text(encoding='utf-8'))
        result = run_isolated(args.source, metrics, tolerances, args.max_bytes, args.timeout)
    except Exception as error:
        result = {'status': 'ERROR', 'cam_status': 'NOT_CONFIGURED', 'exception': str(error)}
    write_report(output, result)
    print(result['status'], 'CAM:', result['cam_status'])
    return 0 if result['status'] == 'PASS' else 1


if __name__ == '__main__':
    if sys.argv[1:] == ['--worker']:
        worker_main()
    else:
        raise SystemExit(main())
