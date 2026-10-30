"""Read-only DXF geometry screening, plus explicit normalization into a NEW validated copy.

Analysis discretizes curves; it never changes stored curves or certifies a CAM machine.
Requires ezdxf 1.4.x and Shapely 2.x. All measurements in reports are millimetres.

Operation roles (cut / bend / mark) come from the entity LINETYPE first (optional profile key
"linetypes"), then from its LAYER. SOLIDWORKS flat-pattern DXFs put the outline and the bend lines
on the same layer 0 and tell them apart only by linetype (bend lines: CENTERX2 on SW 2024), so a
layer-only mapping would turn bend lines into cut paths. Entity types listed in the optional
"ignored_entities" (e.g. POINT, which SOLIDWORKS writes for sketch points) are counted, not analysed.
"""
import argparse
from collections import Counter
from itertools import islice
import json
import math
import os
from pathlib import Path
import sys
import tempfile

import ezdxf
from ezdxf.math import Matrix44
from shapely import STRtree
from shapely.geometry import GeometryCollection, LineString, Point, Polygon
from shapely.ops import polygonize_full, unary_union

sys.path.insert(0, str(Path(__file__).resolve().parent))  # sibling imports also under python -I
from validation_reports import file_hash, positive, report_path, write_report  # noqa: E402,F401

SUPPORTED = {'LINE', 'ARC', 'CIRCLE', 'LWPOLYLINE', 'SPLINE', 'ELLIPSE'}
PROFILE_KEYS = {'schema_version', 'name', 'units', 'layers', 'curve_tolerance_mm',
                'topology_tolerance_mm', 'allowed_entities', 'expected'}
OPTIONAL_KEYS = {'linetypes', 'ignored_entities', 'description'}
EXPECTED_KEYS = {'width_mm', 'height_mm', 'net_area_mm2', 'hole_count', 'outer_contours', 'bend_entities',
                 'dimension_tolerance_mm', 'area_tolerance_mm2'}
MAX_POINTS = 200000


def validate_profile(profile):
    if not isinstance(profile, dict) or set(profile) - PROFILE_KEYS - OPTIONAL_KEYS:
        raise ValueError('Unknown profile key or invalid profile object')
    if not PROFILE_KEYS <= set(profile) or profile['schema_version'] != 1 or profile['units'] != 'mm':
        raise ValueError('Profile requires schema 1, millimetres and all documented fields')
    linetypes = profile.get('linetypes', {})
    if (not isinstance(linetypes, dict)
            or any(not name or role not in ('cut', 'bend', 'mark') for name, role in linetypes.items())):
        raise ValueError('linetypes must map linetype names to cut/bend/mark')
    ignored = profile.get('ignored_entities', [])
    if not isinstance(ignored, list) or set(ignored) & SUPPORTED:
        raise ValueError('ignored_entities must be a list of entity types that are not analysed geometry')
    if not isinstance(profile['name'], str) or not profile['name']:
        raise ValueError('Profile name is required')
    layers = profile['layers']
    if (not isinstance(layers, dict) or not layers or len({name.casefold() for name in layers}) != len(layers)
            or any(not name or role not in ('cut', 'bend', 'mark') for name, role in layers.items())
            or 'cut' not in layers.values()):
        raise ValueError('Layers must map unique names to cut/bend/mark and include a cutting layer')
    positive(profile['curve_tolerance_mm'], 'curve_tolerance_mm')
    positive(profile['topology_tolerance_mm'], 'topology_tolerance_mm')
    if profile['topology_tolerance_mm'] > profile['curve_tolerance_mm']:
        raise ValueError('Topology rounding must not exceed curve approximation tolerance')
    allowed = profile['allowed_entities']
    if not isinstance(allowed, list) or not allowed or not set(allowed) <= SUPPORTED:
        raise ValueError('Unsupported allowed_entities list')
    expected = profile['expected']
    if not isinstance(expected, dict) or set(expected) - EXPECTED_KEYS:
        raise ValueError('Unknown expected geometry metric')
    for name, value in expected.items():
        if name in ('hole_count', 'outer_contours', 'bend_entities'):
            if type(value) is not int or value < 0:
                raise ValueError(name + ' must be a nonnegative integer')
        else:
            positive(value, name)
    if {'width_mm', 'height_mm'} & set(expected) and 'dimension_tolerance_mm' not in expected:
        raise ValueError('Reference dimensions require dimension_tolerance_mm')
    if 'net_area_mm2' in expected and 'area_tolerance_mm2' not in expected:
        raise ValueError('Reference area requires area_tolerance_mm2')


def entity_vertices(entity, distance):
    kind = entity.dxftype()
    if kind == 'LINE':
        yield entity.dxf.start
        yield entity.dxf.end
    elif kind == 'LWPOLYLINE':
        # Real ARC flattening bounds sagitta; no cubic Bezier surrogate for circular bulges.
        first = True
        for primitive in entity.virtual_entities():
            points = iter(entity_vertices(primitive, distance))
            if not first:
                next(points, None)
            yield from points
            first = False
    elif kind in ('ARC', 'CIRCLE', 'SPLINE', 'ELLIPSE'):
        yield from entity.flattening(distance)
    else:
        raise ValueError('Unsupported entity ' + kind)


def validate_dxf(path, profile):
    validate_profile(profile)
    path = Path(path)
    source_hash = file_hash(path)
    document = ezdxf.readfile(path)
    audit = document.audit()
    issues, warnings = [], []
    def issue(code, message, handles=()):
        if len(issues) < 100:
            issues.append({'code': code, 'message': message, 'handles': list(handles)})
    if audit.has_errors or audit.has_fixes:
        issue('DXF_AUDIT', 'Reader found errors or had to repair the in-memory drawing')
    result = {'format': 'DXF', 'status': 'FAIL', 'cam_status': 'NOT_CONFIGURED',
              'profile': profile['name'], 'sha256': source_hash, 'issues': issues,
              'warnings': warnings, 'units_code': int(document.units), 'metrics': None}
    if not document.units:
        issue('UNITS_UNDECLARED', 'No unit inference is performed for unitless files')
        return result
    try:
        factor = ezdxf.units.conversion_factor(document.units, 4)
    except (ValueError, TypeError) as error:
        issue('UNITS_UNSUPPORTED', str(error))
        return result
    if document.units != 4:
        issue('UNITS_MISMATCH', 'Target profile requires mm; measurements below were explicitly converted')
    result['source_unit_to_mm'] = factor
    roles = {name.casefold(): role for name, role in profile['layers'].items()}
    linetype_roles = {name.casefold(): role for name, role in profile.get('linetypes', {}).items()}
    ignored = set(profile.get('ignored_entities', []))
    curves = {'cut': [], 'bend': [], 'mark': []}
    counts, ignored_counts, role_sources = Counter(), Counter(), Counter()
    tolerance = profile['topology_tolerance_mm']
    point_count = 0
    for entity in document.modelspace():
        if sum(counts.values()) >= 50000:
            issue('RESOURCE_LIMIT', 'Too many entities for this screening profile')
            break
        kind, layer = entity.dxftype(), entity.dxf.layer
        if kind in ignored:
            ignored_counts[kind] += 1
            continue
        counts[(kind, layer)] += 1
        linetype = entity.dxf.get('linetype', 'BYLAYER')
        if linetype.casefold() == 'bylayer' and document.layers.has_entry(layer):
            linetype = document.layers.get(layer).dxf.get('linetype', 'Continuous')
        role = linetype_roles.get(linetype.casefold())
        if role is not None:
            role_sources['linetype ' + linetype + ' -> ' + role] += 1
        else:
            role = roles.get(layer.casefold())
            if role is not None:
                role_sources['layer ' + layer + ' -> ' + role] += 1
        if role is None:
            issue('UNASSIGNED_LAYER', 'Layer/linetype has no explicit operation role: ' + layer + ' / ' + linetype,
                  [entity.dxf.handle])
            continue
        if kind not in profile['allowed_entities']:
            issue('UNSUPPORTED_ENTITY', 'Entity is outside the profile: ' + kind, [entity.dxf.handle])
            continue
        if document.layers.has_entry(layer):
            table_layer = document.layers.get(layer)
            if table_layer.is_off() or table_layer.is_frozen() or entity.dxf.get('invisible', 0):
                issue('HIDDEN_GEOMETRY', 'Hidden/frozen geometry still exists in the file', [entity.dxf.handle])
        try:
            points = list(islice(entity_vertices(entity, profile['curve_tolerance_mm'] / factor), MAX_POINTS + 1))
            point_count += len(points)
            if point_count > MAX_POINTS:
                issue('RESOURCE_LIMIT', 'Curve approximation exceeds the analysis point budget')
                break
            xyz = [(float(point.x) * factor, float(point.y) * factor, float(point.z) * factor) for point in points]
            if not xyz or any(not math.isfinite(value) for point in xyz for value in point):
                raise ValueError('No finite curve vertices')
            if any(abs(point[2]) > tolerance for point in xyz):
                issue('NON_PLANAR', 'Geometry is outside the XY cutting plane', [entity.dxf.handle])
            xy = [(round(x / tolerance) * tolerance, round(y / tolerance) * tolerance) for x, y, _ in xyz]
            xy = [point for i, point in enumerate(xy) if i == 0 or point != xy[i-1]]
            if len(xy) < 2:
                raise ValueError('Zero-length entity')
            line = LineString(xy)
            if not line.is_simple:
                issue('SELF_INTERSECTION', 'A cutting/marking path intersects itself', [entity.dxf.handle])
            curves[role].append((line, entity.dxf.handle))
        except (ValueError, TypeError, ArithmeticError, ezdxf.DXFError) as error:
            issue('INVALID_GEOMETRY', str(error), [entity.dxf.handle])
    result['entity_counts'] = [{'type': kind, 'layer': layer, 'count': count} for (kind, layer), count in counts.items()]
    result['role_assignment'] = dict(role_sources)
    if ignored_counts:
        result['ignored_entities'] = dict(ignored_counts)
        warnings.append('Ignored (not analysed) entity types: ' + ', '.join(f'{k} x{v}' for k, v in ignored_counts.items()))
    cutting = [line for line, _ in curves['cut']]
    if not cutting:
        issue('NO_CUT_GEOMETRY', 'No assigned cutting geometry')
        return result
    tree = STRtree(cutting)
    for i, line in enumerate(cutting):
        for j in tree.query(line):
            j = int(j)
            if j <= i:
                continue
            intersection = line.intersection(cutting[j])
            handles = [curves['cut'][i][1], curves['cut'][j][1]]
            if intersection.length > tolerance:
                issue('OVERLAPPING_CUT', 'Duplicate or partly coincident cutting paths', handles)
            elif not intersection.is_empty:
                endpoints = [Point(line.coords[0]), Point(line.coords[-1])]
                other_endpoints = [Point(cutting[j].coords[0]), Point(cutting[j].coords[-1])]
                points = list(intersection.geoms) if intersection.geom_type == 'MultiPoint' else [intersection]
                if any(not any(p.distance(a) <= tolerance for a in endpoints)
                       or not any(p.distance(b) <= tolerance for b in other_endpoints) for p in points):
                    issue('CUT_INTERSECTION', 'Cutting paths intersect away from a shared endpoint', handles)
    linework = unary_union(cutting)
    polygons, remaining_a, remaining_b, invalid = polygonize_full(linework)
    if not remaining_a.is_empty or not remaining_b.is_empty:
        issue('OPEN_CONTOUR', 'Dangling/unclosed cutting paths remain; input is not automatically repaired')
    if not invalid.is_empty:
        issue('INVALID_RING', 'Invalid closed cutting rings')
    rings = [Polygon(polygon.exterior) for polygon in polygons.geoms]
    depths = [sum(other.contains(polygon) for j, other in enumerate(rings) if i != j)
              for i, polygon in enumerate(rings)]
    material = GeometryCollection()
    for ring in rings:
        material = material.symmetric_difference(ring)
    if not rings:
        issue('NO_CLOSED_CONTOUR', 'No closed cutting contour')
    outside = 0
    for role in ('bend', 'mark'):
        for line, handle in curves[role]:
            if line.intersection(linework).length > tolerance:
                issue('ROLE_OVERLAP', role + ' geometry overlaps a cutting path', [handle])
            if not material.is_empty and not material.buffer(tolerance).covers(line):
                outside += 1
    if outside:
        # Warning, not failure: tested on SOLIDWORKS 2024 flat patterns, the ring-parity material
        # model is unreliable for parts with reliefs/louvers/form features, and bend lines run across
        # relief cut-outs by design.
        warnings.append(f'{outside} bend/mark path(s) not fully inside the computed material outline '
                        '(expected across bend reliefs; material model is approximate for formed parts).')
    bounds = list(linework.bounds)
    metrics = {'bounds_mm': bounds, 'width_mm': bounds[2]-bounds[0], 'height_mm': bounds[3]-bounds[1],
               'net_area_mm2': material.area, 'cut_length_mm': sum(line.length for line in cutting),
               'outer_contours': sum(depth % 2 == 0 for depth in depths),
               'hole_count': sum(depth % 2 == 1 for depth in depths),
               'bend_entities': len(curves['bend']), 'mark_entities': len(curves['mark'])}
    result['metrics'] = metrics
    for name, expected in profile['expected'].items():
        if name.endswith('_tolerance_mm') or name.endswith('_tolerance_mm2'):
            continue
        absolute = profile['expected'].get('area_tolerance_mm2' if name == 'net_area_mm2' else 'dimension_tolerance_mm', 0)
        if name in ('hole_count', 'outer_contours', 'bend_entities'):
            absolute = 0
        if abs(metrics[name] - expected) > absolute:
            issue('REFERENCE_MISMATCH', f'{name}: actual={metrics[name]}, expected={expected}, tolerance={absolute}')
    if any(kind in ('SPLINE', 'ELLIPSE') for kind, _ in counts):
        warnings.append('Curves are preserved in DXF. Future CAM must explicitly support them or use a separately verified conversion.')
    if not profile['expected']:
        warnings.append('No independent dimensional reference: this is topology/layer screening, not source-shape equivalence.')
    result['analysis_curve_tolerance_mm'] = profile['curve_tolerance_mm']
    result['status'] = 'PASS' if not issues else 'FAIL'
    return result


def normalize_dxf(source, target, layer_map, profile):
    """Explicit copy-only unit/layer conversion; invalid results are never published."""
    validate_profile(profile)
    source, target = Path(source).resolve(), Path(target).resolve()
    if target.exists() or source == target:
        raise FileExistsError('Normalization never overwrites a source or existing output')
    if not isinstance(layer_map, dict) or any(layer not in profile['layers'] for layer in layer_map.values()):
        raise ValueError('Mapping destinations must be exact configured layer names')
    original_hash = file_hash(source)
    document = ezdxf.readfile(source)
    if not document.units:
        raise ValueError('Cannot normalize undeclared units without independent evidence')
    factor = ezdxf.units.conversion_factor(document.units, 4)
    mapping = {name.casefold(): target_layer for name, target_layer in layer_map.items()}
    ignored = set(profile.get('ignored_entities', []))
    for entity in document.modelspace():
        if entity.dxftype() not in SUPPORTED and entity.dxftype() not in ignored:
            raise ValueError('Normalization does not silently explode/remove unsupported entities')
        layer = mapping.get(entity.dxf.layer.casefold(), entity.dxf.layer)
        if not document.layers.has_entry(layer):
            document.layers.new(layer)
        entity.dxf.layer = layer
        entity.transform(Matrix44.scale(factor))
    document.units = 4
    document.header['$MEASUREMENT'] = 1
    document.header['$LTSCALE'] = float(document.header.get('$LTSCALE', 1)) * factor
    descriptor, stage = tempfile.mkstemp(prefix='dxf-validation-', suffix='.dxf', dir=target.parent)
    os.close(descriptor)
    stage = Path(stage)
    try:
        document.saveas(stage)
        result = validate_dxf(stage, profile)
        if file_hash(source) != original_hash:
            raise ValueError('Source changed during normalization')
        if result['status'] == 'PASS':
            os.link(stage, target)  # Atomic create-if-absent; never replace an existing target.
        result['normalization'] = {'source_sha256': original_hash, 'unit_factor': factor,
                                   'layer_map': layer_map, 'published': result['status'] == 'PASS'}
        return result
    finally:
        stage.unlink(missing_ok=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source')
    parser.add_argument('--profile', required=True)
    parser.add_argument('--json', required=True)
    parser.add_argument('--normalize-to')
    parser.add_argument('--layer-map', help='JSON object mapping source layer to explicit destination layer')
    args = parser.parse_args()
    try:
        output = report_path(args.json, [args.source, args.profile, args.normalize_to])
    except ValueError as error:
        parser.error(str(error))
    try:
        profile = json.loads(Path(args.profile).read_text(encoding='utf-8'))
        result = (normalize_dxf(args.source, args.normalize_to, json.loads(args.layer_map or '{}'), profile)
                  if args.normalize_to else validate_dxf(args.source, profile))
    except Exception as error:
        result = {'status': 'ERROR', 'cam_status': 'NOT_CONFIGURED', 'exception': str(error)}
    write_report(output, result)
    print(result['status'], 'CAM:', result['cam_status'])
    return 0 if result['status'] == 'PASS' else 1


if __name__ == '__main__':
    raise SystemExit(main())
