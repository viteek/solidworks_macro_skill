#!/usr/bin/env python3
"""Verify independently captured JSON observations; never executes or certifies CAD code.

--output must name a new file. Existing inputs, evidence and reports are never overwritten.
"""
import argparse
from datetime import datetime, timezone
import hashlib
import json
import math
from pathlib import Path
import sys

INPUT_PHASES = ('static', 'api', 'project_compile', 'runtime', 'safety')
PHASES = INPUT_PHASES + ('result',)
STATUSES = ('PASS', 'FAIL', 'NOT_TESTED')
ENV_KEYS = ('solidworks_version', 'service_pack', 'api_revision', 'architecture', 'ui_language')


def read_json(path):
    def unique(pairs):
        result = {}
        for key, value in pairs:
            if key in result:
                raise ValueError('Duplicate JSON key: ' + key)
            result[key] = value
        return result
    def invalid(value):
        raise ValueError('Non-finite JSON constant: ' + value)
    return json.loads(Path(path).read_text(encoding='utf-8'), object_pairs_hook=unique,
                      parse_constant=invalid)


def spec_hash(spec):
    return hashlib.sha256(json.dumps(spec, sort_keys=True, allow_nan=False,
                                     separators=(',', ':')).encode()).hexdigest()


def confined(root, name):
    if not isinstance(name, str) or not name or Path(name).is_absolute() or ':' in name:
        raise ValueError('Artifact must be a relative path: ' + repr(name))
    root = Path(root).resolve()
    target = (root / name).resolve()
    if not target.is_relative_to(root):
        raise ValueError('Artifact escapes evidence root: ' + name)
    return target


def identity(document):
    if document.get('schema_version') != 1:
        raise ValueError('Expected schema_version 1')
    if not isinstance(document.get('task_id'), str) or not document['task_id']:
        raise ValueError('Missing task_id')
    fixture = document.get('fixture', {})
    if not all(isinstance(fixture.get(k), str) and fixture[k] for k in ('id', 'revision')):
        raise ValueError('Fixture requires nonempty id and revision')


def environment(document):
    env = document.get('environment', {})
    if not all(isinstance(env.get(k), str) and env[k].strip() for k in ENV_KEYS):
        raise ValueError('Environment requires version, service pack, API revision and architecture')
    return env


def provenance(document):
    value = document.get('provenance', {})
    if not isinstance(value.get('run_id'), str) or not value['run_id']:
        raise ValueError('Provenance requires run_id')
    for key in ('macro_sha256', 'fixture_sha256'):
        digest = value.get(key, '')
        if not isinstance(digest, str) or len(digest) != 64 or any(c not in '0123456789abcdef' for c in digest):
            raise ValueError('Provenance requires SHA256: ' + key)
    try:
        captured = datetime.fromisoformat(value['collected_at'].replace('Z', '+00:00'))
        if captured.tzinfo is None or captured > datetime.now(timezone.utc):
            raise ValueError('Timestamp must be timezone-aware and not in the future')
    except (KeyError, TypeError, AttributeError) as exc:
        raise ValueError('Provenance requires ISO8601 collected_at') from exc
    return value


def validate_spec(spec):
    identity(spec)
    assertions = spec.get('assertions')
    if not isinstance(assertions, list) or not assertions:
        raise ValueError('Task requires assertions')
    seen = set()
    for item in assertions:
        key = item.get('id')
        if not isinstance(key, str) or not key or key in seen:
            raise ValueError('Missing or duplicate assertion id')
        seen.add(key)
        if item.get('op') not in ('equals', 'number', 'exists', 'unordered_equals', 'file'):
            raise ValueError('Unknown assertion operation')
        if item['op'] != 'file' and not isinstance(item.get('path'), str):
            raise ValueError('Assertion requires JSON pointer path')
        if item['op'] in ('equals', 'number', 'unordered_equals') and 'expected' not in item:
            raise ValueError('Assertion requires expected value')
        if item['op'] == 'number':
            for k in ('expected', 'abs_tol', 'rel_tol'):
                v = item.get(k, 0)
                if type(v) not in (int, float) or not math.isfinite(v) or (k != 'expected' and v < 0):
                    raise ValueError('Invalid numeric assertion: ' + key)
        if item['op'] == 'unordered_equals' and not isinstance(item['expected'], list):
            raise ValueError('unordered_equals expects a list')
        if item['op'] == 'file':
            if not isinstance(item.get('file'), str):
                raise ValueError('File assertion requires relative file')
            if type(item.get('min_bytes', 1)) is not int or item.get('min_bytes', 1) < 1:
                raise ValueError('min_bytes must be positive')


def pointer(values, path):
    if path == '':
        return values
    if not path.startswith('/'):
        raise ValueError('Expected JSON pointer: ' + path)
    current = values
    for part in path[1:].split('/'):
        part = part.replace('~1', '/').replace('~0', '~')
        if isinstance(current, list):
            if not part.isdigit():
                raise KeyError(part)
            current = current[int(part)]
        elif isinstance(current, dict):
            current = current[part]
        else:
            raise KeyError(part)
    return current


def strict_equal(left, right):
    # Python's True == 1 must not pass an object/count assertion.
    return json.dumps(left, sort_keys=True, allow_nan=False) == json.dumps(right, sort_keys=True, allow_nan=False)


def evaluate(item, values, root):
    result = {'id': item['id'], 'status': 'FAIL'}
    op = item['op']
    if op == 'file':
        target = confined(root, item['file'])
        if not target.is_file():
            return dict(result, status='NOT_TESTED', reason='Artifact missing', file=item['file'])
        data = target.read_bytes()
        result.update(file=item['file'], bytes=len(data), sha256=hashlib.sha256(data).hexdigest())
        passed = len(data) >= item.get('min_bytes', 1)
        passed = passed and all(s.encode('utf-8') in data for s in item.get('contains', []))
        if 'sha256' in item:
            passed = passed and result['sha256'] == item['sha256']
    else:
        try:
            value = pointer(values, item['path'])
        except (KeyError, IndexError):
            return dict(result, status='NOT_TESTED', reason='Observation missing', path=item['path'])
        result['path'] = item['path']
        if op == 'exists':
            passed = value is not None
        elif op == 'number':
            passed = type(value) in (int, float) and math.isfinite(value) and math.isclose(
                value, item['expected'], abs_tol=item.get('abs_tol', 0), rel_tol=item.get('rel_tol', 0))
        elif op == 'unordered_equals':
            passed = isinstance(value, list) and sorted(json.dumps(v, sort_keys=True) for v in value) == sorted(
                json.dumps(v, sort_keys=True) for v in item['expected'])
        else:
            try:
                passed = strict_equal(value, item['expected'])
            except ValueError:
                passed = False
        # Do not emit NaN, even when a caller supplied an in-memory invalid value.
        result['expected'] = item.get('expected')
    result['status'] = 'PASS' if passed else 'FAIL'
    return result


def aggregate(statuses):
    statuses = list(statuses)
    return 'FAIL' if 'FAIL' in statuses else ('PASS' if statuses and all(s == 'PASS' for s in statuses) else 'NOT_TESTED')


def verify(spec, observations, evidence_root):
    validate_spec(spec)
    identity(observations)
    if any(spec[k] != observations[k] for k in ('task_id', 'fixture')):
        raise ValueError('Task/fixture identity mismatch')
    env = environment(observations)
    captured = provenance(observations)
    for kind in ('macro', 'fixture'):
        target = confined(evidence_root, captured.get(kind + '_artifact'))
        if not target.is_file() or not target.stat().st_size or hashlib.sha256(target.read_bytes()).hexdigest() != captured[kind + '_sha256']:
            raise ValueError('Missing/changed provenance artifact: ' + kind)
    phases = {}
    for phase in INPUT_PHASES:
        evidence = observations.get('evidence', {}).get(phase, {'status': 'NOT_TESTED'})
        status = evidence.get('status')
        if status not in STATUSES:
            raise ValueError('Invalid evidence status: ' + phase)
        artifacts = evidence.get('artifacts', [])
        if not isinstance(artifacts, list) or len(artifacts) != len(set(artifacts)):
            raise ValueError('Invalid/duplicate evidence artifacts: ' + phase)
        files = []
        for name in artifacts:
            target = confined(evidence_root, name)
            if target.is_file() and target.stat().st_size:
                files.append({'file': name, 'sha256': hashlib.sha256(target.read_bytes()).hexdigest()})
        backed = bool(artifacts) and len(files) == len(artifacts)
        phases[phase] = {'status': status if backed or status == 'NOT_TESTED' else 'NOT_TESTED',
                         'artifacts': files}
        if status != 'NOT_TESTED' and not backed:
            phases[phase]['reason'] = 'Claim has missing/empty evidence artifacts'
    results = [evaluate(a, observations.get('values', {}), evidence_root) for a in spec['assertions']]
    result_status = aggregate(r['status'] for r in results)
    if result_status == 'PASS' and phases['runtime']['status'] != 'PASS':
        result_status = 'NOT_TESTED'
    phases['result'] = {'status': result_status}
    return {'schema_version': 1, 'task_id': spec['task_id'], 'fixture': spec['fixture'],
            'spec_sha256': spec_hash(spec), 'environment': env, 'provenance': captured, 'phases': phases,
            'assertions': results, 'status': aggregate(p['status'] for p in phases.values()),
            'scope': 'Evidence verification only; artifact provenance and collector correctness require independent review.'}


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--spec', required=True, type=Path)
    parser.add_argument('--observations', required=True, type=Path)
    parser.add_argument('--evidence-root', required=True, type=Path)
    parser.add_argument('--output', required=True, type=Path,
                        help='new report path; refuses any existing file or directory')
    args = parser.parse_args(argv)
    try:
        report = verify(read_json(args.spec), read_json(args.observations), args.evidence_root)
        payload = json.dumps(report, indent=2, allow_nan=False) + '\n'
        args.output.parent.mkdir(parents=True, exist_ok=True)
        # Exclusive creation protects inputs/evidence even if a path appears after validation.
        with args.output.open('x', encoding='utf-8') as output:
            output.write(payload)
        print(report['status'])
        return {'PASS': 0, 'FAIL': 1, 'NOT_TESTED': 2}[report['status']]
    except (ValueError, OSError, TypeError) as exc:
        print('Invalid verification input: ' + str(exc), file=sys.stderr)
        return 3


if __name__ == '__main__':
    sys.exit(main())
