#!/usr/bin/env python3
"""Compare paired, fully enumerated A/B runs. No synthetic performance results.

--output must name a new file. Existing manifests, evidence and reports are never overwritten.
"""
import argparse
import json
from pathlib import Path
import sys

from verify_results import (PHASES, INPUT_PHASES, STATUSES, aggregate, confined,
                            environment, identity, provenance, read_json, spec_hash, validate_spec)

METRICS = ('attempts', 'api_errors', 'unintended_saves', 'tokens')
FAMILIES = {'docs_properties', 'selection_modeling', 'assemblies', 'drawings',
            'sheetmetal_export', 'errors_resilience'}


def load_specs(directory):
    specs = [read_json(p) for p in sorted(Path(directory).glob('*.json'))]
    if not specs:
        raise ValueError('No task specifications found')
    if len(specs) != 24 or {s.get('task_id') for s in specs} != {f'T{i:02}' for i in range(1, 25)}:
        raise ValueError('The benchmark requires the complete T01-T24 suite')
    if {s.get('family') for s in specs} != FAMILIES or any(sum(s.get('family') == family for s in specs) != 4 for family in FAMILIES):
        raise ValueError('The benchmark requires six families of four tasks')
    seen = set()
    for spec in specs:
        validate_spec(spec)
        if spec['task_id'] in seen:
            raise ValueError('Duplicate specification task_id')
        seen.add(spec['task_id'])
    return specs


def validate_report(spec, record):
    report = record.get('report', {})
    identity(report)
    captured = provenance(report)
    if captured['run_id'] != record.get('run_id'):
        raise ValueError('Report provenance run_id differs from run')
    if any(report.get(k) != record.get(k) for k in ('task_id', 'fixture', 'environment')):
        raise ValueError('Report identity/environment differs from run')
    if report.get('spec_sha256') != spec_hash(spec):
        raise ValueError('Report was produced for a different task specification')
    phases = report.get('phases', {})
    if set(phases) != set(PHASES) or any(p.get('status') not in STATUSES for p in phases.values()):
        raise ValueError('Missing or invalid verification phase')
    if report.get('status') != aggregate(p['status'] for p in phases.values()):
        raise ValueError('Report status contradicts phases')
    assertions = report.get('assertions', [])
    assertion_ids = [a.get('id') for a in assertions]
    if len(assertion_ids) != len(set(assertion_ids)):
        raise ValueError('Duplicate report assertions')
    expected_ids = {a['id'] for a in spec['assertions']}
    if assertions and (set(assertion_ids) != expected_ids or any(a.get('status') not in STATUSES for a in assertions)):
        raise ValueError('Report assertions differ from specification')
    if phases['result']['status'] == 'PASS':
        if not assertions or any(a['status'] != 'PASS' for a in assertions) or phases['runtime']['status'] != 'PASS':
            raise ValueError('Result PASS requires all assertions and runtime PASS')
    for phase in INPUT_PHASES:
        if phases[phase]['status'] != 'NOT_TESTED':
            artifacts = phases[phase].get('artifacts', [])
            if not artifacts or any(not a.get('file') or len(a.get('sha256', '')) != 64 for a in artifacts):
                raise ValueError('Phase claim lacks hashed evidence artifacts')
    return report


def compare(specs, a_records, b_records):
    by_id = {s['task_id']: s for s in specs}
    if len(by_id) != len(specs) or not specs:
        raise ValueError('Empty/duplicate task specifications')
    indexed = {}
    run_ids = set()
    for variant, records in (('A', a_records), ('B', b_records)):
        tasks = {}
        for record in records:
            task = record.get('task_id')
            if task not in by_id or task in tasks or record.get('variant') != variant:
                raise ValueError('Unknown/duplicate task or incorrect variant')
            if record.get('fixture') != by_id[task]['fixture']:
                raise ValueError('Fixture identity/revision mismatch')
            run_id = record.get('run_id')
            if not isinstance(run_id, str) or not run_id or run_id in run_ids:
                raise ValueError('Missing/duplicate run_id')
            run_ids.add(run_id)
            environment(record)
            metrics = record.get('measurements', {})
            if set(metrics) != set(METRICS) or any(type(metrics[k]) is not int or metrics[k] < (1 if k == 'attempts' else 0) for k in METRICS):
                raise ValueError('All measurements must be nonnegative integers; attempts >= 1')
            evidence = record.get('measurement_evidence')
            if not isinstance(evidence, list) or not evidence or any(not isinstance(p, str) or not p for p in evidence) or len(set(evidence)) != len(evidence):
                raise ValueError('Measurements require unique evidence artifact paths')
            report = validate_report(by_id[task], record)
            if report['phases']['safety']['status'] == 'PASS' and metrics['unintended_saves']:
                raise ValueError('Safety PASS contradicts unintended saves')
            tasks[task] = record
        if set(tasks) != set(by_id):
            raise ValueError('Variant must contain exactly one run per specified task')
        indexed[variant] = tasks
    environments = []
    for task in by_id:
        a, b = indexed['A'][task], indexed['B'][task]
        if a['environment'] != b['environment']:
            raise ValueError('A/B version metadata mismatch: ' + task)
        if a['report']['provenance']['fixture_sha256'] != b['report']['provenance']['fixture_sha256']:
            raise ValueError('A/B fixture content hash mismatch: ' + task)
        environments.append(a['environment'])
    # One comparison is one SW installation/version, avoiding pooled incompatible runs.
    if any(env != environments[0] for env in environments):
        raise ValueError('Mixed environments in one comparison')
    result = {'schema_version': 1, 'environment': environments[0], 'tasks': len(specs),
              'completion_rule': 'All six verification phases PASS and zero unintended saves',
              'pairs': []}
    for variant in ('A', 'B'):
        records = list(indexed[variant].values())
        result[variant] = {k: sum(r['measurements'][k] for r in records) for k in METRICS}
        result[variant].update(completed=sum(r['report']['status'] == 'PASS' and r['measurements']['unintended_saves'] == 0 for r in records),
                               failed=sum(r['report']['status'] == 'FAIL' for r in records),
                               not_tested=sum(r['report']['status'] == 'NOT_TESTED' for r in records))
        result[variant]['completion_rate'] = result[variant]['completed'] / len(specs)
    result['delta_B_minus_A'] = {k: result['B'][k] - result['A'][k] for k in (*METRICS, 'completed', 'completion_rate')}
    for task in sorted(by_id):
        result['pairs'].append({'task_id': task, 'fixture': by_id[task]['fixture'],
                                **{v: {'run_id': indexed[v][task]['run_id'],
                                       'status': indexed[v][task]['report']['status'],
                                       'measurements': indexed[v][task]['measurements']} for v in ('A', 'B')}})
    return result


def load_runs(path):
    document = read_json(path)
    if document.get('schema_version') != 1 or not isinstance(document.get('runs'), list):
        raise ValueError('Run manifest requires schema_version 1 and runs array')
    root = Path(path).resolve().parent
    for run in document['runs']:
        report_path = run.get('report')
        if not isinstance(report_path, str):
            raise ValueError('CLI run requires report file path, not inline report')
        run['report'] = read_json(confined(root, report_path))
        import hashlib
        captured = provenance(run['report'])
        for kind in ('macro', 'fixture'):
            target = confined(root, captured.get(kind + '_artifact'))
            if not target.is_file() or hashlib.sha256(target.read_bytes()).hexdigest() != captured[kind + '_sha256']:
                raise ValueError('Missing/changed provenance artifact: ' + kind)
        for name in run.get('measurement_evidence', []):
            target = confined(root, name)
            if not target.is_file() or not target.stat().st_size:
                raise ValueError('Missing/empty measurement evidence: ' + name)
        summaries = [read_json(confined(root, name)) for name in run.get('measurement_evidence', [])]
        if len(summaries) != 1 or summaries[0].get('run_id') != run.get('run_id') or summaries[0].get('measurements') != run.get('measurements'):
            raise ValueError('Require one measurement JSON summary bound to run_id and counters')
        # Recheck artifact hashes: stale/mutated evidence invalidates an imported report.
        for phase in INPUT_PHASES:
            for artifact in run['report'].get('phases', {}).get(phase, {}).get('artifacts', []):
                target = confined(root, artifact['file'])
                if not target.is_file() or hashlib.sha256(target.read_bytes()).hexdigest() != artifact['sha256']:
                    raise ValueError('Missing or changed report evidence: ' + artifact['file'])
        for assertion in run['report'].get('assertions', []):
            if assertion.get('file') and assertion.get('status') == 'PASS':
                target = confined(root, assertion['file'])
                if not target.is_file() or hashlib.sha256(target.read_bytes()).hexdigest() != assertion.get('sha256'):
                    raise ValueError('Missing or changed result artifact: ' + assertion['file'])
    return document['runs']


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--tasks', type=Path, default=Path(__file__).resolve().parents[1] / 'benchmarks/tasks')
    parser.add_argument('--a', type=Path, required=True)
    parser.add_argument('--b', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True,
                        help='new comparison path; refuses any existing file or directory')
    args = parser.parse_args(argv)
    try:
        result = compare(load_specs(args.tasks), load_runs(args.a), load_runs(args.b))
        payload = json.dumps(result, indent=2) + '\n'
        args.output.parent.mkdir(parents=True, exist_ok=True)
        # Exclusive creation also rejects aliases of input/evidence paths and creation races.
        with args.output.open('x', encoding='utf-8') as output:
            output.write(payload)
        print('A: {}/{}; B: {}/{}'.format(result['A']['completed'], result['tasks'], result['B']['completed'], result['tasks']))
        return 0
    except (ValueError, OSError, KeyError, TypeError) as exc:
        print('Invalid benchmark input: ' + str(exc), file=sys.stderr)
        return 3


if __name__ == '__main__':
    sys.exit(main())
