#!/usr/bin/env python3
"""Run one unbent sheet-metal smoke case on an explicitly copied blank fixture."""
import argparse
import json
import math
from pathlib import Path
import shutil
import subprocess
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))
from runner_safety import code_hash, file_hash
from run_in_solidworks import prepare_code

ROOT = Path(__file__).resolve().parents[1]


def assert_metrics(observation):
    """Independent geometry checks; malformed or non-finite values cannot pass."""
    failures = []
    if not isinstance(observation, dict) or observation.get('units') != 'mm':
        return ['Expected independent metrics in millimeters']
    bodies = observation.get('bodies', []) if isinstance(observation, dict) else []
    if not isinstance(bodies, list) or len(bodies) != 1 or not isinstance(bodies[0], dict):
        return ['Expected exactly one observed solid body']
    volume = bodies[0].get('volume_mm3')
    numeric = lambda value: type(value) in (float, int) and math.isfinite(value)
    if not numeric(volume) or not math.isclose(volume, 10000, abs_tol=.01):
        failures.append('Independent body volume mismatch')
    bounds = bodies[0].get('bounds_mm')
    if not isinstance(bounds, list) or len(bounds) != 6 or not all(numeric(value) for value in bounds):
        return failures + ['Invalid independent bounds']
    sizes = sorted(bounds[index+3] - bounds[index] for index in range(3))
    if any(not math.isclose(actual, expected, abs_tol=1e-5) for actual, expected in zip(sizes, [2, 50, 100])):
        failures.append('Independent dimensions mismatch')
    return failures


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--blank', required=True, type=Path)
    parser.add_argument('--test-root', required=True, type=Path)
    parser.add_argument('--run', action='store_true')
    args = parser.parse_args()
    blank = args.blank.resolve(strict=True)
    original_hash = file_hash(blank)
    folder = args.test_root.resolve()
    folder.mkdir(parents=True, exist_ok=False)
    copied = folder / 'blank-copy.sldprt'
    shutil.copy2(blank, copied)
    macro = ROOT / 'recipes/sheetmetal/create-flat-blank/create_flat_blank.bas'
    metrics_source = ROOT / 'assets/templates/source_metrics_template.bas'
    inspection_source = ROOT / 'recipes/sheetmetal/inspect_sheetmetal_cutlist.bas'
    save_source = ROOT / 'assets/diagnostic-macros/save_fixture_copy.bas'
    sources = [macro, metrics_source, inspection_source, save_source]
    hashes = {str(path): file_hash(path) for path in sources}
    metrics = metrics_source.read_text(encoding='utf-8').replace('Sub main()', 'Private Sub CollectMetrics()', 1)
    inspection = inspection_source.read_text(encoding='utf-8').replace('Public Sub main()', 'Private Sub InspectFixture()', 1).replace('RESULT_PATH', 'INSPECTION_PATH')
    save = save_source.read_text(encoding='utf-8').replace('Sub main()', 'Private Sub SaveFixture()', 1)
    observer = folder / 'observer.bas'
    lines = [line for part in (metrics, inspection, save) for line in part.splitlines()
             if not line.startswith('Attribute ') and line.strip() != 'Option Explicit']
    declarations = [line for line in lines if line.startswith(('Const ', 'Private Const ', 'Public Const '))]
    procedures = [line for line in lines if line not in declarations]
    code = 'Option Explicit\n' + '\n'.join(declarations + procedures) + '\nSub main()\n    CollectMetrics\n    InspectFixture\n    SaveFixture\nEnd Sub\n'
    observer.write_text(code, encoding='utf-8')
    outputs = {key: folder / filename for key, filename in {
        'macro': 'macro.log', 'metrics': 'metrics.json', 'inspection': 'inspection.log',
        'save': 'save.log', 'fixture': 'flat-blank.sldprt', 'runner': 'runner.json'}.items()}
    literal = lambda value: '"' + str(value).replace('"', '""') + '"'
    defines = ['RESULT_PATH=' + literal(outputs['macro'])]
    obs_defines = ['METRICS_PATH=' + literal(outputs['metrics']), 'INSPECTION_PATH=' + literal(outputs['inspection']),
                   'RESULT_PATH=' + literal(outputs['save']), 'EXPECTED_INPUT_PATH=' + literal(copied),
                   'OUTPUT_PATH=' + literal(outputs['fixture']), 'SHOW_MESSAGES=False']
    profile = {'schema_version': 1, 'test_root': str(folder), 'reviewed': args.run,
               'copies_confirmed': True, 'allow_model_changes': True,
               'entry_point': {'module': 'TestMacro', 'procedure': 'main'},
               'macro_sha256': code_hash(prepare_code(macro.read_text(encoding='utf-8'), defines)),
               'observer_sha256': code_hash(prepare_code(code, obs_defines)),
               'inputs': [{'path': str(copied), 'sha256': file_hash(copied)}],
               'outputs': [str(path) for path in outputs.values()], 'allowed_dialogs': []}
    profile_path = folder / 'profile.json'
    profile_path.write_text(json.dumps(profile, indent=2), encoding='utf-8')
    command = [sys.executable, '-B', str(ROOT / 'scripts/run_in_solidworks.py'), str(macro), '--profile', str(profile_path),
               '--open', str(copied), '--confirm-model-changes', '--close-opened', '--timeout', '45',
               '--json', str(outputs['runner']), '--define', defines[0], '--observe', str(observer)]
    for item in obs_defines:
        command += ['--observe-define', item]
    for key in ('metrics', 'inspection', 'fixture'):
        command += ['--expect-file', str(outputs[key])]
    command += ['--expect-text', str(outputs['macro']), 'RESULT create_flat_blank OK',
                '--expect-text', str(outputs['save']), 'RESULT=PASS']
    (folder / 'command.json').write_text(json.dumps(command, indent=2), encoding='utf-8')
    if not args.run:
        print('Prepared only:', folder)
        return 0
    try:
        process = subprocess.run(command, capture_output=True, timeout=80)
    except subprocess.TimeoutExpired:
        # No COM queries, cleanup, retry, or later cases after uncertain execution.
        unknown = {'status': 'UNKNOWN', 'reason': 'Controller exceeded timeout; inspect durable runner session',
                   'runner_report': str(outputs['runner'])}
        (folder / 'evidence.json').write_text(json.dumps(unknown, indent=2), encoding='utf-8')
        print('UNKNOWN: stop; inspect durable runner session:', folder)
        return 2
    (folder / 'controller.log').write_bytes(process.stdout + process.stderr)
    report = json.loads(outputs['runner'].read_text(encoding='utf-8'))
    failures = []
    if report.get('session_state') == 'UNKNOWN':
        print('UNKNOWN: stop; inspect runner session:', outputs['runner'])
        return 2
    if process.returncode != 0 or report.get('status') != 'PASS':
        failures.append('Runner did not pass')
    observation = json.loads(outputs['metrics'].read_text(encoding='utf-8')) if outputs['metrics'].exists() else {}
    failures.extend(assert_metrics(observation))
    inspection_text = outputs['inspection'].read_text(encoding='utf-16') if outputs['inspection'].exists() else ''
    if 'RESULT inspect_sheetmetal_cutlist PASS' not in inspection_text:
        failures.append('Positive sheet-metal inspector did not report PASS')
    unchanged = file_hash(blank) == original_hash and file_hash(copied) == profile['inputs'][0]['sha256']
    unchanged = unchanged and all(file_hash(Path(path)) == digest for path, digest in hashes.items())
    if not unchanged:
        failures.append('Source artifact/input copy changed')
    evidence = {'status': 'FAIL' if failures else 'PASS', 'failures': failures, 'source_sha256': hashes,
                'blank_sha256': original_hash, 'sources_unchanged': unchanged, 'profile': profile,
                'runner': report, 'metrics': observation, 'inspection_result': inspection_text.strip(),
                'outputs': {key: str(path) for key, path in outputs.items()},
                'project_compilation': 'NOT_TESTED', 'scope': 'unbent base flange only; no bend/export certification'}
    (folder / 'evidence.json').write_text(json.dumps(evidence, indent=2), encoding='utf-8')
    print(evidence['status'], folder)
    return 1 if failures else 0


if __name__ == '__main__':
    raise SystemExit(main())
