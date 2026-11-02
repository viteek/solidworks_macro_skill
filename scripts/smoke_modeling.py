#!/usr/bin/env python3
"""Prepare or explicitly run nine modeling/selection smoke cases on private copies.

Requires a saved blank part and an untouched saved Recipe_Block (40x20x10 mm).
Defaults to preparation only. --run drives the already running, empty SOLIDWORKS
through the reviewed runner. Whole-project compilation and A/B scores are NOT tested.
"""
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
CASES = {
    'sketch': ('modeling/sketch-extrude/sketch_extrude.bas', 'blank', 'sketch_extrude'),
    'modify': ('modeling/feature-modification/modify_extrude.bas', 'block', 'modify_extrude'),
    'configurations': ('modeling/configuration-variants/configuration_variants.bas', 'block', 'configuration_variants'),
    'hole_linear': ('modeling/hole-linear-pattern/hole_linear_pattern.bas', 'block', 'hole_linear_pattern'),
    'revolve': ('modeling/revolve/revolve_ring.bas', 'blank', 'revolve_ring'),
    'circular': ('modeling/circular-pattern/circular_pattern.bas', 'blank', 'circular_pattern'),
    'sketch3d': ('modeling/sketch-3d/sketch_3d.bas', 'blank', 'sketch_3d'),
    'plane': ('selection/plane-selection/plane_selection.bas', 'blank', 'plane_selection'),
    'persistent': ('selection/plane-selection/plane_selection.bas', 'blank', 'plane_selection')}
VOLUMES = {'sketch': 8e-6, 'modify': 12e-6, 'configurations': 8e-6,
           'hole_linear': 8e-6 - 3 * math.pi * .002 ** 2 * .01,
           'revolve': math.pi * (.02 ** 2 - .01 ** 2) * .01,
           'circular': math.pi * (.02 ** 2 - 4 * .002 ** 2) * .01}
FEATURES = {'sketch': 'Recipe_Block', 'modify': 'Recipe_Block',
            'hole_linear': 'Recipe_ThreeHoles', 'revolve': 'Recipe_Ring',
            'circular': 'Recipe_FourDiskHoles', 'sketch3d': 'Recipe_3DPath'}


def assert_result(case, observation):
    failures = []
    if observation.get('collection_status') != 'OK':
        return ['Independent observer did not collect state']
    if case in VOLUMES:
        value = observation.get('volume_m3')
        if (type(value) not in (int, float) or not math.isfinite(value)
                or not math.isclose(value, VOLUMES[case], rel_tol=1e-7, abs_tol=1e-11)):
            failures.append('Volume does not match expected geometry')
        if observation.get('solid_body_count') != 1:
            failures.append('Expected exactly one solid body')
    features = observation.get('features', [])
    feature = next((item for item in features if item.get('name') == FEATURES.get(case)), None)
    if case in FEATURES and feature is None:
        failures.append('Expected feature missing: ' + FEATURES[case])
    if case == 'configurations':
        names = observation.get('configurations', [])
        if len(names) != 3 or not {'Recipe_10mm', 'Recipe_20mm'}.issubset(names):
            failures.append('Configuration variants missing')
    if case == 'sketch3d' and feature:
        if feature.get('segment_count') != 3 or not math.isclose(feature.get('length_m', -1), .06, abs_tol=1e-9):
            failures.append('3D polyline readback mismatch')
    if case == 'plane' and observation.get('selection_count') != 1:
        failures.append('Expected one selected plane')
    return failures


def run_case(case, root, sources, live, timeout):
    recipe, fixture, label = CASES[case]
    folder = root / case
    folder.mkdir()
    source = folder / (case + '.sldprt')
    shutil.copy2(sources[fixture], source)
    macro = ROOT / 'recipes' / recipe
    observer = (ROOT / 'recipes/selection/persistent-reference/persistent_roundtrip.bas'
                if case == 'persistent' else ROOT / 'assets/diagnostic-macros/model_state_snapshot.bas')
    log, observed, runner_report = folder / 'macro.log', folder / 'observed.json', folder / 'runner.json'
    defines = ['RESULT_PATH="' + str(log) + '"']
    obs_defines = ['RESULT_PATH="' + str(observed) + '"']
    profile = {'schema_version': 1, 'test_root': str(folder), 'reviewed': live,
               'copies_confirmed': True, 'allow_model_changes': True,
               'entry_point': {'module': 'TestMacro', 'procedure': 'main'},
               'macro_sha256': code_hash(prepare_code(macro.read_text(encoding='utf-8'), defines)),
               'observer_sha256': code_hash(prepare_code(observer.read_text(encoding='utf-8'), obs_defines)),
               'inputs': [{'path': str(source), 'sha256': file_hash(source)}],
               'outputs': [str(p) for p in (log, observed, runner_report)], 'allowed_dialogs': []}
    profile_path = folder / 'profile.json'
    profile_path.write_text(json.dumps(profile, indent=2), encoding='utf-8')
    command = [sys.executable, '-B', str(ROOT / 'scripts/run_in_solidworks.py'), str(macro),
               '--profile', str(profile_path), '--open', str(source), '--confirm-model-changes',
               '--define', defines[0], '--observe', str(observer), '--observe-define', obs_defines[0],
               '--close-opened', '--timeout', str(timeout), '--json', str(runner_report),
               '--expect-result', str(log), 'RESULT ' + label + ' OK']
    if case == 'persistent':
        command += ['--expect-result', str(observed), 'RESULT persistent_roundtrip OK']
    else:
        command += ['--expect-file', str(observed)]
    (folder / 'command.json').write_text(json.dumps(command, indent=2), encoding='utf-8')
    if not live:
        return {'case': case, 'status': 'NOT_TESTED', 'profile': str(profile_path)}
    try:
        process = subprocess.run(command, capture_output=True, timeout=timeout + 30)
    except subprocess.TimeoutExpired:
        return {'case': case, 'status': 'UNKNOWN', 'reason': 'Controller exceeded timeout; inspect durable runner session'}
    (folder / 'controller.log').write_bytes(process.stdout + process.stderr)
    report = json.loads(runner_report.read_text(encoding='utf-8')) if runner_report.exists() else {}
    failures = []
    if process.returncode != 0 or report.get('status') != 'PASS':
        failures.append('Runner failed; inspect runner.json and macro.log')
    observation = None
    if case != 'persistent' and observed.exists():
        try:
            observation = json.loads(observed.read_text(encoding='utf-8'))
            failures.extend(assert_result(case, observation))
        except (ValueError, TypeError, KeyError) as error:
            failures.append('Invalid observation: ' + str(error))
    if file_hash(source) != profile['inputs'][0]['sha256']:
        failures.append('Input copy unexpectedly saved')
    return {'case': case, 'status': 'FAIL' if failures else 'PASS', 'failures': failures,
            'runner_status': report.get('status'), 'observation': observation,
            'macro_sha256': profile['macro_sha256'], 'observer_sha256': profile['observer_sha256'],
            'session_state': report.get('session_state'), 'compilation_status': 'NOT_TESTED'}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--blank', type=Path, required=True)
    parser.add_argument('--block', type=Path, required=True)
    parser.add_argument('--test-root', type=Path, required=True, help='NEW disposable directory')
    parser.add_argument('--run', action='store_true', help='explicitly execute reviewed recipes on copied inputs')
    parser.add_argument('--case', choices=list(CASES), action='append')
    parser.add_argument('--timeout', type=float, default=45)
    args = parser.parse_args()
    sources = {'blank': args.blank.resolve(strict=True), 'block': args.block.resolve(strict=True)}
    hashes = {kind: file_hash(path) for kind, path in sources.items()}
    root = args.test_root.resolve()
    root.mkdir(parents=True, exist_ok=False)
    results = []
    for case in args.case or CASES:
        result = run_case(case, root, sources, args.run, args.timeout)
        results.append(result)
        print(case + ': ' + result['status'], flush=True)
        if result.get('session_state') == 'UNKNOWN' or result['status'] == 'UNKNOWN':
            break  # no automatic recovery or retry of uncertain execution
    unchanged = all(file_hash(sources[kind]) == digest for kind, digest in hashes.items())
    report = {'schema_version': 1, 'scope': 'modeling smoke only, not full project compilation or A/B benchmark',
              'source_hashes': hashes, 'sources_unchanged': unchanged, 'results': results}
    (root / 'smoke-report.json').write_text(json.dumps(report, indent=2), encoding='utf-8')
    return 0 if unchanged and all(result['status'] in ('PASS', 'NOT_TESTED') for result in results) else 1


if __name__ == '__main__':
    raise SystemExit(main())
