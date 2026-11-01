#!/usr/bin/env python3
"""Opt-in validated export of a SOLIDWORKS part: stage -> geometry validation -> publication.

1. Runs assets/templates/safe_export_template.bas (+ source_metrics_template.bas for STEP) in the
   running SOLIDWORKS through run_in_solidworks.py, writing into a private job folder.
2. DXF: optional explicit unit normalization to mm (--normalize-units), then validate_dxf.py with a
   profile (profiles/solidworks-sheetmetal-mm.json for SOLIDWORKS flat patterns).
   STEP: validate_step.py compares the OpenCascade import with SOLIDWORKS source metrics.
3. Only a PASSED file, byte-identical to the validated one, is published to the target. An existing
   target is kept as a .previous-<id> backup (with --overwrite) - never deleted first.

Use on COPIES of parts. No CAM/machine is configured: PASS = declared geometric checks passed.

Usage:
  python validated_export.py part.sldprt out.dxf --profile ../profiles/solidworks-sheetmetal-mm.json --normalize-units --json report.json
  python validated_export.py part.sldprt out.step --tolerances ../profiles/step-tolerances.json --json report.json
"""
import argparse
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import time
import uuid

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))  # sibling imports also under python -I
from validation_reports import file_hash, write_report  # noqa: E402
from runner_safety import code_hash
from run_in_solidworks import prepare_code

PACKAGE = HERE.parent


def publish_validated(stage, target, validator, overwrite=False):
    stage, target = Path(stage).resolve(), Path(target).resolve()
    if stage == target or stage.is_symlink() or not stage.is_file() or stage.stat().st_size == 0:
        raise ValueError('Publication requires a distinct nonempty regular staging file')
    if target.exists() and not overwrite:
        raise FileExistsError('Final output already exists (use --overwrite to keep it as a backup and replace)')
    before = file_hash(stage)
    validation = validator(stage)
    if validation.get('status') != 'PASS':
        return {'status': 'BLOCKED_VALIDATION', 'validation': validation, 'published': False}
    if validation.get('sha256') != before or file_hash(stage) != before:
        raise ValueError('Validated byte identity changed before publication')
    backup = None
    if target.exists():
        original_hash = file_hash(target)
        backup = target.with_name(target.name + '.previous-' + uuid.uuid4().hex[:8])
        with target.open('rb') as source, backup.open('xb') as copy:
            shutil.copyfileobj(source, copy)
        if file_hash(target) != original_hash or file_hash(backup) != original_hash:
            raise ValueError('Previous output changed while making its backup')
        os.replace(stage, target)
    else:
        os.link(stage, target)  # exclusive creation, even if the target appeared during validation
        stage.unlink()
    return {'status': 'PUBLISHED', 'published': True, 'sha256': before,
            'backup': str(backup) if backup else None, 'validation': validation}


def build_validation_macro(with_metrics):
    export = (PACKAGE / 'assets/templates/safe_export_template.bas').read_text(encoding='utf-8')
    if not with_metrics:
        return export
    metrics = (PACKAGE / 'assets/templates/source_metrics_template.bas').read_text(encoding='utf-8')
    anchor = '    ExportToStage swApp, model, stagedPath, extension, message'
    if export.count(anchor) != 1 or metrics.count('Private Sub WriteSourceMetrics') != 1:
        raise ValueError('Macro template changed; review the metric capture hook before proceeding')
    export = export.replace('Option Explicit', 'Option Explicit\nConst METRICS_PATH As String = ""', 1)
    export = export.replace(anchor, '    WriteSourceMetrics model, METRICS_PATH\n' + anchor)
    return export + '\n' + metrics[metrics.index('Private Sub WriteSourceMetrics'):]


def vba_string(path):
    return '"' + str(path).replace('"', '""') + '"'


def execute(args):
    source, target = Path(args.source).resolve(strict=True), Path(args.target).resolve()
    kind = target.suffix.lower()
    if source.suffix.lower() != '.sldprt':
        raise ValueError('Validated pipeline accepts saved native parts (.sldprt) only')
    if kind not in ('.step', '.stp', '.dxf'):
        raise ValueError('Validated pipeline supports STEP and DXF')
    if kind == '.dxf' and not args.profile:
        raise ValueError('DXF requires --profile')
    if kind in ('.step', '.stp') and not args.tolerances:
        raise ValueError('STEP requires --tolerances')
    if target.exists() and not args.overwrite:
        raise FileExistsError('Final target exists; --overwrite keeps it as a backup')
    if not target.parent.is_dir():
        raise ValueError('Target directory does not exist')
    source_hash = file_hash(source)
    job = Path(tempfile.mkdtemp(prefix='.sw-validation-', dir=target.parent))
    test_source = job / source.name
    shutil.copy2(source, test_source)
    with_metrics = kind != '.dxf'
    macro = job / 'job.bas'
    macro.write_text(build_validation_macro(with_metrics), encoding='utf-8')
    raw, metrics, log, runner_report = job / ('raw' + kind), job / 'source-metrics.json', job / 'macro.log', job / 'runner.json'
    command = [sys.executable, '-I', str(HERE / 'run_in_solidworks.py'), str(macro),
               '--open', str(test_source), '--close-opened', '--timeout', str(args.timeout), '--json', str(runner_report),
               '--define', 'SHOW_MESSAGES=False', '--define', 'OUTPUT_PATH=' + vba_string(raw),
               '--define', 'RESULT_PATH=' + vba_string(log)]
    if with_metrics:
        command += ['--define', 'METRICS_PATH=' + vba_string(metrics)]
    if args.configuration:
        command += ['--define', 'CONFIGURATION_NAME=' + vba_string(args.configuration)]
    defines = [command[index + 1] for index, value in enumerate(command) if value == '--define']
    prepared = prepare_code(macro.read_text(encoding='utf-8'), defines)
    test_profile = job / 'test-profile.json'
    # This pipeline owns the generated exporter and its private input/output paths.
    # External references are refused by the runner; they are never copied blindly.
    test_profile.write_text(json.dumps({
        'schema_version': 1, 'test_root': str(job), 'reviewed': True, 'copies_confirmed': True,
        'allow_model_changes': False, 'macro_sha256': code_hash(prepared),
        'entry_point': {'module': 'TestMacro', 'procedure': 'main'},
        'inputs': [{'path': str(test_source), 'sha256': file_hash(test_source)}],
        'outputs': [str(path) for path in (raw, metrics, log, runner_report)], 'allowed_dialogs': []
    }, indent=2), encoding='utf-8')
    command += ['--profile', str(test_profile)]
    started = time.time()
    subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=args.timeout + 60)
    report = {'status': 'BLOCKED_EXPORT', 'cam_status': 'NOT_CONFIGURED', 'source_sha256': source_hash,
              'job_folder': str(job), 'target': str(target)}
    report['runner'] = json.loads(runner_report.read_text(encoding='utf-8')) if runner_report.exists() else {}
    log_text = log.read_text(encoding='cp1250', errors='replace') if log.exists() else ''
    report['macro_log'] = log_text.strip()
    report['source_unchanged'] = file_hash(source) == source_hash
    fresh = raw.exists() and raw.stat().st_mtime >= started - 2
    if report['runner'].get('status') != 'PASS' or 'EXPORTED;' not in log_text or not fresh:
        return report
    if not report['source_unchanged']:
        raise ValueError('Source file changed during export; final target was not touched')
    candidate = raw
    if kind == '.dxf':
        sys.path.insert(0, str(HERE))
        from validate_dxf import normalize_dxf, validate_dxf
        profile = json.loads(Path(args.profile).read_text(encoding='utf-8'))
        if args.normalize_units:
            candidate = job / 'normalized.dxf'
            normalized = normalize_dxf(raw, candidate, {}, profile)
            report['normalization'] = normalized.get('normalization')
            if normalized['status'] != 'PASS':
                report.update(status='BLOCKED_VALIDATION', validation=normalized)
                return report
        validator = lambda path: validate_dxf(path, profile)  # noqa: E731
    else:
        from validate_step import validate_step
        source_metrics = json.loads(metrics.read_text(encoding='ascii'))
        if os.path.normcase(str(Path(source_metrics['source_path']).resolve())) != os.path.normcase(str(test_source)):
            raise ValueError('Source metric identity does not match requested part')
        if args.configuration and source_metrics['configuration'] != args.configuration:
            raise ValueError('Metrics came from a different source configuration')
        report['source_configuration'] = source_metrics['configuration']
        tolerances = json.loads(Path(args.tolerances).read_text(encoding='utf-8'))
        validator = lambda path: validate_step(path, source_metrics, tolerances)  # noqa: E731
    report.update(publish_validated(candidate, target, validator, args.overwrite))
    if report.get('status') == 'PUBLISHED':
        shutil.rmtree(job, ignore_errors=True)   # keep the job folder only when something needs inspection
        report['job_folder'] = None
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('source')
    parser.add_argument('target')
    parser.add_argument('--configuration')
    parser.add_argument('--profile', help='DXF validation profile (JSON)')
    parser.add_argument('--normalize-units', action='store_true', help='DXF: convert declared units to mm in a new copy first')
    parser.add_argument('--tolerances', help='STEP tolerances (JSON)')
    parser.add_argument('--overwrite', action='store_true')
    parser.add_argument('--timeout', type=float, default=300)
    parser.add_argument('--json', required=True)
    args = parser.parse_args()
    report_file = Path(args.json).resolve()
    inputs = [args.source, args.target] + [v for v in (args.profile, args.tolerances) if v]
    if report_file.suffix.lower() != '.json' or report_file in [Path(v).resolve() for v in inputs]:
        parser.error('Report must be a distinct .json file')
    try:
        report = execute(args)
    except Exception as error:  # noqa: BLE001
        report = {'status': 'ERROR', 'cam_status': 'NOT_CONFIGURED', 'exception': str(error)}
    write_report(report_file, report)
    print(report['status'], 'CAM:', report['cam_status'])
    return 0 if report['status'] == 'PUBLISHED' else 1


if __name__ == '__main__':
    raise SystemExit(main())
