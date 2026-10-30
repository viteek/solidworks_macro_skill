"""Shared helpers for the geometry validators: hashing, number checks and fresh JSON reports
that never overwrite CAD inputs or validation profiles. No third-party dependencies."""
import hashlib
import json
import math
import os
from pathlib import Path
import tempfile


def file_hash(path):
    digest = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def positive(value, name):
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or value <= 0:
        raise ValueError(name + ' must be finite and positive')
    return value


def report_path(path, protected):
    path = Path(path).resolve()
    if path.suffix.lower() != '.json':
        raise ValueError('Report must use a .json filename')
    for source in protected:
        if source is None:
            continue
        source = Path(source).resolve()
        if path == source or (path.exists() and source.exists() and os.path.samefile(path, source)):
            raise ValueError('Report cannot overwrite an input, profile or CAD output')
    return path


def write_report(path, value):
    path = Path(path)
    descriptor, temporary = tempfile.mkstemp(prefix='.validation-report-', suffix='.json', dir=path.parent)
    try:
        with os.fdopen(descriptor, 'w', encoding='utf-8') as stream:
            json.dump(value, stream, indent=2, ensure_ascii=False, allow_nan=False)
        os.replace(temporary, path)
    finally:
        Path(temporary).unlink(missing_ok=True)
