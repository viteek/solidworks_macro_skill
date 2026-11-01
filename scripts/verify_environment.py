#!/usr/bin/env python3
"""Report local validation capabilities without starting or connecting to SOLIDWORKS."""
import argparse
import importlib.util
import json
from pathlib import Path
import platform
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))
import check_api


def inspect_environment(sw_dir):
    folder = Path(sw_dir)
    libraries = []
    for name in ('sldworks.tlb', 'swconst.tlb', 'swcommands.tlb'):
        path = folder / name
        record = {'path': str(path), 'exists': path.is_file()}
        if path.is_file() and check_api.pythoncom is not None:
            try:
                library = check_api.pythoncom.LoadTypeLib(str(path))
                guid, lcid, syskind, major, minor, flags = library.GetLibAttr()
                record.update(status='READABLE', guid=str(guid), lcid=lcid,
                              major=major, minor=minor, type_count=library.GetTypeInfoCount())
            except Exception as error:
                record.update(status='UNREADABLE', error=str(error))
        libraries.append(record)
    available = all(item.get('status') == 'READABLE' for item in libraries[:2])
    return {'schema_version': 1, 'platform': platform.platform(), 'python': platform.python_version(),
            'python_executable': sys.executable,
            'modules': {name: importlib.util.find_spec(name) is not None
                        for name in ('pythoncom', 'win32com', 'ezdxf', 'shapely', 'OCP')},
            'type_libraries': libraries, 'api_status': 'AVAILABLE' if available else 'UNAVAILABLE',
            'compilation_status': 'NOT_TESTED', 'live_status': 'NOT_PROBED',
            'note': 'Type-library versions are observed; release/SP/UI language require integration evidence.'}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--sw-dir', default=check_api.default_sw_dir())
    parser.add_argument('--json', type=Path)
    parser.add_argument('--require-api', action='store_true')
    args = parser.parse_args()
    report = inspect_environment(args.sw_dir)
    out = json.dumps(report, indent=2, ensure_ascii=False)
    print(out)
    if args.json:
        args.json.write_text(out, encoding='utf-8')
    return 1 if args.require_api and report['api_status'] != 'AVAILABLE' else 0


if __name__ == '__main__':
    raise SystemExit(main())
