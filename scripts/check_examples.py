#!/usr/bin/env python3
"""Lint every .bas asset and fenced VBA example; optionally check installed SW APIs.

Fragments are linted as fragments: they do not establish compilation or runtime
correctness, and undeclared external context cannot be type-checked.
"""
import argparse
from pathlib import Path
import re
import sys
import tempfile
import textwrap

sys.path.insert(0, str(Path(__file__).resolve().parent))
import check_api
import check_vba

ROOT = Path(__file__).resolve().parents[1]
FRAGMENT_CONTEXT = {
    'swApp': 'SldWorks', 'swModel': 'ModelDoc2', 'swPart': 'PartDoc',
    'swDraw': 'DrawingDoc', 'swDrw': 'ModelDoc2', 'swFeat': 'Feature',
    'swSub': 'Feature', 'swSelMgr': 'SelectionMgr', 'swComp': 'Component2',
    'swView': 'View', 'swModelView': 'ModelView'}


def vba_blocks(path):
    text = Path(path).read_text(encoding='utf-8')
    pattern = r'^([ \t]*)```(?:vb|vba)[ \t]*\n(.*?)^\1```[ \t]*$'
    for match in re.finditer(pattern, text, re.M | re.S | re.I):
        yield text.count('\n', 0, match.start(2)) + 1, textwrap.dedent(match.group(2))


def check_paths(paths, index=None):
    errors = 0
    count = 0
    with tempfile.TemporaryDirectory(prefix='sw-examples-') as directory:
        for path in paths:
            path = Path(path)
            blocks = vba_blocks(path) if path.suffix.lower() == '.md' else [(1, path.read_text(encoding='utf-8'))]
            for line, code in blocks:
                count += 1
                print(f'Checking {path}:{line}')
                if path.suffix.lower() == '.md':
                    # The reference fragments use these documented external names.
                    # This supplies analysis context, not a runnable macro wrapper.
                    context = []
                    for name, kind in FRAGMENT_CONTEXT.items():
                        if re.search(r'\b' + name + r'\b', code, re.I) and not re.search(r'\b' + name + r'\s+As\b', code, re.I):
                            context.append(f'Dim {name} As SldWorks.{kind}')
                    if context:
                        print('Fragment context: ' + ', '.join(context))
                        code = '\n'.join(context) + '\n' + code
                extracted = Path(directory) / (path.stem + '-' + str(line) + '.bas')
                extracted.write_text(code, encoding='utf-8')
                errors += check_vba.check(extracted, snippet=path.suffix.lower() == '.md')
                if index is not None:
                    errors += check_api.check(str(extracted), index)
    print(f'Checked {count} VBA assets/fragments; {errors} error(s). '
          + ('Installed API check enabled.' if index is not None else 'L1 only; API types/compilation/execution not verified.'))
    return errors


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('paths', nargs='*', type=Path)
    parser.add_argument('--api', action='store_true', help='also use installed SOLIDWORKS type libraries')
    parser.add_argument('--sw-dir', default=check_api.default_sw_dir())
    args = parser.parse_args()
    paths = args.paths or (sorted((ROOT / 'assets').rglob('*.bas')) + sorted((ROOT / 'recipes').rglob('*.bas'))
                          + sorted((ROOT / 'references').glob('*.md')) + sorted((ROOT / 'recipes').rglob('*.md'))
                          + [ROOT / 'SKILL.md'])
    index = check_api.TypeLibIndex(args.sw_dir) if args.api else None
    return 1 if check_paths(paths, index) else 0


if __name__ == '__main__':
    raise SystemExit(main())
