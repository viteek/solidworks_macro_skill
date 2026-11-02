#!/usr/bin/env python3
"""Validate SKILL.md structure: frontmatter, metadata and the repository paths it routes to.

This is a static packaging check. It does not show that an agent activates or follows the skill.
"""
import argparse
from pathlib import Path
import re
import sys

import yaml

ROOT = Path(__file__).resolve().parents[1]
ALLOWED_KEYS = {'name', 'description', 'license', 'allowed-tools', 'metadata', 'compatibility'}
PATH_PREFIXES = ('references/', 'recipes/', 'scripts/', 'assets/', 'profiles/', 'benchmarks/', 'tests/')


def frontmatter(text):
    match = re.match(r'---\r?\n(.*?)\r?\n---\r?\n', text, re.DOTALL)
    if not match:
        raise ValueError('SKILL.md must start with a --- delimited YAML frontmatter block')
    data = yaml.safe_load(match.group(1))
    if not isinstance(data, dict):
        raise ValueError('Frontmatter must be a YAML mapping')
    return data, text[match.end():]


def check_skill(root=ROOT):
    """Return a list of problems; an empty list means the structure is valid."""
    skill = Path(root) / 'SKILL.md'
    try:
        data, body = frontmatter(skill.read_text(encoding='utf-8'))
    except (OSError, UnicodeDecodeError, ValueError, yaml.YAMLError) as error:
        return [str(error)]
    problems = []
    unexpected = sorted(set(data) - ALLOWED_KEYS)
    if unexpected:
        problems.append(f'Unexpected frontmatter keys: {", ".join(unexpected)}')
    name = data.get('name')
    if not isinstance(name, str) or not re.fullmatch(r'[a-z0-9]+(-[a-z0-9]+)*', name) or len(name) > 64:
        problems.append(f'name must be kebab-case, at most 64 characters: {name!r}')
    description = data.get('description')
    if not isinstance(description, str) or not description.strip():
        problems.append('description must be a non-empty string')
    else:
        if len(description) > 1024:
            problems.append(f'description has {len(description)} characters; the limit is 1024')
        if '<' in description or '>' in description:
            problems.append('description must not contain angle brackets')
    metadata = data.get('metadata', {})
    version = metadata.get('version') if isinstance(metadata, dict) else None
    if not isinstance(version, str) or not version:
        problems.append('metadata.version must be a non-empty string')
    else:
        changelog = (Path(root) / 'CHANGELOG.md').read_text(encoding='utf-8')
        released = re.findall(r'^## (\d+\.\d+\.\d+\S*)', changelog, re.MULTILINE)
        if not released or released[0] != version:
            problems.append(f'metadata.version {version!r} is not the newest CHANGELOG version {released[:1]}')
    references = set(re.findall(r'\]\(([^)#\s]+)', body))
    references |= {p for p in re.findall(r'`([^`\s]+)`', body) if p.startswith(PATH_PREFIXES)}
    for reference in sorted(references):
        if re.match(r'[a-z]+:', reference):
            continue
        if not (Path(root) / reference).exists():
            problems.append(f'SKILL.md references a missing path: {reference}')
    return problems


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('root', nargs='?', default=str(ROOT))
    args = parser.parse_args()
    problems = check_skill(args.root)
    for problem in problems:
        print('ERROR', problem)
    print('Skill structure: ' + ('INVALID' if problems else 'OK'))
    return 1 if problems else 0


if __name__ == '__main__':
    raise SystemExit(main())
