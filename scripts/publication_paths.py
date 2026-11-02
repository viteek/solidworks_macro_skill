"""Shared final-component checks for publication paths.

Parent directories must be trusted against concurrent modification.
"""
from pathlib import Path


def publication_target(target):
    target = Path(target)
    if target.is_symlink():
        raise ValueError('Final output must not be a symlink')
    target = target.parent.resolve() / target.name
    if target.is_symlink():
        raise ValueError('Final output must not be a symlink')
    return target
