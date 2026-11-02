from pathlib import Path
import shutil
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
import check_skill


class SkillStructureTests(unittest.TestCase):
    def test_repository_skill_is_valid(self):
        self.assertEqual(check_skill.check_skill(ROOT), [])

    def check_variant(self, edit):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            shutil.copy(ROOT / 'CHANGELOG.md', root / 'CHANGELOG.md')
            (root / 'references').mkdir()
            (root / 'references' / 'pitfalls.md').write_text('x')
            text = ('---\nname: solidworks-macros\ndescription: Automate SOLIDWORKS.\n'
                    'metadata:\n  version: "' + check_skill.frontmatter(
                        (ROOT / 'SKILL.md').read_text(encoding='utf-8'))[0]['metadata']['version'] + '"\n'
                    '---\n\nSee `references/pitfalls.md`.\n')
            (root / 'SKILL.md').write_text(edit(text), encoding='utf-8')
            return check_skill.check_skill(root)

    def test_minimal_valid_variant_passes(self):
        self.assertEqual(self.check_variant(lambda text: text), [])

    def test_invalid_structure_is_reported(self):
        cases = {
            'missing frontmatter': lambda t: t.replace('---\n', '', 1),
            'invalid yaml': lambda t: t.replace('name: solidworks-macros', 'name: [unclosed'),
            'bad name': lambda t: t.replace('name: solidworks-macros', 'name: SolidWorks Macros'),
            'missing description': lambda t: t.replace('description: Automate SOLIDWORKS.\n', ''),
            'long description': lambda t: t.replace('Automate SOLIDWORKS.', 'x' * 1025),
            'angle brackets': lambda t: t.replace('Automate SOLIDWORKS.', 'Automate <SW>.'),
            'unknown key': lambda t: t.replace('metadata:', 'owner: me\nmetadata:'),
            'version drift': lambda t: t.replace('version: "', 'version: "9.9.9-'),
            'missing path': lambda t: t.replace('references/pitfalls.md', 'references/missing.md'),
        }
        for label, edit in cases.items():
            with self.subTest(label):
                self.assertNotEqual(self.check_variant(edit), [])


if __name__ == '__main__':
    unittest.main()
