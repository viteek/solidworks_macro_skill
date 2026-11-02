import contextlib
import io
from pathlib import Path
import re
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import check_examples
import run_in_solidworks


class ExampleTests(unittest.TestCase):
    def test_extracts_indented_vba_fences_and_preserves_source_location(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'reference.md'
            path.write_text('# Reference\n\n  ```vb\n  Dim ok As Boolean\n  ```\n```python\ninvalid vba\n```\n')
            blocks = list(check_examples.vba_blocks(path))
            self.assertEqual(blocks, [(4, 'Dim ok As Boolean\n')])

    def test_fails_on_a_bad_boolean_in_documentation(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'reference.md'
            path.write_text('```vb\nDim model As SldWorks.ModelDoc2\n'
                            'Dim ok As Boolean\nok = model.GetSaveFlag\nIf Not ok Then Exit Sub\n```\n')
            with contextlib.redirect_stdout(io.StringIO()):
                self.assertGreater(check_examples.check_paths([path]), 0)

    def test_accepts_a_fragment_without_claiming_it_is_a_complete_macro(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'reference.md'
            path.write_text('```vb\nDim drawing As SldWorks.DrawingDoc\nDim names As Variant\n'
                            'names = drawing.GetSheetNames\n```\n')
            with contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(check_examples.check_paths([path]), 0)


class RecipeResultTests(unittest.TestCase):
    def test_every_bundled_negative_status_fails_a_text_expectation(self):
        root = Path(__file__).resolve().parents[1]
        statuses = set()
        for path in list((root / 'assets').rglob('*.bas')) + list((root / 'recipes').rglob('*.bas')):
            source = path.read_text(encoding='utf-8')
            statuses |= set(re.findall(r'"([A-Z][A-Z_]+)(?:[;: ]|")', source))
            statuses |= set(re.findall(r'"RESULT (?:\w+ )?([A-Z][A-Z_]+)[;: ]', source))
        self.assertTrue({'ERROR', 'ERROR_CLEANUP', 'FAIL', 'CLEANUP_FAILED', 'EXPORTED'} <= statuses)
        for status in sorted(statuses):
            with self.subTest(status=status):
                problem = run_in_solidworks.check_log_text(status + ';detail OK\n', 'OK')
                if status in run_in_solidworks.SUCCESS_STATUSES:
                    self.assertIsNone(problem)
                else:
                    self.assertIsNotNone(problem)

    def test_cleanup_failure_line_cannot_repeat_the_success_status(self):
        # Rebuild each recipe's cleanup-failure line from its VBA literals and run it
        # through the runner's own log checks.
        root = Path(__file__).resolve().parents[1] / 'recipes'
        failure = re.compile(r'resultText = "(FAIL (\w+) cleanup: )" & cleanupText & "; prior=" & '
                             r'Replace\(resultText, "([^"]*)", ""\)')
        found = 0
        for path in sorted(root.rglob('*.bas')):
            source = path.read_text(encoding='utf-8')
            for match in failure.finditer(source):
                found += 1
                head, name, removed = match.groups()
                success = re.search(r'resultText = "(RESULT ' + name + r' OK; [^"]*)"', source)
                self.assertIsNotNone(success, path)
                prior = success.group(1).replace(removed, '')
                line = head + 'restore failed; prior=' + prior + '1\n'
                self.assertNotIn(' OK', ' ' + prior, path)
                for needle in ('OK', name + ' OK', 'RESULT ' + name + ' OK'):
                    with self.subTest(recipe=name, needle=needle):
                        self.assertIsNotNone(run_in_solidworks.check_log_text(line, needle))
                        self.assertIsNotNone(run_in_solidworks.check_log_result(line, needle))
        self.assertEqual(found, len(list(root.rglob('*.bas'))) - len(
            [p for p in root.rglob('*.bas') if 'prior=' not in p.read_text(encoding='utf-8')]))
        self.assertGreater(found, 0)


if __name__ == '__main__':
    unittest.main()
