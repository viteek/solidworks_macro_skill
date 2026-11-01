import contextlib
import io
from pathlib import Path
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import check_examples


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


if __name__ == '__main__':
    unittest.main()
