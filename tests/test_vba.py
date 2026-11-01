import contextlib
import io
from pathlib import Path
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import check_vba


class BooleanTests(unittest.TestCase):
    def check_code(self, declarations, statements):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'macro.bas'
            path.write_text('Option Explicit\nSub main()\n' + declarations + '\n' + statements + '\nEnd Sub\n')
            with contextlib.redirect_stdout(io.StringIO()) as output:
                errors = check_vba.check(path)
            return errors, output.getvalue()

    def test_filesystem_boolean_is_not_a_solidworks_result(self):
        errors, output = self.check_code('Dim fso As Object\nDim exists As Boolean',
            'Set fso = CreateObject("Scripting.FileSystemObject")\nexists = fso.FileExists("x")\nIf Not exists Then Exit Sub')
        self.assertEqual(errors, 0, output)

    def test_typed_com_result_requires_explicit_comparison(self):
        for condition in ('Not ok', 'ok', 'ok = True', 'ok <> True'):
            with self.subTest(condition=condition):
                errors, output = self.check_code('Dim model As SldWorks.ModelDoc2\nDim ok As Boolean',
                    'ok = model.GetSaveFlag\nIf ' + condition + ' Then Exit Sub')
                self.assertGreater(errors, 0, output)

    def test_nonzero_comparisons_and_normalized_values_are_accepted(self):
        for statement in ('ok = model.GetSaveFlag\nIf ok = False Then Exit Sub',
                          'ok = model.GetSaveFlag\nIf ok <> False Then Exit Sub',
                          'ok = (model.GetSaveFlag <> False)\nIf Not ok Then Exit Sub'):
            with self.subTest(statement=statement):
                errors, output = self.check_code('Dim model As SldWorks.ModelDoc2\nDim ok As Boolean', statement)
                self.assertEqual(errors, 0, output)

    def test_numeric_api_result_is_not_a_com_boolean(self):
        errors, output = self.check_code('Dim model As SldWorks.ModelDoc2\nDim count As Long',
            'count = model.GetType\nIf count = 1 Then Exit Sub')
        self.assertEqual(errors, 0, output)

    def test_numeric_return_coerced_to_boolean_is_standard_vba(self):
        errors, output = self.check_code('Dim model As SldWorks.ModelDoc2\nDim ok As Boolean',
            'ok = model.GetType\nIf Not ok Then Exit Sub')
        self.assertEqual(errors, 0, output)

    def test_object_alias_can_become_a_filesystem_object(self):
        errors, output = self.check_code('Dim model As SldWorks.ModelDoc2\nDim fso As Object\nDim ok As Boolean',
            'Set fso = model.Extension\nSet fso = CreateObject("Scripting.FileSystemObject")\n'
            'ok = fso.FileExists("x")\nIf Not ok Then Exit Sub')
        self.assertEqual(errors, 0, output)

    def test_declared_interface_is_preserved_after_object_reassignment(self):
        errors, output = self.check_code('Dim app As SldWorks.SldWorks\nDim model As SldWorks.ModelDoc2',
            'Set model = app.ActiveDoc\nSet model = Nothing\nSet model = GetObject("fixture")\n'
            'If Not model.GetSaveFlag Then Exit Sub')
        self.assertGreater(errors, 0, output)

    def test_unsafe_negation_on_assignment_rhs_is_rejected(self):
        errors, output = self.check_code('Dim model As SldWorks.ModelDoc2\nDim failed As Boolean',
            'failed = Not model.Save3(0, errors, warnings)\nIf failed Then Exit Sub')
        self.assertGreater(errors, 0, output)

    def test_parenthesized_false_comparison_is_accepted(self):
        errors, output = self.check_code('Dim model As SldWorks.ModelDoc2',
            'If model.GetSaveFlag = (False) Then Exit Sub')
        self.assertEqual(errors, 0, output)

    def test_circle_graphics_keyword_cannot_be_a_variable(self):
        errors, output = self.check_code('Dim circle As SldWorks.SketchSegment', '')
        self.assertGreater(errors, 0, output)

    def test_bodyfolder_getbodies_is_current_and_partdoc_getbodies_is_obsolete(self):
        _, output = self.check_code('Dim folder As SldWorks.BodyFolder\nDim bodies As Variant',
                                    'bodies = folder.GetBodies')
        self.assertNotIn('Obsolete API', output)
        _, output = self.check_code('Dim part As SldWorks.PartDoc\nDim bodies As Variant',
                                    'bodies = part.GetBodies')
        self.assertIn('GetBodies2', output)

    def test_aliases_continuations_and_inline_assignments_are_tracked(self):
        errors, output = self.check_code('Dim model As SldWorks.ModelDoc2\nDim aliasModel As Object\nDim ok As Boolean\nDim copied As Boolean',
            'Set aliasModel = model\nok = _\n aliasModel.GetSaveFlag\ncopied = ok: If Not copied Then Exit Sub')
        self.assertGreater(errors, 0, output)

    def test_local_boolean_does_not_inherit_another_procedures_type(self):
        errors, output = self.check_code('Dim model As SldWorks.ModelDoc2\nDim ok As Boolean',
            'ok = model.GetSaveFlag\nIf ok = False Then Exit Sub\nEnd Sub\n'
            'Private Sub CheckFile()\nDim fso As Object\nDim ok As Boolean\n'
            'Set fso = CreateObject("Scripting.FileSystemObject")\nok = fso.FileExists("x")\nIf Not ok Then Exit Sub')
        self.assertEqual(errors, 0, output)


if __name__ == '__main__':
    unittest.main()
