import contextlib
import io
from pathlib import Path
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import check_api


class ApiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        # Read the real installed type libraries; never start SOLIDWORKS.
        cls.index = check_api.TypeLibIndex(check_api.default_sw_dir())

    def check_statement(self, statement):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'macro.bas'
            path.write_text('Option Explicit\nSub main()\n'
                            'Dim model As SldWorks.ModelDoc2\nDim ok As Boolean\nDim title As String\n'
                            'Dim errCode As Long\nDim warnCode As Long\n'
                            + statement + '\nEnd Sub\n')
            with contextlib.redirect_stdout(io.StringIO()) as output:
                errors = check_api.check(str(path), self.index)
            return errors, output.getvalue()

    def test_required_arguments_cannot_be_omitted(self):
        for statement in ('ok = model.Extension.SaveAs', 'model.Extension.SaveAs',
                          'If model.Extension.SaveAs Then ok = False',
                          'If ok Then model.Extension.SaveAs',
                          'ok = model.Extension.SaveAs()'):
            with self.subTest(statement=statement):
                errors, output = self.check_statement(statement)
                self.assertGreater(errors, 0, output)
                self.assertIn('needs at least 6 argument(s)', output)

    def test_valid_properties_and_calls_remain_accepted(self):
        for statement in ('title = model.GetTitle', 'ok = model.GetSaveFlag',
                          'ok = model.Extension.SaveAs("out.step", 0, 3, Nothing, errCode, warnCode)',
                          'If ok Then model.ClearSelection2 True',
                          'If ok Then title = model.GetTitle Else model.ClearSelection2 False',
                          'model.Extension.SaveAs "out.step", 0, 3, Nothing, errCode, warnCode'):
            with self.subTest(statement=statement):
                errors, output = self.check_statement(statement)
                self.assertEqual(errors, 0, output)

    def test_inline_with_calls_keep_their_arguments(self):
        errors, output = self.check_statement(
            'With model\nIf ok Then .ClearSelection2 True Else .ClearSelection2 False\nEnd With')
        self.assertEqual(errors, 0, output)

    def test_inline_else_arguments_belong_to_their_own_call(self):
        for statement in (
            'If ok Then model.ClearSelection2 True Else model.Extension.SaveAs "b.step", 0, 3, Nothing, errCode, warnCode',
            'If ok Then model.Extension.SaveAs "a.step", 0, 3, Nothing, errCode, warnCode Else model.Extension.SaveAs "b.step", 0, 3, Nothing, errCode, warnCode',
        ):
            with self.subTest(statement=statement):
                errors, output = self.check_statement(statement)
                self.assertEqual(errors, 0, output)


if __name__ == '__main__':
    unittest.main()
