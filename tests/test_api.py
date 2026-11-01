import contextlib
import io
from pathlib import Path
import sys
import tempfile
import unittest
import os

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import check_api


class ApiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        # Small documented contract fixture tests the checker, not SW compatibility.
        cls.index = object.__new__(check_api.TypeLibIndex)
        def member(name, nargs=0, ret=None):
            return [{'name': name, 'nargs': nargs, 'nopt': 0, 'kind': 1, 'ret': ret}]
        cls.index.types = {
            'modeldoc2': ('ModelDoc2', {
                'extension': member('Extension', ret='ModelDocExtension'),
                'gettitle': member('GetTitle'), 'getsaveflag': member('GetSaveFlag'),
                'clearselection2': member('ClearSelection2', 1)}),
            'modeldocextension': ('ModelDocExtension', {'saveas': member('SaveAs', 6)}),
            'drawingdoc': ('DrawingDoc', {'getsheetnames': member('GetSheetNames'),
                                        'activatesheet': member('ActivateSheet', 1)}),
            'body2': ('Body2', {'getextremepoint': member('GetExtremePoint', 6)})}
        cls.index.consts, cls.index.enums = {}, {}
        cls.index.owners = {'getsheetnames': {'DrawingDoc'}}

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

    def test_member_on_wrong_interface_is_an_error(self):
        errors, output = self.check_statement('Dim names As Variant\nnames = model.GetSheetNames')
        self.assertGreater(errors, 0, output)
        self.assertIn('DrawingDoc', output)

    def test_drawing_interface_accepts_sheet_names(self):
        errors, output = self.check_statement('Dim drawing As SldWorks.DrawingDoc\n'
            'Dim names As Variant\nSet drawing = model\nnames = drawing.GetSheetNames')
        self.assertEqual(errors, 0, output)

    def test_negative_arguments_and_identifier_prefixes_are_not_operators(self):
        errors, output = self.check_statement('Dim body As SldWorks.Body2\nDim x As Double\n'
            'Dim y As Double\nDim z As Double\nDim drawing As SldWorks.DrawingDoc\nDim originalSheet As String\n'
            'body.GetExtremePoint -1, 0, 0, x, y, z\ndrawing.ActivateSheet originalSheet')
        self.assertEqual(errors, 0, output)


@unittest.skipUnless(check_api.pythoncom is not None and
                     os.path.isfile(os.path.join(check_api.default_sw_dir(), 'sldworks.tlb')),
                     'Installed SOLIDWORKS type libraries are unavailable')
class InstalledApiTests(ApiTests):
    @classmethod
    def setUpClass(cls):
        cls.index = check_api.TypeLibIndex(check_api.default_sw_dir())


if __name__ == '__main__':
    unittest.main()
