import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import validation_reports


class ReportPathTests(unittest.TestCase):
    def test_rejects_existing_file_without_mutation(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'report.json'
            path.write_bytes(b'existing content\n')
            with self.assertRaisesRegex(ValueError, 'already exists'):
                validation_reports.report_path(path, [])
            self.assertEqual(path.read_bytes(), b'existing content\n')
            self.assertEqual(list(Path(folder).iterdir()), [path])

    def test_rejects_existing_directory_without_mutation(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'report.json'
            path.mkdir()
            with self.assertRaisesRegex(ValueError, 'already exists'):
                validation_reports.report_path(path, [])
            self.assertTrue(path.is_dir())
            self.assertEqual(list(path.iterdir()), [])

    def test_rejects_symlinks_including_dangling_without_mutation(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            target = root / 'target'
            target.write_bytes(b'original')
            for name, destination in [('normal.json', target), ('dangling.json', root / 'missing')]:
                path = root / name
                try:
                    path.symlink_to(destination)
                except (OSError, NotImplementedError) as error:
                    self.skipTest('Symlinks unavailable: ' + str(error))
                with self.subTest(name=name):
                    with self.assertRaisesRegex(ValueError, 'already exists'):
                        validation_reports.report_path(path, [])
                    self.assertTrue(path.is_symlink())
                    self.assertEqual(path.readlink(), destination)
            self.assertEqual(target.read_bytes(), b'original')
            self.assertFalse((root / 'missing').exists())

    def test_rejects_protected_future_path_and_invalid_suffix(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            path = root / 'report.json'
            with self.assertRaisesRegex(ValueError, 'cannot overwrite'):
                validation_reports.report_path(path, [root / '.' / 'report.json'])
            with self.assertRaisesRegex(ValueError, '.json'):
                validation_reports.report_path(root / 'report.txt', [])
            self.assertEqual(list(root.iterdir()), [])

    def test_accepts_new_json_path(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'report.json'
            self.assertEqual(validation_reports.report_path(path, []), path.resolve())
            self.assertFalse(path.exists())

    def test_rejects_missing_parent_before_any_side_effect(self):
        with tempfile.TemporaryDirectory() as folder:
            parent = Path(folder) / 'missing'
            with self.assertRaisesRegex(ValueError, 'parent directory must already exist'):
                validation_reports.report_path(parent / 'report.json', [])
            self.assertFalse(parent.exists())


class WriteReportTests(unittest.TestCase):
    def test_writes_complete_json_to_new_path(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'report.json'
            value = {'status': 'PASS', 'message': 'Zażółć'}
            validation_reports.write_report(path, value)
            self.assertEqual(json.loads(path.read_text(encoding='utf-8')), value)
            self.assertEqual(list(Path(folder).iterdir()), [path])

    def test_existing_output_is_unchanged_and_temporary_is_removed(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'report.json'
            original = b'existing content\n'
            path.write_bytes(original)
            with self.assertRaises(FileExistsError):
                validation_reports.write_report(path, {'status': 'FAIL'})
            self.assertEqual(path.read_bytes(), original)
            self.assertEqual(list(Path(folder).iterdir()), [path])

    def test_temporary_is_removed_when_publication_fails(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'report.json'
            with patch.object(validation_reports.os, 'link', side_effect=OSError('link failed')):
                with self.assertRaisesRegex(OSError, 'link failed'):
                    validation_reports.write_report(path, {'status': 'PASS'})
            self.assertEqual(list(Path(folder).iterdir()), [])

    def test_temporary_is_removed_when_serialization_fails(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'report.json'
            with self.assertRaises(ValueError):
                validation_reports.write_report(path, {'value': float('nan')})
            self.assertEqual(list(Path(folder).iterdir()), [])
