import itertools
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import validate_dxf


class NormalizationPublicationTests(unittest.TestCase):
    def test_simulated_symlink_rejected_before_validation_or_parsing(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'target.dxf'
            with patch.object(Path, 'is_symlink', autospec=True, side_effect=lambda path: path == target), \
                    patch.object(validate_dxf, 'validate_profile') as profile_check, \
                    patch.object(validate_dxf.ezdxf, 'readfile') as readfile:
                with self.assertRaisesRegex(ValueError, 'symlink'):
                    validate_dxf.normalize_dxf(Path(directory) / 'source.dxf', target, {}, {})
            profile_check.assert_not_called()
            readfile.assert_not_called()

    def test_dangling_external_symlink_rejected_without_creating_destination(self):
        with tempfile.TemporaryDirectory() as directory, tempfile.TemporaryDirectory() as external:
            target = Path(directory) / 'target.dxf'
            destination = Path(external) / 'missing.dxf'
            try:
                target.symlink_to(destination)
            except OSError as error:
                if getattr(error, 'winerror', None) == 1314:
                    self.skipTest(f'Symlink privilege unavailable: {error}')
                raise
            with patch.object(validate_dxf, 'validate_profile') as profile_check, \
                    patch.object(validate_dxf.ezdxf, 'readfile') as readfile:
                with self.assertRaisesRegex(ValueError, 'symlink'):
                    validate_dxf.normalize_dxf(Path(directory) / 'source.dxf', target, {}, {})
            profile_check.assert_not_called()
            readfile.assert_not_called()
            self.assertTrue(target.is_symlink())
            self.assertFalse(destination.exists())

    def test_regular_target_publication_and_existing_target_rejected(self):
        profile = {'schema_version': 1, 'name': 'test', 'units': 'mm',
                   'layers': {'CUT': 'cut'}, 'curve_tolerance_mm': 0.1,
                   'topology_tolerance_mm': 0.01, 'allowed_entities': ['LWPOLYLINE'], 'expected': {}}
        with tempfile.TemporaryDirectory() as directory:
            source, target = Path(directory) / 'source.dxf', Path(directory) / 'target.dxf'
            document = validate_dxf.ezdxf.new()
            document.units = 4
            document.modelspace().add_lwpolyline([(0, 0), (10, 0), (10, 10), (0, 10)], close=True,
                                                 dxfattribs={'layer': 'CUT'})
            document.saveas(source)
            result = validate_dxf.normalize_dxf(source, target, {}, profile)
            self.assertEqual(result['status'], 'PASS')
            self.assertTrue(result['normalization']['published'])
            original = target.read_bytes()
            with self.assertRaises(FileExistsError):
                validate_dxf.normalize_dxf(source, target, {}, profile)
            self.assertEqual(target.read_bytes(), original)
            with self.assertRaises(FileExistsError):
                validate_dxf.normalize_dxf(source, source, {}, profile)

    def test_target_appearing_during_validation_is_not_replaced(self):
        profile = {'schema_version': 1, 'name': 'test', 'units': 'mm',
                   'layers': {'CUT': 'cut'}, 'curve_tolerance_mm': 0.1,
                   'topology_tolerance_mm': 0.01, 'allowed_entities': ['LINE'], 'expected': {}}
        with tempfile.TemporaryDirectory() as directory:
            source, target = Path(directory) / 'source.dxf', Path(directory) / 'target.dxf'
            document = validate_dxf.ezdxf.new()
            document.units = 4
            document.saveas(source)

            def validate(stage, unused_profile):
                target.write_bytes(b'concurrent')
                return {'status': 'PASS'}

            with patch.object(validate_dxf, 'validate_dxf', side_effect=validate):
                with self.assertRaises(FileExistsError):
                    validate_dxf.normalize_dxf(source, target, {}, profile)
            self.assertEqual(target.read_bytes(), b'concurrent')


class EntityLimitTests(unittest.TestCase):
    def test_ignored_entities_count_towards_limit(self):
        profile = {
            'schema_version': 1, 'name': 'test', 'units': 'mm',
            'layers': {'CUT': 'cut'}, 'curve_tolerance_mm': 0.1,
            'topology_tolerance_mm': 0.01, 'allowed_entities': ['LINE'],
            'ignored_entities': ['POINT'], 'expected': {},
        }
        for total in (validate_dxf.MAX_ENTITIES, validate_dxf.MAX_ENTITIES + 1):
            with self.subTest(total=total):
                entity = Mock()
                entity.dxftype.return_value = 'POINT'
                entity.dxf.layer = 'CUT'
                document = Mock()
                document.units = 4
                document.audit.return_value.has_errors = False
                document.audit.return_value.has_fixes = False
                document.modelspace.return_value = itertools.repeat(entity, total)
                with patch.object(validate_dxf, 'check_file_size'), \
                        patch.object(validate_dxf, 'file_hash', return_value='hash'), \
                        patch.object(validate_dxf.ezdxf, 'readfile', return_value=document):
                    report = validate_dxf.validate_dxf('unused.dxf', profile)
                codes = [issue['code'] for issue in report['issues']]
                self.assertEqual('RESOURCE_LIMIT' in codes, total > validate_dxf.MAX_ENTITIES)
                self.assertEqual(report['ignored_entities']['POINT'], validate_dxf.MAX_ENTITIES)
                self.assertEqual(entity.dxftype.call_count, validate_dxf.MAX_ENTITIES)
                self.assertEqual(report['status'], 'FAIL')

    def test_oversized_input_is_rejected_before_parser_for_both_operations(self):
        profile = {
            'schema_version': 1, 'name': 'test', 'units': 'mm',
            'layers': {'CUT': 'cut'}, 'curve_tolerance_mm': 0.1,
            'topology_tolerance_mm': 0.01, 'allowed_entities': ['LINE'], 'expected': {},
        }
        with patch.object(Path, 'stat', return_value=Mock(st_size=validate_dxf.MAX_FILE_BYTES + 1)), \
                patch.object(Path, 'exists', return_value=False), \
                patch.object(Path, 'is_symlink', return_value=False), \
                patch.object(validate_dxf, 'file_hash') as file_hash, \
                patch.object(validate_dxf.ezdxf, 'readfile') as readfile:
            for operation in (
                lambda: validate_dxf.validate_dxf('source.dxf', profile),
                lambda: validate_dxf.normalize_dxf('source.dxf', 'target.dxf', {}, profile),
            ):
                with self.subTest(operation=operation), self.assertRaisesRegex(ValueError, 'RESOURCE_LIMIT'):
                    operation()
            file_hash.assert_not_called()
            readfile.assert_not_called()

    def test_file_size_limit_is_inclusive(self):
        with patch.object(Path, 'stat', return_value=Mock(st_size=validate_dxf.MAX_FILE_BYTES)):
            validate_dxf.check_file_size(Path('source.dxf'))


if __name__ == '__main__':
    unittest.main()
