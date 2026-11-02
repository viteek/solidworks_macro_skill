import itertools
from pathlib import Path
import sys
import unittest
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import validate_dxf


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
                with patch.object(validate_dxf, 'file_hash', return_value='hash'), \
                        patch.object(validate_dxf.ezdxf, 'readfile', return_value=document):
                    report = validate_dxf.validate_dxf('unused.dxf', profile)
                codes = [issue['code'] for issue in report['issues']]
                self.assertEqual('RESOURCE_LIMIT' in codes, total > validate_dxf.MAX_ENTITIES)
                self.assertEqual(report['ignored_entities']['POINT'], validate_dxf.MAX_ENTITIES)
                self.assertEqual(entity.dxftype.call_count, validate_dxf.MAX_ENTITIES)
                self.assertEqual(report['status'], 'FAIL')


if __name__ == '__main__':
    unittest.main()
