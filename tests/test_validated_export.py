"""Offline regression checks for staging publication."""
import hashlib
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from validated_export import publish_validated


class PublicationTests(unittest.TestCase):
    def test_stage_symlink_check_precedes_resolution(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            stage = root / 'stage.dxf'
            stage.write_bytes(b'geometry')
            with patch.object(Path, 'is_symlink', return_value=True):
                with self.assertRaises(ValueError):
                    publish_validated(stage, root / 'target.dxf', lambda _: self.fail('validator called'))

    def test_symlink_stage_is_rejected_without_publication(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / 'source.dxf'
            source.write_bytes(b'geometry')
            stage = root / 'stage.dxf'
            try:
                stage.symlink_to(source)
            except (OSError, NotImplementedError) as error:
                self.skipTest(f'Symlinks unavailable: {error}')
            target = root / 'target.dxf'
            with self.assertRaises(ValueError):
                publish_validated(stage, target, lambda _: self.fail('validator called'))
            self.assertFalse(target.exists())
            self.assertEqual(source.read_bytes(), b'geometry')

    def test_regular_stage_is_published(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            stage, target = root / 'stage.dxf', root / 'target.dxf'
            stage.write_bytes(b'geometry')
            digest = hashlib.sha256(b'geometry').hexdigest()
            result = publish_validated(stage, target, lambda _: {'status': 'PASS', 'sha256': digest})
            self.assertEqual(result['status'], 'PUBLISHED')
            self.assertEqual(target.read_bytes(), b'geometry')


if __name__ == '__main__':
    unittest.main()
