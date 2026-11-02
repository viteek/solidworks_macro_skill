"""Offline regression checks for staging publication."""
import hashlib
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from validated_export import execute, publish_validated


class PublicationTests(unittest.TestCase):
    def test_simulated_target_symlink_rejected_with_or_without_destination(self):
        for dangling in (False, True):
            for overwrite in (False, True):
                with self.subTest(dangling=dangling, overwrite=overwrite), tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    stage, target = root / 'stage.dxf', root / 'target.dxf'
                    stage.write_bytes(b'geometry')
                    if not dangling:
                        target.write_bytes(b'original')
                    with patch.object(Path, 'is_symlink', autospec=True, side_effect=lambda path: path == target):
                        with self.assertRaisesRegex(ValueError, 'symlink'):
                            publish_validated(stage, target, lambda _: self.fail('validator called'), overwrite)
                    self.assertEqual(stage.read_bytes(), b'geometry')
                    self.assertEqual(target.exists(), not dangling)
                    if not dangling:
                        self.assertEqual(target.read_bytes(), b'original')

    def test_target_symlink_rejected_for_both_overwrite_modes(self):
        for dangling in (False, True):
            for overwrite in (False, True):
                with self.subTest(dangling=dangling, overwrite=overwrite), tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    stage, target, destination = root / 'stage.dxf', root / 'target.dxf', root / 'destination.dxf'
                    stage.write_bytes(b'geometry')
                    if not dangling:
                        destination.write_bytes(b'original')
                    try:
                        target.symlink_to(destination)
                    except (OSError, NotImplementedError) as error:
                        self.skipTest(f'Symlinks unavailable: {error}')
                    with self.assertRaisesRegex(ValueError, 'symlink'):
                        publish_validated(stage, target, lambda _: self.fail('validator called'), overwrite)
                    self.assertTrue(target.is_symlink())
                    self.assertEqual(stage.read_bytes(), b'geometry')
                    if dangling:
                        self.assertFalse(destination.exists())
                    else:
                        self.assertEqual(destination.read_bytes(), b'original')
                    self.assertEqual(list(root.glob('*.previous-*')), [])

    def test_execute_rejects_target_symlink_before_export(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source, target = root / 'source.sldprt', root / 'target.dxf'
            source.write_bytes(b'part')
            try:
                target.symlink_to(root / 'missing.dxf')
            except (OSError, NotImplementedError) as error:
                self.skipTest(f'Symlinks unavailable: {error}')
            with self.assertRaisesRegex(ValueError, 'symlink'):
                execute(SimpleNamespace(source=source, target=target))
            self.assertEqual(sorted(path.name for path in root.iterdir()), ['source.sldprt', 'target.dxf'])

    def test_target_symlink_created_during_validation_is_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            stage, target, destination = root / 'stage.dxf', root / 'target.dxf', root / 'destination.dxf'
            stage.write_bytes(b'geometry')
            destination.write_bytes(b'original')

            def validator(path):
                try:
                    target.symlink_to(destination)
                except (OSError, NotImplementedError) as error:
                    self.skipTest(f'Symlinks unavailable: {error}')
                return {'status': 'PASS', 'sha256': hashlib.sha256(path.read_bytes()).hexdigest()}

            with self.assertRaisesRegex(ValueError, 'symlink'):
                publish_validated(stage, target, validator, overwrite=True)
            self.assertTrue(target.is_symlink())
            self.assertEqual(destination.read_bytes(), b'original')
            self.assertTrue(stage.exists())

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
