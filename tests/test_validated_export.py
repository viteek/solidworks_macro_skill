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
import validated_export


class PublicationTests(unittest.TestCase):
    def test_step_export_validation_uses_killable_timeout_worker(self):
        with patch('validate_step.run_isolated', return_value={'status': 'PASS'}) as isolated:
            result = validated_export.validate_step_export('model.step', {'bodies': []}, {}, 23)
        self.assertEqual(result['status'], 'PASS')
        isolated.assert_called_once_with('model.step', {'bodies': []}, {}, timeout=23)

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

    def test_hash_mismatch_after_publication_is_not_reported_as_success(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            stage, target = root / 'stage.dxf', root / 'target.dxf'
            stage.write_bytes(b'geometry')
            digest = hashlib.sha256(b'geometry').hexdigest()
            real_hash = validated_export.file_hash
            with patch.object(validated_export, 'file_hash', side_effect=lambda path: 'wrong' if path.name == target.name else real_hash(path)):
                with self.assertRaisesRegex(ValueError, 'Published output hash'):
                    publish_validated(stage, target, lambda _: {'status': 'PASS', 'sha256': digest})
            self.assertTrue(stage.exists())
            self.assertTrue(target.exists())  # retained for inspection, not blindly deleted
            self.assertFalse((root / 'target.dxf.publication-lock').exists())

    def test_existing_target_cannot_be_overwritten(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            stage, target = root / 'stage.dxf', root / 'target.dxf'
            stage.write_bytes(b'new')
            target.write_bytes(b'original')
            with self.assertRaisesRegex(ValueError, 'Overwrite.*disabled'):
                publish_validated(stage, target, lambda _: self.fail('validator called'), overwrite=True)
            self.assertEqual(target.read_bytes(), b'original')
            self.assertEqual(list(root.glob('*.previous-*')), [])

    def test_cooperative_lock_refuses_concurrent_publication(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            stage, target = root / 'stage.dxf', root / 'target.dxf'
            stage.write_bytes(b'geometry')
            lock = root / 'target.dxf.publication-lock'
            lock.write_text('busy')
            with self.assertRaises(FileExistsError):
                publish_validated(stage, target, lambda _: self.fail('validator called'))
            self.assertEqual(lock.read_text(), 'busy')
            self.assertFalse(target.exists())

    def test_target_created_during_validation_is_not_replaced(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            stage, target = root / 'stage.dxf', root / 'target.dxf'
            stage.write_bytes(b'geometry')
            def validator(path):
                target.write_bytes(b'concurrent')
                return {'status': 'PASS', 'sha256': hashlib.sha256(path.read_bytes()).hexdigest()}
            with self.assertRaises(FileExistsError):
                publish_validated(stage, target, validator, overwrite=True)
            self.assertEqual(target.read_bytes(), b'concurrent')
            self.assertTrue(stage.exists())

    def test_failed_validation_does_not_publish(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            stage, target = root / 'stage.dxf', root / 'target.dxf'
            stage.write_bytes(b'geometry')
            result = publish_validated(stage, target, lambda _: {'status': 'FAIL'})
            self.assertEqual(result['status'], 'BLOCKED_VALIDATION')
            self.assertFalse(target.exists())
            self.assertTrue(stage.exists())


if __name__ == '__main__':
    unittest.main()
