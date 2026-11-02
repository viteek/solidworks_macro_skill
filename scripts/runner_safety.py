"""Offline preflight and durable execution state. This is not a VBA sandbox."""
import hashlib
import json
import os
import stat
from pathlib import Path
import time


def code_hash(code):
    payload = code if isinstance(code, bytes) else code.replace('\r\n', '\n').encode('utf-8')
    return hashlib.sha256(payload).hexdigest()


def file_hash(path):
    digest = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(block)
    return digest.hexdigest()


def canonical(path):
    return os.path.normcase(str(Path(path).resolve()))


def inside(path, root):
    try:
        return os.path.commonpath([canonical(path), canonical(root)]) == canonical(root) and canonical(path) != canonical(root)
    except ValueError:
        return False


def output_identity(path, root):
    """Validate the lexical path and capture identities of existing components."""
    path = Path(os.path.abspath(path))
    root = Path(root)
    if not inside(path, root):
        raise ValueError(f'Output outside test_root: {path}')
    try:
        relative = path.relative_to(root)
    except ValueError as error:
        raise ValueError(f'Output outside test_root: {path}') from error
    identities = {}
    current = Path(root.anchor)
    components = root.parts[1:] + relative.parts
    for part in (None, *components):
        if part is not None:
            current = current / part
        try:
            info = current.lstat()
        except FileNotFoundError:
            continue
        if stat.S_ISLNK(info.st_mode) or getattr(info, 'st_file_attributes', 0) & 0x400:
            raise ValueError(f'Output path contains a link or reparse point: {current}')
        if current == path and not stat.S_ISREG(info.st_mode):
            raise ValueError(f'Output path is not a regular file: {current}')
        # Directory identity must remain stable. The final regular output file may
        # legitimately be atomically replaced by the macro, so track it by path,
        # reject links, but do not require its inode to remain unchanged.
        if current != path:
            identities[str(current)] = (info.st_dev, info.st_ino)
    return canonical(path), identities


def revalidate_output(path, root, before):
    location, identities = output_identity(path, root)
    if location != before[0] or any(identities.get(p) != identity for p, identity in before[1].items()):
        raise ValueError(f'Output path or identity changed since preflight: {path}')


class TestProfile:
    def __init__(self, data):
        self.data = data
        self.root = Path(data['test_root']).resolve(strict=True)
        if not self.root.is_dir():
            raise ValueError('test_root must be an existing directory')
        self.inputs = {canonical(item['path']): item['sha256'] for item in data['inputs']}
        self.outputs = {canonical(path) for path in data['outputs']}
        if not self.inputs or len(self.inputs) != len(data['inputs']):
            raise ValueError('Profile needs distinct input copies')
        if self.outputs.intersection(self.inputs):
            raise ValueError('Outputs must be distinct from input CAD copies')
        for path in [item['path'] for item in data['inputs']] + data['outputs']:
            if not Path(path).is_absolute() or not inside(path, self.root):
                raise ValueError(f'Profile path must be absolute and inside test_root: {path}')
        self.check_input_hashes()

    @classmethod
    def load(cls, path, code):
        data = json.loads(Path(path).read_text(encoding='utf-8'))
        required = {'schema_version', 'test_root', 'reviewed', 'copies_confirmed',
                    'allow_model_changes', 'macro_sha256', 'inputs', 'outputs', 'allowed_dialogs', 'entry_point'}
        if not isinstance(data, dict) or not required.issubset(data):
            raise ValueError('Incomplete test profile')
        if data['schema_version'] != 1 or data['reviewed'] is not True or data['copies_confirmed'] is not True:
            raise ValueError('Review the macro, paths and copies before marking the profile reviewed')
        if type(data['allow_model_changes']) is not bool:
            raise ValueError('allow_model_changes must be a Boolean')
        if not Path(data['test_root']).is_absolute():
            raise ValueError('test_root must be absolute')
        if data['macro_sha256'] != code_hash(code):
            raise ValueError('Prepared macro changed after profile review (including --define / --pick-folder)')
        for dialog in data['allowed_dialogs']:
            if set(dialog) != {'title', 'text', 'answer'} or dialog['answer'] not in ('ok', 'yes', 'no', 'cancel'):
                raise ValueError('Each allowed dialog needs exact title, text and answer')
            if not all(isinstance(dialog[key], str) and dialog[key] for key in dialog):
                raise ValueError('Empty dialog allowlist fields are forbidden')
        return cls(data)

    def validate_entry(self, module, procedure):
        if self.data['entry_point'] != {'module': module, 'procedure': procedure}:
            raise ValueError('Macro entry point differs from the reviewed profile')

    def check_input_hashes(self):
        for path, expected in self.inputs.items():
            if not inside(path, self.root) or not Path(path).is_file() or file_hash(path) != expected:
                raise ValueError(f'Input copy missing, moved outside test_root, or changed since review: {path}')

    def validate_request(self, opened, outputs, folders, confirmed_changes):
        if not opened or canonical(opened) not in self.inputs:
            raise ValueError('--open must name an input copy declared in the profile')
        for path in outputs:
            if not inside(path, self.root) or canonical(path) not in self.outputs:
                raise ValueError(f'Undeclared output or output outside test_root: {path}')
        for path in folders:
            if canonical(path) != canonical(self.root) and not inside(path, self.root):
                raise ValueError(f'Folder picker outside test_root: {path}')
        if self.data['allow_model_changes'] and not confirmed_changes:
            raise ValueError('This profile modifies CAD copies; --confirm-model-changes is required')

    def validate_dependencies(self, dependencies):
        # ISldWorks.GetDocumentDependencies2(..., True, Searchflag, False)
        # returns alternating document names and full paths (no read-only flags).
        values = list(dependencies or [])
        if len(values) % 2:
            raise ValueError('Malformed dependency list; cannot establish safe inputs')
        for path in values[1::2]:
            if not path or not Path(path).is_absolute() or canonical(path) not in self.inputs or not inside(path, self.root):
                raise ValueError(f'Dependency is not a declared test copy: {path!r}')
        self.check_input_hashes()


class SessionState:
    def __init__(self, path):
        self.path = Path(path)

    def read(self):
        if not self.path.exists():
            return {'state': 'IDLE'}
        try:
            data = json.loads(self.path.read_text(encoding='utf-8'))
            if not isinstance(data, dict) or data.get('state') not in ('IDLE', 'RUNNING', 'UNKNOWN'):
                raise ValueError('Invalid session state')
            return data
        except (OSError, ValueError) as error:
            raise RuntimeError(f'Unreadable session marker {self.path}; inspect it before recovery') from error

    def write(self, data):
        # Called only while the controller lock is held. Replace atomically so a
        # killed controller leaves either the previous state or the new state.
        self.path.parent.mkdir(parents=True, exist_ok=True)
        staged = self.path.with_suffix('.tmp')
        with staged.open('w', encoding='utf-8') as stream:
            json.dump(data, stream, indent=2)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(staged, self.path)

    def ensure_ready(self):
        data = self.read()
        if data['state'] != 'IDLE':
            raise RuntimeError(f"Previous session is {data['state']} (SOLIDWORKS PID {data.get('solidworks_pid')}); "
                               'inspect SOLIDWORKS, then use --recover-session. A timeout does not cancel VBA.')

    def start(self, macro):
        self.ensure_ready()
        self.write({'state': 'RUNNING', 'controller_pid': os.getpid(), 'macro': str(macro),
                    'started': time.time(), 'solidworks_pid': None})

    def set_process(self, pid):
        self.write(dict(self.read(), solidworks_pid=int(pid)))

    def unknown(self, reason):
        self.write(dict(self.read(), state='UNKNOWN', reason=reason))

    def finish(self):
        self.write(dict(self.read(), state='IDLE', finished=time.time()))

    def recover(self, confirmed_stopped, process_alive):
        previous = self.read()
        if previous['state'] == 'IDLE':
            return
        if process_alive is not False and not confirmed_stopped:
            raise ValueError('SOLIDWORKS is alive or its state is unknown; inspect it and pass --confirm-macro-stopped')
        self.write({'state': 'IDLE', 'recovered': time.time(), 'previous': previous,
                    'confirmed_macro_stopped': bool(confirmed_stopped), 'process_alive': process_alive})
