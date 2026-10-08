"""Disposable STA worker and bounded native dialog driver for BuildAs tests."""
from __future__ import annotations

import gc
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import time
import traceback
import re
from contextlib import contextmanager


def write(path, value):
    temporary = path.with_suffix('.tmp')
    temporary.write_text(json.dumps(value, indent=2), encoding='utf-8')
    temporary.replace(path)


def native_json(value):
    return json.loads(value[0] if isinstance(value, tuple) else value)


RELEVANT = (
    'clsVersionControl', 'modAPI', 'modCompatibilitySession', 'modBuild',
    'clsOperation', 'clsInteractionScope', 'clsRootOperationLease',
    'clsVbeErrorTrappingScope', 'modObjects', 'modConstants', 'clsOptions',
    'modFileAccess', 'modFileWinAPI', 'modDatabase', 'modVCSUtility',
    'modErrorHandling', 'modUIUtil', 'clsTranslation', 'clsLog',
    'clsDbProject', 'clsDbQuery', 'clsDbVbeProject', 'clsDbVbeReference',
    'modTestModeOnlyScope',
)


def sha(raw):
    return hashlib.sha256(raw).hexdigest()


def canonical_code(code):
    """Fold VBA identifiers only; preserve strings, comments and other bytes.

    CRLF/CR become LF. No whitespace, blank-line, or literal normalization.
    Rem consumes the rest of its line; apostrophe comments do likewise.
    """
    code = code.replace('\r\n', '\n').replace('\r', '\n')
    token = re.compile(r'"(?:[^"\n]|"")*"|\'[^\n]*|\bRem\b[^\n]*|[A-Za-z_][A-Za-z_0-9]*', re.I)
    return token.sub(lambda m: m[0] if m[0].startswith(('"', "'")) or
                     re.match(r'Rem\b', m[0], re.I) else m[0].lower(), code)


def loaded_provenance(app, path, base, phase):
    """Read the actual library; never open the installed file as a target."""
    root = Path(__file__).resolve().parents[2]
    qualified = root / 'verification/VERIFY-2-2026-10-08-01/evidence'
    receipt_path = qualified / 'full-01-before-project.json'
    receipt = json.loads(receipt_path.read_text(encoding='utf-8'))
    snapshot_path = qualified / 'build-1-snapshot.json'
    snapshot = json.loads(snapshot_path.read_text(encoding='utf-8'))
    target = Path(path).resolve()
    library = str(target.with_suffix(''))
    identity = native_json(app.Run(library + '.APIIdentity'))
    assert identity['success'] and Path(identity['loaded_path']).resolve() == target, identity
    projects = [p for p in app.VBE.VBProjects if Path(str(p.FileName)).resolve() == target]
    assert len(projects) == 1, 'Actual library project is ambiguous'
    project = projects[0]
    qproject = next(p for p in receipt['projects'] if Path(p['path']).resolve() == target)
    qcomponents = {c['name']: c for c in qproject['components']}
    rows = []
    source_root = root / 'msaccess-vcs-addin/Version Control.accda.src/modules'
    for name in RELEVANT:
        component = project.VBComponents(name)
        module = component.CodeModule
        code = str(module.Lines(1, module.CountOfLines))
        candidates = [p for p in source_root.rglob(name + '.*') if p.suffix in ('.bas', '.cls')]
        assert len(candidates) == 1, name
        source = candidates[0]
        source_bytes = source.read_bytes()
        relative = str(source.relative_to(root / 'msaccess-vcs-addin'))
        lines = source_bytes.decode('utf-8-sig').splitlines()
        start = next(i for i, line in enumerate(lines) if line.startswith('Attribute VB_Name'))
        source_code = '\r\n'.join(line for line in lines[start:] if not line.startswith('Attribute '))
        row = dict(name=name, type=component.Type, lines=module.CountOfLines,
                   qualified_lines=qcomponents[name]['lines'], code=code,
                   runtime_sha256=sha(code.encode('utf-8')),
                   qualified_runtime_sha256=qcomponents[name]['code_sha256'],
                   source_path=str(source), source_sha256=sha(source_bytes),
                   qualified_source_sha256=snapshot['files'][relative],
                   runtime_canonical_sha256=sha(canonical_code(code).encode('utf-8')),
                   source_canonical_sha256=sha(canonical_code(source_code).encode('utf-8')))
        row['exact_qualified_runtime_match'] = row['runtime_sha256'] == row['qualified_runtime_sha256']
        row['qualified_source_bytes_match'] = row['source_sha256'] == row['qualified_source_sha256']
        row['identifier_case_and_line_ending_equivalent'] = row['runtime_canonical_sha256'] == row['source_canonical_sha256']
        rows.append(row)
        module = component = None
    value = dict(api_identity=identity, project_path=str(project.FileName),
                 installed_sha256=sha(target.read_bytes()),
                 qualified_receipt=str(receipt_path), qualified_receipt_sha256=sha(receipt_path.read_bytes()),
                 qualified_snapshot=str(snapshot_path), qualified_snapshot_sha256=sha(snapshot_path.read_bytes()),
                 qualified_head=snapshot['head'], components=rows,
                 normalization='Remove source export scaffolding/Attribute lines; normalize CRLF/CR to LF; lowercase identifier tokens only, preserving strings and comments. No whitespace normalization.')
    value['relevant_component_equivalence'] = all(r['qualified_source_bytes_match'] and
        r['identifier_case_and_line_ending_equivalent'] and r['lines'] == r['qualified_lines'] for r in rows)
    write(base / ('provenance-' + phase + '.json'), value)
    project = projects = None
    return value


def fixture(base):
    source = base / 'source'
    (source / 'queries').mkdir(parents=True)
    for name, value in {
        'project.json': {'Info': {'Class': 'clsDbProject'}, 'Items': {'FileFormat': 12}},
        'vbe-project.json': {'Info': {'Class': 'clsDbVbeProject'}, 'Items': {'Name': 'PickerFixture', 'FileName': 'original.accdb'}},
        'vbe-references.json': {'Info': {'Class': 'clsDbVbeReference'}, 'Items': {'DAO': {'GUID': '{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}', 'Version': '12.0'}}},
        'vcs-options.json': {'Options': {'ExportFolder': '', 'ExportFormatVersion': '5.1.0', 'SaveQuerySQL': True}},
        'queries/qryPickerProof.json': {'Info': {'Class': 'clsDbQuery'}, 'Items': {'QueryType': 0}},
    }.items():
        (source / name).write_text(json.dumps(value), encoding='utf-8')
    (source / 'queries/qryPickerProof.sql').write_text('SELECT 731 AS PickerProof;\n')
    (base / 'invalid-source').mkdir()
    return source


def worker(base, case):
    import pythoncom
    import win32api
    import win32com.client
    from unittest.mock import patch
    from msaccess_vcs_mcp.addin_integration import VCSAddinIntegration
    from msaccess_vcs_mcp.compatibility_session import ensure_session
    from msaccess_vcs_mcp.config import get_default_addin_path
    from msaccess_vcs_mcp.access_com.process_qos import pid_from_access_app
    # Reuse exact raw FILETIME handling from the existing native test harness.
    from tests.policy_live_host import raw_creation

    os.environ['ACCESS_VCS_LOG_DIR'] = str(base / 'usage')
    os.environ['ACCESS_VCS_DIAGNOSTIC_LOG_DIR'] = str(base / 'diagnostics')
    os.environ['ACCESS_VCS_OWNED_INSTANCES_PATH'] = str(base / 'owned.json')
    app = addin = db = query = rows = None
    result = {'case': case, 'passed': False}
    pythoncom.CoInitialize()
    try:
        app = win32com.client.DispatchEx('Access.Application')
        pid = pid_from_access_app(app)
        handle = win32api.OpenProcess(0x1000 | 0x100000, False, pid)
        try:
            identity = {'pid': pid, 'creation_FILETIME': raw_creation(handle)}
        finally:
            win32api.CloseHandle(handle)
        write(base / 'launched.json', identity)
        host = base / 'picker-host.accdb'
        output = base / 'picked-output.accdb'
        app.NewCurrentDatabase(str(host))
        app.CurrentDb().CreateQueryDef('qryHostSentinel', 'SELECT 19 AS Sentinel')
        app.Visible = True
        app.UserControl = True
        if case == 'sentinel':
            write(base / 'ready.json', identity)
            deadline = time.monotonic() + 300
            while not (base / 'stop').exists():
                assert time.monotonic() < deadline, 'Sentinel lifetime exceeded bound'
                time.sleep(.1)
            result['sentinel_sql'] = app.CurrentDb().QueryDefs('qryHostSentinel').SQL
            assert result['sentinel_sql'] == 'SELECT 19 AS Sentinel;\r\n'
            result['passed'] = True
            return 0
        addin = VCSAddinIntegration(get_default_addin_path())
        result['installed_path'] = addin.addin_path
        result['installed_sha256_before'] = sha(Path(addin.addin_path).read_bytes())
        def probe(probe_app, db_path, timeout):
            assert probe_app is app and db_path is None
            addin._probe_session = ensure_session(app, addin.addin_path)
        with patch.object(addin, '_probe_with_timeout', side_effect=probe):
            addin.load_addin(app)
        result['session'] = ensure_session(app, addin.addin_path)
        from msaccess_vcs_mcp import __version__
        from msaccess_vcs_mcp.compatibility import ADDIN_REQUIREMENT, Requirement, compatibility_result
        result['server_version'] = __version__
        result['workflow_admission'] = compatibility_result(__version__, Requirement('0.3.0', '0.4.0', ('0.3.0-dev.18',)))
        assert result['workflow_admission']['success'], result
        result['provenance_before'] = loaded_provenance(app, addin.addin_path, base, 'before')
        result['addin_admission'] = compatibility_result(result['provenance_before']['api_identity']['addin_version'], ADDIN_REQUIREMENT)
        assert result['addin_admission']['success'], result
        assert json.loads(result['session'])['addin_instance'] == result['provenance_before']['api_identity']['addin_instance']
        # A source mismatch is retained as evidence and fails this case; no rebuild.
        assert result['provenance_before']['relevant_component_equivalence'], 'Loaded BuildAs-path component/source mismatch'
        mode = json.loads(addin.call_sync('SetInteractionMode', 0))
        assert mode['success'] and mode['effective_mode'] == 0, mode
        library = str(Path(addin.addin_path).with_suffix(''))
        result['before'] = native_json(app.Run(library + '.ModeOnlyScopeProbe', 'snapshot'))
        write(base / 'ready.json', {**identity, 'source': str(base / 'source'), 'output': str(output)})
        result['return'] = addin.call_sync('BuildAs')  # deliberately zero path arguments
        result['current_database'] = app.CurrentDb().Name
        result['after'] = native_json(app.Run(library + '.ModeOnlyScopeProbe', 'snapshot'))
        result['provenance_after'] = loaded_provenance(app, addin.addin_path, base, 'after')
        assert result['provenance_after']['api_identity'] == result['provenance_before']['api_identity']
        assert result['provenance_after']['relevant_component_equivalence']
        assert [r['runtime_sha256'] for r in result['provenance_after']['components']] == [r['runtime_sha256'] for r in result['provenance_before']['components']]
        if case == 'success':
            assert output.is_file()
            assert Path(result['current_database']).resolve() == output.resolve()
            db = app.CurrentDb()
            query = db.QueryDefs('qryPickerProof')
            rows = query.OpenRecordset()
            result['query_sql'] = query.SQL
            result['query_value'] = rows.Fields('PickerProof').Value
            assert result['query_value'] == 731
        else:
            assert not output.exists(), 'Cancelled/invalid selection created output'
            assert Path(result['current_database']).resolve() == host.resolve()
            db = app.CurrentDb()
            query = db.QueryDefs('qryHostSentinel')
            rows = query.OpenRecordset()
            result['sentinel_value'] = rows.Fields('Sentinel').Value
            assert result['sentinel_value'] == 19
        # The real test probe records scope state; no fabricated state or fault seam.
        assert result['after']['token'] == '', result
        assert result['after']['status'] == 'ready', result
        assert result['after']['mode'] == result['before']['mode'], result
        assert result['after']['force_unattended'] == result['before']['force_unattended'], result
        assert result['after']['result'] == (1 if case == 'success' else 4), result
        assert result['after']['policy'] == result['before']['policy'], result
        assert not result['after']['blocked'] and not result['after']['decisions'], result
        result['passed'] = True
    except Exception:
        result['error'] = traceback.format_exc()
    finally:
        if rows is not None:
            rows.Close()
        rows = query = db = None
        if app is not None:
            try:
                result['cleanup_current_database'] = app.CurrentDb().Name
                app.CloseCurrentDatabase()
                app.UserControl = False
                app.Quit(2)
                result['quit_requested'] = True
            except Exception:
                result['cleanup_error'] = traceback.format_exc()
        addin = app = None
        gc.collect()
        pythoncom.CoUninitialize()
        if addin_path := result.get('installed_path'):
            result['installed_sha256_after'] = sha(Path(addin_path).read_bytes())
            if result['installed_sha256_after'] != result['installed_sha256_before']:
                result['passed'] = False
                result['installed_binary_changed'] = True
        write(base / 'result.json', result)
    return 0 if result['passed'] and not result.get('cleanup_error') else 1


class DialogDriver:
    """Addresses observed controls only within an exact owned Access identity."""
    def __init__(self, identity, handle):
        from msaccess_vcs_mcp.dialog_recovery import Win32Backend
        self.identity, self.handle = identity, handle
        self.backend = Win32Backend()
        self.events = []

    def verify(self, hwnd=None):
        import win32process
        import win32event
        from tests.policy_live_host import raw_creation
        assert raw_creation(self.handle) == self.identity['creation_FILETIME']
        assert win32event.WaitForSingleObject(self.handle, 0) == win32event.WAIT_TIMEOUT
        if hwnd:
            assert win32process.GetWindowThreadProcessId(hwnd)[1] == self.identity['pid']

    def find(self, title):
        self.verify()
        matches = [w for w in self.backend.list_windows()
                   if w.pid == self.identity['pid'] and w.title == title]
        assert len(matches) <= 1, f'Ambiguous dialog: {title}'
        return matches[0] if matches else None

    def wait(self, title, timeout=30):
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            window = self.find(title)
            if window:
                self.events.append({'observed': title, 'hwnd': window.hwnd})
                return window
            time.sleep(.1)
        raise AssertionError(f'Dialog did not appear: {title}')

    def button(self, title, text):
        from msaccess_vcs_mcp.dialog_recovery import CLICK_DELIVERED
        window = self.wait(title)
        buttons = [b for b in window.buttons if b.text.replace('&', '').strip().lower() == text.lower()]
        assert len(buttons) == 1, (title, text, window)
        self.verify(window.hwnd)
        assert self.backend.click(buttons[0], expected_pid=self.identity['pid'], timeout_ms=3000) == CLICK_DELIVERED
        self.events.append({'clicked': text, 'dialog': title})

    def path(self, title, control_id, path):
        import ctypes
        from ctypes import wintypes
        import win32gui
        import win32con
        window = self.wait(title)
        controls = []
        def collect(hwnd, _):
            if win32gui.GetClassName(hwnd) == 'Edit' and win32gui.GetDlgCtrlID(hwnd) == control_id:
                controls.append(hwnd)
        win32gui.EnumChildWindows(window.hwnd, collect, None)
        assert len(controls) == 1, (title, control_id, controls)
        self.verify(controls[0])
        send = ctypes.WinDLL('user32', use_last_error=True).SendMessageTimeoutW
        send.argtypes = [wintypes.HWND, wintypes.UINT, ctypes.c_size_t, ctypes.c_ssize_t,
                         wintypes.UINT, wintypes.UINT, ctypes.POINTER(ctypes.c_size_t)]
        send.restype = ctypes.c_ssize_t
        value = ctypes.create_unicode_buffer(str(path))
        answer = ctypes.c_size_t()
        assert send(controls[0], win32con.WM_SETTEXT, 0, ctypes.addressof(value),
                    win32con.SMTO_ABORTIFHUNG, 3000, ctypes.byref(answer))
        text = ctypes.create_unicode_buffer(32768)
        assert send(controls[0], win32con.WM_GETTEXT, len(text), ctypes.addressof(text),
                    win32con.SMTO_ABORTIFHUNG, 3000, ctypes.byref(answer))
        assert text.value == str(path), (text.value, str(path))
        self.events.append({'entered': str(path), 'dialog': title})

    def source(self, path):
        self.path('Select Source Folder', 1152, path)
        self.button('Select Source Folder', 'Select Source Files Folder')
        # The shell may navigate first and require a second confirmation.
        deadline = time.monotonic() + 5
        while time.monotonic() < deadline:
            window = self.find('Select Source Folder')
            if window is None:
                return
            time.sleep(.1)
        self.button('Select Source Folder', 'Select Source Files Folder')

    def drive(self, base, case):
        self.wait('Select Source Folder')
        if case == 'cancel_source':
            self.button('Select Source Folder', 'Cancel')
            return
        self.source(base / ('invalid-source' if case == 'invalid_source' else 'source'))
        if case == 'invalid_source':
            warning = self.wait('Version Control Add-in')
            message = ' '.join(warning.texts)
            assert 'Required source files were not found' in message, message
            self.events.append({'invalid_source_message': message})
            self.button('Version Control Add-in', 'OK')
            return
        self.wait('Build New Database File')
        if case == 'cancel_output':
            self.button('Build New Database File', 'Cancel')
        else:
            self.path('Build New Database File', 1001, base / 'picked-output.accdb')
            self.button('Build New Database File', 'Build Here')


def run_case(base, case):
    import win32api
    import win32event
    from tests.policy_live_host import raw_creation
    from msaccess_vcs_mcp.access_com.process_qos import list_access_pids_or_none
    base = Path(base).resolve()
    base.mkdir(parents=True, exist_ok=True)
    fixture(base)
    inventory = list_access_pids_or_none()
    assert inventory is not None, 'Unconfirmed process inventory'
    unrelated = []
    for pid in inventory:
        handle = win32api.OpenProcess(0x1000 | 0x100000, False, pid)
        unrelated.append((pid, handle, raw_creation(handle)))
    handle = None
    driver = None
    outcome = {'case': case, 'passed': False}
    import tests.policy_live_host as host_module
    env = dict(os.environ, PYTHONDONTWRITEBYTECODE='1',
               PYTHONPATH=str(Path(host_module.__file__).resolve().parents[1]))
    with (base / 'worker.txt').open('w', encoding='utf-8') as capture:
        proc = subprocess.Popen([sys.executable, '-B', __file__, '--worker', str(base), case],
                                stdout=capture, stderr=subprocess.STDOUT, env=env)
        try:
            deadline = time.monotonic() + 45
            while not (base / 'ready.json').exists():
                if handle is None and (base / 'launched.json').exists():
                    identity = json.loads((base / 'launched.json').read_text())
                    handle = win32api.OpenProcess(0x1000 | 0x100000 | 1, False, identity['pid'])
                    assert raw_creation(handle) == identity['creation_FILETIME']
                if proc.poll() is not None or time.monotonic() > deadline:
                    diagnostic = (base / 'result.json').read_text() if (base / 'result.json').exists() else (base / 'worker.txt').read_text()
                    raise AssertionError('Worker did not become ready: ' + diagnostic)
                time.sleep(.1)
            identity = json.loads((base / 'ready.json').read_text())
            if handle is None:
                handle = win32api.OpenProcess(0x1000 | 0x100000 | 1, False, identity['pid'])
            assert raw_creation(handle) == identity['creation_FILETIME']
            outcome['identity'] = identity
            driver = DialogDriver(identity, handle)
            driver.drive(base, case)
            proc.wait(timeout=60)
            outcome['worker'] = json.loads((base / 'result.json').read_text())
            assert proc.returncode == 0 and outcome['worker']['passed'], outcome
            capture.flush()
            diagnostic = (base / 'worker.txt').read_text(encoding='utf-8')
            assert 'fatal exception' not in diagnostic.lower() and 'Traceback' not in diagnostic, diagnostic
            exited = win32event.WaitForSingleObject(handle, 15000) == win32event.WAIT_OBJECT_0
            outcome['graceful_exit'] = exited
            assert exited, 'Owned Access did not exit after worker apartment shutdown'
            outcome['passed'] = True
        finally:
            if proc.poll() is None:
                proc.terminate()
                proc.wait(timeout=10)
            if handle is None and (base / 'launched.json').exists():
                identity = json.loads((base / 'launched.json').read_text())
                handle = win32api.OpenProcess(0x1000 | 0x100000 | 1, False, identity['pid'])
                assert raw_creation(handle) == identity['creation_FILETIME']
            if handle is not None:
                if win32event.WaitForSingleObject(handle, 0) != win32event.WAIT_OBJECT_0:
                    try:
                        win32api.TerminateProcess(handle, 1)
                        outcome['emergency_termination'] = True
                    except Exception:
                        if win32event.WaitForSingleObject(handle, 5000) != win32event.WAIT_OBJECT_0:
                            raise
                outcome['original_handle_exit_confirmed'] = win32event.WaitForSingleObject(handle, 5000) == win32event.WAIT_OBJECT_0
                win32api.CloseHandle(handle)
            outcome['unrelated'] = []
            for pid, other, stamp in unrelated:
                preserved = win32event.WaitForSingleObject(other, 0) == win32event.WAIT_TIMEOUT and raw_creation(other) == stamp
                outcome['unrelated'].append({'pid': pid, 'creation_FILETIME': stamp, 'preserved': preserved})
                win32api.CloseHandle(other)
            outcome['dialogs'] = driver.events if driver else []
            write(base / 'receipt.json', outcome)
    assert all(item['preserved'] for item in outcome['unrelated']), outcome
    return outcome


@contextmanager
def sentinel(base):
    """A separate Access worker whose COM object is never retained by pytest."""
    import tests.policy_live_host as host_module
    from tests.policy_live_host import raw_creation
    import win32api
    import win32event
    base = Path(base).resolve()
    base.mkdir(parents=True, exist_ok=True)
    env = dict(os.environ, PYTHONDONTWRITEBYTECODE='1',
               PYTHONPATH=str(Path(host_module.__file__).resolve().parents[1]))
    identity = None
    handle = None
    receipt = {'passed': False}
    with (base / 'worker.txt').open('w', encoding='utf-8') as capture:
        proc = subprocess.Popen([sys.executable, '-B', __file__, '--worker', str(base), 'sentinel'],
                                stdout=capture, stderr=subprocess.STDOUT, env=env)
        try:
            deadline = time.monotonic() + 30
            while not (base / 'ready.json').exists():
                assert proc.poll() is None and time.monotonic() < deadline, 'Sentinel did not start'
                time.sleep(.1)
            identity = json.loads((base / 'ready.json').read_text())
            handle = win32api.OpenProcess(0x1000 | 0x100000 | 1, False, identity['pid'])
            assert raw_creation(handle) == identity['creation_FILETIME']
            receipt['identity'] = identity
            yield identity
            (base / 'stop').touch()
            proc.wait(timeout=30)
            receipt['worker'] = json.loads((base / 'result.json').read_text())
            assert proc.returncode == 0 and receipt['worker']['passed'], receipt
            receipt['graceful_exit'] = win32event.WaitForSingleObject(handle, 15000) == win32event.WAIT_OBJECT_0
            assert receipt['graceful_exit'], 'Sentinel did not exit'
            receipt['passed'] = True
        finally:
            if proc.poll() is None:
                proc.terminate()
                proc.wait(timeout=10)
            if handle is not None:
                if win32event.WaitForSingleObject(handle, 0) != win32event.WAIT_OBJECT_0:
                    win32api.TerminateProcess(handle, 1)
                    receipt['emergency_termination'] = True
                receipt['original_handle_exit_confirmed'] = win32event.WaitForSingleObject(handle, 5000) == win32event.WAIT_OBJECT_0
                win32api.CloseHandle(handle)
            write(base / 'receipt.json', receipt)


if __name__ == '__main__':
    sys.exit(worker(Path(sys.argv[2]), sys.argv[3]))
