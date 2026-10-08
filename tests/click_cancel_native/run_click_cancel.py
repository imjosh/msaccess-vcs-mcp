"""Single unattended isolated entry point (interactive Windows desktop required).

Run with msaccess-vcs-mcp/venv/Scripts/python.exe -B run_click_cancel.py.
Invoked by the opt-in pytest adapter; no server operation replay, dialog-tool changes,
installed-library writes, or manually disconnected UI calls.
Timeouts: readiness 60s, dialog 30s, click 3s, completion 30s, cleanup 10s.
An uncertain dispatch is never replayed. Failure cleanup terminates only the
verified original disposable process handle if graceful worker cleanup stalls.
Receipts/binaries are retained, including failures. This mutex serializes these
isolated runs; production MCP gates are process-local, so other Access work must
stay quiescent. Empty native entry inventory is mandatory.
"""
import sys
sys.dont_write_bytecode = True
import gc
import hashlib
import json
import os
from pathlib import Path
import subprocess
import time
import traceback
from datetime import datetime
from native_cancel import write

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[2]
ARTIFACTS = ROOT / 'verification/click-cancel-integration'
sys.path.insert(0, str(ROOT / 'msaccess-vcs-mcp'))


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def final_pass(worker, unchanged, exited, errors):
    return (worker.get('outcome') == 'behavior_passed' and
            worker.get('installed_unchanged') is True and unchanged and exited and
            not errors and worker.get('ownership', {}).get('original_handle_exit_confirmed') is True)


def finish(work, receipt, worker_result):
    receipt['worker'] = worker_result
    receipt['outcome'] = 'passed' if final_pass(worker_result,
        receipt.get('installed_unchanged', False),
        receipt.get('original_handle_exit_confirmed', False), receipt['errors']) else 'failed'
    write(work / 'result.json', receipt)
    return 0 if receipt['outcome'] == 'passed' else 1


def wait_for(predicate, seconds, worker):
    end = time.monotonic() + seconds
    while time.monotonic() < end:
        value = predicate()
        if value:
            return value
        if worker.poll() is not None:
            raise RuntimeError('Worker exited before expected readiness/completion')
        time.sleep(.1)
    raise TimeoutError('Bounded fixture readiness/operation completion timeout')


def run(work):
    import win32api
    import win32event
    from tests.policy_live_host import raw_creation
    from msaccess_vcs_mcp import __version__
    from msaccess_vcs_mcp.compatibility import workflow_requirement
    from msaccess_vcs_mcp.config import get_default_addin_path
    from msaccess_vcs_mcp.access_com.process_qos import list_access_pids_or_none
    expected_python = ROOT / 'msaccess-vcs-mcp/venv/Scripts/python.exe'
    assert Path(sys.executable).resolve() == expected_python.resolve(), 'Use MCP project venv'
    receipt = dict(outcome='failed', work=str(work), errors=[], unattended=True,
        driver='isolated fresh HWND/MSAA native Cancel control; no coordinates',
        python=sys.executable, mcp_version=__version__, timeouts_seconds={
            'readiness': 60, 'dialog': 30, 'click': 3, 'completion': 30, 'cleanup': 10})
    mutex = process_handle = worker = output = None
    installed = stamp = identity = None
    worker_result = {}
    def capture():
        nonlocal process_handle, identity
        if process_handle is None and (work / 'ownership.json').exists():
            identity = json.loads((work / 'ownership.json').read_text())
            assert identity['path'] == str(work / 'cancel-fixture.accdb')
            assert isinstance(identity['creation_FILETIME'], str), 'FILETIME must not pass through JS Number'
            process_handle = win32api.OpenProcess(0x1000 | 0x100000 | 1, False, identity['pid'])
            assert raw_creation(process_handle) == int(identity['creation_FILETIME'])
            receipt['ownership'] = identity
        return process_handle
    def verify():
        assert capture() is not None, 'No original process handle'
        assert win32event.WaitForSingleObject(process_handle, 0) == win32event.WAIT_TIMEOUT
        assert raw_creation(process_handle) == int(identity['creation_FILETIME'])
        assert list_access_pids_or_none() == {identity['pid']}, 'Concurrent Access work detected; no click'
    try:
        mutex = win32event.CreateMutex(None, False, 'Local\\VERIFY2_IsolatedClickCancel')
        lock = win32event.WaitForSingleObject(mutex, 0)
        assert lock in (win32event.WAIT_OBJECT_0, 0x80), 'Another isolated run owns Access work'
        receipt['mutex_acquired'] = True
        import ctypes
        user32 = ctypes.WinDLL('user32', use_last_error=True)
        user32.OpenInputDesktop.restype = ctypes.c_void_p
        user32.CloseDesktop.argtypes = [ctypes.c_void_p]
        desktop = user32.OpenInputDesktop(0, False, 0x0100)
        assert desktop, 'Interactive input desktop unavailable'
        assert user32.CloseDesktop(desktop), 'Input desktop handle cleanup failed'
        receipt['interactive_desktop_confirmed'] = True
        assert workflow_requirement().reason(__version__) is None
        assert list_access_pids_or_none() == set(), 'Require confirmed empty Access inventory'
        installed = Path(get_default_addin_path())
        stamp = sha(installed)
        receipt.update(installed=str(installed), installed_sha_before=stamp)
        receipt['source_repositories'] = {
            repo: subprocess.run(['git', '-C', str(ROOT / repo), 'rev-parse', 'HEAD'],
                capture_output=True, text=True, check=True, timeout=5).stdout.strip()
            for repo in ('msaccess-vcs-mcp', 'msaccess-vcs-addin')}
        # Read-only compatibility metadata runs with isolated environment too.
        env = os.environ.copy()
        for key, relative in {'ACCESS_VCS_LOG_DIR': 'usage',
                'ACCESS_VCS_DIAGNOSTIC_LOG_DIR': 'diagnostics',
                'ACCESS_VCS_OWNED_INSTANCES_PATH': 'owned.json'}.items():
            env[key] = str(work / relative)
        output = (work / 'worker-output.txt').open('w', encoding='utf-8')
        worker = subprocess.Popen([sys.executable, '-B', str(HERE / 'com_worker.py'),
                                   '--worker', str(work)], env=env, stdout=output, stderr=subprocess.STDOUT)
        wait_for(capture, 60, worker)
        for case in ('export', 'import'):
            path = work / (case + '-pending.json')
            wait_for(path.exists, 60 if case == 'export' else 30, worker)
            pending = json.loads(path.read_text())
            assert pending == dict(case=case, pid=identity['pid'], creation_FILETIME=identity['creation_FILETIME'],
                database=identity['path'], work=str(work))
            verify()
            # Hard outer bound includes MSAA calls, not just polling. No retries.
            ui = subprocess.Popen([sys.executable, '-B', str(HERE / 'native_cancel.py'),
                                   str(path)], env=env, stdout=output, stderr=subprocess.STDOUT)
            try:
                ui.wait(timeout=35)
            except BaseException:
                ui.kill()
                ui.wait(timeout=5)
                click_path = work / (case + '-click.json')
                uncertain = json.loads(click_path.read_text()) if click_path.exists() else dict(pending=pending)
                uncertain.update(status='uncertain', error='UI subprocess did not complete; no replay')
                write(click_path, uncertain)
                raise
            assert ui.returncode == 0, 'Native UI observation/click failed; no replay'
            wait_for((work / (case + '-operation-result.json')).exists, 30, worker)
        worker.wait(timeout=30)
        assert worker.returncode == 0, 'Worker behavior/cleanup failed'
        worker_result = json.loads((work / 'worker-result.json').read_text())
    except BaseException:
        receipt['errors'].append(traceback.format_exc())
    finally:
        # No second operation/click on failure; preserve uncertain results.
        try:
            if worker is not None:
                capture()
                if process_handle is not None:
                    exited = win32event.WaitForSingleObject(process_handle, 1000) == win32event.WAIT_OBJECT_0
                    if not exited:
                        assert raw_creation(process_handle) == int(identity['creation_FILETIME'])
                        receipt['cleanup_method'] = 'TerminateProcess on verified original disposable handle; never PID lookup'
                        # This is failed-run disposal, never behavior evidence.
                        receipt['errors'].append('Graceful fixture cleanup incomplete; owned disposal required')
                        win32api.TerminateProcess(process_handle, 1)
                        exited = win32event.WaitForSingleObject(process_handle, 10000) == win32event.WAIT_OBJECT_0
                    receipt['original_handle_exit_confirmed'] = exited
                    assert exited, 'Original handle exit not confirmed'
                else:
                    receipt['original_handle_exit_confirmed'] = False
                    receipt['errors'].append('No exact original Access handle; uncertain host left untouched')
                try:
                    worker.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    worker.kill()  # only our Python subprocess
                    worker.wait(timeout=10)
                    receipt['errors'].append('Python worker cleanup timed out')
                receipt['worker_exit_code'] = worker.returncode
                if (work / 'worker-result.json').exists():
                    worker_result = json.loads((work / 'worker-result.json').read_text())
        except BaseException:
            receipt['errors'].append('Cleanup failure: ' + traceback.format_exc())
        finally:
            if worker is not None and worker.poll() is None:
                try:
                    worker.kill()
                    worker.wait(timeout=5)
                    receipt['errors'].append('Worker forcibly stopped after incomplete cleanup')
                except BaseException:
                    receipt['errors'].append('Worker cleanup failure: ' + traceback.format_exc())
            if output:
                output.close()
            try:
                if process_handle:
                    win32api.CloseHandle(process_handle)
                if mutex:
                    if receipt.get('mutex_acquired'):
                        win32event.ReleaseMutex(mutex)
                    win32api.CloseHandle(mutex)
            except BaseException:
                receipt['errors'].append('Handle cleanup failure: ' + traceback.format_exc())
            try:
                receipt['installed_sha_after'] = sha(installed) if installed else None
                receipt['installed_unchanged'] = stamp is not None and receipt['installed_sha_after'] == stamp
            except BaseException:
                receipt['errors'].append('Installed preservation unconfirmed: ' + traceback.format_exc())
            finish(work, receipt, worker_result)
            print(json.dumps(dict(receipt=str(work / 'result.json'), outcome=receipt['outcome'])), flush=True)
    return 0 if receipt['outcome'] == 'passed' else 1


def entry(work=None):
    ARTIFACTS.mkdir(parents=True, exist_ok=True)
    work = Path(work).resolve() if work else ARTIFACTS / datetime.now().strftime('run-%Y%m%d-%H%M%S-%f')
    assert work.parent == ARTIFACTS.resolve(), 'Use designated integration artifacts'
    work.mkdir()
    settings = {'ACCESS_VCS_LOG_DIR': str(work / 'usage'),
                'ACCESS_VCS_DIAGNOSTIC_LOG_DIR': str(work / 'diagnostics'),
                'ACCESS_VCS_OWNED_INSTANCES_PATH': str(work / 'owned.json'),
                'ACCESS_VCS_PROJECT_DIR': str(ROOT / 'msaccess-vcs-mcp')}
    original = {key: os.environ.get(key) for key in settings}
    try:
        os.environ.update(settings)
        return run(work)
    finally:
        for key, value in original.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value


if __name__ == '__main__':
    raise SystemExit(entry(sys.argv[1] if len(sys.argv) == 2 else None))
