"""Exact ownership and original-handle cleanup for disposable policy fixtures."""
import ctypes
import gc
from ctypes import wintypes
import json
import os
from pathlib import Path


def raw_creation(handle):
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.GetProcessTimes.argtypes = [wintypes.HANDLE] + [ctypes.POINTER(wintypes.FILETIME)] * 4
    kernel.GetProcessTimes.restype = wintypes.BOOL
    times = [wintypes.FILETIME() for _ in range(4)]
    if not kernel.GetProcessTimes(int(handle), *[ctypes.byref(t) for t in times]):
        raise ctypes.WinError(ctypes.get_last_error())
    return (times[0].dwHighDateTime << 32) | times[0].dwLowDateTime


class OwnedHost:
    def __init__(self, app, path):
        import win32api
        from msaccess_vcs_mcp.access_com.process_qos import pid_from_access_app, list_access_pids_or_none
        self.app, self.path, self.addin = app, Path(path).resolve(), None
        self.pid = pid_from_access_app(app)
        assert self.pid
        self.handle = win32api.OpenProcess(0x1000 | 0x100000, False, self.pid)
        self.creation = raw_creation(self.handle)
        self.receipt = dict(pid=self.pid, creation_FILETIME=self.creation, path=str(self.path), checks=[])
        self.unrelated = {}
        pids = list_access_pids_or_none()
        assert pids is not None, "Native inventory unconfirmed"
        for pid in pids - {self.pid}:
            handle = win32api.OpenProcess(0x1000 | 0x100000, False, pid)
            self.unrelated[pid] = (handle, raw_creation(handle))

    def check(self, label, actual, expected, native=True):
        self.receipt['checks'].append(dict(label=label, actual=actual, expected=expected,
            passed=actual == expected, evidence_kind='native' if native else 'controlled'))
        assert actual == expected, (label, actual, expected)

    def exited(self):
        import win32event
        result = win32event.WaitForSingleObject(self.handle, 0)
        assert result in (win32event.WAIT_OBJECT_0, win32event.WAIT_TIMEOUT)
        return result == win32event.WAIT_OBJECT_0

    def verified(self, closed=False):
        import win32api
        from msaccess_vcs_mcp.access_com.process_qos import pid_from_access_app
        if self.exited():
            return False
        handle = win32api.OpenProcess(0x1000 | 0x100000, False, self.pid)
        try:
            return (raw_creation(handle) == self.creation and
                    raw_creation(self.handle) == self.creation and
                    pid_from_access_app(self.app) == self.pid and
                    (not self.app.CurrentProject.FullName if closed else
                     Path(self.app.CurrentDb().Name).resolve() == self.path))
        finally:
            win32api.CloseHandle(handle)

    def close(self, release):
        import win32api
        import win32event
        try:
            # Drop temporary DAO/COM cycles while their server is still alive.
            gc.collect()
            already_exited = self.exited()
            self.receipt['already_exited_before_cleanup'] = already_exited
            if not already_exited:
                assert self.verified(), "Uncertain identity/path; fixture left untouched"
                if self.addin is not None:
                    self.check('teardown_owner_clear', json.loads(self.addin.call_sync('ClearOperationPolicy')), {'success': True})
                assert self.exited() or self.verified(), "Changed identity/path; left untouched"
                if not self.exited():
                    # With UserControl false, Access exits gracefully when the
                    # final automation reference is released. Release while the
                    # server is alive rather than Quit then release dead proxies.
                    self.app.UserControl = False
                    release()
                    self.receipt['cleanup_method'] = 'UserControl=False; final automation reference release'
            exited = win32event.WaitForSingleObject(self.handle, 10000) == win32event.WAIT_OBJECT_0
            self.receipt['original_handle_exit_confirmed'] = exited
            assert exited, "Owned fixture still live; no uncertain/force cleanup"
            preserved = []
            for pid, (handle, stamp) in self.unrelated.items():
                live = win32event.WaitForSingleObject(handle, 0) == win32event.WAIT_TIMEOUT
                preserved.append(dict(pid=pid, creation_FILETIME=stamp, original_handle_live=live))
                assert live and raw_creation(handle) == stamp
            self.receipt['unrelated_preservation'] = preserved
        finally:
            win32api.CloseHandle(self.handle)
            for handle, _ in self.unrelated.values():
                win32api.CloseHandle(handle)

    def save(self, test):
        print('M54 receipt:', json.dumps(self.receipt), flush=True)
        directory = os.environ.get('M54_RECEIPT_DIR')
        if directory:
            path = Path(directory) / (os.environ['M54_RUN_LABEL'] + '-' + test + '.json')
            with path.open('x', encoding='utf-8') as stream:
                json.dump(self.receipt, stream, indent=2)
