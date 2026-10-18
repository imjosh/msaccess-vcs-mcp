"""Prefer full-power cores for MCP-launched Access processes.

Access is single-threaded. On hybrid CPUs, a windowless COM-launched
``MSACCESS.EXE`` is a good candidate for EcoQoS / LP-E scheduling, which
makes the same VBA work slower than a foreground ribbon run. This module
turns execution-speed power throttling off and raises the process to
Above Normal. It does not set CPU affinity: pinning fights the scheduler
and is wrong on machines without hybrid cores.

Only processes the server created, or that appeared after an MCP rebuild
launched, are touched. User-owned Access that was already running is left
alone. Failures are swallowed; older Windows without
``SetProcessInformation`` still gets the priority change.
"""

from __future__ import annotations

import csv
import ctypes
import subprocess
import sys
import time
from io import StringIO
from typing import Any

PROCESS_SET_INFORMATION = 0x0200
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
PROCESS_TERMINATE = 0x0001
SYNCHRONIZE = 0x00100000
PROCESS_POWER_THROTTLING = 4
PROCESS_POWER_THROTTLING_CURRENT_VERSION = 1
PROCESS_POWER_THROTTLING_EXECUTION_SPEED = 0x1
ABOVE_NORMAL_PRIORITY_CLASS = 0x00008000


class PROCESS_POWER_THROTTLING_STATE(ctypes.Structure):
    _fields_ = [
        ("Version", ctypes.c_uint32),
        ("ControlMask", ctypes.c_uint32),
        ("StateMask", ctypes.c_uint32),
    ]


_kernel32_mod = None


def _kernel32():
    global _kernel32_mod
    if _kernel32_mod is not None:
        return _kernel32_mod

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.OpenProcess.argtypes = [
        ctypes.c_uint32,
        ctypes.c_bool,
        ctypes.c_uint32,
    ]
    kernel32.OpenProcess.restype = ctypes.c_void_p
    kernel32.SetProcessInformation.argtypes = [
        ctypes.c_void_p,
        ctypes.c_int,
        ctypes.c_void_p,
        ctypes.c_uint32,
    ]
    kernel32.SetProcessInformation.restype = ctypes.c_bool
    kernel32.SetPriorityClass.argtypes = [ctypes.c_void_p, ctypes.c_uint32]
    kernel32.SetPriorityClass.restype = ctypes.c_bool
    kernel32.TerminateProcess.argtypes = [ctypes.c_void_p, ctypes.c_uint32]
    kernel32.TerminateProcess.restype = ctypes.c_bool
    kernel32.CloseHandle.argtypes = [ctypes.c_void_p]
    kernel32.CloseHandle.restype = ctypes.c_bool
    kernel32.WaitForSingleObject.argtypes = [ctypes.c_void_p, ctypes.c_uint32]
    kernel32.WaitForSingleObject.restype = ctypes.c_uint32
    _kernel32_mod = kernel32
    return kernel32


class ExactProcess:
    """Observe and terminate one creation-time-verified process handle."""

    def __init__(self, kernel32, handle):
        self._kernel32 = kernel32
        self._handle = handle

    def exited(self) -> bool | None:
        try:
            result = self._kernel32.WaitForSingleObject(self._handle, 0)
        except Exception:
            return None
        if result == 0:
            return True
        if result == 258:
            return False
        return None

    def force(self) -> bool:
        try:
            return bool(self._kernel32.TerminateProcess(self._handle, 1))
        except Exception:
            return False

    def close(self) -> None:
        if self._handle is not None:
            self._kernel32.CloseHandle(self._handle)
            self._handle = None


def open_exact_process(pid: int, create_time: int | None) -> ExactProcess | None:
    """Pin identity before observation or force. Unknown identity is refused."""
    if sys.platform != "win32" or not pid or create_time is None:
        return None
    from .instance_registry import FILETIME, _coerce_filetime

    kernel32 = _kernel32()
    handle = kernel32.OpenProcess(
        SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION | PROCESS_TERMINATE,
        False, int(pid),
    )
    if not handle:
        return None
    try:
        try:
            import win32process
            times = win32process.GetProcessTimes(handle)
            stamp = _coerce_filetime(times["CreationTime"] if isinstance(times, dict) else times[0])
        except ImportError:
            kernel32.GetProcessTimes.argtypes = [ctypes.c_void_p] + [ctypes.POINTER(FILETIME)] * 4
            kernel32.GetProcessTimes.restype = ctypes.c_bool
            creation, exit_time, kernel, user = (FILETIME() for _ in range(4))
            if not kernel32.GetProcessTimes(handle, ctypes.byref(creation), ctypes.byref(exit_time),
                                            ctypes.byref(kernel), ctypes.byref(user)):
                return None
            stamp = (int(creation.dwHighDateTime) << 32) | int(creation.dwLowDateTime)
        if stamp != create_time:
            return None
        process = ExactProcess(kernel32, handle)
        handle = None
        return process
    except Exception:
        return None
    finally:
        if handle is not None:
            kernel32.CloseHandle(handle)


def prefer_full_power_pid(pid: int) -> bool:
    """Disable EcoQoS and raise priority on ``pid``. Best-effort.

    Returns True when at least the priority change succeeded.
    """
    if sys.platform != "win32" or not pid or pid <= 0:
        return False

    kernel32 = _kernel32()
    access = PROCESS_SET_INFORMATION | PROCESS_QUERY_LIMITED_INFORMATION
    handle = kernel32.OpenProcess(access, False, int(pid))
    if not handle:
        return False

    qos_ok = False
    priority_ok = False
    try:
        state = PROCESS_POWER_THROTTLING_STATE(
            Version=PROCESS_POWER_THROTTLING_CURRENT_VERSION,
            ControlMask=PROCESS_POWER_THROTTLING_EXECUTION_SPEED,
            StateMask=0,
        )
        qos_ok = bool(
            kernel32.SetProcessInformation(
                handle,
                PROCESS_POWER_THROTTLING,
                ctypes.byref(state),
                ctypes.sizeof(state),
            )
        )
        priority_ok = bool(
            kernel32.SetPriorityClass(handle, ABOVE_NORMAL_PRIORITY_CLASS)
        )
    except Exception:
        return False
    finally:
        kernel32.CloseHandle(handle)

    if qos_ok or priority_ok:
        try:
            from ..usage_logging import log_diagnostic_event

            log_diagnostic_event(
                "access_full_power_qos",
                pid=int(pid),
                ecoqos_off=qos_ok,
                above_normal=priority_ok,
            )
        except Exception:
            pass
    return bool(priority_ok or qos_ok)


def pid_from_access_app(app: Any) -> int | None:
    """Return the process id for an Access Application COM object."""
    try:
        hwnd = app.hWndAccessApp()
    except Exception:
        return None
    if isinstance(hwnd, bool) or not isinstance(hwnd, int):
        return None
    if hwnd <= 0:
        return None

    try:
        import win32process

        _tid, pid = win32process.GetWindowThreadProcessId(hwnd)
        return int(pid) if pid else None
    except Exception:
        pass

    if sys.platform != "win32":
        return None
    try:
        pid = ctypes.c_ulong(0)
        ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        return int(pid.value) if pid.value else None
    except Exception:
        return None


def prefer_full_power_app(app: Any) -> bool:
    """Apply full-power QoS to the process hosting ``app``."""
    pid = pid_from_access_app(app)
    if not pid:
        return False
    return prefer_full_power_pid(pid)


def prefer_full_power_if_created(app: Any) -> bool:
    """Promote a COM Access instance only when it has no current database.

    ``EnsureDispatch`` can return the user's instance. A database already
    open means this is not an MCP-created empty process.
    """
    try:
        existing = app.CurrentDb()
    except Exception:
        existing = None
    if existing is not None:
        return False
    return prefer_full_power_app(app)


def list_access_pids_or_none() -> set[int] | None:
    """Return PIDs of running ``MSACCESS.EXE``, or None when unknowable.

    ``None`` means the query itself failed, which is not the same as "no
    Access is running".  The owned-instance registry prunes on this set,
    and treating a failed ``tasklist`` as an empty result would forget
    every window the server created.
    """
    if sys.platform != "win32":
        return set()
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq MSACCESS.EXE", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=5,
            check=False,
        )
    except Exception:
        return None

    if result.returncode != 0:
        return None

    stdout = result.stdout or ""
    if not stdout.strip() or "No tasks" in stdout or stdout.lstrip().startswith("INFO:"):
        return set()

    pids: set[int] = set()
    for line in stdout.splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            row = next(csv.reader(StringIO(line)))
        except StopIteration:
            continue
        if len(row) < 2:
            continue
        try:
            pids.add(int(row[1]))
        except ValueError:
            continue
    return pids


def list_access_pids() -> set[int]:
    """Return PIDs of running ``MSACCESS.EXE`` processes. Best-effort.

    A failed query reports an empty set. Callers that must not confuse
    "query failed" with "nothing running" use ``list_access_pids_or_none``.
    """
    pids = list_access_pids_or_none()
    return set() if pids is None else pids


def process_is_alive(pid: int) -> bool | None:
    """True/False when ``pid`` liveness is known, None when it is not."""
    pids = list_access_pids_or_none()
    if pids is None:
        return None
    return int(pid) in pids


def terminate_pid(pid: int) -> bool:
    """Force-terminate ``pid``. Returns True when the process is gone.

    Last resort for a server-created Access instance that failed both a
    recovery probe and a time-bounded ``Quit``. Callers must verify the
    PID belongs to an instance the server created before calling this.
    """
    if sys.platform != "win32" or not pid or pid <= 0:
        return False

    kernel32 = _kernel32()
    handle = kernel32.OpenProcess(PROCESS_TERMINATE, False, int(pid))
    if not handle:
        # Already gone is success; anything else we cannot confirm.
        return process_is_alive(pid) is False
    try:
        ok = bool(kernel32.TerminateProcess(handle, 1))
    except Exception:
        return False
    finally:
        kernel32.CloseHandle(handle)

    if not ok:
        return False

    # TerminateProcess is asynchronous; confirm the process actually exited.
    deadline = time.monotonic() + 5.0
    while time.monotonic() < deadline:
        alive = process_is_alive(pid)
        if alive is False:
            return True
        if alive is None:
            return False
        time.sleep(0.2)
    return False


def prefer_full_power_new_access(known_pids: set[int]) -> set[int]:
    """Promote Access processes that were not in ``known_pids``.

    Returns the updated set of PIDs that have been seen (promoted or not)
    so later calls skip them. Pre-existing user instances stay in
    ``known_pids`` and are never touched.
    """
    current = list_access_pids()
    for pid in current - known_pids:
        prefer_full_power_pid(pid)
    return known_pids | current
