"""Hung gate-exempt workers stay bounded and never starve the Access gate (M24)."""

from __future__ import annotations

import asyncio
import os
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from contextlib import contextmanager

import pytest

from msaccess_vcs_mcp import exempt_workers, tools
from msaccess_vcs_mcp.dialog_recovery import (
    CLICK_NOT_SENT,
    CLICK_UNCERTAIN,
    ButtonInfo,
    Win32Backend,
)
from msaccess_vcs_mcp.exempt_workers import (
    MAX_WAITING_PER_TOOL,
    MAX_WORKERS_PER_TOOL,
    ExemptWorkers,
    WorkerCapacityUnavailable,
)

from .test_dialog_dispatch import Blocker, _install_backend, public_tools  # noqa: F401
from .test_dialog_recovery import FakeBackend, _button, _win

DEADLINE = 1.0
FAILSAFE = 5.0
DIALOG_TIMEOUT = 0.3
SLACK = 0.5


class HangingBackend(FakeBackend):
    """Every window listing blocks until released, like a hung Access UI thread."""

    def __init__(self, pid=10):
        super().__init__([
            _win(hwnd=1, pid=pid, title="Northwind : Database", class_name="OMain"),
            _win(hwnd=2, pid=pid, title="VCS Probe", texts=("Hello",), buttons=(_button(21, "OK"),)),
        ])
        self.release = threading.Event()
        self._lock = threading.Lock()
        self.active = 0
        self.peak = 0
        self.entered = 0
        self.threads: set[str] = set()

    def list_windows(self):
        with self._lock:
            self.active += 1
            self.entered += 1
            self.peak = max(self.peak, self.active)
            self.threads.add(threading.current_thread().name)
        try:
            self.release.wait(FAILSAFE)
        finally:
            with self._lock:
                self.active -= 1
        return super().list_windows()


def _list_dialogs(db):
    return tools.vcs_list_dialogs(db, pid=10, create_time=1000, timeout_seconds=DIALOG_TIMEOUT)


async def _timed(coro):
    started = time.monotonic()
    result = await coro
    return result, time.monotonic() - started


def _exempt_threads() -> list[str]:
    return [t.name for t in threading.enumerate() if t.name.startswith("vcs-exempt-")]


@pytest.fixture
def short_ceiling(monkeypatch):
    # The ceiling is then the dialog timeout itself.
    monkeypatch.setattr(tools, "EXEMPT_WORKER_MARGIN_SEC", 0.0)


@pytest.mark.parametrize("gate_busy", [False, True], ids=["gate_free", "gate_busy"])
def test_hung_dialog_workers_do_not_starve_gate_on_one_thread_executor(
    public_tools, monkeypatch, short_ceiling, gate_busy,  # noqa: F811
):
    """Before M24 each abandoned dialog worker kept the default executor's only
    thread, and the gate's slot wait queued behind it with no timeout running."""
    db, addin = public_tools
    backend = HangingBackend()
    _install_backend(monkeypatch, backend)
    gate_blocker = Blocker()

    def get_option(*_args):
        gate_blocker.wait()
        return "fake option value"

    async def scenario():
        asyncio.get_running_loop().set_default_executor(ThreadPoolExecutor(max_workers=1))
        tasks = []
        try:
            if gate_busy:
                addin.call_sync.side_effect = get_option
                tasks.append(asyncio.create_task(tools.vcs_get_option(db, "ShowDebug")))
                assert await asyncio.to_thread(gate_blocker.entered.wait, DEADLINE)
            for _ in range(MAX_WORKERS_PER_TOOL):
                hung = await asyncio.wait_for(_list_dialogs(db), DEADLINE)
                assert hung["error_pattern"] == "tool_timeout"
            assert backend.active == MAX_WORKERS_PER_TOOL

            gated, elapsed = await asyncio.wait_for(
                _timed(tools.vcs_get_option(db, "ShowDebug")), DEADLINE
            )
            if gate_busy:
                assert gated["error_pattern"] == "server_busy"
                assert gated["busy_with"]["tool"] == "vcs_get_option"
            else:
                assert gated == {
                    "success": True, "option": "ShowDebug", "value": "fake option value",
                }
            assert elapsed < 0.05 + SLACK
            assert backend.active == MAX_WORKERS_PER_TOOL, "the hung workers must still be hung"
        finally:
            backend.release.set()
            gate_blocker.release.set()
            await asyncio.gather(*tasks, return_exceptions=True)

    asyncio.run(scenario())


def test_repeated_hung_calls_have_bounded_workers_and_waiters(
    public_tools, monkeypatch, short_ceiling,  # noqa: F811
):
    db, _addin = public_tools
    backend = HangingBackend()
    _install_backend(monkeypatch, backend)
    calls = MAX_WORKERS_PER_TOOL + MAX_WAITING_PER_TOOL + 6

    async def scenario():
        try:
            results = await asyncio.wait_for(
                asyncio.gather(*(_timed(_list_dialogs(db)) for _ in range(calls))),
                DIALOG_TIMEOUT + DEADLINE,
            )
            patterns = [result["error_pattern"] for result, _elapsed in results]
            assert patterns.count("tool_timeout") == MAX_WORKERS_PER_TOOL
            assert patterns.count("worker_capacity_unavailable") == calls - MAX_WORKERS_PER_TOOL
            refused = [(r, e) for r, e in results if r["error_pattern"] == "worker_capacity_unavailable"]
            assert all(r["recoverable"] is True and r["success"] is False for r, _e in refused)
            # Waiters use the same deadline; the rest are refused at once.
            assert all(e < DIALOG_TIMEOUT + SLACK for _r, e in results)
            assert sum(1 for _r, e in refused if e < DIALOG_TIMEOUT / 2) == (
                calls - MAX_WORKERS_PER_TOOL - MAX_WAITING_PER_TOOL
            )
            assert backend.entered == MAX_WORKERS_PER_TOOL
            assert backend.peak == MAX_WORKERS_PER_TOOL
            assert len(_exempt_threads()) == MAX_WORKERS_PER_TOOL

            # Still hung: a later call waits out its own deadline and starts no thread.
            late, elapsed = await asyncio.wait_for(_timed(_list_dialogs(db)), DEADLINE)
            assert late["error_pattern"] == "worker_capacity_unavailable"
            assert DIALOG_TIMEOUT - 0.05 <= elapsed < DIALOG_TIMEOUT + SLACK
            assert backend.entered == MAX_WORKERS_PER_TOOL

            # Another exempt tool has its own budget.
            recent = await asyncio.wait_for(tools.vcs_get_recent_calls(), DEADLINE)
            assert recent.get("error_pattern") not in {"worker_capacity_unavailable", "tool_timeout"}
        finally:
            backend.release.set()

        assert await asyncio.to_thread(exempt_workers.wait_idle, DEADLINE)
        again = await asyncio.wait_for(_list_dialogs(db), DEADLINE)
        assert again["success"] is True
        assert [item["dialog_id"] for item in again["dialogs"]] == ["hwnd:2"]

    asyncio.run(scenario())
    # Off the event loop and off the COM apartment thread.
    assert backend.threads == {"vcs-exempt-vcs_list_dialogs"}


def test_worker_error_propagates_and_frees_its_place():
    workers = ExemptWorkers("probe", max_workers=1, max_waiting=0)

    def boom():
        raise ValueError("boom")

    async def scenario():
        with pytest.raises(ValueError, match="boom"):
            await workers.run(boom, (), {}, time.monotonic() + DEADLINE)
        assert workers.running == 0
        assert await workers.run(lambda: "ok", (), {}, time.monotonic() + DEADLINE) == "ok"

    asyncio.run(scenario())


def test_abandoned_worker_keeps_its_place_until_it_returns():
    workers = ExemptWorkers("probe", max_workers=1, max_waiting=0)
    release = threading.Event()

    async def scenario():
        try:
            with pytest.raises(asyncio.TimeoutError):
                await workers.run(release.wait, (FAILSAFE,), {}, time.monotonic() + 0.1)
            assert workers.running == 1
            with pytest.raises(WorkerCapacityUnavailable):
                await workers.run(lambda: "ok", (), {}, time.monotonic() + DEADLINE)
        finally:
            release.set()
        await asyncio.to_thread(exempt_workers.wait_idle, DEADLINE)
        assert workers.running == 0
        assert await workers.run(lambda: "ok", (), {}, time.monotonic() + DEADLINE) == "ok"

    asyncio.run(scenario())


# The real click timeout: a button window whose thread never pumps messages.

win32 = pytest.mark.skipif(sys.platform != "win32", reason="Win32 message delivery")


@contextmanager
def _hung_button():
    win32api = pytest.importorskip("win32api")
    win32con = pytest.importorskip("win32con")
    win32gui = pytest.importorskip("win32gui")
    ready = threading.Event()
    release = threading.Event()
    created: dict[str, int] = {}

    def owner():
        created["hwnd"] = win32gui.CreateWindow(
            "BUTTON", "OK", win32con.WS_POPUP, 0, 0, 10, 10, 0, 0,
            win32api.GetModuleHandle(None), None,
        )
        ready.set()
        release.wait(FAILSAFE * 6)
        win32gui.DestroyWindow(created["hwnd"])

    thread = threading.Thread(target=owner, daemon=True)
    thread.start()
    assert ready.wait(FAILSAFE)
    try:
        yield created["hwnd"]
    finally:
        release.set()
        thread.join(FAILSAFE)


@win32
def test_real_bm_click_on_hung_window_returns_uncertain_at_its_timeout():
    with _hung_button() as hwnd:
        started = time.monotonic()
        outcome = Win32Backend().click(
            ButtonInfo(hwnd, "OK"), expected_pid=os.getpid(), timeout_ms=300
        )
        elapsed = time.monotonic() - started
    assert outcome == CLICK_UNCERTAIN
    # SendMessageTimeout waited for the button and gave up at its own timeout.
    assert 0.25 <= elapsed < 0.3 + SLACK


@win32
def test_real_netui_press_on_hung_window_is_not_sent(monkeypatch):
    from msaccess_vcs_mcp import msaa

    def press(*_args, **_kwargs):
        raise AssertionError("a hung window is not asked for its accessible object")

    monkeypatch.setattr(msaa, "press", press)
    with _hung_button() as hwnd:
        started = time.monotonic()
        outcome = Win32Backend().click(
            ButtonInfo(hwnd, "OK", path=(1,)), expected_pid=os.getpid(), timeout_ms=300
        )
        elapsed = time.monotonic() - started
    assert outcome == CLICK_NOT_SENT
    assert 0.25 <= elapsed < 0.3 + SLACK


class RealClickBackend(FakeBackend):
    """Fake windows, real Win32 button message."""

    def click(self, button, *, expected_pid=None, timeout_ms=5000):
        self.clicked.append(button.hwnd)
        return Win32Backend().click(button, expected_pid=expected_pid, timeout_ms=timeout_ms)


@win32
def test_public_dismiss_on_hung_button_is_bounded_by_the_real_click_timeout(
    public_tools, monkeypatch,  # noqa: F811
):
    db, _addin = public_tools
    pid = os.getpid()

    with _hung_button() as hwnd:
        backend = RealClickBackend([
            _win(hwnd=1, pid=pid, title="Northwind : Database", class_name="OMain"),
            _win(hwnd=2, pid=pid, title="VCS Probe", texts=("Hello",), buttons=(_button(hwnd, "OK"),)),
        ])
        _install_backend(monkeypatch, backend)

        async def scenario():
            return await asyncio.wait_for(
                _timed(tools.vcs_dismiss_dialog(
                    db, "hwnd:2", button="OK", pid=pid, create_time=1000,
                    timeout_seconds=DIALOG_TIMEOUT,
                )),
                tools.EXEMPT_WORKER_MARGIN_SEC + DIALOG_TIMEOUT,
            )

        result, elapsed = asyncio.run(scenario())

    # The worker answered on its own, well inside the ceiling: no tool_timeout.
    assert result["success"] is False
    assert result["error_pattern"] == "dismiss_uncertain"
    assert backend.clicked == [hwnd]
    assert DIALOG_TIMEOUT - 0.05 <= elapsed < 2 * DIALOG_TIMEOUT + SLACK
