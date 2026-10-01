"""M45: the version probe owns its COM sequence off the server loop."""
import asyncio
import threading
import time
from unittest.mock import MagicMock

import pytest
import pythoncom

from msaccess_vcs_mcp import exempt_workers, tools, validation
from .test_dialog_dispatch import public_tools, _install_backend, _dialog_call, DIALOG_TOOLS
from .test_dialog_recovery import FakeBackend, _button, _win


@pytest.mark.parametrize("owned", [False, True], ids=["attached", "created"])
@pytest.mark.parametrize("completion", ["completed", "caller_cancelled", "timeout"])
def test_version_probe_keeps_recovery_responsive_and_cleanup_on_owner_thread(
    public_tools, monkeypatch, owned, completion,
):
    if completion == "timeout":
        monkeypatch.setattr(tools, "VERSION_INFO_WORKER_TIMEOUT_SEC", 0.1)
    db, _gated_addin = public_tools
    addin = MagicMock()
    entered = threading.Event()
    release = threading.Event()
    stages = []
    app = MagicMock()
    app.CurrentDb.return_value.Name = db

    def record(stage):
        stages.append((stage, threading.get_ident()))

    def attach(*args):
        record("attach")
        entered.set()
        assert release.wait(6), "version probe blocked the server event loop"
        if owned:
            raise RuntimeError("no running target")
        return app

    def create():
        record("create")
        return app

    def access_info(*args):
        record("use")
        return {}

    def apartment(stage):
        # The overlapping gate call initializes its own apartment as well.
        if threading.current_thread().name.startswith("vcs-exempt-"):
            record(stage)

    monkeypatch.setattr(pythoncom, "CoInitialize", lambda: apartment("initialize"))
    monkeypatch.setattr(pythoncom, "CoUninitialize", lambda: apartment("uninitialize"))
    monkeypatch.setattr(validation.win32com.client, "GetObject", attach)
    monkeypatch.setattr(validation, "create_isolated_access_app", create)
    monkeypatch.setattr(validation, "get_config", lambda: {"ACCESS_VCS_DATABASE": db})
    monkeypatch.setattr(validation, "ensure_access_visible", lambda *_: record("visible"))
    monkeypatch.setattr(validation, "open_current_database", lambda *_: record("open"))
    monkeypatch.setattr(validation, "get_access_info", access_info)
    monkeypatch.setattr(validation, "VCSAddinIntegration", lambda *_: addin)
    addin.verify_addin_exists.return_value = True
    addin.load_addin.side_effect = lambda *a, **kw: record("load")
    addin.get_version_info.side_effect = lambda *_: (record("version") or {"vcs_version": "test"})
    app.CloseCurrentDatabase.side_effect = lambda: record("close")
    app.Quit.side_effect = lambda: record("quit")
    backend = FakeBackend([
        _win(hwnd=1, title="Northwind : Database", class_name="OMain"),
        _win(hwnd=2, title="VCS Probe", texts=("Hello",), buttons=(_button(21, "OK"),)),
    ])
    _install_backend(monkeypatch, backend)

    async def scenario():
        server_thread = threading.get_ident()
        task = asyncio.create_task(tools.vcs_get_version_info())
        try:
            assert await asyncio.to_thread(entered.wait, 1)
            # The gate can operate on its own apartment while the probe is blocked.
            result = await asyncio.wait_for(tools.vcs_get_option(db, "ShowDebug"), 1)
            assert result["success"] is True
            for tool in DIALOG_TOOLS:
                started = time.monotonic()
                result = await asyncio.wait_for(_dialog_call(tool, db, timeout=0.5), 0.5)
                assert time.monotonic() - started < 0.5
                assert result["success"] is True
            result = await asyncio.wait_for(tools.vcs_cancel_operation("no-such-operation"), 0.5)
            assert result.get("error_pattern") != "tool_timeout"
            await asyncio.sleep(3)
            if completion == "timeout":
                assert (await task)["error_pattern"] == "tool_timeout"
            else:
                assert not task.done()
            if completion == "caller_cancelled":
                task.cancel()
                with pytest.raises(asyncio.CancelledError):
                    await task
        finally:
            release.set()
            if completion == "completed":
                result = await task
                assert result["vcs_version"] == "test"
            assert await asyncio.to_thread(exempt_workers.wait_idle, 2)
        assert stages[0][0] == "initialize"
        assert stages[-1][0] == "uninitialize"
        assert len({thread for _, thread in stages}) == 1
        assert stages[0][1] != server_thread

    asyncio.run(scenario())
    assert app.Quit.call_count == int(owned)
    assert app.CloseCurrentDatabase.call_count == int(owned)
