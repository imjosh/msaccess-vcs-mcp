"""M45/X16: isolated metadata stays off the loop and never attaches to Access."""
import asyncio
import os
from pathlib import Path
import sys
import threading
import time
from unittest.mock import MagicMock

import pytest

from msaccess_vcs_mcp import compatibility, exempt_workers, tools, validation
from .test_dialog_dispatch import public_tools, _install_backend, _dialog_call, DIALOG_TOOLS
from .test_dialog_recovery import FakeBackend, _button, _win


@pytest.mark.parametrize("completion", ["completed", "caller_cancelled", "timeout"])
def test_version_probe_keeps_recovery_responsive_without_touching_access(
    public_tools, monkeypatch, completion,
):
    if completion == "timeout":
        monkeypatch.setattr(tools, "VERSION_INFO_WORKER_TIMEOUT_SEC", 0.1)
    db, _gated_addin = public_tools
    entered, release = threading.Event(), threading.Event()
    stages = []
    access = MagicMock(side_effect=AssertionError("Metadata touched Access"))
    monkeypatch.setattr(validation, "create_isolated_access_app", access)
    monkeypatch.setattr(validation.win32com.client, "GetObject", access)
    def inspect(*, fresh=False):
        # Option admission must remain independent of the metadata worker.
        if threading.current_thread().name.startswith("vcs-exempt-vcs_get_version_info"):
            assert fresh is True
            stages.append(("entered", threading.get_ident()))
            entered.set()
            assert release.wait(6), "version probe blocked the server loop"
            stages.append(("returned", threading.get_ident()))
        return {**compatibility.compatibility_result("6.0.0"), "addin_path": "mock-install.accda"}
    monkeypatch.setattr(tools, "inspect_installed_addin", inspect)
    backend = FakeBackend([
        _win(hwnd=1, title="Northwind : Database", class_name="OMain"),
        _win(hwnd=2, title="VCS Probe", texts=("Hello",), buttons=(_button(21, "OK"),)),
    ])
    _install_backend(monkeypatch, backend)
    async def scenario():
        loop_thread = threading.get_ident()
        task = asyncio.create_task(tools.vcs_get_version_info())
        try:
            assert await asyncio.to_thread(entered.wait, 1)
            assert (await asyncio.wait_for(tools.vcs_get_option(db, "ShowDebug"), 1))["success"]
            for tool in DIALOG_TOOLS:
                started = time.monotonic()
                result = await asyncio.wait_for(_dialog_call(tool, db, timeout=0.5), 0.5)
                assert time.monotonic() - started < 0.5
                assert result["success"]
            if completion == "timeout":
                result = await task
                assert result["error_pattern"] == "tool_timeout"
                assert result["mcp_version"] == "0.3.0-dev.18"
                assert result["server_runtime"]["pid"] == os.getpid()
                assert result["server_runtime"]["python_executable"] == sys.executable
                assert result["server_runtime"]["package_root"] == str(Path(compatibility.__file__).resolve().parent)
                assert result["supported_addin_range"] == compatibility.ADDIN_REQUIREMENT.text
            else:
                assert not task.done()
            if completion == "caller_cancelled":
                task.cancel()
                with pytest.raises(asyncio.CancelledError):
                    await task
        finally:
            release.set()
            if completion == "completed":
                assert (await task)["vcs_version"] == "6.0.0"
            assert await asyncio.to_thread(exempt_workers.wait_idle, 2)
        assert [stage for stage, _ in stages] == ["entered", "returned"]
        assert len({thread for _, thread in stages}) == 1
        assert stages[0][1] != loop_thread
    asyncio.run(scenario())
    assert access.call_count == 0
