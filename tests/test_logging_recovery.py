"""Inspect actual records after rotation refusal and file unavailability."""

import ctypes
import asyncio
import json
import os
import subprocess
import sys
from pathlib import Path
from unittest.mock import patch

import pytest

from msaccess_vcs_mcp import usage_logging as logs


@pytest.fixture(autouse=True)
def logging_state(tmp_path, monkeypatch):
    logs.reset_logging()
    monkeypatch.setenv("ACCESS_VCS_ENABLE_LOGGING", "true")
    monkeypatch.setenv("ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG", "false")
    monkeypatch.setenv("ACCESS_VCS_LOG_DIR", str(tmp_path))
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(tmp_path))
    yield
    logs.reset_logging()


def records(path):
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines()]


def windows_read_handle(path, share_mode):
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.CreateFileW.argtypes = [ctypes.c_wchar_p, ctypes.c_uint32, ctypes.c_uint32,
                                  ctypes.c_void_p, ctypes.c_uint32, ctypes.c_uint32,
                                  ctypes.c_void_p]
    kernel.CreateFileW.restype = ctypes.c_void_p
    kernel.CloseHandle.argtypes = [ctypes.c_void_p]
    handle = kernel.CreateFileW(str(path), 0x80000000, share_mode, None, 3, 0, None)
    assert handle != ctypes.c_void_p(-1).value
    return kernel, handle


@pytest.mark.skipif(os.name != "nt", reason="Windows rename sharing semantics")
@pytest.mark.parametrize("diagnostic", [False, True])
def test_windows_rename_refusal_recovers_without_restart(diagnostic, capsys):
    if diagnostic:
        logs.log_diagnostic_event("before_lock")
        handler = logs._diagnostic_handler
        path = logs.get_diagnostic_log_path()
    else:
        logs.log_tool_call("before_lock", {})
        handler = logs._log_handler
        path = logs.get_log_file_path()
    handler.maxBytes = 1
    # Share reads/writes, but deny deletion (and therefore rename).
    kernel, lock = windows_read_handle(path, 3)
    try:
        logs._write_log_entry({"event": "locked_rotation"}, handler)
        logs._write_log_entry({"event": "next_record"}, handler)
        assert [r["event"] for r in records(path)][-2:] == ["locked_rotation", "next_record"]
        assert "rotation" in capsys.readouterr().err.lower()
    finally:
        kernel.CloseHandle(lock)
    # Advance retry time without sleeping or restarting the handler.
    with patch.object(logs.time, "monotonic", return_value=10**12):
        logs._write_log_entry({"event": "after_unlock"}, handler)
    backup = Path(str(path) + ".1")
    assert backup.exists()
    assert records(backup)[-1]["event"] == "after_unlock"
    handler.maxBytes = 0
    logs._write_log_entry({"event": "still_writing"}, handler)
    assert records(path)[-1]["event"] == "still_writing"


@pytest.mark.skipif(os.name != "nt", reason="Windows write sharing semantics")
@pytest.mark.parametrize("diagnostic", [False, True])
def test_windows_lock_denies_reopen_then_recovers(diagnostic, capsys):
    if diagnostic:
        logs.log_diagnostic_event("before_lock")
        handler = logs._diagnostic_handler
        path = logs.get_diagnostic_log_path()
    else:
        logs.log_tool_call("before_lock", {})
        handler = logs._log_handler
        path = logs.get_log_file_path()
    handler.maxBytes = 1
    rotate = handler.rotate
    handles = []

    def lock_then_rotate(source, dest):
        # Rollover has closed the writer, so Windows permits a reader which
        # denies both writes and deletion. Keep it through subsequent calls.
        handles.append(windows_read_handle(path, 1))
        rotate(source, dest)

    try:
        with patch.object(handler, "rotate", side_effect=lock_then_rotate):
            logs._write_log_entry({"event": "rotation_trigger"}, handler)
            logs._write_log_entry({"event": "native_fallback"}, handler)
        output = capsys.readouterr().err
        assert "rotation" in output and "reopen" in output
        assert "WinError 32" in output
        fallback = [json.loads(line) for line in output.splitlines() if line.startswith("{")]
        assert [r["event"] for r in fallback] == ["native_fallback"]
        assert records(path)[-1]["event"] == "rotation_trigger"
    finally:
        for kernel, handle in handles:
            kernel.CloseHandle(handle)
    handler.maxBytes = 0
    with patch.object(logs.time, "monotonic", return_value=10**12):
        logs._write_log_entry({"event": "native_recovery"}, handler)
    assert records(path)[-1]["event"] == "native_recovery"
    assert capsys.readouterr().err == ""


@pytest.mark.parametrize("diagnostic", [False, True])
def test_unavailable_file_falls_back_and_recovers_with_bounded_retries(diagnostic, capsys):
    if diagnostic:
        logs.log_diagnostic_event("before_failure")
        handler = logs._diagnostic_handler
        path = logs.get_diagnostic_log_path()
        write = lambda i: logs.log_diagnostic_event("unavailable", index=i)
    else:
        logs.log_tool_call("before_failure", {})
        handler = logs._log_handler
        path = logs.get_log_file_path()

        @logs.with_logging("unavailable")
        def tool(index):
            return {"success": True, "index": index}

        def write(i):
            assert tool(i) == {"success": True, "index": i}

    handler.maxBytes = 1
    # Model the interval after doRollover closes the current file, with both
    # rename and reopening unavailable. Inspect stderr records independently.
    with patch.object(logs.time, "monotonic", return_value=100), \
         patch.object(handler, "rotate", side_effect=PermissionError("rename denied")), \
         patch.object(handler, "_open", side_effect=PermissionError("write denied")) as reopen:
        write(-1)  # Flushed before rotation; must not be duplicated to stderr.
        for i in range(20):
            write(i)
        assert reopen.call_count == 1
        output = capsys.readouterr().err.splitlines()
        warnings = [line for line in output if line.startswith("Warning:")]
        assert len(warnings) == 2
        assert any("rotation" in line and str(path) in line for line in warnings)
        assert any("reopen" in line and "write denied" in line for line in warnings)
        fallback = [json.loads(line) for line in output if line.startswith("{")]
        assert len(fallback) == 20
        assert [r.get("index", r.get("parameters", {}).get("index")) for r in fallback] == list(range(20))
        assert all(r["instance_id"] == logs._INSTANCE_ID for r in fallback)
        with patch.object(logs.time, "monotonic", return_value=105):
            write(20)
        assert reopen.call_count == 2
        assert "write denied" in capsys.readouterr().err
    handler.maxBytes = 0
    with patch.object(logs.time, "monotonic", return_value=110):
        write(21)
    assert records(path)[-1].get("index", records(path)[-1].get("parameters", {}).get("index")) == 21
    assert capsys.readouterr().err == ""


def test_write_error_falls_back_then_reopens(tmp_path, capsys):
    logs.log_tool_call("before_failure", {})
    handler = logs._log_handler
    path = logs.get_log_file_path()
    with patch.object(logs.time, "monotonic", return_value=100), \
         patch.object(handler.stream, "write", side_effect=OSError("disk unavailable")):
        logs.log_tool_call("failed_write", {"index": 1})
        logs.log_tool_call("during_cooldown", {"index": 2})
    output = capsys.readouterr().err.splitlines()
    assert sum(line.startswith("Warning:") for line in output) == 1
    assert "disk unavailable" in output[0]
    fallback = [json.loads(line) for line in output if line.startswith("{")]
    assert [r["tool"] for r in fallback] == ["failed_write", "during_cooldown"]
    with patch.object(logs.time, "monotonic", return_value=105):
        logs.log_tool_call("recovered_write", {})
    assert records(path)[-1]["tool"] == "recovered_write"
    assert capsys.readouterr().err == ""


@pytest.mark.parametrize("diagnostic", [False, True])
def test_initially_unavailable_file_keeps_path_and_recovers(diagnostic, capsys):
    with patch.object(logs.time, "monotonic", return_value=100), \
         patch.object(logs._RecoveringRotatingFileHandler, "_open", side_effect=PermissionError("open denied")) as reopen:
        if diagnostic:
            logs.log_diagnostic_event("initially_unavailable")
            path = logs.get_diagnostic_log_path()
        else:
            logs.log_tool_call("initially_unavailable", {})
            path = logs.get_log_file_path()
        assert path is not None
        assert reopen.call_count == 1
    output = capsys.readouterr().err
    assert "open denied" in output
    assert "initially_unavailable" in output
    with patch.object(logs.time, "monotonic", return_value=105):
        if diagnostic:
            logs.log_diagnostic_event("initially_recovered")
        else:
            logs.log_tool_call("initially_recovered", {})
    assert records(path)[-1].get("tool", records(path)[-1]["event"]) == "initially_recovered"


def test_two_processes_have_distinct_reported_paths_and_rotated_records(tmp_path):
    script = '''
import json, sys
from msaccess_vcs_mcp import usage_logging as logs
logs.log_tool_call("start", {})
logs.log_diagnostic_event("server_start")
paths = [logs.get_log_file_path(), logs.get_diagnostic_log_path()]
print(json.dumps({"paths": [str(p) for p in paths], "instance_id": logs._INSTANCE_ID}), flush=True)
sys.stdin.readline()
for handler in (logs._log_handler, logs._diagnostic_handler):
    handler.maxBytes = 1
    handler.backupCount = 30
for i in range(10):
    logs.log_tool_call("probe", {"index": i})
    logs.log_diagnostic_event("probe", index=i)
logs.reset_logging()
logs.log_tool_call("after_reload", {})
logs.log_diagnostic_event("after_reload")
assert paths == [logs.get_log_file_path(), logs.get_diagnostic_log_path()]
logs.reset_logging()
'''
    children = []
    try:
        children = [subprocess.Popen([sys.executable, "-c", script], stdin=subprocess.PIPE,
                                    stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
                    for _ in range(2)]
        reports = [json.loads(child.stdout.readline()) for child in children]
        assert reports[0]["instance_id"] != reports[1]["instance_id"]
        assert set(reports[0]["paths"]).isdisjoint(reports[1]["paths"])
        for child in children:
            child.stdin.write("rotate\n")
            child.stdin.flush()
        for child in children:
            _, stderr = child.communicate(timeout=20)
            assert child.returncode == 0, stderr
            assert stderr == ""
        for report in reports:
            for name in report["paths"]:
                path = Path(name)
                assert path.parent == tmp_path
                assert report["instance_id"] in path.name
                entries = [entry for file in path.parent.glob(path.name + "*") for entry in records(file)]
                probes = [r for r in entries if r.get("tool", r["event"]) == "probe"]
                assert sorted(r.get("index", r.get("parameters", {}).get("index")) for r in probes) == list(range(10))
                assert all(r["instance_id"] == report["instance_id"] for r in entries)
                assert len({r["server_pid"] for r in entries}) == 1
                assert records(path)[-1].get("tool", records(path)[-1]["event"]) == "after_reload"
    finally:
        for child in children:
            if child.poll() is None:
                child.kill()
            child.wait(timeout=10)


def test_two_stdio_servers_report_paths_and_complete_calls_during_failure(tmp_path):
    from mcp import ClientSession, StdioServerParameters
    from mcp.client.stdio import stdio_client

    # Only Access inspection and callback startup are controlled. Run the real
    # server entry point, stdio transport, metadata tool and logging decorator.
    script = '''
from unittest.mock import patch
from msaccess_vcs_mcp import main, tools, usage_logging as logs
main.get_config = lambda: {"ACCESS_VCS_ENABLE_ASYNC": False}
main.validate_access_installation = lambda: None
main.validate_all_components = lambda: {"success": True}
main._start_callback_server = lambda config: None
tools._probe_version_info = lambda: {"success": True}

@tools.mcp.tool()
@logs.with_logging("logging_probe")
def logging_probe(index: int, unavailable: bool = False) -> dict:
    if index == 0:
        logs._log_handler.maxBytes = 1
        logs._diagnostic_handler.maxBytes = 1
        logs._log_handler.backupCount = 30
        logs._diagnostic_handler.backupCount = 30
    if unavailable:
        handler = logs._log_handler
        if handler.stream is not None:
            handler.stream.close()
        handler.stream = None
        with patch.object(handler, "_open", side_effect=PermissionError("test file unavailable")):
            logs.log_tool_call("unavailable_probe", {"index": index})
    logs.log_diagnostic_event("probe", index=index)
    return {"success": True, "index": index}

main.main()
'''

    async def run_server(number):
        stderr_path = tmp_path / f"server-{number}.stderr"
        with stderr_path.open("w", encoding="utf-8") as stderr:
            params = StdioServerParameters(command=sys.executable, args=["-c", script], env=dict(os.environ))
            async with stdio_client(params, errlog=stderr) as (reader, writer):
                async with ClientSession(reader, writer) as session:
                    await session.initialize()
                    result = await session.call_tool("vcs_get_version_info", {})
                    assert not result.isError
                    report = json.loads(result.content[0].text)
                    for index in range(4):
                        result = await session.call_tool("logging_probe", {"index": index, "unavailable": index == 3})
                        assert not result.isError
                        assert json.loads(result.content[0].text) == {"success": True, "index": index}
                    return report

    async def run_both():
        return await asyncio.wait_for(asyncio.gather(run_server(0), run_server(1)), timeout=30)

    reports = asyncio.run(run_both())
    for key in ("usage_log_path", "diagnostic_log_path"):
        assert reports[0][key] != reports[1][key]
    for number, report in enumerate(reports):
        stderr = (tmp_path / f"server-{number}.stderr").read_text(encoding="utf-8")
        for key in ("usage_log_path", "diagnostic_log_path"):
            path = Path(report[key])
            assert path.parent == tmp_path
            assert str(path) in stderr
            entries = [r for file in path.parent.glob(path.name + "*") for r in records(file)]
            assert len({r["instance_id"] for r in entries}) == 1
            probes = [r for r in entries if r.get("tool", r["event"]) in ("logging_probe", "probe")]
            expected = [0, 1, 2] if key == "usage_log_path" else [0, 1, 2, 3]
            assert sorted(r.get("index", r.get("parameters", {}).get("index")) for r in probes) == expected
        assert "test file unavailable" in stderr
        fallback = [json.loads(line) for line in stderr.splitlines() if line.startswith("{")]
        assert any(r.get("tool") == "unavailable_probe" and r["parameters"]["index"] == 3 for r in fallback)
        assert any(r.get("tool") == "logging_probe" and r["parameters"]["index"] == 3 for r in fallback)
