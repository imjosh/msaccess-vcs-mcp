"""The rebuild launcher and child build own distinct registered identities."""
import asyncio
import json
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock, Mock

import pytest
from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.operation_manager import OperationManager


def unwrap(fn):
    while hasattr(fn, "__wrapped__"):
        fn = fn.__wrapped__
    return fn


@pytest.fixture
def rebuild(monkeypatch, tmp_path):
    source = tmp_path / "Version Control.accda.src"
    source.mkdir()
    (tmp_path / "Version Control.accda").write_bytes(b"closed mock host")
    order, registrations = [], []
    class Manager(OperationManager):
        def register_operation(self, **kwargs):
            identity, queue = super().register_operation(**kwargs)
            registrations.append((identity, queue, kwargs))
            order.append("register:" + kwargs["command"])
            return identity, queue
    manager = Manager()
    async def exclusive(tool, database, fn, asynchronous, *args, **kwargs):
        return await fn(*args, **kwargs) if asynchronous else fn(*args, **kwargs)
    monkeypatch.setattr(tools, "get_access_gate", lambda: Mock(run_exclusive=exclusive))
    monkeypatch.setattr(tools, "get_config", lambda: {})
    monkeypatch.setattr(tools, "check_write_permission", lambda _: None)
    monkeypatch.setattr(tools, "_is_installed_addin_path", lambda _: False)
    monkeypatch.setattr(tools, "close_owned_instances_holding", lambda *args: [])
    monkeypatch.setattr(tools, "list_access_pids", lambda: set())
    monkeypatch.setattr(tools, "get_callback_url", lambda: "http://127.0.0.1:54321/callback")
    monkeypatch.setattr(tools, "_get_operation_manager", lambda: manager)
    return source, manager, order, registrations


@pytest.mark.parametrize("status,success", [("complete", True), ("compile-failed", False)])
def test_child_progress_keeps_outer_final_identity_and_status_authority(rebuild, monkeypatch, status, success):
    source, manager, order, registrations = rebuild
    captured = {}
    def launch(database, function, arguments, timeout=None, *, on_admitted=None, operation_id=None):
        if on_admitted is not None:
            arguments = on_admitted(arguments)
        identity = operation_id() if callable(operation_id) else operation_id
        assert order == ["register:RebuildAddIn", "register:Build"]
        assert manager.pending_count() == 2
        outer, child = registrations
        assert identity == outer[0] and identity != child[0]
        assert outer[2]["database_path"] == str(source.parent / "Version Control.accda")
        assert "database_path" not in child[2]
        callback = json.loads(arguments[2])
        assert arguments[:2] == ["RebuildAddIn", str(source)]
        assert callback["operation_id"] == child[0]
        captured.update(outer=identity, child=child[0])
        order.append("launch")
        return {"success": True, "result": json.dumps({"success": True, "status": "launched", "phaseStarted": "fresh", "operation_id": identity, "build_operation_id": child[0]})}
    monkeypatch.setattr(tools, "_execute_call_vba", launch)
    async def watch(*args, **kwargs):
        assert order[-1] == "launch"
        child_queue = registrations[1][1]
        await child_queue.put({"type": "log", "operation_id": captured["child"], "message": "Importing source"})
        await child_queue.put({"type": "complete", "operation_id": captured["child"], "message": "child done", "log_path": "child-build.log"})
        await asyncio.sleep(0)
        return {"success": success, "status": status, "phaseStarted": "fresh", "watcher_field": "preserved"}
    monkeypatch.setattr(tools, "wait_for_rebuild_status", watch)
    ctx = Mock(report_progress=AsyncMock())
    result = asyncio.run(unwrap(tools.vcs_rebuild_addin)(str(source), ctx=ctx))
    assert result["operation_id"] == captured["outer"]
    assert result["build_operation_id"] == captured["child"]
    assert result["success"] is success and result["status"] == status
    assert result["watcher_field"] == "preserved"
    assert result["log_messages"] == ["Importing source"]
    assert result["log_path"] == "child-build.log"
    assert any("child done" in call.kwargs["message"] for call in ctx.report_progress.await_args_list)
    assert manager.pending_count() == 0


def test_launch_refusal_has_no_watch_or_fallback_and_releases_both_ids(rebuild, monkeypatch):
    source, manager, order, registrations = rebuild
    calls = []
    def launch(database, function, arguments, timeout=None, *, on_admitted=None, operation_id=None):
        if on_admitted:
            arguments = on_admitted(arguments)
        identity = operation_id() if callable(operation_id) else operation_id
        calls.append(identity)
        return {"success": True, "result": json.dumps({"success": False, "status": "refused", "error_pattern": "timer_unavailable", "error": "queue full", "operation_id": identity})}
    monkeypatch.setattr(tools, "_execute_call_vba", launch)
    watcher = AsyncMock()
    monkeypatch.setattr(tools, "wait_for_rebuild_status", watcher)
    result = asyncio.run(unwrap(tools.vcs_rebuild_addin)(str(source)))
    assert len(calls) == 1
    assert result["error_pattern"] == "timer_unavailable"
    assert result["operation_id"] == registrations[0][0]
    assert result["build_operation_id"] == registrations[1][0]
    watcher.assert_not_awaited()
    assert manager.pending_count() == 0


def test_no_listener_keeps_outer_identity_without_claiming_child_stream(rebuild, monkeypatch):
    source, manager, order, registrations = rebuild
    monkeypatch.setattr(tools, "get_callback_url", lambda: None)
    def launch(database, function, arguments, timeout=None, *, on_admitted=None, operation_id=None):
        if on_admitted:
            arguments = on_admitted(arguments)
        identity = operation_id() if callable(operation_id) else operation_id
        assert len(registrations) == 1
        assert arguments == ["RebuildAddIn", str(source)]
        return {"success": True, "result": json.dumps({"status": "refused", "success": False, "error": "source refused", "operation_id": identity})}
    monkeypatch.setattr(tools, "_execute_call_vba", launch)
    result = asyncio.run(unwrap(tools.vcs_rebuild_addin)(str(source)))
    assert result["operation_id"] == registrations[0][0]
    assert "build_operation_id" not in result
    assert manager.pending_count() == 0


def test_raw_launch_uses_supplied_outer_id_after_admission(tmp_path, monkeypatch):
    from tests.test_call_vba_guards import _fake_access
    library = tmp_path / "Version Control.accda"
    library.touch()
    monkeypatch.setattr(tools, "get_config", lambda: {"ACCESS_VCS_ADDIN_PATH": str(library)})
    monkeypatch.setattr(tools, "validate_database_path", Path)
    monkeypatch.setattr(tools, "get_call_vba_timeout", lambda _: 5)
    app = _fake_access(monkeypatch, run_result='{"success":true,"status":"launched"}')
    admitted = []
    production = hasattr(tools, "ensure_session") and hasattr(tools, "request_policy_contract_confirmed")
    kwargs = {"operation_id": "registered-outer-id"}
    if production:
        monkeypatch.setattr(tools, "ensure_session", lambda *args: "accepted-envelope")
        monkeypatch.setattr(tools, "request_policy_refusal", lambda *args: None)
        monkeypatch.setattr(tools, "request_policy_contract_confirmed", lambda _: True)
        def after_admission(arguments):
            admitted.append(True)
            return arguments
        kwargs.update(on_admitted=after_admission, operation_id=lambda: "registered-outer-id")
    result = tools._execute_call_vba(str(tmp_path / "host.accda"), "VCS.API", ["RebuildAddIn", str(tmp_path / "source"), '{"operation_id":"registered-child-id"}'], **kwargs)
    assert result["success"] is True
    wire = app.Run.call_args.args
    if production:
        assert admitted == [True]
        assert wire[0].endswith(".APIExecuteMCP")
        assert wire[1:4] == ("accepted-envelope", "registered-outer-id", "RebuildAddIn")
    else:
        assert wire[0].endswith(".APIMCP")
        assert wire[1:3] == ("registered-outer-id", "RebuildAddIn")
    assert json.loads(wire[-1])["operation_id"] == "registered-child-id"


@pytest.mark.parametrize("callback_url", [None, "http://127.0.0.1:54321/callback"])
def test_busy_gate_preserves_prior_database_operation(rebuild, monkeypatch, callback_url):
    import threading
    from msaccess_vcs_mcp import access_gate
    source, manager, order, registrations = rebuild
    host = str(source.parent / "Version Control.accda")
    prior_id, _ = manager.register_operation(database_path=host, command="Build")
    prior = manager.get_active_operation_for_database(host)
    registrations.clear()
    order.clear()
    monkeypatch.setattr(tools, "get_callback_url", lambda: callback_url)
    monkeypatch.setattr(access_gate, "_read_busy_wait_sec", lambda: 0.01)
    monkeypatch.setattr(access_gate, "_init_com_apartment", lambda: None)
    gate = access_gate.AccessGate()
    monkeypatch.setattr(tools, "get_access_gate", lambda: gate)
    execute = Mock(side_effect=AssertionError("busy rebuild must not dispatch"))
    close = Mock(side_effect=AssertionError("busy rebuild must not close Access"))
    monkeypatch.setattr(tools, "_execute_call_vba", execute)
    monkeypatch.setattr(tools, "close_owned_instances_holding", close)
    started, release = threading.Event(), threading.Event()
    def prior_work():
        started.set()
        assert release.wait(5), "test failed to release prior operation"
    async def exercise():
        running = asyncio.create_task(gate.run_exclusive("vcs_build_database", host, prior_work, False))
        try:
            assert await asyncio.to_thread(started.wait, 2)
            in_flight = gate.current_in_flight()
            assert in_flight.tool == "vcs_build_database"
            result = await unwrap(tools.vcs_rebuild_addin)(str(source))
            assert result["error_pattern"] == "server_busy"
            assert "operation_id" not in result and "build_operation_id" not in result
            assert registrations == []
            assert manager.get_active_operation_for_database(host) is prior
            assert manager.get_operation(prior_id) is prior
            assert manager.pending_count() == 1
            assert gate.current_in_flight() is in_flight
            execute.assert_not_called()
            close.assert_not_called()
        finally:
            release.set()
            await running
    try:
        asyncio.run(exercise())
    finally:
        gate._executor.shutdown(wait=True)
        manager.unregister_operation(prior_id)


def test_admitted_gate_registers_both_ids_before_close_and_actual_run(rebuild, monkeypatch):
    from msaccess_vcs_mcp import access_gate
    from tests.test_call_vba_guards import _fake_access
    source, manager, order, registrations = rebuild
    monkeypatch.setattr(access_gate, "_init_com_apartment", lambda: None)
    gate = access_gate.AccessGate()
    monkeypatch.setattr(tools, "get_access_gate", lambda: gate)
    monkeypatch.setattr(tools, "get_config", lambda: {"ACCESS_VCS_ADDIN_PATH": str(source.parent / "Version Control.accda")})
    monkeypatch.setattr(tools, "validate_database_path", Path)
    monkeypatch.setattr(tools, "get_call_vba_timeout", lambda _: 5)
    def close(*args):
        assert gate.current_in_flight().tool == "vcs_rebuild_addin"
        assert order == ["register:RebuildAddIn", "register:Build"]
        assert manager.pending_count() == 2
        order.append("close")
        return []
    monkeypatch.setattr(tools, "close_owned_instances_holding", close)
    def run(endpoint, outer_id, method, source_path, callback_info):
        assert gate.current_in_flight().tool == "vcs_rebuild_addin"
        assert order == ["register:RebuildAddIn", "register:Build", "close"]
        assert endpoint.endswith(".APIMCP") and method == "RebuildAddIn"
        assert outer_id == registrations[0][0]
        child_id = json.loads(callback_info)["operation_id"]
        assert child_id == registrations[1][0] and child_id != outer_id
        assert source_path == str(source)
        order.append("run")
        return json.dumps({"success": False, "status": "refused", "error": "controlled refusal"})
    app = _fake_access(monkeypatch, run_side_effect=run)
    try:
        result = asyncio.run(unwrap(tools.vcs_rebuild_addin)(str(source)))
        assert result["success"] is False
        assert result["operation_id"] == registrations[0][0]
        assert result["build_operation_id"] == registrations[1][0]
        app.Run.assert_called_once()
        assert order[-1] == "run"
        assert manager.pending_count() == 0
        assert gate.current_in_flight() is None
    finally:
        gate._executor.shutdown(wait=True)
