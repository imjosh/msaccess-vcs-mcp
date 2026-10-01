"""M37: vcs_cancel_operation records a request; callers see only a confirmed cancel.

Every test runs a real OperationManager behind a real CallbackServer, so the
add-in's view (HTTP ``/cancel-status``, ``/callback``) and the original call's
result are checked together.
"""

import asyncio
import json
import urllib.request
from contextlib import contextmanager
from unittest.mock import AsyncMock, patch

import pytest

from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.callback_server import CallbackServer
from msaccess_vcs_mcp.operation_manager import OperationManager
from tests.test_run_tests import _build_mocks, _unwrap
from tests.test_scoped_sync import _patch_import_tool, _unwrap_sync

DECISIONS = [{"object": "Form1", "type": "conflict", "resolution": "source"}]


@pytest.fixture()
def seam():
    """A real OperationManager served by a real CallbackServer."""
    manager = OperationManager()
    server = CallbackServer(
        callback_router=manager.route_callback,
        cancel_checker=manager.is_cancelled,
        cancel_requester=manager.request_cancel,
    )
    server.start()
    base = f"http://{server.host}:{server.port}"
    try:
        yield manager, base
    finally:
        server.stop()


def _get(url):
    with urllib.request.urlopen(url, timeout=5) as response:
        return json.loads(response.read())


def _post_callback(base, operation_id, **payload):
    request = urllib.request.Request(
        f"{base}/callback",
        data=json.dumps({"operation_id": operation_id, **payload}).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=5) as response:
        return json.loads(response.read())


def _poll(base, operation_id):
    """What the add-in's poller sees for this operation."""
    return _get(f"{base}/cancel-status/{operation_id}")["cancelled"]


def _request_cancel(manager, operation_id):
    """Call the public tool, as an agent would."""
    with patch.object(tools, "_get_operation_manager", return_value=manager):
        return _unwrap_sync(tools.vcs_cancel_operation)(operation_id)


def test_tool_reports_a_request_not_a_stop(seam):
    manager, _base = seam
    operation_id, _queue = manager.register_operation()

    result = _request_cancel(manager, operation_id)

    assert result["success"] is True
    assert result["cancel_requested"] is True
    assert result["operation_id"] == operation_id
    assert "cancelled" not in result
    assert "error_pattern" not in result
    assert "has not stopped" in result["message"]
    assert "cancelled: true only if the add-in confirms" in result["message"]


def test_tool_fails_for_an_unknown_operation(seam):
    manager, _base = seam

    result = _request_cancel(manager, "no-such-operation")

    assert result["success"] is False
    assert "cancel_requested" not in result
    assert "error_pattern" not in result
    assert result["operation_id"] == "no-such-operation"


def test_tool_fails_without_an_operation_manager():
    with patch.object(tools, "_get_operation_manager", return_value=None):
        result = _unwrap_sync(tools.vcs_cancel_operation)("op")

    assert result["success"] is False
    assert "cancel_requested" not in result


def test_status_is_scoped_to_the_requested_operation(seam):
    manager, base = seam
    target, _q1 = manager.register_operation(database_path=r"C:\a.accdb")
    other, _q2 = manager.register_operation(database_path=r"C:\b.accdb")

    assert _poll(base, target) is False
    _request_cancel(manager, target)

    assert _poll(base, target) is True
    assert _poll(base, other) is False
    assert _poll(base, "never-registered") is False


def test_flag_is_cleared_when_the_operation_completes(seam):
    manager, base = seam
    path = r"C:\shared.accdb"
    first, _q = manager.register_operation(database_path=path)
    _request_cancel(manager, first)
    assert _poll(base, first) is True

    async def _finish():
        manager.set_event_loop(asyncio.get_running_loop())
        await asyncio.to_thread(_post_callback, base, first, type="complete", message="done")
        # The poller stops seeing the request as soon as the terminal callback lands.
        assert await asyncio.to_thread(_poll, base, first) is False
        return await manager.wait_for_completion(first, timeout_seconds=2)

    result = asyncio.run(_finish())

    assert result["success"] is True
    assert result["cancel_not_honored"] is True
    assert "cancelled" not in result
    assert "error_pattern" not in result or result["error_pattern"] is None
    assert _poll(base, first) is False

    # A later operation on the same database starts clean.
    second, _q = manager.register_operation(database_path=path)
    assert _poll(base, second) is False


def test_request_after_the_terminal_callback_is_refused(seam):
    manager, base = seam
    operation_id, _queue = manager.register_operation()

    async def _finish():
        manager.set_event_loop(asyncio.get_running_loop())
        await asyncio.to_thread(_post_callback, base, operation_id, type="complete")
        late = _request_cancel(manager, operation_id)
        return late, await manager.wait_for_completion(operation_id, timeout_seconds=2)

    late, result = asyncio.run(_finish())

    assert late["success"] is False
    assert late["error"] == "Operation not found or already completed"
    assert result["success"] is True
    assert "cancel_not_honored" not in result


@pytest.mark.parametrize("terminal", ["complete", "error"])
def test_unhonored_request_marks_the_original_result(seam, terminal):
    manager, base = seam
    operation_id, _queue = manager.register_operation()
    _request_cancel(manager, operation_id)

    async def _finish():
        manager.set_event_loop(asyncio.get_running_loop())
        await asyncio.to_thread(_post_callback, base, operation_id, type=terminal, message="ran on")
        return await manager.wait_for_completion(operation_id, timeout_seconds=2)

    result = asyncio.run(_finish())

    assert result["success"] is (terminal == "complete")
    assert result["cancel_not_honored"] is True
    assert "cancelled" not in result
    assert not result.get("error_pattern")


def test_confirmed_cancel_is_cancelled_without_the_marker(seam):
    manager, base = seam
    operation_id, _queue = manager.register_operation()
    _request_cancel(manager, operation_id)

    async def _finish():
        manager.set_event_loop(asyncio.get_running_loop())
        await asyncio.to_thread(
            _post_callback, base, operation_id,
            type="cancelled", message="Stopped at the next safe point", decisions=DECISIONS,
        )
        return await manager.wait_for_completion(operation_id, timeout_seconds=2)

    result = asyncio.run(_finish())

    assert result["success"] is False
    assert result["cancelled"] is True
    assert result["decisions"] == DECISIONS
    assert "cancel_not_honored" not in result
    assert not result.get("error_pattern")


def test_no_request_means_no_cancel_fields(seam):
    manager, base = seam
    operation_id, _queue = manager.register_operation()

    async def _finish():
        manager.set_event_loop(asyncio.get_running_loop())
        await asyncio.to_thread(_post_callback, base, operation_id, type="complete")
        return await manager.wait_for_completion(operation_id, timeout_seconds=2)

    result = asyncio.run(_finish())

    assert "cancel_not_honored" not in result
    assert "cancelled" not in result


def _run_tests_with(tmp_path, manager, base, *, request_cancel, payload):
    """Run vcs_run_tests; the add-in's COM call requests a cancel, then posts ``payload``."""
    mock_app, mock_conn, mock_addin = _build_mocks(tmp_path)
    seen = {}

    def _start(callback_info, command, *args):
        operation_id = json.loads(callback_info)["operation_id"]
        seen["before"] = _poll(base, operation_id)
        if request_cancel:
            seen["tool"] = _request_cancel(manager, operation_id)
            seen["during"] = _poll(base, operation_id)
        _post_callback(base, operation_id, **payload)
        seen["after"] = _poll(base, operation_id)
        return {"async": True, "timeout_ms": 5000}

    mock_addin.call_async.side_effect = _start
    db_path = str(tmp_path / "test.accdb")
    (tmp_path / "test.accdb").touch()

    with (
        patch("msaccess_vcs_mcp.tools.AccessConnection", return_value=mock_conn),
        patch("msaccess_vcs_mcp.tools.VCSAddinIntegration", return_value=mock_addin),
        patch("msaccess_vcs_mcp.tools.validate_database_path", return_value=tmp_path / "test.accdb"),
        patch("msaccess_vcs_mcp.tools.get_config", return_value={"ACCESS_VCS_ADDIN_PATH": str(tmp_path / "Version Control.accda")}),
        patch("msaccess_vcs_mcp.tools.get_callback_url", return_value=f"{base}/callback"),
        patch("msaccess_vcs_mcp.tools._get_operation_manager", return_value=manager),
        patch("msaccess_vcs_mcp.tools._check_database_busy", return_value=None),
    ):
        result = asyncio.run(_unwrap(tools.vcs_run_tests)(db_path))

    assert manager.pending_count() == 0
    return result, seen


def test_run_tests_completing_despite_a_request_reports_its_outcome(seam, tmp_path):
    manager, base = seam
    results = {
        "summary": {"subs": 2, "assertions": 2, "passed": 2, "failed": 0, "errored": 0, "empty": 0},
        "allPassed": True,
        "tests": {},
    }
    path = tmp_path / "TestResults_1.json"
    path.write_text(json.dumps(results), encoding="utf-8")

    result, seen = _run_tests_with(
        tmp_path, manager, base, request_cancel=True,
        payload={"type": "complete", "message": "done", "results_path": str(path)},
    )

    assert seen["before"] is False and seen["during"] is True and seen["after"] is False
    assert seen["tool"]["cancel_requested"] is True
    assert result["success"] is True
    assert result["allPassed"] is True
    assert result["cancel_not_honored"] is True
    assert "cancelled" not in result
    assert "error_pattern" not in result


def test_run_tests_confirmed_cancel_is_cancelled(seam, tmp_path):
    manager, base = seam

    result, seen = _run_tests_with(
        tmp_path, manager, base, request_cancel=True,
        payload={"type": "cancelled", "message": "Cancelled by request"},
    )

    assert seen["during"] is True and seen["after"] is False
    assert result["success"] is False
    assert result["cancelled"] is True
    assert "cancel_not_honored" not in result
    assert "error_pattern" not in result


def test_run_tests_without_a_request_has_no_cancel_fields(seam, tmp_path):
    manager, base = seam
    results = {"summary": {"passed": 1, "failed": 0, "errored": 0, "empty": 0}, "allPassed": True}
    path = tmp_path / "TestResults_1.json"
    path.write_text(json.dumps(results), encoding="utf-8")

    result, seen = _run_tests_with(
        tmp_path, manager, base, request_cancel=False,
        payload={"type": "complete", "results_path": str(path)},
    )

    assert seen["before"] is False and seen["after"] is False
    assert result["success"] is True
    assert "cancel_not_honored" not in result
    assert "cancelled" not in result


@pytest.fixture()
def _merge_env(monkeypatch, tmp_path):
    from msaccess_vcs_mcp.access_gate import reset_access_gate
    from msaccess_vcs_mcp.dialog_recovery import reset_interruptions

    reset_access_gate()
    reset_interruptions()
    monkeypatch.setattr(tools, "_ensure_env_loaded", AsyncMock())
    monkeypatch.setattr(tools, "load_config", lambda: {})
    monkeypatch.setenv("ACCESS_VCS_ENABLE_LOGGING", "false")
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(tmp_path / "diag"))


def _merge_with(monkeypatch, tmp_path, manager, base, payload):
    def _start(callback_info, command, *args):
        operation_id = json.loads(callback_info)["operation_id"]
        _request_cancel(manager, operation_id)
        assert _poll(base, operation_id) is True
        _post_callback(base, operation_id, **payload)
        assert _poll(base, operation_id) is False
        return {"async": True, "timeout_ms": 5000}

    # patch.object, not monkeypatch: these override _patch_import_tool's own
    # patches, and monkeypatch would restore those mocks after the test.
    with (
        _patch_import_tool(tmp_path) as (db, src, addin, _ops),
        patch.object(tools, "_get_operation_manager", return_value=manager),
        patch.object(tools, "get_callback_url", return_value=f"{base}/callback"),
    ):
        addin.call_async.side_effect = _start
        result = asyncio.run(tools.vcs_import_objects(db, str(src), decision_policy="prefer_source"))
    assert manager.pending_count() == 0
    return result


def test_merge_confirmed_cancel_keeps_decisions(seam, _merge_env, monkeypatch, tmp_path):
    manager, base = seam

    result = _merge_with(
        monkeypatch, tmp_path, manager, base,
        {"type": "cancelled", "message": "Merge cancelled", "decisions": DECISIONS},
    )

    assert result["success"] is False
    assert result["cancelled"] is True
    assert result["decisions"] == DECISIONS
    assert "cancel_not_honored" not in result
    assert "error_pattern" not in result


def test_merge_completing_despite_a_request_reports_the_marker(seam, _merge_env, monkeypatch, tmp_path):
    manager, base = seam

    result = _merge_with(
        monkeypatch, tmp_path, manager, base,
        {"type": "complete", "message": "Merged", "decisions": DECISIONS},
    )

    assert result["success"] is True
    assert result["decisions"] == DECISIONS
    assert result["cancel_not_honored"] is True
    assert "cancelled" not in result
    assert "error_pattern" not in result


@contextmanager
def _no_log_lookup():
    with patch.object(tools, "_newest_log", return_value=None):
        yield


def test_export_and_build_results_carry_the_cancel_outcome(tmp_path):
    """Export and build hand their completion to _attach_log_context."""
    with _no_log_lookup():
        cancelled = tools._attach_log_context(
            {"success": False, "error": "Operation cancelled"},
            tmp_path, "Export", {"success": False, "cancelled": True},
        )
        ignored = tools._attach_log_context(
            {"success": True}, tmp_path, "Build", {"success": True, "cancel_not_honored": True},
        )
        plain = tools._attach_log_context({"success": True}, tmp_path, "Build", {"success": True})

    assert cancelled["cancelled"] is True and "cancel_not_honored" not in cancelled
    assert ignored["cancel_not_honored"] is True and "cancelled" not in ignored
    assert "cancelled" not in plain and "cancel_not_honored" not in plain
