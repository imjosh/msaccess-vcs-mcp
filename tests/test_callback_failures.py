"""M34: a failed terminal callback keeps its refusal pattern and decision fields.

Each payload is posted through a real ``OperationManager`` (as the add-in's
callback would) into the public tool: ``vcs_run_tests``, ``vcs_export_database``
and ``vcs_rebuild_database``. ``operation_already_running`` is covered here at
unit level only; X01 found a live second call hangs the MCP.
"""

import asyncio
import json
from contextlib import contextmanager
from unittest.mock import MagicMock, patch

import pytest

from msaccess_vcs_mcp.operation_manager import OperationManager
from tests.test_run_tests import _run_through_callback

DECISIONS = [{"object": "Form1", "type": "conflict", "resolution": "declined"}]
RUN_LOG = r"C:\src\logs\Run_1.log"


def _unwrap(tool_fn):
    fn = tool_fn
    while hasattr(fn, "__wrapped__"):
        fn = fn.__wrapped__
    return fn


@pytest.fixture(autouse=True)
def _diag_dir(monkeypatch, tmp_path_factory):
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(tmp_path_factory.mktemp("diag")))


def _error(**fields):
    return {"type": "error", "message": "Operation failed", **fields}


REFUSALS = {
    "operation_already_running": _error(
        message="Another VCS operation is already running.",
        error_pattern="operation_already_running",
    ),
    "invalid_decision_policy": _error(
        message="Unknown decision policy 'maybe'. Valid: block, prefer_source.",
        error_pattern="invalid_decision_policy",
    ),
}

# Payloads every tool must return whole. Each is (payload, expected fields).
PAYLOADS = {
    "refusal_operation_already_running": (
        REFUSALS["operation_already_running"],
        {
            "error": "Another VCS operation is already running.",
            "error_pattern": "operation_already_running",
        },
    ),
    "refusal_invalid_decision_policy": (
        REFUSALS["invalid_decision_policy"],
        {
            "error": "Unknown decision policy 'maybe'. Valid: block, prefer_source.",
            "error_pattern": "invalid_decision_policy",
        },
    ),
    "decision_required": (
        _error(
            message="Decision required",
            decision_required=True,
            error_pattern="decision_required",
            decisions=DECISIONS,
        ),
        {
            "decision_required": True,
            "error_pattern": "decision_required",
            "decisions": DECISIONS,
        },
    ),
    "decision_flag_alone": (
        _error(message="Decision required", decision_required=True, decisions=DECISIONS),
        {
            "decision_required": True,
            "error_pattern": "decision_required",
            "decisions": DECISIONS,
        },
    ),
    "decision_required_on_a_complete": (
        {
            "type": "complete",
            "message": "Operation completed",
            "decision_required": True,
            "decisions": DECISIONS,
        },
        {"decision_required": True, "error_pattern": "decision_required", "decisions": DECISIONS},
    ),
    "decisions_journal_on_a_failure": (
        _error(message="Merge failed after two conflicts", decisions=DECISIONS),
        {"error": "Merge failed after two conflicts", "decisions": DECISIONS},
    ),
    "cancelled": (
        {
            "type": "cancelled",
            "message": "Operation was cancelled",
            "decisions": DECISIONS,
        },
        {"cancelled": True, "error": "Operation was cancelled", "decisions": DECISIONS},
    ),
    "runtime_error": (
        _error(
            message="Object variable not set",
            runtime_error="Object variable not set",
            errorNumber=91,
        ),
        {
            "error": "Object variable not set",
            "runtime_error": "Object variable not set",
            "errorNumber": 91,
        },
    ),
}

# Today's results: payloads without the new fields.
PLAIN_FAILURE = _error(message="busy", log_path=RUN_LOG)
NEW_FIELDS = (
    "error_pattern", "decisions", "decision_required", "cancelled", "runtime_error", "errorNumber",
)


def _drive(manager, payload):
    """An ``APIAsync`` stand-in that posts ``payload`` as the add-in's terminal callback."""
    def _post(callback_info, command, *args):
        operation_id = json.loads(callback_info)["operation_id"]
        callback = json.loads(json.dumps({"operation_id": operation_id, **payload}))
        assert manager.route_callback(operation_id, callback)
        return {"async": True, "timeout_ms": 5000}

    return _post


def _conn():
    conn = MagicMock()
    conn.__enter__ = MagicMock(return_value=conn)
    conn.__exit__ = MagicMock(return_value=False)
    conn.connect.return_value = (MagicMock(), MagicMock())
    return conn


def _config(tmp_path):
    return {"ACCESS_VCS_ADDIN_PATH": str(tmp_path / "Version Control.accda")}


def _export_through_callback(tmp_path, payload):
    from msaccess_vcs_mcp.tools import vcs_export_database

    manager = OperationManager()
    addin = MagicMock()
    addin.call_async.side_effect = _drive(manager, payload)
    db = tmp_path / "test.accdb"
    db.touch()
    out = tmp_path / "export"
    out.mkdir()
    with (
        patch("msaccess_vcs_mcp.tools.AccessConnection", return_value=_conn()),
        patch("msaccess_vcs_mcp.tools.VCSAddinIntegration", return_value=addin),
        patch("msaccess_vcs_mcp.tools.validate_database_path", return_value=db),
        patch("msaccess_vcs_mcp.tools.validate_export_directory", return_value=out),
        patch("msaccess_vcs_mcp.tools._resolve_export_folder", return_value=(out, None)),
        patch("msaccess_vcs_mcp.tools._check_database_busy", return_value=None),
        patch("msaccess_vcs_mcp.tools.get_callback_url", return_value="http://localhost:1/cb"),
        patch("msaccess_vcs_mcp.tools._get_operation_manager", return_value=manager),
        patch("msaccess_vcs_mcp.tools.get_config", return_value=_config(tmp_path)),
    ):
        result = asyncio.run(_unwrap(vcs_export_database)(str(db), str(out)))
    addin.export_source.assert_not_called()
    assert manager.pending_count() == 0
    return result, out


def _build_through_callback(tmp_path, payload):
    from msaccess_vcs_mcp.tools import vcs_rebuild_database

    manager = OperationManager()
    addin = MagicMock()
    addin.build_as_paths_refusal.return_value = None
    addin.call_async.side_effect = _drive(manager, payload)
    src = tmp_path / "src"
    src.mkdir()
    output = str(tmp_path / "out.accdb")
    with (
        patch("msaccess_vcs_mcp.tools.check_write_permission", return_value=None),
        patch("msaccess_vcs_mcp.tools.validate_source_directory", return_value=src),
        patch("msaccess_vcs_mcp.tools.close_owned_instances_holding", return_value=[]),
        patch("msaccess_vcs_mcp.tools._check_database_busy", return_value=None),
        patch("msaccess_vcs_mcp.tools.ensure_dispatch", return_value=MagicMock()),
        patch("msaccess_vcs_mcp.tools.prefer_full_power_if_created"),
        patch("msaccess_vcs_mcp.tools.ensure_access_visible"),
        patch("msaccess_vcs_mcp.tools.VCSAddinIntegration", return_value=addin),
        patch("msaccess_vcs_mcp.tools.get_callback_url", return_value="http://localhost:1/cb"),
        patch("msaccess_vcs_mcp.tools._get_operation_manager", return_value=manager),
        patch("msaccess_vcs_mcp.tools.get_config", return_value=_config(tmp_path)),
    ):
        result = asyncio.run(_unwrap(vcs_rebuild_database)(str(src), output))
    addin.build_from_source.assert_not_called()
    assert manager.pending_count() == 0
    return result


def _assert_fields(result, expected):
    assert result["success"] is False
    for key, value in expected.items():
        assert result[key] == value, key


@pytest.mark.parametrize("name", PAYLOADS)
def test_run_tests_keeps_callback_failure_fields(tmp_path, name):
    payload, expected = PAYLOADS[name]
    _assert_fields(_run_through_callback(tmp_path, payload), expected)


@pytest.mark.parametrize("name", PAYLOADS)
def test_export_keeps_callback_failure_fields(tmp_path, name):
    payload, expected = PAYLOADS[name]
    result, out = _export_through_callback(tmp_path, payload)
    _assert_fields(result, expected)
    assert result["exported_count"] == 0
    assert result["export_path"] == str(out)
    assert result["objects_by_type"] == {}


@pytest.mark.parametrize("name", PAYLOADS)
def test_build_keeps_callback_failure_fields(tmp_path, name):
    payload, expected = PAYLOADS[name]
    result = _build_through_callback(tmp_path, payload)
    _assert_fields(result, expected)
    assert result["output_path"] is None


@pytest.mark.parametrize("path", ["tests", "export", "build"])
def test_decision_required_wins_over_an_other_pattern(tmp_path, path):
    payload = _error(
        message="Decision required",
        decision_required=True,
        error_pattern="operation_cancelled",
        decisions=DECISIONS,
    )
    result = {
        "tests": lambda: _run_through_callback(tmp_path, payload),
        "export": lambda: _export_through_callback(tmp_path, payload)[0],
        "build": lambda: _build_through_callback(tmp_path, payload),
    }[path]()
    assert result["error_pattern"] == "decision_required"
    assert result["decision_required"] is True


def test_run_tests_callback_refusal_matches_the_inline_refusal(tmp_path):
    """The same refusal reads alike whether it arrives by callback or inline."""
    from msaccess_vcs_mcp.tools import _parse_test_runner_json

    payload = REFUSALS["operation_already_running"]
    inline = _parse_test_runner_json({
        "success": False,
        "error_pattern": payload["error_pattern"],
        "error": payload["message"],
    })
    by_callback = _run_through_callback(tmp_path, payload)
    assert by_callback == inline


def test_run_tests_plain_failure_is_unchanged(tmp_path):
    assert _run_through_callback(tmp_path, PLAIN_FAILURE) == {
        "success": False,
        "error": "busy",
        "log_path": RUN_LOG,
    }


def test_export_plain_failure_is_unchanged(tmp_path):
    result, out = _export_through_callback(tmp_path, PLAIN_FAILURE)
    assert result == {
        "success": False,
        "error": "busy",
        "exported_count": 0,
        "export_path": str(out),
        "objects_by_type": {},
        "log_path": RUN_LOG,
    }


def test_build_plain_failure_is_unchanged(tmp_path):
    assert _build_through_callback(tmp_path, PLAIN_FAILURE) == {
        "success": False,
        "error": "busy",
        "output_path": None,
        "log_path": RUN_LOG,
    }


@pytest.mark.parametrize("path", ["export", "build"])
def test_plain_failure_gains_none_of_the_new_fields(tmp_path, path):
    result = (
        _export_through_callback(tmp_path, PLAIN_FAILURE)[0]
        if path == "export"
        else _build_through_callback(tmp_path, PLAIN_FAILURE)
    )
    assert not [key for key in NEW_FIELDS if key in result]
