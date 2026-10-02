"""M51: public operations use only explicit or current, correctly named logs."""

import asyncio
import json
import os
from unittest.mock import AsyncMock, Mock, patch

import pytest

from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.access_gate import AccessGate
from msaccess_vcs_mcp.operation_manager import OperationManager

from .test_export_build_fallbacks import FakeAccess, _addin, _build_env, _export_env


OPERATIONS = ["Export", "FullExport", "Build", "BuildAs", "MergeBuild"]
ROUTES = ["no-callback", "unknown", "exception", "inline"]
CALL_START = 1000
JOURNAL = [{"kind": "conflict", "object": "Form1", "resolution": "blocked"}]


def _family(operation):
    return "Merge" if operation == "MergeBuild" else "Build" if operation.startswith("Build") else "Export"


def _write_log(path, body, mtime):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(body, encoding="utf-8")
    os.utime(path, (mtime, mtime))
    return path


@pytest.fixture
def public_operation(tmp_path, monkeypatch):
    """Decorated wrapper and real gate/integration/manager over fake COM calls."""
    gate = AccessGate()
    monkeypatch.setattr(tools, "get_access_gate", lambda: gate)
    monkeypatch.setattr(tools, "_ensure_env_loaded", AsyncMock())
    monkeypatch.setattr(tools, "load_config", lambda: {})
    monkeypatch.setattr(tools.time, "time", lambda: CALL_START)
    monkeypatch.setenv("ACCESS_VCS_ENABLE_LOGGING", "false")
    monkeypatch.setenv("ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG", "true")

    def run(operation, raw, route="no-callback", prepare=None):
        async_start = {
            "unknown": '{}', "exception": RuntimeError("async unavailable"),
            "inline": json.dumps({"sync": True, "result": raw}), "refused": raw,
        }.get(route)
        fake = FakeAccess(api=(raw, None), async_start=async_start)
        fake.export_folder = str(tmp_path / "export") + os.sep
        manager = OperationManager()
        manager.unregister_operation = Mock(wraps=manager.unregister_operation)
        manager.wait_for_completion = AsyncMock(wraps=manager.wait_for_completion)
        callback = route != "no-callback"
        addin = _addin(tmp_path, fake)

        def setup(src):
            for name in ("Export.log", "Build.log", "Merge.log"):
                _write_log(src / name, "Unrelated legacy operation", 1)
            if prepare:
                prepare(src)

        if operation.startswith("Build"):
            output = str(tmp_path / "out.accdb") if operation == "BuildAs" else None
            with _build_env(tmp_path, addin, manager, callback) as src:
                setup(src)
                result = asyncio.run(tools.vcs_rebuild_database(str(src), output))
        else:
            with _export_env(tmp_path, addin, manager, callback) as (db, src):
                setup(src)
                if operation == "MergeBuild":
                    with patch.object(tools, "validate_source_directory", return_value=src):
                        result = asyncio.run(tools.vcs_import_objects(str(db), str(src)))
                else:
                    result = asyncio.run(tools.vcs_export_database(
                        str(db), str(src), full_export=operation == "FullExport",
                    ))

        assert manager.pending_count() == 0
        manager.wait_for_completion.assert_not_awaited()
        if callback:
            manager.unregister_operation.assert_called_once()
        assert fake.dispatched == (
            ([("APIAsync", operation)] if callback else [])
            + ([] if route in ("inline", "refused") else [("API", operation)])
        )
        return result, src

    yield run
    gate._executor.shutdown(wait=True)


@pytest.mark.parametrize("operation", OPERATIONS)
@pytest.mark.parametrize("route", ["no-callback", "unknown", "exception", "refused"])
def test_refusals_never_borrow_legacy_or_current_logs(
    public_operation, addin_start_refusal, operation, route,
):
    def prepare(src):
        _write_log(src / "logs" / f"{_family(operation)}_20261002_120000_000.log", "Other run", 1001)

    raw = addin_start_refusal["raw"]
    if addin_start_refusal["expected"].get("decision_required"):
        raw = json.dumps({**json.loads(raw), "started": False})
    result, _ = public_operation(operation, raw, route, prepare)
    for key, value in addin_start_refusal["expected"].items():
        if key != "api_refused":  # Full import exposes refusal text/pattern.
            assert result[key] == value
    assert result["log_path"] is None
    assert "log_excerpt" not in result
    assert "completion_unconfirmed" not in result


@pytest.mark.parametrize("operation", OPERATIONS)
@pytest.mark.parametrize("route", ROUTES)
@pytest.mark.parametrize("log_key", ["log_path", "logPath"])
@pytest.mark.parametrize("outcome", ["refusal", "decision", "unconfirmed"])
def test_explicit_operation_logs_and_aliases_win(
    public_operation, tmp_path, operation, route, log_key, outcome,
):
    own_log = _write_log(tmp_path / "own.log", "Explicit operation diagnostic", 1)
    payload = {"success": outcome == "unconfirmed", log_key: str(own_log)}
    if outcome == "refusal":
        payload.update(error="Another API command is still running.",
                       error_pattern="operation_already_running", api_refused=True)
    elif outcome == "decision":
        payload.update(decision_required=True, decisions=JOURNAL, started=False)
    else:
        payload["started"] = True
        if operation == "MergeBuild" and route == "inline":
            # An inline merge normally supplies a final payload, unlike the
            # start-only export/build methods. Explicit uncertainty stays so.
            payload["completion_unconfirmed"] = True
    result, _ = public_operation(operation, json.dumps(payload), route)
    assert result["success"] is False
    assert result["log_path"] == str(own_log)
    assert result["log_excerpt"] == "Explicit operation diagnostic"
    if outcome == "refusal":
        assert result["error_pattern"] == "operation_already_running"
        assert "completion_unconfirmed" not in result
    elif outcome == "decision":
        assert result["error_pattern"] == "decision_required"
        assert result["decisions"] == JOURNAL
        assert "completion_unconfirmed" not in result
    else:
        assert result["started"] is True
        assert result["completion_unconfirmed"] is True
        assert "error_pattern" not in result


@pytest.mark.parametrize("operation", OPERATIONS)
@pytest.mark.parametrize("outcome", ["unconfirmed", "decision-ran", "decision-unstarted"])
@pytest.mark.parametrize("mtime,accepted", [(999, False), (1000, True), (1001, True)])
def test_fallback_requires_execution_family_and_timestamp(
    public_operation, operation, outcome, mtime, accepted,
):
    family = _family(operation)
    filename = f"{family}_20261002_120000_000.log"

    def prepare(src):
        # Fresh legacy and wrong-family files still cannot supply attribution.
        _write_log(src / f"{family}.log", "Fresh legacy log", 1001)
        _write_log(src / filename, "Wrong directory", 1001)
        _write_log(src / "logs" / "Other_20261002_235959_999.log", "Wrong family", 1001)
        _write_log(src / "logs" / f"{family}_20261001_120000_000.log", "Earlier run", 1)
        _write_log(src / "logs" / filename, "Current operation diagnostic", mtime)

    if outcome == "unconfirmed":
        raw = '{"success": true, "started": true}' if operation == "MergeBuild" else None
    else:
        raw = json.dumps({
            "success": False, "decision_required": True, "decisions": JOURNAL,
            "started": outcome == "decision-ran",
        })
    result, src = public_operation(operation, raw, prepare=prepare)
    attributable = accepted and outcome != "decision-unstarted"
    assert result["log_path"] == (str(src / "logs" / filename) if attributable else None)
    assert ("log_excerpt" in result) == attributable
    if attributable:
        assert result["log_excerpt"] == "Current operation diagnostic"
    assert result["success"] is False
    if outcome == "unconfirmed":
        assert result["started"] is True
        assert result["completion_unconfirmed"] is True
        assert "error_pattern" not in result
    else:
        assert result["error_pattern"] == "decision_required"
        assert result["decisions"] == JOURNAL
        assert "completion_unconfirmed" not in result


@pytest.mark.parametrize("log_key", ["log_path", "logPath"])
def test_export_vba_helper_preserves_explicit_log(tmp_path, log_key):
    _write_log(tmp_path / "Export.log", "Legacy log", 1)
    own_log = _write_log(tmp_path / "own.log", "Explicit diagnostic", 1)
    fake = FakeAccess(api=json.dumps({
        "success": False, "error": "Refused", log_key: str(own_log),
    }))
    result = _addin(tmp_path, fake).export_vba(str(tmp_path / "db.accdb"), str(tmp_path))
    assert result["log_path"] == str(own_log)
    assert fake.dispatched == [("API", "ExportVBA")]


@pytest.mark.parametrize("operation", OPERATIONS)
@pytest.mark.parametrize("log_key", ["log_path", "logPath"])
def test_unreadable_explicit_log_does_not_revert_to_disk_fallback(
    public_operation, tmp_path, operation, log_key,
):
    missing = tmp_path / "missing-operation.log"

    def prepare(src):
        _write_log(src / "logs" / f"{_family(operation)}_20261002_120000_000.log", "Other run", 1001)

    result, _ = public_operation(operation, json.dumps({
        "success": False, "error": "Refused", "started": False, log_key: str(missing),
    }), prepare=prepare)
    assert result["log_path"] == str(missing)
    assert "log_excerpt" not in result
    assert result["success"] is False
    assert "completion_unconfirmed" not in result
