"""A35's shared completion contract through vcs_run_tests and the real manager."""
import json

import pytest

from tests.test_run_tests import _call_run_tests, _run_through_callback


COMPILE_ERROR = "Project has compile errors. Please resolve before running tests."
DECISION_ERROR = "A required decision was not covered by the decision policy."
JOURNAL = [{"kind": "decision_required", "message": "Proceed?", "resolution": "decision_required"}]
ZERO = {"subs": 0, "assertions": 0, "passed": 0, "failed": 0, "errored": 0, "empty": 0}


@pytest.fixture(autouse=True)
def diagnostic_folder(monkeypatch, tmp_path):
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(tmp_path / "diagnostics"))


def deliver(tmp_path, transport, payload, saved):
    if transport == "sync":
        return _call_run_tests(tmp_path, json.dumps(payload))[0]
    callback = {"type": "error", "message": payload["error"], **payload}
    if saved is not None:
        path = tmp_path / "partial.json"
        path.write_text(json.dumps(saved), encoding="utf-8")
        callback["results_path"] = str(path)
        for key in ("summary", "tests", "allPassed"):
            callback.pop(key, None)
    return _run_through_callback(tmp_path, callback)


@pytest.mark.parametrize("transport", ["sync", "callback"])
@pytest.mark.parametrize("persistence", ["saved", "inline", "write_error"])
@pytest.mark.parametrize("construction_error", [False, True])
def test_combined_failure_retains_all_diagnostics(tmp_path, transport, persistence, construction_error):
    payload = {
        "success": False, "cancelled": False, "allPassed": False,
        "decision_required": True, "error_pattern": "decision_required", "error": DECISION_ERROR,
        "run_error": COMPILE_ERROR, "run_error_pattern": "project_not_compiled", "decisions": JOURNAL,
        "summary": ZERO, "tests": {}, "log_path": str(tmp_path / "own.log"),
    }
    if persistence == "write_error":
        payload["results_error"] = "Write to file failed"
    if construction_error:
        payload.update(completion_error="Completion payload fixture", completion_error_number=5)
    saved = {"success": False, "cancelled": False, "allPassed": False, "error": COMPILE_ERROR,
             "error_pattern": "project_not_compiled", "summary": ZERO, "tests": {}}
    result = deliver(tmp_path, transport, payload, saved if persistence == "saved" else None)
    for key, value in payload.items():
        assert result[key] == value, key
    assert result["logPath"] == payload["log_path"]
    assert "runtime_error" not in result and "errorNumber" not in result
    clean = _call_run_tests(tmp_path, json.dumps({"allPassed": True, "summary": {**ZERO, "subs": 1, "passed": 1}, "tests": {}}))[0]
    assert clean["success"] is True
    assert not set(clean).intersection({"run_error", "run_error_pattern", "completion_error", "decisions"})


@pytest.mark.parametrize("transport", ["sync", "callback"])
@pytest.mark.parametrize("saved", [False, True])
def test_compile_failure_alone_is_not_cancellation(tmp_path, transport, saved):
    payload = {"success": False, "cancelled": False, "allPassed": False, "error": COMPILE_ERROR,
               "error_pattern": "project_not_compiled", "summary": ZERO, "tests": {}}
    result = deliver(tmp_path, transport, payload, payload if saved else None)
    assert all(result[key] == value for key, value in payload.items())
    assert not set(result).intersection({"run_error", "run_error_pattern", "decision_required"})


@pytest.mark.parametrize("saved", [False, True])
def test_completion_construction_failure_overrides_passing_results(tmp_path, saved):
    partial = {"allPassed": True, "summary": {**ZERO, "subs": 1, "passed": 1}, "tests": {}}
    payload = {**partial, "success": False, "error": "Completion payload fixture",
               "completion_error": "Completion payload fixture", "completion_error_number": 5}
    result = deliver(tmp_path, "callback", payload, partial if saved else None)
    assert result["success"] is False
    assert result["error"] == "Completion payload fixture"
    assert result["completion_error_number"] == 5
    assert result["summary"]["passed"] == 1
