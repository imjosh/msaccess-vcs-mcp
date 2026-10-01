"""Tests for vcs_run_tests tool."""

import asyncio
import json
import os
from unittest.mock import Mock, MagicMock, patch, call

import pytest


SAMPLE_RESULTS_ALL_PASS = {
    "runAt": "2026-05-21 14:30:00",
    "databasePath": "C:\\db.accdb",
    "addinVersion": "4.1.0",
    "durationMs": 350,
    "summary": {
        "subs": 5,
        "assertions": 12,
        "passed": 12,
        "failed": 0,
        "errored": 0,
        "empty": 0,
    },
    "tests": {
        "modTestFoo.TestOne": {
            "status": "PASSED",
            "durationMs": 10,
            "tags": ["unit"],
            "assertions": [{"seq": 1, "passed": True, "context": "works"}],
        },
    },
}

SAMPLE_RESULTS_WITH_FAILURE = {
    "runAt": "2026-05-21 14:31:00",
    "databasePath": "C:\\db.accdb",
    "addinVersion": "4.1.0",
    "durationMs": 200,
    "summary": {
        "subs": 3,
        "assertions": 6,
        "passed": 4,
        "failed": 2,
        "errored": 0,
        "empty": 0,
    },
    "tests": {
        "modTestFoo.TestOne": {
            "status": "PASSED",
            "durationMs": 10,
            "tags": [],
            "assertions": [{"seq": 1, "passed": True}],
        },
        "modTestFoo.TestTwo": {
            "status": "FAILED",
            "durationMs": 15,
            "tags": [],
            "assertions": [{"seq": 1, "passed": False, "context": "expected 1"}],
        },
    },
}

SAMPLE_RESULTS_WITH_ERROR = {
    "runAt": "2026-05-21 14:32:00",
    "databasePath": "C:\\db.accdb",
    "addinVersion": "4.1.0",
    "durationMs": 100,
    "summary": {
        "subs": 2,
        "assertions": 1,
        "passed": 1,
        "failed": 0,
        "errored": 1,
        "empty": 0,
    },
    "tests": {
        "modTestFoo.TestOk": {
            "status": "PASSED",
            "durationMs": 5,
            "tags": [],
            "assertions": [{"seq": 1, "passed": True}],
        },
        "modTestFoo.TestBoom": {
            "status": "ERRORED",
            "durationMs": 3,
            "tags": [],
            "errorMessage": "Error 91: Object variable not set",
            "assertions": [],
        },
    },
}

SAMPLE_RESULTS_EMPTY = {
    "runAt": "2026-05-21 14:33:00",
    "databasePath": "C:\\db.accdb",
    "addinVersion": "4.1.0",
    "durationMs": 50,
    "summary": {
        "subs": 0,
        "assertions": 0,
        "passed": 0,
        "failed": 0,
        "errored": 0,
        "empty": 0,
    },
    "tests": {},
}


def _build_mocks(tmp_path, call_sync_return=None):
    """Build the mock objects needed for vcs_run_tests tests.

    Returns (mock_app, mock_addin, addin_path) where call_sync is
    pre-configured with the given return value.
    """
    addin_file = tmp_path / "Version Control.accda"
    addin_file.touch()

    mock_app = Mock()
    mock_db = Mock()
    mock_db.Name = str(tmp_path / "test.accdb")

    mock_conn = MagicMock()
    mock_conn.connect.return_value = (mock_app, mock_db)
    mock_conn.__enter__ = Mock(return_value=mock_conn)
    mock_conn.__exit__ = Mock(return_value=False)

    mock_addin = Mock()
    mock_addin.addin_path = str(addin_file)
    mock_addin.load_addin.return_value = True
    mock_addin.call_sync.return_value = call_sync_return
    mock_addin.call_sync.side_effect = lambda command, *args: (
        json.dumps({"success": True}) if command == "SetOption" else call_sync_return
    )

    return mock_app, mock_conn, mock_addin


@pytest.fixture(autouse=True)
def _patch_tool_infra(monkeypatch, tmp_path_factory):
    """Disable the @vcs_tool decorator's config/logging layers for unit tests."""
    diag_dir = tmp_path_factory.mktemp("diag")
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(diag_dir))


def _unwrap(tool_fn):
    fn = tool_fn
    while hasattr(fn, "__wrapped__"):
        fn = fn.__wrapped__
    return fn


def _call_run_tests(
    tmp_path, call_sync_return, filter_value=None, call_sync_side_effect=None, **tool_kwargs
):
    """Call vcs_run_tests with fully mocked COM layer (sync fallback).

    Patches AccessConnection, VCSAddinIntegration, validate_database_path,
    and get_config so the tool body runs against mock objects. Callbacks
    are disabled so the old add-in path (call_sync RunFilteredTests) is used.
    """
    mock_app, mock_conn, mock_addin = _build_mocks(
        tmp_path, call_sync_return=call_sync_return
    )
    if call_sync_side_effect is not None:
        mock_addin.call_sync.side_effect = call_sync_side_effect

    db_path = str(tmp_path / "test.accdb")
    (tmp_path / "test.accdb").touch()

    with (
        patch(
            "msaccess_vcs_mcp.tools.AccessConnection",
            return_value=mock_conn,
        ),
        patch(
            "msaccess_vcs_mcp.tools.VCSAddinIntegration",
            return_value=mock_addin,
        ),
        patch(
            "msaccess_vcs_mcp.tools.validate_database_path",
            return_value=tmp_path / "test.accdb",
        ),
        patch(
            "msaccess_vcs_mcp.tools.get_config",
            return_value={"ACCESS_VCS_ADDIN_PATH": str(tmp_path / "Version Control.accda")},
        ),
        patch("msaccess_vcs_mcp.tools.get_callback_url", return_value=None),
        patch("msaccess_vcs_mcp.tools._check_database_busy", return_value=None),
    ):
        from msaccess_vcs_mcp.tools import vcs_run_tests

        result = asyncio.run(
            _unwrap(vcs_run_tests)(db_path, filter=filter_value, **tool_kwargs)
        )

    return result, mock_app, mock_addin


class TestRunTestsSuccess:
    """Tests for successful test runs."""

    def test_all_tests_pass(self, tmp_path):
        result, mock_app, mock_addin = _call_run_tests(
            tmp_path,
            call_sync_return=json.dumps(SAMPLE_RESULTS_ALL_PASS),
        )

        assert result["success"] is True
        assert result["summary"]["subs"] == 5
        assert result["summary"]["failed"] == 0

    def test_no_filter_clears_default_test_filter(self, tmp_path):
        result, mock_app, mock_addin = _call_run_tests(
            tmp_path,
            call_sync_return=json.dumps(SAMPLE_RESULTS_ALL_PASS),
            filter_value=None,
        )

        mock_addin.call_sync.assert_any_call("SetOption", "DefaultTestFilter", "")
        mock_addin.call_sync.assert_any_call("RunFilteredTests", "block")

    def test_filter_sets_default_test_filter(self, tmp_path):
        result, mock_app, mock_addin = _call_run_tests(
            tmp_path,
            call_sync_return=json.dumps(SAMPLE_RESULTS_ALL_PASS),
            filter_value="modTestFoo,-slow",
        )

        mock_addin.call_sync.assert_any_call(
            "SetOption", "DefaultTestFilter", "modTestFoo,-slow"
        )
        mock_addin.call_sync.assert_any_call("RunFilteredTests", "block")

    def test_noninteractive_policy_is_scoped_argument(self, tmp_path):
        """The run passes a decision policy and does not stick the process in silent mode."""
        result, mock_app, mock_addin = _call_run_tests(
            tmp_path,
            call_sync_return=json.dumps(SAMPLE_RESULTS_ALL_PASS),
        )

        mock_app.Run.assert_not_called()
        mock_addin.call_sync.assert_any_call("RunFilteredTests", "block")
        assert result["success"] is True


class TestRunTestsFailure:
    """Tests for test runs that report failures or errors."""

    def test_failed_assertions(self, tmp_path):
        result, _, _ = _call_run_tests(
            tmp_path,
            call_sync_return=json.dumps(SAMPLE_RESULTS_WITH_FAILURE),
        )

        assert result["success"] is False
        assert result["summary"]["failed"] == 2

    def test_errored_tests(self, tmp_path):
        result, _, _ = _call_run_tests(
            tmp_path,
            call_sync_return=json.dumps(SAMPLE_RESULTS_WITH_ERROR),
        )

        assert result["success"] is False
        assert result["summary"]["errored"] == 1

    def test_no_tests_found(self, tmp_path):
        result, _, _ = _call_run_tests(
            tmp_path,
            call_sync_return=json.dumps(SAMPLE_RESULTS_EMPTY),
        )

        assert result["success"] is False
        assert result["summary"]["subs"] == 0


class TestRunTestsEdgeCases:
    """Tests for edge cases and error handling."""

    def test_empty_result(self, tmp_path):
        result, _, _ = _call_run_tests(
            tmp_path,
            call_sync_return="",
        )

        assert result["success"] is False
        assert "no results" in result["error"].lower()

    def test_none_result(self, tmp_path):
        result, _, _ = _call_run_tests(
            tmp_path,
            call_sync_return=None,
        )

        assert result["success"] is False
        assert "no results" in result["error"].lower()

    def test_invalid_json_result(self, tmp_path):
        result, _, _ = _call_run_tests(
            tmp_path,
            call_sync_return="not valid json {{{",
        )

        assert result["success"] is False
        assert "parse" in result["error"].lower()

    def test_com_exception(self, tmp_path):
        """COM error during call_sync is caught and returned."""
        mock_app, mock_conn, mock_addin = _build_mocks(tmp_path)
        mock_addin.call_sync.side_effect = Exception("COM disconnected")

        db_path = str(tmp_path / "test.accdb")
        (tmp_path / "test.accdb").touch()

        with (
            patch("msaccess_vcs_mcp.tools.AccessConnection", return_value=mock_conn),
            patch("msaccess_vcs_mcp.tools.VCSAddinIntegration", return_value=mock_addin),
            patch("msaccess_vcs_mcp.tools.validate_database_path", return_value=tmp_path / "test.accdb"),
            patch("msaccess_vcs_mcp.tools.get_config", return_value={"ACCESS_VCS_ADDIN_PATH": "test"}),
            patch("msaccess_vcs_mcp.tools.get_callback_url", return_value=None),
            patch("msaccess_vcs_mcp.tools._check_database_busy", return_value=None),
        ):
            from msaccess_vcs_mcp.tools import vcs_run_tests

            result = asyncio.run(_unwrap(vcs_run_tests)(db_path))

        assert result["success"] is False
        assert "COM disconnected" in result["error"]

    def test_dict_result_passthrough(self, tmp_path):
        """When add-in returns a dict directly (not JSON string), it's handled."""
        result, _, _ = _call_run_tests(
            tmp_path,
            call_sync_return=SAMPLE_RESULTS_ALL_PASS,
        )

        assert result["success"] is True
        assert result["summary"]["subs"] == 5


class TestRunTestsCallOrder:
    """Verify the COM call sequence."""

    def test_call_order(self, tmp_path):
        """SetOption -> RunFilteredTests(policy). Interaction mode is not set globally."""
        result, mock_app, mock_addin = _call_run_tests(
            tmp_path,
            call_sync_return=json.dumps(SAMPLE_RESULTS_ALL_PASS),
            filter_value="SQL,-slow",
        )

        mock_app.Run.assert_not_called()
        calls = mock_addin.call_sync.call_args_list
        assert len(calls) == 2
        assert calls[0] == call("SetOption", "DefaultTestFilter", "SQL,-slow")
        assert calls[1] == call("RunFilteredTests", "block")

        mock_addin.load_addin.assert_called_once()

    @pytest.mark.parametrize("callback_url", [None, "http://localhost:1/cb"])
    def test_interactive_request_is_refused_before_any_add_in_call(self, callback_url):
        """X11: an automation test run is always headless, so False is refused up front."""
        from msaccess_vcs_mcp.tools import vcs_run_tests

        with (
            patch("msaccess_vcs_mcp.tools.AccessConnection") as connection,
            patch("msaccess_vcs_mcp.tools.VCSAddinIntegration") as addin,
            patch("msaccess_vcs_mcp.tools.validate_database_path") as validate,
            patch("msaccess_vcs_mcp.tools._check_database_busy") as busy,
            patch("msaccess_vcs_mcp.tools.get_callback_url", return_value=callback_url),
        ):
            result = asyncio.run(vcs_run_tests(
                r"C:\db.accdb", filter="modTestFoo", noninteractive=False,
                decision_policy="prefer_source",
            ))

        assert result["success"] is False
        assert result["error_pattern"] == "interactive_tests_unsupported"
        assert "headless" in result["error"]
        for untouched in (connection, addin, validate, busy):
            untouched.assert_not_called()


def _call_run_tests_async(
    tmp_path, *, completion, filter_value=None, results=None, async_result=None
):
    """Call vcs_run_tests on the APIAsync callback path."""
    mock_app, mock_conn, mock_addin = _build_mocks(tmp_path)
    mock_addin.call_async.return_value = async_result or {"async": True, "timeout_ms": 1000}

    payload = dict(completion)
    if results is not None:
        results_path = tmp_path / "TestResults.json"
        results_path.write_text(json.dumps(results), encoding="utf-8")
        payload["results_path"] = str(results_path)

    op_manager = MagicMock()
    op_manager.register_operation.return_value = ("op-1", MagicMock())
    op_manager.create_callback_info.return_value = "{}"

    async def _wait(*_a, **_k):
        return payload

    op_manager.wait_for_completion = MagicMock(side_effect=_wait)

    db_path = str(tmp_path / "test.accdb")
    (tmp_path / "test.accdb").touch()

    with (
        patch("msaccess_vcs_mcp.tools.AccessConnection", return_value=mock_conn),
        patch("msaccess_vcs_mcp.tools.VCSAddinIntegration", return_value=mock_addin),
        patch("msaccess_vcs_mcp.tools.validate_database_path", return_value=tmp_path / "test.accdb"),
        patch(
            "msaccess_vcs_mcp.tools.get_config",
            return_value={"ACCESS_VCS_ADDIN_PATH": str(tmp_path / "Version Control.accda")},
        ),
        patch("msaccess_vcs_mcp.tools.get_callback_url", return_value="http://localhost:1/cb"),
        patch("msaccess_vcs_mcp.tools._get_operation_manager", return_value=op_manager),
        patch("msaccess_vcs_mcp.tools._check_database_busy", return_value=None),
    ):
        from msaccess_vcs_mcp.tools import vcs_run_tests

        result = asyncio.run(_unwrap(vcs_run_tests)(db_path, filter=filter_value))

    return result, mock_app, mock_addin, op_manager


class TestRunTestsAsync:
    """Tests for the APIAsync callback path."""

    def test_loads_results_path_on_complete(self, tmp_path):
        result, _, mock_addin, _ = _call_run_tests_async(
            tmp_path,
            completion={"success": True, "message": "Operation completed successfully"},
            results=SAMPLE_RESULTS_ALL_PASS,
        )

        assert result["success"] is True
        assert result["summary"]["subs"] == 5
        mock_addin.call_async.assert_called_once()
        assert mock_addin.call_async.call_args[0][1] == "RunFilteredTests"
        assert mock_addin.call_async.call_args[0][2] == "block"
        mock_addin.call_sync.assert_called_once_with("SetOption", "DefaultTestFilter", "")

    def test_failed_suite_still_returns_json(self, tmp_path):
        """eorFailed posts type error; results_path still carries the suite JSON."""
        result, _, _, _ = _call_run_tests_async(
            tmp_path,
            completion={"success": False, "error": "Operation failed"},
            results=SAMPLE_RESULTS_WITH_FAILURE,
        )

        assert result["success"] is False
        assert result["summary"]["failed"] == 2
        assert "tests" in result
        assert "error" not in result or result.get("summary")

    def test_sync_fallback_parses_wrapped_result(self, tmp_path):
        mock_app, mock_conn, mock_addin = _build_mocks(tmp_path)
        mock_addin.call_async.return_value = {
            "sync": True,
            "result": json.dumps(SAMPLE_RESULTS_ALL_PASS),
        }

        op_manager = MagicMock()
        op_manager.register_operation.return_value = ("op-1", MagicMock())
        op_manager.create_callback_info.return_value = "{}"

        db_path = str(tmp_path / "test.accdb")
        (tmp_path / "test.accdb").touch()

        with (
            patch("msaccess_vcs_mcp.tools.AccessConnection", return_value=mock_conn),
            patch("msaccess_vcs_mcp.tools.VCSAddinIntegration", return_value=mock_addin),
            patch("msaccess_vcs_mcp.tools.validate_database_path", return_value=tmp_path / "test.accdb"),
            patch(
                "msaccess_vcs_mcp.tools.get_config",
                return_value={"ACCESS_VCS_ADDIN_PATH": str(tmp_path / "Version Control.accda")},
            ),
            patch("msaccess_vcs_mcp.tools.get_callback_url", return_value="http://localhost:1/cb"),
            patch("msaccess_vcs_mcp.tools._get_operation_manager", return_value=op_manager),
            patch("msaccess_vcs_mcp.tools._check_database_busy", return_value=None),
        ):
            from msaccess_vcs_mcp.tools import vcs_run_tests

            result = asyncio.run(_unwrap(vcs_run_tests)(db_path))

        assert result["success"] is True
        assert result["summary"]["subs"] == 5
        op_manager.unregister_operation.assert_called_once_with("op-1")
        op_manager.wait_for_completion.assert_not_called()


RUNTIME_ERROR = "Error 91: Object variable not set"
DECISIONS = [{"prompt": "Install TestAssert?", "type": "confirm"}]


class TestRunTestsStartRefusal:
    """A refused start is returned once and the run is not retried."""

    @pytest.mark.parametrize(
        "refusal, expected_decisions",
        [
            ({"success": False, "error_pattern": "operation_already_running", "error": "busy"}, None),
            (
                {
                    "success": False,
                    "error_pattern": "decision_required",
                    "decision_required": True,
                    "decisions": json.dumps(DECISIONS),
                    "error": "busy",
                },
                DECISIONS,
            ),
        ],
    )
    def test_refused_start_is_returned_and_not_run_again(
        self, tmp_path, refusal, expected_decisions
    ):
        result, _, mock_addin, op_manager = _call_run_tests_async(
            tmp_path, completion={"success": True}, async_result=dict(refusal)
        )

        assert result["success"] is False
        assert result["error_pattern"] == refusal["error_pattern"]
        assert result["error"] == "busy"
        assert result.get("decisions") == expected_decisions
        mock_addin.call_async.assert_called_once()
        commands = [c.args[0] for c in mock_addin.call_sync.call_args_list]
        assert "RunFilteredTests" not in commands
        # Unregistered so the same refusal arriving on the callback is dropped.
        op_manager.unregister_operation.assert_called_once_with("op-1")
        op_manager.wait_for_completion.assert_not_called()


class TestRunTestsRuntimeError:
    """runtime_error survives every branch of a completion that carries it."""

    def test_runtime_error_survives_results_file(self, tmp_path):
        result, _, _, _ = _call_run_tests_async(
            tmp_path,
            completion={"success": False, "error": "Operation failed", "runtime_error": RUNTIME_ERROR},
            results=SAMPLE_RESULTS_ALL_PASS,
        )

        assert result["runtime_error"] == RUNTIME_ERROR
        assert result["success"] is False
        assert result["summary"]["subs"] == 5

    def test_decision_required_wins_over_runtime_error(self, tmp_path):
        result, _, _, _ = _call_run_tests_async(
            tmp_path,
            completion={
                "success": False,
                "error_pattern": "decision_required",
                "decision_required": True,
                "decisions": DECISIONS,
                "error": "A prompt was blocked",
                "runtime_error": RUNTIME_ERROR,
            },
            results=SAMPLE_RESULTS_ALL_PASS,
        )

        assert result["success"] is False
        assert result["decision_required"] is True
        assert result["error_pattern"] == "decision_required"
        assert result["error"] == "A prompt was blocked"
        assert result["decisions"] == DECISIONS
        assert result["runtime_error"] == RUNTIME_ERROR

    def test_runtime_error_survives_decision_required_without_results_file(self, tmp_path):
        result, _, _, _ = _call_run_tests_async(
            tmp_path,
            completion={
                "success": False,
                "error_pattern": "decision_required",
                "decision_required": True,
                "decisions": DECISIONS,
                "error": "A prompt was blocked",
                "runtime_error": RUNTIME_ERROR,
                "errorNumber": 91,
            },
        )

        assert result["success"] is False
        assert result["decision_required"] is True
        assert result["error_pattern"] == "decision_required"
        assert result["error"] == "A prompt was blocked"
        assert result["decisions"] == DECISIONS
        assert result["runtime_error"] == RUNTIME_ERROR

    def test_runtime_error_survives_plain_error_without_results_file(self, tmp_path):
        result, _, _, _ = _call_run_tests_async(
            tmp_path,
            completion={
                "success": False,
                "error": "Operation failed",
                "runtime_error": RUNTIME_ERROR,
                "errorNumber": 91,
                "log_path": r"C:\logs\Tests.log",
            },
        )

        assert result["success"] is False
        assert result["error"] == "Operation failed"
        assert result["runtime_error"] == RUNTIME_ERROR
        assert result["errorNumber"] == 91
        assert result["log_path"] == r"C:\logs\Tests.log"


def _run_through_callback(tmp_path, payload):
    """Post ``payload`` through a real OperationManager and return the tool result."""
    from msaccess_vcs_mcp.operation_manager import OperationManager

    mock_app, mock_conn, mock_addin = _build_mocks(tmp_path)
    manager = OperationManager()

    def _post_callback(callback_info, command, *args):
        # The add-in posts its terminal callback while the COM call runs.
        operation_id = json.loads(callback_info)["operation_id"]
        manager.route_callback(operation_id, {"operation_id": operation_id, **payload})
        return {"async": True, "timeout_ms": 5000}

    mock_addin.call_async.side_effect = _post_callback
    db_path = str(tmp_path / "test.accdb")
    (tmp_path / "test.accdb").touch()

    with (
        patch("msaccess_vcs_mcp.tools.AccessConnection", return_value=mock_conn),
        patch("msaccess_vcs_mcp.tools.VCSAddinIntegration", return_value=mock_addin),
        patch("msaccess_vcs_mcp.tools.validate_database_path", return_value=tmp_path / "test.accdb"),
        patch(
            "msaccess_vcs_mcp.tools.get_config",
            return_value={"ACCESS_VCS_ADDIN_PATH": str(tmp_path / "Version Control.accda")},
        ),
        patch("msaccess_vcs_mcp.tools.get_callback_url", return_value="http://localhost:1/cb"),
        patch("msaccess_vcs_mcp.tools._get_operation_manager", return_value=manager),
        patch("msaccess_vcs_mcp.tools._check_database_busy", return_value=None),
    ):
        from msaccess_vcs_mcp.tools import vcs_run_tests

        result = asyncio.run(_unwrap(vcs_run_tests)(db_path))

    mock_addin.call_sync.assert_called_once_with("SetOption", "DefaultTestFilter", "")
    assert manager.pending_count() == 0
    return result


class TestRunTestsRuntimeErrorThroughCallback:
    """A callback payload travels through the real OperationManager into the result."""

    def _run(self, tmp_path, payload):
        return _run_through_callback(tmp_path, payload)

    def test_plain_error_callback(self, tmp_path):
        result = self._run(tmp_path, {
            "type": "error",
            "message": RUNTIME_ERROR,
            "runtime_error": RUNTIME_ERROR,
            "errorNumber": 91,
            "log_path": r"C:\logs\TestRun_1.log",
        })

        assert result["success"] is False
        assert result["error"] == RUNTIME_ERROR
        assert result["runtime_error"] == RUNTIME_ERROR
        assert result["errorNumber"] == 91
        assert result["log_path"] == r"C:\logs\TestRun_1.log"

    def test_decision_required_callback(self, tmp_path):
        result = self._run(tmp_path, {
            "type": "error",
            "message": "Decision required",
            "decision_required": True,
            "error_pattern": "decision_required",
            "decisions": DECISIONS,
            "runtime_error": RUNTIME_ERROR,
            "errorNumber": 91,
        })

        assert result["success"] is False
        assert result["error_pattern"] == "decision_required"
        assert result["decisions"] == DECISIONS
        assert result["runtime_error"] == RUNTIME_ERROR
        assert result["errorNumber"] == 91


    def test_results_file_not_written(self, tmp_path):
        detail = r"Error writing file: C:\src\logs\TestResults_1.json (Write to file failed.)"
        result = self._run(tmp_path, {
            "type": "complete",
            "message": "Operation completed successfully",
            "results_error": detail,
            "log_path": r"C:\logs\TestRun_1.log",
        })

        assert result["success"] is False
        assert detail in result["error"]
        assert "modTestAssert" not in result["error"]
        assert result["results_error"] == detail
        assert result["log_path"] == r"C:\logs\TestRun_1.log"

    def test_results_error_does_not_hide_decision_required(self, tmp_path):
        result = self._run(tmp_path, {
            "type": "error",
            "message": "Decision required",
            "decision_required": True,
            "error_pattern": "decision_required",
            "decisions": DECISIONS,
            "results_error": "Error writing file",
        })

        assert result["error_pattern"] == "decision_required"
        assert result["decisions"] == DECISIONS

    def test_saved_results_win_over_a_stale_results_error(self, tmp_path):
        results = tmp_path / "TestResults_1.json"
        results.write_text(json.dumps({"summary": {"passed": 1, "failed": 0, "errored": 0, "subs": 1}}))
        result = self._run(tmp_path, {
            "type": "complete",
            "message": "Operation completed successfully",
            "results_path": str(results),
            "results_error": "left over",
        })

        assert result["success"] is True
        assert "results_error" not in result


PASSING_SUBSET = {"subs": 2, "assertions": 2, "passed": 2, "failed": 0, "errored": 0, "empty": 0}
ALL_EMPTY = {"subs": 2, "assertions": 0, "passed": 0, "failed": 0, "errored": 0, "empty": 2}
PASSING_AND_EMPTY = {"subs": 3, "assertions": 2, "passed": 2, "failed": 0, "errored": 0, "empty": 1}
SAVE_ERROR = r"Error writing file: C:\src\logs\TestResults_1.json (Write to file failed.)"
ADDIN_SAVE_ERROR_TEXT = "The test run finished, but its results file could not be written."
ACKNOWLEDGED = [{"prompt": "Overwrite TestResults?", "resolution": "acknowledged"}]
RUN_LOG = r"C:\src\logs\TestRun_1.log"


def _results(summary, **fields):
    """Add-in test results JSON: the sync return value and the TestResults file."""
    return {
        "durationMs": 20,
        **fields,
        "summary": dict(summary),
        "tests": {"modTestFoo.TestOne": {"status": "PASSED", "assertions": []}},
    }


# add-in sync JSON, callback payload ("file" is the results file it points to),
# expected on every transport, expected on the sync transports only.
VERDICT_CASES = [
    pytest.param(
        _results(PASSING_SUBSET, cancelled=True, allPassed=False),
        {"type": "cancelled", "message": "Operation cancelled",
         "file": _results(PASSING_SUBSET, cancelled=True, allPassed=False)},
        {"success": False, "cancelled": True, "summary": PASSING_SUBSET},
        {},
        id="cancelled-passing-subset",
    ),
    pytest.param(
        _results(PASSING_SUBSET, cancelled=False, allPassed=False, success=False,
                 results_error=SAVE_ERROR, error=ADDIN_SAVE_ERROR_TEXT),
        {"type": "error", "message": "Operation failed", "results_error": SAVE_ERROR},
        {"success": False, "results_error": SAVE_ERROR},
        {"summary": PASSING_SUBSET},
        id="results-not-written",
    ),
    pytest.param(
        _results(PASSING_SUBSET, allPassed=True, success=False, error="Refused by the add-in"),
        {"type": "complete", "message": "Operation completed successfully",
         "file": _results(PASSING_SUBSET, allPassed=True, success=False, error="Refused by the add-in")},
        {"success": False, "error": "Refused by the add-in"},
        {},
        id="explicit-success-false",
    ),
    pytest.param(
        _results(ALL_EMPTY, cancelled=False, allPassed=False),
        {"type": "error", "message": "Operation failed",
         "file": _results(ALL_EMPTY, cancelled=False, allPassed=False)},
        {"success": False, "summary": ALL_EMPTY},
        {},
        id="all-empty",
    ),
    pytest.param(
        _results(ALL_EMPTY),
        {"type": "error", "message": "Operation failed", "file": _results(ALL_EMPTY)},
        {"success": False, "summary": ALL_EMPTY},
        {},
        id="all-empty-older-addin",
    ),
    pytest.param(
        _results(PASSING_AND_EMPTY, cancelled=False, allPassed=True),
        {"type": "complete", "message": "Operation completed successfully",
         "file": _results(PASSING_AND_EMPTY, cancelled=False, allPassed=True)},
        {"success": True},
        {},
        id="passing-and-empty",
    ),
    pytest.param(
        _results(PASSING_SUBSET),
        {"type": "complete", "message": "Operation completed successfully",
         "file": _results(PASSING_SUBSET)},
        {"success": True},
        {},
        id="passing-older-addin",
    ),
    pytest.param(
        _results(PASSING_SUBSET, allPassed=True, decisions=json.dumps(ACKNOWLEDGED)),
        {"type": "complete", "message": "Operation completed successfully",
         "decisions": ACKNOWLEDGED, "file": _results(PASSING_SUBSET, allPassed=True)},
        {"success": True, "decisions": ACKNOWLEDGED},
        {},
        id="acknowledged-decisions-kept",
    ),
    pytest.param(
        _results(PASSING_SUBSET, allPassed=True, logPath=RUN_LOG),
        {"type": "complete", "message": "Operation completed successfully", "log_path": RUN_LOG,
         "file": _results(PASSING_SUBSET, allPassed=True, logPath=RUN_LOG)},
        {"success": True, "log_path": RUN_LOG, "logPath": RUN_LOG},
        {},
        id="log-path-both-spellings",
    ),
    pytest.param(
        _results(PASSING_SUBSET, cancelled=True, allPassed=False, success=False,
                 decision_required=True, error_pattern="decision_required",
                 decisions=json.dumps(DECISIONS), error="A prompt was blocked"),
        {"type": "error", "message": "A prompt was blocked", "decision_required": True,
         "error_pattern": "decision_required", "decisions": DECISIONS,
         "file": _results(PASSING_SUBSET, cancelled=True, allPassed=False)},
        {"success": False, "decision_required": True, "error_pattern": "decision_required",
         "decisions": DECISIONS, "error": "A prompt was blocked"},
        {},
        id="decision-required-wins-over-cancel",
    ),
]


def _run_sync_fallback(tmp_path, addin_json, _callback):
    result, _, _ = _call_run_tests(tmp_path, call_sync_return=json.dumps(addin_json))
    return result


def _run_inline_sync(tmp_path, addin_json, _callback):
    result, _, _, op_manager = _call_run_tests_async(
        tmp_path, completion={}, async_result={"sync": True, "result": json.dumps(addin_json)}
    )
    op_manager.wait_for_completion.assert_not_called()
    return result


def _run_callback(tmp_path, _addin_json, callback):
    payload = dict(callback)
    results = payload.pop("file", None)
    if results is not None:
        results_path = tmp_path / "TestResults_1.json"
        results_path.write_text(json.dumps(results), encoding="utf-8")
        payload["results_path"] = str(results_path)
    return _run_through_callback(tmp_path, payload)


SYNC_TRANSPORTS = [
    pytest.param(_run_sync_fallback, id="sync-fallback"),
    pytest.param(_run_inline_sync, id="inline-sync"),
]
ALL_TRANSPORTS = [*SYNC_TRANSPORTS, pytest.param(_run_callback, id="callback")]


class TestRunTestsOneVerdict:
    """The same add-in outcome gives the same verdict on every transport."""

    @pytest.mark.parametrize("run", ALL_TRANSPORTS)
    @pytest.mark.parametrize("addin_json, callback, expected, sync_expected", VERDICT_CASES)
    def test_verdict(self, tmp_path, run, addin_json, callback, expected, sync_expected):
        result = run(tmp_path, addin_json, callback)

        for key, value in expected.items():
            assert result.get(key) == value, key
        if run is not _run_callback:
            for key, value in sync_expected.items():
                assert result.get(key) == value, key
        if not result["success"]:
            assert result.get("error")
        if "decision_required" not in expected:
            assert "decision_required" not in result
            assert result.get("error_pattern") is None

    @pytest.mark.parametrize("run", ALL_TRANSPORTS)
    def test_cancelled_run_keeps_its_tests(self, tmp_path, run):
        cancelled = _results(PASSING_SUBSET, cancelled=True, allPassed=False)
        result = run(tmp_path, cancelled, {"type": "cancelled", "file": cancelled})

        assert result["success"] is False
        assert "modTestFoo.TestOne" in result["tests"]

    @pytest.mark.parametrize("run", ALL_TRANSPORTS)
    def test_all_empty_error_names_the_empty_count(self, tmp_path, run):
        all_empty = _results(ALL_EMPTY, cancelled=False, allPassed=False)
        result = run(tmp_path, all_empty, {"type": "error", "file": all_empty})

        assert "EMPTY" in result["error"]
        assert "2 " in result["error"]
        assert "error_pattern" not in result

    @pytest.mark.parametrize("run", SYNC_TRANSPORTS)
    def test_unsaved_results_use_the_x07_message(self, tmp_path, run):
        unsaved = _results(PASSING_SUBSET, cancelled=False, allPassed=False, success=False,
                           results_error=SAVE_ERROR, error=ADDIN_SAVE_ERROR_TEXT)
        result = run(tmp_path, unsaved, {})

        assert SAVE_ERROR in result["error"]
        assert "modTestAssert" not in result["error"]
