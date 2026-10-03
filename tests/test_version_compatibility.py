"""X16 SemVer, discovery and public MCP/CLI admission; controlled fakes."""

import asyncio
import json
import subprocess
import threading
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import pytest

from msaccess_vcs_mcp import compatibility as c, tools, cli, version_probe
from msaccess_vcs_mcp.access_gate import reset_access_gate

pytestmark = pytest.mark.version_gate


@pytest.mark.parametrize("version,reason", [
    ("5.9.99", "below_minimum"), ("6.0.0", None), ("6.0.1", None),
    ("6.10.0", None), ("7.0.0", "unsupported_boundary"),
    ("7.0.0-rc.1", "unsupported_boundary"),
    ("6.0.0-dev.17", None), ("6.0.0-dev.17+local.009", None),
    ("6.0.0-rc.1", "prerelease_not_admitted"),
    ("6.1.0-beta.1", "prerelease_not_admitted"), ("6.0.0+build.001", None),
    ("6.00.0", "invalid_version"), ("v6.0.0", "invalid_version"),
    ("5.2", "invalid_version"), ("6.0.0.1", "invalid_version"),
    ("6.0.0-01", "invalid_version"), ("6.0.0+", "invalid_version"),
    (" 6.0.0", "invalid_version"), ("6.0.0\n", "invalid_version"),
    (None, "unknown_version"), ("", "unknown_version"), (5.2, "invalid_version"),
])
def test_addin_matrix(version, reason):
    result = c.compatibility_result(version)
    assert result["compatibility_reason"] == reason
    assert result["success"] == (reason is None)
    assert result["component"] == "addin"
    assert result["minimum_version"] == "6.0.0"
    if reason:
        assert result["error_pattern"] in {"version_unconfirmed", "version_incompatible"}
        assert "required" in result["error"]
        assert "recovery_action" in result


def test_semver_precedence_and_build_metadata():
    versions = ["1.0.0-alpha", "1.0.0-alpha.1", "1.0.0-alpha.beta", "1.0.0-beta",
                "1.0.0-beta.2", "1.0.0-beta.11", "1.0.0-rc.1", "1.0.0"]
    parsed = [c.Version.parse(v) for v in versions]
    assert all(left < right for left, right in zip(parsed, parsed[1:]))
    assert c.Version.parse("1.9.0") < c.Version.parse("1.10.0")
    assert c.Version.parse("1.0.0+x") == c.Version.parse("1.0.0+y")


@pytest.mark.parametrize("value,reason", [
    ("0.1.9", "below_minimum"), ("0.2.0", None), ("0.2.10", None),
    ("0.3.0", "unsupported_boundary"), ("1.0.0", "unsupported_boundary"),
    ("0.2.0-dev.16", None), ("0.2.0-dev.17", "prerelease_not_admitted"),
])
def test_major_zero(value, reason):
    requirement = c.Requirement("0.2.0", "0.3.0", ("0.2.0-dev.16",))
    assert requirement.reason(value) == reason


@pytest.mark.parametrize("low,high,pre", [
    ("0.2.0", "1.0.0", ()), ("1.0.0", "3.0.0", ()),
    ("1.0.0-alpha", "2.0.0", ()), ("1.0.0", "2.0.0", ("2.0.0-rc.1",)),
])
def test_invalid_requirements(low, high, pre):
    with pytest.raises(ValueError):
        c.Requirement(low, high, pre)


def test_newer_major_message_does_not_request_another_upgrade():
    result = c.compatibility_result("7.1.0")
    assert "Update the" not in result["recovery_action"]
    assert "consumer release" in result["recovery_action"]


@pytest.fixture
def discovery(monkeypatch, tmp_path):
    library = tmp_path / "installed.accda"
    library.write_bytes(b"first installation")
    monkeypatch.setattr(c, "configured_addin_path", lambda: str(library))
    return library


def test_replacement_and_configured_path_are_rechecked(discovery, monkeypatch):
    read = MagicMock(return_value="6.0.0")
    monkeypatch.setattr(c, "_read_installed_version", read)
    assert c.inspect_installed_addin()["success"]
    discovery.write_bytes(b"replacement")
    read.return_value = "5.9.0"
    assert not c.inspect_installed_addin()["success"]
    other = discovery.with_name("another.accde")
    other.write_bytes(b"different path")
    read.return_value = "7.0.0"
    monkeypatch.setattr(c, "configured_addin_path", lambda: str(other))
    assert c.inspect_installed_addin()["compatibility_reason"] == "unsupported_boundary"
    assert read.call_count == 3
    assert read.call_args.args == (str(other),)


def test_change_during_discovery_refused(discovery, monkeypatch):
    def changed(_):
        discovery.write_bytes(b"changed during read")
        return "6.0.0"
    monkeypatch.setattr(c, "_read_installed_version", changed)
    result = c.inspect_installed_addin()
    assert result["error_pattern"] == "version_unconfirmed"
    assert "changed during" in result["discovery_error"]


def test_missing_and_timed_out_library(discovery, monkeypatch):
    monkeypatch.setattr(c, "_read_installed_version", MagicMock(side_effect=subprocess.TimeoutExpired("probe", 10)))
    assert c.inspect_installed_addin()["error_pattern"] == "version_unconfirmed"
    discovery.unlink()
    assert c.inspect_installed_addin()["installed_version"] is None


def test_subprocess_uses_active_interpreter_bounded_hidden_read(monkeypatch):
    run = MagicMock(return_value=SimpleNamespace(stdout='{"version":"6.0.0"}'))
    monkeypatch.setattr(c.subprocess, "run", run)
    assert c._read_installed_version("install.accda") == "6.0.0"
    assert run.call_args.args[0] == [c.sys.executable, "-m", "msaccess_vcs_mcp.version_probe", "install.accda"]
    assert run.call_args.kwargs["timeout"] == 10
    assert run.call_args.kwargs["stdin"] == subprocess.DEVNULL


def test_dao_opens_only_installed_library_read_only(monkeypatch):
    import pythoncom
    import win32com.client
    app = MagicMock()
    engine = MagicMock()
    engine.OpenDatabase.return_value.Properties.return_value.Value = "6.0.0"
    monkeypatch.setattr(win32com.client, "Dispatch", lambda name: engine if name == "DAO.DBEngine.120" else app)
    assert version_probe.read_version("install.accda") == "6.0.0"
    engine.OpenDatabase.assert_called_once_with("install.accda", False, True)
    engine.OpenDatabase.return_value.Close.assert_called_once()
    assert not app.mock_calls


@pytest.fixture
def public(monkeypatch, discovery):
    monkeypatch.setattr(tools, "_ensure_env_loaded", AsyncMock())
    monkeypatch.setattr(tools, "load_config", lambda: {})
    monkeypatch.setattr(tools, "get_config", lambda: {})
    monkeypatch.setenv("ACCESS_VCS_ENABLE_LOGGING", "false")
    monkeypatch.setenv("ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG", "true")
    reset_access_gate()
    yield
    reset_access_gate()


# Use the registered wrappers, with valid argument shapes. The first guard is
# the real installed-library admission, not an injected failure in a handler.
CALLS = {
    "vcs_export_database": {"database_path": "target.accdb"},
    "vcs_import_objects": {"database_path": "target.accdb", "source_dir": "source"},
    "vcs_rebuild_database": {"source_dir": "source", "output_path": "output.accdb"},
    "vcs_rebuild_addin": {"source_dir": "source"},
    "vcs_check_vba_compiled": {"database_path": "target.accdb"},
    "vcs_compile_vba": {"database_path": "target.accdb"},
    "vcs_export_object": {"database_path": "target.accdb", "object_type": "module", "object_name": "Module1"},
    "vcs_import_object": {"database_path": "target.accdb", "object_type": "module", "object_name": "Module1"},
    "vcs_execute_sql": {"database_path": "target.accdb", "sql": "SELECT 1"},
    "vcs_call_vba": {"database_path": "target.accdb", "function_name": "Example"},
    "vcs_run_vba": {"database_path": "target.accdb", "code": "Debug.Print 1"},
    "vcs_set_option": {"database_path": "target.accdb", "option_name": "ShowDebug", "value": True},
    "vcs_get_option": {"database_path": "target.accdb", "option_name": "ShowDebug"},
    "vcs_get_log": {"database_path": "target.accdb"},
    "vcs_run_tests": {"database_path": "target.accdb"},
    "vcs_end_session": {"database_path": "target.accdb"},
}


@pytest.mark.parametrize("name", sorted(CALLS))
@pytest.mark.parametrize("installed", ["5.9.0", "7.0.0", "garbage", None])
def test_every_public_addin_path_refuses_before_any_dependent_work(public, monkeypatch, name, installed):
    monkeypatch.setattr(c, "_read_installed_version", lambda _: installed)
    connect = MagicMock(side_effect=AssertionError("target opened"))
    addin = MagicMock(side_effect=AssertionError("addin constructed or dispatched"))
    policy = MagicMock(side_effect=AssertionError("policy setup"))
    manager = MagicMock(side_effect=AssertionError("callback registration"))
    mutation = MagicMock(side_effect=AssertionError("target mutation"))
    monkeypatch.setattr(tools, "AccessConnection", connect)
    monkeypatch.setattr(tools, "VCSAddinIntegration", addin)
    monkeypatch.setattr(tools, "noninteractive_policy", policy)
    monkeypatch.setattr(tools, "_get_operation_manager", manager)
    monkeypatch.setattr(tools, "_open_build_host", mutation)
    result = asyncio.run(getattr(tools, name)(**CALLS[name]))
    assert not result["success"]
    assert result["component"] == "addin"
    assert result["installed_version"] == installed
    assert result["error_pattern"] in {"version_incompatible", "version_unconfirmed"}
    assert [mock.call_count for mock in (connect, addin, policy, manager, mutation)] == [0]*5


def test_admitted_public_tool_dispatches_normally(public, monkeypatch, tmp_path):
    monkeypatch.setattr(c, "_read_installed_version", lambda _: "6.0.0")
    db = tmp_path / "target.accdb"
    db.touch()
    conn, addin = MagicMock(), MagicMock()
    conn.return_value.__enter__.return_value.connect.return_value = (object(), object())
    addin.call_sync.return_value = "option value"
    monkeypatch.setattr(tools, "AccessConnection", conn)
    monkeypatch.setattr(tools, "VCSAddinIntegration", lambda _: addin)
    result = asyncio.run(tools.vcs_get_option(str(db), "ShowDebug"))
    assert result["success"]
    addin.load_addin.assert_called_once()
    addin.call_sync.assert_called_once_with("GetOption", "ShowDebug")


def test_cli_refusal_has_nonzero_exit_and_same_fields(public, monkeypatch, capsys):
    monkeypatch.setattr(c, "_read_installed_version", lambda _: "5.9.0")
    async def dispatch(name, arguments, progress):
        payload = await getattr(tools, name)(**arguments)
        return SimpleNamespace(isError=False, content=[SimpleNamespace(type="text", text=json.dumps(payload))])
    assert cli.main(["merge", "target.accdb", "source"], session_factory=dispatch) == 1
    output = capsys.readouterr().out
    assert '"installed_version": "5.9.0"' in output
    assert '"component": "addin"' in output
    assert "Update the MSAccess VCS add-in" in output


def test_metadata_survives_incompatible_or_missing_access(public, monkeypatch):
    monkeypatch.setattr(c, "_read_installed_version", lambda _: "5.9.0")
    connect = MagicMock(side_effect=AssertionError("Access started"))
    monkeypatch.setattr(tools, "AccessConnection", connect)
    result = asyncio.run(tools.vcs_get_version_info())
    assert result["mcp_version"] == "0.3.0-dev.17"
    assert result["vcs_version"] == "5.9.0"
    assert not result["addin_compatibility"]["success"]
    assert result["supported_addin_range"] == c.ADDIN_REQUIREMENT.text
    assert "access_version" in result and "bitness" in result
    assert connect.call_count == 0


def test_slow_admission_keeps_status_and_metadata_responsive(public, monkeypatch):
    entered, release = threading.Event(), threading.Event()
    gate_thread = []
    def slow(_):
        gate_thread.append(threading.get_ident())
        entered.set()
        assert release.wait(3)
        return "5.9.0"
    monkeypatch.setattr(c, "_read_installed_version", slow)
    async def scenario():
        loop_thread = threading.get_ident()
        task = asyncio.create_task(tools.vcs_run_tests("target.accdb"))
        try:
            assert await asyncio.to_thread(entered.wait, 1)
            result = await asyncio.wait_for(tools.vcs_cancel_operation("missing"), 0.5)
            assert result.get("error_pattern") != "tool_timeout"
            assert not task.done()
            assert gate_thread[0] != loop_thread
        finally:
            release.set()
            assert (await task)["error_pattern"] == "version_incompatible"
    asyncio.run(scenario())


def test_declared_tool_set_is_exhaustively_exercised():
    assert set(CALLS) == c.ADDIN_DEPENDENT_TOOLS


def test_current_distribution_delivers_workflow_owned_preflight():
    requirement = c.workflow_requirement()
    assert requirement.text == ">=0.3.0 <0.4.0; prerelease=0.3.0-dev.17"
    assert requirement.reason("0.1.0") == "below_minimum"
    assert requirement.reason("0.3.0-dev.17") is None
    assert requirement.reason("0.4.0") == "unsupported_boundary"
    instructions = tools.mcp._mcp_server.instructions
    assert requirement.text in instructions
    assert "Before database work" in instructions
    assert "after reconnecting" in instructions


def test_invalid_discovery_envelope_is_unconfirmed(discovery, monkeypatch):
    monkeypatch.setattr(c.subprocess, "run", lambda *a, **kw: SimpleNamespace(stdout="[]"))
    assert c.inspect_installed_addin()["error_pattern"] == "version_unconfirmed"


def test_shutdown_does_not_bypass_version_gate(discovery, monkeypatch):
    from msaccess_vcs_mcp import main
    import msaccess_vcs_mcp.access_com.connection as connection
    monkeypatch.setenv("ACCESS_VCS_SKIP_SESSION_CLEANUP", "false")
    monkeypatch.setenv("ACCESS_VCS_SESSION_ID", "x16-test")
    monkeypatch.setenv("ACCESS_VCS_DATABASE", "target.accdb")
    monkeypatch.setattr(main, "get_config", lambda: {"ACCESS_VCS_ADDIN_PATH": str(discovery)})
    monkeypatch.setattr(c, "_read_installed_version", lambda _: "5.9.0")
    monkeypatch.setattr(connection, "access_instance_is_live", lambda _: True)
    connect = MagicMock(side_effect=AssertionError("cleanup connected"))
    monkeypatch.setattr(connection, "AccessConnection", connect)
    main._cleanup_session()
    assert connect.call_count == 0
