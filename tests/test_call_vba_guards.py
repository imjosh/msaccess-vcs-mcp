"""Tests for vcs_call_vba's dispatcher and rebuild-status correlation.

Deadline behaviour runs through the real gate in ``test_call_vba_deadline.py``.

Refusing the installed add-in as a target is the ``vcs_tool`` wrapper's job now and
lives in ``test_installed_addin_guard.py``.

Paths here are synthetic. They only have to look like a real install and a real
repository to the code under test; nothing is read from disk.
"""

from __future__ import annotations

import json
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

import msaccess_vcs_mcp.tools as tools_module


INSTALLED_ADDIN = r"C:\Users\Example\AppData\Roaming\MSAccessVCS\Version Control.accda"
REPO_ADDIN = r"C:\Projects\msaccess-vcs-addin\Version Control.accda"
SOURCE_DIR = r"C:\Projects\msaccess-vcs-addin\Version Control.accda.src"
RESOLVED_API = r"C:\Users\Example\AppData\Roaming\MSAccessVCS\Version Control.API"


def _call_vba(*args, **kwargs):
    fn = tools_module.vcs_call_vba
    while hasattr(fn, "__wrapped__"):
        fn = fn.__wrapped__
    return fn(*args, **kwargs)


def _fake_access(monkeypatch, run_result=None, run_side_effect=None):
    """Wire the connection to one fake Application and return it.

    The target is reached through the connection's own proxy, on the thread that
    validated the session, so callers can assert on what it received.
    """
    monkeypatch.setattr(tools_module.VCSAddinIntegration, "load_addin", lambda *args, **kwargs: True)
    app = MagicMock()
    if run_side_effect is not None:
        app.Run.side_effect = run_side_effect
    else:
        app.Run.return_value = run_result

    class _Conn:
        def __init__(self, *_args, **_kwargs):
            pass

        def __enter__(self):
            return self

        def __exit__(self, *args):
            return False

        def connect(self):
            return app, MagicMock()

    monkeypatch.setattr(tools_module, "AccessConnection", _Conn)
    return app


@pytest.fixture(autouse=True)
def _bypass_access_gate(monkeypatch):
    """Unit tests call tool bodies directly — skip the real gate."""

    async def _immediate(tool, database, fn, is_async, /, *args, **kwargs):
        if is_async:
            return await fn(*args, **kwargs)
        return fn(*args, **kwargs)

    monkeypatch.setattr(
        tools_module,
        "get_access_gate",
        lambda: MagicMock(run_exclusive=_immediate),
    )


@pytest.fixture
def _config_installed_addin(monkeypatch):
    monkeypatch.setattr(
        tools_module,
        "get_config",
        lambda: {"ACCESS_VCS_ADDIN_PATH": INSTALLED_ADDIN},
    )


def test_development_copy_dispatches(_config_installed_addin, monkeypatch):
    """The development copy shares the install's file name but is a different file."""
    monkeypatch.setattr(
        tools_module,
        "validate_database_path",
        lambda p: Path(p),
    )
    monkeypatch.setattr(tools_module, "get_call_vba_timeout", lambda _t=None: 5.0)
    _fake_access(monkeypatch, run_result='{"version":"5.1.2"}')

    result = _call_vba(
        REPO_ADDIN,
        "VCS.API",
        ["GetVCSVersion"],
    )

    assert result["success"] is True


def test_rebuild_snapshot_captured_before_dispatch(
    _config_installed_addin, monkeypatch, tmp_path
):
    source = tmp_path / "Version Control.accda.src"
    logs = source / "logs"
    logs.mkdir(parents=True)
    status_file = logs / "rebuild-status.json"
    status_file.write_text(
        json.dumps({
            "status": "complete",
            "updated": "2026-08-17 16:46:20",
            "phaseStarted": "2026-08-17 16:45:36",
        }),
        encoding="utf-8",
    )

    host = tmp_path / "host.accdb"
    host.write_bytes(b"x")
    monkeypatch.setattr(
        tools_module,
        "validate_database_path",
        lambda p: Path(p),
    )
    monkeypatch.setattr(tools_module, "get_call_vba_timeout", lambda _t=None: 5.0)
    _fake_access(monkeypatch, run_result='{"status":"launched"}')

    result = _call_vba(
        str(host),
        "VCS.API",
        ["RebuildAddIn", str(source)],
    )

    assert result["success"] is True
    assert result["rebuild_status_file"] == str(status_file)
    assert result["rebuild_status_before"]["status"] == "complete"
    assert result["rebuild_status_before"]["updated"] == "2026-08-17 16:46:20"
    assert "mtime" in result["rebuild_status_before"]


def test_is_addin_api_resolved_name():
    with patch.object(
        tools_module,
        "get_config",
        return_value={"ACCESS_VCS_ADDIN_PATH": INSTALLED_ADDIN},
    ):
        assert tools_module._is_addin_api_resolved_name(RESOLVED_API)
        assert not tools_module._is_addin_api_resolved_name("SomeModule.Foo")


def test_describe_rebuild_attempt_correlates_phase_started():
    describe = tools_module._describe_rebuild_attempt
    launched = json.dumps({"status": "launched", "phaseStarted": "2026-08-21 09:15:00"})

    # The attempt stamped its own start, so the record on disk is now this run's.
    assert describe(launched, {"phaseStarted": "2026-08-17 16:45:36"}) == {
        "rebuild_phase_started": "2026-08-21 09:15:00",
        "rebuild_status_superseded": True,
    }

    # Same value as the snapshot: nothing new was written, so whatever is on disk
    # belongs to an earlier run and must not be read as this attempt's verdict.
    assert describe(launched, {"phaseStarted": "2026-08-21 09:15:00"}) == {
        "rebuild_phase_started": "2026-08-21 09:15:00",
        "rebuild_status_superseded": False,
    }

    assert describe(launched, None) == {"rebuild_phase_started": "2026-08-21 09:15:00"}
    assert describe(json.dumps({"status": "refused"}), None) == {}
    assert describe("not json", None) == {}
    assert describe(None, None) == {}


def test_snapshot_rebuild_status_reads_bom_encoded_file(tmp_path):
    """The add-in writes rebuild-status.json as UTF-8 with a BOM.

    Reading it as plain "utf-8" leaves the BOM in the string, json.load rejects it,
    and the snapshot degrades to read_error -- which silently disables the
    phaseStarted comparison that tells this attempt's record from a stale one.
    """
    logs = tmp_path / "logs"
    logs.mkdir()
    payload = {
        "status": "complete",
        "phaseStarted": "2026-08-21 13:12:15",
        "updated": "2026-08-21 13:13:04",
    }
    (logs / "rebuild-status.json").write_text(
        json.dumps(payload), encoding="utf-8-sig"
    )

    status_file, snapshot = tools_module._snapshot_rebuild_status(str(tmp_path))

    assert status_file == str(logs / "rebuild-status.json")
    assert "read_error" not in snapshot
    assert snapshot["status"] == "complete"
    assert snapshot["phaseStarted"] == "2026-08-21 13:12:15"

    # And the value has to be usable for the comparison it exists to serve.
    described = tools_module._describe_rebuild_attempt(
        json.dumps({"status": "launched", "phaseStarted": "2026-08-21 13:30:00"}),
        snapshot,
    )
    assert described["rebuild_status_superseded"] is True


def test_snapshot_rebuild_status_missing_and_malformed(tmp_path):
    logs = tmp_path / "logs"
    logs.mkdir()

    # No file at all: nothing to compare against, reported as absent rather than broken.
    assert tools_module._snapshot_rebuild_status(str(tmp_path))[1] is None

    (logs / "rebuild-status.json").write_text("{ truncated", encoding="utf-8")
    assert tools_module._snapshot_rebuild_status(str(tmp_path))[1]["read_error"] is True


def test_is_installed_addin_path_ignores_the_extension():
    with patch.object(
        tools_module,
        "get_config",
        return_value={"ACCESS_VCS_ADDIN_PATH": INSTALLED_ADDIN},
    ):
        check = tools_module._is_installed_addin_path

        assert check(INSTALLED_ADDIN)
        # A compiled install is the .accde built from the same .accda, and only one of
        # the two is ever named in ACCESS_VCS_ADDIN_PATH.
        assert check(INSTALLED_ADDIN.replace(".accda", ".accde"))
        assert check(INSTALLED_ADDIN.upper())
        # Same file name, different folder.
        assert not check(REPO_ADDIN)
        # A folder beside the install is not the install.
        assert not check(SOURCE_DIR)


def test_raw_target_call_uses_accepted_dispatcher(_config_installed_addin, monkeypatch):
    monkeypatch.setattr(tools_module, "validate_database_path", lambda p: Path(p))
    app = _fake_access(monkeypatch, run_result="echo")
    captured = {}
    def bounded(_permit, _app, procedure, args):
        captured.update(procedure=procedure, args=args)
        return {"success": True, "result": ["msaccess-vcs.call-vba-result/1", "returned", "echo"]}
    monkeypatch.setattr(tools_module, "_run_target", bounded)
    result = _call_vba(REPO_ADDIN, "TargetEcho", ["hello"])
    assert result["success"]
    assert captured["procedure"].endswith(".APIExecute")
    assert captured["args"][1:] == ["CallVBAResult", "TargetEcho", '["hello"]']


def test_rebuild_admission_failure_registers_no_callback(_config_installed_addin, monkeypatch):
    from msaccess_vcs_mcp.compatibility_session import AdmissionError, session_failure
    monkeypatch.setattr(tools_module, "validate_database_path", lambda p: Path(p))
    app = _fake_access(monkeypatch)
    def refuse(*_args):
        raise AdmissionError(session_failure("stale_instance"))
    monkeypatch.setattr(tools_module, "ensure_session", refuse)
    register = MagicMock()
    with pytest.raises(AdmissionError):
        tools_module._execute_call_vba(REPO_ADDIN, "VCS.API", ["RebuildAddIn", SOURCE_DIR], on_admitted=register)
    register.assert_not_called()
    app.Run.assert_not_called()
