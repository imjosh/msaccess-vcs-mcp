"""Public connection seam: closed target cannot execute before loaded admission."""
from pathlib import Path
from unittest.mock import Mock

import pytest

from msaccess_vcs_mcp import compatibility_session as sessions
from msaccess_vcs_mcp.access_com import connection, instance_registry, process_qos
from msaccess_vcs_mcp.addin_integration import VCSAddinIntegration

pytestmark = pytest.mark.compatibility_session


@pytest.fixture
def host(monkeypatch, tmp_path):
    target = tmp_path / "target.accdb"
    target.write_bytes(b"target startup sentinel")
    app = Mock()
    events = []
    app.NewCurrentDatabase.side_effect = lambda path: (events.append("neutral"), Path(path).touch())
    app.OpenCurrentDatabase.side_effect = lambda path: events.append("target-autoexec")
    app.CloseCurrentDatabase.side_effect = lambda: events.append("close-neutral")
    app.Quit.side_effect = lambda *args: events.append("quit-owned")
    app.CurrentDb.return_value.Name = str(target)
    monkeypatch.setattr(connection, "create_isolated_access_app", lambda: app)
    monkeypatch.setattr(connection, "COM_AVAILABLE", True)
    monkeypatch.setattr(VCSAddinIntegration, "_find_access_in_rot", lambda path: None)
    monkeypatch.setattr(process_qos, "pid_from_access_app", lambda app: 73)
    monkeypatch.setattr(process_qos, "process_is_alive", lambda pid: False)
    monkeypatch.setattr(process_qos, "prefer_full_power_app", lambda app: None)
    monkeypatch.setattr(instance_registry, "process_create_time", lambda pid: 100)
    monkeypatch.setenv("ACCESS_VCS_OWNED_INSTANCES_PATH", str(tmp_path / "owned.json"))
    monkeypatch.setattr(instance_registry, "list_access_pids_or_none", lambda: {73})
    folder = tmp_path / "neutral"
    folder.mkdir()
    monkeypatch.setattr(connection.tempfile, "mkdtemp", lambda **kwargs: str(folder))
    monkeypatch.setattr(sessions, "ensure_session", lambda *args: events.append("validate-before-open"))
    return app, events, target, folder


@pytest.mark.parametrize("reason", ["server_incompatible", "named_api_missing"])
def test_refused_bootstrap_never_opens_or_mutates_target(host, monkeypatch, reason):
    app, events, target, folder = host
    def refuse(self, *args, **kwargs):
        events.append("handshake-refused")
        raise sessions.AdmissionError(sessions.session_failure(reason))
    monkeypatch.setattr(VCSAddinIntegration, "load_addin", refuse)
    before = target.read_bytes()
    with sessions.target_admission("library.accda"):
        with pytest.raises(sessions.AdmissionError):
            connection.AccessConnection(str(target)).connect()
    assert events == ["neutral", "handshake-refused", "quit-owned"]
    app.OpenCurrentDatabase.assert_not_called()
    assert target.read_bytes() == before
    assert not folder.exists()


def test_accepted_bootstrap_admits_and_validates_before_target_autoexec(host, monkeypatch):
    app, events, target, folder = host
    monkeypatch.setattr(VCSAddinIntegration, "load_addin", lambda *a, **kw: events.append("mutual-handshake"))
    with sessions.target_admission("library.accda"):
        conn = connection.AccessConnection(str(target))
        assert conn.get_app() is app
    assert events == ["neutral", "mutual-handshake", "close-neutral", "validate-before-open", "target-autoexec"]
    assert not folder.exists()


def test_existing_user_host_is_admitted_without_open_or_quit(host, monkeypatch):
    app, events, target, folder = host
    monkeypatch.setattr(VCSAddinIntegration, "_find_access_in_rot", lambda path: app)
    monkeypatch.setattr(VCSAddinIntegration, "load_addin", lambda *a, **kw: events.append("mutual-handshake"))
    with sessions.target_admission("library.accda"):
        assert connection.AccessConnection(str(target)).get_app() is app
    assert events == ["mutual-handshake"]
    app.NewCurrentDatabase.assert_not_called()
    app.OpenCurrentDatabase.assert_not_called()
    app.Quit.assert_not_called()


def test_probe_timeout_does_not_make_blocking_com_cleanup_calls(host, monkeypatch):
    app, events, target, folder = host
    stalled = False
    def identity(app):
        if stalled:
            pytest.fail("COM after timeout")
        return 73
    monkeypatch.setattr(process_qos, "pid_from_access_app", identity)
    def timeout(*args, **kwargs):
        # Once the probe stalls, even a PID read through app.hWndAccessApp
        # can block; cleanup must use the identity captured before probing.
        nonlocal stalled
        stalled = True
        raise TimeoutError("bounded probe stalled")
    monkeypatch.setattr(VCSAddinIntegration, "load_addin", timeout)
    with sessions.target_admission("library.accda"):
        with pytest.raises(TimeoutError):
            connection.AccessConnection(str(target)).get_app()
    app.OpenCurrentDatabase.assert_not_called()
    app.Quit.assert_not_called()
    records = instance_registry.list_owned()
    assert len(records) == 1 and records[0].create_time == 100
    assert Path(records[0].database_path) == folder / "bootstrap.accdb"
