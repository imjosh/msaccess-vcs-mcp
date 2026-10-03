"""Controlled X17 public transport boundaries; these are not live VBA evidence."""
import json
import uuid
from collections import Counter
from unittest.mock import Mock

import pytest

from msaccess_vcs_mcp import compatibility as c
from msaccess_vcs_mcp import compatibility_session as s
from msaccess_vcs_mcp.addin_integration import VCSAddinIntegration

pytestmark = pytest.mark.compatibility_session


class AddinTransport:
    def __init__(self):
        self.calls = Counter()
        self.instance = str(uuid.uuid4())
        self.version = "6.0.0-dev.17"
        self.sessions = {}
        self.payloads = []
        self.reject = False
        self.fail_command = False

    def Run(self, procedure, *args):
        method = procedure.rsplit(".", 1)[1]
        self.calls[method] += 1
        self.payloads.append((method, args))
        if method == "APIIdentity":
            return json.dumps(dict(success=True, addin_version=self.version,
                                   addin_instance=self.instance, protocol=s.PROTOCOL))
        if method == "APIHandshake":
            if self.reject:
                return json.dumps(s.session_failure("server_refused"))
            request = json.loads(args[0])
            self.sessions[request["session_id"]] = request
            return json.dumps(dict(success=True, **request))
        if method == "APIValidateSession":
            request = json.loads(args[0])
            entry = self.sessions.get(request["session_id"], {})
            return json.dumps(dict(success=entry.get("addin_instance") == self.instance and
                                   all(entry.get(k) == v for k, v in request.items())))
        if method in {"APIExecute", "APIExecuteAsync"}:
            if self.fail_command:
                raise RuntimeError("COM disconnected after possible dispatch")
            return "true"
        raise AssertionError(method)


@pytest.fixture
def transport(monkeypatch, tmp_path):
    from msaccess_vcs_mcp.access_com import instance_registry, process_qos
    s._sessions.clear()
    s._connections.clear()
    c._discovery_cache.clear()
    monkeypatch.setattr(process_qos, "pid_from_access_app", lambda app: 73)
    monkeypatch.setattr(instance_registry, "process_create_time", lambda pid: 100)
    path = tmp_path / "library.accda"
    path.write_bytes(b"installed")
    return AddinTransport(), str(path)


def test_cached_calls_send_no_versions_and_discover_once(transport, monkeypatch):
    app, path = transport
    read = Mock(return_value="6.0.0-dev.17")
    monkeypatch.setattr(c, "_read_installed_version", read)
    integration = VCSAddinIntegration(path)
    integration._app = app
    integration._addin_loaded = True
    app.CurrentDb = lambda: type("Database", (), {"Name": "disposable"})()
    for _ in range(10):
        assert c.inspect_installed_addin(path)["success"]
        assert integration.call_sync("GetOption", "ShowDebug") == "true"
    assert read.call_count == 1
    assert app.calls["APIIdentity"] == app.calls["APIHandshake"] == 1
    assert app.calls["APIExecute"] == 10
    assert app.calls["APIValidateSession"] == 9
    for method, args in app.payloads:
        if method != "APIHandshake":
            assert "server_version" not in str(args)
    for method, args in app.payloads:
        if method == "APIExecute":
            assert "session_id" in json.loads(args[0])


@pytest.mark.parametrize("change", ["reset", "reload", "access_restart", "pid_reuse", "server_restart", "path", "library", "requirement", "protocol"])
def test_invalidation_cannot_resurrect_an_old_pair(transport, monkeypatch, change):
    from msaccess_vcs_mcp.access_com import instance_registry
    from pathlib import Path
    app, path = transport
    first = s.ensure_session(app, path)
    if change in {"reset", "reload"}:
        app.instance = str(uuid.uuid4())
        app.sessions.clear()
    elif change in {"access_restart", "pid_reuse"}:
        monkeypatch.setattr(instance_registry, "process_create_time", lambda pid: 101)
    elif change == "server_restart":
        monkeypatch.setattr(s, "SERVER_INSTANCE", str(uuid.uuid4()))
    elif change == "path":
        path = str(Path(path).with_name("other.accda"))
        Path(path).write_bytes(b"other")
    elif change == "library":
        Path(path).write_bytes(b"replaced file")
    elif change == "requirement":
        monkeypatch.setattr(s, "ADDIN_REQUIREMENT", c.Requirement("6.0.1", "7.0.0"))
        app.version = "6.0.1"
    elif change == "protocol":
        monkeypatch.setattr(s, "PROTOCOL", "msaccess-vcs.session/2")
    second = s.ensure_session(app, path)
    assert first != second
    assert app.calls["APIHandshake"] == 2


def test_connections_do_not_share_sessions(transport):
    app, path = transport
    one, two = object(), object()
    with s.connection_scope(one):
        first = s.ensure_session(app, path)
    with s.connection_scope(two):
        second = s.ensure_session(app, path)
    with s.connection_scope(one):
        assert s.ensure_session(app, path) == first
    assert first != second
    assert app.calls["APIHandshake"] == 2


@pytest.mark.parametrize("version", [None, "", "garbage", "5.9.9", "7.0.0", "6.0.0-rc.1"])
def test_server_refusal_dispatches_nothing(transport, version):
    app, path = transport
    app.version = version
    with pytest.raises(s.AdmissionError) as failure:
        s.ensure_session(app, path)
    assert failure.value.result["success"] is False
    assert app.calls["APIHandshake"] == 0
    assert app.calls["APIExecute"] == 0
    assert not s._sessions


def test_addin_refusal_leaves_no_partial_server_admission(transport):
    app, path = transport
    app.reject = True
    with pytest.raises(s.AdmissionError):
        s.ensure_session(app, path)
    assert not s._sessions
    assert not app.sessions


def test_unavailable_identity_never_reuses_cache(transport, monkeypatch):
    from msaccess_vcs_mcp.access_com import instance_registry
    app, path = transport
    s.ensure_session(app, path)
    monkeypatch.setattr(instance_registry, "process_create_time", lambda pid: None)
    with pytest.raises(s.AdmissionError):
        s.ensure_session(app, path)
    assert app.calls["APIHandshake"] == 1
    assert app.calls["APIValidateSession"] == 0


def test_missing_named_api_refuses_without_legacy_dispatch(transport):
    app, path = transport
    app.Run = Mock(side_effect=RuntimeError("2517 procedure not found"))
    with pytest.raises(s.AdmissionError):
        s.ensure_session(app, path)
    assert app.Run.call_count == 1
    assert app.Run.call_args.args[0].endswith(".APIIdentity")


def test_command_exception_never_replays(transport):
    app, path = transport
    app.fail_command = True
    app.CurrentDb = lambda: type("Database", (), {"Name": "fixture"})()
    integration = VCSAddinIntegration(path)
    integration._app = app
    integration._addin_loaded = True
    with pytest.raises(RuntimeError, match="possible dispatch"):
        integration.call_sync("SetOption", "ShowDebug", False)
    assert app.calls["APIExecute"] == 1
    assert app.calls["APIHandshake"] == 1


@pytest.mark.parametrize("asynchronous", [False, True])
def test_addin_boundary_refusal_preserves_fields_and_never_replays(transport, asynchronous):
    app, path = transport
    original = app.Run
    refusal = s.session_failure("reset_before_dispatch")
    def run(procedure, *args):
        if procedure.endswith((".APIExecute", ".APIExecuteAsync")):
            app.calls["rejected_dispatch_attempt"] += 1
            return json.dumps(refusal)
        return original(procedure, *args)
    app.Run = run
    app.CurrentDb = lambda: type("Database", (), {"Name": "disposable"})()
    integration = VCSAddinIntegration(path)
    integration._app = app
    integration._addin_loaded = True
    with pytest.raises(s.AdmissionError) as failure:
        if asynchronous:
            integration.call_async("{}", "Export")
        else:
            integration.call_sync("GetOption", "ShowDebug")
    assert failure.value.result == refusal
    assert app.calls["rejected_dispatch_attempt"] == 1
    assert app.calls["APIHandshake"] == 1
