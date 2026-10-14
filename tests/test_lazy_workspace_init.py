"""Tests for lazy MCP workspace-roots .env discovery in tools.py."""

import asyncio
from contextlib import asynccontextmanager
import json
import os
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import pytest
import anyio
from mcp import types
from mcp.server.models import InitializationOptions
from mcp.server.session import ServerSession
from mcp.shared.message import SessionMessage
from pydantic import ValidationError

import msaccess_vcs_mcp.config as config_module
import msaccess_vcs_mcp.tools as tools_module
import msaccess_vcs_mcp.usage_logging as logging_module
from msaccess_vcs_mcp.tools import _ensure_env_loaded, _file_uri_to_path


@asynccontextmanager
async def _wire_session(responses):
    """Feed raw JSON-RPC results through the SDK's receive/validation boundary."""
    to_server, server_read = anyio.create_memory_object_stream(1)
    server_write, from_server = anyio.create_memory_object_stream(1)
    requests = []

    async def client():
        async with from_server, to_server:
            async for message in from_server:
                request = message.message.root
                requests.append(request)
                response = responses[min(len(requests) - 1, len(responses) - 1)]
                await to_server.send(SessionMessage(message=types.JSONRPCMessage(
                    types.JSONRPCResponse(jsonrpc="2.0", id=request.id, result=response)
                )))

    async with anyio.create_task_group() as tasks:
        tasks.start_soon(client)
        async with ServerSession(server_read, server_write, InitializationOptions(
            server_name="m58-test", server_version="1", capabilities=types.ServerCapabilities(),
        ), stateless=True) as session:
            try:
                yield session, requests
            finally:
                tasks.cancel_scope.cancel()


def _metadata_tool(monkeypatch, session):
    monkeypatch.setattr(tools_module.mcp, "get_context", lambda: SimpleNamespace(session=session))
    monkeypatch.setattr(tools_module, "inspect_installed_addin", lambda **kw: {
        "installed_version": "6.0.0", "addin_path": "mock-install.accda",
    })
    return tools_module.mcp._tool_manager.get_tool("vcs_get_version_info").fn


class TestSDKWorkspaceRecovery:
    @pytest.mark.parametrize("native_format", ["backslashes", "slashes"])
    def test_native_windows_root_loads_project_via_metadata(self, tmp_path, monkeypatch, native_format):
        monkeypatch.chdir(tmp_path)
        workspace = tmp_path / "project with spaces % #"
        workspace.mkdir()
        database = str(workspace / "intended.accdb")
        (workspace / ".env").write_text(f"ACCESS_VCS_DATABASE='{database}'\n", encoding="utf-8")

        async def exercise():
            uri = str(workspace) if native_format == "backslashes" else workspace.as_posix()
            async with _wire_session([{"roots": [{"uri": uri, "name": "project"}]}]) as (session, requests):
                # Confirm this is the actual Cursor incompatibility, not a permissive mock.
                with pytest.raises(ValidationError, match="file"):
                    await session.list_roots()
                tool = _metadata_tool(monkeypatch, session)
                result = await tool()
                assert result["target_database"] == database
                assert result["project_configuration"] == {
                    "project_root": str(workspace), "resolution_method": "mcp_workspace_roots",
                }
                assert requests[-1].method == "roots/list"
        asyncio.run(exercise())

    def test_failed_response_recovers_on_later_metadata_call(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        workspace = tmp_path / "project"
        workspace.mkdir()
        (workspace / ".env").write_text("ACCESS_VCS_DATABASE=recovered.accdb\n", encoding="utf-8")

        async def exercise():
            async with _wire_session([
                {"roots": [{"uri": "https://invalid.example/project"}]},
                {"roots": [{"uri": workspace.as_uri()}]},
            ]) as (session, requests):
                tool = _metadata_tool(monkeypatch, session)
                assert (await tool())["target_database"] == ""
                # Advance the retry clock without sleeping, once the policy exists.
                monkeypatch.setattr(tools_module, "_lazy_init_next_retry", 0.0, raising=False)
                assert (await tool())["target_database"] == "recovered.accdb"
                assert len(requests) == 2
                events = _read_diagnostic_events()
                assert any(e["event"] == "list_roots_failed" and e["reason"] == "invalid_client_response" for e in events)
                assert any(e["event"] == "lazy_init_loaded" and e["recovered"] for e in events)
        asyncio.run(exercise())

    @pytest.mark.parametrize("root", [
        {"uri": "https://invalid.example/project"}, {"uri": "C:relative"},
        {"uri": "relative/path"}, {"uri": 123}, {"name": "missing uri"},
        {"uri": "C:\\bad\x00path"}, {"uri": "C:\\bad?path"},
        {"uri": "C:\\absolute", "name": 123},
    ])
    def test_invalid_roots_remain_rejected_by_sdk(self, root, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)

        async def exercise():
            async with _wire_session([{"roots": [root]}]) as (session, requests):
                result = await _metadata_tool(monkeypatch, session)()
                assert result["target_database"] == ""
                assert len(requests) == 1
                assert any(e["event"] == "list_roots_failed" and e["error_type"] == "ValidationError"
                           for e in _read_diagnostic_events())
        asyncio.run(exercise())

    def test_file_uri_selection_is_stable_through_reload(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        workspace = tmp_path / "file uri with spaces"
        workspace.mkdir()
        env = workspace / ".env"
        env.write_text("ACCESS_VCS_DATABASE=first.accdb\n", encoding="utf-8")
        other = tmp_path / "other"
        other.mkdir()
        (other / ".env").write_text("ACCESS_VCS_DATABASE=wrong.accdb\n", encoding="utf-8")

        async def exercise():
            async with _wire_session([
                {"roots": [{"uri": workspace.as_uri()}]},
                {"roots": [{"uri": other.as_uri()}]},
            ]) as (session, requests):
                tool = _metadata_tool(monkeypatch, session)
                assert (await tool())["target_database"] == "first.accdb"
                assert (await tool())["target_database"] == "first.accdb"
                old_mtime = env.stat().st_mtime
                env.write_text("ACCESS_VCS_DATABASE=edited.accdb\n", encoding="utf-8")
                os.utime(env, (old_mtime + 10, old_mtime + 10))
                assert (await tool())["target_database"] == "edited.accdb"
                assert config_module.get_project_root_info()["project_root"] == str(workspace)
                assert len(requests) == 1
        asyncio.run(exercise())

    @pytest.mark.parametrize("has_env", [False, True])
    def test_explicit_project_never_switches_to_client_root(self, tmp_path, monkeypatch, has_env):
        monkeypatch.chdir(tmp_path)
        selected = tmp_path / "selected"
        selected.mkdir()
        monkeypatch.setenv("ACCESS_VCS_PROJECT_DIR", str(selected))
        if has_env:
            (selected / ".env").write_text("ACCESS_VCS_DATABASE=base.accdb\n", encoding="utf-8")
            (selected / ".env.local").write_text("ACCESS_VCS_DATABASE=personal.accdb\n", encoding="utf-8")
        other = tmp_path / "other"
        other.mkdir()
        (other / ".env").write_text("ACCESS_VCS_DATABASE=wrong.accdb\n", encoding="utf-8")
        monkeypatch.setenv("ACCESS_VCS_DATABASE", "launch.accdb")

        async def exercise():
            async with _wire_session([{"roots": [{"uri": other.as_uri()}]}]) as (session, requests):
                result = await _metadata_tool(monkeypatch, session)()
                assert result["target_database"] == ("personal.accdb" if has_env else "launch.accdb")
                assert result["project_configuration"]["resolution_method"] == "ACCESS_VCS_PROJECT_DIR"
                assert config_module.get_project_root_info()["project_root"] == str(selected)
                assert not requests
        asyncio.run(exercise())

    def test_retry_cooldown_and_diagnostics_are_bounded(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)

        async def exercise():
            async with _wire_session([{"roots": []}]) as (session, requests):
                tool = _metadata_tool(monkeypatch, session)
                for attempt in range(7):
                    monkeypatch.setattr(tools_module, "_lazy_init_next_retry", 0.0)
                    assert (await tool())["target_database"] == ""
                    before = len(_read_diagnostic_events())
                    for _ in range(5):
                        await tool()
                    # Tool usage may be logged separately; discovery remains quiet.
                    assert len(_read_diagnostic_events()) == before
                assert len(requests) == 7
                retries = [e["retry_after_seconds"] for e in _read_diagnostic_events()
                           if e["event"] == "lazy_init_retry_scheduled"]
                assert retries == [2, 4, 8, 16, 30, 30, 30]
                assert any(e.get("reason") == "unresolved_project" for e in _read_diagnostic_events())
        asyncio.run(exercise())

    def test_missing_env_can_appear_on_later_call(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        workspace = tmp_path / "later"
        workspace.mkdir()

        async def exercise():
            async with _wire_session([{"roots": [{"uri": workspace.as_uri()}]}]) as (session, requests):
                tool = _metadata_tool(monkeypatch, session)
                assert (await tool())["target_database"] == ""
                assert any(e.get("reason") == "missing_env_file" for e in _read_diagnostic_events())
                (workspace / ".env.local").write_text("ACCESS_VCS_DATABASE=later.accdb\n", encoding="utf-8")
                monkeypatch.setattr(tools_module, "_lazy_init_next_retry", 0.0)
                assert (await tool())["target_database"] == "later.accdb"
                assert len(requests) == 2
        asyncio.run(exercise())

    def test_timeout_recovers_without_concurrent_roots_requests(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        monkeypatch.setattr(tools_module, "WORKSPACE_ROOTS_TIMEOUT_SEC", 0.05)

        async def exercise():
            entered = asyncio.Event()
            async def hung(*args):
                entered.set()
                await asyncio.Event().wait()
            session = SimpleNamespace(send_request=AsyncMock(side_effect=hung))
            tool = _metadata_tool(monkeypatch, session)
            first = asyncio.create_task(tool())
            await entered.wait()
            assert (await tool())["target_database"] == ""
            assert (await first)["target_database"] == ""
            assert session.send_request.call_count == 1
            assert not tools_module._lazy_init_in_flight
            assert any(e.get("error_type") == "TimeoutError" for e in _read_diagnostic_events())
            (tmp_path / "workspace").mkdir()
            workspace = tmp_path / "workspace"
            (workspace / ".env").write_text("ACCESS_VCS_DATABASE=after-timeout.accdb\n", encoding="utf-8")
            session.send_request = AsyncMock(return_value=types.ListRootsResult(
                roots=[types.Root(uri=workspace.as_uri())],
            ))
            monkeypatch.setattr(tools_module, "_lazy_init_next_retry", 0.0)
            assert (await tool())["target_database"] == "after-timeout.accdb"
        asyncio.run(exercise())


@pytest.fixture(autouse=True)
def _reset_state(monkeypatch, tmp_path_factory, tmp_path):
    """Clear the lazy-init flag and cached config state between tests.

    Also redirects the diagnostic stream to a per-test temp directory
    so the always-on diagnostic log doesn't write to the user's real
    ``~/.msaccess-vcs-mcp/logs/`` during the suite. Tests can read
    ``logging_module._diagnostic_file`` to inspect the events that
    fired.
    """
    tools_module._lazy_init_attempted = False
    tools_module._lazy_init_skip_logged = False
    monkeypatch.setattr(tools_module, "_lazy_init_failures", 0, raising=False)
    monkeypatch.setattr(tools_module, "_lazy_init_next_retry", 0.0, raising=False)
    monkeypatch.setattr(tools_module, "_lazy_init_in_flight", False, raising=False)
    # Stop discovery at this disposable project, never the checkout's real .env.
    (tmp_path / "pyproject.toml").write_text("", encoding="utf-8")
    config_module._env_loaded = False
    config_module._is_reload = False
    config_module._project_root = None
    config_module._project_root_method = None
    config_module._env_mtimes = {}

    logging_module._diagnostic_handler = None
    logging_module._diagnostic_file = None
    logging_module._diagnostic_initialized = False
    logging_module._diagnostic_enabled = False
    logging_module._diagnostic_disabled_reason = None

    for key in list(os.environ):
        if key.startswith("ACCESS_VCS_"):
            monkeypatch.delenv(key, raising=False)

    diag_dir = tmp_path_factory.mktemp("diag")
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(diag_dir))

    yield

    tools_module._lazy_init_attempted = False
    tools_module._lazy_init_skip_logged = False
    config_module._env_loaded = False
    config_module._is_reload = False
    config_module._project_root = None
    config_module._project_root_method = None
    config_module._env_mtimes = {}

    if logging_module._diagnostic_handler is not None:
        try:
            logging_module._diagnostic_handler.close()
        except Exception:
            pass
    logging_module._diagnostic_handler = None
    logging_module._diagnostic_file = None
    logging_module._diagnostic_initialized = False
    logging_module._diagnostic_enabled = False
    logging_module._diagnostic_disabled_reason = None


def _read_diagnostic_events() -> list[dict]:
    """Return all events recorded in the active diagnostic log file.

    Reads from the configured ``ACCESS_VCS_DIAGNOSTIC_LOG_DIR`` rather
    than ``logging_module._diagnostic_file`` so the helper still works
    after ``reset_logging()`` has cleared the in-memory handle (which
    happens whenever ``initialize_from_workspace`` is called and
    triggers a logging reload mid-test).
    """
    diag_dir_str = os.environ.get("ACCESS_VCS_DIAGNOSTIC_LOG_DIR")
    if not diag_dir_str:
        return []
    path = Path(diag_dir_str) / f"vcs-mcp-diagnostic-{logging_module._INSTANCE_ID}.jsonl"
    if not path.exists():
        return []
    return [
        json.loads(line)
        for line in path.read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]


def _make_ctx(root_uris: list[str]) -> MagicMock:
    """Build a fake Context exposing ``session.list_roots()`` -> roots."""
    roots = [SimpleNamespace(uri=uri) for uri in root_uris]
    list_roots_result = SimpleNamespace(roots=roots)

    session = MagicMock()
    session.send_request = AsyncMock(return_value=list_roots_result)

    ctx = MagicMock()
    ctx.session = session
    return ctx


class TestFileUriToPath:
    def test_converts_unix_uri(self):
        assert _file_uri_to_path("file:///home/me/proj") == Path("/home/me/proj")

    def test_converts_windows_uri(self):
        # file:///C:/foo -> /C:/foo, leading slash stripped on Windows-style paths.
        result = _file_uri_to_path("file:///C:/Users/me/proj")
        assert str(result) in (r"C:\Users\me\proj", "C:/Users/me/proj")

    def test_decodes_percent_encoding(self):
        result = _file_uri_to_path("file:///tmp/with%20space")
        assert "with space" in str(result)

    def test_returns_none_for_non_file_uri(self):
        assert _file_uri_to_path("http://example.com") is None


class TestEnsureEnvLoaded:
    def test_loads_env_from_first_root_with_env(self, tmp_path, monkeypatch):
        elsewhere = tmp_path / "elsewhere"
        elsewhere.mkdir()
        monkeypatch.chdir(elsewhere)

        # Two workspace roots; only the second has an .env file.
        empty_root = tmp_path / "empty_workspace"
        empty_root.mkdir()
        real_root = tmp_path / "real_workspace"
        real_root.mkdir()
        (real_root / ".env").write_text(
            "ACCESS_VCS_ENABLE_LOGGING=true\n"
            "ACCESS_VCS_DATABASE=ws.accdb\n",
            encoding="utf-8",
        )

        ctx = _make_ctx([
            empty_root.as_uri(),
            real_root.as_uri(),
        ])

        asyncio.run(_ensure_env_loaded(ctx))

        assert os.environ.get("ACCESS_VCS_ENABLE_LOGGING") == "true"
        assert os.environ.get("ACCESS_VCS_DATABASE") == "ws.accdb"
        assert tools_module._lazy_init_attempted is True

    def test_skips_when_ctx_is_none(self):
        asyncio.run(_ensure_env_loaded(None))
        # No exception, flag NOT flipped (so a real ctx call later still works).
        assert tools_module._lazy_init_attempted is False

    def test_only_runs_once(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        (tmp_path / ".env").write_text("ACCESS_VCS_DATABASE=selected.accdb\n", encoding="utf-8")
        ctx = _make_ctx([tmp_path.as_uri()])
        asyncio.run(_ensure_env_loaded(ctx))
        assert tools_module._lazy_init_attempted is True

        # Second call must not invoke list_roots again.
        ctx.session.send_request.reset_mock()
        asyncio.run(_ensure_env_loaded(ctx))
        ctx.session.send_request.assert_not_called()

    def test_skips_when_startup_env_already_found(
        self, tmp_path, monkeypatch,
    ):
        """If .env was already discovered at startup (CWD walk), don't override."""
        startup_project = tmp_path / "startup"
        startup_project.mkdir()
        (startup_project / ".env").write_text(
            "ACCESS_VCS_DATABASE=startup.accdb\n", encoding="utf-8",
        )
        monkeypatch.chdir(startup_project)

        # Prime config so _project_root points at the startup dir.
        from msaccess_vcs_mcp.config import load_config
        load_config()
        assert os.environ.get("ACCESS_VCS_DATABASE") == "startup.accdb"

        # Now simulate a workspace root with a different .env.
        workspace = tmp_path / "workspace"
        workspace.mkdir()
        (workspace / ".env").write_text(
            "ACCESS_VCS_DATABASE=workspace.accdb\n", encoding="utf-8",
        )

        ctx = _make_ctx([workspace.as_uri()])
        asyncio.run(_ensure_env_loaded(ctx))

        # The startup-discovered value should still be in place.
        assert os.environ["ACCESS_VCS_DATABASE"] == "startup.accdb"
        ctx.session.send_request.assert_not_called()

    def test_handles_list_roots_failure_gracefully(
        self, tmp_path, monkeypatch,
    ):
        """list_roots() failure is captured as a diagnostic event with the
        underlying error type, instead of a stderr-only print.
        """
        monkeypatch.chdir(tmp_path)

        ctx = MagicMock()
        ctx.session = MagicMock()
        ctx.session.send_request = AsyncMock(side_effect=RuntimeError("not supported"))

        asyncio.run(_ensure_env_loaded(ctx))

        events = _read_diagnostic_events()
        event_names = [e["event"] for e in events]
        assert "lazy_init_started" in event_names
        assert "list_roots_failed" in event_names
        failure = next(e for e in events if e["event"] == "list_roots_failed")
        assert failure["error_type"] == "RuntimeError"
        assert "not supported" in failure["error"]
        assert tools_module._lazy_init_attempted is False


class TestDiagnosticEvents:
    """Verify every branch of _ensure_env_loaded emits the right event."""

    def test_emits_no_ctx_skip(self):
        """ctx=None records lazy_init_skipped(reason='no_ctx')."""
        asyncio.run(_ensure_env_loaded(None))

        events = _read_diagnostic_events()
        skipped = [e for e in events if e["event"] == "lazy_init_skipped"]
        assert any(e.get("reason") == "no_ctx" for e in skipped)

    def test_emits_already_attempted_skip(self, tmp_path, monkeypatch):
        """A repeated call records lazy_init_skipped(reason='already_attempted')."""
        monkeypatch.chdir(tmp_path)
        (tmp_path / ".env").write_text("ACCESS_VCS_DATABASE=selected.accdb\n", encoding="utf-8")
        ctx = _make_ctx([])

        asyncio.run(_ensure_env_loaded(ctx))
        # Re-run; the flag is already set so we should hit the early
        # 'already_attempted' branch.
        asyncio.run(_ensure_env_loaded(ctx))

        events = _read_diagnostic_events()
        skipped = [e for e in events if e["event"] == "lazy_init_skipped"]
        assert any(e.get("reason") == "already_attempted" for e in skipped)

    def test_emits_lazy_init_started_and_list_roots_response(
        self, tmp_path, monkeypatch,
    ):
        """A successful path emits started + list_roots_response + loaded."""
        elsewhere = tmp_path / "elsewhere"
        elsewhere.mkdir()
        monkeypatch.chdir(elsewhere)

        workspace = tmp_path / "workspace"
        workspace.mkdir()
        (workspace / ".env").write_text("ACCESS_VCS_DATABASE=ws.accdb\n", encoding="utf-8")

        ctx = _make_ctx([workspace.as_uri()])
        asyncio.run(_ensure_env_loaded(ctx))

        events = _read_diagnostic_events()
        names = [e["event"] for e in events]
        assert "lazy_init_started" in names
        assert "list_roots_response" in names
        assert "lazy_init_loaded" in names

        started = next(e for e in events if e["event"] == "lazy_init_started")
        assert started["startup_root_has_env"] is False

        response = next(e for e in events if e["event"] == "list_roots_response")
        assert workspace.as_uri() in response["roots"]

        loaded = next(e for e in events if e["event"] == "lazy_init_loaded")
        assert loaded["workspace"] == str(workspace)

    def test_emits_no_env_in_roots(self, tmp_path, monkeypatch):
        """When list_roots returns workspaces without .env, we record
        lazy_init_no_env_in_roots with the URIs we tried.
        """
        elsewhere = tmp_path / "elsewhere"
        elsewhere.mkdir()
        monkeypatch.chdir(elsewhere)

        empty_a = tmp_path / "empty_a"
        empty_b = tmp_path / "empty_b"
        empty_a.mkdir()
        empty_b.mkdir()

        ctx = _make_ctx([empty_a.as_uri(), empty_b.as_uri()])
        asyncio.run(_ensure_env_loaded(ctx))

        events = _read_diagnostic_events()
        names = [e["event"] for e in events]
        assert "lazy_init_no_env_in_roots" in names
        assert "lazy_init_loaded" not in names

        no_env = next(e for e in events if e["event"] == "lazy_init_no_env_in_roots")
        assert empty_a.as_uri() in no_env["roots"]
        assert empty_b.as_uri() in no_env["roots"]

    def test_emits_startup_env_present_skip(self, tmp_path, monkeypatch):
        """If startup discovery already found an .env, we skip with
        reason='startup_env_present' (and never call list_roots).
        """
        startup_project = tmp_path / "startup"
        startup_project.mkdir()
        (startup_project / ".env").write_text(
            "ACCESS_VCS_DATABASE=startup.accdb\n", encoding="utf-8",
        )
        monkeypatch.chdir(startup_project)

        from msaccess_vcs_mcp.config import load_config
        load_config()

        workspace = tmp_path / "workspace"
        workspace.mkdir()
        (workspace / ".env").write_text("X=1\n", encoding="utf-8")

        ctx = _make_ctx([workspace.as_uri()])
        asyncio.run(_ensure_env_loaded(ctx))

        events = _read_diagnostic_events()
        skipped = [e for e in events if e["event"] == "lazy_init_skipped"]
        assert any(e.get("reason") == "startup_env_present" for e in skipped)
        ctx.session.send_request.assert_not_called()

    def test_emits_no_session_skip_for_request_unavailable_ctx(self, tmp_path, monkeypatch):
        """Context.request_context raises ValueError outside a request.

        Simulates ``mcp.get_context()`` being called outside a real MCP
        request (e.g. unit tests calling the wrapper directly). The
        helper records ``reason='no_session'`` and the lazy-init flag is
        NOT flipped, so a later real request can still drive discovery.
        """
        monkeypatch.chdir(tmp_path)

        ctx = MagicMock()
        type(ctx).session = property(
            lambda self: (_ for _ in ()).throw(
                ValueError("Context is not available outside of a request")
            )
        )

        asyncio.run(_ensure_env_loaded(ctx))

        events = _read_diagnostic_events()
        skipped = [e for e in events if e["event"] == "lazy_init_skipped"]
        assert any(e.get("reason") == "no_session" for e in skipped)
        assert tools_module._lazy_init_attempted is False


class TestVcsToolWrapper:
    """Verify ``vcs_tool`` wires lazy init for *every* registered tool.

    The contract: regardless of whether the wrapped function is sync or
    async, and regardless of whether it declares a ``ctx: Context``
    parameter, the wrapper must:
      1. Pull the active request context via ``mcp.get_context()``.
      2. Invoke ``_ensure_env_loaded`` with it.
      3. Refresh ``load_config()``.
      4. Dispatch the tool body (await for async, call for sync).

    These tests stub ``mcp.get_context`` and ``_ensure_env_loaded`` to
    verify the orchestration without actually calling MCP/Access COM.
    """

    def _get_registered_callable(self, fn_name: str):
        """Pull a FastMCP tool by registered name.

        ``mcp.tool()`` defaults the tool name to the wrapped callable's
        ``__name__`` (which ``functools.wraps`` preserves), so we look up
        by the wrapped function's name.
        """
        manager = tools_module.mcp._tool_manager
        for tool in manager.list_tools():
            if tool.name == fn_name:
                return tool.fn
        raise AssertionError(f"tool {fn_name!r} not registered")

    def test_sync_tool_triggers_lazy_init(self, monkeypatch):
        """A sync tool registered via @vcs_tool still drives _ensure_env_loaded."""
        ensure_calls = []

        async def fake_ensure(ctx):
            ensure_calls.append(ctx)

        monkeypatch.setattr(tools_module, "_ensure_env_loaded", fake_ensure)

        sentinel_ctx = MagicMock(name="ctx")
        monkeypatch.setattr(
            tools_module.mcp, "get_context", lambda: sentinel_ctx
        )
        monkeypatch.setattr(tools_module, "load_config", lambda: {})

        body_calls = []

        @tools_module.vcs_tool("test_sync_tool")
        def test_sync_tool_fn(arg: str = "default") -> dict:
            body_calls.append(arg)
            return {"ok": True, "arg": arg}

        wrapper = self._get_registered_callable("test_sync_tool_fn")
        result = asyncio.run(wrapper(arg="hello"))

        assert result == {"ok": True, "arg": "hello"}
        assert ensure_calls == [sentinel_ctx]
        assert body_calls == ["hello"]

    def test_async_tool_triggers_lazy_init(self, monkeypatch):
        """Async tools (with or without ctx param) drive _ensure_env_loaded."""
        ensure_calls = []

        async def fake_ensure(ctx):
            ensure_calls.append(ctx)

        monkeypatch.setattr(tools_module, "_ensure_env_loaded", fake_ensure)

        sentinel_ctx = MagicMock(name="ctx")
        monkeypatch.setattr(
            tools_module.mcp, "get_context", lambda: sentinel_ctx
        )
        monkeypatch.setattr(tools_module, "load_config", lambda: {})

        @tools_module.vcs_tool("test_async_tool")
        async def test_async_tool_fn(arg: str = "default") -> dict:
            return {"ok": True, "arg": arg}

        wrapper = self._get_registered_callable("test_async_tool_fn")
        result = asyncio.run(wrapper(arg="world"))

        assert result == {"ok": True, "arg": "world"}
        assert ensure_calls == [sentinel_ctx]

    def test_wrapper_handles_get_context_failure(self, monkeypatch):
        """If mcp.get_context() raises, wrapper passes None and proceeds."""
        ensure_calls = []

        async def fake_ensure(ctx):
            ensure_calls.append(ctx)

        monkeypatch.setattr(tools_module, "_ensure_env_loaded", fake_ensure)

        def boom():
            raise RuntimeError("no request context")

        monkeypatch.setattr(tools_module.mcp, "get_context", boom)
        monkeypatch.setattr(tools_module, "load_config", lambda: {})

        @tools_module.vcs_tool("test_no_ctx_tool")
        def test_no_ctx_tool_fn() -> dict:
            return {"ok": True}

        wrapper = self._get_registered_callable("test_no_ctx_tool_fn")
        result = asyncio.run(wrapper())

        assert result == {"ok": True}
        assert ensure_calls == [None]

    def test_repeated_skip_logged_only_once(self, tmp_path, monkeypatch):
        """After lazy init runs, repeated 'already_attempted' skips don't
        spam the diagnostic log on every tool call. The first repeat
        emits an event (so operators can verify the cache is working);
        subsequent repeats are silent.
        """
        monkeypatch.chdir(tmp_path)

        (tmp_path / ".env").write_text("ACCESS_VCS_DATABASE=selected.accdb\n", encoding="utf-8")
        ctx = _make_ctx([])
        asyncio.run(_ensure_env_loaded(ctx))

        # Repeat enough times to dwarf any noisy implementation.
        for _ in range(5):
            asyncio.run(_ensure_env_loaded(ctx))

        events = _read_diagnostic_events()
        skipped = [
            e for e in events
            if e["event"] == "lazy_init_skipped"
            and e.get("reason") == "already_attempted"
        ]
        assert len(skipped) == 1, (
            f"expected exactly one 'already_attempted' diagnostic; got {len(skipped)}"
        )
