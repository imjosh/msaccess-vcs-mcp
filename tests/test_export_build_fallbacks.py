"""M35: whole-export and database-build fallbacks never invent success.

Export, FullExport, ExportVBA, Build and BuildAs are Subs or form starts, so
their API return is Empty. Any path without a terminal callback therefore reports M15's
``completion_unconfirmed`` result, never ``success: true``. The tools run
against a real ``VCSAddinIntegration`` over a fake ``Application.Run``, so each
test asserts the add-in method that was actually dispatched.
"""

import asyncio
import json
import os
from contextlib import contextmanager
from unittest.mock import MagicMock, Mock, patch

import pytest

from msaccess_vcs_mcp.addin_integration import (
    VCSAddinIntegration,
    start_only_result,
    unconfirmed_start_result,
)
from msaccess_vcs_mcp.operation_manager import OperationManager

REFUSED = "VCS_API_REFUSED: VCS Export refused a call to 'Export' while another API command was running."
ASYNC_STARTED = '{"async": true, "timeout_ms": 5000}'
CAPABLE = '{"success": true, "capabilities": ["build_as_paths"]}'

EXPORT_METHODS = [(False, "Export"), (True, "FullExport")]


def _unwrap(tool_fn):
    fn = tool_fn
    while hasattr(fn, "__wrapped__"):
        fn = fn.__wrapped__
    return fn


@pytest.fixture(autouse=True)
def _diag_dir(monkeypatch, tmp_path_factory):
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(tmp_path_factory.mktemp("diag")))


class FakeAccess:
    """``Application.Run`` stand-in that records the add-in method each call reached.

    ``api`` and ``async_start`` are the raw returns of ``API`` and ``APIAsync``;
    an exception instance is raised instead. ``on_async`` runs after a
    successful ``APIAsync`` start, to post a terminal callback.
    ``capabilities`` answers a direct run of ``APICapabilities``, which is
    counted in ``probes`` rather than ``dispatched``; ``probe_names`` holds
    the procedure names run. ``arguments`` holds each dispatched method's
    arguments.
    """

    def __init__(self, api=None, async_start=ASYNC_STARTED, on_async=None, capabilities=CAPABLE):
        self.api = api
        self.async_start = async_start
        self.on_async = on_async
        self.capabilities = capabilities
        self.dispatched: list[tuple[str, str]] = []
        self.arguments: list[tuple] = []
        self.probes = 0
        self.probe_names: list[str] = []
        self.export_folder: object = ""
        self.folder_queries = 0

    def run(self, name, *args):
        entry = name.rsplit(".", 1)[1]
        if entry == "APICapabilities":
            self.probes += 1
            self.probe_names.append(name)
            if isinstance(self.capabilities, Exception):
                raise self.capabilities
            return self.capabilities
        method = args[1] if entry == "APIAsync" else args[0]
        if entry == "API" and method == "GetExportFolder":
            self.folder_queries += 1
            if isinstance(self.export_folder, Exception):
                raise self.export_folder
            return self.export_folder
        self.dispatched.append((entry, method))
        self.arguments.append(args[2:] if entry == "APIAsync" else args[1:])
        outcome = self.async_start if entry == "APIAsync" else self.api
        if isinstance(outcome, Exception):
            raise outcome
        if entry == "APIAsync" and self.on_async:
            self.on_async(args[0])
        return outcome


def _addin(tmp_path, fake):
    addin_file = tmp_path / "Version Control.accda"
    addin_file.touch()
    addin = VCSAddinIntegration(str(addin_file))
    addin._app = Mock()
    addin._app.Run = Mock(side_effect=fake.run)
    addin._addin_loaded = True
    return addin


def _complete(manager, **fields):
    def post(callback_info):
        operation_id = json.loads(callback_info)["operation_id"]
        assert manager.route_callback(
            operation_id, {"operation_id": operation_id, "type": "complete", **fields}
        )

    return post


def _conn():
    conn = MagicMock()
    conn.__enter__ = MagicMock(return_value=conn)
    conn.__exit__ = MagicMock(return_value=False)
    conn.connect.return_value = (MagicMock(), MagicMock())
    return conn


@contextmanager
def _export_env(tmp_path, addin, manager, callback):
    db = tmp_path / "test.accdb"
    db.touch()
    out = tmp_path / "export"
    out.mkdir(exist_ok=True)
    with (
        patch("msaccess_vcs_mcp.tools.AccessConnection", return_value=_conn()),
        patch("msaccess_vcs_mcp.tools.VCSAddinIntegration", return_value=addin),
        patch("msaccess_vcs_mcp.tools.validate_database_path", return_value=db),
        patch("msaccess_vcs_mcp.tools.validate_export_directory", return_value=out),
        patch("msaccess_vcs_mcp.tools._check_database_busy", return_value=None),
        patch(
            "msaccess_vcs_mcp.tools.get_callback_url",
            return_value="http://localhost:1/cb" if callback else None,
        ),
        patch("msaccess_vcs_mcp.tools._get_operation_manager", return_value=manager),
        patch(
            "msaccess_vcs_mcp.tools.get_config",
            return_value={"ACCESS_VCS_ADDIN_PATH": str(tmp_path / "Version Control.accda")},
        ),
    ):
        yield db, out


def _export(tmp_path, fake, *, full_export=False, callback=True, complete=False):
    from msaccess_vcs_mcp.tools import vcs_export_database

    manager = OperationManager()
    if complete:
        fake.on_async = _complete(manager)
    if fake.export_folder == "":
        fake.export_folder = str(tmp_path / "export") + os.sep
    addin = _addin(tmp_path, fake)
    with _export_env(tmp_path, addin, manager if callback else None, callback) as (db, out):
        result = asyncio.run(
            _unwrap(vcs_export_database)(str(db), str(out), full_export=full_export)
        )
    assert manager.pending_count() == 0
    return result, out


@contextmanager
def _build_env(tmp_path, addin, manager, callback, app=None):
    src = tmp_path / "src"
    src.mkdir(exist_ok=True)
    with (
        patch("msaccess_vcs_mcp.tools.check_write_permission", return_value=None),
        patch("msaccess_vcs_mcp.tools.validate_source_directory", return_value=src),
        patch("msaccess_vcs_mcp.tools.close_owned_instances_holding", return_value=[]),
        patch("msaccess_vcs_mcp.tools._check_database_busy", return_value=None),
        patch("msaccess_vcs_mcp.tools.create_isolated_access_app", return_value=app or MagicMock()),
        patch("msaccess_vcs_mcp.tools.prefer_full_power_if_created"),
        patch("msaccess_vcs_mcp.tools.ensure_access_visible"),
        patch("msaccess_vcs_mcp.tools.VCSAddinIntegration", return_value=addin),
        patch(
            "msaccess_vcs_mcp.tools.get_callback_url",
            return_value="http://localhost:1/cb" if callback else None,
        ),
        patch("msaccess_vcs_mcp.tools._get_operation_manager", return_value=manager),
        patch(
            "msaccess_vcs_mcp.tools.get_config",
            return_value={"ACCESS_VCS_ADDIN_PATH": str(tmp_path / "Version Control.accda")},
        ),
    ):
        yield src


def _build(tmp_path, fake, *, callback=True, complete=False, completion=None, output=None, app=None):
    from msaccess_vcs_mcp.tools import vcs_rebuild_database

    manager = OperationManager()
    if complete:
        fake.on_async = _complete(manager, **(completion or {}))
    addin = _addin(tmp_path, fake)
    output = str(tmp_path / "out.accdb") if output is None else output
    with _build_env(tmp_path, addin, manager if callback else None, callback, app) as src:
        result = asyncio.run(_unwrap(vcs_rebuild_database)(str(src), output))
    assert manager.pending_count() == 0
    return result, src, output


def _assert_unconfirmed(result):
    assert result["success"] is False
    assert result["started"] is True
    assert result["completion_unconfirmed"] is True
    assert "error_pattern" not in result
    assert "log_path" in result
    for guidance in ("may have succeeded or failed", "log_path", "vcs_get_recent_calls()"):
        assert guidance in result["error"]


# ---------------------------------------------------------------------------
# start_only_result: what a start-only return proves
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("raw", [None, "", "   ", "not json", "[1]", '{"success": true}',
                                 '{"started": true}', {"success": True}, 7])
def test_a_return_that_proves_nothing_is_unconfirmed(raw):
    result = start_only_result(raw, "export", "Export")
    assert result == unconfirmed_start_result("export", "Export")
    assert result["success"] is False
    assert "error_pattern" not in result


def test_dispatcher_refusal_fails_with_its_pattern():
    result = start_only_result(REFUSED, "export", "Export")
    assert result["success"] is False
    assert result["error_pattern"] == "operation_already_running"
    assert "completion_unconfirmed" not in result


def test_self_dispatch_refusal_keeps_its_own_pattern():
    refused = "VCS_API_REFUSED: the call arrived back in the project that sent it"
    assert start_only_result(refused, "build", "Build")["error_pattern"] == "api_self_dispatch"


def test_failure_object_keeps_pattern_and_decisions():
    raw = json.dumps({
        "success": False, "error_pattern": "decision_required",
        "decision_required": True, "decisions": [{"object": "a"}], "error": "needs a decision",
    })
    result = start_only_result(raw, "export", "Export")
    assert result["success"] is False
    assert result["error_pattern"] == "decision_required"
    assert result["decision_required"] is True
    assert result["decisions"] == [{"object": "a"}]
    assert "completion_unconfirmed" not in result


def test_undecided_prompt_beats_a_success_flag():
    result = start_only_result({"success": True, "decision_required": True}, "build", "Build")
    assert result["success"] is False
    assert result["error_pattern"] == "decision_required"


def test_failure_without_text_is_named_by_operation():
    assert start_only_result({"success": False}, "build", "Build")["error"] == "The build failed"


def test_refusal_wrapped_in_a_failure_envelope_is_recognised():
    result = start_only_result({"success": False, "error": REFUSED}, "export", "Export")
    assert result["error_pattern"] == "operation_already_running"
    assert result["api_refused"] is True


def test_guidance_names_the_operation_and_its_log_family():
    error = unconfirmed_start_result("build", "Build")["error"]
    assert error.startswith("The build started")
    assert 'vcs_get_log(log_type="Build")' in error


# ---------------------------------------------------------------------------
# Integration helpers never return success: true
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("helper, expected", [
    (lambda a, p: a.export_source(p), "Export"),
    (lambda a, p: a.export_source(p, full_export=True), "FullExport"),
    (lambda a, p: a.export_vba(p), "ExportVBA"),
])
def test_export_helpers_dispatch_the_requested_method(tmp_path, helper, expected):
    fake = FakeAccess()
    result = helper(_addin(tmp_path, fake), str(tmp_path / "test.accdb"))
    assert fake.dispatched == [("API", expected)]
    _assert_unconfirmed(result)


def test_build_helper_passes_source_and_output_and_is_unconfirmed(tmp_path):
    fake = FakeAccess()
    addin = _addin(tmp_path, fake)
    src, out = str(tmp_path / "src"), str(tmp_path / "out.accdb")
    result = addin.build_from_source(src, out)
    assert fake.dispatched == [("API", "BuildAs")]
    assert fake.arguments == [(src, out)]
    _assert_unconfirmed(result)
    # Nothing reported where it wrote: the request is kept apart.
    assert result["output_path"] is None
    assert result["requested_output_path"] == out


def test_build_helper_without_an_output_builds_the_source(tmp_path):
    fake = FakeAccess()
    src = str(tmp_path / "src")
    result = _addin(tmp_path, fake).build_from_source(src)
    assert fake.dispatched == [("API", "Build")]
    assert fake.arguments == [(src,)]
    _assert_unconfirmed(result)
    assert "requested_output_path" not in result


@pytest.mark.parametrize("capabilities", [
    RuntimeError("Microsoft Access cannot find the procedure 'APICapabilities.'"),
    None,
    "",
    "not json",
    '{"success": false}',
    '{"success": true}',
    '{"success": true, "capabilities": "build_as_paths"}',
    '{"success": true, "capabilities": ["something_else"]}',
], ids=["raises", "empty", "blank", "not-json", "failed", "no-list", "not-a-list", "missing-name"])
def test_build_as_paths_refusal_fails_closed(tmp_path, capabilities):
    refusal = _addin(tmp_path, FakeAccess(capabilities=capabilities)).build_as_paths_refusal()
    assert refusal["success"] is False
    assert refusal["error_pattern"] == "build_output_unsupported"
    assert "build_as_paths" in refusal["error"]


def test_build_as_paths_confirmed_is_no_refusal(tmp_path):
    fake = FakeAccess()
    assert _addin(tmp_path, fake).build_as_paths_refusal() is None
    # Run by name, never through API: API on a missing method leaves a
    # modal runtime-error dialog in Access instead of raising.
    lib = os.path.splitext(os.path.abspath(tmp_path / "Version Control.accda"))[0]
    assert fake.probe_names == [f"{lib}.APICapabilities"]
    assert fake.dispatched == []


def test_build_as_paths_reads_an_early_bound_tuple(tmp_path):
    fake = FakeAccess(capabilities=(CAPABLE,))
    assert _addin(tmp_path, fake).build_as_paths_refusal() is None


def test_build_as_paths_refusal_without_a_loaded_addin(tmp_path):
    fake = FakeAccess()
    addin = _addin(tmp_path, fake)
    addin._addin_loaded = False
    refusal = addin.build_as_paths_refusal()
    assert refusal["error_pattern"] == "build_output_unsupported"
    assert fake.probes == 0


@pytest.mark.parametrize("call", [
    lambda a, tp: a.export_source(str(tp / "t.accdb")),
    lambda a, tp: a.export_vba(str(tp / "t.accdb")),
    lambda a, tp: a.build_from_source(str(tp / "src")),
])
def test_a_com_exception_still_fails(tmp_path, call):
    result = call(_addin(tmp_path, FakeAccess(api=RuntimeError("boom"))), tmp_path)
    assert result["success"] is False
    assert "boom" in result["error"]
    assert "completion_unconfirmed" not in result


def test_a_refused_start_fails_in_the_helper(tmp_path):
    result = _addin(tmp_path, FakeAccess(api=REFUSED)).export_source(str(tmp_path / "t.accdb"))
    assert result["success"] is False
    assert result["error_pattern"] == "operation_already_running"


# ---------------------------------------------------------------------------
# vcs_export_database
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("full_export, method", EXPORT_METHODS)
def test_export_callback_completion_is_unchanged(tmp_path, full_export, method):
    fake = FakeAccess()
    result, out = _export(tmp_path, fake, full_export=full_export, complete=True)
    assert result["success"] is True
    assert result["export_path"] == str(out)
    assert "completion_unconfirmed" not in result
    assert fake.dispatched == [("APIAsync", method)]


@pytest.mark.parametrize("full_export, method", EXPORT_METHODS)
def test_export_inline_empty_result_is_unconfirmed(tmp_path, full_export, method):
    fake = FakeAccess(async_start='{"sync": true, "result": null}')
    result, out = _export(tmp_path, fake, full_export=full_export)
    _assert_unconfirmed(result)
    assert result["export_path"] == str(out)
    # The add-in already ran it inline: never run it twice.
    assert fake.dispatched == [("APIAsync", method)]


@pytest.mark.parametrize("full_export, method", EXPORT_METHODS)
def test_export_inline_refusal_fails_with_its_pattern(tmp_path, full_export, method):
    fake = FakeAccess(async_start=json.dumps({"sync": True, "result": REFUSED}))
    result, out = _export(tmp_path, fake, full_export=full_export)
    assert result["success"] is False
    assert result["error_pattern"] == "operation_already_running"
    assert "completion_unconfirmed" not in result
    assert result["exported_count"] == 0
    assert result["export_path"] == str(out)
    assert fake.dispatched == [("APIAsync", method)]


def test_export_inline_failure_object_keeps_decision_fields(tmp_path):
    nested = {
        "success": False, "error_pattern": "decision_required",
        "decision_required": True, "decisions": [{"object": "a"}],
    }
    fake = FakeAccess(async_start=json.dumps({"sync": True, "result": json.dumps(nested)}))
    result, _ = _export(tmp_path, fake, full_export=True)
    assert result["success"] is False
    assert result["error_pattern"] == "decision_required"
    assert result["decisions"] == [{"object": "a"}]


@pytest.mark.parametrize("full_export, method", EXPORT_METHODS)
@pytest.mark.parametrize("start", [
    '{"unexpected": true}',
    RuntimeError("async start failed"),
], ids=["neither-marker", "async-raises"])
def test_export_sync_fallback_keeps_the_requested_method(tmp_path, full_export, method, start):
    fake = FakeAccess(async_start=start)
    result, out = _export(tmp_path, fake, full_export=full_export)
    _assert_unconfirmed(result)
    assert result["export_path"] == str(out)
    assert fake.dispatched == [("APIAsync", method), ("API", method)]


@pytest.mark.parametrize("full_export, method", EXPORT_METHODS)
def test_export_without_a_callback_server_is_unconfirmed(tmp_path, full_export, method):
    fake = FakeAccess()
    result, _ = _export(tmp_path, fake, full_export=full_export, callback=False)
    _assert_unconfirmed(result)
    assert fake.dispatched == [("API", method)]


def test_export_sync_fallback_refusal_and_exception_fail(tmp_path):
    refused, _ = _export(tmp_path, FakeAccess(api=REFUSED), callback=False)
    assert refused["success"] is False
    assert refused["error_pattern"] == "operation_already_running"
    assert refused["exported_count"] == 0

    raised, _ = _export(tmp_path, FakeAccess(api=RuntimeError("boom")), full_export=True, callback=False)
    assert raised["success"] is False
    assert "boom" in raised["error"]
    assert "completion_unconfirmed" not in raised


# ---------------------------------------------------------------------------
# vcs_rebuild_database
# ---------------------------------------------------------------------------


def test_build_callback_reports_the_add_ins_output_path(tmp_path):
    reported = str(tmp_path / "reported.accdb")
    fake = FakeAccess()
    result, src, output = _build(tmp_path, fake, complete=True, completion={"output_path": reported})
    assert result["success"] is True
    assert result["output_path"] == reported
    assert result["source_dir"] == str(src)
    assert "completion_unconfirmed" not in result
    assert fake.dispatched == [("APIAsync", "BuildAs")]
    assert fake.arguments == [(str(src), output)]
    assert fake.probes == 1


def test_build_completion_without_a_path_does_not_echo_the_request(tmp_path):
    result, _, _ = _build(tmp_path, FakeAccess(), complete=True)
    assert result["success"] is True
    assert result["output_path"] is None


@pytest.mark.parametrize("callback", [True, False], ids=["callback", "no-callback"])
def test_build_refused_before_start_without_the_capability(tmp_path, callback):
    fake = FakeAccess(capabilities=RuntimeError("Method not found"))
    result, src, _ = _build(tmp_path, fake, callback=callback)
    assert result["success"] is False
    assert result["error_pattern"] == "build_output_unsupported"
    assert result["output_path"] is None
    assert result["source_dir"] == str(src)
    assert fake.dispatched == []


def test_build_relative_output_is_refused_before_access_starts(tmp_path):
    app = MagicMock()
    fake = FakeAccess()
    result, _, _ = _build(tmp_path, fake, output="out.accdb", app=app)
    assert result["success"] is False
    assert result["error_pattern"] == "invalid_build_path"
    assert result["output_path"] is None
    assert fake.probes == 0
    assert fake.dispatched == []
    app.NewCurrentDatabase.assert_not_called()


def test_build_host_database_is_closed_and_removed(tmp_path):
    app = MagicMock()
    _build(tmp_path, FakeAccess(), complete=True, app=app)
    host_path = app.NewCurrentDatabase.call_args.args[0]
    assert os.path.basename(host_path) == "BuildHost.accdb"
    app.CloseCurrentDatabase.assert_called_once_with()
    app.Quit.assert_called_once_with()
    assert not os.path.exists(os.path.dirname(host_path))


def test_build_host_folder_removal_is_retried(tmp_path, monkeypatch):
    from msaccess_vcs_mcp import tools

    host = tmp_path / "host"
    host.mkdir()
    calls = []
    real_rmtree = tools.shutil.rmtree

    def held_once(path, ignore_errors=False):
        calls.append(path)
        if len(calls) > 1:
            real_rmtree(path, ignore_errors=ignore_errors)

    monkeypatch.setattr(tools.shutil, "rmtree", held_once)
    tools._close_build_host(MagicMock(), str(host), delay=0)
    assert len(calls) == 2
    assert not host.exists()


def test_build_inline_empty_result_is_unconfirmed_and_not_run_twice(tmp_path):
    fake = FakeAccess(async_start='{"sync": true, "result": null}')
    result, src, output = _build(tmp_path, fake)
    _assert_unconfirmed(result)
    assert result["output_path"] is None
    assert result["requested_output_path"] == output
    assert result["source_dir"] == str(src)
    assert [entry for entry, _ in fake.dispatched] == ["APIAsync"]


def test_build_inline_refusal_fails_with_its_pattern(tmp_path):
    fake = FakeAccess(async_start=json.dumps({"sync": True, "result": REFUSED}))
    result, _, _ = _build(tmp_path, fake)
    assert result["success"] is False
    assert result["error_pattern"] == "operation_already_running"
    assert result["output_path"] is None
    assert "completion_unconfirmed" not in result
    assert [entry for entry, _ in fake.dispatched] == ["APIAsync"]


def test_build_refused_async_start_is_not_dispatched_again(tmp_path):
    fake = FakeAccess(async_start=json.dumps({"success": False, "error": "start refused"}))
    result, _, _ = _build(tmp_path, fake)
    assert result["success"] is False
    assert result["error"] == "start refused"
    assert result["output_path"] is None
    assert "completion_unconfirmed" not in result
    assert fake.dispatched == [("APIAsync", "BuildAs")]


@pytest.mark.parametrize("start", [
    '{"unexpected": true}',
    RuntimeError("async start failed"),
], ids=["neither-marker", "async-raises"])
def test_build_sync_fallback_uses_build_as_and_is_unconfirmed(tmp_path, start):
    fake = FakeAccess(async_start=start)
    result, src, output = _build(tmp_path, fake)
    _assert_unconfirmed(result)
    assert result["output_path"] is None
    assert result["requested_output_path"] == output
    assert fake.dispatched == [("APIAsync", "BuildAs"), ("API", "BuildAs")]
    assert fake.arguments == [(str(src), output), (str(src), output)]


def test_build_without_a_callback_server_is_unconfirmed(tmp_path):
    fake = FakeAccess()
    result, src, output = _build(tmp_path, fake, callback=False)
    _assert_unconfirmed(result)
    assert fake.dispatched == [("API", "BuildAs")]
    assert fake.arguments == [(str(src), output)]


def test_build_sync_fallback_refusal_and_exception_fail(tmp_path):
    refused, _, _ = _build(tmp_path, FakeAccess(api=REFUSED), callback=False)
    assert refused["success"] is False
    assert refused["error_pattern"] == "operation_already_running"
    assert refused["output_path"] is None

    raised, _, _ = _build(tmp_path, FakeAccess(api=RuntimeError("boom")), callback=False)
    assert raised["success"] is False
    assert "boom" in raised["error"]
    assert "completion_unconfirmed" not in raised


# --- X10: the export destination is the add-in's configured folder ---------------


def _export_to(tmp_path, fake, output_dir, *, object_types=None, callback=True):
    from msaccess_vcs_mcp.tools import vcs_export_database

    manager = OperationManager()
    fake.on_async = _complete(manager)
    addin = _addin(tmp_path, fake)
    with _export_env(tmp_path, addin, manager if callback else None, callback) as (db, out):
        kwargs = {"object_types": object_types} if object_types else {}
        args = (str(db),) if output_dir is None else (str(db), str(output_dir))
        result = asyncio.run(_unwrap(vcs_export_database)(*args, **kwargs))
    return result, out


@pytest.mark.parametrize("object_types", [None, ["queries"]])
def test_export_to_another_folder_is_refused_before_anything_is_exported(tmp_path, object_types):
    fake = FakeAccess()
    fake.export_folder = str(tmp_path / "configured") + os.sep
    elsewhere = tmp_path / "elsewhere"
    result, _ = _export_to(tmp_path, fake, elsewhere, object_types=object_types)
    assert result["success"] is False
    assert result["error_pattern"] == "export_folder_mismatch"
    assert result["export_path"] is None
    assert result["exported_count"] == 0
    assert result["configured_export_folder"] == fake.export_folder
    assert result["requested_output_dir"] == str(elsewhere)
    assert fake.dispatched == []
    assert not elsewhere.exists()


@pytest.mark.parametrize("object_types", [None, ["queries"]])
@pytest.mark.parametrize("given", ["same", "omitted", "case_and_slash"])
def test_export_to_the_configured_folder_reports_it_as_export_path(tmp_path, object_types, given):
    fake = FakeAccess(api='{"success": true}')
    configured = tmp_path / "export"
    fake.export_folder = str(configured) + os.sep
    output_dir = {
        "same": configured,
        "omitted": None,
        "case_and_slash": str(configured).upper() + os.sep,
    }[given]
    result, out = _export_to(tmp_path, fake, output_dir, object_types=object_types)
    assert result["success"] is True
    assert result["export_path"] == str(out)
    assert fake.dispatched[0][1] == ("ExportByType" if object_types else "Export")


@pytest.mark.parametrize("folder", ["", "{}", REFUSED, RuntimeError("boom")])
def test_export_with_no_readable_folder_is_refused(tmp_path, folder):
    fake = FakeAccess()
    fake.export_folder = folder
    result, _ = _export_to(tmp_path, fake, None)
    assert result["success"] is False
    assert result["error_pattern"] == "export_folder_unavailable"
    assert result["export_path"] is None
    assert fake.dispatched == []
