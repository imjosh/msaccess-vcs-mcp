"""M43: a rebuild owns only the isolated instance it starts."""
import asyncio
from unittest.mock import MagicMock, patch

import pytest

from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.access_com import connection
from tests.test_export_build_fallbacks import _unwrap


@pytest.mark.parametrize("failure", [None, "dispatch", "power", "visible", "host", "load", "build"])
def test_rebuild_never_touches_existing_access(tmp_path, failure):
    existing = MagicMock()
    existing.CurrentDb.return_value.Name = "users-unsaved.accdb"
    existing.UserControl = True
    owned = MagicMock()
    raw = MagicMock()
    addin = MagicMock()
    addin.build_as_paths_refusal.return_value = None
    addin.build_from_source.return_value = {"success": False, "error": "build refused"}
    src = tmp_path / "src"
    src.mkdir()
    def stage(name):
        if failure == name:
            raise RuntimeError(name)
    owned.NewCurrentDatabase.side_effect = lambda *_: stage("host")
    addin.load_addin.side_effect = lambda *_, **kw: stage("load")
    if failure == "build":
        addin.build_from_source.side_effect = RuntimeError("build")
    with (
        patch.object(connection.pythoncom, "CoCreateInstanceEx",
              side_effect=RuntimeError("dispatch") if failure == "dispatch" else None,
              return_value=(raw,)) as isolated,
        patch.object(connection.win32com.client, "Dispatch", return_value=owned),
        patch("msaccess_vcs_mcp.tools.ensure_dispatch", return_value=existing, create=True),
        patch.object(tools, "get_config", return_value={}),
        patch.object(tools, "check_write_permission"),
        patch.object(tools, "validate_source_directory", return_value=src),
        patch.object(tools, "close_owned_instances_holding") as close_existing,
        patch.object(tools, "_check_database_busy", return_value=None),
        patch.object(tools, "get_callback_url", return_value=None),
        patch.object(tools, "_get_operation_manager", return_value=None),
        patch.object(tools, "prefer_full_power_if_created", side_effect=lambda *_: stage("power")),
        patch.object(tools, "ensure_access_visible", side_effect=lambda *_: stage("visible")),
        patch.object(tools, "VCSAddinIntegration", return_value=addin),
        patch.object(tools, "_close_build_host", wraps=tools._close_build_host) as close,
    ):
        result = asyncio.run(_unwrap(tools.vcs_rebuild_database)(str(src), str(tmp_path / "out.accdb")))
    assert result["success"] is False
    close_existing.assert_not_called()
    existing.NewCurrentDatabase.assert_not_called()
    existing.CloseCurrentDatabase.assert_not_called()
    existing.Quit.assert_not_called()
    assert existing.mock_calls == []
    assert existing.UserControl is True
    isolated.assert_called_once_with(
        "Access.Application", None, connection.pythoncom.CLSCTX_SERVER,
        None, (connection.pythoncom.IID_IDispatch,),
    )
    if failure == "dispatch":
        close.assert_not_called()
    else:
        assert close.call_count == 1
        assert close.call_args.args[0] is owned
        owned.CloseCurrentDatabase.assert_called_once()
        owned.Quit.assert_called_once()


@pytest.mark.parametrize("probe", ["installation", "version", "version_fallback"])
def test_standalone_probes_do_not_quit_existing_windows(tmp_path, probe):
    from msaccess_vcs_mcp import config, validation
    from msaccess_vcs_mcp.access_com import connection

    existing = MagicMock()
    existing.CurrentDb.return_value = None  # Even an empty shell is user-owned.
    owned = MagicMock()
    raw = MagicMock()
    addin = MagicMock()
    addin.verify_addin_exists.return_value = True
    target = tmp_path / "probe.accdb"
    target.touch()
    with (
        patch.object(connection.pythoncom, "CoCreateInstanceEx", return_value=(raw,)) as isolated,
        patch.object(connection.win32com.client, "Dispatch", return_value=owned),
        patch.object(connection, "ensure_dispatch", return_value=existing),
        patch.object(validation.win32com.client, "GetObject", side_effect=RuntimeError("not open")),
        patch.object(validation, "get_config", return_value={
            "ACCESS_VCS_DATABASE": str(target) if probe == "version_fallback" else None,
        }),
        patch.object(validation, "get_access_info", return_value={}),
        patch.object(validation, "VCSAddinIntegration", return_value=addin),
    ):
        if probe == "installation":
            config.validate_access_installation()
        else:
            assert validation.validate_components(load_addin=False)["success"] is True
    isolated.assert_called_once_with(
        "Access.Application", None, connection.pythoncom.CLSCTX_SERVER,
        None, (connection.pythoncom.IID_IDispatch,),
    )
    owned.Quit.assert_called_once()
    existing.Quit.assert_not_called()
    existing.CloseCurrentDatabase.assert_not_called()
