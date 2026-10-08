"""M43 live rebuild beside a user-controlled, disposable Access database."""
import asyncio
import json
import sys
from pathlib import Path
from unittest.mock import patch

import pytest

from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.callback_server import CallbackServer
from msaccess_vcs_mcp.config import get_default_addin_path
from msaccess_vcs_mcp.operation_manager import OperationManager
from tests.test_export_build_fallbacks import _unwrap
from tests.policy_live_host import OwnedHost, raw_creation


@pytest.mark.integration
@pytest.mark.skipif(sys.platform != "win32", reason="Requires Microsoft Access")
def test_rebuild_preserves_other_access_database(tmp_path, monkeypatch):
    import pythoncom
    import win32api
    import win32event
    from msaccess_vcs_mcp.access_com.connection import create_isolated_access_app, ensure_access_visible
    from msaccess_vcs_mcp.access_com.instance_registry import process_create_time
    from msaccess_vcs_mcp.access_com.process_qos import pid_from_access_app

    # Keep all diagnostics and registry claims local to this disposable run.
    monkeypatch.setenv("ACCESS_VCS_OWNED_INSTANCES_PATH", str(tmp_path / "owned.json"))
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(tmp_path / "diag"))
    source = tmp_path / "source"
    source.mkdir()
    (source / "project.json").write_text(json.dumps({
        "Info": {"Class": "clsDbProject", "Description": "Project"},
        "Items": {"FileFormat": 12, "RemovePersonalInformation": False},
    }))
    (source / "vcs-options.json").write_text(json.dumps({"Options": {"ExportFolder": ""}}))
    (source / "vbe-project.json").write_text(json.dumps({
        "Info": {"Class": "clsDbVbeProject", "Description": "VBE Project"},
        "Items": {"Name": "M43Fixture", "FileName": "fixture.accdb"},
    }))
    modules = source / "modules"
    modules.mkdir()
    (modules / "modM43Fixture.bas").write_text(
        'Attribute VB_Name = "modM43Fixture"\nOption Explicit\n'
        'Public Function M43Value() As Long\n    M43Value = 43\nEnd Function\n'
    )
    user_path = str(tmp_path / "user-session.accdb")
    output = str(tmp_path / "rebuilt.accdb")
    manager = OperationManager()
    server = CallbackServer(manager.route_callback)
    app = None
    owned = None
    build_handles = []
    error_trapping = None
    created = []
    pythoncom.CoInitialize()
    try:
        # This fixture owns this window; the rebuild under test does not.
        app = create_isolated_access_app()
        owned = OwnedHost(app, user_path)
        error_trapping = app.GetOption("Error Trapping")
        app.SetOption("Error Trapping", 2)
        app.NewCurrentDatabase(user_path)
        ensure_access_visible(app)
        app.UserControl = True
        app.CurrentDb().CreateQueryDef("qryM43Sentinel", "SELECT 43 AS SentinelValue")
        pid = pid_from_access_app(app)
        stamp = process_create_time(pid)
        assert pid and stamp
        print("M43 fixture pid:", pid, flush=True)
        server.start()

        def create_host():
            host = create_isolated_access_app()
            created.append(pid_from_access_app(host))
            handle = win32api.OpenProcess(0x1000 | 0x100000, False, created[-1])
            build_handles.append((created[-1], handle, raw_creation(handle)))
            print("M43 build pid:", created[-1], flush=True)
            return host

        with (
            patch.object(tools, "create_isolated_access_app", side_effect=create_host),
            patch.object(tools, "get_config", return_value={"ACCESS_VCS_ADDIN_PATH": get_default_addin_path()}),
            patch.object(tools, "_get_operation_manager", return_value=manager),
            patch.object(tools, "get_callback_url", return_value=server.callback_url),
        ):
            result = asyncio.run(_unwrap(tools.vcs_rebuild_database)(str(source), output))
        print("M43 rebuild:", json.dumps(result))
        assert result["success"] is True, result
        assert Path(result["output_path"]).resolve() == Path(output).resolve()
        assert Path(output).is_file()
        assert len(created) == 1 and created[0] and created[0] != pid
        assert pid_from_access_app(app) == pid
        assert process_create_time(pid) == stamp
        assert Path(app.CurrentDb().Name).resolve() == Path(user_path).resolve()
        assert app.UserControl is True
        assert app.CurrentDb().QueryDefs("qryM43Sentinel").SQL == "SELECT 43 AS SentinelValue;\r\n"
        owned.receipt['build_processes'] = []
        for build_pid, handle, creation in build_handles:
            exited = win32event.WaitForSingleObject(handle, 10000) == win32event.WAIT_OBJECT_0
            owned.receipt['build_processes'].append(dict(pid=build_pid, creation_FILETIME=creation,
                                                        original_handle_exit_confirmed=exited))
            assert exited
        owned.check('sentinel_preserved', owned.verified(), True)
        # Inspect the built file through DAO without opening another Access UI.
        built = app.DBEngine.OpenDatabase(output, False, True)
        try:
            assert built.Containers("Modules").Documents("modM43Fixture").Name == "modM43Fixture"
        finally:
            built.Close()
            built = None
    finally:
        server.stop()
        try:
            if owned is not None:
                if not owned.exited():
                    assert owned.verified(), 'Uncertain sentinel identity/path; left untouched'
                    if error_trapping is not None:
                        app.SetOption("Error Trapping", error_trapping)
                def release():
                    nonlocal app
                    owned.app = None
                    app = None
                owned.close(release)
        finally:
            if owned is not None:
                owned.save('build-preservation')
            for _, handle, _ in build_handles:
                win32api.CloseHandle(handle)
            pythoncom.CoUninitialize()
