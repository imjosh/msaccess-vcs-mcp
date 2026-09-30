"""Consume the real A24 contract from an isolated, test-owned Access instance."""

import json
import sys
from pathlib import Path
from unittest.mock import patch

import pytest

from msaccess_vcs_mcp.addin_integration import VCSAddinIntegration
from msaccess_vcs_mcp.config import get_default_addin_path
from msaccess_vcs_mcp.decision_policy import call_under_policy, select_interactive_mode
from msaccess_vcs_mcp.usage_logging import reset_logging
from tests.interaction_mode_contract import INTERACTIVE_CONFIRMED, INTERACTIVE_REFUSED


@pytest.mark.integration
@pytest.mark.skipif(sys.platform != "win32", reason="Requires Microsoft Access")
def test_real_confirmation_refusal_and_failed_cleanup(tmp_path, monkeypatch):
    import pythoncom
    import win32com.client
    from msaccess_vcs_mcp.access_com.connection import ensure_access_visible

    monkeypatch.setenv("ACCESS_VCS_LOG_DIR", str(tmp_path / "usage"))
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(tmp_path / "diag"))
    reset_logging()
    pythoncom.CoInitialize()
    app = None
    addin = None
    loaded = False
    try:
        # DispatchEx creates our own instance; never attach or close user Access.
        app = win32com.client.DispatchEx("Access.Application")
        db_path = str(tmp_path / "interactive-contract.accdb")
        app.NewCurrentDatabase(db_path)
        ensure_access_visible(app)
        app.CurrentDb().CreateQueryDef("qryInteractiveContract", "SELECT 1 AS ResultValue")
        addin = VCSAddinIntegration(get_default_addin_path())
        # NewCurrentDatabase has no ROT file moniker. Probe in this apartment;
        # this test concerns the VBA contract, not worker-thread attachment.
        api = str(Path(addin.addin_path).with_suffix("")) + ".API"
        with patch.object(addin, "_probe_with_timeout", side_effect=lambda *_: app.Run(api, "GetVCSVersion")):
            addin.load_addin(app)
        loaded = True

        confirmed = addin.call_sync("SetInteractionMode", 0)
        print("confirmed:", confirmed)
        assert json.loads(confirmed) == INTERACTIVE_CONFIRMED
        assert select_interactive_mode(addin, None) is None

        # Simulate transport failure of cleanup, retaining the actual VBA scope.
        real_call = addin.call_sync
        events = []

        def call(command, *args):
            events.append(command)
            if command == "ClearOperationPolicy":
                raise RuntimeError("cleanup transport failed")
            return real_call(command, *args)

        with patch.object(addin, "call_sync", side_effect=call):
            result, state = call_under_policy(
                addin, "block", "ExportObject", "query", "qryInteractiveContract", parse_result=json.loads,
            )
            assert state == "completed"
            assert result["success"] is True
            assert result["policy_cleanup_error"] == "cleanup transport failed"
            events.clear()
            refused, state = call_under_policy(
                addin, None, "ExportObject", "query", "qryInteractiveContract", parse_result=json.loads,
            )
            print("refused:", json.dumps(refused))
            assert state == "refused"
            assert refused == INTERACTIVE_REFUSED
            assert events == ["SetInteractionMode"]

        # The failed interactive selection left the actual caller-owned scope open.
        assert json.loads(real_call("SetInteractionMode", 0)) == INTERACTIVE_REFUSED
        real_call("ClearOperationPolicy")
        assert json.loads(real_call("SetInteractionMode", 0)) == INTERACTIVE_CONFIRMED

        # pywin32 exposes VBA Empty as None, the older add-in's return value.
        empty = win32com.client.VARIANT(pythoncom.VT_EMPTY, None).value
        assert empty is None
        with patch.object(addin, "call_sync", return_value=empty):
            assert select_interactive_mode(addin, None)["error_pattern"] == "interaction_mode_unconfirmed"
    finally:
        if app is not None:
            try:
                if loaded:
                    addin.call_sync("ClearOperationPolicy")
            finally:
                app.CloseCurrentDatabase()
                app.Quit(2)
        pythoncom.CoUninitialize()
        reset_logging()
