"""Consume the real A24 contract from an isolated, test-owned Access instance."""

import gc
import json
import sys
from pathlib import Path
from unittest.mock import patch

import pytest

from msaccess_vcs_mcp.addin_integration import VCSAddinIntegration
from msaccess_vcs_mcp.compatibility_session import ensure_session
from tests.policy_live_host import OwnedHost
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
    monkeypatch.setenv("ACCESS_VCS_OWNED_INSTANCES_PATH", str(tmp_path / "owned.json"))
    reset_logging()
    pythoncom.CoInitialize()
    app = None
    addin = None
    host = None
    try:
        # DispatchEx creates our own instance; never attach or close user Access.
        app = win32com.client.DispatchEx("Access.Application")
        db_path = str(tmp_path / "interactive-contract.accdb")
        host = OwnedHost(app, db_path)
        app.NewCurrentDatabase(db_path)
        ensure_access_visible(app)
        app.CurrentDb().CreateQueryDef("qryInteractiveContract", "SELECT 1 AS ResultValue")
        addin = VCSAddinIntegration(get_default_addin_path())
        # NewCurrentDatabase has no ROT file moniker. Probe in this apartment;
        # this test concerns the VBA contract, not worker-thread attachment.
        def apartment_probe(probe_app, db_path, timeout):
            # Use the production named read-only handshake in this owning STA.
            # Its real accepted return is forwarded exactly as the worker does.
            # A returned compatibility refusal raises and is never bypassed.
            assert probe_app is app and db_path is None
            addin._probe_session = ensure_session(probe_app, addin.addin_path)

        with patch.object(addin, "_probe_with_timeout", side_effect=apartment_probe):
            addin.load_addin(app)
        host.addin = addin
        library = str(Path(addin.addin_path).with_suffix(""))
        identity = app.Run(library + ".APIIdentity")
        host.receipt["loaded_identity"] = json.loads(identity[0] if isinstance(identity, tuple) else identity)
        host.receipt["session"] = json.loads(ensure_session(app, addin.addin_path))

        confirmed = addin.call_sync("SetInteractionMode", 0)
        print("confirmed:", confirmed)
        host.check("confirmed", json.loads(confirmed), INTERACTIVE_CONFIRMED)
        assert select_interactive_mode(addin, None) is None

        # Simulate transport failure of cleanup, retaining the actual VBA scope.
        real_call = addin.call_sync
        events = []
        original = {}

        def call(command, *args):
            events.append(command)
            if command == "ClearOperationPolicy":
                raise RuntimeError("cleanup transport failed")
            raw = real_call(command, *args)
            if command == "ExportObject":
                original.update(json.loads(raw))
            return raw

        with patch.object(addin, "call_sync", side_effect=call):
            result, state = call_under_policy(
                addin, "block", "ExportObject", "query", "qryInteractiveContract", parse_result=json.loads,
            )
            assert state == "completed"
            assert result["success"] is True
            assert result["policy_cleanup_error"] == "cleanup transport failed"
            host.check("failed_cleanup_outcome", result, {**original, "policy_cleanup_error": "cleanup transport failed"}, native=False)
            host.check("failed_cleanup_commands", events.copy(), ["SetOperationPolicy", "ExportObject", "ClearOperationPolicy"], native=False)
            events.clear()
            refused, state = call_under_policy(
                addin, None, "ExportObject", "query", "qryInteractiveContract", parse_result=json.loads,
            )
            print("refused:", json.dumps(refused))
            assert state == "refused"
            host.check("refusal", refused, INTERACTIVE_REFUSED)
            host.check("refusal_commands", events.copy(), ["SetInteractionMode"])

        # The failed interactive selection left the actual caller-owned scope open.
        host.check("owner_still_active", json.loads(real_call("SetInteractionMode", 0)), INTERACTIVE_REFUSED)
        host.check("owner_cleanup", json.loads(real_call("ClearOperationPolicy")), {"success": True})
        host.check("owner_cleanup_mode", json.loads(real_call("SetInteractionMode", 0)), INTERACTIVE_CONFIRMED)

        with patch.object(addin, "call_sync", wraps=real_call) as calls:
            clean, state = call_under_policy(addin, None, "ExportObject", "query", "qryInteractiveContract", parse_result=json.loads)
        host.check("clean_interactive_state", state, "completed")
        host.check("clean_interactive_success", clean["success"], True)
        host.check("clean_interactive_cleanup_error", "policy_cleanup_error" in clean, False)
        host.check("clean_interactive_commands", [c.args[0] for c in calls.call_args_list], ["SetInteractionMode", "ExportObject"])

        # pywin32 exposes VBA Empty as None, the older add-in's return value.
        empty = win32com.client.VARIANT(pythoncom.VT_EMPTY, None).value
        assert empty is None
        with patch.object(addin, "call_sync", return_value=empty):
            host.check("controlled_empty", select_interactive_mode(addin, None)["error_pattern"], "interaction_mode_unconfirmed", native=False)
    finally:
        try:
            if host is not None:
                def release():
                    nonlocal app
                    if addin is not None:
                        addin._app = None
                    host.app = None
                    app = None
                    gc.collect()
                host.close(release)
        finally:
            if host is not None:
                host.save("confirmation")
            if addin is not None:
                addin._app = None  # release the apartment proxy before CoUninitialize
            if host is not None:
                host.app = None
            app = None
            gc.collect()
            pythoncom.CoUninitialize()
            reset_logging()


@pytest.mark.integration
@pytest.mark.skipif(sys.platform != "win32", reason="Requires Microsoft Access")
def test_real_policy_set_and_clear_acknowledgments(tmp_path, monkeypatch):
    """The real add-in's SetOperationPolicy/ClearOperationPolicy JSON passes the M40 checks."""
    import pythoncom
    import win32com.client
    from msaccess_vcs_mcp.access_com.connection import ensure_access_visible

    monkeypatch.setenv("ACCESS_VCS_LOG_DIR", str(tmp_path / "usage"))
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(tmp_path / "diag"))
    monkeypatch.setenv("ACCESS_VCS_OWNED_INSTANCES_PATH", str(tmp_path / "owned.json"))
    reset_logging()
    pythoncom.CoInitialize()
    app = None
    addin = None
    host = None
    try:
        app = win32com.client.DispatchEx("Access.Application")
        host = OwnedHost(app, tmp_path / "policy-contract.accdb")
        app.NewCurrentDatabase(str(host.path))
        ensure_access_visible(app)
        app.CurrentDb().CreateQueryDef("qryPolicyContract", "SELECT 1 AS ResultValue")
        addin = VCSAddinIntegration(get_default_addin_path())
        def apartment_probe(probe_app, db_path, timeout):
            # Use the production named read-only handshake in this owning STA.
            # Its real accepted return is forwarded exactly as the worker does.
            # A returned compatibility refusal raises and is never bypassed.
            assert probe_app is app and db_path is None
            addin._probe_session = ensure_session(probe_app, addin.addin_path)

        with patch.object(addin, "_probe_with_timeout", side_effect=apartment_probe):
            addin.load_addin(app)
        host.addin = addin
        library = str(Path(addin.addin_path).with_suffix(""))
        identity = app.Run(library + ".APIIdentity")
        host.receipt["loaded_identity"] = json.loads(identity[0] if isinstance(identity, tuple) else identity)
        host.receipt["session"] = json.loads(ensure_session(app, addin.addin_path))

        host.check("set_ack", json.loads(addin.call_sync("SetOperationPolicy", "Block")), {"success": True, "policy": "block"})
        host.check("clear_ack", json.loads(addin.call_sync("ClearOperationPolicy")), {"success": True})
        host.check("idempotent_clear", json.loads(addin.call_sync("ClearOperationPolicy")), {"success": True})

        with patch.object(addin, "call_sync", wraps=addin.call_sync) as calls:
            result, state = call_under_policy(
                addin, "block", "ExportObject", "query", "qryPolicyContract", parse_result=json.loads,
            )
        host.check("policy_commands", [c.args[0] for c in calls.call_args_list], ["SetOperationPolicy", "ExportObject", "ClearOperationPolicy"])
        print("result:", json.dumps(result))
        host.check("policy_state", state, "completed")
        host.check("policy_success", result["success"], True)
        host.check("policy_cleanup_error", "policy_cleanup_error" in result, False)
        host.check("confirmed_after_policy", json.loads(addin.call_sync("SetInteractionMode", 0)), INTERACTIVE_CONFIRMED)

        with patch.object(addin, "call_sync", wraps=addin.call_sync) as calls:
            invalid, state = call_under_policy(
                addin, "bogus", "ExportObject", "query", "qryPolicyContract", parse_result=json.loads,
            )
        host.check("invalid_commands", [c.args[0] for c in calls.call_args_list], ["SetOperationPolicy"])
        host.check("invalid_state", state, "refused")
        host.check("invalid_pattern", invalid["error_pattern"], "invalid_decision_policy")
    finally:
        try:
            if host is not None:
                def release():
                    nonlocal app
                    if addin is not None:
                        addin._app = None
                    host.app = None
                    app = None
                    gc.collect()
                host.close(release)
        finally:
            if host is not None:
                host.save("acknowledgments")
            if addin is not None:
                addin._app = None  # release the apartment proxy before CoUninitialize
            if host is not None:
                host.app = None
            app = None
            gc.collect()
            pythoncom.CoUninitialize()
            reset_logging()
