"""Shared test fixtures."""

from __future__ import annotations

import json

import pytest

from msaccess_vcs_mcp import exempt_workers


@pytest.fixture(autouse=True)
def _mock_existing_unit_installation(request, monkeypatch):
    """Existing unit seams use fake Access/add-ins, never the user's installation.

    X16 boundary tests opt out and exercise real admission with controlled DAO
    replies. Integration tests always inspect the real installation.
    """
    if request.node.get_closest_marker("integration") or request.node.get_closest_marker("version_gate"):
        return
    from msaccess_vcs_mcp import compatibility, tools
    def accepted():
        return {**compatibility.compatibility_result("5.2.0"), "addin_path": "mock-install.accda"}
    monkeypatch.setattr(tools, "inspect_installed_addin", accepted)


@pytest.fixture(params=[
    "raw-busy", "raw-self-dispatch", "wrapped-busy", "wrapped-self-dispatch",
    "busy-with-decisions", "decision-required",
])
def addin_start_refusal(request):
    """Fresh wire response and expected fields for M48, M49 and M51.

    Expected fields are independent of the production refusal parser. Consumers
    can use the same wire response on API or APIAsync without sharing a fake
    dispatcher or mutating another test's journal.
    """
    case = request.param
    self_dispatch = "self-dispatch" in case
    error = (
        "The call arrived back in the project that sent it."
        if self_dispatch else "Another API command is still running."
    )
    expected = {
        "success": False,
        "error": error,
        "error_pattern": "api_self_dispatch" if self_dispatch else "operation_already_running",
        "api_refused": True,
    }
    marked = "VCS_API_REFUSED: " + error
    if case.startswith("raw-"):
        raw = marked
    elif case.startswith("wrapped-"):
        raw = json.dumps({"success": False, "error": marked})
    else:
        expected["decisions"] = [{
            "kind": "confirmation", "object": "Form1", "resolution": "blocked",
        }]
        if case == "decision-required":
            expected.pop("api_refused")
            expected.update(
                error="A required decision was not covered by the decision policy.",
                error_pattern="decision_required", decision_required=True,
            )
        raw = json.dumps({**expected, "error": marked} if case == "busy-with-decisions" else expected)
    return {"raw": raw, "expected": expected}


@pytest.fixture(autouse=True)
def _drain_exempt_workers(monkeypatch):
    """Let released gate-exempt workers return before a test's patches are undone.

    Those workers are daemon threads, so ``asyncio.run`` does not wait for them.
    Depending on ``monkeypatch`` puts this teardown before the patches are undone.
    """
    yield
    exempt_workers.wait_idle(5.0)
