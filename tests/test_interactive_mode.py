"""Omitting a policy never mutates process-wide interaction state."""
from unittest.mock import Mock
import pytest
from msaccess_vcs_mcp.decision_policy import select_interactive_mode, clear_operation_policy

@pytest.mark.parametrize("policy", [None, "block", "skip"])
def test_compatibility_helpers_do_not_call_the_addin(policy):
    addin = Mock()
    assert select_interactive_mode(addin, policy) is None
    assert clear_operation_policy(addin) is None
    addin.call_sync.assert_not_called()
