"""Final MSAA action validation and accessibility pointer ownership."""

from contextlib import nullcontext
from unittest.mock import Mock

import pytest

from msaccess_vcs_mcp import msaa


@pytest.mark.parametrize(
    "state,role,name,path,expected,released",
    [
        (None, msaa.ROLE_PUSHBUTTON, "OK", (0, 0), False, [9, 1, 8, 2, 3]),
        (0, msaa.ROLE_PUSHBUTTON, "OK", (0, 0), True, [9, 1, 8, 2, 3]),
        (msaa.STATE_INVISIBLE, msaa.ROLE_PUSHBUTTON, "OK", (0, 0), False, [9, 1, 8, 2, 3]),
        (msaa.STATE_UNAVAILABLE, msaa.ROLE_PUSHBUTTON, "OK", (0, 0), False, [9, 1, 8, 2, 3]),
        (0, msaa.ROLE_PUSHBUTTON, "Cancel", (0, 0), False, [9, 1, 8, 2, 3]),
        (0, msaa.ROLE_STATICTEXT, "OK", (0, 0), False, [9, 1, 8, 2, 3]),
        (0, msaa.ROLE_PUSHBUTTON, "OK", (0, 4), False, [9, 1, 3, 8, 2]),
        (0, msaa.ROLE_PUSHBUTTON, "OK", (-1,), False, [2, 9, 1]),
        (0, msaa.ROLE_PUSHBUTTON, "OK", (1,), False, [2, 9, 1]),
    ],
    ids=["unreadable", "zero", "invisible", "unavailable", "changed-name",
         "non-button", "out-of-range", "negative-index", "simple-child"],
)
def test_press_validates_control_and_releases_tree(
    monkeypatch, state, role, name, path, expected, released
):
    action = Mock(return_value=0)
    release = Mock()
    methods = Mock(return_value=action)
    fetch = Mock(return_value=1)
    monkeypatch.setattr(msaa, "_ComScope", nullcontext)
    monkeypatch.setattr(msaa, "_from_window", fetch)
    monkeypatch.setattr(msaa, "_children", lambda ptr: {1: [2, None, 9], 2: [3, 8]}[ptr])
    monkeypatch.setattr(msaa, "_release", release)
    monkeypatch.setattr(msaa, "_int_property", lambda ptr, slot: state if slot == msaa._ACC_STATE else role)
    monkeypatch.setattr(msaa, "_name", lambda ptr: name)
    monkeypatch.setattr(msaa, "_method", methods)

    assert msaa.press(123, path, "OK", timeout_ms=500) is expected

    fetch.assert_called_once_with(123, 500)
    assert [call.args[0] for call in release.call_args_list] == released
    if expected:
        methods.assert_called_once_with(3, msaa._ACC_DO_DEFAULT_ACTION, msaa._VARIANT)
        action.assert_called_once()
        assert action.call_args.args[0] == 3
    else:
        methods.assert_not_called()
        action.assert_not_called()
