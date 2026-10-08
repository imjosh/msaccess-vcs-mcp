"""Controlled ownership failures must never cause another COM close."""
from pathlib import Path
from types import SimpleNamespace
import sys

import pytest

from tests import policy_live_host as live


def host(monkeypatch, *, exited=False):
    actions = []
    monkeypatch.setitem(sys.modules, 'win32api', SimpleNamespace(
        CloseHandle=lambda h: actions.append(('release_handle', h)),
        OpenProcess=lambda *args: 20))
    monkeypatch.setitem(sys.modules, 'win32event', SimpleNamespace(
        WAIT_OBJECT_0=0, WAIT_TIMEOUT=258,
        WaitForSingleObject=lambda h, ms: 0 if exited else 258))
    result = live.OwnedHost.__new__(live.OwnedHost)
    result.pid, result.creation, result.handle = 10, 123, 10
    result.path = Path('owned.accdb').resolve()
    result.receipt, result.unrelated, result.addin = {}, {}, None
    result.app = SimpleNamespace(CurrentDb=lambda: SimpleNamespace(Name=str(result.path)))
    monkeypatch.setattr(live, 'raw_creation', lambda h: 123)
    from msaccess_vcs_mcp.access_com import process_qos
    monkeypatch.setattr(process_qos, 'pid_from_access_app', lambda app: 10)
    return result, actions


def test_confirmed_already_exited_never_calls_com_or_quit(monkeypatch):
    owned, actions = host(monkeypatch, exited=True)
    owned.app = None  # Any attempted COM access would fail.
    owned.close(lambda: actions.append(('release_com',)))
    assert owned.receipt['already_exited_before_cleanup'] is True
    assert owned.receipt['original_handle_exit_confirmed'] is True
    assert actions == [('release_handle', 10)]


@pytest.mark.parametrize('uncertainty', ['creation', 'path', 'pid', 'query'])
def test_uncertain_identity_or_path_is_left_untouched(monkeypatch, uncertainty):
    owned, actions = host(monkeypatch)
    if uncertainty == 'creation':
        monkeypatch.setattr(live, 'raw_creation', lambda h: 456)
    elif uncertainty == 'path':
        owned.app.CurrentDb = lambda: SimpleNamespace(Name='unrelated.accdb')
    elif uncertainty == 'pid':
        from msaccess_vcs_mcp.access_com import process_qos
        monkeypatch.setattr(process_qos, 'pid_from_access_app', lambda app: 99)
    else:
        def fail(handle):
            raise OSError('identity unreadable')
        monkeypatch.setattr(live, 'raw_creation', fail)
    with pytest.raises((AssertionError, OSError)):
        owned.close(lambda: actions.append(('release_com',)))
    assert all(action[0] == 'release_handle' for action in actions)


def test_identity_is_rechecked_after_owner_clear(monkeypatch):
    owned, actions = host(monkeypatch)
    confirmations = iter([True, False])
    owned.verified = lambda: next(confirmations)
    owned.addin = SimpleNamespace(call_sync=lambda command: '{"success":true}')
    owned.receipt['checks'] = []
    with pytest.raises(AssertionError, match='Changed identity/path'):
        owned.close(lambda: actions.append(('release_com',)))
    assert actions == [('release_handle', 10)]
