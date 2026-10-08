"""Fixture-only Win32 Cancel control driver; never imported by production tools.

Fresh HWND enumeration avoids the historical accessibility element cache.
BM_CLICK / fresh MSAA default action activates the observed native Cancel, without coordinates,
keyboard shortcuts, close-error injection or cancellation flags.
"""
import json
import time
from pathlib import Path
from dataclasses import asdict

PROMPT = "Do you want to save changes to the design of form 'frmSideCancel'?"


def write(path, value):
    path = Path(path)
    temp = path.with_suffix('.tmp')
    temp.write_text(json.dumps(value, indent=2), encoding='utf-8')
    temp.replace(path)


def validate_dialog(dialog, pid):
    assert dialog['pid'] == pid
    assert dialog['class'] in ('#32770', 'NUIDialog')
    assert dialog['visible']
    assert PROMPT in dialog['texts']
    buttons = dialog['buttons']
    assert {c['text'].replace('&', '') for c in buttons} == {'Yes', 'No', 'Cancel'}
    cancel = [c for c in buttons if c['text'].replace('&', '') == 'Cancel']
    assert len(cancel) == 1
    return cancel[0]


def observe(pid):
    import win32gui
    import win32process
    from msaccess_vcs_mcp.dialog_recovery import _netui_contents, _child_contents
    found = []

    def visit(hwnd, _):
        if (win32process.GetWindowThreadProcessId(hwnd)[1] != pid or
                win32gui.GetClassName(hwnd) not in ('#32770', 'NUIDialog') or
                not win32gui.IsWindowVisible(hwnd)):
            return
        cls = win32gui.GetClassName(hwnd)
        buttons, texts = (_netui_contents(hwnd) if cls == 'NUIDialog' else _child_contents(hwnd))
        dialog = dict(hwnd=hwnd, pid=pid, visible=True,
                      buttons=[asdict(b) for b in buttons], texts=texts,
                      owner=win32gui.GetWindow(hwnd, 4))
        if not dialog['owner'] or win32process.GetWindowThreadProcessId(dialog['owner'])[1] != pid:
            return
        dialog['class'] = cls
        try:
            validate_dialog(dialog, pid)
        except AssertionError:
            return
        found.append(dialog)
    win32gui.EnumWindows(visit, None)
    assert len(found) <= 1, 'Ambiguous fixture save-design dialogs'
    return found[0] if found else None


def click(work, pending, verify_identity, timeout=30):
    import win32gui
    import win32process
    from msaccess_vcs_mcp.dialog_recovery import Win32Backend, ButtonInfo, CLICK_DELIVERED
    end = time.monotonic() + timeout
    dialog = None
    while time.monotonic() < end:
        verify_identity()
        dialog = observe(pending['pid'])
        if dialog:
            break
        time.sleep(.1)
    assert dialog, 'Native dialog appearance timeout'
    case = pending['case']
    write(work / (case + '-dialog.json'), dict(pending=pending, dialog=dialog))
    verify_identity()
    fresh = observe(pending['pid'])
    assert fresh == dialog, 'Native dialog changed before dispatch'
    button = validate_dialog(fresh, pending['pid'])
    assert win32process.GetWindowThreadProcessId(button['hwnd'])[1] == pending['pid']
    assert win32gui.IsChild(dialog['hwnd'], button['hwnd'])
    receipt = dict(pending=pending, dialog=dialog, button=button,
                   action='literal native Cancel button default action', status='dispatching')
    target = work / (case + '-click.json')
    assert not target.exists(), 'Never replay a dispatched or uncertain click'
    write(target, receipt)
    try:
        # SMTO_ABORTIFHUNG; one dispatch only, even if delivery is uncertain.
        delivery = Win32Backend().click(ButtonInfo(**button), expected_pid=pending['pid'], timeout_ms=3000)
        receipt['delivery'] = delivery
        assert delivery == CLICK_DELIVERED, 'Cancel delivery not confirmed; never replay'
        receipt['status'] = 'click returned'
    except BaseException as exc:
        receipt.update(status='uncertain', error=repr(exc))
        raise
    finally:
        write(target, receipt)


if __name__ == '__main__':
    import sys
    sys.dont_write_bytecode = True
    sys.path.insert(0, str(Path(__file__).resolve().parents[3] / 'msaccess-vcs-mcp'))
    import win32api
    import win32event
    from tests.policy_live_host import raw_creation
    from msaccess_vcs_mcp.access_com.process_qos import list_access_pids_or_none
    pending_path = Path(sys.argv[1]).resolve()
    pending = json.loads(pending_path.read_text())
    work = pending_path.parent
    assert work.parent == (Path(__file__).resolve().parents[3] / 'verification/click-cancel-integration').resolve()
    assert pending['work'] == str(work)
    assert pending['database'] == str(work / 'cancel-fixture.accdb')
    assert isinstance(pending['creation_FILETIME'], str)
    handle = win32api.OpenProcess(0x1000 | 0x100000, False, pending['pid'])
    try:
        def verify():
            assert raw_creation(handle) == int(pending['creation_FILETIME'])
            assert win32event.WaitForSingleObject(handle, 0) == win32event.WAIT_TIMEOUT
            assert list_access_pids_or_none() == {pending['pid']}
        click(work, pending, verify)
    finally:
        win32api.CloseHandle(handle)
