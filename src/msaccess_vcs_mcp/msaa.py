"""Read and press Office NetUI dialog controls through MSAA (IAccessible).

Access draws some message boxes as a top-level ``NUIDialog`` window whose only
child is a ``NetUIHWND``. That child has no Win32 ``Static`` or ``Button``
windows; its text and buttons are visible only through accessibility. Seen on
Microsoft 365 Access: ``MsgBox`` with the ``@``-separated bold form (the
add-in's ``MsgBox2``) and Access's own error dialogs such as a missing form.

NetUI's ``IDispatch::Invoke`` returns ``E_NOTIMPL``, so pywin32 late binding
cannot call ``IAccessible``. This module calls the vtable directly with ctypes.
The object is fetched with ``SendMessageTimeout(WM_GETOBJECT)`` so a hung
window is skipped rather than waited on.
"""

from __future__ import annotations

import ctypes
from ctypes import wintypes
from dataclasses import dataclass

WM_GETOBJECT = 0x003D
OBJID_CLIENT = -4
SMTO_ABORTIFHUNG = 0x0002
ROLE_PUSHBUTTON = 43
ROLE_STATICTEXT = 41
STATE_INVISIBLE = 0x8000
STATE_UNAVAILABLE = 0x0001
# A NetUI message box is a few levels deep with a dozen controls. The caps stop a
# pathological tree from turning one inspection into a long crawl.
MAX_DEPTH = 6
MAX_NODES = 200

_VT_I4 = 3
_VT_DISPATCH = 9

# IAccessible vtable slots (after IUnknown's 3 and IDispatch's 4).
_ACC_CHILD_COUNT = 8
_ACC_NAME = 10
_ACC_ROLE = 13
_ACC_STATE = 14
_ACC_DO_DEFAULT_ACTION = 25


class _GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", ctypes.c_ulong),
        ("Data2", ctypes.c_ushort),
        ("Data3", ctypes.c_ushort),
        ("Data4", ctypes.c_ubyte * 8),
    ]


# {618736E0-3C3D-11CF-810C-00AA00389B71}
_IID_IACCESSIBLE = _GUID(
    0x618736E0, 0x3C3D, 0x11CF, (ctypes.c_ubyte * 8)(0x81, 0x0C, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71)
)


class _VariantValue(ctypes.Union):
    _fields_ = [("lVal", ctypes.c_long), ("pointer", ctypes.c_void_p), ("_pad", ctypes.c_byte * 16)]


class _VARIANT(ctypes.Structure):
    _fields_ = [
        ("vt", ctypes.c_ushort),
        ("r1", ctypes.c_ushort),
        ("r2", ctypes.c_ushort),
        ("r3", ctypes.c_ushort),
        ("value", _VariantValue),
    ]


def _self_id() -> _VARIANT:
    variant = _VARIANT()
    variant.vt = _VT_I4
    variant.value.lVal = 0  # CHILDID_SELF
    return variant


def _method(ptr: int, slot: int, *argtypes):
    vtable = ctypes.cast(ptr, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
    return ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, *argtypes)(vtable[slot])


def _release(ptr: int) -> None:
    _method(ptr, 2)(ptr)


def _name(ptr: int) -> str:
    out = ctypes.c_void_p()
    call = _method(ptr, _ACC_NAME, _VARIANT, ctypes.POINTER(ctypes.c_void_p))
    if call(ptr, _self_id(), ctypes.byref(out)) != 0 or not out.value:
        return ""
    try:
        return ctypes.wstring_at(out.value)
    finally:
        ctypes.windll.oleaut32.SysFreeString(ctypes.c_void_p(out.value))


def _int_property(ptr: int, slot: int) -> int | None:
    out = _VARIANT()
    call = _method(ptr, slot, _VARIANT, ctypes.POINTER(_VARIANT))
    if call(ptr, _self_id(), ctypes.byref(out)) != 0:
        return None
    if out.vt == _VT_I4:
        return int(out.value.lVal)
    ctypes.windll.oleaut32.VariantClear(ctypes.byref(out))
    return None


def _children(ptr: int) -> list[int | None]:
    """Child IAccessible pointers in order. ``None`` holds the place of a simple element.

    The caller releases every pointer returned.
    """
    count = ctypes.c_long()
    if _method(ptr, _ACC_CHILD_COUNT, ctypes.POINTER(ctypes.c_long))(ptr, ctypes.byref(count)) != 0:
        return []
    if count.value <= 0:
        return []
    items = (_VARIANT * count.value)()
    got = ctypes.c_long()
    if ctypes.windll.oleacc.AccessibleChildren(
        ctypes.c_void_p(ptr), 0, count.value, items, ctypes.byref(got)
    ) < 0:
        return []
    found: list[int | None] = []
    for item in items[: got.value]:
        if item.vt == _VT_DISPATCH and item.value.pointer:
            dispatch = item.value.pointer
            accessible = ctypes.c_void_p()
            query = _method(dispatch, 0, ctypes.POINTER(_GUID), ctypes.POINTER(ctypes.c_void_p))
            ok = query(dispatch, ctypes.byref(_IID_IACCESSIBLE), ctypes.byref(accessible)) == 0
            _release(dispatch)
            found.append(accessible.value if ok and accessible.value else None)
        else:
            ctypes.windll.oleaut32.VariantClear(ctypes.byref(item))
            found.append(None)
    return found


def _from_window(hwnd: int, timeout_ms: int) -> int | None:
    result = ctypes.c_size_t()
    sent = ctypes.windll.user32.SendMessageTimeoutW(
        wintypes.HWND(hwnd),
        WM_GETOBJECT,
        0,
        OBJID_CLIENT,
        SMTO_ABORTIFHUNG,
        int(timeout_ms),
        ctypes.byref(result),
    )
    if not sent or not result.value:
        return None
    accessible = ctypes.c_void_p()
    hr = ctypes.windll.oleacc.ObjectFromLresult(
        ctypes.c_size_t(result.value),
        ctypes.byref(_IID_IACCESSIBLE),
        ctypes.c_size_t(0),
        ctypes.byref(accessible),
    )
    return accessible.value if hr == 0 and accessible.value else None


class _ComScope:
    """COM initialised on this thread for the duration, and undone if we did it."""

    def __enter__(self) -> "_ComScope":
        # S_OK and S_FALSE both need a matching CoUninitialize; a mode clash does not.
        self._initialised = ctypes.windll.ole32.CoInitializeEx(None, 2) in (0, 1)
        return self

    def __exit__(self, *_exc) -> None:
        if self._initialised:
            ctypes.windll.ole32.CoUninitialize()


@dataclass(frozen=True)
class Control:
    """One visible text or push button, with its child-index path from the window's client object."""

    role: int
    name: str
    path: tuple[int, ...]


def read_controls(hwnd: int, timeout_ms: int = 2000) -> list[Control]:
    """Visible static texts and push buttons under ``hwnd``, in tree order.

    Empty when the window does not answer ``WM_GETOBJECT`` in time or exposes
    no accessible object.
    """
    with _ComScope():
        root = _from_window(hwnd, timeout_ms)
        if root is None:
            return []
        found: list[Control] = []
        visited = [0]

        def _walk(ptr: int, path: tuple[int, ...]) -> None:
            for index, child in enumerate(_children(ptr)):
                if child is None:
                    continue
                try:
                    visited[0] += 1
                    if visited[0] > MAX_NODES:
                        return
                    state = _int_property(child, _ACC_STATE) or 0
                    if state & STATE_INVISIBLE:
                        continue
                    role = _int_property(child, _ACC_ROLE)
                    name = _name(child)
                    here = (*path, index)
                    if role in (ROLE_PUSHBUTTON, ROLE_STATICTEXT) and name.strip():
                        found.append(Control(role, name, here))
                    elif len(here) < MAX_DEPTH:
                        _walk(child, here)
                finally:
                    _release(child)

        try:
            _walk(root, ())
        finally:
            _release(root)
        return found


def press(hwnd: int, path: tuple[int, ...], expected_name: str, timeout_ms: int = 2000) -> bool:
    """Run the default action of the push button at ``path`` under ``hwnd``.

    True only when a readable state confirms the control there is still a
    visible, enabled push button named ``expected_name`` and accepted the action.
    """
    with _ComScope():
        current = _from_window(hwnd, timeout_ms)
        if current is None:
            return False
        try:
            for index in path:
                children = _children(current)
                chosen = children[index] if 0 <= index < len(children) else None
                for other in children:
                    if other is not None and other != chosen:
                        _release(other)
                _release(current)
                current = chosen
                if current is None:
                    return False
            state = _int_property(current, _ACC_STATE)
            if state is None or state & (STATE_INVISIBLE | STATE_UNAVAILABLE):
                return False
            if _int_property(current, _ACC_ROLE) != ROLE_PUSHBUTTON or _name(current) != expected_name:
                return False
            return _method(current, _ACC_DO_DEFAULT_ACTION, _VARIANT)(current, _self_id()) == 0
        finally:
            if current is not None:
                _release(current)
