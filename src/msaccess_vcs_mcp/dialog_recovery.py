"""Inspect and dismiss Access, VBA, and add-in dialogs from outside the Access process.

Access COM calls block while a modal dialog is open or VBA is in break mode.
This module uses Win32 window enumeration and button messages so recovery can
run while another MCP request is waiting on that COM call. Office NetUI dialogs
(``NUIDialog``) have no Win32 buttons; their text and buttons are read and
pressed through MSAA (see ``msaa.py``). It never sends keystrokes or clicks by
screen coordinate.

A dialog tool result describes what was found and what was clicked. Dismissing
an error dialog does not make the failed Access operation a success.
"""

from __future__ import annotations

import itertools
import os
import threading
import time
from dataclasses import dataclass
from typing import Any, Callable, Iterable, Protocol

from .access_com.instance_registry import list_owned, process_create_time
from .access_gate import get_access_gate
from .config import get_config

DEFAULT_TIMEOUT_SEC = 5.0
MAX_TIMEOUT_SEC = 30.0
MAX_AUTO_DIALOGS = 5

# Button messages. BM_CLICK is delivered to a specific button window.
BM_CLICK = 0x00F5
WM_CLOSE = 0x0010
WM_NULL = 0x0000
SMTO_ABORTIFHUNG = 0x0002

ACCESS_EXE_NAMES = {"msaccess.exe", "msaccess"}
# Top-level classes of a standard dialog box. ``#32770`` is the Win32 dialog.
# ``NUIDialog`` is Office NetUI: Access uses it for ``MsgBox`` with the
# ``@``-separated bold form (the add-in's ``MsgBox2``) and for its own error
# dialogs. Both get the same positive-signature rules.
WIN32_DIALOG_CLASS = "#32770"
NETUI_DIALOG_CLASS = "NUIDialog"
NETUI_HOST_CLASS = "NetUIHWND"
DIALOG_CLASSES = {WIN32_DIALOG_CLASS, NETUI_DIALOG_CLASS}
ADDIN_CAPTIONS = {"msaccessvcs", "version control system"}
VBA_RUNTIME_ERROR_KIND = "vba_runtime_error"
VBA_COMPILE_ERROR_KIND = "vba_compile_error"
VBA_BREAK_KIND = "vba_break"
ACCESS_DIALOG_KIND = "access_dialog"
ADDIN_WINDOW_KIND = "addin_window"
# An ordinary Access window: not a dialog, never reported.
IGNORED_KIND = "ignored"
FAILURE_KINDS = {VBA_RUNTIME_ERROR_KIND, VBA_COMPILE_ERROR_KIND}
# A kind other than "unknown" is assigned only by a positive signature (see
# classify_window). Automatic recovery never clicks "unknown".
UNKNOWN_KIND = "unknown"
# A standard dialog whose only actionable button is OK, under any caption: a VBA
# MsgBox with a custom title. Positive signature is the button set alone.
MSGBOX_KIND = "vba_msgbox"
BLOCKING_KINDS = {
    VBA_RUNTIME_ERROR_KIND,
    VBA_COMPILE_ERROR_KIND,
    ACCESS_DIALOG_KIND,
    MSGBOX_KIND,
    UNKNOWN_KIND,
    VBA_BREAK_KIND,
}
# Kinds whose OK-only form the ``safe`` policy may acknowledge. Access error and
# warning dialogs are report-only, and compile errors belong to
# ``end_runtime_error``.
SAFE_OK_KINDS = {MSGBOX_KIND}
POLL_INTERVAL_SEC = 0.05
DIALOG_ID_PREFIX = "hwnd:"
_DESTRUCTIVE_TEXT = (
    "save changes",
    "discard",
    "delete",
    "overwrite",
    "remove all",
    "uninstall",
    "without saving",
    "lose your",
    "lose changes",
)

# What ``WindowBackend.click`` reports. ``not_sent`` means nothing reached the
# window (wrong owner, unresponsive, or the button refused the action).
# ``uncertain`` means it may have: the send timed out or failed part way.
CLICK_DELIVERED = "delivered"
CLICK_NOT_SENT = "not_sent"
CLICK_UNCERTAIN = "uncertain"
# How long past the click timeout a finished gated call waits for a click still
# in flight: the Win32 click may spend one timeout on the responsiveness probe
# and another on the press.
CLICK_SETTLE_MARGIN_SEC = 1.0

# Reservation token -> record of an End/cancel done from outside Access. A record
# is reserved (``pending``) before the action is delivered, because Access can
# resume and finish the blocked call before the click returns. The action then
# settles it: ``confirmed``, ``uncertain`` (kept only for a gated call), or gone
# when nothing was sent. A delivered cancel close is only a ``request``: the
# add-in may ask to confirm and resume, so it settles as ``requested`` (kept only
# for a gated call) and that call's own result decides it. A record made while a
# gated call was in flight on the same database carries that call (``busy_with``
# and ``call_id``) and is used up when the call finishes; otherwise a confirmed
# record stays as ``last_interruption`` until the process identity changes or the
# next gated call on its database starts.
_interruptions: dict[int, dict[str, Any]] = {}
_interruptions_changed = threading.Condition()
_reservation_tokens = itertools.count(1)
_INTERNAL_FIELDS = {"call_id", "identity", "state", "deadline", "request"}


def _same_path(a: str | None, b: str | None) -> bool:
    if not a or not b:
        return False
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def _interrupts_execution(kind: str, button_text: str | None) -> bool:
    """The one rule for what makes a dismissal an interruption.

    A runtime or compile error dialog, or any click on End. Both
    ``dismiss_dialog`` and ``recover_dialogs`` use it. An ordinary Access dialog
    is not a failure by its wording.
    """
    return kind in FAILURE_KINDS or _button_label(button_text or "") == "end"


def _reserve_interruption(
    report: dict[str, Any],
    database_path: str | None,
    item: dict[str, Any],
    button_text: str | None,
    settle_within_sec: float,
    message: str | None = None,
    request: bool = False,
) -> int:
    """Reserve an interruption for ``item`` before the action that causes it.

    Attaches the in-flight gated call when it is on the same database. Returns
    the token ``_settle_interruption`` takes once the action has been tried.
    ``request`` marks an action that only asks Access to stop (a posted cancel
    close), so its delivery does not confirm the interruption.
    """
    identity = (int(report["pid"]), report.get("create_time"))
    record: dict[str, Any] = {
        "dialog_id": item["dialog_id"],
        "kind": item["kind"],
        "button": button_text,
        "title": item.get("title"),
        "message": message if message is not None else item.get("message"),
        "database_path": database_path,
        "busy_with": None,
        "identity": identity,
        "state": "pending",
        "deadline": time.monotonic() + settle_within_sec,
        "request": request,
    }
    current = get_access_gate().current_in_flight()
    if current is not None and _same_path(current.database, database_path):
        record["busy_with"] = {"tool": current.tool, "database": current.database}
        record["call_id"] = current.call_id
    token = next(_reservation_tokens)
    with _interruptions_changed:
        # Any other identity on this PID is a dead process; forget it.
        for key in [
            k for k, v in _interruptions.items()
            if v["identity"][0] == identity[0] and v["identity"] != identity
        ]:
            del _interruptions[key]
        _interruptions[token] = record
    return token


def _settle_interruption(token: int, outcome: str) -> None:
    """Record how the reserved action went. A reservation already used up stays gone."""
    with _interruptions_changed:
        record = _interruptions.get(token)
        if record is None:
            return
        if outcome == CLICK_DELIVERED and record["request"]:
            # Only the call's result can confirm a request; with no call, nothing will.
            if record.get("call_id") is not None:
                record["state"] = "requested"
            else:
                del _interruptions[token]
        elif outcome == CLICK_DELIVERED:
            record["state"] = "confirmed"
            if record["busy_with"] is None:
                # Only the newest free record is reported.
                for key in [
                    k for k, v in _interruptions.items()
                    if k != token and v["busy_with"] is None and v["identity"] == record["identity"]
                ]:
                    del _interruptions[key]
        elif outcome == CLICK_UNCERTAIN and record.get("call_id") is not None:
            record["state"] = "uncertain"
        else:
            del _interruptions[token]
        _interruptions_changed.notify_all()


def _interruption_for(pid: int | None, create_time: int | None) -> dict[str, Any] | None:
    if not pid:
        return None
    with _interruptions_changed:
        confirmed = [
            (token, record) for token, record in _interruptions.items()
            if record["identity"] == (int(pid), create_time) and record["state"] == "confirmed"
        ]
    if not confirmed:
        return None
    _token, record = max(confirmed, key=lambda pair: pair[0])
    return {k: v for k, v in record.items() if k not in _INTERNAL_FIELDS}


def begin_gated_call(database_path: str | None, call_id: int | None) -> None:
    """Gated call ``call_id`` on ``database_path`` starts: other records there are stale.

    That covers unattached records and one reserved against an earlier call just
    as the gate released it, after that call had used up its records.
    """
    with _interruptions_changed:
        for key in [
            k for k, v in _interruptions.items()
            if v.get("call_id") != call_id and _same_path(v["database_path"], database_path)
        ]:
            del _interruptions[key]


def _await_reserved_actions(call_id: int) -> None:
    """Wait, bounded by their deadlines, for actions reserved against ``call_id`` to settle.

    The caller holds ``_interruptions_changed``.
    """
    while True:
        deadlines = [
            v["deadline"] for v in _interruptions.values()
            if v.get("call_id") == call_id and v["state"] == "pending"
        ]
        remaining = max(deadlines, default=0.0) - time.monotonic()
        if remaining <= 0:
            return
        _interruptions_changed.wait(remaining)


def finish_gated_call(call_id: int | None, result: Any) -> Any:
    """Use up the interruption records of a finished gated call.

    A finished handler (``result`` is not None) first waits for any dialog
    action reserved against it that is still being delivered, at most until that
    action's deadline. An interrupted result is ``success: false`` with
    ``execution_interrupted: true``. The primary ``error_pattern`` follows the
    precedence ``decision_required`` > ``execution_interrupted`` > plain error >
    ``interruption_uncertain``, so a decision-required result keeps its pattern,
    decisions and error text and only gains the flag. An action whose delivery is
    uncertain, or did not settle in time, adds ``interruption_uncertain: true``
    and never leaves the result a success. ``result`` None uses the records up
    without waiting: the call raised, or the gate is releasing it.

    A delivered cancel close is a request, and the result decides it.
    ``cancelled: true`` (the add-in's confirmation, from its sync result or its
    ``cancelled`` terminal callback) makes it an interruption as above. A
    success stands and gains ``cancel_not_honored: true``: the person answered
    No and the operation completed. Any other failure gains
    ``interruption_uncertain: true``, because a result without ``cancelled``
    (a whole-export sync fallback, for one) cannot say whether the cancel
    stopped it.
    """
    if call_id is None:
        return result
    with _interruptions_changed:
        if result is not None:
            _await_reserved_actions(call_id)
        matched = [k for k, v in _interruptions.items() if v.get("call_id") == call_id]
        records = [_interruptions.pop(k) for k in matched]
    if not records or not isinstance(result, dict):
        return result
    decision_required = (
        bool(result.get("decision_required")) or result.get("error_pattern") == "decision_required"
    )
    cancelled = result.get("cancelled") is True
    confirmed = [
        record for record in records
        if record["state"] == "confirmed" or (record["state"] == "requested" and cancelled)
    ]
    if not confirmed:
        unsettled = [record for record in records if record["state"] != "requested"]
        failed = result.get("success") is False
        if not (unsettled or failed or decision_required):
            # Only cancel requests, and the call completed: the add-in resumed.
            result["cancel_not_honored"] = True
            return result
        result["success"] = False
        result["interruption_uncertain"] = True
        if failed or decision_required:
            return result
        result["error_pattern"] = "interruption_uncertain"
        if not result.get("error"):
            detail = unsettled[-1].get("message") or unsettled[-1].get("kind")
            result["error"] = (
                "A dialog action that interrupts this call was sent while it ran, but its "
                f"delivery was not confirmed ({detail}). The call may not have completed; "
                "check its effects before relying on it."
            )
        return result
    record = confirmed[-1]
    result["success"] = False
    result["execution_interrupted"] = True
    if decision_required:
        return result
    result["error_pattern"] = "execution_interrupted"
    if not result.get("error"):
        detail = record.get("message") or "a runtime or compile error"
        result["error"] = (
            f"Execution was interrupted: a dialog was dismissed while this call ran ({detail})."
        )
    return result


@dataclass(frozen=True)
class ButtonInfo:
    """A dialog button.

    For a Win32 button ``hwnd`` is the button window and ``path`` is None. For a
    NetUI button ``hwnd`` is the ``NetUIHWND`` that hosts it and ``path`` is its
    MSAA child-index path from that window's client object.
    """

    hwnd: int
    text: str
    path: tuple[int, ...] | None = None


@dataclass(frozen=True)
class WindowInfo:
    hwnd: int
    pid: int
    title: str
    class_name: str
    texts: tuple[str, ...] = ()
    buttons: tuple[ButtonInfo, ...] = ()


@dataclass(frozen=True)
class ProcessIdentity:
    """What the backend could read about a PID. ``None`` means unknown."""

    name: str | None = None
    create_time: int | None = None
    running: bool | None = None

    @property
    def is_access(self) -> bool:
        return bool(self.name) and self.name.lower() in ACCESS_EXE_NAMES

    @property
    def confirmed(self) -> bool:
        """Access by name and a readable creation time: the only state that may be acted on."""
        return self.is_access and self.create_time is not None


class WindowBackend(Protocol):
    def list_windows(self) -> list[WindowInfo]: ...

    def process_identity(self, pid: int) -> ProcessIdentity: ...

    def click(
        self, button: ButtonInfo, *, expected_pid: int | None = None, timeout_ms: int = 5000
    ) -> str:
        """Press ``button``. ``CLICK_DELIVERED`` only when it was delivered within
        ``timeout_ms``; ``CLICK_NOT_SENT`` when nothing reached the window;
        otherwise ``CLICK_UNCERTAIN``."""
        ...

    def close(self, hwnd: int) -> None: ...

    def responsive(self, hwnd: int, timeout_ms: int) -> bool: ...


def dialog_id_for(hwnd: int) -> str:
    """The only place a dialog id is built."""
    return f"{DIALOG_ID_PREFIX}{int(hwnd)}"


def parse_dialog_id(dialog_id: str) -> int | None:
    """The only place a dialog id is read. None when it is not one of ours."""
    text = (dialog_id or "").strip()
    if not text.lower().startswith(DIALOG_ID_PREFIX):
        return None
    try:
        return int(text[len(DIALOG_ID_PREFIX):])
    except ValueError:
        return None


def dialog_timeout_sec(override: float | None = None) -> float:
    """Bound a dialog wait. The environment supplies the default."""
    if override is None:
        try:
            override = float(get_config().get("ACCESS_VCS_DIALOG_TIMEOUT_SEC", DEFAULT_TIMEOUT_SEC))
        except (TypeError, ValueError):
            override = DEFAULT_TIMEOUT_SEC
    if override <= 0:
        override = DEFAULT_TIMEOUT_SEC
    return min(override, MAX_TIMEOUT_SEC)


def classify_window(window: WindowInfo) -> str:
    """Return a dialog kind, or ``IGNORED_KIND`` for ordinary Access windows.

    A known kind comes only from a positive signature: a recognised caption,
    text, or button set. A standard dialog box that matches none is ``unknown``,
    never a known kind by elimination.
    """
    title = window.title or ""
    title_l = title.lower()
    blob = " ".join([title, *window.texts]).lower()
    buttons = {_button_label(button.text) for button in window.buttons if button.text.strip()}
    class_name = window.class_name or ""

    if "[break]" in title_l:
        return VBA_BREAK_KIND
    if "run-time error" in blob or "runtime error" in blob:
        return VBA_RUNTIME_ERROR_KIND
    if "compile error" in blob:
        return VBA_COMPILE_ERROR_KIND
    if "end" in buttons and "debug" in buttons:
        return VBA_RUNTIME_ERROR_KIND

    # The window class decides before the caption: an add-in MsgBox2 box keeps the
    # add-in caption but is a real dialog (#32770 or NUIDialog) and blocks. Only a
    # window of another class with an add-in caption is the add-in's own form.
    caption_key = title_l.strip()
    if class_name not in DIALOG_CLASSES and (
        caption_key in ADDIN_CAPTIONS or caption_key.startswith("msaccessvcs")
    ):
        return ADDIN_WINDOW_KIND

    if class_name in DIALOG_CLASSES:
        if title_l.startswith("microsoft access"):
            return ACCESS_DIALOG_KIND
        if title_l.startswith("microsoft visual basic"):
            return ACCESS_DIALOG_KIND
        # Help is not a choice, so OK + Help is still a single-button box.
        actionable = {label for label in buttons if label != "help"}
        if actionable == {"ok"} and len(_actionable_buttons(window)) == 1:
            return MSGBOX_KIND
        return UNKNOWN_KIND
    return IGNORED_KIND


def _actionable_buttons(window: WindowInfo) -> list[ButtonInfo]:
    return [
        button
        for button in window.buttons
        if button.text.strip() and _button_label(button.text) != "help"
    ]


def _destructive_text(window: WindowInfo) -> bool:
    blob = " ".join([window.title, *window.texts]).lower()
    return any(phrase in blob for phrase in _DESTRUCTIVE_TEXT)


def _button_label(text: str) -> str:
    """Button caption without the Win32 accelerator marker.

    ``&End`` and ``End`` are the same choice. A doubled ``&&`` is a literal
    ampersand, not an accelerator.
    """
    raw = text or ""
    out: list[str] = []
    index = 0
    while index < len(raw):
        char = raw[index]
        if char == "&" and index + 1 < len(raw):
            nxt = raw[index + 1]
            out.append(nxt)
            index += 2
            continue
        out.append(char)
        index += 1
    return "".join(out).strip().lower()


def _button_named(window: WindowInfo, name: str) -> ButtonInfo | None:
    wanted = _button_label(name)
    if not wanted:
        return None
    for button in window.buttons:
        if _button_label(button.text) == wanted:
            return button
    return None


def auto_button(window: WindowInfo, kind: str, policy: str) -> str | None:
    """Return a button caption the policy may click, or None.

    Debug is never chosen. Destructive confirmations (save, discard, delete,
    overwrite) are never acknowledged automatically, including in a
    single-button ``vba_msgbox``. End is used only for a runtime-error dialog
    when the policy is ``end_runtime_error``, and does not depend on the
    destructive-text check: that guards OK clicks, and a runtime error's own
    text may say "deleted". ``safe`` clicks OK only on a ``vba_msgbox``;
    ``end_runtime_error`` also clicks OK on a compile error. An
    ``access_dialog`` or ``unknown`` dialog is never clicked.
    """
    if policy not in {"safe", "end_runtime_error"}:
        return None
    if kind == VBA_RUNTIME_ERROR_KIND:
        if policy == "end_runtime_error" and _button_named(window, "End") is not None:
            return "End"
        return None
    ok_kinds = SAFE_OK_KINDS
    if policy == "end_runtime_error":
        ok_kinds = SAFE_OK_KINDS | {VBA_COMPILE_ERROR_KIND}
    if kind not in ok_kinds or _destructive_text(window):
        return None
    # OK-only, ignoring a Help button that we do not click.
    actionable = [button.text.strip() for button in _actionable_buttons(window)]
    if len(actionable) == 1 and _button_label(actionable[0]) == "ok":
        return actionable[0]
    return None


def _policy_button(window: WindowInfo, kind: str, policy: str) -> ButtonInfo | None:
    """The button ``policy`` would click on ``window``, or None."""
    caption = auto_button(window, kind, policy)
    return None if caption is None else _button_named(window, caption)


def _filename_tokens(database_path: str) -> set[str]:
    base = os.path.basename(database_path)
    stem, _ext = os.path.splitext(base)
    tokens = set()
    if base:
        tokens.add(base.lower())
    if stem:
        tokens.add(stem.lower())
    return tokens


def _title_matches(title: str, database_path: str) -> bool:
    title_l = (title or "").lower()
    return any(token and token in title_l for token in _filename_tokens(database_path))


def resolve_target(
    windows: Iterable[WindowInfo],
    database_path: str,
    backend: WindowBackend,
    pid: int | None = None,
    create_time: int | None = None,
) -> dict[str, Any]:
    """Pick one Access process. Ambiguous or mismatched identity is an error.

    Success carries ``identity_confirmed``. Only a confirmed identity (Access by
    name, creation time readable and equal to ``create_time`` when given) may be
    acted on; an unconfirmed one may still be inspected. ``create_time`` in the
    result is the observed value, fixed for the rest of the call.
    """
    listed = list(windows)
    identities: dict[int, ProcessIdentity] = {}

    def _ident(candidate: int) -> ProcessIdentity:
        if candidate not in identities:
            identities[candidate] = backend.process_identity(candidate)
        return identities[candidate]

    def _describe(candidate: int) -> dict[str, Any]:
        ident = _ident(candidate)
        return {"pid": candidate, "create_time": ident.create_time, "process_name": ident.name}

    def _accept(candidate: int, matched_by: str) -> dict[str, Any]:
        ident = _ident(candidate)
        if ident.name and not ident.is_access:
            return {
                "success": False,
                "error_pattern": "not_access_process",
                "error": f"PID {candidate} is not MSACCESS.EXE ({ident.name}).",
                "pid": candidate,
            }
        if (
            create_time is not None
            and ident.create_time is not None
            and int(ident.create_time) != int(create_time)
        ):
            return {
                "success": False,
                "error_pattern": "process_identity_mismatch",
                "error": "PID was reused or does not match the expected process creation time.",
                "pid": candidate,
                "create_time": ident.create_time,
                "expected_create_time": create_time,
            }
        return {
            "success": True,
            **_describe(candidate),
            "identity_confirmed": ident.confirmed,
            "running": ident.running,
            "matched_by": matched_by,
        }

    if pid is not None:
        return _accept(pid, "pid")

    def _titled(window: WindowInfo) -> bool:
        return _title_matches(window.title, database_path) or any(
            _title_matches(text, database_path) for text in window.texts
        )

    titled = {window.pid for window in listed if _titled(window)}
    title_pids = {item for item in titled if _ident(item).is_access}
    unconfirmed_pids = {item for item in titled if not _ident(item).name}
    try:
        owned_records = [
            record for record in list_owned() if _same_path(record.database_path, database_path)
        ]
    except Exception:
        owned_records = []
    dead_owned = {record.pid for record in owned_records if _ident(record.pid).running is False}
    # Pruning retains uncertain records for bookkeeping. Only a matching,
    # readable creation stamp proves that the recorded database still owns this PID.
    owned_pids = {
        record.pid
        for record in owned_records
        if _ident(record.pid).is_access
        and record.create_time is not None
        and _ident(record.pid).create_time is not None
        and record.create_time == _ident(record.pid).create_time
    }

    # Prefer the intersection when both sources agree on a live window.
    both = title_pids & owned_pids
    if len(both) == 1:
        candidates = both
    elif title_pids:
        candidates = title_pids
    else:
        candidates = owned_pids

    if len(candidates) == 0:
        if dead_owned and not unconfirmed_pids:
            return {
                "success": False,
                "ready": False,
                "error_pattern": "access_not_running",
                "error": "The Access instance for this database is no longer running. Relaunch it.",
                "candidates": [_describe(item) for item in sorted(dead_owned)],
            }
        if unconfirmed_pids:
            return {
                "success": False,
                "error_pattern": "identity_unconfirmed",
                "error": (
                    "A window matches this database but its process name could not be read, "
                    "so it is not confirmed as Access. Nothing was touched."
                ),
                "candidates": [_describe(item) for item in sorted(unconfirmed_pids)],
            }
        return {
            "success": False,
            "error_pattern": "access_not_found",
            "error": "No Access window for this database was found. Pass pid when the window title does not contain the file name.",
        }
    if len(candidates) > 1:
        return {
            "success": False,
            "error_pattern": "ambiguous_instance",
            "error": "More than one Access window matches this database. Pass pid to choose one.",
            "candidates": [_describe(item) for item in sorted(candidates)],
        }
    return _accept(next(iter(candidates)), "database")


def _identity_refusal(report: dict[str, Any]) -> dict[str, Any] | None:
    """Refuse to act unless the target's identity is confirmed."""
    if report.get("identity_confirmed"):
        return None
    return {
        **report,
        "success": False,
        "error_pattern": "identity_unconfirmed",
        "error": (
            "The process name or creation time could not be confirmed. "
            "Inspection is reported; nothing was clicked or closed."
        ),
        "dismissed": False,
    }


def _signature(window: WindowInfo) -> tuple[Any, ...]:
    """What must not change between inspection and action.

    The kind alone is not enough: "Hello" and "Delete all records?" are both an
    OK-only ``vba_msgbox``.
    """
    return (window.class_name, window.title, window.texts, classify_window(window), window.buttons)


def _reverify(
    backend: WindowBackend, report: dict[str, Any], inspected: WindowInfo
) -> tuple[WindowInfo, str] | None:
    """Re-list windows right before acting. The fresh window and its kind, or None.

    None on any mismatch: the process identity changed, the handle is gone or
    belongs to another PID, or the window's signature differs from the inspected
    one. Windows reuses a closed box's handle, so a different box can sit behind
    the same handle and button handles. Only ``_press_fresh`` and
    ``_close_fresh`` call it, and they act only on the window it returns.
    """
    pid = int(report["pid"])
    ident = backend.process_identity(pid)
    if not ident.confirmed or ident.create_time != report.get("create_time"):
        return None
    for window in backend.list_windows():
        if window.hwnd != inspected.hwnd:
            continue
        if window.pid != pid or _signature(window) != _signature(inspected):
            return None
        return window, classify_window(window)
    return None


def _press_fresh(
    backend: WindowBackend,
    report: dict[str, Any],
    inspected: WindowInfo,
    choose: Callable[[WindowInfo, str], ButtonInfo | None],
    click_timeout_sec: float | None,
    database_path: str,
) -> tuple[str, dict[str, Any] | None, ButtonInfo | None]:
    """Re-verify, choose the button on the fresh window, and click that button.

    The only place a dialog button is clicked. Returns ``(status, record,
    button)``: ``dialog_changed`` when re-verification fails or ``choose`` picks
    nothing on the fresh window (nothing was clicked), ``dismiss_uncertain`` when
    the click was not delivered, otherwise ``clicked``. ``record`` describes the
    fresh window, so a result never reports text from a box other than the one
    clicked. A click that interrupts execution is reserved against the in-flight
    call before it is sent and settled by its delivery outcome.
    """
    fresh = _reverify(backend, report, inspected)
    if fresh is None:
        return "dialog_changed", None, None
    window, kind = fresh
    button = choose(window, kind)
    if button is None:
        return "dialog_changed", None, None
    record = _dialog_record(window, kind)
    timeout_sec = dialog_timeout_sec(click_timeout_sec)
    token = None
    if _interrupts_execution(kind, button.text):
        token = _reserve_interruption(
            report, database_path, record, button.text, 2 * timeout_sec + CLICK_SETTLE_MARGIN_SEC
        )
    outcome = CLICK_UNCERTAIN
    try:
        outcome = backend.click(
            button, expected_pid=int(report["pid"]), timeout_ms=int(timeout_sec * 1000)
        )
    finally:
        if token is not None:
            _settle_interruption(token, outcome)
    if outcome != CLICK_DELIVERED:
        return "dismiss_uncertain", record, button
    return "clicked", record, button


def _close_fresh(
    backend: WindowBackend,
    report: dict[str, Any],
    inspected: WindowInfo,
    database_path: str,
    interruption: str | None = None,
) -> dict[str, Any] | None:
    """Re-verify, then close the fresh window. Its record, or None when it changed.

    With ``interruption`` (its message) the close is a cancel request, reserved
    against the in-flight call before the close is posted. Posting it does not
    confirm the cancel: the add-in can ask to confirm and resume.
    """
    fresh = _reverify(backend, report, inspected)
    if fresh is None:
        return None
    window, kind = fresh
    record = _dialog_record(window, kind)
    token = None
    if interruption is not None:
        token = _reserve_interruption(
            report,
            database_path,
            record,
            None,
            CLICK_SETTLE_MARGIN_SEC,
            interruption,
            request=True,
        )
    outcome = CLICK_NOT_SENT
    try:
        backend.close(window.hwnd)
        outcome = CLICK_DELIVERED
    finally:
        if token is not None:
            _settle_interruption(token, outcome)
    return record


def _dialog_changed(report: dict[str, Any], item: dict[str, Any]) -> dict[str, Any]:
    return {
        **report,
        "success": False,
        "dismissed": False,
        "error_pattern": "dialog_changed",
        "error": (
            "The dialog, its content, or its process changed since it was inspected. "
            "Nothing was clicked. Re-inspect and decide again."
        ),
        "dialog_id": item.get("dialog_id"),
    }


def _readiness(
    *,
    identity_confirmed: bool,
    running: bool | None,
    responsive: bool | None,
    dialogs: list[dict[str, Any]],
    gate_busy_here: bool,
) -> tuple[bool, str | None]:
    """Readiness and, when not ready, the first reason as an ``error_pattern``.

    Every condition must be positively confirmed: an unknown answer (identity,
    liveness, responsiveness) is not ready.
    """
    if running is False:
        return False, "access_not_running"
    if running is None or not identity_confirmed:
        return False, "identity_unconfirmed"
    if responsive is None:
        return False, "no_windows_to_probe"
    if responsive is False:
        return False, "access_unresponsive"
    if any(item.get("kind") == VBA_BREAK_KIND for item in dialogs):
        return False, "vba_break"
    if any(item.get("kind") in BLOCKING_KINDS for item in dialogs):
        return False, "blocking_dialog"
    if gate_busy_here:
        return False, "server_busy"
    return True, None


def _gate_busy_here(gate: dict[str, Any]) -> bool:
    """The gate is busy with this database, or with one we cannot name."""
    if not gate.get("gate_busy"):
        return False
    operation = gate.get("operation") or {}
    return bool(operation.get("same_database")) or not operation.get("database")


def _dialog_record(window: WindowInfo, kind: str) -> dict[str, Any]:
    message = "\n".join(text for text in window.texts if text and text != window.title)
    return {
        "dialog_id": dialog_id_for(window.hwnd),
        "hwnd": window.hwnd,
        "kind": kind,
        "is_dialog": kind not in {VBA_BREAK_KIND, ADDIN_WINDOW_KIND, IGNORED_KIND},
        "title": window.title,
        "message": message,
        "buttons": [button.text for button in window.buttons],
        "class_name": window.class_name,
        "pid": window.pid,
    }


def _dialog_records(windows: Iterable[WindowInfo], pid: int) -> list[dict[str, Any]]:
    """Dialog records for the windows of one process, ordinary windows left out."""
    records = []
    for window in windows:
        if window.pid != pid:
            continue
        kind = classify_window(window)
        if kind == IGNORED_KIND:
            continue
        records.append(_dialog_record(window, kind))
    return records


def _with_readiness(report: dict[str, Any], dialogs: list[dict[str, Any]]) -> dict[str, Any]:
    """``report`` with ``dialogs`` and the fields derived from them and the gate.

    ``report`` must already carry pid, create_time, identity_confirmed, running,
    responsive, gate_busy and operation.
    """
    ready, _reason = _readiness(
        identity_confirmed=bool(report.get("identity_confirmed")),
        running=report.get("running"),
        responsive=report.get("responsive"),
        dialogs=dialogs,
        gate_busy_here=_gate_busy_here(
            {"gate_busy": report.get("gate_busy"), "operation": report.get("operation")}
        ),
    )
    interruption = _interruption_for(report.get("pid"), report.get("create_time"))
    return {
        **report,
        "dialogs": dialogs,
        "break_mode": any(item.get("kind") == VBA_BREAK_KIND for item in dialogs),
        "blocking_dialog": any(item.get("kind") in BLOCKING_KINDS for item in dialogs),
        "ready": ready,
        "execution_interrupted": interruption is not None,
        "last_interruption": interruption,
    }


def _gate_snapshot(database_path: str | None) -> dict[str, Any]:
    current = get_access_gate().current_in_flight()
    if current is None:
        return {"gate_busy": False, "operation": None}
    return {
        "gate_busy": True,
        "operation": {
            "tool": current.tool,
            "database": current.database,
            "elapsed_ms": round((time.perf_counter() - current.started_at) * 1000, 2),
            "same_database": _same_path(database_path, current.database),
        },
    }


def inspect_windows(
    windows: Iterable[WindowInfo],
    database_path: str,
    *,
    backend: WindowBackend,
    pid: int | None = None,
    create_time: int | None = None,
    responsive: bool | None = None,
) -> dict[str, Any]:
    """Build an inspection report for one Access instance. Does not click."""
    listed = list(windows)
    target = resolve_target(listed, database_path, backend, pid=pid, create_time=create_time)
    if not target.get("success"):
        target["dialogs"] = []
        return target

    chosen = int(target["pid"])
    gate = _gate_snapshot(database_path)
    return _with_readiness(
        {
            "success": True,
            "database_path": database_path,
            "pid": chosen,
            "create_time": target.get("create_time"),
            "process_name": target.get("process_name") or None,
            "identity_confirmed": target.get("identity_confirmed", False),
            "running": target.get("running"),
            "matched_by": target.get("matched_by"),
            "responsive": responsive,
            "gate_busy": gate["gate_busy"],
            "operation": gate["operation"],
            "note": (
                "ready means this process is confirmed as a running, responsive Access, "
                "is not in break mode, has no blocking dialog, and the MCP Access gate "
                "is free for this database. Dismissing a dialog does not retry the "
                "operation that was waiting."
            ),
        },
        _dialog_records(listed, chosen),
    )


def _find_dialog(report: dict[str, Any], dialog_id: str) -> dict[str, Any] | None:
    hwnd = parse_dialog_id(dialog_id)
    if hwnd is None:
        return None
    for item in report.get("dialogs") or []:
        if item.get("hwnd") == hwnd:
            return item
    return None


def _uncertain_click(before: dict[str, Any], item: dict[str, Any]) -> dict[str, Any]:
    return {
        **before,
        "success": False,
        "dismissed": False,
        "uncertain": True,
        "error_pattern": "dismiss_uncertain",
        "error": (
            "The click was not delivered before the timeout, or was refused because the "
            "button no longer belongs to the target process. It was not retried."
        ),
        "dialog_id": item.get("dialog_id"),
    }


def _windows_for_pid(windows: Iterable[WindowInfo], pid: int) -> dict[int, WindowInfo]:
    return {window.hwnd: window for window in windows if window.pid == pid}


def dismiss_dialog_in_windows(
    windows: list[WindowInfo],
    database_path: str,
    dialog_id: str,
    *,
    button: str | None = None,
    action: str = "close",
    pid: int | None = None,
    create_time: int | None = None,
    responsive: bool | None = None,
    backend: WindowBackend,
    click_timeout_sec: float | None = None,
) -> dict[str, Any]:
    """Click one button or close one add-in window belonging to the target pid."""
    before = inspect_windows(
        windows,
        database_path,
        backend=backend,
        pid=pid,
        create_time=create_time,
        responsive=responsive,
    )
    if not before.get("success"):
        return before
    item = _find_dialog(before, dialog_id)
    if item is None:
        return {
            "success": False,
            "error_pattern": "dialog_not_found",
            "error": f"No open dialog with id {dialog_id} on the target Access instance.",
            "pid": before.get("pid"),
            "dialogs": before.get("dialogs"),
        }
    refused = _identity_refusal(before)
    if refused:
        return refused
    if responsive is False:
        return {
            **before,
            "success": False,
            "error_pattern": "access_unresponsive",
            "error": "The window did not respond within the timeout. No button was clicked.",
            "dismissed": False,
        }

    kind = item["kind"]
    action_l = (action or "close").strip().lower()
    # The item was read from these windows, so its window is always here.
    window = _windows_for_pid(windows, int(before["pid"]))[int(item["hwnd"])]
    closed: list[dict[str, Any]] = []
    interrupted = False
    cancel_requested = False
    failure_dialog = kind in FAILURE_KINDS

    if button:
        if _button_label(button) == "debug":
            return {
                **before,
                "success": False,
                "error_pattern": "debug_refused",
                "error": "Debug is never selected. Pass button=End to stop a failed VBA call.",
                "dismissed": False,
            }
        if _button_named(window, button) is None:
            return {
                **before,
                "success": False,
                "error_pattern": "button_not_found",
                "error": f"Button {button!r} is not on that dialog.",
                "buttons": item.get("buttons"),
                "dismissed": False,
            }
        status, record, pressed = _press_fresh(
            backend,
            before,
            window,
            lambda fresh, _kind: _button_named(fresh, button),
            click_timeout_sec,
            database_path,
        )
        if status == "dialog_changed":
            return _dialog_changed(before, item)
        assert record is not None and pressed is not None
        if status == "dismiss_uncertain":
            return _uncertain_click(before, record)
        closed.append({**record, "button": pressed.text, "how": "click"})
        interrupted = _interrupts_execution(record["kind"], pressed.text)
    elif action_l == "cancel":
        if kind != "addin_window":
            return {
                **before,
                "success": False,
                "error_pattern": "unsupported_action",
                "error": "cancel closes an add-in progress window. For a VBA dialog pass button.",
                "dismissed": False,
            }
        record = _close_fresh(
            backend, before, window, database_path, "Add-in window close requested as cancellation."
        )
        if record is None:
            return _dialog_changed(before, item)
        closed.append({**record, "how": "cancel"})
        cancel_requested = True
    elif action_l == "close":
        if kind == "addin_window" and (before.get("operation") or {}).get("same_database"):
            return {
                **before,
                "success": False,
                "error_pattern": "operation_in_progress",
                "error": (
                    "This add-in window belongs to a running operation. "
                    "close is only for a finished results window. "
                    "Pass action=cancel to interrupt it."
                ),
                "dismissed": False,
                "interrupted": False,
            }
        if kind == "vba_break":
            return {
                **before,
                "success": False,
                "error_pattern": "vba_break",
                "error": "VBA is in break mode. That is not a dialog and was not closed.",
                "dismissed": False,
            }
        if kind != "addin_window":
            return {
                **before,
                "success": False,
                "error_pattern": "button_required",
                "error": "Pass button to choose a dialog button. close is for a finished add-in window.",
                "buttons": item.get("buttons"),
                "dismissed": False,
            }
        record = _close_fresh(backend, before, window, database_path)
        if record is None:
            return _dialog_changed(before, item)
        closed.append({**record, "how": "close"})
    else:
        return {
            **before,
            "success": False,
            "error_pattern": "unsupported_action",
            "error": "Supported actions are close and cancel. Pass button to click a dialog button.",
            "dismissed": False,
        }

    result = {
        "success": True,
        "dismissed": True,
        "closed": closed,
        "interrupted": interrupted,
        "failure_dialog_dismissed": failure_dialog,
        "pid": before.get("pid"),
        "create_time": before.get("create_time"),
        "operation": before.get("operation"),
        "note": (
            "The dialog action completed. If failure_dialog_dismissed or interrupted "
            "is true, the Access operation that was waiting did not succeed."
        ),
    }
    if cancel_requested:
        result["cancel_requested"] = True
        result["note"] = (
            "The close was posted to the add-in window as a cancel request. The add-in "
            "may ask to confirm, and keeps its window open while the operation runs. "
            "The waiting call's own result says whether it was cancelled."
        )
    return result


def recover_windows(
    windows: list[WindowInfo],
    database_path: str,
    *,
    policy: str = "report",
    pid: int | None = None,
    create_time: int | None = None,
    responsive: bool | None = None,
    backend: WindowBackend,
    click_timeout_sec: float | None = None,
) -> dict[str, Any]:
    """Apply a bounded automatic policy. Unknown dialogs are reported, not clicked."""
    policy_l = (policy or "report").strip().lower()
    if policy_l not in {"report", "safe", "end_runtime_error"}:
        return {
            "success": False,
            "error_pattern": "invalid_recovery_policy",
            "error": "policy must be report, safe, or end_runtime_error.",
        }
    report = inspect_windows(
        windows,
        database_path,
        backend=backend,
        pid=pid,
        create_time=create_time,
        responsive=responsive,
    )
    if not report.get("success"):
        return report
    if policy_l == "report":
        report["policy"] = policy_l
        report["automatic"] = False
        report["closed"] = []
        return report
    refused = _identity_refusal(report)
    if refused:
        refused["closed"] = []
        return refused
    if responsive is False:
        report["success"] = False
        report["error_pattern"] = "access_unresponsive"
        report["error"] = "The window did not respond. No dialog was clicked."
        report["closed"] = []
        return report

    closed: list[dict[str, Any]] = []
    skipped: list[dict[str, Any]] = []
    uncertain = False
    changed = False
    by_hwnd = _windows_for_pid(windows, int(report["pid"]))
    handled = 0
    for item in list(report["dialogs"]):
        if handled >= MAX_AUTO_DIALOGS:
            skipped.append({**item, "reason": "max_auto_dialogs"})
            continue
        window = by_hwnd.get(int(item["hwnd"]))
        if window is None:
            continue
        caption = auto_button(window, item["kind"], policy_l)
        if caption is None:
            skipped.append({
                **item,
                "reason": "not_covered_by_policy",
                "actions": item.get("buttons") or [],
            })
            continue
        if _button_named(window, caption) is None:
            skipped.append({**item, "reason": "button_missing"})
            continue
        # The policy runs again on the fresh window; its choice is what gets clicked.
        status, record, pressed = _press_fresh(
            backend,
            report,
            window,
            lambda fresh, kind: _policy_button(fresh, kind, policy_l),
            click_timeout_sec,
            database_path,
        )
        if status == "dialog_changed":
            changed = True
            skipped.append({**item, "reason": "dialog_changed"})
            continue
        assert record is not None and pressed is not None
        if status == "dismiss_uncertain":
            # Never retry an undelivered click, and stop touching this instance.
            skipped.append({**record, "reason": "dismiss_uncertain"})
            uncertain = True
            break
        handled += 1
        closed.append({**record, "button": pressed.text, "how": "auto"})

    result = {
        "success": True,
        "policy": policy_l,
        "automatic": True,
        "pid": report.get("pid"),
        "create_time": report.get("create_time"),
        "closed": closed,
        "skipped": skipped,
        "interrupted": any(_interrupts_execution(item["kind"], item.get("button")) for item in closed),
        "failure_dialog_dismissed": any(item.get("kind") in FAILURE_KINDS for item in closed),
        "operation": report.get("operation"),
        "note": (
            "Automatic recovery clicks only recognized dialogs covered by the policy. "
            "Skipped dialogs are unchanged. This does not retry the blocked operation."
        ),
    }
    if uncertain:
        result["uncertain"] = True
        result["error_pattern"] = "dismiss_uncertain"
    elif changed:
        result["error_pattern"] = "dialog_changed"
    return result


class Win32Backend:
    """Live window backend. Imported only when a tool actually enumerates windows."""

    def list_windows(self) -> list[WindowInfo]:
        import win32gui
        import win32process

        found: list[WindowInfo] = []
        access_pids: dict[int, bool] = {}

        def _is_access(pid: int) -> bool:
            if pid not in access_pids:
                access_pids[pid] = _process_basename(pid).lower() in ACCESS_EXE_NAMES
            return access_pids[pid]

        def _callback(hwnd: int, _extra: Any) -> bool:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            title = win32gui.GetWindowText(hwnd) or ""
            class_name = win32gui.GetClassName(hwnd) or ""
            _thread, pid = win32process.GetWindowThreadProcessId(hwnd)
            # Other Office apps draw NUIDialogs too; only Access's are read through MSAA.
            if class_name == NETUI_DIALOG_CLASS and _is_access(int(pid)):
                buttons, texts = _netui_contents(hwnd)
            else:
                buttons, texts = _child_contents(hwnd)
            found.append(
                WindowInfo(
                    hwnd=int(hwnd),
                    pid=int(pid),
                    title=title,
                    class_name=class_name,
                    texts=tuple(texts),
                    buttons=tuple(buttons),
                )
            )
            return True

        win32gui.EnumWindows(_callback, None)
        return found

    def process_identity(self, pid: int) -> ProcessIdentity:
        return ProcessIdentity(
            name=_process_basename(pid) or None,
            create_time=process_create_time(pid),
            running=_process_running(pid),
        )

    def click(
        self, button: ButtonInfo, *, expected_pid: int | None = None, timeout_ms: int = 5000
    ) -> str:
        import ctypes

        import win32process

        hwnd = int(button.hwnd)
        if expected_pid is not None:
            try:
                _thread, owner_pid = win32process.GetWindowThreadProcessId(hwnd)
            except Exception:
                return CLICK_NOT_SENT
            if int(owner_pid) != int(expected_pid):
                return CLICK_NOT_SENT
        if button.path is not None:
            # A hung window is not asked for its accessible object.
            if not self.responsive(hwnd, timeout_ms):
                return CLICK_NOT_SENT
            from . import msaa

            try:
                pressed = msaa.press(hwnd, button.path, button.text, timeout_ms)
            except OSError:
                return CLICK_UNCERTAIN
            # False: the button was not found as expected or refused the action.
            return CLICK_DELIVERED if pressed else CLICK_NOT_SENT
        result = ctypes.c_ulong()
        sent = ctypes.windll.user32.SendMessageTimeoutW(
            hwnd,
            BM_CLICK,
            0,
            0,
            SMTO_ABORTIFHUNG,
            int(timeout_ms),
            ctypes.byref(result),
        )
        # A timed-out send may still be processed later.
        return CLICK_DELIVERED if sent else CLICK_UNCERTAIN

    def close(self, hwnd: int) -> None:
        import win32gui

        win32gui.PostMessage(hwnd, WM_CLOSE, 0, 0)

    def responsive(self, hwnd: int, timeout_ms: int) -> bool:
        import ctypes

        result = ctypes.c_ulong()
        sent = ctypes.windll.user32.SendMessageTimeoutW(
            int(hwnd),
            WM_NULL,
            0,
            0,
            SMTO_ABORTIFHUNG,
            int(timeout_ms),
            ctypes.byref(result),
        )
        return bool(sent)


def _child_contents(hwnd: int) -> tuple[list[ButtonInfo], list[str]]:
    import win32gui

    buttons: list[ButtonInfo] = []
    texts: list[str] = []

    def _callback(child: int, _extra: Any) -> bool:
        class_name = win32gui.GetClassName(child) or ""
        text = win32gui.GetWindowText(child) or ""
        if text:
            texts.append(text)
        if class_name == "Button" and text:
            buttons.append(ButtonInfo(int(child), text))
        return True

    try:
        win32gui.EnumChildWindows(hwnd, _callback, None)
    except Exception:
        return buttons, texts
    return buttons, texts


def _netui_contents(hwnd: int) -> tuple[list[ButtonInfo], list[str]]:
    """Buttons and texts of a ``NUIDialog``, read through MSAA from its ``NetUIHWND`` hosts."""
    import win32gui

    from . import msaa

    hosts: list[int] = []

    def _callback(child: int, _extra: Any) -> bool:
        if (win32gui.GetClassName(child) or "") == NETUI_HOST_CLASS:
            hosts.append(int(child))
        return True

    buttons: list[ButtonInfo] = []
    texts: list[str] = []
    try:
        win32gui.EnumChildWindows(hwnd, _callback, None)
        for host in hosts:
            for control in msaa.read_controls(host, int(min(dialog_timeout_sec(), 2) * 1000)):
                if control.role == msaa.ROLE_PUSHBUTTON:
                    buttons.append(ButtonInfo(host, control.name, control.path))
                else:
                    texts.append(control.name)
    except Exception:
        return buttons, texts
    return buttons, texts


def _process_basename(pid: int) -> str:
    if pid <= 0:
        return ""
    try:
        import win32api
        import win32con
        import win32process

        handle = win32api.OpenProcess(win32con.PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        try:
            path = win32process.GetModuleFileNameEx(handle, 0)
        finally:
            win32api.CloseHandle(handle)
        return os.path.basename(path or "")
    except Exception:
        return ""


def _process_running(pid: int) -> bool | None:
    if pid <= 0:
        return False
    try:
        import win32api
        import win32con
        import win32process

        handle = win32api.OpenProcess(win32con.PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        try:
            return win32process.GetExitCodeProcess(handle) == 259  # STILL_ACTIVE
        finally:
            win32api.CloseHandle(handle)
    except Exception:
        return None


def _default_backend() -> WindowBackend:
    # Kept as a function: tests patch it to swap the live Win32 backend for a fake.
    return Win32Backend()


def _responsive_for(windows: list[WindowInfo], pid: int, backend: WindowBackend, timeout_sec: float) -> bool | None:
    hwnds = [window.hwnd for window in windows if window.pid == pid]
    if not hwnds:
        return None
    timeout_ms = max(200, int(min(timeout_sec, 2) * 1000))
    return any(backend.responsive(hwnd, timeout_ms) for hwnd in hwnds[:8])


def _open_ids(windows: Iterable[WindowInfo], pid: int, ids: set[str]) -> set[str]:
    return {dialog_id_for(w.hwnd) for w in windows if w.pid == pid} & ids


def _settle_and_reinspect(
    live: WindowBackend,
    database_path: str,
    preview: dict[str, Any],
    acted: set[str],
    timeout: float,
) -> tuple[dict[str, Any], set[str]]:
    """The post-action sequence shared by dismiss and recover: wait, then re-inspect.

    Polls until none of the dialogs this call acted on is still open, or the
    timeout elapses. Dialogs the call did not act on (for example ones the
    policy does not cover) never extend the wait. Then inspects once for the
    returned report. Returns that report and the acted-on dialogs still open at
    the deadline, which the caller reports as ``dismiss_uncertain``.
    """
    pid = int(preview["pid"])
    deadline = time.monotonic() + timeout
    windows = live.list_windows()
    while acted and _open_ids(windows, pid, acted) and time.monotonic() < deadline:
        time.sleep(POLL_INTERVAL_SEC)
        windows = live.list_windows()
    pending = _open_ids(windows, pid, acted)
    follow = inspect_windows(
        windows,
        database_path,
        backend=live,
        pid=pid,
        create_time=preview.get("create_time"),
        responsive=_responsive_for(windows, pid, live, timeout),
    )
    return _with_readiness(follow, _dialog_records(windows, pid)), pending


def list_dialogs(
    database_path: str,
    pid: int | None = None,
    create_time: int | None = None,
    timeout_seconds: float | None = None,
    backend: WindowBackend | None = None,
) -> dict[str, Any]:
    """Inspect dialogs for one database. Pass a backend in tests."""
    timeout = dialog_timeout_sec(timeout_seconds)
    live = backend or _default_backend()
    windows = live.list_windows()
    target_pid = pid
    if target_pid is None:
        preview = resolve_target(windows, database_path, live, pid=None, create_time=create_time)
        if not preview.get("success"):
            preview["timeout_seconds"] = timeout
            return preview
        target_pid = int(preview["pid"])
    responsive = _responsive_for(windows, int(target_pid), live, timeout)
    report = inspect_windows(
        windows,
        database_path,
        backend=live,
        pid=target_pid,
        create_time=create_time,
        responsive=responsive,
    )
    report["timeout_seconds"] = timeout
    return report


def dismiss_dialog(
    database_path: str,
    dialog_id: str,
    *,
    button: str | None = None,
    action: str = "close",
    pid: int | None = None,
    create_time: int | None = None,
    timeout_seconds: float | None = None,
    backend: WindowBackend | None = None,
) -> dict[str, Any]:
    """Dismiss one dialog, then report what is still open."""
    timeout = dialog_timeout_sec(timeout_seconds)
    live = backend or _default_backend()
    windows = live.list_windows()
    preview = resolve_target(windows, database_path, live, pid=pid, create_time=create_time)
    if not preview.get("success"):
        return preview
    responsive = _responsive_for(windows, int(preview["pid"]), live, timeout)
    result = dismiss_dialog_in_windows(
        windows,
        database_path,
        dialog_id,
        button=button,
        action=action,
        pid=int(preview["pid"]),
        create_time=preview.get("create_time"),
        responsive=responsive,
        backend=live,
        click_timeout_sec=timeout,
    )
    if not result.get("success"):
        result["timeout_seconds"] = timeout
        return result
    acted = {str(item["dialog_id"]) for item in result.get("closed") or []}
    follow, pending = _settle_and_reinspect(live, database_path, preview, acted, timeout)
    uncertain = bool(pending)
    follow.update({
        "success": not uncertain,
        "dismissed": not uncertain,
        "uncertain": uncertain,
        "closed": result.get("closed"),
        "interrupted": result.get("interrupted"),
        "failure_dialog_dismissed": result.get("failure_dialog_dismissed"),
        "timeout_seconds": timeout,
        "note": result.get("note"),
    })
    if result.get("cancel_requested"):
        follow["cancel_requested"] = True
    if uncertain:
        follow["error_pattern"] = "dismiss_uncertain"
        follow["error"] = "The dialog was still open when the timeout elapsed. The operation was not retried."
    return follow


def recover_dialogs(
    database_path: str,
    policy: str = "report",
    pid: int | None = None,
    create_time: int | None = None,
    timeout_seconds: float | None = None,
    backend: WindowBackend | None = None,
) -> dict[str, Any]:
    """Inspect, and optionally click recognized dialogs covered by ``policy``."""
    timeout = dialog_timeout_sec(timeout_seconds)
    live = backend or _default_backend()
    windows = live.list_windows()
    preview = resolve_target(windows, database_path, live, pid=pid, create_time=create_time)
    if not preview.get("success"):
        return preview
    responsive = _responsive_for(windows, int(preview["pid"]), live, timeout)
    result = recover_windows(
        windows,
        database_path,
        policy=policy,
        pid=int(preview["pid"]),
        create_time=preview.get("create_time"),
        responsive=responsive,
        backend=live,
        click_timeout_sec=timeout,
    )
    if policy.strip().lower() == "report" or not result.get("success"):
        result["timeout_seconds"] = timeout
        return result
    acted = {str(item["dialog_id"]) for item in result.get("closed") or []}
    follow, pending = _settle_and_reinspect(live, database_path, preview, acted, timeout)
    if result.get("uncertain") or pending:
        follow["success"] = False
        follow["uncertain"] = True
        follow["error_pattern"] = "dismiss_uncertain"
        follow["error"] = (
            "A click was not delivered, or a clicked dialog was still open when the "
            "timeout elapsed. Nothing was retried."
        )
    elif result.get("error_pattern") == "dialog_changed":
        follow["success"] = False
        follow["error_pattern"] = "dialog_changed"
        follow["error"] = (
            "A dialog changed after it was inspected and was not clicked. "
            "Re-inspect and decide again."
        )
    follow.update({
        "policy": result.get("policy"),
        "automatic": True,
        "closed": result.get("closed"),
        "skipped": result.get("skipped"),
        "interrupted": result.get("interrupted"),
        "failure_dialog_dismissed": result.get("failure_dialog_dismissed"),
        "timeout_seconds": timeout,
        "note": result.get("note"),
    })
    return follow


def automation_status(
    database_path: str,
    pid: int | None = None,
    create_time: int | None = None,
    timeout_seconds: float | None = None,
    backend: WindowBackend | None = None,
) -> dict[str, Any]:
    """Readiness check that does not click and does not use the Access COM gate."""
    report = list_dialogs(
        database_path,
        pid=pid,
        create_time=create_time,
        timeout_seconds=timeout_seconds,
        backend=backend,
    )
    if not report.get("success"):
        report["ready"] = False
        return report
    _ready, pattern = _readiness(
        identity_confirmed=bool(report.get("identity_confirmed")),
        running=report.get("running"),
        responsive=report.get("responsive"),
        dialogs=report.get("dialogs") or [],
        gate_busy_here=_gate_busy_here(
            {"gate_busy": report.get("gate_busy"), "operation": report.get("operation")}
        ),
    )
    if pattern:
        report["error_pattern"] = pattern
    if pattern in {"access_not_running", "no_windows_to_probe"}:
        report["success"] = False
        report["error"] = (
            "The Access process is not running. Relaunch it."
            if pattern == "access_not_running"
            else "The Access process has no windows to probe yet. Wait and check again."
        )
    return report


def reset_interruptions() -> None:
    """Clear recorded End/cancel diagnostics. Used by tests."""
    with _interruptions_changed:
        _interruptions.clear()
