"""Inspect and dismiss Access, VBA, and add-in dialogs from outside the Access process.

Access COM calls block while a modal dialog is open or VBA is in break mode.
This module uses Win32 window enumeration and button messages so recovery can
run while another MCP request is waiting on that COM call. It never sends
keystrokes or clicks by screen coordinate.

A dialog tool result describes what was found and what was clicked. Dismissing
an error dialog does not make the failed Access operation a success.
"""

from __future__ import annotations

import os
import time
from dataclasses import dataclass
from typing import Any, Iterable, Protocol

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
ADDIN_CAPTIONS = {"msaccessvcs", "version control system"}
FAILURE_KINDS = {"vba_runtime_error", "vba_compile_error"}
# A kind other than "unknown" is assigned only by a positive signature (see
# classify_window). Automatic recovery never clicks "unknown".
UNKNOWN_KIND = "unknown"
# A standard dialog whose only actionable button is OK, under any caption: a VBA
# MsgBox with a custom title. Positive signature is the button set alone.
MSGBOX_KIND = "vba_msgbox"
BLOCKING_KINDS = {
    "vba_runtime_error",
    "vba_compile_error",
    "access_dialog",
    MSGBOX_KIND,
    UNKNOWN_KIND,
    "vba_break",
}
# Kinds whose OK-only form the ``safe`` policy may acknowledge.
SAFE_OK_KINDS = {"vba_compile_error", "access_dialog", MSGBOX_KIND}
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

# pid -> last time we ended or reset execution from outside Access
_interruptions: dict[int, dict[str, Any]] = {}


@dataclass(frozen=True)
class ButtonInfo:
    hwnd: int
    text: str


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

    def click(self, hwnd: int, *, expected_pid: int | None = None, timeout_ms: int = 5000) -> bool:
        """Send BM_CLICK. True only when the message was handled within ``timeout_ms``."""
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
    """Return a dialog kind, or ``ignored`` for ordinary Access windows.

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
        return "vba_break"
    if "run-time error" in blob or "runtime error" in blob:
        return "vba_runtime_error"
    if "compile error" in blob:
        return "vba_compile_error"
    if "end" in buttons and "debug" in buttons:
        return "vba_runtime_error"

    caption_key = title_l.strip()
    if caption_key in ADDIN_CAPTIONS or caption_key.startswith("msaccessvcs"):
        return "addin_window"

    if class_name == "#32770":
        if title_l.startswith("microsoft access"):
            return "access_dialog"
        if title_l.startswith("microsoft visual basic"):
            return "access_dialog"
        # Help is not a choice, so OK + Help is still a single-button box.
        actionable = {label for label in buttons if label != "help"}
        if actionable == {"ok"} and len(_actionable_buttons(window)) == 1:
            return MSGBOX_KIND
        return UNKNOWN_KIND
    return "ignored"


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
    single-button ``vba_msgbox``. End is used only for a
    runtime-error dialog when the policy is ``end_runtime_error``. An
    ``unknown`` dialog is never clicked.
    """
    if policy in {"", "report"}:
        return None
    if kind != "vba_runtime_error" and kind not in SAFE_OK_KINDS:
        return None
    if _destructive_text(window):
        return None
    if policy not in {"safe", "end_runtime_error"}:
        return None
    if kind == "vba_runtime_error" and policy == "end_runtime_error":
        if _button_named(window, "End") is not None:
            return "End"
        return None
    # OK-only, ignoring a Help button that we do not click.
    actionable = [button.text.strip() for button in _actionable_buttons(window)]
    if len(actionable) == 1 and _button_label(actionable[0]) == "ok":
        return actionable[0]
    return None


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
    owned_pids: set[int] = set()
    try:
        wanted = os.path.normcase(os.path.abspath(database_path))
        for record in list_owned():
            if os.path.normcase(os.path.abspath(record.database_path)) == wanted:
                owned_pids.add(record.pid)
    except Exception:
        owned_pids = set()
    dead_owned = {item for item in owned_pids if _ident(item).running is False}
    owned_pids = {item for item in owned_pids if _ident(item).is_access}

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


def _still_same_target(
    backend: WindowBackend,
    report: dict[str, Any],
    dialog_hwnd: int,
    button_hwnd: int | None = None,
) -> bool:
    """Re-list windows right before acting: same process, same dialog, same button."""
    pid = int(report["pid"])
    ident = backend.process_identity(pid)
    if not ident.confirmed or ident.create_time != report.get("create_time"):
        return False
    for window in backend.list_windows():
        if window.hwnd != dialog_hwnd:
            continue
        if window.pid != pid:
            return False
        return button_hwnd is None or any(b.hwnd == button_hwnd for b in window.buttons)
    return False


def _dialog_changed(report: dict[str, Any], item: dict[str, Any]) -> dict[str, Any]:
    return {
        **report,
        "success": False,
        "dismissed": False,
        "error_pattern": "dialog_changed",
        "error": "The dialog or its process changed since it was inspected. Nothing was clicked. Re-inspect and decide again.",
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
    if any(item.get("kind") == "vba_break" for item in dialogs):
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
        "is_dialog": kind not in {"vba_break", "addin_window", "ignored"},
        "title": window.title,
        "message": message,
        "buttons": [button.text for button in window.buttons],
        "class_name": window.class_name,
        "pid": window.pid,
    }


def _gate_snapshot(database_path: str | None) -> dict[str, Any]:
    current = get_access_gate().current_in_flight()
    if current is None:
        return {"gate_busy": False, "operation": None}
    same = False
    if database_path and current.database:
        same = os.path.normcase(os.path.abspath(database_path)) == os.path.normcase(
            os.path.abspath(current.database)
        )
    return {
        "gate_busy": True,
        "operation": {
            "tool": current.tool,
            "database": current.database,
            "elapsed_ms": round((time.perf_counter() - current.started_at) * 1000, 2),
            "same_database": same,
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
    dialogs = []
    for window in listed:
        if window.pid != chosen:
            continue
        kind = classify_window(window)
        if kind == "ignored":
            continue
        dialogs.append(_dialog_record(window, kind))

    break_mode = any(item["kind"] == "vba_break" for item in dialogs)
    blocking = [item for item in dialogs if item["kind"] in BLOCKING_KINDS]
    gate = _gate_snapshot(database_path)
    ready, _reason = _readiness(
        identity_confirmed=bool(target.get("identity_confirmed")),
        running=target.get("running"),
        responsive=responsive,
        dialogs=dialogs,
        gate_busy_here=_gate_busy_here(gate),
    )
    interruption = _interruptions.get(chosen)
    return {
        "success": True,
        "database_path": database_path,
        "pid": chosen,
        "create_time": target.get("create_time"),
        "process_name": target.get("process_name") or None,
        "identity_confirmed": target.get("identity_confirmed", False),
        "running": target.get("running"),
        "matched_by": target.get("matched_by"),
        "dialogs": dialogs,
        "break_mode": break_mode,
        "blocking_dialog": bool(blocking),
        "responsive": responsive,
        "gate_busy": gate["gate_busy"],
        "operation": gate["operation"],
        "ready": ready,
        "execution_interrupted": interruption is not None,
        "last_interruption": interruption,
        "note": (
            "ready means this process is confirmed as a running, responsive Access, "
            "is not in break mode, has no blocking dialog, and the MCP Access gate "
            "is free for this database. Dismissing a dialog does not retry the "
            "operation that was waiting."
        ),
    }


def _find_dialog(report: dict[str, Any], dialog_id: str) -> dict[str, Any] | None:
    hwnd = parse_dialog_id(dialog_id)
    if hwnd is None:
        return None
    for item in report.get("dialogs") or []:
        if item.get("hwnd") == hwnd:
            return item
    return None


def _click_timeout_ms(click_timeout_sec: float | None) -> int:
    return int(dialog_timeout_sec(click_timeout_sec) * 1000)


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


def dismiss_one(
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
    by_hwnd = _windows_for_pid(windows, int(before["pid"]))
    window = by_hwnd.get(int(item["hwnd"]))
    closed: list[dict[str, Any]] = []
    interrupted = False
    failure_dialog = kind in FAILURE_KINDS or (
        kind == "access_dialog" and "error" in (item.get("message") or "").lower()
    )

    if button:
        if _button_label(button) == "debug":
            return {
                **before,
                "success": False,
                "error_pattern": "debug_refused",
                "error": "Debug is never selected. Pass button=End to stop a failed VBA call.",
                "dismissed": False,
            }
        if window is None or _button_named(window, button) is None:
            return {
                **before,
                "success": False,
                "error_pattern": "button_not_found",
                "error": f"Button {button!r} is not on that dialog.",
                "buttons": item.get("buttons"),
                "dismissed": False,
            }
        target_button = _button_named(window, button)
        assert target_button is not None
        if not _still_same_target(backend, before, int(item["hwnd"]), target_button.hwnd):
            return _dialog_changed(before, item)
        if not backend.click(
            target_button.hwnd,
            expected_pid=int(before["pid"]),
            timeout_ms=_click_timeout_ms(click_timeout_sec),
        ):
            return _uncertain_click(before, item)
        closed.append({**item, "button": target_button.text, "how": "click"})
        if _button_label(target_button.text) == "end" or failure_dialog:
            interrupted = True
            _interruptions[int(before["pid"])] = {
                "dialog_id": dialog_id,
                "kind": kind,
                "button": target_button.text,
                "title": item.get("title"),
                "message": item.get("message"),
            }
    elif action_l == "cancel":
        if kind != "addin_window":
            return {
                **before,
                "success": False,
                "error_pattern": "unsupported_action",
                "error": "cancel closes an add-in progress window. For a VBA dialog pass button.",
                "dismissed": False,
            }
        if not _still_same_target(backend, before, int(item["hwnd"])):
            return _dialog_changed(before, item)
        backend.close(int(item["hwnd"]))
        closed.append({**item, "how": "cancel"})
        interrupted = True
        _interruptions[int(before["pid"])] = {
            "dialog_id": dialog_id,
            "kind": kind,
            "button": None,
            "title": item.get("title"),
            "message": "Add-in window close requested as cancellation.",
        }
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
        if not _still_same_target(backend, before, int(item["hwnd"])):
            return _dialog_changed(before, item)
        backend.close(int(item["hwnd"]))
        closed.append({**item, "how": "close"})
    else:
        return {
            **before,
            "success": False,
            "error_pattern": "unsupported_action",
            "error": "Supported actions are close and cancel. Pass button to click a dialog button.",
            "dismissed": False,
        }

    return {
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
    if policy_l in {"report"}:
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
        button = _button_named(window, caption)
        if button is None:
            skipped.append({**item, "reason": "button_missing"})
            continue
        if not _still_same_target(backend, report, int(item["hwnd"]), button.hwnd):
            skipped.append({**item, "reason": "dialog_changed"})
            continue
        if not backend.click(
            button.hwnd,
            expected_pid=int(report["pid"]),
            timeout_ms=_click_timeout_ms(click_timeout_sec),
        ):
            # Never retry an undelivered click, and stop touching this instance.
            skipped.append({**item, "reason": "dismiss_uncertain"})
            uncertain = True
            break
        handled += 1
        record = {**item, "button": button.text, "how": "auto"}
        closed.append(record)
        if item["kind"] in FAILURE_KINDS or _button_label(button.text) == "end":
            _interruptions[int(report["pid"])] = {
                "dialog_id": item["dialog_id"],
                "kind": item["kind"],
                "button": button.text,
                "title": item.get("title"),
                "message": item.get("message"),
            }

    result = {
        "success": True,
        "policy": policy_l,
        "automatic": True,
        "pid": report.get("pid"),
        "create_time": report.get("create_time"),
        "closed": closed,
        "skipped": skipped,
        "interrupted": any(
            _button_label(item.get("button") or "") == "end" or item.get("kind") in FAILURE_KINDS
            for item in closed
        ),
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
    return result


def note_ready(report: dict[str, Any], remaining: list[dict[str, Any]]) -> dict[str, Any]:
    """Recompute readiness after a click, using the dialogs that are still open."""
    break_mode = any(item.get("kind") == "vba_break" for item in remaining)
    blocking = [item for item in remaining if item.get("kind") in BLOCKING_KINDS]
    pid = report.get("pid")
    interruption = _interruptions.get(int(pid)) if pid else None
    ready, _reason = _readiness(
        identity_confirmed=bool(report.get("identity_confirmed")),
        running=report.get("running"),
        responsive=report.get("responsive"),
        dialogs=remaining,
        gate_busy_here=_gate_busy_here(
            {"gate_busy": report.get("gate_busy"), "operation": report.get("operation")}
        ),
    )
    report = dict(report)
    report["dialogs"] = remaining
    report["break_mode"] = break_mode
    report["blocking_dialog"] = bool(blocking)
    report["ready"] = ready
    report["execution_interrupted"] = interruption is not None
    report["last_interruption"] = interruption
    return report


def filter_remaining(after_windows: Iterable[WindowInfo], pid: int) -> list[dict[str, Any]]:
    remaining = []
    for window in after_windows:
        if window.pid != pid:
            continue
        kind = classify_window(window)
        if kind == "ignored":
            continue
        remaining.append(_dialog_record(window, kind))
    return remaining


class Win32Backend:
    """Live window backend. Imported only when a tool actually enumerates windows."""

    def list_windows(self) -> list[WindowInfo]:
        import win32gui
        import win32process

        found: list[WindowInfo] = []

        def _callback(hwnd: int, _extra: Any) -> bool:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            title = win32gui.GetWindowText(hwnd) or ""
            class_name = win32gui.GetClassName(hwnd) or ""
            _thread, pid = win32process.GetWindowThreadProcessId(hwnd)
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

    def click(self, hwnd: int, *, expected_pid: int | None = None, timeout_ms: int = 5000) -> bool:
        import ctypes

        import win32process

        if expected_pid is not None:
            try:
                _thread, owner_pid = win32process.GetWindowThreadProcessId(int(hwnd))
            except Exception:
                return False
            if int(owner_pid) != int(expected_pid):
                return False
        result = ctypes.c_ulong()
        sent = ctypes.windll.user32.SendMessageTimeoutW(
            int(hwnd),
            BM_CLICK,
            0,
            0,
            SMTO_ABORTIFHUNG,
            int(timeout_ms),
            ctypes.byref(result),
        )
        return bool(sent)

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
    return note_ready(follow, filter_remaining(windows, pid)), pending


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
    result = dismiss_one(
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
    _interruptions.clear()
