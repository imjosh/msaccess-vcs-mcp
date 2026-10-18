"""
Access COM connection management.

Adapted from db-inspector-mcp with ownership tracking and cleanup patterns.
Manages COM connections to Access databases with proper resource cleanup.

Key Complexity Areas:
- Multiple connection strategies based on whether Access is already running
- Ownership tracking to avoid interfering with user's Access instances
- COM cleanup challenges and garbage collection issues
"""

import os
import re
import shutil
import tempfile
import sys
import threading
import time
import traceback
from dataclasses import asdict, dataclass, replace
from typing import Any, Callable

try:
    import pythoncom
    import win32com.client
    from win32com.client import gencache
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False

_GEN_PY_MODULE_RE = re.compile(r"win32com\.gen_py\.([^\s'\"]+)")


def _extract_gen_py_folder_from_error(message: str) -> str | None:
    """Return the gen_py type-library folder name embedded in an error message."""
    match = _GEN_PY_MODULE_RE.search(message)
    return match.group(1) if match else None


def _gen_py_folder_incomplete(folder_name: str) -> bool:
    """True when a gen_py folder exists but lacks the wrapper package files."""
    if not COM_AVAILABLE:
        return False
    folder_path = os.path.join(gencache.GetGeneratePath(), folder_name)
    if not os.path.isdir(folder_path):
        return False
    return not os.path.isfile(os.path.join(folder_path, "__init__.py"))


def _should_heal_gen_py_cache(exc: BaseException) -> bool:
    """Detect a stale pywin32 gen_py wrapper that can be rebuilt safely."""
    if not isinstance(exc, AttributeError):
        return False
    message = str(exc)
    if "win32com.gen_py." not in message:
        return False
    if "CLSIDToClassMap" in message:
        return True
    folder = _extract_gen_py_folder_from_error(message)
    return bool(folder and _gen_py_folder_incomplete(folder))


def _purge_gen_py_cache_folder(folder_name: str, *, prog_id: str) -> None:
    """Delete one corrupted gen_py type-library folder and drop cached imports."""
    from ..usage_logging import log_diagnostic_event

    folder_path = os.path.join(gencache.GetGeneratePath(), folder_name)
    removed = False
    if os.path.isdir(folder_path):
        shutil.rmtree(folder_path, ignore_errors=True)
        removed = True

    module_prefix = f"win32com.gen_py.{folder_name}"
    for name in list(sys.modules):
        if name == module_prefix or name.startswith(f"{module_prefix}."):
            del sys.modules[name]

    try:
        gencache.Rebuild()
    except Exception:
        pass

    log_diagnostic_event(
        "gen_py_cache_rebuilt",
        prog_id=prog_id,
        folder=folder_name,
        path=folder_path,
        removed=removed,
    )
    print(
        f"Rebuilt pywin32 gen_py cache for {prog_id} "
        f"(removed stale folder {folder_name})",
        file=sys.stderr,
    )


def ensure_dispatch(prog_id: str):
    """Early-bound COM dispatch with one-shot gen_py cache self-heal.

    ``gencache.EnsureDispatch`` can fail when a type-library folder under
    ``%TEMP%\\gen_py`` is left in a half-built state (for example only
    ``__pycache__`` remains).  Delete that folder, rebuild the cache, and
    retry once so MCP startup does not die on a recoverable local error.
    """
    if not COM_AVAILABLE:
        raise ImportError(
            "pywin32 is required for COM automation. "
            "Install it with: pip install pywin32"
        )

    for attempt in range(2):
        try:
            return gencache.EnsureDispatch(prog_id)
        except AttributeError as exc:
            if attempt == 0 and _should_heal_gen_py_cache(exc):
                folder = _extract_gen_py_folder_from_error(str(exc))
                if folder:
                    _purge_gen_py_cache_folder(folder, prog_id=prog_id)
                    continue
            raise


def create_isolated_access_app():
    """Start a new Access COM server, never attach to an existing window.

    Keep DispatchEx's native creation semantics, but retain the raw object while
    repairing a corrupt generated wrapper. Retry wrapping once on that same
    object; never create a second host or fall back to EnsureDispatch.
    """
    if not COM_AVAILABLE:
        raise ImportError("pywin32 is required for Access COM automation")
    dispatch = pythoncom.CoCreateInstanceEx(
        "Access.Application", None, pythoncom.CLSCTX_SERVER,
        None, (pythoncom.IID_IDispatch,),
    )[0]
    for attempt in range(2):
        try:
            return win32com.client.Dispatch(
                dispatch, "Access.Application", clsctx=pythoncom.CLSCTX_SERVER,
            )
        except AttributeError as exc:
            if attempt or not _should_heal_gen_py_cache(exc):
                raise
            folder = _extract_gen_py_folder_from_error(str(exc))
            typelib, _ = dispatch.GetTypeInfo().GetContainingTypeLib()
            attributes = typelib.GetLibAttr()
            identity = attributes[0], attributes[1], attributes[3], attributes[4]
            if folder != gencache.GetGeneratedFileName(*identity):
                # Only the native object's own type-library cache may be removed.
                raise
            _purge_gen_py_cache_folder(folder, prog_id="Access.Application")
            if gencache.EnsureModule(*identity) is None:
                raise RuntimeError("Access type-library wrapper generation unavailable")


def _paths_match(a: str, b: str) -> bool:
    """Case-insensitive, normalised path comparison."""
    try:
        return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))
    except (OSError, ValueError):
        return False


def leave_access_open() -> bool:
    """True unless ACCESS_VCS_LEAVE_ACCESS_OPEN is an explicit false.

    Default is to leave a COM-created Access window running so later
    calls reattach instead of spawning a new process. Set the env var
    to ``false`` / ``0`` / ``no`` to restore quit-per-call.
    """
    raw = os.environ.get("ACCESS_VCS_LEAVE_ACCESS_OPEN", "true").strip().lower()
    return raw not in ("0", "false", "no")


def ensure_access_visible(app) -> bool:
    """Show the Access window for any instance we drive a database through.

    A hidden instance strands the user: an error dialog, a VBA break, or a
    trust prompt blocks every later call with nothing on screen to explain
    why, and nobody can dismiss what they cannot see.  Visibility is
    therefore not optional for database work -- it applies to instances we
    created and to ones we attached to, since either can raise a dialog.

    Call this *after* the database is open.  Making the window visible can
    set ``UserControl``, and startup code reads that flag to decide whether
    a person is watching (see ``open_current_database``).

    Best-effort by design: failing to show the window is never worth failing
    an operation over.
    """
    try:
        app.Visible = True
        return True
    except Exception as e:
        print(f"Could not make Access visible: {e}", file=sys.stderr)
        return False


def open_current_database(app, db_path: str) -> None:
    """Open ``db_path`` as ``app``'s current database, as automation.

    ``OpenCurrentDatabase`` runs the target's AutoExec.  Startup code
    commonly branches on ``Application.UserControl`` to decide whether a
    person is watching -- the VCS add-in's own ``AutoRun`` opens its
    installer form when it believes one is, which strands the instance we
    are about to automate.  The flag gets set both deliberately (see
    ``AccessConnection._create_isolated_instance``, which needs the process
    to outlive a client teardown) and as a side effect of making the window
    visible, so lower it across the open and restore it afterwards.

    Use this in place of a bare ``OpenCurrentDatabase`` call: it also leaves
    the instance visible, which is the invariant every database-holding
    instance owes the user.
    """
    try:
        was_set = bool(app.UserControl)
    except Exception:
        was_set = False

    if was_set:
        try:
            app.UserControl = False
        except Exception:
            was_set = False
    try:
        app.OpenCurrentDatabase(db_path)
    finally:
        if was_set:
            try:
                app.UserControl = True
            except Exception:
                pass
    ensure_access_visible(app)


def access_instance_is_live(db_path: str) -> bool:
    """True when an Access process already has ``db_path`` open.

    Uses the Running Object Table only -- it does not bind a file moniker
    and therefore cannot spawn a new instance.
    """
    from ..addin_integration import VCSAddinIntegration

    return VCSAddinIntegration._find_access_in_rot(db_path) is not None


DEFAULT_CLOSE_TIMEOUT_SEC = 10.0


def get_close_timeout() -> float:
    """Seconds to wait for a graceful Quit before force-terminating."""
    try:
        value = float(
            os.environ.get(
                "ACCESS_VCS_CLOSE_TIMEOUT_SEC", str(DEFAULT_CLOSE_TIMEOUT_SEC)
            )
        )
    except ValueError:
        return DEFAULT_CLOSE_TIMEOUT_SEC
    return value if value > 0 else DEFAULT_CLOSE_TIMEOUT_SEC


def _resolve_instance_for_record(record) -> Any:
    """Return the COM object for ``record``, or None if it is not that process.

    Both lookups resolve by path, which can land on a *different* Access
    instance than the one recorded -- two processes can hold the same file,
    and a moniker bind can even start a new one. The PID and creation stamp
    are therefore re-checked before the caller is allowed to close anything.
    """
    from .instance_registry import process_create_time
    from .process_qos import pid_from_access_app

    if not COM_AVAILABLE:
        return None

    app = None
    try:
        app = win32com.client.GetObject(record.database_path)
    except Exception:
        from ..addin_integration import VCSAddinIntegration

        app = VCSAddinIntegration._find_access_in_rot(record.database_path)
    if app is None:
        return None

    pid = pid_from_access_app(app)
    if pid != record.pid:
        return None
    if process_create_time(pid) != record.create_time:
        return None
    return app


@dataclass(frozen=True)
class _CloseEvidence:
    quit_returned_at: float | None = None
    apartment_released: bool = False
    worker_completed: bool = False
    process_exited: bool = False
    force_requested: bool = False
    force_accepted: bool = False
    errors: tuple[str, ...] = ()


class _QuitAttempt:
    def __init__(self, app):
        self.evidence = _CloseEvidence()
        self._lock = threading.Lock()
        self._handoff = {}
        self.thread = None
        try:
            self._handoff['stream'] = pythoncom.CoMarshalInterThreadInterfaceInStream(
                pythoncom.IID_IDispatch, app._oleobj_)
        except Exception as error:
            self._error(error)

    def _update(self, **fields):
        with self._lock:
            self.evidence = replace(self.evidence, **fields)

    def _error(self, error):
        with self._lock:
            self.evidence = replace(self.evidence, errors=self.evidence.errors +
                                    (f"{type(error).__name__}: {error}",))

    def _release_unused(self):
        stream = self._handoff.pop('stream', None)
        if stream is not None:
            try:
                pythoncom.CoReleaseMarshalData(stream)
            except Exception as error:
                self._error(error)
                return False
        return True

    def start(self):
        if not self._handoff:
            self._update(worker_completed=True, apartment_released=True)
            return
        self.thread = threading.Thread(target=self._worker, daemon=True, name='vcs-access-quit')
        try:
            self.thread.start()
        except Exception as error:
            self._error(error)
            self.thread = None
            released = self._release_unused()
            self._update(worker_completed=True, apartment_released=released)

    def _worker(self):
        worker_app = interface = stream = None
        initialized = False
        released = False
        try:
            pythoncom.CoInitialize()
            initialized = True
            stream = self._handoff.pop('stream')
            interface = pythoncom.CoGetInterfaceAndReleaseStream(stream, pythoncom.IID_IDispatch)
            stream = None
            worker_app = win32com.client.Dispatch(interface)
            interface = None
            try:
                worker_app.CloseCurrentDatabase()
            except Exception as error:
                self._error(error)
            worker_app.Quit()
            self._update(quit_returned_at=time.monotonic())
        except Exception as error:
            self._error(error)
        finally:
            worker_app = interface = stream = None
            if initialized:
                try:
                    pythoncom.CoUninitialize()
                    released = True
                except Exception as error:
                    self._error(error)
            else:
                released = True
            self._update(apartment_released=released)

    def join(self, timeout):
        if self.thread is not None:
            self.thread.join(timeout=max(0.0, timeout))
            if self.thread.is_alive():
                return False
        released = self._release_unused()
        self._update(worker_completed=True,
                     apartment_released=self.evidence.apartment_released and released)
        return self.evidence.apartment_released


def _quit_with_timeout(app, timeout_sec: float) -> bool:
    if not COM_AVAILABLE:
        return False
    attempt = _QuitAttempt(app)
    app = None
    attempt.start()
    return attempt.join(timeout_sec) and attempt.evidence.quit_returned_at is not None


@dataclass
class _PendingClose:
    target: str
    record: Any
    process: Any = None
    attempt: _QuitAttempt | None = None
    admission_worker: Any = None


_pending_closes: dict[tuple[str, int, int | None], _PendingClose] = {}
_recycle_closed: dict[str, list[dict[str, Any]]] = {}
NATURAL_EXIT_SECONDS = 60.0
FORCE_SETTLEMENT_SECONDS = 5.0


def _pending_for(target):
    return [pending for pending in _pending_closes.values() if _paths_match(pending.target, target)]


def owned_recovery_pending(target: str) -> bool:
    return bool(_pending_for(target) or
                os.path.normcase(os.path.abspath(target)) in _recycle_closed)


def _close_owned_record(record, *, target=None, app=None) -> dict[str, Any] | None:
    from .instance_registry import unregister_owned
    from .process_qos import open_exact_process

    target = target or record.database_path
    key = (os.path.normcase(os.path.abspath(target)), record.pid, record.create_time)
    pending = _pending_closes.get(key)
    if pending is None:
        pending = _PendingClose(target, record)
        _pending_closes[key] = pending
    if pending.admission_worker is not None:
        if pending.admission_worker.is_alive():
            return None
        pending.admission_worker = None
    if pending.process is None:
        pending.process = open_exact_process(record.pid, record.create_time)
        if pending.process is None:
            return None
    evidence = _CloseEvidence(process_exited=True)
    if pending.attempt is None and pending.process.exited() is not True:
        if app is None:
            app = _resolve_instance_for_record(record)
        if app is None:
            return None
        pending.attempt = _QuitAttempt(app)
        app = None
        pending.attempt.start()
    if pending.attempt is not None:
        attempt = pending.attempt
        attempt.join(get_close_timeout())
        returned = attempt.evidence.quit_returned_at
        if returned is not None and not attempt.evidence.force_requested:
            while time.monotonic() < returned + NATURAL_EXIT_SECONDS:
                exited = pending.process.exited()
                if exited is not False:
                    break
                attempt.join(min(0.05, returned + NATURAL_EXIT_SECONDS - time.monotonic()))
                time.sleep(0.05)
        exited = pending.process.exited()
        if exited is None:
            return None
        if not exited and not attempt.evidence.force_requested:
            accepted = pending.process.force()
            attempt._update(force_requested=True, force_accepted=accepted)
        deadline = time.monotonic() + FORCE_SETTLEMENT_SECONDS
        while True:
            settled = attempt.join(0)
            exited = pending.process.exited()
            if exited is True and settled:
                break
            if exited is None or time.monotonic() >= deadline:
                return None
            attempt.join(min(0.05, deadline - time.monotonic()))
            time.sleep(0.05)
        evidence = replace(attempt.evidence, process_exited=True)
    pending.process.close()
    del _pending_closes[key]
    unregister_owned(record.pid, record.create_time)
    result = {"pid": record.pid, "database_path": record.database_path,
              "quit": evidence.quit_returned_at is not None,
              "terminated": evidence.force_accepted, "evidence": asdict(evidence)}
    _log_instance_event('owned_instance_closed', **result)
    return result


def _log_instance_event(event: str, **fields: Any) -> None:
    try:
        from ..usage_logging import log_diagnostic_event

        log_diagnostic_event(event, **fields)
    except Exception:
        pass


def close_owned_instances_holding(
    paths: list[str] | tuple[str, ...],
    addin_paths: list[str] | tuple[str, ...] = (),
) -> list[dict[str, Any]]:
    """Quit server-created Access instances holding any of the given files.

    ``paths`` matches an instance's open database; ``addin_paths`` also
    matches a loaded add-in library, which locks its file even when the
    instance has an unrelated database open. User-owned windows are never
    touched. Used before rebuilds that must replace a file, and when
    recycling a stuck owned instance.
    """
    from .instance_registry import owned_records_for_paths

    closed: list[dict[str, Any]] = []
    for record in owned_records_for_paths(paths, addin_paths):
        result = _close_owned_record(record)
        if result is not None:
            closed.append(result)
    return closed


def recycle_owned_instance(database_path: str, *, settle_predecessor: Callable[[], bool] | None = None) -> bool:
    """Replace only after every exact host and predecessor worker has settled."""
    from .instance_registry import owned_records_for_paths

    records = {(r.pid, r.create_time): r for r in owned_records_for_paths([database_path])}
    for pending in _pending_for(database_path):
        records[(pending.record.pid, pending.record.create_time)] = pending.record
    target_key = os.path.normcase(os.path.abspath(database_path))
    closed = _recycle_closed.setdefault(target_key, [])
    if not records and not closed:
        _recycle_closed.pop(target_key, None)
        return False
    results = [_close_owned_record(record, target=database_path) for record in records.values()]
    closed.extend(result for result in results if result is not None)
    if any(result is None for result in results):
        return False
    if settle_predecessor is not None and not settle_predecessor():
        return False
    _recycle_closed.pop(target_key, None)
    try:
        with AccessConnection(database_path):
            pass
    except Exception as error:
        _recycle_closed[target_key] = closed
        _log_instance_event('owned_instance_recycle_refused', database=database_path,
                            error=str(error), closed=closed)
        return False
    _recycle_closed.pop(target_key, None)
    return True


class AccessConnection:
    """
    Manages COM connection to Access database.
    
    Uses ownership tracking to determine cleanup responsibility:
    - If connecting to user's existing Access instance, we don't close it
    - If we create our own Access instance, we're responsible for cleanup
      (quit only when ACCESS_VCS_LEAVE_ACCESS_OPEN is explicitly false, or
      when a rebuild / recycle must replace the file)
    """
    
    def __init__(self, db_path: str):
        """
        Initialize Access COM connection manager.
        
        Args:
            db_path: Path to Access database file (.accdb, .accda, .mdb)
        """
        if not COM_AVAILABLE:
            raise ImportError(
                "pywin32 is required for COM automation. "
                "Install it with: pip install pywin32"
            )
        
        self._db_path = db_path
        self._app = None
        self._db = None
        
        # Ownership tracking flags - determine cleanup responsibility
        self._owns_app = False  # True only if we created Access via Dispatch
        self._owns_db = False   # True only if we opened db via DBEngine
        self._db_opened_via_getobject = False  # True if connected to existing instance
        self._db_opened_as_current = False  # True if we called OpenCurrentDatabase
    
    def _get_access_app(self):
        """
        Get or create Access COM application.
        
        Uses a multi-step approach:
        1. Try GetObject(path) to connect to our database if already open
        2. Try EnsureDispatch -- but check whether it returned an existing
           user instance (has a different database open) vs a fresh one
        3. If EnsureDispatch attached to a user instance with a different
           DB, fall back to DispatchEx to create a guaranteed-isolated
           Access process
        
        IMPORTANT: Access can only have ONE database open at a time.
        - If user has OUR database open -> connect to their instance
        - If user has DIFFERENT database open -> create our OWN isolated instance
        - If no Access running -> create new instance
        
        The ownership check after EnsureDispatch prevents the dangerous
        scenario where close() calls Quit() on the user's Access window
        (ported from db-inspector-mcp's DispatchEx fallback pattern).

        Whichever route wins, the instance ends up visible: a dialog or a
        VBA break in a hidden window is unresolvable by the person who has
        to resolve it.  That happens last, once the database is open, for
        the ``UserControl`` reason in ``open_current_database``.  Instances
        we created also get ``UserControl = True`` afterwards so the window
        is a normal interactive Access app, not an automation ghost.
        """
        if self._app is None:
            if owned_recovery_pending(self._db_path):
                raise RuntimeError("Owned recovery is still pending for this target")
            from ..compatibility_session import target_admission_path

            library = target_admission_path()
            if library is not None:
                return self._get_admitted_access_app(library)
            created_now = False
            try:
                # A moniker bind *launches* Access when nothing has the file
                # open, so whether an instance was already running decides
                # whether the process below is ours. The ROT lookup only
                # inspects the table -- it cannot start anything itself.
                was_running = access_instance_is_live(self._db_path)
                self._app = win32com.client.GetObject(self._db_path)
                self._db_opened_via_getobject = True
                created_now = not was_running
            except Exception:
                self._app, created_now = self._create_or_reuse_instance()
                self._db_opened_via_getobject = False
                # Never displace a user's database: only open on instances
                # we just created (empty or isolated).
                self._owns_app = created_now
                self._open_as_current_database(self._app)
            self._resolve_ownership(created_now)
            if self._owns_app:
                from .process_qos import prefer_full_power_app

                prefer_full_power_app(self._app)
            ensure_access_visible(self._app)
            if self._owns_app:
                self._ensure_owned_instance_interactive()
        return self._app

    def _get_admitted_access_app(self, library):
        """Admit in a neutral host before opening a previously closed target.

        ROT lookup only attaches to an already open target. Never bind a file
        moniker here: that can launch Access and run AutoExec before admission.
        The loaded library survives CloseCurrentDatabase on the neutral host.
        """
        from ..addin_integration import VCSAddinIntegration
        from ..compatibility_session import ensure_session
        from .instance_registry import OwnedInstance, process_create_time
        from .process_qos import pid_from_access_app, prefer_full_power_app

        if not os.path.isfile(self._db_path):
            raise FileNotFoundError(self._db_path)
        existing = VCSAddinIntegration._find_access_in_rot(self._db_path)
        if existing is not None:
            VCSAddinIntegration(library).load_addin(existing, db_path=self._db_path)
            self._app = existing
            self._db_opened_via_getobject = True
            self._resolve_ownership(False)
            ensure_access_visible(existing)
            return existing

        app = create_isolated_access_app()
        pid = pid_from_access_app(app)
        created = process_create_time(pid) if pid else None
        if not created:
            # No target or bootstrap has been opened; identity loss fails closed.
            from ..compatibility_session import AdmissionError, session_failure
            app = None
            raise AdmissionError(session_failure("access_identity_unconfirmed"))
        self._app = app
        self._resolve_ownership(True)
        prefer_full_power_app(app)
        folder = tempfile.mkdtemp(prefix="vcs-admission-")
        neutral = os.path.join(folder, "bootstrap.accdb")
        try:
            app.UserControl = False
            app.NewCurrentDatabase(neutral)
            ensure_access_visible(app)
            VCSAddinIntegration(library).load_addin(app)
            app.CloseCurrentDatabase()
            # This is validation only, before any target startup code executes.
            ensure_session(app, library)
            open_current_database(app, self._db_path)
            self._db_opened_as_current = True
            self._ensure_owned_instance_interactive()
            return app
        except Exception as error:
            failure = error
            timed_out = isinstance(error, TimeoutError)
            chain = [error]
            seen = set()
            while chain:
                item = chain.pop()
                if id(item) in seen:
                    continue
                seen.add(id(item))
                chain.extend(e for e in (item.__context__, item.__cause__) if e is not None)
                traceback.clear_frames(item.__traceback__)
                item.__traceback__ = item.__context__ = item.__cause__ = None
            record = OwnedInstance(pid, neutral, created)
            self._app = None
            self._owns_app = False
            if timed_out:
                from .instance_registry import register_owned
                register_owned(pid, neutral, create_time=created)
                key = (os.path.normcase(os.path.abspath(self._db_path)), pid, created)
                _pending_closes[key] = _PendingClose(
                    self._db_path, record,
                    admission_worker=VCSAddinIntegration._active_probe_thread)
                app = None
            else:
                from .process_qos import open_exact_process
                process = open_exact_process(pid, created)
                if process is not None:
                    attempt = _QuitAttempt(app)
                    key = (os.path.normcase(os.path.abspath(self._db_path)), pid, created)
                    _pending_closes[key] = _PendingClose(self._db_path, record, process, attempt)
                    app = None
                    attempt.start()
                    try:
                        _close_owned_record(record, target=self._db_path)
                    except Exception as cleanup_error:
                        _log_instance_event('bootstrap_cleanup_deferred', pid=pid,
                                            error=str(cleanup_error))
                else:
                    app = None
                    key = (os.path.normcase(os.path.abspath(self._db_path)), pid, created)
                    _pending_closes[key] = _PendingClose(self._db_path, record)
            raise failure from None
        finally:
            # Exact files in our mkdtemp directory; no recursive deletion.
            for path in (neutral, os.path.splitext(neutral)[0] + ".laccdb"):
                try:
                    os.unlink(path)
                except FileNotFoundError:
                    pass
                except OSError as exc:
                    _log_instance_event("bootstrap_cleanup_deferred", path=path, error=str(exc))
            try:
                os.rmdir(folder)
            except OSError:
                pass

    def _resolve_ownership(self, created_now: bool) -> None:
        """Set ``_owns_app`` from creation or the durable PID registry.

        GetObject on a left-open window would otherwise reclassify a
        server-created instance as user-owned. The registry survives
        process restarts; create_time rejects PID reuse.
        """
        from .instance_registry import is_owned, process_create_time, register_owned
        from .process_qos import pid_from_access_app

        pid = pid_from_access_app(self._app)
        create_time = process_create_time(pid) if pid else None
        if created_now:
            self._owns_app = True
            if pid:
                register_owned(pid, self._db_path, create_time=create_time)
            return
        self._owns_app = bool(pid and is_owned(pid, create_time))

    def _ensure_owned_instance_interactive(self) -> None:
        """Give a COM-created Access instance a normal interactive window.

        ``Visible = True`` alone often leaves an automation-created process
        out of the desktop. ``UserControl`` after the database is open does
        not retrigger AutoExec -- that already ran with the flag down.
        Best-effort: never fail the operation over this.
        """
        try:
            self._app.UserControl = True
        except Exception as e:
            print(f"Could not set Access UserControl: {e}", file=sys.stderr)

    def _open_as_current_database(self, app):
        """Make ``self._db_path`` the instance's current database.

        ``GetObject(path)`` binds a file moniker, which resolves through the
        COM registration for the file's extension.  That registration opens
        .accdb and .mdb as the current database, but not .accda -- Access
        treats it as an add-in and moniker binding fails.  Without this
        fallback the instance has no current database, so add-in calls and
        the Running Object Table lookup in ``_find_access_in_rot`` (which
        matches on ``CurrentDb().Name``) both come up empty.

        A failed open is not fatal: leaving ``_db_opened_as_current`` False
        lets ``_get_current_db`` fall through to its DAO strategies, which
        still serve read-only callers when Access itself cannot take the
        file as the current database.
        """
        try:
            existing = app.CurrentDb()
        except Exception:
            existing = None

        if existing is not None:
            if _paths_match(existing.Name, self._db_path):
                self._db_opened_as_current = True
                return
            # A different database is open.  _create_or_reuse_instance
            # routes that case to an isolated process, so this only happens
            # on an instance we own; never displace a user's database.
            if not self._owns_app:
                return

        try:
            open_current_database(app, self._db_path)
            self._db_opened_as_current = True
        except Exception as e:
            print(
                f"[{os.path.basename(self._db_path)}] OpenCurrentDatabase "
                f"failed ({e}) -- falling back to DAO",
                file=sys.stderr,
            )

    def _create_or_reuse_instance(self):
        """Create or attach to an Access instance with correct ownership.

        EnsureDispatch("Access.Application") can silently return an
        existing user-owned instance instead of creating a new one.
        Returns ``(app, created_now)`` so the caller can register a new
        process or look the PID up in the durable ownership registry.

        Creation is decided by comparing the resulting PID against a
        snapshot taken beforehand, not by whether the instance has a
        database open. A user can leave Access sitting at an empty shell
        with no current database, and that window is still theirs.
        """
        from .process_qos import list_access_pids_or_none, pid_from_access_app

        before = list_access_pids_or_none()
        app = ensure_dispatch("Access.Application")

        try:
            existing_db = app.CurrentDb()
        except Exception:
            existing_db = None

        if existing_db is not None:
            if _paths_match(existing_db.Name, self._db_path):
                return app, False
            db_name = os.path.basename(self._db_path)
            other_name = os.path.basename(existing_db.Name)
            print(
                f"[{db_name}] EnsureDispatch returned instance with "
                f"'{other_name}' open -- launching isolated instance via DispatchEx",
                file=sys.stderr,
            )
            return self._create_isolated_instance(), True

        if before is None:
            # Cannot prove we created it; treat as the user's.
            return app, False
        pid = pid_from_access_app(app)
        return app, bool(pid and pid not in before)

    def _create_isolated_instance(self):
        """Create a guaranteed-isolated Access process via DispatchEx.

        Unlike EnsureDispatch, DispatchEx always spawns a new COM server
        process independent of any existing Access instance.  ``UserControl``
        keeps that process alive if the client goes away mid-operation; see
        ``open_current_database`` for why it has to be lowered again while a
        database opens.
        """
        app = create_isolated_access_app()
        app.UserControl = True
        return app
    
    def _get_current_db(self):
        """
        Get database object for DAO operations.
        
        Strategy:
        1. If database was opened via GetObject (Access already had it open), use CurrentDb()
        2. Otherwise, use DBEngine.OpenDatabase() which is more reliable
        
        IMPORTANT: This method tracks whether we OWN the database connection:
        - If we use CurrentDb() (user's database), we do NOT own it
        - If we open via DBEngine, we OWN it and are responsible for closing
        """
        if self._db is None:
            app = self._get_access_app()
            
            # If Access already had the database open, CurrentDb() should work
            if self._db_opened_via_getobject or self._db_opened_as_current:
                try:
                    db = app.CurrentDb()
                    if db is not None:
                        self._db = db
                        self._owns_db = False  # User's database - do NOT close it
                        return self._db
                except Exception:
                    pass
            
            # Use DBEngine.OpenDatabase() - more reliable
            try:
                dbe = app.DBEngine
                # Open database in shared mode (Exclusive=False, ReadOnly=False)
                self._db = dbe.OpenDatabase(self._db_path, False, False)
                self._owns_db = True  # We opened this - we're responsible for closing
            except Exception:
                # Try read-only mode
                try:
                    dbe = app.DBEngine
                    self._db = dbe.OpenDatabase(self._db_path, False, True)
                    self._owns_db = True
                except Exception:
                    # Last resort: try direct DAO without Access
                    try:
                        dbe = win32com.client.Dispatch("DAO.DBEngine.120")
                        self._db = dbe.OpenDatabase(self._db_path, False, True)
                        self._owns_db = True
                    except Exception as e:
                        raise RuntimeError(
                            f"Failed to open database '{self._db_path}' via COM. "
                            f"Ensure the database file exists and is not corrupted. Error: {e}"
                        )
        
        return self._db
    
    def connect(self):
        """
        Connect to Access database via COM.
        
        Returns:
            Tuple of (Application, Database) objects
        """
        app = self._get_access_app()
        db = self._get_current_db()
        return app, db
    
    def get_app(self):
        """Get Access Application object."""
        return self._get_access_app()
    
    def get_db(self):
        """Get DAO Database object."""
        return self._get_current_db()
    
    def close(self):
        """
        Close the database connection and cleanup COM resources.
        
        IMPORTANT: This method respects ownership:
        - If we connected to an existing Access instance (via GetObject), we do NOT
          close Access or the database - the user is still using them!
        - If we created our own Access instance, we leave it running by default
          so later calls reattach. Set ACCESS_VCS_LEAVE_ACCESS_OPEN=false to
          restore quit-per-call. Rebuild and recycle close owned instances
          when they must replace the file.
        """
        # Only close the database if we opened it ourselves
        if self._db is not None and self._owns_db:
            try:
                self._db.Close()
            except Exception:
                pass
        self._db = None
        
        # Only quit Access if we created it ourselves AND the operator
        # opted back into quit-per-call.
        if self._app is not None and self._owns_app and not leave_access_open():
            from .instance_registry import process_create_time, unregister_owned
            from .process_qos import pid_from_access_app, process_is_alive

            pid = pid_from_access_app(self._app)
            create_time = process_create_time(pid) if pid else None
            try:
                self._app.CloseCurrentDatabase()
            except Exception:
                pass
            try:
                self._app.Quit()
            except Exception:
                pass
            # Only drop the claim once the process is confirmed gone. A
            # failed Quit that also forgot the PID would leave a live
            # instance nobody can close or recycle later.
            if pid and process_is_alive(pid) is False:
                unregister_owned(pid, create_time)
        self._app = None
        
        # Reset ownership flags
        self._owns_app = False
        self._owns_db = False
        self._db_opened_via_getobject = False
        self._db_opened_as_current = False
    
    def __enter__(self):
        """Context manager entry."""
        self.connect()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        self.close()
        return False
