from __future__ import annotations

import shutil
import tempfile
import threading
import time
from pathlib import Path

from .access_gate import _retain_build_host
from .access_com import instance_registry as registry
from .access_com.connection import get_close_timeout, NATURAL_EXIT_SECONDS, FORCE_SETTLEMENT_SECONDS
from .access_com.process_qos import pid_from_access_app, open_exact_process

_observers = set()
_observers_lock = threading.Lock()
_exiting = threading.Event()


def stop_observers(timeout: float = 1.0) -> list[str]:
    """Stop reconciliation observers; return operation IDs still running.

    A stopped observer performs no further cleanup. The durable build claim,
    retained process and published cleanup state stay as they were, so a
    later server reports the host as owned rather than cleaned up.
    """
    with _observers_lock:
        owners = list(_observers)
    for owner in owners:
        owner._stop.set()
    deadline = time.monotonic() + timeout
    running = []
    for owner in owners:
        owner._observer.join(max(0.0, deadline - time.monotonic()))
        if owner._observer.is_alive():
            running.append(owner.operation_id)
    return running


def _interpreter_exit():
    _exiting.set()
    stop_observers()


# Interpreter exit waits for the gate's call in flight before atexit
# handlers run, and a pending build's retire() runs inside that call. These
# hooks run newest first, so this one runs before the gate's wait, which
# access_gate registered when it was imported above.
threading._register_atexit(_interpreter_exit)


class _BuildOwner:
    def __init__(self, manager, operation_id):
        self.manager = manager
        self.operation_id = operation_id
        self.app = None
        self.process = None
        self.record = None
        self.host_dir = None
        self.outcome = None
        self.dispatched = False
        self.retired = False
        self.reconciling = False
        self.response = None
        self.cleanup = {"state": "preparing"}
        # Pinned so a late observer never resolves a different registry.
        self._registry_path = registry.get_registry_path()
        self._stop = threading.Event()
        self._observer = None
        self._finished = False
        self._rmtree = shutil.rmtree
        self.operation = manager.attach_build_owner(operation_id, self.cleanup)
        try:
            _retain_build_host(self.retire)
        except Exception:
            # Refused before any host exists, so nothing needs cleanup.
            manager.update_build_cleanup(operation_id, {"state": "complete"}, release=True)
            raise

    def capture(self, app, database_path):
        self.app = app
        pid = pid_from_access_app(app)
        stamp = registry.process_create_time(pid) if pid else None
        self.process = open_exact_process(pid, stamp) if pid else None
        if self.process is None:
            raise RuntimeError("Build refused because exact host identity could not be pinned")
        try:
            self.record = registry.register_build_owned(
                pid, database_path, stamp, self.operation_id, self._registry_path
            )
        except Exception as error:
            # The write may have landed before its confirmation failed.
            try:
                registry.remove_build_owned(pid, stamp, self._registry_path, self.operation_id)
            except Exception:
                pass
            raise RuntimeError(f"Build refused because host ownership could not be persisted: {error}") from error
        self._publish("prepared")

    def open_host(self):
        self.host_dir = tempfile.mkdtemp(prefix="vcs-build-host-")
        self.record = registry._record_build_directory(self.record, self.host_dir, self._registry_path)
        self.app.NewCurrentDatabase(str(Path(self.host_dir) / "BuildHost.accdb"))

    def describe(self, response):
        response["operation_id"] = self.operation_id
        response["cleanup"] = dict(self.cleanup)
        if not self._finished:
            # Kept current until finish(); a later result is a snapshot that
            # no other thread rewrites after it is returned.
            self.response = response
        return response

    def _publish(self, state, **fields):
        self.cleanup = {"state": state, **fields}
        if self.record:
            self.cleanup.update(pid=self.record.pid, create_time=self.record.create_time)
        self.manager.update_build_cleanup(self.operation_id, self.cleanup)
        if self.response is not None:
            self.response["cleanup"] = dict(self.cleanup)

    def finish(self):
        if not self.dispatched or (self.outcome is not None and self.outcome.kind in {"terminal", "refused"}):
            self.retire()
        else:
            self.manager.detach_waiter(self.operation_id)
            self._publish("pending")
        self.response = None
        self._finished = True

    def retire(self):
        if self.retired or self.reconciling:
            return
        try:
            while (self.dispatched and not self.operation.settlement.is_set()
                   and (self.outcome is None or self.outcome.kind not in {"terminal", "refused"})):
                if _exiting.is_set():
                    # Leave the claim and the pending state for the next server.
                    return
                if self.process is not None and self.process.exited() is True:
                    break
                self.operation.settlement.wait(0.05)
            if self.app is None and not self.dispatched:
                self._publish("complete")
                self.manager.update_build_cleanup(self.operation_id, self.cleanup, release=True)
                self.retired = True
                return
            if self.process is None:
                # Identity is pinned before the host opens anything or a build
                # is dispatched, so this host never held the target.
                try:
                    self.app.Quit()
                except Exception:
                    pass
                self._publish("recovery_required", error="Exact process identity unavailable; Quit was requested but exit is unconfirmed")
                self.manager.update_build_cleanup(self.operation_id, self.cleanup, release=True)
                self.retired = True
                return
            self._publish("closing")
            if self.process.exited() is not True:
                self._close_settled_host()
            if self.process.exited() is not True:
                self._publish("recovery_required", error="Exact host exit was not confirmed")
                return
            self.app = None
            self._remove_artifacts()
        except Exception as error:
            self._publish("cleanup_failed", error=str(error))
        finally:
            self.app = None
            if self.process is not None and not self.retired and not _exiting.is_set():
                self.reconciling = True
                self._observer = threading.Thread(target=self._reconcile, name="vcs-build-reconcile", daemon=True)
                with _observers_lock:
                    _observers.add(self)
                self._observer.start()

    def _remove_artifacts(self):
        if self.host_dir:
            folder = Path(self.host_dir)
            for attempt in range(10):
                try:
                    self._rmtree(folder)
                    break
                except FileNotFoundError:
                    break
                except OSError:
                    if attempt == 9:
                        raise RuntimeError(f"Generated host directory could not be removed: {folder}")
                    time.sleep(0.2)
        if self.record:
            registry.remove_build_owned(self.record.pid, self.record.create_time, self._registry_path)
        self._publish("complete")
        self.manager.update_build_cleanup(self.operation_id, self.cleanup, release=True)
        self.process.close()
        self.process = None
        self.retired = True

    def _reconcile(self):
        try:
            while not self.retired and not self._stop.wait(0.1):
                try:
                    if self.process.exited() is True:
                        self._remove_artifacts()
                except Exception as error:
                    self._publish("cleanup_failed", error=str(error))
                    self._stop.wait(0.9)
        finally:
            with _observers_lock:
                _observers.discard(self)

    def _close_settled_host(self):
        finished = threading.Event()
        quit_returned = threading.Event()
        process = self.process

        def bound_cleanup():
            if not quit_returned.wait(get_close_timeout()):
                if not finished.is_set() and process.exited() is False:
                    process.force()
            elif not finished.wait(NATURAL_EXIT_SECONDS):
                if process.exited() is False:
                    process.force()
            if not finished.wait(FORCE_SETTLEMENT_SECONDS):
                self._publish("recovery_required", error="Cleanup has not settled after its bound")

        watchdog = threading.Thread(target=bound_cleanup, name="vcs-build-cleanup-bound", daemon=True)
        watchdog.start()
        try:
            try:
                self.app.CloseCurrentDatabase()
            except Exception:
                pass
            try:
                self.app.Quit()
            except Exception:
                pass
            else:
                quit_returned.set()
            self.app = None
            grace = NATURAL_EXIT_SECONDS if quit_returned.is_set() else get_close_timeout()
            deadline = time.monotonic() + grace + FORCE_SETTLEMENT_SECONDS
            while process.exited() is False and time.monotonic() < deadline:
                time.sleep(0.05)
        finally:
            finished.set()
            quit_returned.set()
            watchdog.join(get_close_timeout())
