"""X17 mutual admission. No operation is replayed after a dispatch/COM error."""
from __future__ import annotations

import contextvars
import json
import threading
import uuid
from collections import OrderedDict
from contextlib import contextmanager

from . import __version__
from .compatibility import ADDIN_REQUIREMENT, _installation_identity, compatibility_result

PROTOCOL = "msaccess-vcs.session/1"
SERVER_INSTANCE = str(uuid.uuid4())
_local_connection = str(uuid.uuid4())
_connection = contextvars.ContextVar("compatibility_connection", default=_local_connection)
_target_library = contextvars.ContextVar("compatibility_target_library", default=None)
_connections = OrderedDict()
_sessions = OrderedDict()
_lock = threading.RLock()
MAX_SESSIONS = 128


@contextmanager
def target_admission(path):
    """Require loaded mutual admission before a connection opens its target."""
    token = _target_library.set(path)
    try:
        yield
    finally:
        _target_library.reset(token)


def target_admission_path():
    return _target_library.get()


class AdmissionError(RuntimeError):
    def __init__(self, result):
        self.result = result
        super().__init__(result["error"])


def session_failure(reason, detail=None):
    action = "Reconnect and complete a mutual handshake before dispatch; do not replay an uncertain operation."
    return dict(success=False, error_pattern="compatibility_session_invalid", component="session",
                installed_version=None, required_range=ADDIN_REQUIREMENT.text,
                minimum_version=ADDIN_REQUIREMENT.minimum, compatibility_reason=reason,
                requirement_status="assigned_unpublished", started=False,
                error=detail or f"Compatibility session refused: {reason}. {action}", recovery_action=action)


@contextmanager
def connection_scope(connection):
    """Keep a strong reference so object/PID reuse cannot resurrect a connection."""
    with _lock:
        if connection is None:
            identity = _local_connection
        else:
            key = id(connection)
            saved = _connections.get(key)
            if saved is None or saved[0] is not connection:
                saved = (connection, str(uuid.uuid4()))
                _connections[key] = saved
            _connections.move_to_end(key)
            while len(_connections) > MAX_SESSIONS:
                _, (_, expired) = _connections.popitem(last=False)
                for session_key in list(_sessions):
                    if session_key[0] == expired:
                        del _sessions[session_key]
            identity = saved[1]
    token = _connection.set(identity)
    try:
        yield
    finally:
        _connection.reset(token)


def _reply(app, library, procedure, *args):
    raw = app.Run(library + "." + procedure, *args)
    if isinstance(raw, tuple):
        raw = raw[0]
    value = json.loads(raw) if isinstance(raw, str) else raw
    if not isinstance(value, dict):
        raise AdmissionError(session_failure("invalid_handshake_reply"))
    return value


def ensure_session(app, path):
    """Cheap validation first; handshake only after certain non-execution.

    The cache contains strings/identities, never an apartment-bound COM proxy.
    Failure of file/process/loaded identity detection fails closed.
    """
    from .access_com.process_qos import pid_from_access_app
    from .access_com.instance_registry import process_create_time
    import os

    try:
        pid = pid_from_access_app(app)
        created = process_create_time(pid) if pid else None
        if not created:
            raise AdmissionError(session_failure("access_identity_unconfirmed"))
        installed = _installation_identity(path)
        key = (_connection.get(), SERVER_INSTANCE, pid, created, installed,
               ADDIN_REQUIREMENT.text, PROTOCOL)
        library = os.path.splitext(os.path.abspath(path))[0]
        with _lock:
            cached = _sessions.get(key)
        if cached:
            verdict = _reply(app, library, "APIValidateSession", cached)
            if verdict.get("success") is True:
                return cached
            # Validation-only: execution has definitely not been dispatched.
            with _lock:
                _sessions.pop(key, None)
        identity = _reply(app, library, "APIIdentity")
        if identity.get("success") is False:
            raise AdmissionError(identity)
        admission = compatibility_result(identity.get("addin_version"), ADDIN_REQUIREMENT)
        if not admission["success"]:
            raise AdmissionError(admission)
        if identity.get("protocol") != PROTOCOL or not identity.get("addin_instance"):
            raise AdmissionError(session_failure("protocol_or_instance_unconfirmed"))
        request = dict(server_instance=SERVER_INSTANCE, connection_id=_connection.get(),
                       server_version=__version__, protocol=PROTOCOL,
                       addin_instance=identity["addin_instance"], session_id=str(uuid.uuid4()))
        accepted = _reply(app, library, "APIHandshake", json.dumps(request))
        if accepted.get("success") is not True:
            raise AdmissionError(accepted)
        if any(accepted.get(k) != request[k] for k in
               ("session_id", "server_instance", "connection_id", "addin_instance", "protocol")):
            raise AdmissionError(session_failure("handshake_identity_mismatch"))
        if _installation_identity(path) != installed:
            raise AdmissionError(session_failure("installation_changed_during_handshake"))
        envelope = json.dumps({k: request[k] for k in
                               ("session_id", "server_instance", "connection_id", "addin_instance")})
        with _lock:
            _sessions[key] = envelope
            _sessions.move_to_end(key)
            while len(_sessions) > MAX_SESSIONS:
                _sessions.popitem(last=False)
        return envelope
    except AdmissionError:
        raise
    except Exception as exc:
        raise AdmissionError(session_failure("identity_or_handshake_unavailable", str(exc))) from exc
