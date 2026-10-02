"""
Integration with MSAccess VCS Add-in via COM automation.

This module provides a lightweight wrapper around the MSAccess VCS add-in,
delegating all export/import/build operations to the battle-tested add-in
rather than reimplementing them in Python.
"""

import json
import os
import threading
import time
from pathlib import Path
from typing import Any, Optional

try:
    import pythoncom
    import win32com.client
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False

from .decision_policy import is_decision_required, normalize_terminal_result
from .usage_logging import log_addin_probe

# Prefix the add-in puts on a refused (re-entrant) call. Keep in sync with
# modAPI.API_REFUSED_PREFIX.
API_REFUSED_PREFIX = "VCS_API_REFUSED: "
# Same meaning as the add-in's own refusal of a nested start: another API
# command is still running.
API_REFUSED_PATTERN = "operation_already_running"
# The add-in's self-dispatch refusal: the call arrived back in the project that
# sent it. This is an add-in defect, so waiting and retrying cannot help.
API_SELF_DISPATCH_PATTERN = "api_self_dispatch"
_SELF_DISPATCH_TEXT = "arrived back in the project that sent it"


def api_refusal_payload(value: Any) -> dict[str, Any] | None:
    """The structured failure for a dispatcher refusal string, else None."""
    if isinstance(value, str) and value.startswith(API_REFUSED_PREFIX):
        text = value[len(API_REFUSED_PREFIX):]
        return {
            "success": False,
            "error": text,
            "error_pattern": (
                API_SELF_DISPATCH_PATTERN if _SELF_DISPATCH_TEXT in text else API_REFUSED_PATTERN
            ),
            "api_refused": True,
        }
    return None


_UNCONFIRMED_START_ERROR = (
    "The {operation} started, but no completion callback exists on this path to "
    "report its outcome. It may have succeeded or failed. Read log_path, or "
    "call vcs_get_recent_calls() and vcs_get_log(log_type=\"{log_type}\")."
)


def unconfirmed_start_result(operation: str, log_type: str) -> dict[str, Any]:
    """The result of a start whose outcome this path cannot follow.

    Same shape as the merge's (M15): neither a success nor a failure, no
    ``error_pattern``, so the caller must not retry blindly.
    """
    return {
        "success": False,
        "started": True,
        "completion_unconfirmed": True,
        "error": _UNCONFIRMED_START_ERROR.format(operation=operation, log_type=log_type),
    }


def start_only_result(raw: Any, operation: str, log_type: str) -> dict[str, Any]:
    """Interpret what a start-only API call (Export, FullExport, ExportVBA, Build) returned.

    These add-in methods are Subs or form starts, so a normal return is Empty and
    says nothing about the outcome. Only an explicit failure is reported as one:
    a dispatcher refusal, or a JSON object with ``success: false`` or an
    undecided ``decision_required``, which keeps its ``error_pattern``. Anything
    else, including a ``success: true`` the method has no contract to give, is an
    unconfirmed start. This never returns ``success: true``.
    """
    refusal = api_refusal_payload(raw)
    if refusal:
        return refusal
    value = raw
    if isinstance(value, str) and value.strip():
        try:
            value = json.loads(value)
        except json.JSONDecodeError:
            value = None
    if isinstance(value, dict):
        wrapped = api_refusal_payload(value.get("error"))
        if wrapped:
            result = {**value, **wrapped}
        elif value.get("success") is False or is_decision_required(value):
            result = normalize_terminal_result(
                {**value, "success": False}, f"The {operation} failed"
            )
            if "started" in value:
                result["started"] = value["started"]
        else:
            result = unconfirmed_start_result(operation, log_type)
        # Log metadata is independent of the verdict. A start-only return can
        # identify its own log even when its completion remains unconfirmed.
        log_path = value.get("log_path") or value.get("logPath")
        if log_path:
            result["log_path"] = log_path
            if "logPath" in value:
                result["logPath"] = log_path
        return result
    return unconfirmed_start_result(operation, log_type)


# APICapabilities names this when BuildAs(source, output) builds with no picker and
# its completion callback reports output_path. A rebuild does not change the
# add-in's version, so the version cannot say whether the build has it.
# APICapabilities is a module procedure called by name, not an API method: Access
# refuses a missing procedure at once (error 2517), while API on a missing method
# stops in a modal "Run-time error '438'" inside Access.
CAPABILITIES_PROCEDURE = "APICapabilities"
BUILD_AS_PATHS = "build_as_paths"
BUILD_OUTPUT_UNSUPPORTED_PATTERN = "build_output_unsupported"


def build_output_unsupported_result(detail: str) -> dict[str, Any]:
    """The refusal for an add-in that cannot build to a given output path."""
    return {
        "success": False,
        "error_pattern": BUILD_OUTPUT_UNSUPPORTED_PATTERN,
        "error": (
            f"The add-in did not confirm the {BUILD_AS_PATHS} capability ({detail}); "
            "no build started. Building to an output path needs an add-in whose "
            "APICapabilities lists it, so that BuildAs takes the source folder and "
            "output file without opening a picker. Upgrade the add-in."
        ),
    }


def get_access_info(app) -> dict[str, Any]:
    """
    Get Access application version and bitness.
    
    Args:
        app: Access Application COM object
        
    Returns:
        Dictionary with access_version and bitness
    """
    try:
        # Get Access version (e.g., "16.0", "15.0", "14.0")
        access_version = app.Version
        
        # Detect bitness - Access itself reports if it's 64-bit
        # We can check the build number or system architecture
        try:
            # Try to check if we're running in 64-bit mode
            # In 64-bit Access, VBE version is typically 7.1, in 32-bit it's 7.0
            import platform
            import sys
            
            # Check Python's architecture (which should match Access if they're compatible)
            is_64bit = sys.maxsize > 2**32
            bitness = "64-bit" if is_64bit else "32-bit"
        except Exception:
            bitness = "unknown"
        
        return {
            "access_version": access_version,
            "bitness": bitness
        }
    except Exception as e:
        return {
            "access_version": "unknown",
            "bitness": "unknown",
            "error": str(e)
        }


class VCSAddinIntegration:
    """
    Integration layer for MSAccess VCS add-in.
    
    This class handles:
    - Loading the VCS add-in into Access
    - Calling add-in API functions via Application.Run
    - Translating between MCP and add-in formats
    - Parsing results from add-in operations
    """
    
    # Class-level guard against zombie probe threads piling up if Access
    # stops responding (VBA break mode, modal dialog, true hang).  Per-
    # instance state would be useless: every tool call constructs a fresh
    # VCSAddinIntegration, so a hung probe on one instance is invisible to
    # the next instance unless we track it on the class.
    _active_probe_thread: Optional[threading.Thread] = None
    
    def __init__(self, addin_path: Optional[str] = None):
        """
        Initialize add-in integration.
        
        Args:
            addin_path: Path to VCS add-in file (.accda). If None, uses default location.
        """
        if not COM_AVAILABLE:
            raise ImportError(
                "pywin32 is required for COM automation. "
                "Install it with: pip install pywin32"
            )
        
        self.addin_path = addin_path or self._get_default_addin_path()
        self._app = None
        self._addin_loaded = False
    
    def _get_default_addin_path(self) -> str:
        """
        Get default VCS add-in installation path.
        
        Returns:
            Path to add-in file (may not exist)
        """
        # Default installation location: %AppData%\MSAccessVCS\Version Control.accda
        appdata = os.environ.get("APPDATA", "")
        return os.path.join(appdata, "MSAccessVCS", "Version Control.accda")
    
    def verify_addin_exists(self) -> bool:
        """
        Check if add-in file exists at configured path.
        
        Returns:
            True if add-in file exists
        """
        return os.path.isfile(self.addin_path)
    
    def load_addin(self, app, db_path: Optional[str] = None) -> bool:
        """
        Verify add-in is accessible via the new API method.
        
        Probes the add-in by calling ``GetVCSVersion`` through
        ``Application.Run`` with a hard timeout, surfacing dialog-blocked,
        VBA-break-mode, or hung Access instances as a clear lifecycle error
        before any real work is dispatched.  Idempotent: a second call on
        the same instance is a no-op once the probe has succeeded.
        
        Args:
            app: Access Application COM object (with a database open).
            db_path: Optional path to the target database.  When provided,
                two things change:
                  * a fast ``os.path.isfile`` pre-flight catches stale or
                    typo paths in ~1ms instead of burning the full timeout;
                  * the worker thread re-acquires Access via the Running
                    Object Table for proper cross-apartment timeout
                    enforcement.  When ``None``, the worker falls back to
                    marshaling the supplied ``app`` proxy into the worker
                    apartment, preserving the exact host instance.
        
        Returns:
            True if add-in can be called successfully.
        
        Raises:
            RuntimeError: If add-in cannot be accessed (missing file,
                untrusted, missing database, dialog-blocked, etc.).
            TimeoutError: If the probe exceeds
                ``ACCESS_VCS_PROBE_TIMEOUT_SEC`` (default 10s) -- typically
                because Access is in VBA break mode or has a modal dialog
                open.
        """
        # Idempotent: skip the probe if we've already loaded the add-in
        # against this instance.  validate_components() and get_version_info
        # both call load_addin and we don't want to double-pay.
        if self._addin_loaded:
            return True
        
        if not self.verify_addin_exists():
            raise RuntimeError(
                f"VCS add-in not found at: {self.addin_path}\n"
                f"Please install the MSAccess VCS add-in or set ACCESS_VCS_ADDIN_PATH.\n"
                f"Download from: https://github.com/joyfullservice/msaccess-vcs-integration/releases"
            )
        
        # Pre-flight: fast file-existence check.  When db_path is provided,
        # catching a missing/moved/typo path here in ~1ms avoids burning the
        # full probe timeout in ROT lookup for a database that isn't there.
        if db_path is not None and not os.path.isfile(db_path):
            raise RuntimeError(
                f"Database file not found at: {db_path}.  The Access "
                f"instance may have closed it, the file was moved/deleted, "
                f"or the path is incorrect."
            )
        
        try:
            timeout_sec = float(os.environ.get("ACCESS_VCS_PROBE_TIMEOUT_SEC", "10"))
        except ValueError:
            timeout_sec = 10.0
        
        probe_start = time.perf_counter()
        probe_error: Optional[str] = None
        timed_out = False
        try:
            self._probe_with_timeout(app, db_path, timeout_sec)
        except TimeoutError as e:
            timed_out = True
            probe_error = str(e)
            raise
        except RuntimeError as e:
            probe_error = str(e)
            raise
        except Exception as e:
            probe_error = str(e)
            raise RuntimeError(
                f"Failed to load VCS add-in: {e}\n"
                f"Ensure a database is open and the add-in is trusted by Access."
            )
        finally:
            duration_ms = round((time.perf_counter() - probe_start) * 1000, 2)
            log_addin_probe(
                addin_path=self.addin_path,
                duration_ms=duration_ms,
                success=probe_error is None,
                timed_out=timed_out,
                error=probe_error,
            )
        
        self._app = app
        self._addin_loaded = True
        self._note_addin_locked(app)
        return True

    def _note_addin_locked(self, app) -> None:
        """Record that this instance now holds the add-in file open.

        A successful probe means the add-in is loaded as a library, which
        locks the file. A rebuild that replaces it has to close every
        server-created instance holding it, whatever database is open --
        and this is the one place every load path passes through.
        """
        try:
            from .access_com.instance_registry import note_loaded_addin
            from .access_com.process_qos import pid_from_access_app

            pid = pid_from_access_app(app)
            if pid:
                note_loaded_addin(pid, self.addin_path)
        except Exception:
            pass
    
    def _probe_with_timeout(self, app, db_path: Optional[str], timeout_sec: float) -> None:
        """
        Run ``GetVCSVersion`` in a daemon worker thread with a hard timeout.
        
        Adapted from db-inspector-mcp's ``_run_dao_with_timeout`` (see that
        project's DECISIONS.md for the full rationale).  The short version:
        Access COM has no native timeout knob, ``CoCancelCall`` requires
        server-side cooperation Jet/ACE doesn't implement, and killing
        ``MSACCESS.EXE`` would lose the user's unsaved work.  A daemon
        thread + ``thread.join(timeout)`` is the only practical way to
        recover responsiveness when Access is stuck.
        
        Raises:
            TimeoutError: If the worker doesn't finish within ``timeout_sec``.
            RuntimeError: If a previous probe is still pending, or if the
                worker can't acquire an Access instance.
        """
        cls = type(self)
        if cls._active_probe_thread is not None and cls._active_probe_thread.is_alive():
            raise RuntimeError(
                "A previous VCS add-in probe is still pending against Access. "
                "This usually means Access is in VBA break mode or has an "
                "open modal dialog.  Resume execution in the VBE, dismiss "
                "any dialog, or restart Access, then retry."
            )
        
        addin_path_abs = os.path.abspath(self.addin_path)
        addin_lib_name = os.path.splitext(addin_path_abs)[0]
        api_function_name = f'{addin_lib_name}.API'
        
        # A new build host may not have registered in the ROT. Marshal its
        # exact application proxy rather than sharing an STA proxy across
        # apartments or looking up an unrelated Access window.
        app_stream = None
        if db_path is None:
            app_stream = pythoncom.CoMarshalInterThreadInterfaceInStream(
                pythoncom.IID_IDispatch, app._oleobj_
            )

        result_box: dict[str, Any] = {}
        
        def worker() -> None:
            try:
                pythoncom.CoInitialize()
                try:
                    if db_path:
                        worker_app = self._find_access_in_rot(db_path)
                    else:
                        worker_app = win32com.client.Dispatch(
                            pythoncom.CoGetInterfaceAndReleaseStream(
                                app_stream, pythoncom.IID_IDispatch
                            )
                        )
                    if worker_app is None:
                        raise RuntimeError(
                            f"Cannot find Access instance for {db_path} "
                            f"from worker thread.  The Access application "
                            f"may have been closed."
                        )
                    worker_app.Run(api_function_name, "GetVCSVersion")
                    result_box["ok"] = True
                finally:
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass
            except Exception as exc:
                result_box["error"] = exc
        
        thread = threading.Thread(
            target=worker, daemon=True, name="vcs-addin-probe"
        )
        cls._active_probe_thread = thread
        thread.start()
        thread.join(timeout=timeout_sec)
        
        if thread.is_alive():
            # Leave _active_probe_thread set so the next call's guard
            # detects the lingering worker and short-circuits with a clear
            # message.  The thread is daemon, so it will be reaped on
            # process exit if Access never responds.
            raise TimeoutError(
                f"VCS add-in probe timed out after {timeout_sec}s "
                f"(no response from Access).  Access is likely in VBA "
                f"break mode (check the VBE), blocked on a modal dialog, "
                f"or hung.  The probe thread will complete naturally once "
                f"Access responds -- no data was lost.  To recover, "
                f"resume execution in the VBE, dismiss any dialog, or "
                f"close and reopen the database."
            )
        
        cls._active_probe_thread = None
        
        if "error" in result_box:
            raise result_box["error"]
    
    @staticmethod
    def _find_access_in_rot(db_path: str):
        """
        Find an Access instance in the Running Object Table that has
        ``db_path`` open.
        
        Two-tier strategy lifted with light edits from db-inspector-mcp's
        ``_find_existing_instance``:
          * Tier 1 -- direct file moniker lookup (~1ms).  ``GetObject`` only
            inspects the ROT; it does NOT fall through to moniker binding,
            so there's no risk of triggering a file-open or password
            dialog.
          * Tier 2 -- enumerate all ROT entries, call ``CurrentDb`` on each
            (~10-50ms).  Catches Access instances that opened the database
            via ``OpenCurrentDatabase`` from a COM client and therefore
            don't appear under a file moniker.  Non-Access entries fail
            on ``CurrentDb`` and are silently skipped.
        
        Returns:
            An Access Application COM object, or None if no match found.
        """
        try:
            pythoncom.CreateBindCtx(0)
            rot = pythoncom.GetRunningObjectTable(0)
        except Exception:
            return None
        
        # Tier 1: direct file moniker lookup
        try:
            moniker = pythoncom.CreateFileMoniker(os.path.abspath(db_path))
            obj = rot.GetObject(moniker)
            return win32com.client.Dispatch(
                obj.QueryInterface(pythoncom.IID_IDispatch)
            )
        except Exception:
            pass
        
        # Tier 2: enumerate ROT entries, check CurrentDb on each
        try:
            enum = rot.EnumRunning()
            target = os.path.normpath(os.path.abspath(db_path)).lower()
            while True:
                monikers = enum.Next(1)
                if not monikers:
                    break
                try:
                    obj = rot.GetObject(monikers[0])
                    dispatch = win32com.client.Dispatch(
                        obj.QueryInterface(pythoncom.IID_IDispatch)
                    )
                    cdb = dispatch.CurrentDb()
                    if cdb is not None:
                        cdb_path = os.path.normpath(
                            os.path.abspath(cdb.Name)
                        ).lower()
                        if cdb_path == target:
                            return dispatch
                except Exception:
                    continue
        except Exception:
            pass
        
        return None
    
    def _call_addin_function(self, function_name: str, *args) -> Any:
        """
        Call a function in the VCS add-in using the new API method.
        
        Note: A database must be open in the Access application before calling
        add-in functions. The add-in loads automatically when called.
        
        Args:
            function_name: Name of function to call (e.g., "GetVCSVersion", "HandleRibbonCommand")
            *args: Arguments to pass to the function
            
        Returns:
            Result from add-in function
            
        Raises:
            RuntimeError: If call fails
        """
        # Gate on add-in load state so callers get a clear lifecycle error
        # instead of a downstream COM message.
        if not self._addin_loaded or not self._app:
            raise RuntimeError("VCS add-in not loaded. Call load_addin() first.")

        try:
            # New API format: Path without extension + ".API", then function name as first argument
            # Example: app.Run("C:\Path\Version Control.API", "GetVCSVersion")
            addin_path_abs = os.path.abspath(self.addin_path)
            addin_lib_name = os.path.splitext(addin_path_abs)[0]
            api_function_name = f'{addin_lib_name}.API'

            # Verify database is open (required for add-in to work)
            try:
                current_db = self._app.CurrentDb()
                if not current_db:
                    raise RuntimeError("No database is currently open in Access. The add-in requires a database to be open.")
                # Force Access to recognize the current database by accessing a property
                # This ensures the database context is fully established
                _ = current_db.Name
            except Exception as db_error:
                raise RuntimeError(f"Cannot access current database: {db_error}. Ensure a database is open before calling add-in functions.")

            # With early binding (gencache.EnsureDispatch), Run returns a tuple
            # where the first element is the actual result.
            try:
                if args:
                    result = self._app.Run(api_function_name, function_name, *args)
                else:
                    result = self._app.Run(api_function_name, function_name)
                if isinstance(result, tuple) and len(result) > 0:
                    return result[0]
                return result
            except Exception as run_error:
                # First call may fail while Access is loading/initializing the
                # add-in. Retry once -- the add-in should now be resident.
                try:
                    if args:
                        result = self._app.Run(api_function_name, function_name, *args)
                    else:
                        result = self._app.Run(api_function_name, function_name)
                    if isinstance(result, tuple) and len(result) > 0:
                        return result[0]
                    return result
                except Exception as second_run_error:
                    raise RuntimeError(
                        f"Failed to call add-in function '{function_name}': {second_run_error}\n"
                        f"First attempt error: {run_error}\n"
                        f"API path used: {api_function_name}\n"
                        f"Add-in path: {self.addin_path}\n"
                        f"Ensure a database is open and the add-in path is correct."
                    )

        except RuntimeError:
            raise
        except Exception as e:
            raise RuntimeError(
                f"Failed to call add-in function '{function_name}': {e}\n"
                f"Ensure a database is open and the add-in path is correct."
            )
    
    def _get_export_folder(self, db_path: str, source_folder: Optional[str] = None) -> str:
        """
        Determine export folder path.
        
        Args:
            db_path: Path to database file
            source_folder: Optional explicit source folder path
            
        Returns:
            Path to export folder
        """
        if source_folder:
            return source_folder
        
        # Default: database_name.src folder next to database
        db_file = Path(db_path)
        return str(db_file.parent / f"{db_file.stem}.src")
    
    def export_source(
        self,
        db_path: str,
        source_folder: Optional[str] = None,
        full_export: bool = False
    ) -> dict[str, Any]:
        """
        Export database to source files using VCS add-in.
        
        Note: This is the synchronous fallback. The preferred approach is to use
        call_async() with the "Export" command, which spawns via timer and allows
        the UI to show while posting progress callbacks.
        
        Args:
            db_path: Path to Access database
            source_folder: Optional custom export folder (default: db_name.src)
            full_export: If True, force full export; if False, use fast save
            
        Returns:
            Dictionary with the start result, never ``success: true``: Export and
            FullExport return Empty, so the outcome is unknown.
            - success: False
            - started / completion_unconfirmed: True when the call returned
              normally; ``error`` then says to read the log
            - error / message: The failure, for a refusal or a COM exception
            - export_path: Path the export was asked to write to
            - log_path: The add-in's explicit operation log, when supplied
        """
        export_path = self._get_export_folder(db_path, source_folder)
        command = "FullExport" if full_export else "Export"
        result = self._start_only(command, "export", "Export")
        result["export_path"] = export_path
        return result
    
    def export_vba(self, db_path: str, source_folder: Optional[str] = None) -> dict[str, Any]:
        """
        Export only VBA components (modules, class modules).
        
        Note: This is the synchronous fallback. The preferred approach is to use
        call_async() with the "ExportVBA" command.
        
        Args:
            db_path: Path to Access database
            source_folder: Optional custom export folder
            
        Returns:
            Dictionary with the start result, as for ``export_source``
        """
        export_path = self._get_export_folder(db_path, source_folder)
        result = self._start_only("ExportVBA", "VBA export", "Export")
        result["export_path"] = export_path
        return result
    
    def merge_build(
        self,
        db_path: str,
        source_folder: Optional[str] = None,
        decision_policy: Optional[str] = None,
    ) -> dict[str, Any]:
        """
        Merge source files into existing database.
        
        This updates modified objects without rebuilding the entire database.
        
        Note: This is the synchronous fallback. The preferred approach is to use
        call_async() with the "MergeBuild" command.
        
        Args:
            db_path: Path to Access database
            source_folder: Optional custom source folder
            
        Returns:
            Dictionary with build results:
            - success: Boolean
            - database_path: Path to database
            - log_path: The add-in's explicit operation log, when supplied
            - message: Status message
        """
        try:
            if decision_policy:
                raw = self.call_sync("MergeBuild", decision_policy)
            else:
                raw = self.call_sync("MergeBuild")
        except Exception as e:
            return {
                "success": False,
                "database_path": db_path,
                "log_path": None,
                "message": f"Merge build failed: {e}"
            }

        # The add-in's return is its own start result: a refusal, or a
        # ``started`` marker. Pass it through; never invent a success.
        parsed: Any = raw
        if isinstance(raw, str):
            try:
                parsed = json.loads(raw)
            except json.JSONDecodeError:
                parsed = None
        if not isinstance(parsed, dict):
            return {
                "success": False,
                "database_path": db_path,
                "log_path": None,
                "message": "Merge build returned no parseable result; "
                           "the add-in build may be too old.",
            }
        result = dict(parsed)
        wrapped = api_refusal_payload(result.get("error"))
        if wrapped:
            result.update(wrapped)
        result.setdefault("database_path", db_path)
        log_path = result.get("log_path") or result.get("logPath")
        if log_path:
            result["log_path"] = log_path
        # The public log resolver checks execution and timestamps. A legacy
        # Build.log here would look explicit and bypass those attribution checks.
        return result

    def build_from_source(
        self,
        source_folder: str,
        output_path: Optional[str] = None
    ) -> dict[str, Any]:
        """
        Build database from source files.
        
        Creates a fresh database from source files.
        
        Args:
            source_folder: Path to source files folder
            output_path: Optional path for new database (default: the name
                the source files record). Given, it is passed to
                ``BuildAs(source, output)``; confirm ``build_as_paths_refusal``
                first, since an older add-in opens pickers for BuildAs.

        Returns:
            Dictionary with the start result, never ``success: true``: no
            callback follows this call, so the outcome is unknown.
            - success: False
            - started / completion_unconfirmed: True when the call returned
              normally; ``error`` then says to read the log
            - error / message: The failure, for a refusal or a COM exception
            - output_path: None. Only a completion callback reports where a
              build wrote; ``requested_output_path`` keeps the request on an
              unconfirmed start
            - log_path: The add-in's explicit operation log, when supplied
        """
        if output_path:
            args: tuple[str, ...] = (source_folder, output_path)
            command = "BuildAs"
        else:
            args = (source_folder,)
            command = "Build"
        result = self._start_only(command, "build", "Build", *args)
        result["output_path"] = None
        if output_path and result.get("completion_unconfirmed"):
            result["requested_output_path"] = output_path
        return result

    def build_as_paths_refusal(self) -> dict[str, Any] | None:
        """None when the add-in confirms ``build_as_paths``, else the refusal.

        Runs ``APICapabilities`` by name, never through ``API``: an older
        add-in without it makes ``Application.Run`` raise at once, where
        ``API`` would leave a modal runtime-error dialog in Access.

        Fails closed: a raise, a reply that is not
        ``{success: true, capabilities: [...]}``, or a list without the name
        refuses the build before it starts.
        """
        if not self._addin_loaded or not self._app:
            return build_output_unsupported_result("the add-in is not loaded")
        procedure = (
            f"{os.path.splitext(os.path.abspath(self.addin_path))[0]}"
            f".{CAPABILITIES_PROCEDURE}"
        )
        try:
            raw = self._app.Run(procedure)
        except Exception as e:
            return build_output_unsupported_result(f"{CAPABILITIES_PROCEDURE} failed: {e}")
        if isinstance(raw, tuple):
            raw = raw[0] if raw else None
        reply: Any = raw
        if isinstance(raw, str):
            try:
                reply = json.loads(raw)
            except json.JSONDecodeError:
                reply = None
        if not isinstance(reply, dict) or reply.get("success") is not True:
            return build_output_unsupported_result(
                f"{CAPABILITIES_PROCEDURE} returned {raw!r}"
            )
        capabilities = reply.get("capabilities")
        if not isinstance(capabilities, list) or BUILD_AS_PATHS not in capabilities:
            return build_output_unsupported_result(f"capabilities: {capabilities!r}")
        return None

    def _start_only(
        self,
        command: str,
        operation: str,
        log_type: str,
        *args: Any,
    ) -> dict[str, Any]:
        """Dispatch a start-only API method and report only what its return proves.

        Disk fallback belongs to the public resolver, which checks whether the
        operation ran, the log family and its timestamp. Synthesizing a legacy
        path here would bypass those checks as if the add-in supplied it.
        """
        try:
            raw = self._call_addin_function(command, *args)
        except Exception as e:
            message = f"The {operation} failed: {e}"
            return {
                "success": False, "log_path": None,
                "error": message, "message": message,
            }
        result = start_only_result(raw, operation, log_type)
        result.setdefault("log_path", None)
        result["message"] = result.get("error")
        return result
    
    def parse_log_file(self, log_path: str) -> dict[str, Any]:
        """
        Parse add-in log file for detailed results.
        
        Args:
            log_path: Path to Export.log or Build.log
            
        Returns:
            Dictionary with parsed log information
        """
        if not os.path.exists(log_path):
            return {"found": False}
        
        try:
            with open(log_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            return {
                "found": True,
                "content": content,
                "path": log_path
            }
            
        except Exception as e:
            return {
                "found": False,
                "error": str(e)
            }
    
    # =========================================================================
    # Sync/Async API Methods (for callback-enabled operations)
    # =========================================================================
    
    def call_sync(self, command: str, *args) -> Any:
        """
        Call a VBA function synchronously using the API entry point.
        
        This calls the existing API function which returns results immediately.
        Use for quick operations that don't need progress reporting.
        
        Args:
            command: Command name (e.g., "GetVCSVersion", "GetOptions")
            *args: Additional arguments to pass
            
        Returns:
            Result from the VBA function
            
        Raises:
            RuntimeError: If call fails

        A dispatcher refusal (``VCS_API_REFUSED: ...``) is returned as a JSON
        failure object with ``error_pattern`` ``operation_already_running`` (or
        ``api_self_dispatch`` for the add-in's self-dispatch defect), so
        every consumer sees ``success: false`` instead of a truthy string.
        """
        result = self._call_addin_function(command, *args)
        refusal = api_refusal_payload(result)
        return json.dumps(refusal) if refusal else result
    
    def call_async(self, callback_info: str, command: str, *args) -> dict[str, Any]:
        """
        Call a VBA function asynchronously using the APIAsync entry point.
        
        This calls the new APIAsync function which:
        - Spawns a detached process for long-running operations
        - Returns immediately with async marker and timeout hint
        - Sends progress updates via HTTP callbacks
        - Sends completion or error when done
        
        Args:
            callback_info: JSON string with callback_url and operation_id
            command: Command name (e.g., "Export", "Build", "MergeBuild")
            *args: Additional arguments to pass
            
        Returns:
            Dict with either:
            - {"sync": true, "result": ...} for quick operations
            - {"async": true, "timeout_ms": ...} for async operations
            
        Raises:
            RuntimeError: If call fails
        """
        import json
        
        try:
            # Call APIAsync entry point
            # Format: APIAsync(strCallbackInfo, strCommand, [args...])
            addin_path_abs = os.path.abspath(self.addin_path)
            addin_lib_name = os.path.splitext(addin_path_abs)[0]
            api_async_name = f'{addin_lib_name}.APIAsync'
            
            # Ensure app reference is set
            if not self._app:
                raise RuntimeError("Access Application object not set.")
            
            # Call APIAsync with callback info as first arg
            if args:
                result = self._app.Run(api_async_name, callback_info, command, *args)
            else:
                result = self._app.Run(api_async_name, callback_info, command)
            
            # Handle tuple return from early-bound Run method
            if isinstance(result, tuple) and len(result) > 0:
                result = result[0]
            
            # A dispatcher refusal is a marked string, not JSON.
            refusal = api_refusal_payload(result)
            if refusal:
                return refusal

            # Parse JSON response from VBA
            if isinstance(result, str):
                envelope = json.loads(result)
                # The add-in wraps its refusal as {success: false, error: <marked>}.
                if isinstance(envelope, dict):
                    wrapped = api_refusal_payload(envelope.get("error"))
                    if wrapped:
                        return {**envelope, **wrapped}
                return envelope
            else:
                # Unexpected return type
                return {"sync": True, "result": result}
                
        except json.JSONDecodeError as e:
            raise RuntimeError(f"Invalid JSON response from APIAsync: {e}")
        except Exception as e:
            raise RuntimeError(f"Failed to call APIAsync '{command}': {e}")
    
    def get_version_info(self, app) -> dict[str, Any]:
        """
        Get comprehensive version information for VCS add-in and Access.
        
        Note: A database must be open in the Access application before calling this.
        
        Args:
            app: Access Application COM object (with database open)
            
        Returns:
            Dictionary with vcs_version, access_version, bitness, and paths
        """
        result = {
            "success": True,
            "addin_path": self.addin_path,
        }
        
        # Get Access application info first -- this works even if the add-in
        # itself is unhealthy, so callers always get at least the Access
        # version/bitness fields.
        access_info = get_access_info(app)
        result.update(access_info)
        
        # Load the add-in via the lifecycle gate.  Idempotent: if a caller
        # (e.g. validate_components) already loaded it, this is an early
        # return.  db_path=None: get_version_info doesn't have the DB path
        # in scope, so the probe falls back to the main thread's app proxy.
        try:
            self.load_addin(app)
        except Exception as load_error:
            result["vcs_version"] = None
            result["vcs_error"] = f"Failed to load VCS add-in: {load_error}"
            result["success"] = False
            return result
        
        # Get VCS add-in version
        try:
            vcs_version = self._call_addin_function("GetVCSVersion")
            result["vcs_version"] = vcs_version
        except Exception as e:
            result["vcs_version"] = None
            result["vcs_error"] = f"Failed to get VCS version: {e}"
            result["success"] = False
        
        return result
