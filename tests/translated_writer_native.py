"""Native translated ExportObject worker; invoked only by the opt-in pytest case.

Process exit discards config, logging handlers, compatibility and COM caches.
Uses OwnedHost and a real owning-STA admission handshake, never fake admission.
"""
import os
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]

import gc
import hashlib
import json
import shutil
import traceback
import re
import uuid
import winreg
import ctypes
from unittest.mock import patch


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def user_language():
    """Read the real user's value, including absence and registry type."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                r"Software\VB and VBA Program Settings\MSAccessVCS\Language") as key:
            value, kind = winreg.QueryValueEx(key, "Language")
            return {"present": True, "value": value, "type": kind}
    except FileNotFoundError:
        return {"present": False}


def run(work):
    work = Path(work).resolve()
    work.mkdir(exist_ok=True)
    env_before = {name: os.environ.get(name) for name in (
        "ACCESS_VCS_LOG_DIR", "ACCESS_VCS_DIAGNOSTIC_LOG_DIR",
        "ACCESS_VCS_OWNED_INSTANCES_PATH", "ACCESS_VCS_ADDIN_PATH")}
    for name, value in {
        "ACCESS_VCS_LOG_DIR": work / "usage",
        "ACCESS_VCS_DIAGNOSTIC_LOG_DIR": work / "diagnostics",
        "ACCESS_VCS_OWNED_INSTANCES_PATH": work / "owned.json",
    }.items():
        os.environ[name] = str(value)

    import pythoncom
    import win32api
    import win32com.client
    import win32con
    import win32file
    from msaccess_vcs_mcp import __version__
    from msaccess_vcs_mcp.access_com.connection import ensure_access_visible
    from msaccess_vcs_mcp.access_com.process_qos import list_access_pids_or_none
    from msaccess_vcs_mcp.addin_integration import VCSAddinIntegration
    from msaccess_vcs_mcp.compatibility import workflow_requirement
    from msaccess_vcs_mcp.compatibility_session import ensure_session
    from msaccess_vcs_mcp.config import get_default_addin_path
    from msaccess_vcs_mcp.decision_policy import call_under_policy
    from tests.policy_live_host import OwnedHost

    receipt = {"mcp_version": __version__, "checks": [], "work": str(work)}
    app = addin = host = sentinel = sentinel_host = None
    registry_active = False
    registry_owned = False
    registry_path = "Software\\MSAccessVCS-Writer-" + uuid.uuid4().hex
    receipt["registry_namespace"] = registry_path
    receipt["user_language_before"] = user_language()
    source_library = Path(get_default_addin_path()).resolve()
    receipt["installed_library"] = str(source_library)
    receipt["installed_sha256_before"] = sha(source_library)
    def check(label, actual, expected):
        row = {"label": label, "actual": actual, "expected": expected, "passed": actual == expected}
        receipt["checks"].append(row)
        assert actual == expected, row

    pythoncom.CoInitialize()
    try:
        # Call the real read-only metadata tool before any database work.
        import asyncio
        from msaccess_vcs_mcp.tools import vcs_get_version_info
        preflight = asyncio.run(vcs_get_version_info())
        receipt["preflight"] = preflight
        assert workflow_requirement().reason(preflight.get("mcp_version")) is None, preflight
        assert preflight.get("addin_compatibility", {}).get("success") is True, preflight
        # Require a quiet starting point so no open library is copied and no
        # concurrent native run can interfere with this experiment.
        inventory = list_access_pids_or_none()
        assert inventory == set(), ("Need confirmed empty Access inventory", inventory)
        private_library = work / "Private VCS.accda"
        shutil.copy2(source_library, private_library)
        receipt["copied_library_sha256"] = sha(private_library)
        assert receipt["copied_library_sha256"] == receipt["installed_sha256_before"]
        os.environ["ACCESS_VCS_ADDIN_PATH"] = str(private_library)

        # An independent instance is never loaded with the private library or
        # redirected. Its original process handle must survive fixture teardown.
        sentinel = win32com.client.DispatchEx("Access.Application")
        sentinel_database = work / "unrelated-sentinel.accdb"
        sentinel_host = OwnedHost(sentinel, sentinel_database)
        sentinel.NewCurrentDatabase(str(sentinel_database))
        ensure_access_visible(sentinel)
        sentinel.CurrentDb().CreateQueryDef("qrySentinel", "SELECT 91 AS Untouched")
        sentinel_sql = sentinel.CurrentDb().QueryDefs("qrySentinel").SQL

        app = win32com.client.DispatchEx("Access.Application")
        database = work / "writer-fixture.accdb"
        host = OwnedHost(app, database)
        app.NewCurrentDatabase(str(database))
        ensure_access_visible(app)
        app.CurrentDb().CreateQueryDef("qryWriterLocale", "SELECT 17 AS FixtureValue")

        # Seed only the private copy's data tables. No VBA components in the
        # library are changed, and the installed library is never loaded here.
        language = "fr_SIDE"
        english = "Export completed with errors. Check the log for details."
        french = 'Export termine avec des erreurs. Voir le journal "C:\\source".\r\nDetails.'
        db = app.DBEngine.OpenDatabase(str(private_library))
        try:
            def sql(value):
                return "'" + value.replace("'", "''") + "'"
            db.Execute("DELETE FROM tblTranslation WHERE Language=" + sql(language), 128)
            db.Execute("DELETE FROM tblLanguages WHERE ID=" + sql(language), 128)
            db.Execute("INSERT INTO tblLanguages (ID) VALUES (" + sql(language) + ")", 128)
            rs = db.OpenRecordset("SELECT ID FROM tblStrings WHERE msgid=" + sql(english))
            if rs.EOF:
                receipt["completion_string_inserted"] = True
                rs.Close()
                db.Execute("INSERT INTO tblStrings (msgid) VALUES (" + sql(english) + ")", 128)
                rs = db.OpenRecordset("SELECT ID FROM tblStrings WHERE msgid=" + sql(english))
            try:
                assert not rs.EOF, "Expected production export failure string"
                string_id = int(rs.Fields("ID").Value)
            finally:
                rs.Close()
                rs = None
            db.Execute("INSERT INTO tblTranslation (Language,StringID,Translation) VALUES ("
                       + sql(language) + "," + str(string_id) + "," + sql(french) + ")", 128)
            rs = db.OpenRecordset("SELECT Translation FROM tblTranslation WHERE Language="
                                  + sql(language) + " AND StringID=" + str(string_id))
            try:
                receipt["translation_oracle"] = {"language": language, "msgid": english,
                                                  "string_id": string_id,
                                                  "translation": rs.Fields("Translation").Value}
                check("independent DAO French text", receipt["translation_oracle"]["translation"], french)
            finally:
                rs.Close()
                rs = None
        finally:
            db.Close()
            db = None

        ref = app.References.AddFromFile(str(private_library))
        project_name = ref.Name
        ref = None
        receipt["private_reference_project"] = project_name
        helper = (
            'Attribute VB_Name = "modSideWriterFixture"\n'
            'Option Explicit\n'
            'Public Function SideTranslate(ByVal value As String) As String\n'
            f'    SideTranslate = {project_name}.T(value)\n'
            'End Function\n'
            'Public Function SideConfigured(ByVal ignored As String) As String\n'
            f'    SideConfigured = {project_name}.Translation.Language\n'
            'End Function\n'
            'Public Function SideIdentity(ByVal ignored As String) As String\n'
            f'    SideIdentity = {project_name}.APIIdentity()\n'
            'End Function\n'
            'Public Function SideCurrent(ByVal ignored As String) As String\n'
            f'    SideCurrent = {project_name}.Translation.GetCurrentLanguage\n'
            'End Function\n'
            'Public Function SideConfigure(ByVal value As String) As String\n'
            '    On Error GoTo Failed\n'
            f'    {project_name}.Translation.Language = value\n'
            f'    {project_name}.ReleaseObjects\n'
            f'    SideConfigure = {project_name}.Translation.GetCurrentLanguage\n'
            '    Exit Function\n'
            'Failed:\n'
            '    SideConfigure = "CONFIGURATION ERROR " & Err.Number & ": " & Err.Description\n'
            'End Function\n'
        )
        helper_path = work / "modSideWriterFixture.bas"
        helper_path.write_text(helper, encoding="ascii")
        component = app.VBE.ActiveVBProject.VBComponents.Import(str(helper_path))
        component = None
        component = app.VBE.ActiveVBProject.VBComponents.Import(str(HERE / "fixtures" / "translated_writer_registry.bas"))
        component = None

        addin = VCSAddinIntegration(str(private_library))
        def apartment_probe(probe_app, db_path, timeout):
            assert probe_app is app and db_path is None
            addin._probe_session = ensure_session(probe_app, addin.addin_path)
        # Match the qualified M54 fixture: real handshake in the owning STA,
        # preserving every returned refusal rather than fabricating admission.
        with patch.object(addin, "_probe_with_timeout", side_effect=apartment_probe):
            addin.load_addin(app)
        host.addin = addin
        receipt["session"] = json.loads(ensure_session(app, str(private_library)))

        def vba(procedure, argument):
            raw = app.Run(procedure, argument)
            return raw[0] if isinstance(raw, tuple) else raw
        receipt["helper_identity"] = json.loads(vba("SideIdentity", ""))
        receipt["configured_language"] = vba("SideConfigured", "")
        check("helper uses admitted private library", receipt["helper_identity"]["addin_instance"], receipt["session"]["addin_instance"])
        check("helper uses private path", str(Path(receipt["helper_identity"]["loaded_path"]).resolve()), str(private_library.resolve()))
        receipt["binary_source"] = {}
        for project in app.VBE.VBProjects:
            if Path(project.FileName).resolve() != private_library.resolve():
                continue
            module = project.VBComponents("modBuild").CodeModule
            binary = module.Lines(1, module.CountOfLines)
            source_path = ROOT / "msaccess-vcs-addin/Version Control.accda.src/modules/Core/modBuild.bas"
            source = source_path.read_text(encoding="utf-8-sig")
            pattern = r"(?ms)^Public Function FinishSingleObjectExport\(.*?^End Function"
            binary_proc = re.search(pattern, binary.replace("\r\n", "\n")).group(0)
            source_proc = re.search(pattern, source).group(0)
            receipt["binary_source"] = {"path": str(project.FileName), "completion": binary_proc,
                                        "source_sha256": sha(source_path),
                                        "matches_current_source_ignoring_vba_case": binary_proc.casefold() == source_proc.casefold()}
            module = None
        project = None
        check("binary completion matches source ignoring VBA case", receipt["binary_source"]["matches_current_source_ignoring_vba_case"], True)
        check("process registry redirected", vba("SideRegistryBegin", registry_path), 0)
        registry_active = True
        registry_owned = True
        check("configured French reloads after release", vba("SideConfigure", language), language)
        check("persistent isolated language", vba("SideConfigured", ""), language)
        check("real user language unchanged during test", user_language(), receipt["user_language_before"])
        check("real translation table loaded", vba("SideTranslate", english), french)

        def export(label):
            # Never reselect an in-memory cache: completion must reload the
            # configured language on its own after releasing shared objects.
            check(label + " configured French", vba("SideConfigured", ""), language)
            result, state = call_under_policy(addin, "block", "ExportObject", "query", "qryWriterLocale", parse_result=json.loads)
            receipt[label] = {"result": result, "state": state}
            receipt[label]["language_after_completion"] = vba("SideCurrent", "")
            check(label + " French survives completion", receipt[label]["language_after_completion"], language)
            print(label + ": " + json.dumps(result), flush=True)
            check(label + " reached completion", state, "completed")
            assert not result.get("policy_cleanup_error"), result
            log = Path(result.get("logPath") or result.get("log_path") or "")
            assert log.is_file() and log.is_relative_to(work), result
            receipt[label]["log"] = str(log)
            receipt[label]["log_text"] = log.read_text(encoding="utf-8-sig", errors="replace")
            return result

        control = export("control")
        check("normal writer succeeds", control["success"], True)
        output = Path(str(database) + ".src")
        sql_files = list(output.rglob("qryWriterLocale.sql"))
        assert len(sql_files) == 1, [str(p) for p in output.rglob("*")]
        writer_file = sql_files[0]
        receipt["writer_file"] = str(writer_file)
        before = writer_file.read_bytes()
        receipt["writer_sha256_before"] = sha(writer_file)
        # WriteFile intentionally skips identical content. Change the fixture
        # query so the second export must attempt a real write to the locked file.
        app.CurrentDb().QueryDefs("qryWriterLocale").SQL = "SELECT 18 AS FixtureValue"
        query_before = app.CurrentDb().QueryDefs("qryWriterLocale").SQL
        receipt["query_sql_before_failure"] = query_before

        # A native sharing violation at the actual output file, with logs and
        # the export folder still writable. No raised VBA error or mocked writer.
        handle = win32file.CreateFile(str(writer_file), win32con.GENERIC_READ,
                                     win32con.FILE_SHARE_READ, None,
                                     win32con.OPEN_EXISTING, 0, None)
        try:
            failed = export("locked_writer")
            check("query preserved while output locked", app.CurrentDb().QueryDefs("qryWriterLocale").SQL, query_before)
            check("locked output preserved", writer_file.read_bytes() == before, True)
            receipt["writer_sha256_while_locked"] = sha(writer_file)
        finally:
            win32api.CloseHandle(handle)
        check("actual writer reports failure", failed["success"], False)
        check("caller gets exact French failure", failed.get("error"), french)
        check("failure is not cancellation", bool(failed.get("cancelled")), False)
        check("failure is not a decision refusal", failed.get("error_pattern") == "decision_required", False)
        check("no decision required flag", bool(failed.get("decision_required")), False)
        check("writer decision count", len(failed.get("decisions", [])), 1)
        decision = failed["decisions"][0]
        check("writer acknowledgment fields", {k: decision[k] for k in ("kind", "resolution", "title", "detail")},
              {"kind": "acknowledged", "resolution": "acknowledged", "title": "", "detail": ""})
        check("decision names actual native writer", "Error 3004: Write to file failed. Source: modFileAccess.WriteFile" in decision["message"], True)
        check("query preserved", app.CurrentDb().QueryDefs("qryWriterLocale").SQL, query_before)
        text = receipt["locked_writer"]["log_text"]
        check("log identifies actual native writer", "Error 3004: Write to file failed. Source: modFileAccess.WriteFile" in text, True)
        check("log names locked query output", str(writer_file) in text, True)
        check("failure has its own log", receipt["locked_writer"]["log"] != receipt["control"]["log"], True)
        clean = export("after_unlock")
        check("next export succeeds", clean["success"], True)
        check("next export writes changed SQL", writer_file.read_bytes() != before, True)
        check("changed SQL contains new value", "18 AS FixtureValue" in writer_file.read_text(encoding="utf-8-sig"), True)
        check("query preserved after retry", app.CurrentDb().QueryDefs("qryWriterLocale").SQL, query_before)
        receipt["writer_sha256_after_unlock"] = sha(writer_file)
        receipt["query_sql_after_retry"] = app.CurrentDb().QueryDefs("qryWriterLocale").SQL
        receipt["outcome"] = "passed"
    except Exception:
        receipt["outcome"] = "failed"
        receipt["traceback"] = traceback.format_exc()
        # A retained COM exception traceback can itself hold the automation
        # proxy alive during owned-process release.
        traceback.clear_frames(sys.exc_info()[2])
    finally:
        cleanup_errors = []
        def cleanup_step(label, action):
            try:
                action()
            except Exception:
                cleanup_errors.append({"step": label, "traceback": traceback.format_exc()})
                traceback.clear_frames(sys.exc_info()[2])

        try:
            if registry_active:
                cleanup_step("registry mapping", lambda: check("process registry mapping restored", vba("SideRegistryEnd", ""), 0))
                registry_active = bool(cleanup_errors)
            if host is not None:
                def release():
                    nonlocal app
                    if addin is not None:
                        addin._app = None
                    host.app = None
                    app = None
                    gc.collect()
                cleanup_step("owned host", lambda: host.close(release))
            if sentinel is not None:
                if "sentinel_sql" in locals():
                    cleanup_step("sentinel readability", lambda: check("unrelated database remains readable", sentinel.CurrentDb().QueryDefs("qrySentinel").SQL, sentinel_sql))
                def release_sentinel():
                    nonlocal sentinel
                    # The sentinel loads no library and has no operation state.
                    # Explicit graceful Quit avoids depending on Access's idle
                    # window lifetime after final automation release.
                    sentinel.Quit()
                    sentinel_host.app = None
                    sentinel = None
                    gc.collect()
                if sentinel_host is not None:
                    cleanup_step("sentinel host", lambda: sentinel_host.close(release_sentinel))
        finally:
            if host is not None:
                receipt["ownership"] = host.receipt
                host.app = None
            if addin is not None:
                addin._app = None
            app = None
            gc.collect()
            pythoncom.CoUninitialize()
            # Remove only this run's unique registry tree after the original
            # owned process has exited. Keep it if ownership/exit is uncertain.
            def remove_registry_namespace():
                api = ctypes.WinDLL("advapi32", use_last_error=True)
                api.RegDeleteTreeW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p]
                api.RegDeleteTreeW.restype = ctypes.c_long
                status = api.RegDeleteTreeW(ctypes.c_void_p(0xFFFFFFFF80000001), registry_path)
                assert status in (0, 2), status
                try:
                    winreg.DeleteKey(winreg.HKEY_CURRENT_USER, registry_path)
                except FileNotFoundError:
                    pass
                try:
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, registry_path):
                        raise AssertionError("Temporary registry key survived deletion")
                except FileNotFoundError:
                    receipt["registry_namespace_removed"] = True
            if registry_owned and host is not None and host.receipt.get("original_handle_exit_confirmed"):
                cleanup_step("registry namespace", remove_registry_namespace)
            receipt["user_language_after"] = user_language()
            receipt["user_language_unchanged"] = receipt["user_language_after"] == receipt["user_language_before"]
            for name, value in env_before.items():
                if value is None:
                    os.environ.pop(name, None)
                else:
                    os.environ[name] = value
            receipt["environment_restored"] = all(os.environ.get(k) == v for k, v in env_before.items())
            if sentinel_host is not None:
                receipt["sentinel_ownership"] = sentinel_host.receipt
            receipt["installed_sha256_after"] = sha(source_library)
            receipt["installed_unchanged"] = receipt["installed_sha256_after"] == receipt["installed_sha256_before"]
            if cleanup_errors:
                receipt["cleanup_errors"] = cleanup_errors
                receipt["outcome"] = "failed"
            if not (receipt["installed_unchanged"] and receipt["user_language_unchanged"] and receipt["environment_restored"]):
                receipt["outcome"] = "failed"
            # Retain failed fixtures. After both original handles signal exit,
            # discard only the successful run's exact, owned database files;
            # query exports, operation logs and source diagnostics stay reviewable.
            if receipt.get("outcome") == "passed":
                def dispose_databases():
                    assert host.receipt["original_handle_exit_confirmed"]
                    assert sentinel_host.receipt["original_handle_exit_confirmed"]
                    targets = (private_library, database, sentinel_database)
                    for target in targets:
                        assert target.resolve().parent == work.resolve()
                        target.unlink()
                    receipt["disposed_databases"] = [str(p) for p in targets]
                    check("successful owned databases disposed", all(not p.exists() for p in targets), True)
                cleanup_step("database disposal", dispose_databases)
                if cleanup_errors:
                    receipt["cleanup_errors"] = cleanup_errors
                    receipt["outcome"] = "failed"
            (work / "result.json").write_text(json.dumps(receipt, indent=2), encoding="utf-8")
            print("Receipt: " + str(work / "result.json"), flush=True)
    assert receipt["installed_unchanged"], "Installed library changed"
    assert receipt["user_language_unchanged"] and receipt["environment_restored"]
    assert receipt["outcome"] == "passed", "Focused test failed; see " + str(work / "result.json")
    print("PASS: isolated native translated writer failure", flush=True)
    return receipt


if __name__ == "__main__":
    work = Path(sys.argv[1]).resolve()
    try:
        run(work)
    except BaseException as exc:
        # Also retain setup/cleanup failures outside run's inner exception trap.
        path = work / "result.json"
        receipt = json.loads(path.read_text()) if path.exists() else {"work": str(work)}
        receipt["outcome"] = "failed"
        receipt["worker_traceback"] = traceback.format_exc()
        if hasattr(exc, "result"):
            receipt["refusal"] = exc.result
        path.write_text(json.dumps(receipt, indent=2), encoding="utf-8")
        raise
