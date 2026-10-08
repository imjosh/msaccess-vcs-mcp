"""Isolated native category Cancel experiment; no pytest or suite discovery.

Internal worker: invoke run_click_cancel.py with the MCP project venv, -B.
The parent observes and clicks the exact native Cancel control independently.
All databases, private library copies, helper modules and receipts are new files.
"""
import sys
sys.dont_write_bytecode = True
import gc
import hashlib
import json
import os
from pathlib import Path
import shutil
import traceback
import time
from datetime import datetime

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[2]
ARTIFACTS = ROOT / 'verification/click-cancel-integration'
sys.path.insert(0, str(ROOT / 'msaccess-vcs-mcp'))


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def run():
    work = Path(sys.argv[2]).resolve()
    assert work.parent == ARTIFACTS.resolve() and work.is_dir()
    for name, value in {'ACCESS_VCS_LOG_DIR': work / 'usage',
                        'ACCESS_VCS_DIAGNOSTIC_LOG_DIR': work / 'diagnostics',
                        'ACCESS_VCS_OWNED_INSTANCES_PATH': work / 'owned.json'}.items():
        os.environ[name] = str(value)
    import pythoncom
    import win32com.client
    import win32api
    import win32event
    from msaccess_vcs_mcp import __version__
    from msaccess_vcs_mcp.compatibility import workflow_requirement
    from msaccess_vcs_mcp.compatibility_session import ensure_session
    from msaccess_vcs_mcp.config import get_default_addin_path
    from msaccess_vcs_mcp.access_com.connection import ensure_access_visible
    from msaccess_vcs_mcp.access_com.process_qos import list_access_pids_or_none
    from tests.policy_live_host import OwnedHost
    from native_cancel import write

    assert workflow_requirement().reason(__version__) is None
    import asyncio
    from msaccess_vcs_mcp.tools import vcs_get_version_info
    metadata = asyncio.run(vcs_get_version_info())
    write(work / 'runtime-preflight.json', metadata)
    assert workflow_requirement().reason(metadata.get('mcp_version')) is None
    assert metadata.get('addin_compatibility', {}).get('success') is True
    assert list_access_pids_or_none() == set(), 'Require confirmed empty Access inventory'
    installed = Path(get_default_addin_path())
    receipt = {'mcp_version': __version__, 'installed': str(installed),
               'installed_sha_before': sha(installed), 'checks': [], 'work': str(work),
               'runtime_preflight': metadata}
    private = work / 'Private VCS.accda'
    shutil.copy2(installed, private)
    receipt['private_initial_sha256'] = sha(private)
    assert receipt['private_initial_sha256'] == receipt['installed_sha_before']
    database = work / 'cancel-fixture.accdb'
    app = host = None
    pythoncom.CoInitialize()
    try:
        app = win32com.client.DispatchEx('Access.Application')
        host = OwnedHost(app, database)
        host.receipt['creation_FILETIME'] = str(host.creation)
        write(work / 'ownership.json', dict(host.receipt, creation_FILETIME=str(host.creation)))
        app.NewCurrentDatabase(str(database))
        ensure_access_visible(app)
        app.UserControl = True
        ref = app.References.AddFromFile(str(private))
        lib = str(ref.Name)
        ref = None
        receipt['session'] = json.loads(ensure_session(app, str(private)))
        raw_identity = app.Run(str(private.with_suffix('')) + '.APIIdentity')
        receipt['addin_identity'] = json.loads(raw_identity[0] if isinstance(raw_identity, tuple) else raw_identity)
        qualified_path = ROOT / 'verification/VERIFY-2-2026-10-08-01/evidence/build-1-built-project.json'
        qualified = json.loads(qualified_path.read_text(encoding='utf-8-sig'))
        prior = next(p for p in qualified['projects'] if 'suite-scratch' in p['path'])
        code_checks = []
        project = next(p for p in app.VBE.VBProjects if str(p.Name) == lib)
        for name in ('clsVersionControl', 'modDatabase', 'modAPI', 'clsOperation'):
            expected = next(c for c in prior['components'] if c['name'] == name)['code_sha256']
            cm = project.VBComponents(name).CodeModule
            code = str(cm.Lines(1, cm.CountOfLines))
            actual = hashlib.sha256(code.encode('utf-8')).hexdigest()
            code_checks.append(dict(component=name, actual=actual, qualified=expected, matches=actual == expected))
            cm = None
        project = None
        receipt['qualified_source_relationship'] = dict(evidence=str(qualified_path), evidence_sha256=sha(qualified_path),
            components=code_checks, note='Loaded cancellation-path code comparison; Access binary storage hashes may differ.')
        assert all(c['matches'] for c in code_checks), code_checks
        helper = f'''Attribute VB_Name = "modSideCancel"
Option Explicit
Public Function SideSetup() As String
    Dim f As Form, n As String
    Set f = Application.CreateForm
    n = f.Name
    f.Caption = "Saved baseline caption"
    DoCmd.Close acForm, n, acSaveYes
    Set f = Nothing
    DoCmd.Rename "frmSideCancel", acForm, n
    {lib}.VCS.SetOperationPolicy "block"
    SideSetup = {lib}.VCS.ExportByType("forms", True)
    {lib}.VCS.ClearOperationPolicy
End Function
Public Function SideDirty() As String
    DoCmd.OpenForm "frmSideCancel", acDesign
    Forms("frmSideCancel").Caption = "Unsaved caption kept by Cancel"
    SideDirty = Forms("frmSideCancel").Caption
End Function
Public Function SideMode() As String
    SideMode = {lib}.VCS.SetInteractionMode(0)
End Function
Public Function SideCancel(ByVal exporting As Boolean) As String
    {lib}.Operation.Source = {lib}.eosUserInterface
    {lib}.Operation.ForceUnattended = False
    If exporting Then
        SideCancel = {lib}.VCS.ExportByType("forms", True)
    Else
        SideCancel = {lib}.VCS.ImportByType("forms", True)
    End If
End Function
Public Function SideState() As String
    SideState = CStr(CurrentProject.AllForms("frmSideCancel").IsLoaded) & "|" & _
        Forms("frmSideCancel").Caption & "|" & CStr({lib}.Operation.Status) & "|" & _
        CStr({lib}.Log.Active) & "|" & CStr({lib}.Operation.Status = {lib}.eosRunning)
End Function
Public Function SideDiscard() As Boolean
    DoCmd.Close acForm, "frmSideCancel", acSaveNo
    SideDiscard = True
End Function
'''
        module_path = work / 'modSideCancel.bas'
        module_path.write_text(helper, encoding='ascii')
        component = app.VBE.ActiveVBProject.VBComponents.Import(str(module_path))
        component = None
        # Save the new fixture's helper before running it: no worker save is
        # necessary while this helper is itself on the VBA execution stack.
        app.DoCmd.RunCommand(126)  # acCmdCompileAndSaveAllModules

        def call(name, *args):
            raw = app.Run(name, *args)
            return raw[0] if isinstance(raw, tuple) else raw

        def check(label, actual, expected):
            row = dict(label=label, actual=actual, expected=expected, passed=actual == expected)
            receipt['checks'].append(row)
            assert row['passed'], row

        baseline = json.loads(call('SideSetup'))
        receipt['baseline'] = baseline
        check('baseline export succeeds', baseline['success'], True)
        source = list(Path(str(database) + '.src').rglob('frmSideCancel.*'))
        assert len(source) == 1, source
        source = source[0]
        receipt['source'] = str(source)
        for case in ('export', 'import'):
            if case == 'import':
                text = source.read_text(encoding='utf-8-sig')
                assert 'Saved baseline caption' in text
                source.write_text(text.replace('Saved baseline caption', 'Import source caption'), encoding='utf-8-sig')
            source_hash = sha(source)
            check(case + ' dirty form prepared', call('SideDirty'), 'Unsaved caption kept by Cancel')
            ack = json.loads(call('SideMode'))
            receipt[case + '_mode_ack'] = ack
            check(case + ' mode success', ack.get('success') is True, True)
            check(case + ' effective interactive mode', type(ack.get('effective_mode')) is int and ack['effective_mode'] == 0, True)
            assert host.verified(), 'Fixture ownership changed'
            pending = dict(case=case, pid=host.pid, creation_FILETIME=str(host.creation),
                           database=str(database), work=str(work))
            write(work / (case + '-pending.json'), pending)
            print('WAITING_FOR_NATIVE_CANCEL ' + json.dumps(pending), flush=True)
            # Blocks in Access's actual native save prompt; no injected 2501.
            result = json.loads(call('SideCancel', case == 'export'))
            receipt[case] = result
            print(case + ' result: ' + json.dumps(result), flush=True)
            check(case + ' fails rather than succeeds', result.get('success'), False)
            check(case + ' classified as cancellation', result.get('cancelled'), True)
            check(case + ' exact cancellation message', result.get('error'), 'Operation was canceled.')
            check(case + ' no decision refusal', 'decision_required' in result, False)
            write(work / (case + '-operation-result.json'), dict(pending=pending, result=result))
            log = Path(result['logPath'])
            prior_logs = {baseline['logPath']}
            if case == 'import':
                prior_logs.add(receipt['export']['logPath'])
            check(case + ' own existing log', log.is_file() and
                  log.resolve().is_relative_to(work) and
                  log.name.startswith('Export_' if case == 'export' else 'Merge_') and
                  str(log) not in prior_logs, True)
            receipt[case + '_log'] = log.read_text(encoding='utf-8-sig', errors='replace')
            state = str(call('SideState')).split('|')
            receipt[case + '_state'] = state
            check(case + ' form remains open', state[0], 'True')
            check(case + ' unsaved caption preserved', state[1], 'Unsaved caption kept by Cancel')
            check(case + ' log released', state[3], 'False')
            check(case + ' operation released', state[4], 'False')
            check(case + ' source unchanged', sha(source), source_hash)
            click_path = work / (case + '-click.json')
            for _ in range(30):
                if click_path.is_file() and json.loads(click_path.read_text())['status'] == 'click returned':
                    break
                time.sleep(0.1)
            click = json.loads(click_path.read_text())
            check(case + ' literal click returned', click['status'], 'click returned')
            check(case + ' click matches ownership', click['pending'], pending)
            dialog = json.loads((work / (case + '-dialog.json')).read_text())
            check(case + ' dialog matches click', dialog['dialog'], click['dialog'])
            call('SideDiscard')
        receipt['outcome'] = 'behavior_passed'
    except Exception:
        receipt['outcome'] = 'failed'
        receipt['traceback'] = traceback.format_exc()
        raise
    finally:
        try:
            if host is not None:
                try:
                    if not host.exited():
                        assert host.verified(), 'Uncertain test host; do not close'
                        app.Quit(2)  # acQuitSaveNone; only this disposable fixture
                    host.app = None
                    app = None
                    gc.collect()
                    exited = win32event.WaitForSingleObject(host.handle, 10000) == win32event.WAIT_OBJECT_0
                    host.receipt['original_handle_exit_confirmed'] = exited
                    assert exited, 'Disposable host exit unconfirmed'
                    assert not host.unrelated, 'Entry inventory was required empty'
                finally:
                    win32api.CloseHandle(host.handle)
                    for handle, _ in host.unrelated.values():
                        win32api.CloseHandle(handle)
        except BaseException:
            receipt['outcome'] = 'failed'
            receipt['cleanup_error'] = traceback.format_exc()
            raise
        finally:
            if host is not None:
                receipt['ownership'] = host.receipt
            receipt['installed_sha_after'] = sha(installed)
            receipt['installed_unchanged'] = receipt['installed_sha_before'] == receipt['installed_sha_after']
            write(work / 'worker-result.json', receipt)
            pythoncom.CoUninitialize()
            print('WORKER_RECEIPT ' + str(work / 'worker-result.json'), flush=True)
    assert receipt['installed_unchanged']


if __name__ == '__main__':
    assert len(sys.argv) == 3 and sys.argv[1] == '--worker', 'Use run_click_cancel.py'
    run()
