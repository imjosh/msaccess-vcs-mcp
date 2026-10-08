# Isolated native category Cancel

This support harness is adapted from
`verification/isolated-click-cancel-2026-10-08/INTEGRATION-HANDOFF.md`.
Pytest discovers one adapter in `tests/test_click_cancel_live.py`; the COM
worker and UI drivers are separate subprocesses, not standalone pytest cases.

From the MCP repository, with an interactive Windows desktop and Microsoft
Access installed:

```powershell
$env:ACCESS_VCS_RUN_CLICK_CANCEL = '1'
try {
    ./venv/Scripts/python.exe -B -m pytest tests/test_click_cancel_live.py -m integration -v -s
} finally {
    Remove-Item Env:ACCESS_VCS_RUN_CLICK_CANCEL
}
```

Reserve Access exclusively for this invocation. Close Access yourself and stop
competing native suites/MCP database operations before launching. Select only
this test; mixed selections and xdist are refused. The named mutex serializes
this harness and the original isolated harness. It does not lock production MCP
processes, whose gates are process-local. Empty inventory must be confirmed at
entry; fresh inventory must contain only the exact fixture before each click.
Unconfirmed inventory or unavailable input desktop fails, rather than passing.

The coordinator exclusively creates a fresh `run-*` under
`../verification/click-cancel-integration/`. All fixture/library/log/ownership
files stay there, with an adapter output file alongside the run directory.
Failed runs and successful disposable binaries remain as evidence. The worker
requires the read-only qualified-project receipt at
`../verification/VERIFY-2-2026-10-08-01/evidence/build-1-built-project.json`.
This is intentionally a cross-repository qualification test, not a test that
can run from an MCP checkout alone. A missing qualification receipt fails.

Run the focused controlled checks (no Access instances) separately:

```powershell
./venv/Scripts/python.exe -B tests/click_cancel_native/check_harness.py
```

The original readiness/dialog/click/completion/cleanup bounds, including the
35-second UI subprocess bound, remain intact. Uncertain clicks are never
replayed. Teardown acts only through the verified original disposable Access
handle; uncertain ownership leaves Access untouched. Forced disposal or cleanup
errors fail the final receipt and pytest invocation. No adapter timeout kills
the coordinator and strands its ownership/teardown logic.

The case uses the real version preflight, mutual session admission, interactive
mode acknowledgement and native save-design Cancel. It retains independent
export/import checks for public results, form/source preservation, distinct
operation logs and released state. Production native save/discard dialog policy
remains report-only. Four loaded code-component hashes establish cancellation
path provenance against the qualification receipt; this does not establish
whole-project or binary equivalence, external MCP transport coverage, or
VERIFY-2 acceptance.
