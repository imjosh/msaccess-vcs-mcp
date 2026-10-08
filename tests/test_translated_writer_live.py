"""Opt-in public ExportObject failure with an exact synthetic French result.

Run serially with ACCESS_VCS_RUN_TRANSLATED_WRITER=1 and pytest -m integration.
Receipts survive in verification/translated-writer-integration/run-*; failed
databases remain for recovery. This case does not establish VERIFY-2 acceptance.
"""
import json
import os
from pathlib import Path
import subprocess
import sys
from datetime import datetime

import pytest


@pytest.mark.integration
@pytest.mark.skipif(sys.platform != "win32", reason="Requires Microsoft Access")
@pytest.mark.skipif(
    os.environ.get("ACCESS_VCS_RUN_TRANSLATED_WRITER") != "1",
    reason="Set ACCESS_VCS_RUN_TRANSLATED_WRITER=1 for this isolated native case",
)
def test_public_export_object_translated_writer_failure():
    import win32api
    import win32event

    assert not os.environ.get("PYTEST_XDIST_WORKER"), "Run native tests serially without xdist"
    # Refuse competing executions instead of copying a potentially open library.
    mutex = win32event.CreateMutex(None, False, r"Local\MSAccessVCS-TranslatedWriterTest")
    acquired = False
    try:
        status = win32event.WaitForSingleObject(mutex, 0)
        acquired = status in (win32event.WAIT_OBJECT_0, win32event.WAIT_ABANDONED)
        assert acquired, "Another translated writer native test is running"
        root = Path(__file__).resolve().parents[2]
        work = root / "verification" / "translated-writer-integration" / datetime.now().strftime("run-%Y%m%d-%H%M%S-%f")
        work.mkdir(parents=True)
        print("Translated writer receipt:", work / "result.json", flush=True)
        # Same venv as pytest. No fixture imports or environment mutations in
        # pytest: every package singleton/handler/session stays in this worker.
        with (work / "worker-output.txt").open("x", encoding="utf-8") as output:
            process = subprocess.run(
                [sys.executable, "-B", "-m", "tests.translated_writer_native", str(work)],
                cwd=root / "msaccess-vcs-mcp", stdout=output, stderr=subprocess.STDOUT,
            )
        print((work / "worker-output.txt").read_text(encoding="utf-8"), flush=True)
        assert process.returncode == 0, f"Native worker failed; see {work}"
        receipt = json.loads((work / "result.json").read_text(encoding="utf-8"))
        assert receipt["outcome"] == "passed", receipt
        assert receipt["ownership"]["original_handle_exit_confirmed"]
        assert receipt["sentinel_ownership"]["original_handle_exit_confirmed"]
        assert receipt["registry_namespace_removed"]
        assert receipt["environment_restored"] and receipt["user_language_unchanged"]
        assert receipt["installed_unchanged"] and not receipt.get("cleanup_errors")
    finally:
        if acquired:
            win32event.ReleaseMutex(mutex)
        win32api.CloseHandle(mutex)
