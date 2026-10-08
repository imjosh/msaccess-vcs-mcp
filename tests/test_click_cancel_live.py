"""One opt-in, serialized native invocation for category export and import."""
from datetime import datetime
import json
import os
from pathlib import Path
import subprocess
import sys

import pytest


@pytest.mark.integration
@pytest.mark.skipif(sys.platform != "win32", reason="Requires Windows Access and an interactive desktop")
@pytest.mark.skipif(
    os.environ.get("ACCESS_VCS_RUN_CLICK_CANCEL") != "1",
    reason="Set ACCESS_VCS_RUN_CLICK_CANCEL=1 for the isolated native Cancel case",
)
def test_category_export_and_import_literal_native_cancel(request):
    assert not os.environ.get("PYTEST_XDIST_WORKER"), "Run without xdist"
    assert not getattr(request.config.option, "numprocesses", None), "Run without xdist"
    assert len(request.session.items) == 1, "Select only this test; no competing Access suite"
    root = Path(__file__).resolve().parents[2]
    assert Path(sys.executable).resolve() == (root / "msaccess-vcs-mcp/venv/Scripts/python.exe").resolve(), "Use MCP project venv"
    work = root / "verification/click-cancel-integration" / datetime.now().strftime("run-%Y%m%d-%H%M%S-%f")
    work.parent.mkdir(parents=True, exist_ok=True)
    # The coordinator exclusively creates work, owns teardown and retains the
    # original Access handle. Do not kill it with an adapter timeout: its bounded
    # UI child and cleanup waits contain uncertain native dispatch safely.
    output_path = work.parent / (work.name + "-adapter-output.txt")
    print("Native Cancel receipt:", work / "result.json", flush=True)
    with output_path.open("x", encoding="utf-8") as output:
        process = subprocess.run(
            [sys.executable, "-B", str(Path(__file__).parent / "click_cancel_native/run_click_cancel.py"), str(work)],
            cwd=root, stdout=output, stderr=subprocess.STDOUT,
        )
    print(output_path.read_text(encoding="utf-8"), flush=True)
    receipt_path = work / "result.json"
    assert receipt_path.is_file(), f"Missing final receipt; see {output_path}"
    receipt = json.loads(receipt_path.read_text(encoding="utf-8"))
    assert process.returncode == 0 and receipt["outcome"] == "passed", receipt
    assert receipt["work"] == str(work)
    assert receipt["installed_unchanged"] and receipt["original_handle_exit_confirmed"]
    assert not receipt["errors"]
    assert receipt["worker"]["ownership"]["original_handle_exit_confirmed"]
