"""Child interpreter for the exit test: a pending build, then normal exit.

argv: marker_path registry_path. No COM: the apartment initializer is
replaced and the pinned process is a fake that never exits.
"""
import asyncio
import atexit
import os
import sys
from unittest.mock import MagicMock

marker, registry_path = sys.argv[1], sys.argv[2]
os.environ["ACCESS_VCS_OWNED_INSTANCES_PATH"] = registry_path
os.environ["ACCESS_VCS_DIAGNOSTIC_LOG_DIR"] = os.path.dirname(registry_path)

from msaccess_vcs_mcp import access_gate, build_owner  # noqa: E402
from msaccess_vcs_mcp.operation_manager import DispatchOutcome, OperationManager  # noqa: E402

access_gate._init_com_apartment = lambda: None


def note(text):
    with open(marker, "a", encoding="utf-8") as handle:
        handle.write(text + "\n")


atexit.register(lambda: note("atexit-ran"))

process = MagicMock()
process.exited.return_value = False
build_owner.pid_from_access_app = lambda app: 771
build_owner.registry.process_create_time = lambda pid: 4321
build_owner.open_exact_process = lambda pid, stamp: process
manager = OperationManager()


async def body():
    operation_id, _ = manager.register_operation(database_path="out.accdb")
    owner = build_owner._BuildOwner(manager, operation_id)
    owner.capture(MagicMock(), "out.accdb")
    owner.dispatched = True
    owner.outcome = DispatchOutcome("uncertain", {})
    owner.finish()
    return dict(owner.cleanup)


result = asyncio.run(access_gate.AccessGate().run_exclusive("build", None, body, True))
note("returned " + result["state"])
