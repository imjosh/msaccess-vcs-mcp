import os
import shutil
import tempfile
import threading
from contextlib import contextmanager
from pathlib import Path
from unittest.mock import MagicMock, patch

from msaccess_vcs_mcp import build_owner
from msaccess_vcs_mcp.access_gate import AccessGate

# Tests patch shutil.rmtree to hold the host folder; fixture cleanup must not.
_remove_tree = shutil.rmtree


@contextmanager
def owned_build(app):
    process = MagicMock()
    gone = threading.Event()
    process.exited.side_effect = gone.is_set
    process.force.side_effect = lambda: gone.set() or True
    app.Quit.side_effect = gone.set
    gate = AccessGate()
    folder = tempfile.mkdtemp(prefix='vcs-test-registry-')
    with (
        patch.dict(os.environ, {'ACCESS_VCS_OWNED_INSTANCES_PATH': str(Path(folder) / 'owned.json')}),
        patch.object(build_owner, 'pid_from_access_app', return_value=771),
        patch.object(build_owner.registry, 'process_create_time', return_value=4321),
        patch.object(build_owner, 'open_exact_process', return_value=process),
    ):
        try:
            yield gate, gone
        finally:
            gone.set()
            gate._executor.shutdown(wait=True)
            assert build_owner.stop_observers(5.0) == []
            _remove_tree(folder, ignore_errors=True)
