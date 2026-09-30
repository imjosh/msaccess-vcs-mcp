"""Shared test fixtures."""

from __future__ import annotations

import pytest

from msaccess_vcs_mcp import exempt_workers


@pytest.fixture(autouse=True)
def _drain_exempt_workers(monkeypatch):
    """Let released gate-exempt workers return before a test's patches are undone.

    Those workers are daemon threads, so ``asyncio.run`` does not wait for them.
    Depending on ``monkeypatch`` puts this teardown before the patches are undone.
    """
    yield
    exempt_workers.wait_idle(5.0)
