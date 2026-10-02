"""M48: refused whole exports terminate through the decorated public wrapper."""

import asyncio
import json
import os
from unittest.mock import AsyncMock, Mock

import pytest

from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.access_gate import AccessGate
from msaccess_vcs_mcp.operation_manager import OperationManager

from .test_export_build_fallbacks import (
    EXPORT_METHODS, FakeAccess, _addin, _assert_unconfirmed, _export_env,
)


@pytest.fixture
def public_export(tmp_path, monkeypatch):
    """Real wrapper, gate, integration and callback manager; disposable COM/files."""
    gate = AccessGate()
    monkeypatch.setattr(tools, "get_access_gate", lambda: gate)
    monkeypatch.setattr(tools, "_ensure_env_loaded", AsyncMock())
    monkeypatch.setattr(tools, "load_config", lambda: {})
    monkeypatch.setenv("ACCESS_VCS_ENABLE_LOGGING", "false")
    monkeypatch.setenv("ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG", "true")

    def run(fake, full_export, *, refusal_callback=None):
        manager = OperationManager()
        manager.unregister_operation = Mock(wraps=manager.unregister_operation)
        manager.wait_for_completion = AsyncMock(wraps=manager.wait_for_completion)
        operation_ids = []

        def on_async(callback_info):
            operation_id = json.loads(callback_info)["operation_id"]
            operation_ids.append(operation_id)
            if refusal_callback:
                assert manager.route_callback(operation_id, {
                    **refusal_callback, "operation_id": operation_id, "type": "error",
                    "error": "Duplicate callback must not replace the start refusal.",
                })

        fake.on_async = on_async
        fake.export_folder = str(tmp_path / "export") + os.sep
        with _export_env(tmp_path, _addin(tmp_path, fake), manager, True) as (db, out):
            # No unwrapping: config, logging, gate and result finalization run.
            result = asyncio.run(tools.vcs_export_database(
                str(db), str(out), full_export=full_export,
            ))
        assert manager.pending_count() == 0
        return result, manager, operation_ids, out

    yield run
    gate._executor.shutdown(wait=True)


@pytest.mark.parametrize("full_export, method", EXPORT_METHODS)
@pytest.mark.parametrize("second_success", [True, False], ids=["retry-would-succeed", "retry-would-refuse"])
def test_public_export_returns_start_refusal_once(
    public_export, addin_start_refusal, full_export, method, second_success,
):
    expected = addin_start_refusal["expected"]
    fake = FakeAccess(
        async_start=addin_start_refusal["raw"],
        api=json.dumps({"success": second_success, "error": "Second dispatch replaced the refusal"}),
    )
    result, manager, operation_ids, out = public_export(
        fake, full_export, refusal_callback=expected,
    )

    # Prove the exact symptom first: a retry that would succeed never dispatches.
    assert fake.dispatched == [("APIAsync", method)]
    for key, value in expected.items():
        assert result[key] == value
    assert result["exported_count"] == 0
    assert result["objects_by_type"] == {}
    assert result["export_path"] == str(out)
    assert "started" not in result
    assert "completion_unconfirmed" not in result
    operation_id, = operation_ids
    manager.unregister_operation.assert_called_once_with(operation_id)
    manager.wait_for_completion.assert_not_awaited()
    assert manager.is_database_busy(str(out.parent / "test.accdb")) is False
    assert manager.route_callback(operation_id, {
        **expected, "operation_id": operation_id, "type": "error",
    }) is False
    assert fake.dispatched == [("APIAsync", method)]


@pytest.mark.parametrize("full_export, method", EXPORT_METHODS)
@pytest.mark.parametrize("start", [
    '{"unexpected": true}', RuntimeError("async start failed"),
    '{"sync": true, "result": null}',
], ids=["unknown-start", "async-exception", "inline-empty"])
def test_public_export_preserves_unknown_exception_and_inline_paths(
    public_export, full_export, method, start,
):
    fake = FakeAccess(async_start=start)
    result, manager, _, _ = public_export(fake, full_export)
    _assert_unconfirmed(result)
    inline = isinstance(start, str) and '"sync"' in start
    assert fake.dispatched == [("APIAsync", method)] + ([] if inline else [("API", method)])
    manager.unregister_operation.assert_called_once()
    manager.wait_for_completion.assert_not_awaited()


@pytest.mark.parametrize("full_export, method", EXPORT_METHODS)
def test_public_export_parses_inline_refusal_without_redispatch(
    public_export, addin_start_refusal, full_export, method,
):
    fake = FakeAccess(async_start=json.dumps({"sync": True, "result": addin_start_refusal["raw"]}))
    result, manager, _, _ = public_export(fake, full_export)
    for key, value in addin_start_refusal["expected"].items():
        assert result[key] == value
    assert "completion_unconfirmed" not in result
    assert fake.dispatched == [("APIAsync", method)]
    manager.unregister_operation.assert_called_once()
    manager.wait_for_completion.assert_not_awaited()
