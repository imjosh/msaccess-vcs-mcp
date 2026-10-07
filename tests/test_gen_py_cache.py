"""Tests for pywin32 cache recovery without attaching or starting extra hosts."""

import sys
from unittest.mock import MagicMock, patch

import pytest

from msaccess_vcs_mcp.access_com import connection as conn_mod
from msaccess_vcs_mcp.access_com.connection import (
    _extract_gen_py_folder_from_error,
    _gen_py_folder_incomplete,
    _purge_gen_py_cache_folder,
    _should_heal_gen_py_cache,
    ensure_dispatch,
    create_isolated_access_app,
)
from msaccess_vcs_mcp.usage_logging import _extract_error_pattern

ACCESS_TLB = "4AFFC9A0-5F99-101B-AF4E-00AA003F0F07x0x9x0"
CORRUPT_MSG = (
    f"module 'win32com.gen_py.{ACCESS_TLB}' has no attribute 'CLSIDToClassMap'"
)


class TestGenPyDetection:
    def test_extract_folder_from_error(self):
        assert _extract_gen_py_folder_from_error(CORRUPT_MSG) == ACCESS_TLB

    def test_should_heal_clsid_to_class_map_error(self):
        exc = AttributeError(CORRUPT_MSG)
        assert _should_heal_gen_py_cache(exc) is True

    def test_should_heal_incomplete_folder(self, tmp_path):
        folder = tmp_path / ACCESS_TLB
        folder.mkdir()
        (folder / "__pycache__").mkdir()

        with patch.object(conn_mod.gencache, "GetGeneratePath", return_value=str(tmp_path)):
            assert _gen_py_folder_incomplete(ACCESS_TLB) is True
            assert _should_heal_gen_py_cache(
                AttributeError(f"module 'win32com.gen_py.{ACCESS_TLB}' missing stuff")
            ) is True

    def test_should_not_heal_unrelated_attribute_error(self):
        assert _should_heal_gen_py_cache(AttributeError("no such attribute")) is False


class TestGenPyPurge:
    def test_purge_removes_folder_and_modules(self, tmp_path):
        folder = tmp_path / ACCESS_TLB
        folder.mkdir()
        (folder / "__pycache__").mkdir()
        module_name = f"win32com.gen_py.{ACCESS_TLB}"
        sys.modules[module_name] = MagicMock()
        sys.modules[f"{module_name}._Application"] = MagicMock()

        with (
            patch.object(conn_mod.gencache, "GetGeneratePath", return_value=str(tmp_path)),
            patch.object(conn_mod.gencache, "Rebuild") as mock_rebuild,
            patch("msaccess_vcs_mcp.usage_logging.log_diagnostic_event") as mock_log,
        ):
            _purge_gen_py_cache_folder(ACCESS_TLB, prog_id="Access.Application")

        assert not folder.exists()
        assert module_name not in sys.modules
        assert f"{module_name}._Application" not in sys.modules
        mock_rebuild.assert_called_once()
        mock_log.assert_called_once()
        assert mock_log.call_args.args[0] == "gen_py_cache_rebuilt"


class TestEnsureDispatch:
    def test_retries_after_corrupt_cache(self):
        app = MagicMock()
        corrupt = AttributeError(CORRUPT_MSG)

        with (
            patch.object(conn_mod.gencache, "EnsureDispatch", side_effect=[corrupt, app]),
            patch.object(conn_mod, "_purge_gen_py_cache_folder") as mock_purge,
        ):
            result = ensure_dispatch("Access.Application")

        assert result is app
        mock_purge.assert_called_once_with(ACCESS_TLB, prog_id="Access.Application")

    def test_raises_when_retry_also_fails(self):
        corrupt = AttributeError(CORRUPT_MSG)

        with (
            patch.object(
                conn_mod.gencache,
                "EnsureDispatch",
                side_effect=[corrupt, corrupt],
            ),
            patch.object(conn_mod, "_purge_gen_py_cache_folder"),
        ):
            with pytest.raises(AttributeError, match="CLSIDToClassMap"):
                ensure_dispatch("Access.Application")

    def test_raises_unrelated_attribute_error_without_retry(self):
        with patch.object(
            conn_mod.gencache,
            "EnsureDispatch",
            side_effect=AttributeError("other problem"),
        ) as mock_dispatch:
            with pytest.raises(AttributeError, match="other problem"):
                ensure_dispatch("Access.Application")

        assert mock_dispatch.call_count == 1


class TestGenPyErrorPattern:
    def test_extract_error_pattern_gen_py_cache(self):
        assert _extract_error_pattern(CORRUPT_MSG) == "gen_py_cache"


class TestIsolatedDispatchRecovery:
    @pytest.fixture
    def native(self):
        raw = MagicMock()
        typelib = MagicMock()
        typelib.GetLibAttr.return_value = (
            "{4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}", 0, 1, 9, 0, 0,
        )
        raw.GetTypeInfo.return_value.GetContainingTypeLib.return_value = (typelib, 0)
        with (
            patch.object(conn_mod.pythoncom, "CoCreateInstanceEx", return_value=(raw,)) as create,
            patch.object(conn_mod.win32com.client, "Dispatch") as wrap,
            # Keep the old implementation safe during the red run too.
            patch.object(conn_mod.win32com.client, "DispatchEx", side_effect=AttributeError(CORRUPT_MSG)),
            patch.object(conn_mod.gencache, "EnsureDispatch") as attach,
            patch.object(conn_mod.gencache, "EnsureModule") as generate,
            patch.object(conn_mod, "_purge_gen_py_cache_folder") as purge,
        ):
            yield raw, create, wrap, attach, generate, purge

    def test_repairs_wrapper_on_same_native_object(self, native):
        raw, create, wrap, attach, generate, purge = native
        app = MagicMock()
        wrap.side_effect = [AttributeError(CORRUPT_MSG), app]

        assert create_isolated_access_app() is app

        create.assert_called_once_with(
            "Access.Application", None, conn_mod.pythoncom.CLSCTX_SERVER,
            None, (conn_mod.pythoncom.IID_IDispatch,),
        )
        assert wrap.call_count == 2
        assert all(call.args[0] is raw for call in wrap.call_args_list)
        purge.assert_called_once_with(ACCESS_TLB, prog_id="Access.Application")
        generate.assert_called_once_with(
            "{4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}", 0, 9, 0,
        )
        attach.assert_not_called()
        raw.Quit.assert_not_called()

    def test_healthy_wrapper_does_not_repair(self, native):
        raw, create, wrap, attach, generate, purge = native
        app = MagicMock()
        wrap.return_value = app

        assert create_isolated_access_app() is app

        create.assert_called_once()
        assert wrap.call_args.args[0] is raw
        wrap.assert_called_once()
        purge.assert_not_called()
        generate.assert_not_called()
        attach.assert_not_called()

    def test_unavailable_wrapper_generation_does_not_fall_back(self, native):
        raw, create, wrap, attach, generate, purge = native
        wrap.side_effect = AttributeError(CORRUPT_MSG)
        generate.return_value = None

        with pytest.raises(RuntimeError, match="wrapper generation unavailable"):
            create_isolated_access_app()

        create.assert_called_once()
        wrap.assert_called_once()
        attach.assert_not_called()

    def test_second_wrapper_failure_propagates_without_extra_host(self, native):
        raw, create, wrap, attach, generate, purge = native
        wrap.side_effect = AttributeError(CORRUPT_MSG)

        with pytest.raises(AttributeError, match="CLSIDToClassMap"):
            create_isolated_access_app()

        create.assert_called_once()
        assert wrap.call_count == 2
        purge.assert_called_once()
        generate.assert_called_once()
        attach.assert_not_called()

    def test_unrelated_wrapper_failure_does_not_repair(self, native):
        raw, create, wrap, attach, generate, purge = native
        wrap.side_effect = AttributeError("other problem")

        with pytest.raises(AttributeError, match="other problem"):
            create_isolated_access_app()

        create.assert_called_once()
        wrap.assert_called_once()
        purge.assert_not_called()
        generate.assert_not_called()
        attach.assert_not_called()

    def test_regeneration_failure_propagates_without_retrying_activation(self, native):
        raw, create, wrap, attach, generate, purge = native
        wrap.side_effect = AttributeError(CORRUPT_MSG)
        generate.side_effect = PermissionError("cache write denied")

        with pytest.raises(PermissionError, match="cache write denied"):
            create_isolated_access_app()

        create.assert_called_once()
        wrap.assert_called_once()
        attach.assert_not_called()

    def test_different_typelib_is_not_purged(self, native):
        raw, create, wrap, attach, generate, purge = native
        raw.GetTypeInfo.return_value.GetContainingTypeLib.return_value[0].GetLibAttr.return_value = (
            "{00000000-0000-0000-0000-000000000000}", 0, 1, 9, 0, 0,
        )
        wrap.side_effect = AttributeError(CORRUPT_MSG)

        with pytest.raises(AttributeError, match="CLSIDToClassMap"):
            create_isolated_access_app()

        purge.assert_not_called()
        generate.assert_not_called()
        attach.assert_not_called()

    def test_native_failure_does_not_repair_or_attach(self, native):
        raw, create, wrap, attach, generate, purge = native
        create.side_effect = RuntimeError("native activation failed")

        with pytest.raises(RuntimeError, match="native activation failed"):
            create_isolated_access_app()

        create.assert_called_once()
        wrap.assert_not_called()
        purge.assert_not_called()
        generate.assert_not_called()
        attach.assert_not_called()
