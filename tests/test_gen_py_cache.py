"""Tests for pywin32 gen_py cache self-heal around EnsureDispatch."""

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
