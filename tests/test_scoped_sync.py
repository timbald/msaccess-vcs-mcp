"""Tests for scoped ImportByType / ExportByType routing via object_types."""

import asyncio
import json
from contextlib import contextmanager
from unittest.mock import MagicMock, patch

import pytest

from msaccess_vcs_mcp.tools import _scoped_types_arg


def _unwrap_sync(tool_fn):
    sync_fn = tool_fn
    while hasattr(sync_fn, "__wrapped__"):
        sync_fn = sync_fn.__wrapped__
    return sync_fn


@pytest.fixture(autouse=True)
def _patch_tool_infra(monkeypatch, tmp_path_factory):
    diag_dir = tmp_path_factory.mktemp("diag")
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(diag_dir))


def _write_log(source_dir, name, content="line one\nline two\n"):
    logs = source_dir / "logs"
    logs.mkdir(exist_ok=True)
    path = logs / name
    path.write_text(content, encoding="utf-8")
    return path


@contextmanager
def _patch_import_tool(tmp_path, *, call_sync_result=None, async_result=None):
    mock_conn = MagicMock()
    mock_conn.__enter__ = MagicMock(return_value=mock_conn)
    mock_conn.__exit__ = MagicMock(return_value=False)
    mock_conn.connect.return_value = (MagicMock(), MagicMock())

    mock_addin = MagicMock()
    if call_sync_result is not None:
        mock_addin.call_sync.return_value = call_sync_result
    mock_addin.call_async.return_value = async_result or {"async": True, "timeout_ms": 1000}
    mock_addin.merge_build.return_value = {"success": True, "message": "ok"}

    db_path = tmp_path / "test.accdb"
    db_path.touch()
    src_path = tmp_path / "src"
    src_path.mkdir(exist_ok=True)

    op_manager = MagicMock()
    op_manager.register_operation.return_value = ("op-1", MagicMock())
    op_manager.create_callback_info.return_value = "{}"
    if async_result and async_result.get("async"):
        async def _wait(*_a, **_k):
            return {"success": True, "log_path": None}

        op_manager.wait_for_completion = _wait

    with (
        patch("msaccess_vcs_mcp.tools.AccessConnection", return_value=mock_conn),
        patch("msaccess_vcs_mcp.tools.VCSAddinIntegration", return_value=mock_addin),
        patch("msaccess_vcs_mcp.tools.validate_database_path", return_value=db_path),
        patch("msaccess_vcs_mcp.tools.validate_source_directory", return_value=src_path),
        patch("msaccess_vcs_mcp.tools.check_write_permission", return_value=None),
        patch("msaccess_vcs_mcp.tools._check_database_busy", return_value=None),
        patch("msaccess_vcs_mcp.tools.get_callback_url", return_value="http://localhost:1/cb"),
        patch("msaccess_vcs_mcp.tools._get_operation_manager", return_value=op_manager),
        patch(
            "msaccess_vcs_mcp.tools.get_config",
            return_value={"ACCESS_VCS_ADDIN_PATH": str(tmp_path / "Version Control.accda")},
        ),
    ):
        yield str(db_path), src_path, mock_addin, op_manager


@contextmanager
def _patch_export_tool(tmp_path, *, call_sync_result=None, async_result=None):
    mock_conn = MagicMock()
    mock_conn.__enter__ = MagicMock(return_value=mock_conn)
    mock_conn.__exit__ = MagicMock(return_value=False)
    mock_conn.connect.return_value = (MagicMock(), MagicMock())

    mock_addin = MagicMock()
    if call_sync_result is not None:
        mock_addin.call_sync.return_value = call_sync_result
    mock_addin.call_async.return_value = async_result or {"async": True, "timeout_ms": 1000}
    mock_addin.export_source.return_value = {"success": True, "message": "ok"}

    db_path = tmp_path / "test.accdb"
    db_path.touch()
    export_path = tmp_path / "src"
    export_path.mkdir(exist_ok=True)

    op_manager = MagicMock()
    op_manager.register_operation.return_value = ("op-1", MagicMock())
    op_manager.create_callback_info.return_value = "{}"
    if async_result and async_result.get("async"):
        async def _wait(*_a, **_k):
            return {"success": True, "log_path": None}

        op_manager.wait_for_completion = _wait

    with (
        patch("msaccess_vcs_mcp.tools.AccessConnection", return_value=mock_conn),
        patch("msaccess_vcs_mcp.tools.VCSAddinIntegration", return_value=mock_addin),
        patch("msaccess_vcs_mcp.tools.validate_database_path", return_value=db_path),
        patch("msaccess_vcs_mcp.tools.validate_export_directory", return_value=export_path),
        patch("msaccess_vcs_mcp.tools._check_database_busy", return_value=None),
        patch("msaccess_vcs_mcp.tools.get_callback_url", return_value="http://localhost:1/cb"),
        patch("msaccess_vcs_mcp.tools._get_operation_manager", return_value=op_manager),
        patch(
            "msaccess_vcs_mcp.tools.get_config",
            return_value={"ACCESS_VCS_ADDIN_PATH": str(tmp_path / "Version Control.accda")},
        ),
    ):
        yield str(db_path), export_path, mock_addin, op_manager


class TestScopedTypesArg:
    def test_single_type_is_bare_string(self):
        assert _scoped_types_arg(["modules"]) == "modules"

    def test_multiple_types_remain_list(self):
        assert _scoped_types_arg(["queries", "modules"]) == ["queries", "modules"]

    def test_strips_blanks(self):
        assert _scoped_types_arg(["  forms  ", "", "  "]) == "forms"


class TestImportObjectsScoped:
    def test_object_types_routes_to_import_by_type(self, tmp_path):
        from msaccess_vcs_mcp.tools import vcs_import_objects

        with _patch_import_tool(tmp_path) as (db_path, src_path, mock_addin, _):
            log = _write_log(src_path, "Merge_20260807_120000_000.log")
            mock_addin.call_sync.return_value = json.dumps(
                {"success": True, "logPath": str(log)}
            )
            result = asyncio.run(
                _unwrap_sync(vcs_import_objects)(
                    db_path, str(src_path), object_types=["modules"]
                )
            )

        mock_addin.call_sync.assert_called_once_with("ImportByType", "modules", False)
        mock_addin.call_async.assert_not_called()
        mock_addin.merge_build.assert_not_called()
        assert result["success"] is True
        assert result["log_path"] == str(log)

    def test_full_import_forwarded(self, tmp_path):
        from msaccess_vcs_mcp.tools import vcs_import_objects

        payload = json.dumps({"success": True, "logPath": "C:\\x.log"})
        with _patch_import_tool(tmp_path, call_sync_result=payload) as (
            db_path,
            src_path,
            mock_addin,
            _,
        ):
            asyncio.run(
                _unwrap_sync(vcs_import_objects)(
                    db_path,
                    str(src_path),
                    object_types=["queries", "forms"],
                    full_import=True,
                )
            )

        mock_addin.call_sync.assert_called_once_with(
            "ImportByType", ["queries", "forms"], True
        )

    def test_none_object_types_uses_async_merge(self, tmp_path):
        from msaccess_vcs_mcp.tools import vcs_import_objects

        with _patch_import_tool(
            tmp_path, async_result={"async": True, "timeout_ms": 1000}
        ) as (db_path, src_path, mock_addin, op_manager):
            _write_log(src_path, "Merge_20260807_120000_000.log")
            result = asyncio.run(
                _unwrap_sync(vcs_import_objects)(db_path, str(src_path))
            )

        mock_addin.call_sync.assert_not_called()
        mock_addin.call_async.assert_called_once()
        assert mock_addin.call_async.call_args[0][1] == "MergeBuild"
        assert result["success"] is True

    def test_addin_validation_error_propagates(self, tmp_path):
        from msaccess_vcs_mcp.tools import vcs_import_objects

        payload = json.dumps(
            {
                "success": False,
                "error": "Import not supported for component type(s): Table Data.",
            }
        )
        with _patch_import_tool(tmp_path, call_sync_result=payload) as (
            db_path,
            src_path,
            _,
            _,
        ):
            result = asyncio.run(
                _unwrap_sync(vcs_import_objects)(
                    db_path, str(src_path), object_types=["table_data"]
                )
            )

        assert result["success"] is False
        assert "Table Data" in result["error"]

    def test_empty_object_types_rejected(self, tmp_path):
        from msaccess_vcs_mcp.tools import vcs_import_objects

        with _patch_import_tool(tmp_path, call_sync_result="{}") as (
            db_path,
            src_path,
            mock_addin,
            _,
        ):
            result = asyncio.run(
                _unwrap_sync(vcs_import_objects)(
                    db_path, str(src_path), object_types=["", "  "]
                )
            )

        assert result["success"] is False
        mock_addin.call_sync.assert_not_called()


class TestExportDatabaseScoped:
    def test_object_types_routes_to_export_by_type(self, tmp_path):
        from msaccess_vcs_mcp.tools import vcs_export_database

        with _patch_export_tool(tmp_path) as (db_path, export_path, mock_addin, _):
            log = _write_log(export_path, "Export_20260807_120000_000.log")
            mock_addin.call_sync.return_value = json.dumps(
                {"success": True, "logPath": str(log)}
            )
            result = asyncio.run(
                _unwrap_sync(vcs_export_database)(
                    db_path, str(export_path), object_types=["modules"]
                )
            )

        mock_addin.call_sync.assert_called_once_with("ExportByType", "modules", False)
        mock_addin.call_async.assert_not_called()
        assert result["success"] is True
        assert result["log_path"] == str(log)

    def test_full_export_forwarded_to_scoped(self, tmp_path):
        from msaccess_vcs_mcp.tools import vcs_export_database

        payload = json.dumps({"success": True, "logPath": "C:\\x.log"})
        with _patch_export_tool(tmp_path, call_sync_result=payload) as (
            db_path,
            export_path,
            mock_addin,
            _,
        ):
            asyncio.run(
                _unwrap_sync(vcs_export_database)(
                    db_path,
                    str(export_path),
                    object_types=["queries"],
                    full_export=True,
                )
            )

        mock_addin.call_sync.assert_called_once_with("ExportByType", "queries", True)

    def test_none_object_types_uses_async_export(self, tmp_path):
        from msaccess_vcs_mcp.tools import vcs_export_database

        with _patch_export_tool(
            tmp_path, async_result={"async": True, "timeout_ms": 1000}
        ) as (db_path, export_path, mock_addin, _):
            _write_log(export_path, "Export_20260807_120000_000.log")
            result = asyncio.run(
                _unwrap_sync(vcs_export_database)(db_path, str(export_path))
            )

        mock_addin.call_sync.assert_not_called()
        mock_addin.call_async.assert_called_once()
        assert mock_addin.call_async.call_args[0][1] == "Export"
        assert result["success"] is True

    def test_modules_only_no_longer_uses_export_vba(self, tmp_path):
        """Regression: modules-only must go through ExportByType, not ExportVBA."""
        from msaccess_vcs_mcp.tools import vcs_export_database

        payload = json.dumps({"success": True, "logPath": "C:\\x.log"})
        with _patch_export_tool(tmp_path, call_sync_result=payload) as (
            db_path,
            export_path,
            mock_addin,
            _,
        ):
            asyncio.run(
                _unwrap_sync(vcs_export_database)(
                    db_path, str(export_path), object_types=["modules"]
                )
            )

        mock_addin.call_async.assert_not_called()
        assert mock_addin.call_sync.call_args[0][0] == "ExportByType"

    def test_addin_validation_error_propagates(self, tmp_path):
        from msaccess_vcs_mcp.tools import vcs_export_database

        payload = json.dumps(
            {"success": False, "error": "Unknown object type(s): bogons."}
        )
        with _patch_export_tool(tmp_path, call_sync_result=payload) as (
            db_path,
            export_path,
            _,
            _,
        ):
            result = asyncio.run(
                _unwrap_sync(vcs_export_database)(
                    db_path, str(export_path), object_types=["bogons"]
                )
            )

        assert result["success"] is False
        assert "bogons" in result["error"]
