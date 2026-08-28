"""Tests for monotonic MCP progress reporting."""

from __future__ import annotations

import asyncio
from unittest.mock import AsyncMock, MagicMock

import pytest

from msaccess_vcs_mcp.operation_manager import (
    MonotonicProgressReporter,
    OperationManager,
)


@pytest.fixture(autouse=True)
def _reset_manager():
    OperationManager._instance = None
    yield
    OperationManager._instance = None


def test_format_message_keeps_vba_counts_in_text():
    assert MonotonicProgressReporter.format_message("queries", 28, 30) == (
        "queries (28/30)"
    )
    assert MonotonicProgressReporter.format_message("", 4, 10) == "(4/10)"
    assert MonotonicProgressReporter.format_message("starting") == "starting"
    assert MonotonicProgressReporter.format_message("log", -1, -1) == "log"
    assert MonotonicProgressReporter.format_message("queries", "28", "30") == (
        "queries (28/30)"
    )


def test_reporter_is_strictly_increasing():
    ctx = MagicMock()
    ctx.report_progress = AsyncMock()
    reporter = MonotonicProgressReporter()

    async def _run():
        await reporter.emit(ctx, message="queries", vba_progress=28, vba_total=30)
        await reporter.emit(ctx, message="modules", vba_progress=1, vba_total=50)
        await reporter.emit(ctx, message="exported forms")

    asyncio.run(_run())

    progresses = [call.kwargs["progress"] for call in ctx.report_progress.await_args_list]
    assert progresses == [1.0, 2.0, 3.0]
    assert all(call.kwargs["total"] is None for call in ctx.report_progress.await_args_list)
    assert ctx.report_progress.await_args_list[0].kwargs["message"] == "queries (28/30)"
    assert ctx.report_progress.await_args_list[1].kwargs["message"] == "modules (1/50)"
    assert ctx.report_progress.await_args_list[2].kwargs["message"] == "exported forms"


def test_wait_for_completion_forwards_ctx_and_stays_monotonic():
    manager = OperationManager.get_instance()
    operation_id, _queue = manager.register_operation(timeout_ms=5000)
    ctx = MagicMock()
    ctx.report_progress = AsyncMock()

    async def _run():
        async def _feed():
            await asyncio.sleep(0)
            manager.route_callback(operation_id, {
                "type": "progress",
                "progress": 28,
                "total": 30,
                "message": "queries",
            })
            manager.route_callback(operation_id, {
                "type": "log",
                "message": "Skipping unchanged object",
            })
            manager.route_callback(operation_id, {
                "type": "progress",
                "progress": 1,
                "total": 50,
                "message": "modules",
            })
            manager.route_callback(operation_id, {
                "type": "complete",
                "message": "done",
                "log_path": r"C:\src\logs\Export_1.log",
            })

        feeder = asyncio.create_task(_feed())
        result = await manager.wait_for_completion(operation_id, ctx=ctx, timeout_seconds=2)
        await feeder
        return result

    result = asyncio.run(_run())
    assert result["success"] is True
    assert result["log_messages"] == ["Skipping unchanged object"]
    progresses = [call.kwargs["progress"] for call in ctx.report_progress.await_args_list]
    assert progresses == [1.0, 2.0, 3.0]
    assert progresses == sorted(progresses)
    assert all(call.kwargs["total"] is None for call in ctx.report_progress.await_args_list)


def test_wait_for_completion_without_ctx_still_collects_logs():
    manager = OperationManager.get_instance()
    operation_id, _queue = manager.register_operation()

    async def _run():
        manager.route_callback(operation_id, {"type": "log", "message": "hello"})
        manager.route_callback(operation_id, {"type": "complete", "message": "ok"})
        return await manager.wait_for_completion(operation_id, ctx=None, timeout_seconds=2)

    result = asyncio.run(_run())
    assert result["success"] is True
    assert result["log_messages"] == ["hello"]


def test_wait_for_completion_forwards_results_path_on_error():
    """Failed test runs post type error but still attach results_path."""
    manager = OperationManager.get_instance()
    operation_id, _queue = manager.register_operation(timeout_ms=5000)

    async def _run():
        async def _feed():
            await asyncio.sleep(0)
            manager.route_callback(operation_id, {
                "type": "error",
                "message": "Operation failed",
                "log_path": r"C:\src\logs\TestRun_1.log",
                "results_path": r"C:\src\logs\TestResults_1.json",
            })

        feeder = asyncio.create_task(_feed())
        result = await manager.wait_for_completion(operation_id, timeout_seconds=2)
        await feeder
        return result

    result = asyncio.run(_run())
    assert result["success"] is False
    assert result["error"] == "Operation failed"
    assert result["log_path"] == r"C:\src\logs\TestRun_1.log"
    assert result["results_path"] == r"C:\src\logs\TestResults_1.json"
