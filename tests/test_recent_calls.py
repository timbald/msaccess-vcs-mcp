"""Tests for vcs_get_recent_calls and usage log tail reading."""

from __future__ import annotations

import json
from pathlib import Path

import msaccess_vcs_mcp.tools as tools_module
from msaccess_vcs_mcp.usage_logging import read_recent_tool_calls


def test_read_recent_tool_calls_returns_newest_first_reversed(tmp_path, monkeypatch):
    log_file = tmp_path / "usage.jsonl"
    entries = [
        {"event": "tool_call", "tool": "vcs_list_objects", "success": True},
        {"event": "code_execution", "tool": "vcs_run_vba"},
        {"event": "tool_call", "tool": "vcs_call_vba", "success": False},
        {"event": "tool_call", "tool": "vcs_run_tests", "success": True},
    ]
    log_file.write_text(
        "\n".join(json.dumps(e) for e in entries) + "\n",
        encoding="utf-8",
    )

    monkeypatch.setattr(
        "msaccess_vcs_mcp.usage_logging.get_log_file_path",
        lambda: log_file,
    )

    recent = read_recent_tool_calls(limit=2)
    assert len(recent) == 2
    assert recent[0]["tool"] == "vcs_call_vba"
    assert recent[1]["tool"] == "vcs_run_tests"


def test_read_recent_tool_calls_tolerates_malformed_lines(tmp_path, monkeypatch):
    log_file = tmp_path / "usage.jsonl"
    log_file.write_text(
        '{"event":"tool_call","tool":"ok"}\n'
        "not json\n"
        '{"event":"tool_call","tool":"last"}\n',
        encoding="utf-8",
    )
    monkeypatch.setattr(
        "msaccess_vcs_mcp.usage_logging.get_log_file_path",
        lambda: log_file,
    )

    recent = read_recent_tool_calls(limit=10)
    assert len(recent) == 2
    assert recent[-1]["tool"] == "last"


def test_read_recent_tool_calls_missing_file(monkeypatch):
    monkeypatch.setattr(
        "msaccess_vcs_mcp.usage_logging.get_log_file_path",
        lambda: None,
    )
    assert read_recent_tool_calls() == []


def _call_recent(limit=5):
    fn = tools_module.vcs_get_recent_calls
    while hasattr(fn, "__wrapped__"):
        fn = fn.__wrapped__
    return fn(limit=limit)


def test_vcs_get_recent_calls_tool(monkeypatch, tmp_path):
    log_file = tmp_path / "usage.jsonl"
    log_file.write_text(
        json.dumps({"event": "tool_call", "tool": "vcs_export_database"}) + "\n",
        encoding="utf-8",
    )
    monkeypatch.setattr(
        "msaccess_vcs_mcp.usage_logging.get_log_file_path",
        lambda: log_file,
    )
    monkeypatch.setattr(
        "msaccess_vcs_mcp.usage_logging._initialize_logging",
        lambda: True,
    )

    result = _call_recent(limit=5)
    assert result["success"] is True
    assert result["count"] == 1
    assert result["entries"][0]["tool"] == "vcs_export_database"
    assert result["log_path"] == str(log_file)
