"""Tests for the streaming msaccess-vcs CLI."""

from __future__ import annotations

import json
from types import SimpleNamespace

from msaccess_vcs_mcp.cli import (
    arguments_for,
    build_parser,
    format_progress_line,
    main,
    parse_result_payload,
    print_progress,
    result_succeeded,
    startup_message,
    stdio_server_environment,
)


def test_arguments_for_each_subcommand():
    parser = build_parser()
    export = parser.parse_args(["export", r"C:\db.accdb", r"C:\src", "--full"])
    assert arguments_for(export) == (
        "vcs_export_database",
        {
            "database_path": r"C:\db.accdb",
            "output_dir": r"C:\src",
            "full_export": True,
        },
    )

    merge = parser.parse_args(["merge", r"C:\db.accdb", r"C:\src"])
    assert arguments_for(merge) == (
        "vcs_import_objects",
        {"database_path": r"C:\db.accdb", "source_dir": r"C:\src"},
    )

    rebuild = parser.parse_args([
        "rebuild-database",
        r"C:\src",
        r"C:\out.accdb",
        "--template",
        r"C:\blank.accdb",
    ])
    name, payload = arguments_for(rebuild)
    assert name == "vcs_rebuild_database"
    assert payload["template_path"] == r"C:\blank.accdb"

    addin = parser.parse_args(["rebuild-addin", r"C:\src", "--timeout", "90"])
    assert arguments_for(addin) == (
        "vcs_rebuild_addin",
        {"source_dir": r"C:\src", "timeout_seconds": 90.0},
    )


def test_startup_message():
    assert startup_message("rebuild-addin") == "Starting rebuild-addin..."
    assert startup_message("export") == "Starting export..."


def test_format_progress_line():
    assert format_progress_line(3, None, "building") == "building"
    assert format_progress_line(2, 10, "queries") == "queries"
    assert format_progress_line(1, None, None) == ""
    assert format_progress_line(1, None, "  ") == ""


def test_print_progress_skips_blank(capsys):
    print_progress(1, None, None)
    print_progress(2, None, "   ")
    print_progress(3, None, "Importing modules...")
    assert capsys.readouterr().out.splitlines() == ["Importing modules..."]


def test_result_succeeded_from_json():
    assert result_succeeded({"success": True})
    assert not result_succeeded({"success": False, "error": "nope"})
    assert result_succeeded({"status": "complete"})
    assert not result_succeeded({"cancelled": True})


def test_parse_result_payload():
    assert parse_result_payload('{"success": true}') == {"success": True}
    assert parse_result_payload("not json") == "not json"


def test_stdio_server_skips_access_session_cleanup(monkeypatch):
    monkeypatch.setenv("EXISTING_VALUE", "kept")
    env = stdio_server_environment()
    assert env["EXISTING_VALUE"] == "kept"
    assert env["ACCESS_VCS_SKIP_SESSION_CLEANUP"] == "true"


def test_main_streams_progress_then_json(capsys):
    async def factory(name, arguments, on_progress):
        assert name == "vcs_rebuild_addin"
        assert arguments["source_dir"] == r"C:\src"
        await on_progress(1, None, "Rebuild launched")
        await on_progress(2, None, "Rebuild complete")
        return SimpleNamespace(
            content=[SimpleNamespace(text=json.dumps({"success": True, "status": "complete"}))],
            isError=False,
        )

    code = main(["rebuild-addin", r"C:\src"], session_factory=factory)
    captured = capsys.readouterr()
    lines = [line for line in captured.out.splitlines() if line]
    assert lines[0] == "Starting rebuild-addin..."
    assert lines[1] == "Rebuild launched"
    assert lines[2] == "Rebuild complete"
    assert json.loads("\n".join(lines[3:]))["status"] == "complete"
    assert code == 0


def test_main_failure_exit_code(capsys):
    async def factory(_name, _arguments, on_progress):
        await on_progress(1, None, "Rebuild refused")
        return SimpleNamespace(
            content=[SimpleNamespace(text=json.dumps({
                "success": False,
                "status": "refused",
                "error": "busy",
            }))],
            isError=False,
        )

    code = main(["rebuild-addin", r"C:\src"], session_factory=factory)
    assert code == 1
    out = capsys.readouterr().out
    assert "Starting rebuild-addin..." in out
    assert "Rebuild refused" in out
