"""Tests for vcs_call_vba's timeout and rebuild-status correlation.

Refusing the installed add-in as a target is the ``vcs_tool`` wrapper's job now and
lives in ``test_installed_addin_guard.py``.

Paths here are synthetic. They only have to look like a real install and a real
repository to the code under test; nothing is read from disk.
"""

from __future__ import annotations

import json
import threading
import time
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

import msaccess_vcs_mcp.tools as tools_module


INSTALLED_ADDIN = r"C:\Users\Example\AppData\Roaming\MSAccessVCS\Version Control.accda"
REPO_ADDIN = r"C:\Projects\msaccess-vcs-addin\Version Control.accda"
SOURCE_DIR = r"C:\Projects\msaccess-vcs-addin\Version Control.accda.src"
RESOLVED_API = r"C:\Users\Example\AppData\Roaming\MSAccessVCS\Version Control.API"


def _call_vba(*args, **kwargs):
    fn = tools_module.vcs_call_vba
    while hasattr(fn, "__wrapped__"):
        fn = fn.__wrapped__
    return fn(*args, **kwargs)


def _fake_access(monkeypatch, run_result=None, run_side_effect=None):
    """Wire both the connection and the worker's ROT lookup to one fake Application.

    The dispatch runs on a worker thread that re-acquires Access from the Running
    Object Table rather than reusing the connection's proxy, because a COM pointer
    belongs to the apartment that made it. A test that stubs only AccessConnection
    therefore never reaches its own mock. Returns the fake so callers can assert on
    what it received.
    """
    app = MagicMock()
    if run_side_effect is not None:
        app.Run.side_effect = run_side_effect
    else:
        app.Run.return_value = run_result

    class _Conn:
        def __init__(self, *_args, **_kwargs):
            pass

        def __enter__(self):
            return self

        def __exit__(self, *args):
            return False

        def connect(self):
            return app, MagicMock()

    monkeypatch.setattr(tools_module, "AccessConnection", _Conn)
    monkeypatch.setattr(
        tools_module.VCSAddinIntegration,
        "_find_access_in_rot",
        staticmethod(lambda _db_path: app),
    )
    return app


@pytest.fixture(autouse=True)
def _bypass_access_gate(monkeypatch):
    """Unit tests call tool bodies directly — skip the real gate."""

    async def _immediate(tool, database, fn, is_async, /, *args, **kwargs):
        if is_async:
            return await fn(*args, **kwargs)
        return fn(*args, **kwargs)

    monkeypatch.setattr(
        tools_module,
        "get_access_gate",
        lambda: MagicMock(run_exclusive=_immediate),
    )


@pytest.fixture
def _config_installed_addin(monkeypatch):
    monkeypatch.setattr(
        tools_module,
        "get_config",
        lambda: {"ACCESS_VCS_ADDIN_PATH": INSTALLED_ADDIN},
    )


def test_development_copy_dispatches(_config_installed_addin, monkeypatch):
    """The development copy shares the install's file name but is a different file."""
    monkeypatch.setattr(
        tools_module,
        "validate_database_path",
        lambda p: Path(p),
    )
    monkeypatch.setattr(tools_module, "get_call_vba_timeout", lambda _t=None: 5.0)
    _fake_access(monkeypatch, run_result='{"version":"5.1.2"}')

    result = _call_vba(
        REPO_ADDIN,
        "VCS.API",
        ["GetVCSVersion"],
    )

    assert result["success"] is True


def test_call_vba_timeout_returns_structured_json(_config_installed_addin, monkeypatch, tmp_path):
    db = tmp_path / "host.accdb"
    db.write_bytes(b"x")
    monkeypatch.setattr(
        tools_module,
        "validate_database_path",
        lambda p: Path(p),
    )
    monkeypatch.setattr(tools_module, "get_call_vba_timeout", lambda _t=None: 0.05)

    def slow_run(*_a, **_k):
        time.sleep(0.2)
        return "never"

    _fake_access(monkeypatch, run_side_effect=slow_run)

    result = _call_vba(
        str(db),
        "VCS.API",
        ["GetVCSVersion"],
        timeout_seconds=0.05,
    )

    assert result["success"] is False
    assert result["timed_out"] is True
    assert result["error_pattern"] == "timeout"
    assert result["recoverable"] is True


def test_rebuild_snapshot_captured_before_dispatch(
    _config_installed_addin, monkeypatch, tmp_path
):
    source = tmp_path / "Version Control.accda.src"
    logs = source / "logs"
    logs.mkdir(parents=True)
    status_file = logs / "rebuild-status.json"
    status_file.write_text(
        json.dumps({
            "status": "complete",
            "updated": "2026-08-17 16:46:20",
            "phaseStarted": "2026-08-17 16:45:36",
        }),
        encoding="utf-8",
    )

    host = tmp_path / "host.accdb"
    host.write_bytes(b"x")
    monkeypatch.setattr(
        tools_module,
        "validate_database_path",
        lambda p: Path(p),
    )
    monkeypatch.setattr(tools_module, "get_call_vba_timeout", lambda _t=None: 5.0)
    _fake_access(monkeypatch, run_result='{"status":"launched"}')

    result = _call_vba(
        str(host),
        "VCS.API",
        ["RebuildAddIn", str(source)],
    )

    assert result["success"] is True
    assert result["rebuild_status_file"] == str(status_file)
    assert result["rebuild_status_before"]["status"] == "complete"
    assert result["rebuild_status_before"]["updated"] == "2026-08-17 16:46:20"
    assert "mtime" in result["rebuild_status_before"]


def test_rebuild_snapshot_included_on_timeout(
    _config_installed_addin, monkeypatch, tmp_path
):
    source = tmp_path / "src"
    logs = source / "logs"
    logs.mkdir(parents=True)
    (logs / "rebuild-status.json").write_text(
        json.dumps({"status": "complete", "updated": "old"}),
        encoding="utf-8",
    )

    host = tmp_path / "host.accdb"
    host.write_bytes(b"x")
    monkeypatch.setattr(
        tools_module,
        "validate_database_path",
        lambda p: Path(p),
    )
    monkeypatch.setattr(tools_module, "get_call_vba_timeout", lambda _t=None: 0.05)
    _fake_access(monkeypatch, run_side_effect=lambda *_a, **_k: time.sleep(0.2))

    result = _call_vba(
        str(host),
        "VCS.API",
        ["RebuildAddIn", str(source)],
        timeout_seconds=0.05,
    )

    assert result["timed_out"] is True
    assert result["rebuild_status_before"]["status"] == "complete"


def test_is_addin_api_resolved_name():
    with patch.object(
        tools_module,
        "get_config",
        return_value={"ACCESS_VCS_ADDIN_PATH": INSTALLED_ADDIN},
    ):
        assert tools_module._is_addin_api_resolved_name(RESOLVED_API)
        assert not tools_module._is_addin_api_resolved_name("SomeModule.Foo")


def test_describe_rebuild_attempt_correlates_phase_started():
    describe = tools_module._describe_rebuild_attempt
    launched = json.dumps({"status": "launched", "phaseStarted": "2026-08-21 09:15:00"})

    # The attempt stamped its own start, so the record on disk is now this run's.
    assert describe(launched, {"phaseStarted": "2026-08-17 16:45:36"}) == {
        "rebuild_phase_started": "2026-08-21 09:15:00",
        "rebuild_status_superseded": True,
    }

    # Same value as the snapshot: nothing new was written, so whatever is on disk
    # belongs to an earlier run and must not be read as this attempt's verdict.
    assert describe(launched, {"phaseStarted": "2026-08-21 09:15:00"}) == {
        "rebuild_phase_started": "2026-08-21 09:15:00",
        "rebuild_status_superseded": False,
    }

    assert describe(launched, None) == {"rebuild_phase_started": "2026-08-21 09:15:00"}
    assert describe(json.dumps({"status": "refused"}), None) == {}
    assert describe("not json", None) == {}
    assert describe(None, None) == {}


def test_snapshot_rebuild_status_reads_bom_encoded_file(tmp_path):
    """The add-in writes rebuild-status.json as UTF-8 with a BOM.

    Reading it as plain "utf-8" leaves the BOM in the string, json.load rejects it,
    and the snapshot degrades to read_error -- which silently disables the
    phaseStarted comparison that tells this attempt's record from a stale one.
    """
    logs = tmp_path / "logs"
    logs.mkdir()
    payload = {
        "status": "complete",
        "phaseStarted": "2026-08-21 13:12:15",
        "updated": "2026-08-21 13:13:04",
    }
    (logs / "rebuild-status.json").write_text(
        json.dumps(payload), encoding="utf-8-sig"
    )

    status_file, snapshot = tools_module._snapshot_rebuild_status(str(tmp_path))

    assert status_file == str(logs / "rebuild-status.json")
    assert "read_error" not in snapshot
    assert snapshot["status"] == "complete"
    assert snapshot["phaseStarted"] == "2026-08-21 13:12:15"

    # And the value has to be usable for the comparison it exists to serve.
    described = tools_module._describe_rebuild_attempt(
        json.dumps({"status": "launched", "phaseStarted": "2026-08-21 13:30:00"}),
        snapshot,
    )
    assert described["rebuild_status_superseded"] is True


def test_snapshot_rebuild_status_missing_and_malformed(tmp_path):
    logs = tmp_path / "logs"
    logs.mkdir()

    # No file at all: nothing to compare against, reported as absent rather than broken.
    assert tools_module._snapshot_rebuild_status(str(tmp_path))[1] is None

    (logs / "rebuild-status.json").write_text("{ truncated", encoding="utf-8")
    assert tools_module._snapshot_rebuild_status(str(tmp_path))[1]["read_error"] is True


def test_is_installed_addin_path_ignores_the_extension():
    with patch.object(
        tools_module,
        "get_config",
        return_value={"ACCESS_VCS_ADDIN_PATH": INSTALLED_ADDIN},
    ):
        check = tools_module._is_installed_addin_path

        assert check(INSTALLED_ADDIN)
        # A compiled install is the .accde built from the same .accda, and only one of
        # the two is ever named in ACCESS_VCS_ADDIN_PATH.
        assert check(INSTALLED_ADDIN.replace(".accda", ".accde"))
        assert check(INSTALLED_ADDIN.upper())
        # Same file name, different folder.
        assert not check(REPO_ADDIN)
        # A folder beside the install is not the install.
        assert not check(SOURCE_DIR)
