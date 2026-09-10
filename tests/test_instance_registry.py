"""Tests for the durable owned-Access PID registry."""

from __future__ import annotations

import json

from msaccess_vcs_mcp.access_com import instance_registry as registry


def _isolate_registry(tmp_path, monkeypatch, live_pids=None, create_times=None):
    """Point the registry at a temp file with a scripted process table.

    ``live_pids=None`` means the PID query itself failed, which the
    registry must treat differently from an empty set.
    """
    path = tmp_path / "owned-instances.json"
    monkeypatch.setenv("ACCESS_VCS_OWNED_INSTANCES_PATH", str(path))
    pids = None if live_pids is None else set(live_pids)
    monkeypatch.setattr(registry, "list_access_pids_or_none", lambda: pids)
    times = create_times or {}
    monkeypatch.setattr(registry, "process_create_time", lambda pid: times.get(pid))
    return path


def test_register_round_trip(tmp_path, monkeypatch):
    _isolate_registry(
        tmp_path, monkeypatch, live_pids={42}, create_times={42: 1000}
    )

    record = registry.register_owned(
        42, r"C:\db.accdb", create_time=1000, session_id="abcd"
    )

    assert record is not None
    assert record.pid == 42
    assert registry.is_owned(42, 1000) is True
    listed = registry.list_owned()
    assert len(listed) == 1
    assert listed[0].database_path == r"C:\db.accdb"
    assert listed[0].session_id == "abcd"


def test_pid_reuse_is_rejected(tmp_path, monkeypatch):
    _isolate_registry(
        tmp_path, monkeypatch, live_pids={42}, create_times={42: 2000}
    )
    registry.register_owned(42, r"C:\db.accdb", create_time=1000)

    assert registry.is_owned(42, 2000) is False
    assert registry.list_owned() == []


def test_dead_pid_is_pruned(tmp_path, monkeypatch):
    path = _isolate_registry(
        tmp_path, monkeypatch, live_pids=set(), create_times={}
    )
    path.write_text(
        json.dumps({
            "instances": [
                {
                    "pid": 99,
                    "database_path": r"C:\gone.accdb",
                    "create_time": 1,
                }
            ]
        }),
        encoding="utf-8",
    )

    assert registry.is_owned(99, 1) is False
    assert registry.list_owned() == []


def test_unregister_removes_matching_record(tmp_path, monkeypatch):
    _isolate_registry(
        tmp_path, monkeypatch, live_pids={7}, create_times={7: 5}
    )
    registry.register_owned(7, r"C:\db.accdb", create_time=5)

    assert registry.unregister_owned(7, 5) is True
    assert registry.is_owned(7, 5) is False


def test_addin_paths_match_ignoring_extension(tmp_path, monkeypatch):
    _isolate_registry(
        tmp_path, monkeypatch, live_pids={11}, create_times={11: 3}
    )
    registry.register_owned(
        11,
        r"C:\Repos\addin\Version Control.accda",
        create_time=3,
    )

    matches = registry.owned_records_for_paths(
        [], [r"C:\Repos\addin\Version Control.accde"]
    )
    assert len(matches) == 1
    assert matches[0].pid == 11


def test_database_paths_match_exactly(tmp_path, monkeypatch):
    """Extension-blind matching is for the add-in pair only.

    ``Reports.accdb`` and ``Reports.mdb`` are unrelated databases and a
    rebuild of one must not close a window holding the other.
    """
    _isolate_registry(
        tmp_path, monkeypatch, live_pids={12}, create_times={12: 4}
    )
    registry.register_owned(12, r"C:\Data\Reports.accdb", create_time=4)

    assert registry.owned_records_for_paths([r"C:\Data\Reports.mdb"]) == []
    assert len(registry.owned_records_for_paths([r"C:\Data\Reports.accdb"])) == 1


def test_loaded_addin_holder_is_matched(tmp_path, monkeypatch):
    """An instance with an unrelated database still locks the add-in."""
    _isolate_registry(
        tmp_path, monkeypatch, live_pids={13}, create_times={13: 5}
    )
    installed = r"C:\Users\me\AppData\Roaming\MSAccessVCS\Version Control.accda"
    registry.register_owned(13, r"C:\Data\Unrelated.accdb", create_time=5)
    assert registry.note_loaded_addin(13, installed) is True

    matches = registry.owned_records_for_paths([], [installed])
    assert [m.pid for m in matches] == [13]
    assert registry.owned_records_for_paths([installed]) == []


def test_failed_pid_query_preserves_registry(tmp_path, monkeypatch):
    """A transient tasklist failure must not orphan owned windows.

    Pruning them would make the server treat its own instances as the
    user's, and it would then refuse to close them for a rebuild.
    """
    path = _isolate_registry(
        tmp_path, monkeypatch, live_pids={5}, create_times={5: 9}
    )
    registry.register_owned(5, r"C:\db.accdb", create_time=9)

    monkeypatch.setattr(registry, "list_access_pids_or_none", lambda: None)

    assert registry.is_owned(5, 9) is True
    assert len(registry.list_owned()) == 1
    assert "\"pid\": 5" in path.read_text(encoding="utf-8").replace("'", '"')


def test_missing_live_create_time_does_not_claim_recorded_pid(
    tmp_path, monkeypatch
):
    _isolate_registry(
        tmp_path, monkeypatch, live_pids={8}, create_times={}
    )
    registry.register_owned(8, r"C:\db.accdb", create_time=50)

    assert registry.is_owned(8, None) is False


def test_record_without_create_time_is_never_owned(tmp_path, monkeypatch):
    """Ownership needs positive proof, not the absence of contradiction.

    A record with no creation stamp cannot be distinguished from a reused
    PID, so it must not authorize closing that process.
    """
    path = _isolate_registry(
        tmp_path, monkeypatch, live_pids={21}, create_times={21: 77}
    )
    path.write_text(
        json.dumps(
            {"instances": [{"pid": 21, "database_path": r"C:\db.accdb"}]}
        ),
        encoding="utf-8",
    )

    assert registry.is_owned(21, 77) is False
    assert registry.is_owned(21, None) is False


def test_reused_pid_is_not_unregistered(tmp_path, monkeypatch):
    _isolate_registry(
        tmp_path, monkeypatch, live_pids={31}, create_times={31: 1}
    )
    registry.register_owned(31, r"C:\db.accdb", create_time=1)

    assert registry.unregister_owned(31, 999) is False
    assert registry.is_owned(31, 1) is True


def test_reads_do_not_rewrite_an_unchanged_registry(tmp_path, monkeypatch):
    path = _isolate_registry(
        tmp_path, monkeypatch, live_pids={41}, create_times={41: 2}
    )
    registry.register_owned(41, r"C:\db.accdb", create_time=2)
    before = path.stat().st_mtime_ns

    registry.list_owned()
    registry.is_owned(41, 2)

    assert path.stat().st_mtime_ns == before
