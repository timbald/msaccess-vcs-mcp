"""Durable registry of Access processes this server created.

``_owns_app`` is recomputed on every ``AccessConnection``. GetObject on a
left-open window would otherwise reclassify a server-created instance as
user-owned, after which the server could never close it. Ownership is
therefore recorded on disk, keyed by PID plus process create time so a
reused PID cannot inherit a stale claim. The server restarts often enough
that an in-memory set would orphan windows within a session.
"""

from __future__ import annotations

import ctypes
import json
import os
import sys
import threading
from dataclasses import dataclass, replace
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

from .process_qos import list_access_pids_or_none


PROCESS_QUERY_LIMITED_INFORMATION = 0x1000

_lock = threading.Lock()


class FILETIME(ctypes.Structure):
    _fields_ = [
        ("dwLowDateTime", ctypes.c_uint32),
        ("dwHighDateTime", ctypes.c_uint32),
    ]


@dataclass(frozen=True)
class OwnedInstance:
    pid: int
    database_path: str
    create_time: int | None = None
    created_at: str | None = None
    session_id: str | None = None
    loaded_addin: str | None = None

    def to_dict(self) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "pid": self.pid,
            "database_path": self.database_path,
        }
        if self.create_time is not None:
            payload["create_time"] = self.create_time
        if self.created_at:
            payload["created_at"] = self.created_at
        if self.session_id:
            payload["session_id"] = self.session_id
        if self.loaded_addin:
            payload["loaded_addin"] = self.loaded_addin
        return payload

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> OwnedInstance | None:
        try:
            pid = int(data["pid"])
        except (KeyError, TypeError, ValueError):
            return None
        path = data.get("database_path")
        if not isinstance(path, str) or not path:
            return None
        create_time = data.get("create_time")
        if create_time is not None:
            try:
                create_time = int(create_time)
            except (TypeError, ValueError):
                create_time = None
        created_at = data.get("created_at")
        if created_at is not None and not isinstance(created_at, str):
            created_at = None
        session_id = data.get("session_id")
        if session_id is not None and not isinstance(session_id, str):
            session_id = None
        loaded_addin = data.get("loaded_addin")
        if loaded_addin is not None and not isinstance(loaded_addin, str):
            loaded_addin = None
        return cls(
            pid=pid,
            database_path=path,
            create_time=create_time,
            created_at=created_at,
            session_id=session_id,
            loaded_addin=loaded_addin,
        )


def get_registry_path() -> Path:
    """Return the on-disk registry path.

    Override with ``ACCESS_VCS_OWNED_INSTANCES_PATH`` (used by tests).
    Default is ``~/.msaccess-vcs-mcp/owned-instances.json``.
    """
    override = os.environ.get("ACCESS_VCS_OWNED_INSTANCES_PATH", "").strip()
    if override:
        return Path(override)
    return Path.home() / ".msaccess-vcs-mcp" / "owned-instances.json"


def process_create_time(pid: int) -> int | None:
    """Return the process creation stamp for ``pid``, or None.

    Prefers ``win32process.GetProcessTimes``; falls back to ctypes
    ``GetProcessTimes``. The value is a FILETIME-style integer (100-ns
    intervals since 1601) when that representation is available.
    """
    if not pid or pid <= 0:
        return None
    value = _create_time_via_pywin32(pid)
    if value is not None:
        return value
    return _create_time_via_ctypes(pid)


def _coerce_filetime(value: Any) -> int | None:
    if value is None:
        return None
    if isinstance(value, int) and not isinstance(value, bool):
        return value
    if hasattr(value, "timestamp"):
        try:
            # datetime-like → FILETIME (100-ns since 1601-01-01).
            return int(value.timestamp() * 10_000_000) + 116444736000000000
        except Exception:
            pass
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _create_time_via_pywin32(pid: int) -> int | None:
    try:
        import win32api
        import win32con
        import win32process
    except ImportError:
        return None
    try:
        handle = win32api.OpenProcess(
            win32con.PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid)
        )
    except Exception:
        return None
    try:
        times = win32process.GetProcessTimes(handle)
        if isinstance(times, dict):
            creation = times.get("CreationTime")
        else:
            creation = times[0]
        return _coerce_filetime(creation)
    except Exception:
        return None
    finally:
        try:
            win32api.CloseHandle(handle)
        except Exception:
            pass


def _create_time_via_ctypes(pid: int) -> int | None:
    if sys.platform != "win32":
        return None
    try:
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.OpenProcess.argtypes = [
            ctypes.c_uint32,
            ctypes.c_bool,
            ctypes.c_uint32,
        ]
        kernel32.OpenProcess.restype = ctypes.c_void_p
        kernel32.GetProcessTimes.argtypes = [
            ctypes.c_void_p,
            ctypes.POINTER(FILETIME),
            ctypes.POINTER(FILETIME),
            ctypes.POINTER(FILETIME),
            ctypes.POINTER(FILETIME),
        ]
        kernel32.GetProcessTimes.restype = ctypes.c_bool
        kernel32.CloseHandle.argtypes = [ctypes.c_void_p]
        kernel32.CloseHandle.restype = ctypes.c_bool
    except Exception:
        return None

    handle = kernel32.OpenProcess(
        PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid)
    )
    if not handle:
        return None
    creation = FILETIME()
    exit_t = FILETIME()
    kernel = FILETIME()
    user = FILETIME()
    try:
        ok = kernel32.GetProcessTimes(
            handle,
            ctypes.byref(creation),
            ctypes.byref(exit_t),
            ctypes.byref(kernel),
            ctypes.byref(user),
        )
        if not ok:
            return None
        return (int(creation.dwHighDateTime) << 32) | int(creation.dwLowDateTime)
    except Exception:
        return None
    finally:
        try:
            kernel32.CloseHandle(handle)
        except Exception:
            pass


def _norm_path(path: str) -> str:
    return os.path.normcase(os.path.abspath(path))


def paths_match(a: str, b: str) -> bool:
    """Case-insensitive, normalised path comparison."""
    try:
        return _norm_path(a) == _norm_path(b)
    except (OSError, ValueError):
        return False


def paths_match_ignoring_extension(a: str, b: str) -> bool:
    """Match paths the way the add-in compares install vs development copies."""
    try:
        return os.path.splitext(_norm_path(a))[0] == os.path.splitext(_norm_path(b))[0]
    except (OSError, ValueError):
        return False


def _parse(raw: list[dict[str, Any]]) -> list[OwnedInstance]:
    records: list[OwnedInstance] = []
    for item in raw:
        parsed = OwnedInstance.from_dict(item)
        if parsed is not None:
            records.append(parsed)
    return records


def _load_raw() -> list[dict[str, Any]]:
    path = get_registry_path()
    if not path.is_file():
        return []
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    if isinstance(data, list):
        return [item for item in data if isinstance(item, dict)]
    if isinstance(data, dict) and isinstance(data.get("instances"), list):
        return [item for item in data["instances"] if isinstance(item, dict)]
    return []


def _save(records: list[OwnedInstance]) -> None:
    """Write the registry atomically.

    The temp name carries this process id: the MCP server and a CLI child
    server share one registry file, and a fixed temp name lets one
    process's partial write land under another's ``os.replace``.
    """
    path = get_registry_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {"instances": [record.to_dict() for record in records]}
    tmp = path.with_name(f"{path.name}.{os.getpid()}.tmp")
    try:
        tmp.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        os.replace(tmp, path)
    except OSError:
        try:
            tmp.unlink()
        except OSError:
            pass


def _save_if_changed(
    records: list[OwnedInstance], original: list[OwnedInstance]
) -> None:
    """Persist only a real change, so reads do not rewrite on every call."""
    if records != original:
        _save(records)


def _prune(records: list[OwnedInstance]) -> list[OwnedInstance]:
    """Drop records whose process is gone or whose PID was reused.

    When liveness cannot be determined the records are kept: a transient
    ``tasklist`` failure must not make the server forget every window it
    created and start treating them as the user's.
    """
    live = list_access_pids_or_none()
    if live is None:
        return list(records)
    kept: list[OwnedInstance] = []
    for record in records:
        if record.pid not in live:
            continue
        live_create = process_create_time(record.pid)
        if (
            record.create_time is not None
            and live_create is not None
            and record.create_time != live_create
        ):
            continue
        kept.append(record)
    return kept


def _load_pruned() -> list[OwnedInstance]:
    return _prune(_parse(_load_raw()))


def _identity_confirmed(record: OwnedInstance, create_time: int | None) -> bool:
    """True only when ``create_time`` proves this is the recorded process.

    Anything less is treated as unconfirmed. Ownership drives whether the
    server may close a window, so an unreadable create time has to mean
    "not ours" rather than "probably ours".
    """
    if record.create_time is None or create_time is None:
        return False
    return record.create_time == create_time


def list_owned() -> list[OwnedInstance]:
    """Return live owned-instance records, pruning dead or reused PIDs."""
    with _lock:
        original = _parse(_load_raw())
        records = _prune(original)
        _save_if_changed(records, original)
        return list(records)


def register_owned(
    pid: int,
    database_path: str,
    create_time: int | None = None,
    session_id: str | None = None,
) -> OwnedInstance | None:
    """Record that this server created ``pid`` for ``database_path``."""
    if not pid or pid <= 0:
        return None
    if create_time is None:
        create_time = process_create_time(pid)
    if session_id is None:
        try:
            from ..config import get_session_id

            session_id = get_session_id()
        except Exception:
            session_id = None
    record = OwnedInstance(
        pid=int(pid),
        database_path=database_path,
        create_time=create_time,
        created_at=datetime.now(timezone.utc).isoformat(),
        session_id=session_id,
    )
    with _lock:
        records = [item for item in _load_pruned() if item.pid != record.pid]
        records.append(record)
        _save(records)

    try:
        from ..usage_logging import log_diagnostic_event

        log_diagnostic_event(
            "owned_instance_registered",
            pid=record.pid,
            database=str(database_path),
            create_time=create_time,
        )
    except Exception:
        pass
    return record


def unregister_owned(pid: int, create_time: int | None = None) -> bool:
    """Drop a registry claim. Returns True when a record was removed.

    A ``create_time`` that contradicts the stored one means the PID has
    been reused, so the record belongs to a different process and stays.
    """
    removed = False
    with _lock:
        records = _load_pruned()
        kept: list[OwnedInstance] = []
        for record in records:
            if record.pid != pid:
                kept.append(record)
                continue
            if (
                create_time is not None
                and record.create_time is not None
                and record.create_time != create_time
            ):
                kept.append(record)
                continue
            removed = True
        if removed:
            _save(kept)
    if removed:
        try:
            from ..usage_logging import log_diagnostic_event

            log_diagnostic_event("owned_instance_unregistered", pid=int(pid))
        except Exception:
            pass
    return removed


def is_owned(pid: int, create_time: int | None = None) -> bool:
    """True when ``pid`` is a live server-created Access process.

    Requires a confirmed ``create_time`` match. Without that proof the
    answer is False, which fails in the safe direction: the server treats
    the window as the user's and never closes it.
    """
    if not pid or pid <= 0:
        return False
    with _lock:
        original = _parse(_load_raw())
        records = _prune(original)
        _save_if_changed(records, original)
        for record in records:
            if record.pid == pid:
                return _identity_confirmed(record, create_time)
        return False


def note_loaded_addin(pid: int, addin_path: str) -> bool:
    """Record that ``pid`` has ``addin_path`` loaded as a library.

    A loaded add-in locks its file, so a rebuild that replaces it has to
    close the instances holding it -- whatever database they have open.
    Only updates instances already known to be server-created.
    """
    if not pid or pid <= 0 or not addin_path:
        return False
    updated = False
    with _lock:
        original = _parse(_load_raw())
        records = _prune(original)
        changed: list[OwnedInstance] = []
        for record in records:
            if record.pid == pid and not paths_match(
                record.loaded_addin or "", addin_path
            ):
                changed.append(replace(record, loaded_addin=addin_path))
                updated = True
            else:
                changed.append(record)
        if updated:
            _save(changed)
    return updated


def owned_records_for_paths(
    paths: Iterable[str],
    addin_paths: Iterable[str] = (),
) -> list[OwnedInstance]:
    """Return live owned records that hold any of the given files.

    ``paths`` matches a record's open database exactly. ``addin_paths``
    additionally matches a record's loaded add-in and ignores the file
    extension, mirroring the add-in's own install comparison: a compiled
    install is a ``.accde`` built from the same ``.accda``. Extension-blind
    matching is deliberately not applied to ordinary databases, where
    ``Reports.accdb`` and ``Reports.mdb`` are unrelated files.
    """
    targets = [path for path in paths if path]
    addin_targets = [path for path in addin_paths if path]
    if not targets and not addin_targets:
        return []

    matches: list[OwnedInstance] = []
    for record in list_owned():
        holds_db = any(paths_match(record.database_path, t) for t in targets)
        holds_addin = any(
            paths_match_ignoring_extension(record.database_path, t)
            or (
                record.loaded_addin
                and paths_match_ignoring_extension(record.loaded_addin, t)
            )
            for t in addin_targets
        )
        if holds_db or holds_addin:
            matches.append(record)
    return matches
