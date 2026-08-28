"""Tests for EcoQoS-off / Above Normal promotion of MCP-launched Access."""

from __future__ import annotations

import sys
from unittest.mock import MagicMock, patch

from msaccess_vcs_mcp.access_com import process_qos


def test_prefer_full_power_pid_skips_non_windows(monkeypatch):
    monkeypatch.setattr(process_qos.sys, "platform", "linux")
    assert process_qos.prefer_full_power_pid(1234) is False


def test_prefer_full_power_pid_skips_invalid_pid():
    assert process_qos.prefer_full_power_pid(0) is False
    assert process_qos.prefer_full_power_pid(-1) is False


def test_prefer_full_power_pid_opens_and_sets_qos(monkeypatch):
    monkeypatch.setattr(process_qos.sys, "platform", "win32")
    kernel32 = MagicMock()
    kernel32.OpenProcess.return_value = 42
    kernel32.SetProcessInformation.return_value = 1
    kernel32.SetPriorityClass.return_value = 1
    kernel32.CloseHandle.return_value = 1

    with (
        patch.object(process_qos, "_kernel32", return_value=kernel32),
        patch.object(process_qos, "log_diagnostic_event", create=True),
        patch(
            "msaccess_vcs_mcp.usage_logging.log_diagnostic_event",
            MagicMock(),
        ),
    ):
        assert process_qos.prefer_full_power_pid(99) is True

    kernel32.OpenProcess.assert_called_once()
    assert kernel32.SetProcessInformation.called
    kernel32.SetPriorityClass.assert_called_once_with(
        42, process_qos.ABOVE_NORMAL_PRIORITY_CLASS
    )
    kernel32.CloseHandle.assert_called_once_with(42)


def test_prefer_full_power_pid_open_failure(monkeypatch):
    monkeypatch.setattr(process_qos.sys, "platform", "win32")
    kernel32 = MagicMock()
    kernel32.OpenProcess.return_value = 0
    with patch.object(process_qos, "_kernel32", return_value=kernel32):
        assert process_qos.prefer_full_power_pid(99) is False
    kernel32.SetProcessInformation.assert_not_called()
    kernel32.CloseHandle.assert_not_called()


def test_pid_from_access_app_rejects_mock_hwnd():
    app = MagicMock()
    assert process_qos.pid_from_access_app(app) is None


def test_pid_from_access_app_reads_hwnd(monkeypatch):
    app = MagicMock()
    app.hWndAccessApp.return_value = 12345
    fake = MagicMock()
    fake.GetWindowThreadProcessId.return_value = (1, 4321)
    monkeypatch.setitem(sys.modules, "win32process", fake)
    assert process_qos.pid_from_access_app(app) == 4321


def test_prefer_full_power_if_created_skips_user_instance():
    app = MagicMock()
    app.CurrentDb.return_value = MagicMock()
    with patch.object(process_qos, "prefer_full_power_app") as promote:
        assert process_qos.prefer_full_power_if_created(app) is False
        promote.assert_not_called()


def test_prefer_full_power_if_created_promotes_empty_instance():
    app = MagicMock()
    app.CurrentDb.return_value = None
    with patch.object(process_qos, "prefer_full_power_app", return_value=True) as promote:
        assert process_qos.prefer_full_power_if_created(app) is True
        promote.assert_called_once_with(app)


def test_list_access_pids_parses_tasklist(monkeypatch):
    monkeypatch.setattr(process_qos.sys, "platform", "win32")
    completed = MagicMock()
    completed.stdout = '"MSACCESS.EXE","1111","Console","1","50 K"\n"MSACCESS.EXE","2222","Console","1","50 K"\n'
    with patch.object(process_qos.subprocess, "run", return_value=completed):
        assert process_qos.list_access_pids() == {1111, 2222}


def test_list_access_pids_empty(monkeypatch):
    monkeypatch.setattr(process_qos.sys, "platform", "win32")
    completed = MagicMock()
    completed.stdout = "INFO: No tasks are running which match the specified criteria.\n"
    with patch.object(process_qos.subprocess, "run", return_value=completed):
        assert process_qos.list_access_pids() == set()


def test_prefer_full_power_new_access_skips_known():
    with (
        patch.object(process_qos, "list_access_pids", return_value={10, 20, 30}),
        patch.object(process_qos, "prefer_full_power_pid") as promote,
    ):
        seen = process_qos.prefer_full_power_new_access({10, 20})
        promote.assert_called_once_with(30)
        assert seen == {10, 20, 30}
