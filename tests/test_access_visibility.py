"""Tests for the rule that database-holding Access instances stay visible.

A hidden instance strands the user: an error dialog, a VBA break, or a trust
prompt blocks every later call with nothing on screen to explain why, and
nobody can dismiss what they cannot see.  Visibility has to arrive *after* the
database opens, because showing the window can set ``UserControl`` and startup
code reads that flag to decide whether a person is watching.
"""

from unittest.mock import MagicMock, patch

import pytest

from msaccess_vcs_mcp.access_com import connection as conn_mod
from msaccess_vcs_mcp.access_com.connection import (
    AccessConnection,
    ensure_access_visible,
    open_current_database,
)

DB = r"C:\Projects\other\Some Database.accdb"


class _FakeApp:
    """Access Application stand-in that records the order of state changes."""

    def __init__(self, current_db_name=None, user_control=False):
        self.events: list[str] = []
        self._current_db_name = current_db_name
        self._user_control = user_control
        self._visible = False

    @property
    def UserControl(self):
        return self._user_control

    @UserControl.setter
    def UserControl(self, value):
        self._user_control = value
        self.events.append(f"UserControl={value}")

    @property
    def Visible(self):
        return self._visible

    @Visible.setter
    def Visible(self, value):
        self._visible = value
        self.events.append(f"Visible={value}")

    def CurrentDb(self):
        if self._current_db_name is None:
            return None
        db = MagicMock()
        db.Name = self._current_db_name
        return db

    def OpenCurrentDatabase(self, path):
        self.events.append(f"Open={path}")
        self._current_db_name = path


@pytest.fixture(autouse=True)
def _com_available():
    with patch.object(conn_mod, "COM_AVAILABLE", True):
        yield


class TestOpenCurrentDatabase:
    def test_shows_the_window_only_once_the_database_is_open(self):
        """AutoExec must run as automation; the user sees the window after."""
        app = _FakeApp(user_control=True)

        open_current_database(app, DB)

        assert app.events == [
            "UserControl=False",
            f"Open={DB}",
            "UserControl=True",
            "Visible=True",
        ]

    def test_leaves_user_control_alone_when_never_set(self):
        app = _FakeApp(user_control=False)

        open_current_database(app, DB)

        assert app.events == [f"Open={DB}", "Visible=True"]

    def test_restores_user_control_when_the_open_fails(self):
        """A failed open must not leave the instance looking like automation."""
        app = _FakeApp(user_control=True)
        app.OpenCurrentDatabase = MagicMock(side_effect=Exception("in use"))

        with pytest.raises(Exception, match="in use"):
            open_current_database(app, DB)

        assert app.UserControl is True


class TestEnsureAccessVisible:
    def test_reports_failure_without_raising(self):
        """Showing the window is never worth failing an operation over."""
        app = MagicMock()
        type(app).Visible = property(
            lambda self: False,
            lambda self, value: (_ for _ in ()).throw(Exception("refused")),
        )

        assert ensure_access_visible(app) is False


class TestAccessConnectionVisibility:
    def test_instance_we_create_becomes_visible(self):
        app = _FakeApp()

        with (
            patch.object(conn_mod, "win32com") as mock_win32com,
            patch.object(conn_mod, "gencache") as mock_gencache,
            patch(
                "msaccess_vcs_mcp.access_com.process_qos.list_access_pids_or_none",
                return_value=set(),
            ),
            patch(
                "msaccess_vcs_mcp.access_com.process_qos.pid_from_access_app",
                return_value=4242,
            ),
        ):
            mock_win32com.client.GetObject.side_effect = Exception("no moniker")
            mock_gencache.EnsureDispatch.return_value = app

            AccessConnection(DB)._get_access_app()

        assert app.Visible is True
        assert app.UserControl is True
        assert app.events.index(f"Open={DB}") < app.events.index("Visible=True")
        assert app.events.index("Visible=True") < app.events.index("UserControl=True")
        open_idx = app.events.index(f"Open={DB}")
        assert all(not event.startswith("UserControl=") for event in app.events[:open_idx])

    def test_instance_we_attach_to_becomes_visible(self, monkeypatch):
        """A user's window is shown but left as the user's.

        ``UserControl`` stays untouched: raising it is how the server marks
        an instance it created, and doing so here would be a lie the
        closure path could act on.
        """
        app = _FakeApp(current_db_name=DB)
        monkeypatch.setattr(conn_mod, "access_instance_is_live", lambda _p: True)

        with patch.object(conn_mod, "win32com") as mock_win32com:
            mock_win32com.client.GetObject.return_value = app

            AccessConnection(DB)._get_access_app()

        assert app.Visible is True
        assert app.UserControl is False

    def test_instance_launched_by_moniker_bind_is_ours(self, monkeypatch):
        """Binding a moniker starts Access when nothing has the file open.

        That process is the server's, so it gets the interactive treatment
        and a registry claim -- otherwise nothing could ever close it.
        """
        app = _FakeApp(current_db_name=DB)
        monkeypatch.setattr(conn_mod, "access_instance_is_live", lambda _p: False)

        with (
            patch.object(conn_mod, "win32com") as mock_win32com,
            patch(
                "msaccess_vcs_mcp.access_com.process_qos.pid_from_access_app",
                return_value=4242,
            ),
            patch("msaccess_vcs_mcp.access_com.process_qos.prefer_full_power_app"),
        ):
            mock_win32com.client.GetObject.return_value = app

            conn = AccessConnection(DB)
            conn._get_access_app()

        assert conn._owns_app is True
        assert app.Visible is True
        assert app.UserControl is True


class TestAccessConnectionClose:
    def _owned_connection(self, app):
        conn = AccessConnection.__new__(AccessConnection)
        conn._app = app
        conn._db = None
        conn._owns_app = True
        conn._owns_db = False
        conn._db_opened_via_getobject = False
        conn._db_opened_as_current = False
        return conn

    def test_owned_instance_is_left_open_by_default(self):
        app = MagicMock()
        self._owned_connection(app).close()
        app.CloseCurrentDatabase.assert_not_called()
        app.Quit.assert_not_called()

    def test_leave_access_open_false_quits(self, monkeypatch):
        monkeypatch.setenv("ACCESS_VCS_LEAVE_ACCESS_OPEN", "false")
        app = MagicMock()
        self._owned_connection(app).close()
        app.CloseCurrentDatabase.assert_called_once()
        app.Quit.assert_called_once()

    def test_leave_access_open_true_skips_quit(self, monkeypatch):
        monkeypatch.setenv("ACCESS_VCS_LEAVE_ACCESS_OPEN", "true")
        app = MagicMock()
        self._owned_connection(app).close()
        app.CloseCurrentDatabase.assert_not_called()
        app.Quit.assert_not_called()
