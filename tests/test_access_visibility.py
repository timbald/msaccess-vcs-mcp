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

DB = r"C:\Repos\other\Some Database.accdb"


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
        ):
            mock_win32com.client.GetObject.side_effect = Exception("no moniker")
            mock_gencache.EnsureDispatch.return_value = app

            AccessConnection(DB)._get_access_app()

        assert app.Visible is True
        assert app.events.index(f"Open={DB}") < app.events.index("Visible=True")

    def test_instance_we_attach_to_becomes_visible(self):
        """Moniker binding launches Access hidden when nothing had the file open."""
        app = _FakeApp(current_db_name=DB)

        with patch.object(conn_mod, "win32com") as mock_win32com:
            mock_win32com.client.GetObject.return_value = app

            AccessConnection(DB)._get_access_app()

        assert app.Visible is True
