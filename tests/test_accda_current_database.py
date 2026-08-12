"""Tests for opening .accda files as the Access current database.

``GetObject(path)`` resolves a file moniker through the COM registration for
the file's extension.  That registration opens .accdb and .mdb as the current
database, but not .accda -- Access treats it as an add-in and the bind fails.
``AccessConnection`` therefore falls back to an explicit
``OpenCurrentDatabase``, without which the instance has no current database
and every add-in call fails with "Cannot find Access instance".
"""

from unittest.mock import MagicMock, patch

import pytest

from msaccess_vcs_mcp.access_com import connection as conn_mod
from msaccess_vcs_mcp.access_com.connection import AccessConnection

ACCDA = r"C:\Repos\msaccess-vcs-addin\Version Control.accda"
ACCDB = r"C:\Repos\other\Some Database.accdb"


def _app_with_current_db(path):
    """Build a mock Access Application whose CurrentDb() reports ``path``."""
    app = MagicMock()
    if path is None:
        app.CurrentDb.return_value = None
    else:
        current = MagicMock()
        current.Name = path
        app.CurrentDb.return_value = current
    return app


@pytest.fixture(autouse=True)
def _com_available():
    with patch.object(conn_mod, "COM_AVAILABLE", True):
        yield


class TestAccdaOpensAsCurrentDatabase:
    def test_getobject_failure_falls_back_to_open_current_database(self):
        """The .accda case: moniker bind fails, so open the file explicitly."""
        app = _app_with_current_db(None)

        with (
            patch.object(conn_mod, "win32com") as mock_win32com,
            patch.object(conn_mod, "gencache") as mock_gencache,
        ):
            mock_win32com.client.GetObject.side_effect = Exception(
                "invalid reference to the Parent property"
            )
            mock_gencache.EnsureDispatch.return_value = app

            c = AccessConnection(ACCDA)
            result = c._get_access_app()

        assert result is app
        app.OpenCurrentDatabase.assert_called_once_with(ACCDA)
        assert c._db_opened_as_current is True
        assert c._db_opened_via_getobject is False

    def test_getobject_success_does_not_reopen(self):
        """The .accdb case: the moniker bind already made it current."""
        app = _app_with_current_db(ACCDB)

        with patch.object(conn_mod, "win32com") as mock_win32com:
            mock_win32com.client.GetObject.return_value = app

            c = AccessConnection(ACCDB)
            result = c._get_access_app()

        assert result is app
        app.OpenCurrentDatabase.assert_not_called()
        assert c._db_opened_via_getobject is True
        assert c._db_opened_as_current is False

    def test_instance_already_holding_target_is_not_reopened(self):
        """Reusing an instance that already has our database open."""
        app = _app_with_current_db(ACCDA)

        c = AccessConnection(ACCDA)
        c._owns_app = False
        c._open_as_current_database(app)

        app.OpenCurrentDatabase.assert_not_called()
        assert c._db_opened_as_current is True

    def test_never_displaces_a_user_database(self):
        """A database we don't own must not be closed out from under the user."""
        app = _app_with_current_db(ACCDB)

        c = AccessConnection(ACCDA)
        c._owns_app = False
        c._open_as_current_database(app)

        app.OpenCurrentDatabase.assert_not_called()
        assert c._db_opened_as_current is False

    def test_current_db_uses_currentdb_not_a_second_dao_handle(self):
        """After an explicit open, reuse CurrentDb rather than DBEngine."""
        app = _app_with_current_db(ACCDA)

        c = AccessConnection(ACCDA)
        c._app = app
        c._db_opened_as_current = True

        db = c._get_current_db()

        assert db is app.CurrentDb.return_value
        app.DBEngine.OpenDatabase.assert_not_called()
        assert c._owns_db is False

    def test_close_resets_the_flag(self):
        app = _app_with_current_db(ACCDA)

        c = AccessConnection(ACCDA)
        c._app = app
        c._owns_app = True
        c._db_opened_as_current = True

        c.close()

        assert c._db_opened_as_current is False
