"""Tests for configuration management.

Note: load_config schema and defaults are exercised in
``tests/test_config_env_loading.py`` against the current env-var
schema (ACCESS_VCS_DATABASE, ACCESS_VCS_DISABLE_WRITES, etc.).
"""

import os

import pytest

import msaccess_vcs_mcp.config as config_module
from msaccess_vcs_mcp.config import _strip_quotes, get_default_addin_path


class TestInstalledAddInPath:
    """Resolving the install path from the settings the add-in's installer writes.

    ``HKCU\\Software\\VB and VBA Program Settings\\MSAccessVCS\\Install`` is the only
    place to read this from: the add-in's own ``GetInstalledAddInFileName`` joins
    exactly these two values. Reconstructing the path from ``%AppData%`` alone gets a
    custom install folder wrong, and always guessing ``.accda`` gets a compiled
    install wrong.
    """

    @pytest.fixture(autouse=True)
    def _clear_cache(self):
        config_module.reset_addin_path_cache()
        yield
        config_module.reset_addin_path_cache()

    def _settings(self, monkeypatch, folder=None, compiled=None):
        values = {"Install Folder": folder, "Compile accde": compiled}
        monkeypatch.setattr(
            config_module, "_read_install_setting", lambda name: values.get(name)
        )

    def test_default_folder_when_setting_absent(self, monkeypatch):
        # The installer deletes "Install Folder" when the folder is the default, so
        # its absence means %AppData%\MSAccessVCS -- not "not installed".
        self._settings(monkeypatch)
        monkeypatch.setenv("APPDATA", r"C:\Users\Example\AppData\Roaming")

        assert get_default_addin_path() == os.path.join(
            r"C:\Users\Example\AppData\Roaming", "MSAccessVCS", "Version Control.accda"
        )

    def test_custom_install_folder(self, monkeypatch):
        self._settings(monkeypatch, folder=r"D:\Tools\VCS")

        assert get_default_addin_path() == os.path.join(
            r"D:\Tools\VCS", "Version Control.accda"
        )

    def test_compiled_install_resolves_to_accde(self, monkeypatch):
        # SaveSetting persists a VBA integer, so True arrives as -1.
        self._settings(monkeypatch, folder=r"D:\Tools\VCS", compiled="-1")

        assert get_default_addin_path().endswith("Version Control.accde")

    def test_compile_flag_off_resolves_to_accda(self, monkeypatch):
        self._settings(monkeypatch, folder=r"D:\Tools\VCS", compiled="0")

        assert get_default_addin_path().endswith("Version Control.accda")

    def test_result_is_cached(self, monkeypatch):
        self._settings(monkeypatch, folder=r"D:\Tools\VCS")
        first = get_default_addin_path()

        self._settings(monkeypatch, folder=r"E:\Elsewhere")
        assert get_default_addin_path() == first

        config_module.reset_addin_path_cache()
        assert get_default_addin_path() != first


class TestStripQuotes:
    """Tests for the _strip_quotes helper used on path-valued env vars."""

    def test_unquoted_passthrough(self):
        assert _strip_quotes(r"C:\Projects\db.accdb") == r"C:\Projects\db.accdb"

    def test_double_quoted(self):
        assert _strip_quotes('"C:\\\\Projects\\\\db.accdb"') == "C:\\\\Projects\\\\db.accdb"

    def test_single_quoted(self):
        assert _strip_quotes("'C:\\\\Projects\\\\db.accdb'") == "C:\\\\Projects\\\\db.accdb"

    def test_path_with_spaces(self):
        assert (
            _strip_quotes('"C:\\\\Projects\\\\My Database.accdb"')
            == "C:\\\\Projects\\\\My Database.accdb"
        )

    def test_empty_string(self):
        assert _strip_quotes("") == ""

    def test_mismatched_quotes_untouched(self):
        assert _strip_quotes("\"C:\\\\Projects\\\\db.accdb'") == "\"C:\\\\Projects\\\\db.accdb'"

    def test_single_char_untouched(self):
        assert _strip_quotes('"') == '"'

    def test_empty_quoted_string(self):
        assert _strip_quotes('""') == ""
        assert _strip_quotes("''") == ""


def test_validate_access_installation():
    """Test Access COM validation."""
    from msaccess_vcs_mcp.config import validate_access_installation
    
    # This test will only pass if Access is installed
    # On CI/CD without Access, this would be skipped
    try:
        validate_access_installation()
    except (ImportError, RuntimeError) as e:
        pytest.skip(f"Access not available: {e}")
