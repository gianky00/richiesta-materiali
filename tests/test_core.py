import os
import pytest
import json
import hashlib
from datetime import date
from cryptography.fernet import Fernet
from unittest.mock import MagicMock

# Import modules to test
from src.core import version
from src.core import license_validator
from src.core import license_updater
from src.core import app_updater

# --- Version Tests ---
def test_version_format():
    assert isinstance(version.__version__, str)
    assert len(version.__version__.split('.')) == 3
    assert "netlify.app" in version.UPDATE_URL

# --- License Validator Tests ---
def test_license_validator_hashing(tmp_path):
    test_file = tmp_path / "test.txt"
    test_file.write_text("hello world")

    # sha256 of "hello world"
    expected = hashlib.sha256(b"hello world").hexdigest()
    assert license_validator._calculate_sha256(str(test_file)) == expected

def test_hardware_id_fallback():
    # Since we are on Linux/Docker, it likely uses a fallback or fails gracefully
    hw_id = license_validator.get_hardware_id()
    assert isinstance(hw_id, str)
    assert len(hw_id) > 0

# --- License Updater Tests ---
def test_get_github_token():
    token = license_updater.get_github_token()
    assert isinstance(token, str)
    assert token.startswith("ghp_")

def test_get_license_dir():
    d = license_updater.get_license_dir()
    assert "Licenza" in d

def test_check_grace_period_no_token(mocker):
    # Mock _get_validity_token_path to point to non-existent file
    mocker.patch("src.core.license_updater._get_validity_token_path", return_value="nonexistent.token")

    with pytest.raises(Exception, match="Nessuna validazione online precedente"):
        license_updater.check_grace_period()

# --- App Updater Tests ---
def test_app_updater_check(mocker):
    # Mock requests.get
    mock_resp = MagicMock()
    mock_resp.status_code = 200
    mock_resp.json.return_value = {
        "version": "99.99.99",
        "url": "http://test.com"
    }
    mocker.patch("requests.get", return_value=mock_resp)

    # Mock messagebox
    mocker.patch("tkinter.messagebox.askyesno", return_value=True)
    mock_open = mocker.patch("webbrowser.open")

    app_updater.check_for_updates(silent=False)

    mock_open.assert_called_once_with("http://test.com")
