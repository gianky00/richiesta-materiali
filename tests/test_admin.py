import os
import sys
import pytest
import shutil
import json
from unittest.mock import MagicMock

# Import admin tool
sys.path.append(os.path.join(os.getcwd(), 'admin'))
import admin_license_gui

def test_generate_license(tmp_path, mocker):
    # Setup paths in admin_license_gui to use tmp_path

    # Actually, let's instantiate the class and mock the UI inputs
    root = MagicMock()
    app = admin_license_gui.LicenseAdminApp(root)

    # Mock inputs
    app.ent_disk = MagicMock()
    app.ent_disk.get.return_value = "TEST-HW-ID-123"

    app.ent_name = MagicMock()
    app.ent_name.get.return_value = "TestClient"

    app.ent_date = MagicMock()
    app.ent_date.get.return_value = "2099-12-31"

    # Mock os.path.dirname to force output to tmp_path
    mocker.patch("os.path.abspath", return_value=str(tmp_path / "admin_script.py"))

    # Mock messagebox
    mocker.patch("tkinter.messagebox.showinfo")
    mocker.patch("os.startfile", create=True) # Windows only, so we mock/create

    app.generate()

    # Verify output
    expected_dir = tmp_path / "TestClient" / "Licenza"
    assert expected_dir.exists()
    assert (expected_dir / "config.dat").exists()
    assert (expected_dir / "manifest.json").exists()

    # Verify content validity
    with open(expected_dir / "manifest.json") as f:
        manifest = json.load(f)
        assert "config.dat" in manifest
