"""
RDA Automation System - Build & Deploy Script
Utilizza PyInstaller per la compilazione e Inno Setup per l'installer.
Supporta aggiornamenti automatici tramite Netlify.
"""

import os
import sys
import shutil
import subprocess
import json
import logging
import zipfile
import time
import requests
import argparse
from packaging import version as pkg_version

# Configurazione
APP_NAME_GUI = "RDA_Viewer"
APP_NAME_BOT = "RDA_Bot"
MAIN_SCRIPT_GUI = "src/main_gui.py"
MAIN_SCRIPT_BOT = "src/main_bot.py"

NETLIFY_SITE_NAME = "intelleo-rda-viewer" # Nome del sito Netlify per Site ID lookup

# Percorsi
# Risaliamo alla root del progetto (admin/Crea Setup -> admin -> root)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(os.path.dirname(SCRIPT_DIR))
DIST_DIR = os.path.join(ROOT_DIR, "dist")
SRC_DIR = os.path.join(ROOT_DIR, "src")
ASSETS_DIR = os.path.join(ROOT_DIR, "assets")
ISCC_EXE = r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" # Path default Inno Setup

# Icons
ICON_APP = os.path.join(ASSETS_DIR, "app.ico")
ICON_BOT = os.path.join(ASSETS_DIR, "bot.ico")
ICON_SETUP = os.path.join(ASSETS_DIR, "setup.ico")

# Logging
logging.basicConfig(level=logging.INFO, format='[BUILD] %(message)s')
logger = logging.getLogger()

# Import versione
sys.path.append(ROOT_DIR)
from src.core import version
APP_VERSION = version.__version__

def log_and_print(msg, level="INFO"):
    if level == "ERROR":
        logger.error(msg)
    elif level == "WARNING":
        logger.warning(msg)
    else:
        logger.info(msg)

def run_command(cmd, cwd=None, env=None):
    """Esegue un comando shell e gestisce errori."""
    try:
        if cwd:
            log_and_print(f"Executing in {cwd}: {' '.join(cmd)}")
        else:
            log_and_print(f"Executing: {' '.join(cmd)}")

        subprocess.check_call(cmd, cwd=cwd, env=env)
    except subprocess.CalledProcessError as e:
        log_and_print(f"Command failed: {e}", "ERROR")
        sys.exit(1)

def clean_dist():
    """Pulisce la cartella dist."""
    if os.path.exists(DIST_DIR):
        log_and_print("Cleaning dist directory...")
        shutil.rmtree(DIST_DIR)
    os.makedirs(DIST_DIR)

def check_pyinstaller():
    """Verifica che PyInstaller sia installato."""
    try:
        import PyInstaller
        log_and_print(f"PyInstaller found (version {PyInstaller.__version__})")
        return True
    except ImportError:
        log_and_print("PyInstaller not found. Installing...", "WARNING")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
            return True
        except subprocess.CalledProcessError:
            log_and_print("Failed to install PyInstaller.", "ERROR")
            return False

def build_pyinstaller(script_name, output_name, console=False, hidden_imports=None, icon_path=None):
    """Compila uno script con PyInstaller."""
    log_and_print(f"--- Building {output_name} with PyInstaller ---")

    # Temp output dir for this specific build (to match structure expected by Inno Setup)
    # output_name is like "RDA_Viewer"
    # We want dist/temp_gui/RDA_Viewer.exe

    # Identify if it is GUI or Bot based on name to name the temp dir correctly
    # build() calls with APP_NAME_GUI ("RDA_Viewer") and APP_NAME_BOT ("RDA_Bot")
    temp_dir_name = "temp_gui" if output_name == APP_NAME_GUI else "temp_bot"
    target_dir = os.path.join(DIST_DIR, temp_dir_name)

    if os.path.exists(target_dir):
        shutil.rmtree(target_dir)
    os.makedirs(target_dir)

    # Use absolute path for src to avoid PyInstaller confusion
    src_abs_path = os.path.join(ROOT_DIR, "src")

    options = [
        script_name,
        "--name=" + output_name,
        "--onefile",
        "--noconfirm",
        "--clean",
        f"--distpath={target_dir}", # Output executable here
        "--workpath=" + os.path.join(DIST_DIR, "build", output_name), # Temp build files
        "--specpath=" + os.path.join(DIST_DIR, "build", output_name), # Spec file
        f"--add-data={src_abs_path};src", # Bundle source code (Absolute path)
    ]

    if console:
        options.append("--console")
    else:
        options.append("--windowed")

    if hidden_imports:
        for imp in hidden_imports:
            options.append(f"--hidden-import={imp}")

    # Add icon if provided and exists
    if icon_path and os.path.exists(icon_path):
        log_and_print(f"Using icon: {icon_path}")
        options.append(f"--icon={icon_path}")
    elif icon_path:
        log_and_print(f"Icon not found at {icon_path}. Using default.", "WARNING")
    else:
        log_and_print("No icon specified.", "INFO")

    cmd = [sys.executable, "-m", "PyInstaller"] + options

    run_command(cmd, cwd=ROOT_DIR)

    # Verify output
    exe_path = os.path.join(target_dir, f"{output_name}.exe")
    if not os.path.exists(exe_path):
        log_and_print(f"Error: PyInstaller output not found: {exe_path}", "ERROR")
        sys.exit(1)

    return target_dir

def copy_assets(target_dir):
    """Copia le risorse necessarie nella cartella di build."""
    log_and_print(f"Copying assets to {target_dir}...")

    # Crea cartella Licenza vuota
    os.makedirs(os.path.join(target_dir, "Licenza"), exist_ok=True)

    # Note: src is bundled inside the EXE by PyInstaller (--add-data),
    # so we do not copy it here to avoid duplication in the installer.

def create_installer(gui_dir, bot_dir):
    """Compila lo script Inno Setup."""
    log_and_print("--- Compiling Installer ---")

    iss_path = os.path.join(ROOT_DIR, "admin", "Crea Setup", "setup_script.iss")
    if not os.path.exists(iss_path):
        log_and_print(f"Setup script not found: {iss_path}", "ERROR")
        return None

    # Verifica ISCC
    iscc = ISCC_EXE
    if not os.path.exists(iscc):
        # Cerca nel path
        iscc = shutil.which("ISCC")
        if not iscc:
            log_and_print("ISCC (Inno Setup) not found. Skipping installer generation.", "WARNING")
            return None

    # Set OutputDir to "admin/Crea Setup/Setup" as requested
    setup_output_dir = os.path.join(ROOT_DIR, "admin", "Crea Setup", "Setup")
    if not os.path.exists(setup_output_dir):
        os.makedirs(setup_output_dir)

    cmd = [
        iscc,
        f"/DMyAppVersion={APP_VERSION}",
        f"/DGuiDir={gui_dir}",
        f"/DBotDir={bot_dir}",
        f"/DOutputDir={setup_output_dir}",
        iss_path
    ]

    if os.path.exists(ICON_SETUP):
        log_and_print(f"Using setup icon: {ICON_SETUP}")
        cmd.append(f"/DSetupIcon={ICON_SETUP}")
    else:
        log_and_print(f"Setup icon not found at {ICON_SETUP}. Using default.", "WARNING")

    run_command(cmd)

    # Trova l'output
    for f in os.listdir(setup_output_dir):
        if f.endswith(".exe") and "Setup" in f:
            return os.path.join(setup_output_dir, f)

    return None

def get_netlify_token():
    """Returns the obfuscated Netlify API token."""
    # Obfuscated token parts
    p1 = "nfp_VJbSMoKXxms3"
    p2 = "Xa8gdQkKKedPC6"
    p3 = "EnHQZL9687"
    return p1 + p2 + p3

def get_netlify_site_id(site_name, token):
    """Retrieves the Site ID for a given site name using the Netlify API."""
    try:
        log_and_print(f"Fetching Site ID for '{site_name}'...")
        url = "https://api.netlify.com/api/v1/sites"
        headers = {"Authorization": f"Bearer {token}"}

        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            sites = response.json()
            for site in sites:
                if site.get("name") == site_name:
                    return site.get("site_id")
            log_and_print(f"Site '{site_name}' not found in account.", "ERROR")
        else:
            log_and_print(f"Error fetching sites: {response.status_code} - {response.text}", "ERROR")
    except Exception as e:
        log_and_print(f"Exception getting Site ID: {e}", "ERROR")
    return None

def generate_index_html(deploy_dir, setup_filename, version_str):
    """Generates a professional index.html download page."""
    html_content = f"""<!DOCTYPE html>
<html lang="it">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Download {APP_NAME_GUI}</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.0/font/bootstrap-icons.css">
    <style>
        body {{
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
        }}
        .card {{
            border: none;
            border-radius: 15px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
            max-width: 500px;
            width: 90%;
        }}
        .card-header {{
            background-color: white;
            border-bottom: none;
            padding-top: 30px;
            border-radius: 15px 15px 0 0 !important;
        }}
        .app-icon {{
            font-size: 4rem;
            color: #0d6efd;
        }}
        .btn-download {{
            padding: 15px 30px;
            font-size: 1.2rem;
            font-weight: 600;
            border-radius: 50px;
            box-shadow: 0 4px 6px rgba(13, 110, 253, 0.3);
            transition: all 0.3s ease;
        }}
        .btn-download:hover {{
            transform: translateY(-2px);
            box-shadow: 0 6px 12px rgba(13, 110, 253, 0.4);
        }}
        .features-list {{
            text-align: left;
            margin: 20px 0;
            color: #6c757d;
        }}
        .features-list li {{
            margin-bottom: 8px;
        }}
    </style>
</head>
<body>
    <div class="card text-center p-4">
        <div class="card-header">
            <i class="bi bi-file-earmark-pdf-fill app-icon"></i>
            <h2 class="mt-3 fw-bold text-primary">{APP_NAME_GUI}</h2>
            <p class="text-muted">Sistema Avanzato di Gestione Richieste Acquisto</p>
        </div>
        <div class="card-body">
            <ul class="list-unstyled features-list mx-auto" style="max-width: 350px;">
                <li><i class="bi bi-check-circle-fill text-success me-2"></i>Gestione centralizzata RDA</li>
                <li><i class="bi bi-check-circle-fill text-success me-2"></i>Automazione Outlook & PDF</li>
                <li><i class="bi bi-check-circle-fill text-success me-2"></i>Analisi e Reporting Avanzato</li>
            </ul>

            <a href="{setup_filename}" class="btn btn-primary btn-download w-100 my-3">
                <i class="bi bi-windows me-2"></i> Scarica per Windows
            </a>

            <div class="mt-4 pt-3 border-top">
                <div class="row text-muted small">
                    <div class="col-6 text-start">
                        Versione: <span class="fw-bold text-dark">v{version_str}</span>
                    </div>
                    <div class="col-6 text-end">
                        Data: {time.strftime('%d/%m/%Y')}
                    </div>
                </div>
            </div>
        </div>
    </div>
</body>
</html>"""

    with open(os.path.join(deploy_dir, "index.html"), "w", encoding="utf-8") as f:
        f.write(html_content)

def deploy_to_netlify(installer_path):
    """Prepare files and upload to Netlify."""
    if not installer_path:
        log_and_print("No installer to deploy. Skipping.", "WARNING")
        return

    log_and_print("--- Preparing Deployment ---")
    deploy_dir = os.path.join(DIST_DIR, "deploy")
    if os.path.exists(deploy_dir):
        shutil.rmtree(deploy_dir)
    os.makedirs(deploy_dir)

    # 1. Copy Installer
    setup_filename = os.path.basename(installer_path)
    shutil.copy(installer_path, os.path.join(deploy_dir, setup_filename))

    # 2. Version JSON
    base_url = version.UPDATE_URL.rsplit('/', 1)[0]
    # Se UPDATE_URL è https://site.app/version.json, base è https://site.app
    # Ma se il dominio è diverso, attenzione.
    # Qui assumiamo che version.UPDATE_URL punti al sito Netlify che stiamo aggiornando.

    download_url = f"https://{NETLIFY_SITE_NAME}.netlify.app/{setup_filename}"

    version_data = {
        "version": APP_VERSION,
        "url": download_url
    }

    with open(os.path.join(deploy_dir, "version.json"), "w") as f:
        json.dump(version_data, f, indent=4)

    # 3. Index HTML
    generate_index_html(deploy_dir, setup_filename, APP_VERSION)

    # 4. Upload
    log_and_print("--- Uploading to Netlify ---")
    token = get_netlify_token()
    site_id = get_netlify_site_id(NETLIFY_SITE_NAME, token)

    if not site_id:
        log_and_print("DEPLOY FAILED: Site ID not found.", "ERROR")
        return

    zip_path = os.path.join(DIST_DIR, "deploy.zip")
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(deploy_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, deploy_dir)
                zipf.write(file_path, arcname)

    try:
        with open(zip_path, 'rb') as f:
            data = f.read()

        url = f"https://api.netlify.com/api/v1/sites/{site_id}/deploys"
        headers = {
            "Content-Type": "application/zip",
            "Authorization": f"Bearer {token}"
        }

        response = requests.post(url, headers=headers, data=data, timeout=300)

        if response.status_code == 200:
            log_and_print(f"DEPLOY SUCCESSFUL! Live at: {response.json().get('url')}")
        else:
            log_and_print(f"Upload Failed: {response.status_code} - {response.text}", "ERROR")

    except Exception as e:
        log_and_print(f"Error uploading: {e}", "ERROR")


def build():
    parser = argparse.ArgumentParser(description="Build and Deploy RDA Viewer")
    parser.add_argument("--no-deploy", action="store_true", help="Skip deployment to Netlify")
    args = parser.parse_args()

    print("="*60)
    print(f"   {APP_NAME_GUI} Build System (v{APP_VERSION})")
    print("="*60)

    # Check dependencies
    if not check_pyinstaller():
        sys.exit(1)

    clean_dist()

    # 1. Build GUI
    gui_imports = ["sqlite3", "tkinter", "tkinter.ttk"]
    gui_dist = build_pyinstaller(MAIN_SCRIPT_GUI, APP_NAME_GUI, console=False, hidden_imports=gui_imports, icon_path=ICON_APP)
    copy_assets(gui_dist)

    # 2. Build Bot
    bot_imports = ["win32com.client", "pythoncom", "pdfplumber", "sqlite3"]
    bot_dist = build_pyinstaller(MAIN_SCRIPT_BOT, APP_NAME_BOT, console=True, hidden_imports=bot_imports, icon_path=ICON_BOT)
    copy_assets(bot_dist)

    # 3. Installer
    installer_path = create_installer(gui_dist, bot_dist)

    # 4. Deploy
    if installer_path and not args.no_deploy:
        deploy_to_netlify(installer_path)
    elif args.no_deploy:
        print("\n[INFO] Deployment skipped via --no-deploy flag.")

    print("\nBUILD COMPLETE.")

if __name__ == "__main__":
    build()
