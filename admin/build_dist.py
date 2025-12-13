"""
RDA Automation System - Build & Deploy Script
Utilizza Nuitka per la compilazione e Inno Setup per l'installer.
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
from packaging import version as pkg_version

# Configurazione
APP_NAME_GUI = "RDA_Viewer"
APP_NAME_BOT = "RDA_Bot"
MAIN_SCRIPT_GUI = "main_gui.py"
MAIN_SCRIPT_BOT = "main_bot.py"

NETLIFY_SITE_NAME = "intelleo-rda-viewer" # Nome del sito Netlify per Site ID lookup

# Percorsi
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DIST_DIR = os.path.join(ROOT_DIR, "dist")
SRC_DIR = os.path.join(ROOT_DIR, "src")
ISCC_EXE = r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" # Path default Inno Setup

# Logging
logging.basicConfig(level=logging.INFO, format='[BUILD] %(message)s')
logger = logging.getLogger()

# Import versione
sys.path.append(ROOT_DIR)
import version
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

def build_nuitka(script_name, output_name, console=False):
    """Compila uno script con Nuitka."""
    log_and_print(f"--- Building {output_name} with Nuitka ---")

    cmd = [
        sys.executable, "-m", "nuitka",
        "--standalone",
        "--lto=no", # Velocizza la build, metti 'yes' per produzione finale ottimizzata
        "--follow-imports",
        "--include-package=src",
        f"--output-dir={DIST_DIR}",
        "--output-filename=" + output_name,
    ]

    if not console:
        cmd.append("--windows-disable-console")
        cmd.append("--enable-plugin=tk-inter")

    cmd.append(os.path.join(ROOT_DIR, script_name))

    run_command(cmd, cwd=ROOT_DIR)

    # Rinomina la cartella di output generata da Nuitka
    nuitka_dist_dir = os.path.join(DIST_DIR, f"{script_name.replace('.py', '')}.dist")
    final_target_dir = os.path.join(DIST_DIR, output_name)

    if os.path.exists(final_target_dir):
        shutil.rmtree(final_target_dir)

    if os.path.exists(nuitka_dist_dir):
        shutil.move(nuitka_dist_dir, final_target_dir)
    else:
        log_and_print(f"Errore: Nuitka output dir not found: {nuitka_dist_dir}", "ERROR")
        sys.exit(1)

    return final_target_dir

def copy_assets(target_dir):
    """Copia le risorse necessarie nella cartella di build."""
    log_and_print(f"Copying assets to {target_dir}...")

    # Crea cartella Licenza vuota
    os.makedirs(os.path.join(target_dir, "Licenza"), exist_ok=True)

    # Copia src (anche se Nuitka include i py, a volte servono risorse statiche)
    # Se src contiene solo codice, Nuitka lo gestisce.
    # Ma il codice originale usava --add-data=src;src
    src_dest = os.path.join(target_dir, "src")
    if os.path.exists(SRC_DIR):
        shutil.copytree(SRC_DIR, src_dest, dirs_exist_ok=True)

def merge_builds(gui_dir, bot_dir):
    """
    Unisce le due build in una struttura unica per l'installer.
    Poiché Nuitka standalone crea un folder con tutte le dipendenze,
    unire due build standalone è complesso (conflitti di DLL).

    Strategia: Manteniamo due cartelle separate nell'installer.
    Oppure: In futuro, si può usare una build unica con due entry point, ma è complesso.

    Per ora: Ritorniamo i path, l'installer gestirà le source.
    """
    pass

def create_installer(gui_dir, bot_dir):
    """Compila lo script Inno Setup."""
    log_and_print("--- Compiling Installer ---")

    iss_path = os.path.join(ROOT_DIR, "admin", "setup_script.iss")
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

    cmd = [
        iscc,
        f"/DMyAppVersion={APP_VERSION}",
        f"/DGuiDir={gui_dir}",
        f"/DBotDir={bot_dir}",
        f"/DOutputDir={DIST_DIR}",
        iss_path
    ]

    run_command(cmd)

    # Trova l'output
    for f in os.listdir(DIST_DIR):
        if f.endswith(".exe") and "Setup" in f:
            return os.path.join(DIST_DIR, f)

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
    <style>
        body {{ font-family: 'Segoe UI', sans-serif; background: #f4f4f9; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }}
        .container {{ background: white; padding: 40px; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.1); text-align: center; max-width: 500px; }}
        .btn {{ display: inline-block; background-color: #007bff; color: white; padding: 15px 30px; text-decoration: none; font-size: 18px; border-radius: 6px; margin-top: 20px; }}
        .btn:hover {{ background-color: #0056b3; }}
        .info {{ margin-top: 20px; color: #888; font-size: 14px; }}
    </style>
</head>
<body>
    <div class="container">
        <h1>{APP_NAME_GUI}</h1>
        <p>Sistema di Gestione Richieste di Acquisto e Automazione.</p>
        <a href="{setup_filename}" class="btn">Scarica v{version_str}</a>
        <div class="info">Ultimo aggiornamento: {time.strftime('%d/%m/%Y')}</div>
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
    print("="*60)
    print(f"   {APP_NAME_GUI} Build System (v{APP_VERSION})")
    print("="*60)

    clean_dist()

    # 1. Build GUI
    gui_dist = build_nuitka(MAIN_SCRIPT_GUI, APP_NAME_GUI, console=False)
    copy_assets(gui_dist)

    # 2. Build Bot
    bot_dist = build_nuitka(MAIN_SCRIPT_BOT, APP_NAME_BOT, console=True)
    copy_assets(bot_dist)

    # 3. Installer
    installer_path = create_installer(gui_dist, bot_dist)

    # 4. Deploy
    if installer_path:
        deploy_to_netlify(installer_path)

    print("\nBUILD COMPLETE.")

if __name__ == "__main__":
    build()
