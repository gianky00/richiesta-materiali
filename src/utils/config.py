"""
Configurazione centralizzata per il sistema RDA
Questo file contiene tutte le impostazioni per automazione e GUI
"""

import os
import sys
from src.core import config_manager

# --- DETERMINAZIONE PERCORSI ---
def get_base_path():
    """Determina il percorso base (supporta sia sviluppo che EXE compilato)"""
    return config_manager.get_base_path()

SCRIPT_DIR = get_base_path()

# Carica configurazione dinamica
CONFIG = config_manager.current_config

# --- PATHS ---
# Percorsi definiti in config.json o fallback ai default
EXCEL_DB_PATH = CONFIG.get("excel_path")
# Fallback logic if config path doesn't exist?
# No, let's trust the config manager logic (which defaults to network path).
# User requirement is to allow changing it.

# Derived paths based on config or defaults
# Note: config_manager defaults define full paths.
DATABASE_DIR = CONFIG.get("database_dir", os.path.dirname(EXCEL_DB_PATH))
PDF_SAVE_PATH = CONFIG.get("pdf_folder")
SQLITE_DB_PATH = os.path.join(DATABASE_DIR, "database_RDA.db")

# --- EXCEL SETTINGS ---
SHEET_PASSWORD = "coemi"
TABLE_NAME = "Tabella1"
RDA_REFERENCE_COLUMN_LETTER = "A"
RDA_REFERENCE_COLUMN_NUMBER = 1

# --- OUTLOOK SETTINGS ---
OUTLOOK_FOLDER_ID = 6  # Inbox
TARGET_FOLDER_NAME = "MAGO"
SENDER_EMAIL = "magonet@coemi.it"
ATTACHMENT_NAME = "RDAPerFornitore.pdf"
DAYS_TO_CHECK = 60  # Controlla solo email degli ultimi X giorni

# --- EMAIL ALERT SETTINGS ---
EMAIL_SENDER = "isabsud@coemi.it"
EMAIL_RECIPIENT = "isabsud@coemi.it;concetto.siringo@coemi.it"
EMAIL_SUBJECT = "RIEPILOGO RDA SCADUTE"

# --- LOGGING ---
LOG_FORMAT = '%(asctime)s - %(levelname)s - %(message)s'
LOG_LEVEL = "INFO"


def ensure_directories():
    """Crea le directory necessarie se non esistono"""
    # Create local logs dir
    log_dir = os.path.join(SCRIPT_DIR, "Logs")
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    # We generally don't want to create network paths if they are missing/unreachable,
    # but we can try.
    if os.path.exists(os.path.dirname(PDF_SAVE_PATH)):
        if not os.path.exists(PDF_SAVE_PATH):
            try:
                os.makedirs(PDF_SAVE_PATH)
            except OSError as e:
                print(f"Impossibile creare directory {PDF_SAVE_PATH}: {e}")
