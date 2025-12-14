"""
Configurazione centralizzata per il sistema RDA
Questo file contiene tutte le impostazioni per automazione e GUI
"""

import os
import sys

# --- DETERMINAZIONE PERCORSI ---
def get_base_path():
    """Determina il percorso base (supporta sia sviluppo che EXE compilato)"""
    if getattr(sys, 'frozen', False):
        # Eseguibile PyInstaller
        return os.path.dirname(sys.executable)
    else:
        # Script Python normale
        # Risaliamo da src/utils/config.py fino alla root
        return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

SCRIPT_DIR = get_base_path()

# --- PATHS ---
# Prova prima percorso di rete, poi locale
NETWORK_BASE_PATH = r"\\192.168.11.251\Condivisa\RICHIESTE MATERIALI"
LOCAL_BASE_PATH = SCRIPT_DIR

def get_active_base_path():
    """Restituisce il percorso base attivo (rete se disponibile, altrimenti locale)"""
    if os.path.exists(NETWORK_BASE_PATH):
        return NETWORK_BASE_PATH
    return LOCAL_BASE_PATH

# Percorsi dinamici
BASE_PATH = get_active_base_path()
PDF_SAVE_PATH = os.path.join(BASE_PATH, "RDA_PDF")
DATABASE_DIR = os.path.join(BASE_PATH, "DATABASE")

EXCEL_DB_PATH = os.path.join(DATABASE_DIR, "database_RDA.xlsm")
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
    for path in [PDF_SAVE_PATH, DATABASE_DIR]:
        if not os.path.exists(path):
            try:
                os.makedirs(path)
            except OSError as e:
                print(f"Impossibile creare directory {path}: {e}")
