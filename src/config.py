import os

# --- PATHS ---
# Network paths
NETWORK_BASE_PATH = r"\\192.168.11.251\Condivisa\RICHIESTE MATERIALI"
PDF_SAVE_PATH = os.path.join(NETWORK_BASE_PATH, "RDA_PDF")
DATABASE_DIR = os.path.join(NETWORK_BASE_PATH, "DATABASE")

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
DAYS_TO_CHECK = 60

# --- EMAIL ALERT SETTINGS ---
EMAIL_SENDER = "isabsud@coemi.it"
EMAIL_RECIPIENT = "isabsud@coemi.it;concetto.siringo@coemi.it"
EMAIL_SUBJECT = "RIEPILOGO RDA SCADUTE"
