"""
Script per la sincronizzazione manuale Excel -> SQLite
Utile per aggiornare il database senza eseguire l'intero bot

Uso:
    python run_sync.py
"""

import pythoncom
import sys
import os

# Aggiungi directory al path
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from src.excel_manager import ExcelManager
from src.database import init_db, replace_all_data
from src.utils import logger


def run_sync():
    """Esegue la sincronizzazione Excel -> Database SQLite."""
    logger.info("=" * 50)
    logger.info("Avvio sincronizzazione manuale Excel -> DB")
    logger.info("=" * 50)
    
    # 1. Inizializza Database
    try:
        init_db()
    except Exception as e:
        logger.error(f"Errore inizializzazione database: {e}")
        return 1
    
    # 2. Inizializza COM
    pythoncom.CoInitialize()
    
    excel_mgr = None
    exit_code = 0
    
    try:
        # 3. Apri Excel in sola lettura
        excel_mgr = ExcelManager()
        if not excel_mgr.open():
            logger.error("Impossibile aprire il file Excel.")
            return 1
        
        # 4. Leggi tutti i dati
        logger.info("Lettura dati dal file Excel...")
        all_data = excel_mgr.get_all_data_for_sync()
        
        # 5. Aggiorna database
        if all_data:
            logger.info(f"Trovate {len(all_data)} righe. Aggiornamento DB...")
            replace_all_data(all_data)
            logger.info("Database aggiornato con successo!")
        else:
            logger.warning("Nessun dato trovato nel file Excel.")
        
        # 6. Chiudi senza salvare (lettura only)
        excel_mgr.close(save=False)
        
    except Exception as e:
        logger.error(f"Errore durante la sincronizzazione: {e}")
        exit_code = 1
        if excel_mgr:
            excel_mgr.close(save=False)
    
    finally:
        pythoncom.CoUninitialize()
        logger.info("=" * 50)
        logger.info("Sincronizzazione completata")
        logger.info("=" * 50)
    
    return exit_code


if __name__ == "__main__":
    sys.exit(run_sync())
