import pythoncom
import logging
from src.excel_manager import ExcelManager
from src.database import init_db, replace_all_data
from src.utils import logger

def run_sync():
    logger.info("Avvio sincronizzazione manuale (Excel -> DB)...")

    # 1. Initialize Database
    init_db()

    # 2. Open Excel
    excel_mgr = ExcelManager()
    pythoncom.CoInitialize()

    if not excel_mgr.open():
        logger.error("Impossibile aprire il file Excel.")
        return

    try:
        # 3. Read Data
        logger.info("Lettura dati dal file Excel...")
        all_data = excel_mgr.get_all_data_for_sync()

        # 4. Update DB
        if all_data:
            logger.info(f"Trovate {len(all_data)} righe. Aggiornamento DB in corso...")
            replace_all_data(all_data)
            logger.info("Database aggiornato con successo.")
        else:
            logger.warning("Nessun dato trovato nel file Excel (o errore di lettura).")

        excel_mgr.close(save=False) # No need to save if we only read

    except Exception as e:
        logger.error(f"Errore durante la sincronizzazione: {e}")
        excel_mgr.close(save=False)
    finally:
        pythoncom.CoUninitialize()
        logger.info("--- Sincronizzazione Completata ---")

if __name__ == "__main__":
    run_sync()
