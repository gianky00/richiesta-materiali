"""
RDA Bot - Automazione per l'elaborazione delle Richieste di Acquisto
Questo script deve essere eseguito periodicamente (es. Task Scheduler)

Funzionalità:
- Scansiona Outlook per email con allegati PDF RDA
- Estrae dati dai PDF
- Aggiorna il file Excel di registro
- Sincronizza i dati con il database SQLite
- Invia email di riepilogo per RDA scadute
"""

import pythoncom
import logging
import sys
import os

# Aggiungi la directory ROOT al path per permettere import relativi come 'src.utils...'
# Risaliamo da src/main_bot.py a ROOT
SCRIPT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from src.utils.config import PDF_SAVE_PATH, ensure_directories
from src.data.excel_manager import ExcelManager
from src.services.email_scanner import EmailScanner
from src.services.pdf_parser import extract_rda_data, save_pdf_to_archive
from src.data.database import init_db, replace_all_data
from src.utils.utils import logger

# Moduli Licenza
from src.core import license_updater
from src.core import license_validator


def process_pdf_callback(excel_mgr):
    """
    Crea una callback per processare i PDF trovati nelle email.
    
    Args:
        excel_mgr: Istanza di ExcelManager
    
    Returns:
        Funzione callback
    """
    def callback(temp_pdf_path):
        # Estrai dati dal PDF
        rda_data = extract_rda_data(temp_pdf_path)
        if not rda_data:
            return
        
        # Verifica se RDA già esistente
        if excel_mgr.check_if_exists(rda_data['rda_number_raw']):
            logger.warning(f"RDA {rda_data['rda_number_raw']} già esistente in Excel. Saltata.")
            return
        
        # Salva PDF nell'archivio
        final_path = save_pdf_to_archive(
            temp_pdf_path, 
            rda_data['rda_number_raw'], 
            rda_data['rda_date_str']
        )
        if not final_path:
            return
        
        # Aggiungi percorso PDF ai dati
        rda_data['pdf_final_path'] = final_path
        
        # Inserisci in Excel
        excel_mgr.append_data(rda_data)
        logger.info(f"RDA {rda_data['rda_number_raw']} elaborata con successo")
    
    return callback


def main():
    """
    Funzione principale del bot.
    Esegue l'intero ciclo di elaborazione.
    """
    logger.info("=" * 50)
    logger.info("Avvio RDA Bot...")
    logger.info("=" * 50)

    # -------------------------------------------------------------------------
    # CHECK LICENZA
    # -------------------------------------------------------------------------
    try:
        logger.info("Aggiornamento licenza...")
        license_updater.run_update()

        is_valid, msg = license_validator.verify_license()
        if not is_valid:
            logger.critical(f"LICENZA NON VALIDA: {msg}")
            print(f"FATAL ERROR: {msg}", file=sys.stderr)
            return 1

        logger.info(f"Licenza OK: {msg}")
    except Exception as e:
        logger.error(f"Errore controllo licenza: {e}")
        return 1

    # 1. Inizializza database SQLite
    try:
        init_db()
    except Exception as e:
        logger.error(f"Errore inizializzazione database: {e}")
        return 1
    
    # 2. Verifica/crea directory necessarie
    ensure_directories()
    
    if not os.path.exists(PDF_SAVE_PATH):
        logger.error(f"Directory PDF non accessibile: {PDF_SAVE_PATH}")
        return 1
    
    # 3. Inizializza COM per Windows
    pythoncom.CoInitialize()
    
    excel_mgr = None
    scanner = None
    exit_code = 0
    
    try:
        # 4. Apri Excel
        excel_mgr = ExcelManager()
        if not excel_mgr.open():
            logger.error("Impossibile aprire il file Excel. Verifica che il percorso sia corretto.")
            return 1
        
        logger.info("File Excel aperto correttamente")
        
        # 5. Scansiona email
        scanner = EmailScanner()
        if scanner.connect():
            logger.info("Connesso a Outlook")
            
            messages = scanner.get_messages()
            if messages:
                # Crea callback con riferimento a excel_mgr
                pdf_callback = process_pdf_callback(excel_mgr)
                scanner.process_emails(messages, pdf_callback)
            else:
                logger.info("Nessuna email da processare")
        else:
            logger.warning("Impossibile connettersi a Outlook")
        
        # 6. Aggiorna livelli di alert
        logger.info("Aggiornamento livelli di alert...")
        overdue_items = excel_mgr.update_alerts_and_get_overdue()
        
        # 7. Invia email di riepilogo se ci sono RDA scadute
        if scanner and scanner.outlook and overdue_items:
            logger.info(f"Trovate {len(overdue_items)} RDA scadute. Invio email...")
            scanner.send_summary_email(overdue_items)
        
        # 8. Pulizia Excel
        logger.info("Pulizia file Excel...")
        excel_mgr.delete_empty_rows()
        excel_mgr.fit_columns()
        
        # 9. Sincronizza Excel -> SQLite
        logger.info("Sincronizzazione Excel -> Database SQLite...")
        all_data = excel_mgr.get_all_data_for_sync()
        if all_data:
            replace_all_data(all_data)
            logger.info(f"Database sincronizzato: {len(all_data)} righe")
        else:
            logger.warning("Nessun dato trovato in Excel per la sincronizzazione")
        
        # 10. Salva e chiudi Excel
        excel_mgr.close(save=True)
        logger.info("File Excel salvato e chiuso")
        
    except Exception as e:
        logger.error(f"Errore generale: {e}")
        exit_code = 1
        
        # Tenta chiusura Excel senza salvare in caso di errore
        if excel_mgr:
            try:
                excel_mgr.close(save=False)
            except:
                pass
    
    finally:
        pythoncom.CoUninitialize()
        logger.info("=" * 50)
        logger.info("Elaborazione completata")
        logger.info("=" * 50)
    
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
