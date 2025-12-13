import pythoncom
import logging
from src.config import *
from src.excel_manager import ExcelManager
from src.email_scanner import EmailScanner
from src.pdf_parser import extract_rda_data, save_pdf_to_archive
from src.database import init_db, replace_all_data
from src.utils import logger

def main():
    logger.info("Starting RDA Bot...")

    # 1. Initialize Database (ensure table exists)
    init_db()

    # 2. Check Paths
    if not os.path.exists(PDF_SAVE_PATH):
        try:
            os.makedirs(PDF_SAVE_PATH)
        except OSError:
            logger.error(f"Cannot create path {PDF_SAVE_PATH}")
            return

    # 3. Setup Excel Manager
    excel_mgr = ExcelManager()
    pythoncom.CoInitialize() # Needed for win32com in some contexts

    if not excel_mgr.open():
        logger.error("Could not open Excel database. Aborting.")
        return

    try:
        # Define callback for PDF processing
        def process_pdf(temp_pdf_path):
            rda_data = extract_rda_data(temp_pdf_path)
            if not rda_data: return

            if excel_mgr.check_if_exists(rda_data['rda_number_raw']):
                logger.warning(f"RDA {rda_data['rda_number_raw']} already exists in Excel. Skipping.")
                return

            final_path = save_pdf_to_archive(temp_pdf_path, rda_data['rda_number_raw'], rda_data['rda_date_str'])
            if not final_path: return

            rda_data['pdf_final_path'] = final_path
            excel_mgr.append_data(rda_data)

        # 4. Scan Emails and Process
        scanner = EmailScanner()
        if scanner.connect():
            messages = scanner.get_messages()
            scanner.process_emails(messages, process_pdf)

        # 5. Update Alerts & Send Email
        logger.info("Updating alerts in Excel...")
        overdue_items = excel_mgr.update_alerts_and_get_overdue()
        if scanner.outlook: # Re-use connection
            scanner.send_summary_email(overdue_items)

        # 6. Clean up Excel
        excel_mgr.delete_empty_rows()
        excel_mgr.fit_columns()

        # 7. SYNC: Read final state of Excel and update DB
        logger.info("Syncing Excel to SQLite Database...")
        # Since we modified the sheet, we can read from it directly before closing.
        # However, save first to be safe?
        # excel_mgr.workbook.Save() # Logic inside close handles save, but we need to read NOW.
        # We can read from open workbook.

        all_data = excel_mgr.get_all_data_for_sync()
        if all_data:
            replace_all_data(all_data)
        else:
            logger.warning("No data found in Excel to sync.")

        excel_mgr.close(save=True)
        logger.info("Excel closed and saved.")

    except Exception as e:
        logger.error(f"General error: {e}")
        # Try to close excel if open
        excel_mgr.close(save=False)
    finally:
        pythoncom.CoUninitialize()
        logger.info("--- Process Completed ---")

if __name__ == "__main__":
    main()
