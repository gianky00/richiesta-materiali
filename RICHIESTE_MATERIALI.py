import os
import re
import shutil
import win32com.client
import pythoncom
import pdfplumber
from datetime import datetime, timedelta
import logging
import time

# --- CONFIGURAZIONE ---
SENDER_EMAIL = "magonet@coemi.it"
ATTACHMENT_NAME = "RDAPerFornitore.pdf"
PDF_SAVE_PATH = r"\\192.168.11.251\Condivisa\RICHIESTE MATERIALI\RDA_PDF"
EXCEL_DB_PATH = r"\\192.168.11.251\Condivisa\RICHIESTE MATERIALI\DATABASE\database_RDA.xlsm"
SHEET_PASSWORD = "coemi"
OUTLOOK_FOLDER_ID = 6
TABLE_NAME = "Tabella1"
TARGET_FOLDER_NAME = "MAGO"

# Ottimizzazione Velocità: Controlla solo le email degli ultimi X giorni
DAYS_TO_CHECK = 60 

# La colonna con il riferimento univoco (N° RDA)
RDA_REFERENCE_COLUMN_LETTER = "A"
RDA_REFERENCE_COLUMN_NUMBER = 1

# Impostazioni Email
EMAIL_SENDER = "isabsud@coemi.it"
EMAIL_RECIPIENT = "isabsud@coemi.it;concetto.siringo@coemi.it"
EMAIL_SUBJECT = "RIEPILOGO RDA SCADUTE"
# --- FINE CONFIGURAZIONE ---

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def check_if_data_exists(sheet, rda_number_to_check):
    try:
        last_row = sheet.Cells(sheet.Rows.Count, RDA_REFERENCE_COLUMN_LETTER).End(-4162).Row
        for i in range(2, last_row + 1):
            cell_value = sheet.Cells(i, RDA_REFERENCE_COLUMN_NUMBER).Value
            if cell_value and str(cell_value).strip() == rda_number_to_check:
                return True
        return False
    except Exception as e:
        logging.error(f"Errore verifica esistenza dati: {e}")
        return True

def append_to_excel(sheet, rda_data):
    try:
        first_empty_row = sheet.Cells(sheet.Rows.Count, RDA_REFERENCE_COLUMN_LETTER).End(-4162).Row + 1
        valid_rows = [row for row in rda_data['table'] if any(cell is not None and str(cell).strip() != '' for cell in row)]

        for row_data in valid_rows:
            cleaned_row = [item if item is not None else "" for item in row_data]
            delivery_date_str = cleaned_row[8]
            delivery_date_obj = None
            if delivery_date_str:
                try:
                    delivery_date_obj = datetime.strptime(str(delivery_date_str), '%d/%m/%Y')
                except (ValueError, TypeError):
                    delivery_date_obj = delivery_date_str

            quantity_val = cleaned_row[5]
            try:
                quantity_str = str(quantity_val).replace('.', '').replace(',', '.')
                quantity_val = float(quantity_str)
            except (ValueError, TypeError):
                pass

            new_row_values = [
                rda_data['rda_number_raw'], cleaned_row[1], cleaned_row[2], cleaned_row[3],
                cleaned_row[4], quantity_val, cleaned_row[7],
                f'=HYPERLINK("{rda_data["pdf_final_path"]}", "Apri PDF")',
                rda_data['rda_date_obj'], delivery_date_obj, 0, rda_data.get('requester', '')
            ]

            for i, value in enumerate(new_row_values):
                sheet.Cells(first_empty_row, i + 1).Value = value
            first_empty_row += 1
        logging.info(f"Dati per RDA {rda_data['rda_number_raw']} aggiunti correttamente.")
    except Exception as e:
        logging.error(f"Errore durante l'inserimento dati per RDA {rda_data['rda_number_raw']}: {e}")

def update_alerts_and_finalize(sheet):
    overdue_items = []
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    try:
        last_row = sheet.Cells(sheet.Rows.Count, RDA_REFERENCE_COLUMN_LETTER).End(-4162).Row
        
        for row in range(2, last_row + 1):
            rda_date = None
            rda_date_cell_val = sheet.Cells(row, 9).Value
            
            if hasattr(rda_date_cell_val, 'timestamp'):
                rda_date = datetime.fromtimestamp(rda_date_cell_val.timestamp())
            elif isinstance(rda_date_cell_val, str):
                try:
                    rda_date = datetime.strptime(rda_date_cell_val, '%d/%m/%Y')
                except ValueError:
                    continue

            if rda_date:
                days_diff = (today - rda_date).days
                alert_level = days_diff // 7 if days_diff >= 7 else 0
                sheet.Cells(row, 11).Value = alert_level
                
                if alert_level > 0:
                    delivery_date = None
                    delivery_date_cell_val = sheet.Cells(row, 10).Value

                    if hasattr(delivery_date_cell_val, 'timestamp'):
                        delivery_date = datetime.fromtimestamp(delivery_date_cell_val.timestamp())
                    elif isinstance(delivery_date_cell_val, str):
                        try:
                            delivery_date = datetime.strptime(delivery_date_cell_val, '%d/%m/%Y')
                        except ValueError:
                            pass

                    if delivery_date and delivery_date.replace(hour=0, minute=0, second=0, microsecond=0) <= today:
                        continue
                        
                    item_data = {
                        "N°RDA": sheet.Cells(row, 1).Value, "Data RDA": rda_date.strftime('%d/%m/%Y'),
                        "Commessa": sheet.Cells(row, 2).Value, "Descrizione Materiale": sheet.Cells(row, 4).Value,
                        "Unità di Misura": sheet.Cells(row, 5).Value, "Quantità Richiesta": sheet.Cells(row, 6).Value,
                        "APF": sheet.Cells(row, 7).Value, "richiesta da: (giorni)": days_diff,
                        "Richiedente": sheet.Cells(row, 12).Value
                    }
                    overdue_items.append(item_data)

        if last_row < 2: last_row = 1
        new_range = sheet.Range(f"A1:L{last_row}")
        sheet.ListObjects(TABLE_NAME).Resize(new_range)
        logging.info(f"Tabella '{TABLE_NAME}' aggiornata e dimensionata correttamente.")
        return overdue_items
    except Exception as e:
        logging.error(f"Errore durante l'aggiornamento e finalizzazione: {e}")
        return []

def delete_empty_rows_from_table(excel, sheet):
    try:
        logging.info("Avvio eliminazione righe vuote dalla tabella...")
        table = sheet.ListObjects(TABLE_NAME)
        for i in range(table.ListRows.Count, 0, -1):
            list_row = table.ListRows(i)
            if excel.WorksheetFunction.CountA(list_row.Range) == 0:
                list_row.Delete()
        logging.info("Eliminazione righe vuote completata.")
    except Exception as e:
        logging.error(f"Errore eliminazione righe vuote: {e}")

def send_summary_email(overdue_items):
    if not overdue_items:
        logging.info("Nessuna RDA scaduta da notificare."); return
    logging.info(f"Trovate {len(overdue_items)} RDA da notificare. Invio email...")
    html_body = """<html><head><style>body{font-family:sans-serif}table{border-collapse:collapse;width:100%}th,td{border:1px solid #ddd;text-align:left;padding:8px}th{background-color:#f2f2f2}tr:nth-child(even){background-color:#f9f9f9}</style></head><body><p>Con la presente per comunicare le RDA non evase:</p><table><tr><th>N°RDA</th><th>Data RDA</th><th>Commessa</th><th>Descrizione Materiale</th><th>Unità di Misura</th><th>Quantità Richiesta</th><th>APF</th><th>richiesta da: (giorni)</th><th>Richiedente</th></tr>"""
    for item in overdue_items:
        html_body += "<tr>"
        for key in ["N°RDA", "Data RDA", "Commessa", "Descrizione Materiale", "Unità di Misura", "Quantità Richiesta", "APF", "richiesta da: (giorni)", "Richiedente"]:
            value = item.get(key)
            if key == "Commessa" and isinstance(value, (float, int)):
                value = int(value)
            elif key == "Quantità Richiesta" and isinstance(value, (float, int)):
                if value == int(value): value = int(value)
                else: value = str(value).replace('.', ',')
            html_body += f"<td>{value if value is not None else ''}</td>"
        html_body += "</tr>"
    html_body += "</table></body></html>"
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = EMAIL_RECIPIENT; mail.Subject = EMAIL_SUBJECT
        mail.SentOnBehalfOfName = EMAIL_SENDER; mail.HTMLBody = html_body
        mail.Send()
        logging.info("Email di riepilogo inviata con successo.")
    except Exception as e:
        logging.error(f"Impossibile inviare l'email: {e}")

def process_new_pdf(temp_pdf_path, sheet):
    try:
        with pdfplumber.open(temp_pdf_path) as pdf:
            page = pdf.pages[0]; full_text = page.extract_text(x_tolerance=1, y_tolerance=1)
            rda_match = re.search(r"Richiesta\s+di\s+Acquisto\s+([\d\/]+)", full_text)
            date_match = re.search(r"del\s+([\d\/]+)", full_text)
            requester_match = re.search(r"Richiedente\s*:?\s*(.+)", full_text, re.IGNORECASE)
            
            if not (rda_match and date_match):
                logging.error("N° RDA o Data non trovati nel PDF."); return
            
            requester_name = requester_match.group(1).strip() if requester_match else ""
            if "richiesta di acquisto" in requester_name.lower(): requester_name = ""

            rda_number_raw = rda_match.group(1).strip()
            rda_date_raw = date_match.group(1).strip()
            
            if check_if_data_exists(sheet, rda_number_raw):
                logging.warning(f"Dati per RDA '{rda_number_raw}' già presenti. Salto."); return
            
            logging.info(f"Dati per RDA '{rda_number_raw}' non trovati in Excel. Procedo.")
            rda_number_formatted = rda_number_raw.replace('/', '-')
            date_parts = rda_date_raw.split('/')
            rda_date_formatted = f"{date_parts[0]}-{date_parts[1]}-{date_parts[2][-2:]}"
            new_filename = f"RDA_{rda_number_formatted}_{rda_date_formatted}.pdf"
            final_path = os.path.join(PDF_SAVE_PATH, new_filename)
            
            if not os.path.exists(final_path):
                shutil.copy2(temp_pdf_path, final_path); logging.info(f"Nuovo PDF salvato: '{new_filename}'")
            else:
                logging.info(f"PDF '{new_filename}' già esistente.")

            table = page.extract_table()
            if not table: logging.error("Nessuna tabella trovata nel PDF."); return
            
            rda_data = {
                'rda_number_raw': rda_number_raw, 'rda_date_obj': datetime.strptime(rda_date_raw, '%d/%m/%Y'),
                'table': table[1:], 'pdf_final_path': final_path, 'requester': requester_name
            }
            append_to_excel(sheet, rda_data)
    except Exception as e:
        logging.error(f"Errore grave durante l'elaborazione del PDF: {e}")

def main_process():
    for path in [PDF_SAVE_PATH, os.path.dirname(EXCEL_DB_PATH)]:
        if not os.path.exists(path):
            logging.error(f"Il percorso '{path}' non esiste. Crealo e riavvia."); return

    logging.info("Avvio script...")
    excel = None; workbook = None
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False; excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(EXCEL_DB_PATH)
        sheet = workbook.ActiveSheet
        sheet.Unprotect(Password=SHEET_PASSWORD)
        
        logging.info("Controllo Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        inbox = outlook.GetNamespace("MAPI").GetDefaultFolder(OUTLOOK_FOLDER_ID)
        
        try:
            target_folder = inbox.Folders(TARGET_FOLDER_NAME)
            logging.info(f"Accesso alla cartella '{TARGET_FOLDER_NAME}' riuscito.")
        except Exception:
            logging.error(f"ERRORE: La cartella '{TARGET_FOLDER_NAME}' non è stata trovata. Controllo Posta in arrivo.")
            target_folder = inbox
        
        messages = target_folder.Items
        messages.Sort("[ReceivedTime]", True) # Ordina: recenti -> vecchie
        
        # --- OTTIMIZZAZIONE TEMPO ---
        limit_date = datetime.now() - timedelta(days=DAYS_TO_CHECK)
        logging.info(f"Scansione email (limite temporale: {limit_date.strftime('%d/%m/%Y')})...")

        for message in messages:
            try:
                # 1. Salta elementi non-mail (es. inviti)
                if getattr(message, "Class", 0) != 43: continue
                
                # 2. CONTROLLO DATA: Se l'email è più vecchia del limite, STOP TOTALE
                #    (Usiamo .date() per evitare conflitti di timezone)
                if message.ReceivedTime.date() < limit_date.date():
                    logging.info("Raggiunto limite temporale di scansione. Stop ricerca email.")
                    break
                
                # 3. Procedi solo se ha allegati
                found_attachment = False
                for attachment in message.Attachments:
                    if attachment.FileName.lower() == ATTACHMENT_NAME.lower():
                        # Risoluzione Mittente
                        sender_address = message.SenderEmailAddress
                        try:
                            if message.SenderEmailType == "EX":
                                sender_user = message.Sender.GetExchangeUser()
                                if sender_user: sender_address = sender_user.PrimarySmtpAddress
                        except: pass
                        
                        if sender_address and sender_address.lower() == SENDER_EMAIL.lower():
                            logging.info(f"Trovata RDA da '{sender_address}' del {message.ReceivedTime}")
                            temp_dir = os.path.join(os.environ['TEMP'], "rda_outlook_processing")
                            os.makedirs(temp_dir, exist_ok=True)
                            temp_pdf_path = os.path.join(temp_dir, attachment.FileName)
                            
                            attachment.SaveAsFile(temp_pdf_path)
                            process_new_pdf(temp_pdf_path, sheet)
                            os.remove(temp_pdf_path)
                            found_attachment = True

                if found_attachment and message.UnRead:
                    message.UnRead = False
            
            except Exception as e:
                # Se errore su singola email, continua con la successiva
                logging.warning(f"Errore lettura email (ignorata): {e}")
                continue 
        
        logging.info("Aggiornamento scadenze Excel...")
        overdue_rdas = update_alerts_and_finalize(sheet)
        send_summary_email(overdue_rdas)
        delete_empty_rows_from_table(excel, sheet)
        sheet.Columns.AutoFit()
        
        sheet.Protect(Password=SHEET_PASSWORD)
        workbook.Close(SaveChanges=True)
        logging.info("File Excel salvato e chiuso correttamente.")
        workbook = None

    except Exception as e:
        logging.error(f"Si è verificato un errore generale: {e}")
    finally:
        if workbook: workbook.Close(SaveChanges=False)
        if excel: excel.Quit()
        pythoncom.CoUninitialize()
        logging.info("--- Scansione e aggiornamento completati ---")

if __name__ == "__main__":
    main_process()
