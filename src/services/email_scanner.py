"""
Modulo per la scansione delle email Outlook
"""

import win32com.client
import os
import logging
from datetime import datetime, timedelta
from src.utils.config import (
    OUTLOOK_FOLDER_ID, TARGET_FOLDER_NAME, DAYS_TO_CHECK, 
    ATTACHMENT_NAME, SENDER_EMAIL, EMAIL_SENDER, EMAIL_RECIPIENT, EMAIL_SUBJECT
)
from src.utils.utils import format_number

logger = logging.getLogger("RDA_Bot")


class EmailScanner:
    """
    Scanner per email Outlook con allegati PDF RDA.
    """
    
    def __init__(self):
        self.outlook = None
        self.namespace = None
        self._connected = False
    
    def connect(self):
        """
        Stabilisce connessione con Outlook.
        
        Returns:
            bool: True se connessione riuscita, False altrimenti
        """
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self._connected = True
            return True
        except Exception as e:
            logger.error(f"Errore connessione Outlook: {e}")
            return False
    
    def get_messages(self):
        """
        Recupera i messaggi dalla cartella target.
        
        Returns:
            Collection di messaggi Outlook o lista vuota
        """
        if not self._connected:
            return []
        
        try:
            inbox = self.namespace.GetDefaultFolder(OUTLOOK_FOLDER_ID)
            
            # Prova cartella specifica, fallback su Inbox
            try:
                target_folder = inbox.Folders(TARGET_FOLDER_NAME)
                logger.info(f"Scansione cartella: {TARGET_FOLDER_NAME}")
            except Exception:
                logger.warning(f"Cartella '{TARGET_FOLDER_NAME}' non trovata. Uso Posta in arrivo.")
                target_folder = inbox
            
            messages = target_folder.Items
            messages.Sort("[ReceivedTime]", True)  # Più recenti prima
            
            return messages
            
        except Exception as e:
            logger.error(f"Errore recupero messaggi: {e}")
            return []
    
    def process_emails(self, messages, callback_process_pdf):
        """
        Processa le email cercando allegati PDF RDA.
        
        Args:
            messages: Collection di messaggi Outlook
            callback_process_pdf: Funzione da chiamare per ogni PDF trovato
        """
        if not messages:
            return
        
        limit_date = datetime.now() - timedelta(days=DAYS_TO_CHECK)
        logger.info(f"Scansione email dal {limit_date.strftime('%d/%m/%Y')}")
        
        processed_count = 0
        
        for message in messages:
            try:
                # Salta elementi non-mail (inviti, ecc.)
                if getattr(message, "Class", 0) != 43:
                    continue
                
                # Controllo limite temporale
                received_date = message.ReceivedTime.date()
                if received_date < limit_date.date():
                    logger.info("Raggiunto limite temporale. Stop scansione.")
                    break
                
                # Cerca allegato corretto
                found_attachment = False
                for attachment in message.Attachments:
                    if attachment.FileName.lower() == ATTACHMENT_NAME.lower():
                        # Verifica mittente
                        sender = self._resolve_sender(message)
                        if sender and sender.lower() == SENDER_EMAIL.lower():
                            logger.info(f"Trovata RDA da {sender} - {message.ReceivedTime}")
                            
                            # Salva allegato temporaneamente
                            temp_path = self._save_temp_attachment(attachment)
                            if temp_path:
                                callback_process_pdf(temp_path)
                                self._cleanup_temp(temp_path)
                                found_attachment = True
                                processed_count += 1
                
                # Segna come letta
                if found_attachment and message.UnRead:
                    message.UnRead = False
                    
            except Exception as e:
                logger.warning(f"Errore elaborazione email: {e}")
                continue
        
        logger.info(f"Elaborate {processed_count} email con allegati RDA")
    
    def send_summary_email(self, overdue_items):
        """
        Invia email di riepilogo per RDA scadute.
        
        Args:
            overdue_items: Lista di dict con dati RDA scadute
        """
        if not overdue_items:
            logger.info("Nessuna RDA scaduta da notificare")
            return
        
        if not self._connected:
            logger.error("Non connesso a Outlook")
            return
        
        logger.info(f"Invio email riepilogo per {len(overdue_items)} RDA scadute")
        
        # Costruisci HTML
        html_body = self._build_summary_html(overdue_items)
        
        try:
            mail = self.outlook.CreateItem(0)
            mail.To = EMAIL_RECIPIENT
            mail.Subject = EMAIL_SUBJECT
            mail.SentOnBehalfOfName = EMAIL_SENDER
            mail.HTMLBody = html_body
            mail.Send()
            
            logger.info("Email di riepilogo inviata con successo")
            
        except Exception as e:
            logger.error(f"Errore invio email: {e}")
    
    def _resolve_sender(self, message):
        """Risolve l'indirizzo email del mittente."""
        try:
            sender_address = message.SenderEmailAddress
            
            # Gestisci indirizzi Exchange
            if message.SenderEmailType == "EX":
                try:
                    sender_user = message.Sender.GetExchangeUser()
                    if sender_user:
                        sender_address = sender_user.PrimarySmtpAddress
                except:
                    pass
            
            return sender_address
            
        except:
            return None
    
    def _save_temp_attachment(self, attachment):
        """Salva un allegato in una cartella temporanea."""
        try:
            temp_dir = os.path.join(os.environ.get('TEMP', '/tmp'), "rda_outlook_processing")
            os.makedirs(temp_dir, exist_ok=True)
            
            temp_path = os.path.join(temp_dir, attachment.FileName)
            attachment.SaveAsFile(temp_path)
            
            return temp_path
            
        except Exception as e:
            logger.error(f"Errore salvataggio allegato temporaneo: {e}")
            return None
    
    def _cleanup_temp(self, temp_path):
        """Elimina un file temporaneo."""
        try:
            if os.path.exists(temp_path):
                os.remove(temp_path)
        except:
            pass
    
    def _build_summary_html(self, overdue_items):
        """Costruisce l'HTML per l'email di riepilogo."""
        html = """
        <html>
        <head>
        <style>
            body { font-family: 'Segoe UI', Arial, sans-serif; }
            table { border-collapse: collapse; width: 100%; margin-top: 20px; }
            th, td { border: 1px solid #ddd; text-align: left; padding: 10px; }
            th { background-color: #2563EB; color: white; }
            tr:nth-child(even) { background-color: #f9f9f9; }
            tr:hover { background-color: #f1f1f1; }
            .header { color: #1F2937; margin-bottom: 10px; }
        </style>
        </head>
        <body>
        <h2 class="header">Riepilogo RDA Non Evase</h2>
        <p>Con la presente per comunicare le Richieste di Acquisto non ancora evase:</p>
        <table>
        <tr>
            <th>N° RDA</th>
            <th>Data RDA</th>
            <th>Commessa</th>
            <th>Descrizione Materiale</th>
            <th>UM</th>
            <th>Quantità</th>
            <th>APF</th>
            <th>Giorni Trascorsi</th>
            <th>Richiedente</th>
        </tr>
        """
        
        keys = [
            "N°RDA", "Data RDA", "Commessa", "Descrizione Materiale",
            "Unità di Misura", "Quantità Richiesta", "APF",
            "richiesta da: (giorni)", "Richiedente"
        ]
        
        for item in overdue_items:
            html += "<tr>"
            for key in keys:
                value = item.get(key)
                
                # Formattazione numeri
                if key == "Commessa" and isinstance(value, (float, int)):
                    value = format_number(value)
                elif key == "Quantità Richiesta" and isinstance(value, (float, int)):
                    value = format_number(value)
                
                html += f"<td>{value if value is not None else ''}</td>"
            html += "</tr>"
        
        html += """
        </table>
        <p style="margin-top: 20px; color: #6B7280; font-size: 12px;">
            Email generata automaticamente dal sistema RDA Bot.
        </p>
        </body>
        </html>
        """
        
        return html
