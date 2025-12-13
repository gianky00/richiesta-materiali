import win32com.client
import os
import logging
from datetime import datetime, timedelta
from .config import OUTLOOK_FOLDER_ID, TARGET_FOLDER_NAME, DAYS_TO_CHECK, ATTACHMENT_NAME, SENDER_EMAIL, EMAIL_SENDER, EMAIL_RECIPIENT, EMAIL_SUBJECT

logger = logging.getLogger("RDA_Bot")

class EmailScanner:
    def __init__(self):
        self.outlook = None
        self.namespace = None

    def connect(self):
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            return True
        except Exception as e:
            logger.error(f"Error connecting to Outlook: {e}")
            return False

    def get_messages(self):
        try:
            inbox = self.namespace.GetDefaultFolder(OUTLOOK_FOLDER_ID)
            try:
                target_folder = inbox.Folders(TARGET_FOLDER_NAME)
                logger.info(f"Scanning folder: {TARGET_FOLDER_NAME}")
            except Exception:
                logger.warning(f"Folder '{TARGET_FOLDER_NAME}' not found. Scanning Inbox.")
                target_folder = inbox

            messages = target_folder.Items
            messages.Sort("[ReceivedTime]", True) # Newest first
            return messages
        except Exception as e:
            logger.error(f"Error retrieving messages: {e}")
            return []

    def process_emails(self, messages, callback_process_pdf):
        limit_date = datetime.now() - timedelta(days=DAYS_TO_CHECK)
        logger.info(f"Scanning emails newer than {limit_date.strftime('%d/%m/%Y')}")

        for message in messages:
            try:
                if getattr(message, "Class", 0) != 43: continue # MailItem

                # Timezone naive comparison
                if message.ReceivedTime.date() < limit_date.date():
                    logger.info("Reached time limit. Stopping scan.")
                    break

                found_attachment = False
                for attachment in message.Attachments:
                    if attachment.FileName.lower() == ATTACHMENT_NAME.lower():
                        sender = self._resolve_sender(message)
                        if sender and sender.lower() == SENDER_EMAIL.lower():
                            logger.info(f"Found RDA from {sender} received on {message.ReceivedTime}")

                            temp_dir = os.path.join(os.environ['TEMP'], "rda_outlook_processing")
                            os.makedirs(temp_dir, exist_ok=True)
                            temp_path = os.path.join(temp_dir, attachment.FileName)

                            attachment.SaveAsFile(temp_path)
                            callback_process_pdf(temp_path)

                            try:
                                os.remove(temp_path)
                            except:
                                pass

                            found_attachment = True

                if found_attachment and message.UnRead:
                    message.UnRead = False

            except Exception as e:
                logger.warning(f"Error processing email: {e}")
                continue

    def _resolve_sender(self, message):
        try:
            sender_address = message.SenderEmailAddress
            if message.SenderEmailType == "EX":
                sender_user = message.Sender.GetExchangeUser()
                if sender_user:
                    sender_address = sender_user.PrimarySmtpAddress
            return sender_address
        except:
            return None

    def send_summary_email(self, overdue_items):
        if not overdue_items:
            logger.info("No overdue items to report.")
            return

        logger.info(f"Sending summary email for {len(overdue_items)} overdue items.")

        html_body = """<html><head><style>body{font-family:sans-serif}table{border-collapse:collapse;width:100%}th,td{border:1px solid #ddd;text-align:left;padding:8px}th{background-color:#f2f2f2}tr:nth-child(even){background-color:#f9f9f9}</style></head><body><p>Con la presente per comunicare le RDA non evase:</p><table><tr><th>N°RDA</th><th>Data RDA</th><th>Commessa</th><th>Descrizione Materiale</th><th>Unità di Misura</th><th>Quantità Richiesta</th><th>APF</th><th>richiesta da: (giorni)</th><th>Richiedente</th></tr>"""

        for item in overdue_items:
            html_body += "<tr>"
            keys = ["N°RDA", "Data RDA", "Commessa", "Descrizione Materiale", "Unità di Misura", "Quantità Richiesta", "APF", "richiesta da: (giorni)", "Richiedente"]
            for key in keys:
                value = item.get(key)
                # Formatting logic
                if key == "Commessa" and isinstance(value, (float, int)):
                    value = int(value)
                elif key == "Quantità Richiesta" and isinstance(value, (float, int)):
                    if value == int(value): value = int(value)
                    else: value = str(value).replace('.', ',')

                html_body += f"<td>{value if value is not None else ''}</td>"
            html_body += "</tr>"
        html_body += "</table></body></html>"

        try:
            mail = self.outlook.CreateItem(0)
            mail.To = EMAIL_RECIPIENT
            mail.Subject = EMAIL_SUBJECT
            mail.SentOnBehalfOfName = EMAIL_SENDER
            mail.HTMLBody = html_body
            mail.Send()
            logger.info("Summary email sent.")
        except Exception as e:
            logger.error(f"Failed to send email: {e}")
