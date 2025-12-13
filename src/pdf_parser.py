import pdfplumber
import re
import os
import shutil
import logging
from datetime import datetime
from .config import PDF_SAVE_PATH

logger = logging.getLogger("RDA_Bot")

def extract_rda_data(pdf_path):
    """
    Extracts data from the RDA PDF.
    Returns a dictionary with metadata and the table of items.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            full_text = page.extract_text(x_tolerance=1, y_tolerance=1)

            rda_match = re.search(r"Richiesta\s+di\s+Acquisto\s+([\d\/]+)", full_text)
            date_match = re.search(r"del\s+([\d\/]+)", full_text)
            requester_match = re.search(r"Richiedente\s*:?\s*(.+)", full_text, re.IGNORECASE)

            if not (rda_match and date_match):
                logger.error(f"RDA Number or Date not found in PDF {pdf_path}")
                return None

            requester_name = requester_match.group(1).strip() if requester_match else ""
            if "richiesta di acquisto" in requester_name.lower():
                requester_name = ""

            rda_number_raw = rda_match.group(1).strip()
            rda_date_raw = date_match.group(1).strip()

            try:
                rda_date_obj = datetime.strptime(rda_date_raw, '%d/%m/%Y')
            except ValueError:
                logger.error(f"Invalid date format in PDF: {rda_date_raw}")
                return None

            table = page.extract_table()
            if not table:
                logger.error(f"No table found in PDF {pdf_path}")
                return None

            return {
                'rda_number_raw': rda_number_raw,
                'rda_date_obj': rda_date_obj,
                'rda_date_str': rda_date_raw,
                'requester': requester_name,
                'table': table[1:] # Skip header
            }
    except Exception as e:
        logger.error(f"Error parsing PDF {pdf_path}: {e}")
        return None

def save_pdf_to_archive(source_path, rda_number, rda_date_str):
    """
    Saves the PDF to the network archive with a standardized name.
    Returns the final path.
    """
    try:
        if not os.path.exists(PDF_SAVE_PATH):
            os.makedirs(PDF_SAVE_PATH) # Should exist, but just in case

        rda_number_formatted = rda_number.replace('/', '-')
        date_parts = rda_date_str.split('/')
        # Assuming DD/MM/YYYY
        rda_date_formatted = f"{date_parts[0]}-{date_parts[1]}-{date_parts[2][-2:]}"

        new_filename = f"RDA_{rda_number_formatted}_{rda_date_formatted}.pdf"
        final_path = os.path.join(PDF_SAVE_PATH, new_filename)

        if not os.path.exists(final_path):
            shutil.copy2(source_path, final_path)
            logger.info(f"Saved PDF to archive: {new_filename}")
        else:
            logger.info(f"PDF already exists in archive: {new_filename}")

        return final_path
    except Exception as e:
        logger.error(f"Error saving PDF to archive: {e}")
        return None
