"""
Modulo per il parsing dei PDF RDA
"""

import pdfplumber
import re
import os
import shutil
import logging
from datetime import datetime
from src.utils.config import PDF_SAVE_PATH

logger = logging.getLogger("RDA_Bot")


def extract_rda_data(pdf_path):
    """
    Estrae i dati dal PDF RDA.
    
    Args:
        pdf_path: Percorso del file PDF
    
    Returns:
        dict: Dizionario con metadati e tabella articoli, None se errore
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                logger.error(f"PDF vuoto: {pdf_path}")
                return None
            
            page = pdf.pages[0]
            full_text = page.extract_text(x_tolerance=1, y_tolerance=1)
            
            if not full_text:
                logger.error(f"Impossibile estrarre testo da PDF: {pdf_path}")
                return None
            
            # Estrai numero RDA
            rda_match = re.search(r"Richiesta\s+di\s+Acquisto\s+([\d\/]+)", full_text)
            if not rda_match:
                logger.error(f"Numero RDA non trovato nel PDF: {pdf_path}")
                return None
            
            # Estrai data RDA
            date_match = re.search(r"del\s+([\d\/]+)", full_text)
            if not date_match:
                logger.error(f"Data RDA non trovata nel PDF: {pdf_path}")
                return None
            
            # Estrai richiedente (opzionale)
            requester_match = re.search(r"Richiedente\s*:?\s*(.+)", full_text, re.IGNORECASE)
            requester_name = ""
            if requester_match:
                requester_name = requester_match.group(1).strip()
                # Pulisci se ha catturato testo extra
                if "richiesta di acquisto" in requester_name.lower():
                    requester_name = ""
                # Limita lunghezza
                requester_name = requester_name[:100]
            
            rda_number_raw = rda_match.group(1).strip()
            rda_date_raw = date_match.group(1).strip()
            
            # Valida e converti data
            try:
                rda_date_obj = datetime.strptime(rda_date_raw, '%d/%m/%Y')
            except ValueError:
                logger.error(f"Formato data non valido nel PDF: {rda_date_raw}")
                return None
            
            # Estrai tabella articoli
            table = page.extract_table()
            if not table:
                logger.error(f"Nessuna tabella trovata nel PDF: {pdf_path}")
                return None
            
            return {
                'rda_number_raw': rda_number_raw,
                'rda_date_obj': rda_date_obj,
                'rda_date_str': rda_date_raw,
                'requester': requester_name,
                'table': table[1:]  # Salta header tabella
            }
            
    except Exception as e:
        logger.error(f"Errore parsing PDF {pdf_path}: {e}")
        return None


def save_pdf_to_archive(source_path, rda_number, rda_date_str):
    """
    Salva il PDF nell'archivio con nome standardizzato.
    
    Args:
        source_path: Percorso file PDF sorgente
        rda_number: Numero RDA (es. "25/01812")
        rda_date_str: Data RDA come stringa (es. "09/09/2025")
    
    Returns:
        str: Percorso finale del PDF salvato, None se errore
    """
    try:
        # Crea directory se non esiste
        if not os.path.exists(PDF_SAVE_PATH):
            os.makedirs(PDF_SAVE_PATH)
        
        # Formatta nome file
        # 25/01812 -> 25-01812
        rda_number_formatted = rda_number.replace('/', '-')
        
        # 09/09/2025 -> 09-09-25
        date_parts = rda_date_str.split('/')
        if len(date_parts) == 3:
            rda_date_formatted = f"{date_parts[0]}-{date_parts[1]}-{date_parts[2][-2:]}"
        else:
            rda_date_formatted = rda_date_str.replace('/', '-')
        
        new_filename = f"RDA_{rda_number_formatted}_{rda_date_formatted}.pdf"
        final_path = os.path.join(PDF_SAVE_PATH, new_filename)
        
        # Copia solo se non esiste già
        if not os.path.exists(final_path):
            shutil.copy2(source_path, final_path)
            logger.info(f"PDF salvato in archivio: {new_filename}")
        else:
            logger.info(f"PDF già presente in archivio: {new_filename}")
        
        return final_path
        
    except Exception as e:
        logger.error(f"Errore salvataggio PDF in archivio: {e}")
        return None


def validate_pdf(pdf_path):
    """
    Valida che un file sia un PDF valido.
    
    Args:
        pdf_path: Percorso del file
    
    Returns:
        bool: True se valido, False altrimenti
    """
    if not os.path.exists(pdf_path):
        return False
    
    if not pdf_path.lower().endswith('.pdf'):
        return False
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            return len(pdf.pages) > 0
    except:
        return False
