"""
Utility e funzioni di supporto
"""

import logging
import sys
import os
import re
from datetime import datetime

from src.utils.config import LOG_FORMAT, LOG_LEVEL


def setup_logging(name="RDA_Bot", log_file=None):
    """
    Configura il sistema di logging.
    
    Args:
        name: Nome del logger
        log_file: Percorso opzionale per file di log
    
    Returns:
        Logger configurato
    """
    logger = logging.getLogger(name)
    
    if not logger.handlers:
        logger.setLevel(getattr(logging, LOG_LEVEL, logging.INFO))
        formatter = logging.Formatter(LOG_FORMAT)
        
        # Console handler
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)
        
        # File handler (opzionale)
        if log_file:
            try:
                file_handler = logging.FileHandler(log_file, encoding='utf-8')
                file_handler.setFormatter(formatter)
                logger.addHandler(file_handler)
            except Exception as e:
                logger.warning(f"Impossibile creare file di log: {e}")
    
    return logger


# Logger globale
logger = setup_logging()


def format_number(value):
    """
    Formatta un numero per la visualizzazione:
    - 2.0 diventa "2"
    - 2.5 diventa "2,5"
    - Stringhe come "25/039" restano invariate
    
    Args:
        value: Valore da formattare
    
    Returns:
        Stringa formattata
    """
    if value is None or value == "":
        return ""
    
    str_val = str(value)
    
    # Se contiene caratteri non numerici (esclusi . , - e spazi), lascia invariato
    if re.search(r'[^\d.,\-\s]', str_val):
        return str_val
    
    # Prova conversione numerica
    try:
        num = float(str_val.replace(',', '.'))
        if num == int(num):
            return str(int(num))
        else:
            return str(num).replace('.', ',')
    except (ValueError, TypeError):
        return str_val


def format_date(value, input_format="%d/%m/%Y", output_format="%d/%m/%Y"):
    """
    Formatta una data.
    
    Args:
        value: Data da formattare (stringa o datetime)
        input_format: Formato di input per stringhe
        output_format: Formato di output
    
    Returns:
        Stringa data formattata
    """
    if not value:
        return ""
    
    if isinstance(value, datetime):
        return value.strftime(output_format)
    
    if isinstance(value, str):
        try:
            dt = datetime.strptime(value, input_format)
            return dt.strftime(output_format)
        except ValueError:
            return value
    
    return str(value)


def parse_date(date_str, format="%d/%m/%Y"):
    """
    Converte una stringa in datetime.
    
    Args:
        date_str: Stringa data
        format: Formato della stringa
    
    Returns:
        datetime object o None se parsing fallisce
    """
    if not date_str:
        return None
    
    try:
        return datetime.strptime(str(date_str).strip(), format)
    except ValueError:
        return None


def safe_str(value, default=""):
    """
    Conversione sicura a stringa.
    
    Args:
        value: Valore da convertire
        default: Valore default se None
    
    Returns:
        Stringa
    """
    if value is None:
        return default
    return str(value)


def safe_float(value, default=0.0):
    """
    Conversione sicura a float.
    
    Args:
        value: Valore da convertire
        default: Valore default se conversione fallisce
    
    Returns:
        Float
    """
    if value is None:
        return default
    
    try:
        # Gestisce formato italiano con virgola
        str_val = str(value).replace(',', '.')
        return float(str_val)
    except (ValueError, TypeError):
        return default


def safe_int(value, default=0):
    """
    Conversione sicura a int.
    
    Args:
        value: Valore da convertire
        default: Valore default se conversione fallisce
    
    Returns:
        Int
    """
    if value is None:
        return default
    
    try:
        return int(float(str(value).replace(',', '.')))
    except (ValueError, TypeError):
        return default


def truncate_string(text, max_length=50, suffix="..."):
    """
    Tronca una stringa alla lunghezza massima.
    
    Args:
        text: Testo da troncare
        max_length: Lunghezza massima
        suffix: Suffisso da aggiungere se troncato
    
    Returns:
        Stringa troncata
    """
    if not text:
        return ""
    
    text = str(text)
    if len(text) <= max_length:
        return text
    
    return text[:max_length - len(suffix)] + suffix
