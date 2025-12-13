"""
RDA Automation System - Moduli Sorgente

Questo pacchetto contiene i moduli per:
- Gestione database SQLite
- Automazione Excel via COM
- Parsing PDF RDA
- Scansione email Outlook
- Configurazione e utilità
"""

from .config import *
from .utils import logger, format_number, format_date
from .database import get_connection, init_db, replace_all_data

__version__ = "2.0.0"
__author__ = "RDA Team"

