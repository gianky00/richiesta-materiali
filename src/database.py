"""
Modulo per la gestione del database SQLite
"""

import sqlite3
import logging
import os
from .config import SQLITE_DB_PATH, DATABASE_DIR

logger = logging.getLogger("RDA_Bot")


def get_connection():
    """
    Ottiene una connessione al database SQLite.
    Crea il file database se non esiste.
    """
    try:
        # Assicura che la directory esista
        if not os.path.exists(DATABASE_DIR):
            os.makedirs(DATABASE_DIR)
        
        conn = sqlite3.connect(SQLITE_DB_PATH, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        return conn
    except Exception as e:
        logger.error(f"Errore connessione database: {e}")
        raise


def init_db():
    """
    Inizializza il database creando le tabelle necessarie.
    Sicuro da chiamare più volte (CREATE IF NOT EXISTS).
    """
    conn = get_connection()
    cursor = conn.cursor()
    
    # Tabella principale RDA
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS rda_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rda_number TEXT,
            commessa TEXT,
            descrizione_1 TEXT,
            descrizione_materiale TEXT,
            unita_misura TEXT,
            quantita REAL,
            apf TEXT,
            pdf_path TEXT,
            data_rda TEXT,
            data_consegna TEXT,
            alert_level INTEGER DEFAULT 0,
            richiedente TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    
    # Indici per migliorare le performance delle query
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_rda_number ON rda_data(rda_number)
    """)
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_richiedente ON rda_data(richiedente)
    """)
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_data_rda ON rda_data(data_rda)
    """)
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_alert_level ON rda_data(alert_level)
    """)
    
    conn.commit()
    conn.close()
    logger.info("Database inizializzato correttamente")


def replace_all_data(rows):
    """
    Sostituisce tutti i dati nella tabella con le righe fornite.
    Operazione atomica con transazione.
    
    Args:
        rows: Lista di tuple con i dati (escludendo ID e created_at)
    """
    if not rows:
        logger.warning("Nessun dato da inserire")
        return
    
    conn = get_connection()
    cursor = conn.cursor()
    
    try:
        cursor.execute("BEGIN TRANSACTION")
        
        # Elimina tutti i dati esistenti
        cursor.execute("DELETE FROM rda_data")
        
        # Reset contatore auto-increment
        cursor.execute("DELETE FROM sqlite_sequence WHERE name='rda_data'")
        
        # Inserisci nuovi dati
        cursor.executemany("""
            INSERT INTO rda_data (
                rda_number, commessa, descrizione_1, descrizione_materiale,
                unita_misura, quantita, apf, pdf_path, data_rda,
                data_consegna, alert_level, richiedente
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, rows)
        
        cursor.execute("COMMIT")
        logger.info(f"Database sincronizzato: {len(rows)} righe inserite")
        
    except Exception as e:
        cursor.execute("ROLLBACK")
        logger.error(f"Errore durante l'inserimento dati: {e}")
        raise
    finally:
        conn.close()


def get_all_rows():
    """
    Recupera tutte le righe dal database.
    
    Returns:
        Lista di righe come sqlite3.Row objects
    """
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT * FROM rda_data 
        ORDER BY data_rda DESC, rda_number
    """)
    rows = cursor.fetchall()
    conn.close()
    return rows


def get_statistics():
    """
    Calcola statistiche aggregate sui dati.
    
    Returns:
        Dictionary con varie statistiche
    """
    conn = get_connection()
    cursor = conn.cursor()
    
    stats = {}
    
    # Totale RDA univoche
    cursor.execute("SELECT COUNT(DISTINCT rda_number) FROM rda_data")
    stats['total_rda'] = cursor.fetchone()[0] or 0
    
    # Totale righe
    cursor.execute("SELECT COUNT(*) FROM rda_data")
    stats['total_rows'] = cursor.fetchone()[0] or 0
    
    # RDA con alert attivo
    cursor.execute("SELECT COUNT(DISTINCT rda_number) FROM rda_data WHERE alert_level > 0")
    stats['overdue_rda'] = cursor.fetchone()[0] or 0
    
    # Top richiedenti
    cursor.execute("""
        SELECT richiedente, COUNT(DISTINCT rda_number) as cnt 
        FROM rda_data 
        WHERE richiedente IS NOT NULL AND richiedente != ''
        GROUP BY richiedente 
        ORDER BY cnt DESC
        LIMIT 10
    """)
    stats['by_requester'] = cursor.fetchall()
    
    conn.close()
    return stats


def search_rda(filters):
    """
    Ricerca RDA con filtri multipli.
    
    Args:
        filters: Dictionary con chiavi filtro (rda_number, richiedente, data_from, data_to, apf)
    
    Returns:
        Lista di righe matching
    """
    conn = get_connection()
    cursor = conn.cursor()
    
    query = "SELECT * FROM rda_data WHERE 1=1"
    params = []
    
    if filters.get('rda_number'):
        query += " AND rda_number LIKE ?"
        params.append(f"%{filters['rda_number']}%")
    
    if filters.get('richiedente'):
        query += " AND richiedente LIKE ?"
        params.append(f"%{filters['richiedente']}%")
    
    if filters.get('apf'):
        query += " AND apf LIKE ?"
        params.append(f"%{filters['apf']}%")
    
    if filters.get('only_overdue'):
        query += " AND alert_level > 0"
    
    query += " ORDER BY data_rda DESC"
    
    cursor.execute(query, params)
    rows = cursor.fetchall()
    conn.close()
    
    return rows
