import sqlite3
import logging
from .config import SQLITE_DB_PATH

logger = logging.getLogger("RDA_Bot")

def get_connection():
    try:
        conn = sqlite3.connect(SQLITE_DB_PATH)
        conn.row_factory = sqlite3.Row
        return conn
    except Exception as e:
        logger.error(f"Failed to connect to database: {e}")
        raise

def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    # Using a surrogate ID since RDA Number is not unique (one RDA has multiple items)
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
            alert_level INTEGER,
            richiedente TEXT
        )
    """)
    conn.commit()
    conn.close()

def replace_all_data(rows):
    """
    Replaces all data in the table with the provided list of rows.
    'rows' should be a list of tuples/lists matching the columns (excluding ID).
    """
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM rda_data")
        cursor.execute("DELETE FROM sqlite_sequence WHERE name='rda_data'") # Reset ID counter

        cursor.executemany("""
            INSERT INTO rda_data (
                rda_number, commessa, descrizione_1, descrizione_materiale,
                unita_misura, quantita, apf, pdf_path, data_rda,
                data_consegna, alert_level, richiedente
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, rows)

        conn.commit()
        logger.info(f"Database synchronized. {cursor.rowcount} rows inserted.")
    except Exception as e:
        logger.error(f"Error replacing data in DB: {e}")
        conn.rollback()
    finally:
        conn.close()

def get_all_rows():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM rda_data")
    rows = cursor.fetchall()
    conn.close()
    return rows
