"""
RDA Viewer - Applicazione per la gestione delle Richieste di Acquisto
Versione completamente rifattorizzata con interfaccia moderna e reattiva
"""

import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import os
import sys
import subprocess
import threading
from datetime import datetime, timedelta
from collections import Counter
import re
import traceback

# Percorsi configurazione
# Risaliamo da src/main_gui.py
SCRIPT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

# Moduli Licenza e Aggiornamento
from src.core import app_updater
from src.core import license_updater
from src.core import license_validator
from src.core import config_manager
from src.utils import config

# Setup logging
def setup_logging():
    """Configura il logging per catturare errori in file"""
    # Usa AppData per i log
    data_dir = config_manager.get_data_path()
    log_dir = os.path.join(data_dir, "Logs")

    if not os.path.exists(log_dir):
        try:
            os.makedirs(log_dir)
        except:
            return 

    log_file = os.path.join(log_dir, "app_error.log")
    
    # Redirezione stderr su file
    sys.stderr = open(log_file, "a")
    
    def exception_handler(exc_type, exc_value, exc_traceback):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        error_msg = f"\n[{timestamp}] UNHANDLED EXCEPTION:\n"
        error_msg += "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
        error_msg += "-"*80 + "\n"
        
        # Scrivi su file
        sys.stderr.write(error_msg)
        sys.stderr.flush()
        
        # Mostra messaggio errore se possibile
        try:
            import tkinter.messagebox
            root = tk.Tk()
            root.withdraw()
            tkinter.messagebox.showerror("Errore Critico", f"Si è verificato un errore imprevisto.\nControlla {log_file}\n\n{exc_value}")
            root.destroy()
        except:
            pass
            
    sys.excepthook = exception_handler

setup_logging()

# Prova prima il percorso di rete, poi locale
NETWORK_BASE_PATH = r"\\192.168.11.251\Condivisa\RICHIESTE MATERIALI"
LOCAL_BASE_PATH = SCRIPT_DIR

def get_paths():
    """Determina i percorsi corretti (rete o locale)"""
    if os.path.exists(NETWORK_BASE_PATH):
        base = NETWORK_BASE_PATH
    else:
        base = LOCAL_BASE_PATH
    
    return {
        'database': os.path.join(base, "DATABASE", "database_RDA.db"),
        'excel': os.path.join(base, "DATABASE", "database_RDA.xlsm"),
        'pdf_folder': os.path.join(base, "RDA_PDF")
    }

PATHS = get_paths()


def format_number(value):
    """
    Formatta i numeri: mostra come intero se è un numero intero (2.0 -> 2)
    Mantiene stringhe testuali come '25/039' invariate
    """
    if value is None or value == "":
        return ""
    
    # Se è già una stringa con caratteri non numerici (es. 25/039), la lascia invariata
    str_val = str(value)
    if re.search(r'[^\d.,\-\s]', str_val):
        return str_val
    
    # Prova a convertire in numero
    try:
        num = float(str(value).replace(',', '.'))
        # Se è un intero, mostralo senza decimali
        if num == int(num):
            return str(int(num))
        else:
            # Altrimenti mostra con virgola come separatore decimale italiano
            return str(num).replace('.', ',')
    except (ValueError, TypeError):
        return str_val


def format_date(value):
    """Formatta le date in formato italiano dd/mm/yyyy"""
    if not value:
        return ""
    if isinstance(value, str):
        return value
    try:
        if isinstance(value, datetime):
            return value.strftime("%d/%m/%Y")
    except:
        pass
    return str(value)


class ModernStyle:
    """Stili moderni per l'interfaccia - TEMA CHIARO"""
    
    # Colori tema chiaro
    BG_PRIMARY = "#FFFFFF"
    BG_SECONDARY = "#F5F7FA"
    BG_TERTIARY = "#E8ECF0"
    
    ACCENT_PRIMARY = "#2563EB"      # Blu
    ACCENT_SUCCESS = "#059669"       # Verde
    ACCENT_WARNING = "#D97706"       # Arancio
    ACCENT_DANGER = "#DC2626"        # Rosso
    
    TEXT_PRIMARY = "#1F2937"
    TEXT_SECONDARY = "#6B7280"
    TEXT_MUTED = "#9CA3AF"
    
    BORDER_COLOR = "#E5E7EB"
    
    # Font
    FONT_FAMILY = "Segoe UI"
    FONT_SIZE_NORMAL = 10
    FONT_SIZE_LARGE = 12
    FONT_SIZE_XLARGE = 16
    FONT_SIZE_TITLE = 24
    
    @classmethod
    def apply(cls, root):
        """Applica gli stili ttk"""
        style = ttk.Style()
        
        # Tema base
        style.theme_use('clam')
        
        # Frame
        style.configure("TFrame", background=cls.BG_PRIMARY)
        style.configure("Card.TFrame", background=cls.BG_PRIMARY, relief="flat")
        style.configure("Secondary.TFrame", background=cls.BG_SECONDARY)
        
        # Label
        style.configure("TLabel", 
                       background=cls.BG_PRIMARY, 
                       foreground=cls.TEXT_PRIMARY,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_NORMAL))
        
        style.configure("Title.TLabel",
                       background=cls.BG_PRIMARY,
                       foreground=cls.TEXT_PRIMARY,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_TITLE, "bold"))
        
        style.configure("Subtitle.TLabel",
                       background=cls.BG_PRIMARY,
                       foreground=cls.TEXT_SECONDARY,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_LARGE))
        
        style.configure("Stat.TLabel",
                       background=cls.BG_PRIMARY,
                       foreground=cls.TEXT_PRIMARY,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_XLARGE, "bold"))
        
        style.configure("StatTitle.TLabel",
                       background=cls.BG_PRIMARY,
                       foreground=cls.TEXT_SECONDARY,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_NORMAL))
        
        # Entry
        style.configure("TEntry",
                       fieldbackground=cls.BG_PRIMARY,
                       foreground=cls.TEXT_PRIMARY,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_NORMAL))
        
        # Button
        style.configure("TButton",
                       background=cls.ACCENT_PRIMARY,
                       foreground="white",
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_NORMAL),
                       padding=(15, 8))
        
        style.configure("Accent.TButton",
                       background=cls.ACCENT_PRIMARY,
                       foreground="white")
        
        style.configure("Success.TButton",
                       background=cls.ACCENT_SUCCESS,
                       foreground="white")
        
        # Notebook (Tabs)
        style.configure("TNotebook",
                       background=cls.BG_SECONDARY,
                       tabmargins=[0, 5, 0, 0])
        
        style.configure("TNotebook.Tab",
                       background=cls.BG_TERTIARY,
                       foreground=cls.TEXT_PRIMARY,
                       padding=[20, 10],
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_NORMAL))
        
        style.map("TNotebook.Tab",
                 background=[("selected", cls.BG_PRIMARY)],
                 foreground=[("selected", cls.ACCENT_PRIMARY)])
        
        # Treeview
        style.configure("Treeview",
                       background=cls.BG_PRIMARY,
                       foreground=cls.TEXT_PRIMARY,
                       fieldbackground=cls.BG_PRIMARY,
                       rowheight=30,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_NORMAL))
        
        style.configure("Treeview.Heading",
                       background=cls.BG_SECONDARY,
                       foreground=cls.TEXT_PRIMARY,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_NORMAL, "bold"),
                       padding=10)
        
        style.map("Treeview",
                 background=[("selected", cls.ACCENT_PRIMARY)],
                 foreground=[("selected", "white")])
        
        # LabelFrame
        style.configure("TLabelframe",
                       background=cls.BG_PRIMARY,
                       foreground=cls.TEXT_PRIMARY)
        
        style.configure("TLabelframe.Label",
                       background=cls.BG_PRIMARY,
                       foreground=cls.TEXT_PRIMARY,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_NORMAL, "bold"))
        
        # Progressbar
        style.configure("TProgressbar",
                       background=cls.ACCENT_PRIMARY,
                       troughcolor=cls.BG_TERTIARY)
        
        # Scrollbar
        style.configure("TScrollbar",
                       background=cls.BG_TERTIARY,
                       troughcolor=cls.BG_PRIMARY)


class DatabaseManager:
    """Gestisce connessione e operazioni sul database SQLite"""
    
    def __init__(self, db_path):
        self.db_path = db_path
        self._connection = None
    
    def get_connection(self):
        """Ottiene connessione al database"""
        try:
            conn = sqlite3.connect(self.db_path, check_same_thread=False)
            conn.row_factory = sqlite3.Row
            return conn
        except Exception as e:
            raise Exception(f"Errore connessione database: {e}")
    
    def fetch_all_data(self):
        """Recupera tutti i dati dal database"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT rda_number, commessa, descrizione_materiale, unita_misura, 
                   quantita, apf, data_rda, data_consegna, alert_level, 
                   richiedente, pdf_path 
            FROM rda_data
            ORDER BY data_rda DESC
        """)
        rows = cursor.fetchall()
        conn.close()
        return rows
    
    def get_statistics(self):
        """Calcola statistiche dai dati"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        stats = {}
        
        # Totale RDA
        cursor.execute("SELECT COUNT(DISTINCT rda_number) FROM rda_data")
        stats['total_rda'] = cursor.fetchone()[0] or 0
        
        # Totale righe/articoli
        cursor.execute("SELECT COUNT(*) FROM rda_data")
        stats['total_rows'] = cursor.fetchone()[0] or 0
        
        # RDA scadute (alert_level > 0)
        cursor.execute("SELECT COUNT(DISTINCT rda_number) FROM rda_data WHERE alert_level > 0")
        stats['overdue_rda'] = cursor.fetchone()[0] or 0
        
        # RDA per richiedente
        cursor.execute("""
            SELECT richiedente, COUNT(DISTINCT rda_number) as cnt 
            FROM rda_data 
            WHERE richiedente IS NOT NULL AND richiedente != ''
            GROUP BY richiedente 
            ORDER BY cnt DESC
        """)
        stats['by_requester'] = cursor.fetchall()
        
        # RDA per mese (ultimi 12 mesi)
        cursor.execute("""
            SELECT substr(data_rda, 4, 7) as mese, COUNT(DISTINCT rda_number) as cnt
            FROM rda_data
            WHERE data_rda IS NOT NULL AND data_rda != ''
            GROUP BY mese
            ORDER BY substr(data_rda, 7, 4) DESC, substr(data_rda, 4, 2) DESC
            LIMIT 12
        """)
        stats['by_month'] = cursor.fetchall()
        
        # APF distribution
        cursor.execute("""
            SELECT apf, COUNT(*) as cnt 
            FROM rda_data 
            WHERE apf IS NOT NULL AND apf != ''
            GROUP BY apf
        """)
        stats['by_apf'] = cursor.fetchall()
        
        # Alert level distribution
        cursor.execute("""
            SELECT alert_level, COUNT(*) as cnt 
            FROM rda_data 
            GROUP BY alert_level
            ORDER BY alert_level
        """)
        stats['by_alert'] = cursor.fetchall()
        
        # Top articoli richiesti
        cursor.execute("""
            SELECT descrizione_materiale, SUM(quantita) as tot_qty, COUNT(*) as cnt
            FROM rda_data
            WHERE descrizione_materiale IS NOT NULL AND descrizione_materiale != ''
            GROUP BY descrizione_materiale
            ORDER BY cnt DESC
            LIMIT 10
        """)
        stats['top_materials'] = cursor.fetchall()
        
        conn.close()
        return stats


class RDAViewerApp:
    """Applicazione principale RDA Viewer"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("📋 RDA Viewer - Gestione Richieste di Acquisto")
        self.root.geometry("1400x800")
        self.root.minsize(1200, 600)
        
        # Applica stili
        ModernStyle.apply(root)
        self.root.configure(bg=ModernStyle.BG_SECONDARY)
        
        # Database manager
        self.db = DatabaseManager(PATHS['database'])
        
        # Dati in memoria per performance
        self.all_data = []
        self.filtered_data = []
        self.path_map = {}
        
        # Variabili di stato
        self.loading = False
        self.search_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Avvio in corso...")
        
        # Costruisci interfaccia
        self._build_ui()
        
        # Carica dati automaticamente all'avvio
        self.root.after(100, self._initial_load)
    
    def _build_ui(self):
        """Costruisce l'interfaccia utente"""
        # Container principale
        main_container = ttk.Frame(self.root, style="Secondary.TFrame")
        main_container.pack(fill="both", expand=True)
        
        # Header
        self._build_header(main_container)
        
        # Notebook con tabs
        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Tab 1: Dati RDA
        self.data_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.data_tab, text="  📊 Dati RDA  ")
        self._build_data_tab()
        
        # Tab 2: Dashboard
        self.dashboard_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.dashboard_tab, text="  🏠 Dashboard  ")
        self._build_dashboard_tab()
        
        # Tab 3: Statistiche
        self.stats_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.stats_tab, text="  📈 Statistiche  ")
        self._build_stats_tab()
        
        # Tab 4: RDA Scadute
        self.overdue_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.overdue_tab, text="  ⚠️ RDA Scadute  ")
        self._build_overdue_tab()
        
        # Tab 5: Ricerca Avanzata
        self.search_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.search_tab, text="  🔍 Ricerca Avanzata  ")
        self._build_search_tab()

        # Tab 6: Configurazione
        self.config_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.config_tab, text="  ⚙️ Configurazione  ")
        self._build_config_tab()
        
        # Status bar
        self._build_statusbar(main_container)
    
    def _build_header(self, parent):
        """Costruisce l'header dell'applicazione"""
        header_frame = ttk.Frame(parent, style="TFrame")
        header_frame.pack(fill="x", padx=10, pady=10)
        
        # Titolo
        title_frame = ttk.Frame(header_frame)
        title_frame.pack(side="left")
        
        ttk.Label(title_frame, text="RDA Viewer", style="Title.TLabel").pack(anchor="w")
        ttk.Label(title_frame, text="Sistema di Gestione Richieste di Acquisto", 
                 style="Subtitle.TLabel").pack(anchor="w")
        
        # Info connessione
        info_frame = ttk.Frame(header_frame)
        info_frame.pack(side="right")
        
        db_status = "🟢 Connesso" if os.path.exists(PATHS['database']) else "🔴 Non connesso"
        ttk.Label(info_frame, text=f"Database: {db_status}", 
                 style="TLabel").pack(anchor="e")
        
        # Mostra percorso abbreviato
        short_path = PATHS['database']
        if len(short_path) > 50:
            short_path = "..." + short_path[-47:]
        ttk.Label(info_frame, text=short_path, 
                 style="TLabel", foreground=ModernStyle.TEXT_MUTED).pack(anchor="e")
    
    def _build_data_tab(self):
        """Costruisce il tab principale con i dati"""
        # Frame di ricerca
        search_frame = ttk.Frame(self.data_tab)
        search_frame.pack(fill="x", padx=10, pady=10)
        
        ttk.Label(search_frame, text="🔍 Cerca:").pack(side="left", padx=(0, 10))
        
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, 
                                width=50, font=(ModernStyle.FONT_FAMILY, 12))
        search_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        search_entry.bind("<KeyRelease>", self._on_search)
        
        # Info risultati
        self.results_label = ttk.Label(search_frame, text="", style="TLabel")
        self.results_label.pack(side="right", padx=10)
        
        # Frame tabella
        table_frame = ttk.Frame(self.data_tab)
        table_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Colonne
        columns = (
            "rda_number", "commessa", "desc_materiale", "unita", "qty",
            "apf", "data_rda", "data_consegna", "alert", "richiedente"
        )
        
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        
        # Intestazioni e larghezze
        headers = {
            "rda_number": ("N° RDA", 100),
            "commessa": ("Articolo", 120),
            "desc_materiale": ("Descrizione Materiale", 350),
            "unita": ("UM", 50),
            "qty": ("Quantità", 80),
            "apf": ("APF", 60),
            "data_rda": ("Data RDA", 100),
            "data_consegna": ("Data Consegna", 100),
            "alert": ("Alert", 60),
            "richiedente": ("Richiedente", 150)
        }
        
        for col, (text, width) in headers.items():
            self.tree.heading(col, text=text, command=lambda c=col: self._sort_column(c))
            self.tree.column(col, width=width, minwidth=50)
        
        # Scrollbars
        v_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        h_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        # Grid layout
        self.tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        # Menu contestuale
        self.context_menu = tk.Menu(self.tree, tearoff=0)
        self.context_menu.add_command(label="📄 Apri PDF", command=self._open_pdf)
        self.context_menu.add_command(label="📋 Copia riga", command=self._copy_row)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="🔍 Filtra per questo RDA", command=self._filter_by_rda)
        
        self.tree.bind("<Button-3>", self._show_context_menu)
        self.tree.bind("<Double-1>", lambda e: self._open_pdf())
    
    def _build_dashboard_tab(self):
        """Costruisce il tab dashboard con panoramica"""
        # Container con scroll
        canvas = tk.Canvas(self.dashboard_tab, bg=ModernStyle.BG_PRIMARY, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.dashboard_tab, orient="vertical", command=canvas.yview)
        self.dashboard_frame = ttk.Frame(canvas)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        
        canvas_window = canvas.create_window((0, 0), window=self.dashboard_frame, anchor="nw")
        
        def configure_scroll(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        def configure_width(event):
            canvas.itemconfig(canvas_window, width=event.width)
        
        self.dashboard_frame.bind("<Configure>", configure_scroll)
        canvas.bind("<Configure>", configure_width)
        
        # Titolo
        ttk.Label(self.dashboard_frame, text="Dashboard", 
                 style="Title.TLabel").pack(anchor="w", padx=20, pady=(20, 5))
        ttk.Label(self.dashboard_frame, text="Panoramica delle Richieste di Acquisto", 
                 style="Subtitle.TLabel").pack(anchor="w", padx=20, pady=(0, 20))
        
        # Frame per le card statistiche
        self.stats_cards_frame = ttk.Frame(self.dashboard_frame)
        self.stats_cards_frame.pack(fill="x", padx=20, pady=10)
        
        # Placeholder per le card (verranno popolate dopo il caricamento)
        self.dashboard_cards = {}
    
    def _build_stats_tab(self):
        """Costruisce il tab statistiche dettagliate"""
        # Canvas scrollabile
        canvas = tk.Canvas(self.stats_tab, bg=ModernStyle.BG_PRIMARY, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.stats_tab, orient="vertical", command=canvas.yview)
        self.stats_frame = ttk.Frame(canvas)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        
        canvas_window = canvas.create_window((0, 0), window=self.stats_frame, anchor="nw")
        
        def configure_scroll(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        def configure_width(event):
            canvas.itemconfig(canvas_window, width=event.width)
        
        self.stats_frame.bind("<Configure>", configure_scroll)
        canvas.bind("<Configure>", configure_width)
        
        # Titolo
        ttk.Label(self.stats_frame, text="Statistiche", 
                 style="Title.TLabel").pack(anchor="w", padx=20, pady=(20, 5))
        ttk.Label(self.stats_frame, text="Analisi dettagliata dei dati", 
                 style="Subtitle.TLabel").pack(anchor="w", padx=20, pady=(0, 20))
        
        # Container per grafici testuali
        self.stats_content_frame = ttk.Frame(self.stats_frame)
        self.stats_content_frame.pack(fill="both", expand=True, padx=20, pady=10)
    
    def _build_overdue_tab(self):
        """Costruisce il tab per le RDA scadute"""
        # Header
        header = ttk.Frame(self.overdue_tab)
        header.pack(fill="x", padx=10, pady=10)
        
        ttk.Label(header, text="⚠️ RDA Scadute", 
                 style="Title.TLabel").pack(side="left")
        
        # Tabella RDA scadute
        table_frame = ttk.Frame(self.overdue_tab)
        table_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        columns = ("rda_number", "desc_materiale", "data_rda", "alert", "richiedente", "giorni")
        
        self.overdue_tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        
        headers = {
            "rda_number": ("N° RDA", 100),
            "desc_materiale": ("Descrizione", 400),
            "data_rda": ("Data RDA", 100),
            "alert": ("Livello Alert", 100),
            "richiedente": ("Richiedente", 150),
            "giorni": ("Giorni Trascorsi", 120)
        }
        
        for col, (text, width) in headers.items():
            self.overdue_tree.heading(col, text=text)
            self.overdue_tree.column(col, width=width)
        
        # Tag per colorazione
        self.overdue_tree.tag_configure("high", background="#FEE2E2", foreground="#991B1B")
        self.overdue_tree.tag_configure("medium", background="#FEF3C7", foreground="#92400E")
        self.overdue_tree.tag_configure("low", background="#DBEAFE", foreground="#1E40AF")
        
        v_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.overdue_tree.yview)
        self.overdue_tree.configure(yscrollcommand=v_scroll.set)
        
        self.overdue_tree.pack(side="left", fill="both", expand=True)
        v_scroll.pack(side="right", fill="y")
    
    def _build_search_tab(self):
        """Costruisce il tab ricerca avanzata"""
        # Frame filtri
        filters_frame = ttk.LabelFrame(self.search_tab, text="Filtri Avanzati", padding=15)
        filters_frame.pack(fill="x", padx=10, pady=10)
        
        # Grid per filtri
        row = 0
        
        # Filtro RDA
        ttk.Label(filters_frame, text="N° RDA:").grid(row=row, column=0, padx=5, pady=5, sticky="e")
        self.adv_rda_var = tk.StringVar()
        ttk.Entry(filters_frame, textvariable=self.adv_rda_var, width=20).grid(
            row=row, column=1, padx=5, pady=5, sticky="w")
        
        # Filtro Richiedente
        ttk.Label(filters_frame, text="Richiedente:").grid(row=row, column=2, padx=5, pady=5, sticky="e")
        self.adv_requester_var = tk.StringVar()
        self.adv_requester_combo = ttk.Combobox(filters_frame, textvariable=self.adv_requester_var, width=25)
        self.adv_requester_combo.grid(row=row, column=3, padx=5, pady=5, sticky="w")
        
        row += 1
        
        # Filtro Data Da
        ttk.Label(filters_frame, text="Data Da:").grid(row=row, column=0, padx=5, pady=5, sticky="e")
        self.adv_date_from_var = tk.StringVar()
        ttk.Entry(filters_frame, textvariable=self.adv_date_from_var, width=20).grid(
            row=row, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(filters_frame, text="(gg/mm/aaaa)").grid(row=row, column=1, padx=(150, 0), pady=5, sticky="w")
        
        # Filtro Data A
        ttk.Label(filters_frame, text="Data A:").grid(row=row, column=2, padx=5, pady=5, sticky="e")
        self.adv_date_to_var = tk.StringVar()
        ttk.Entry(filters_frame, textvariable=self.adv_date_to_var, width=20).grid(
            row=row, column=3, padx=5, pady=5, sticky="w")
        
        row += 1
        
        # Filtro APF
        ttk.Label(filters_frame, text="APF:").grid(row=row, column=0, padx=5, pady=5, sticky="e")
        self.adv_apf_var = tk.StringVar()
        self.adv_apf_combo = ttk.Combobox(filters_frame, textvariable=self.adv_apf_var, width=20)
        self.adv_apf_combo.grid(row=row, column=1, padx=5, pady=5, sticky="w")
        
        # Solo scadute
        self.adv_overdue_var = tk.BooleanVar()
        ttk.Checkbutton(filters_frame, text="Solo RDA scadute", 
                       variable=self.adv_overdue_var).grid(row=row, column=2, columnspan=2, padx=5, pady=5, sticky="w")
        
        row += 1
        
        # Pulsanti
        btn_frame = ttk.Frame(filters_frame)
        btn_frame.grid(row=row, column=0, columnspan=4, pady=15)
        
        ttk.Button(btn_frame, text="🔍 Cerca", command=self._advanced_search).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="🔄 Reset", command=self._reset_advanced_search).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="📥 Esporta CSV", command=self._export_csv).pack(side="left", padx=5)
        
        # Risultati
        results_frame = ttk.Frame(self.search_tab)
        results_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        columns = ("rda_number", "commessa", "desc_materiale", "qty", "data_rda", "richiedente")
        
        self.adv_tree = ttk.Treeview(results_frame, columns=columns, show="headings")
        
        for col in columns:
            self.adv_tree.heading(col, text=col.replace("_", " ").title())
            self.adv_tree.column(col, width=150)
        
        v_scroll = ttk.Scrollbar(results_frame, orient="vertical", command=self.adv_tree.yview)
        self.adv_tree.configure(yscrollcommand=v_scroll.set)
        
        self.adv_tree.pack(side="left", fill="both", expand=True)
        v_scroll.pack(side="right", fill="y")
    
    def _build_statusbar(self, parent):
        """Costruisce la barra di stato"""
        status_frame = ttk.Frame(parent, style="Secondary.TFrame")
        status_frame.pack(fill="x", side="bottom")
        
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var, 
                                      style="TLabel", padding=(10, 5))
        self.status_label.pack(side="left")
        
        # Progress bar (nascosta di default)
        self.progress = ttk.Progressbar(status_frame, mode="indeterminate", length=150)
        self.progress.pack(side="right", padx=10, pady=5)
        self.progress.pack_forget()
        
        # Info aggiornamento
        self.update_label = ttk.Label(status_frame, text="", style="TLabel", padding=(10, 5))
        self.update_label.pack(side="right")
    
    def _initial_load(self):
        """Caricamento iniziale dei dati"""
        self._sync_and_load()
    
    def _sync_and_load(self):
        """Sincronizza con Excel e carica i dati"""
        def sync_task():
            self.root.after(0, lambda: self._set_loading(True, "Sincronizzazione in corso..."))
            
            try:
                # Tenta sincronizzazione con Excel se disponibile
                if os.path.exists(PATHS['excel']):
                    try:
                        self._sync_excel_to_db()
                        self.root.after(0, lambda: self.status_var.set("Sincronizzato con Excel"))
                    except Exception as e:
                        self.root.after(0, lambda: self.status_var.set(f"Sync Excel fallita: {str(e)[:30]}..."))
                
                # Carica i dati dal database
                self.root.after(0, lambda: self.status_var.set("Caricamento dati..."))
                self.all_data = self.db.fetch_all_data()
                self.filtered_data = list(self.all_data)
                
                # Aggiorna interfaccia nel thread principale
                self.root.after(0, self._update_ui_after_load)
                
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Errore", f"Errore caricamento: {e}"))
            finally:
                self.root.after(0, lambda: self._set_loading(False, "Pronto"))
        
        # Esegui in thread separato
        threading.Thread(target=sync_task, daemon=True).start()
    
    def _sync_excel_to_db(self):
        """Sincronizza i dati da Excel al database SQLite"""
        import pythoncom
        
        pythoncom.CoInitialize()
        try:
            import win32com.client
            
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            try:
                workbook = excel.Workbooks.Open(PATHS['excel'], ReadOnly=True)
                sheet = workbook.ActiveSheet
                
                # Leggi ultima riga
                last_row = sheet.Cells(sheet.Rows.Count, "A").End(-4162).Row
                
                if last_row < 2:
                    workbook.Close(SaveChanges=False)
                    return
                
                # Leggi tutti i dati
                raw_values = sheet.Range(f"A2:L{last_row}").Value
                formulas_col8 = sheet.Range(f"H2:H{last_row}").Formula
                
                workbook.Close(SaveChanges=False)
                
                # Normalizza struttura dati
                if not isinstance(raw_values, tuple):
                    raw_values = ((raw_values,),)
                elif len(raw_values) > 0 and not isinstance(raw_values[0], (tuple, list)):
                    raw_values = (raw_values,)
                
                if isinstance(formulas_col8, str):
                    formulas_col8 = ((formulas_col8,),)
                elif isinstance(formulas_col8, tuple) and len(formulas_col8) > 0 and not isinstance(formulas_col8[0], (tuple, list)):
                    formulas_col8 = tuple((f,) for f in formulas_col8) if isinstance(formulas_col8[0], str) else (formulas_col8,)
                
                # Converti in formato database
                data_rows = []
                for i, row_val in enumerate(raw_values):
                    if row_val is None or all(v is None for v in row_val):
                        continue
                    
                    # Estrai path PDF dalla formula HYPERLINK
                    pdf_path = ""
                    try:
                        formula = formulas_col8[i][0] if isinstance(formulas_col8[i], tuple) else formulas_col8[i]
                        match = re.search(r'HYPERLINK\("([^"]+)"', str(formula))
                        if match:
                            pdf_path = match.group(1)
                    except:
                        pass
                    
                    def fmt_date(d):
                        if isinstance(d, datetime):
                            return d.strftime("%d/%m/%Y")
                        return str(d) if d else ""
                    
                    item = (
                        str(row_val[0]) if row_val[0] else "",
                        str(row_val[1]) if row_val[1] else "",
                        str(row_val[2]) if row_val[2] else "",
                        str(row_val[3]) if row_val[3] else "",
                        str(row_val[4]) if row_val[4] else "",
                        row_val[5] if row_val[5] else 0.0,
                        str(row_val[6]) if row_val[6] else "",
                        pdf_path,
                        fmt_date(row_val[8]),
                        fmt_date(row_val[9]),
                        int(row_val[10]) if row_val[10] else 0,
                        str(row_val[11]) if row_val[11] else ""
                    )
                    data_rows.append(item)
                
                # Salva nel database
                if data_rows:
                    conn = self.db.get_connection()
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM rda_data")
                    cursor.executemany("""
                        INSERT INTO rda_data (
                            rda_number, commessa, descrizione_1, descrizione_materiale,
                            unita_misura, quantita, apf, pdf_path, data_rda,
                            data_consegna, alert_level, richiedente
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, data_rows)
                    conn.commit()
                    conn.close()
                
            finally:
                excel.Quit()
        finally:
            pythoncom.CoUninitialize()
    
    def _update_ui_after_load(self):
        """Aggiorna l'interfaccia dopo il caricamento dei dati"""
        # Aggiorna tabella principale
        self._refresh_table(self.filtered_data)
        
        # Aggiorna statistiche
        self._update_dashboard()
        self._update_stats()
        self._update_overdue()
        self._update_advanced_filters()
        
        # Info risultati
        self.results_label.config(text=f"{len(self.filtered_data)} risultati")
        
        # Timestamp aggiornamento
        self.update_label.config(text=f"Ultimo aggiornamento: {datetime.now().strftime('%H:%M:%S')}")
    
    def _set_loading(self, loading, message=""):
        """Imposta stato di caricamento"""
        self.loading = loading
        self.status_var.set(message)
        
        if loading:
            self.progress.pack(side="right", padx=10, pady=5)
            self.progress.start(10)
        else:
            self.progress.stop()
            self.progress.pack_forget()
    
    def _refresh_table(self, data):
        """Aggiorna la tabella con i dati"""
        self.tree.delete(*self.tree.get_children())
        self.path_map.clear()
        
        for row in data:
            # Formatta i valori
            values = [
                row[0],  # RDA Number
                format_number(row[1]),  # Articolo/Commessa - formato corretto
                row[2],  # Descrizione
                row[3],  # UM
                format_number(row[4]),  # Quantità - formato corretto
                row[5],  # APF
                row[6],  # Data RDA
                row[7],  # Data Consegna
                format_number(row[8]) if row[8] else "",  # Alert
                row[9]   # Richiedente
            ]
            
            item_id = self.tree.insert("", "end", values=values)
            self.path_map[item_id] = row[10] if len(row) > 10 else ""
    
    def _on_search(self, event=None):
        """Gestisce la ricerca in tempo reale"""
        query = self.search_var.get().lower().strip()
        
        if not query:
            self.filtered_data = list(self.all_data)
        else:
            self.filtered_data = []
            for row in self.all_data:
                for field in row[:10]:
                    if query in str(field).lower():
                        self.filtered_data.append(row)
                        break
        
        self._refresh_table(self.filtered_data)
        self.results_label.config(text=f"{len(self.filtered_data)} risultati")
    
    def _sort_column(self, col):
        """Ordina la tabella per colonna"""
        # Mappa colonne a indici
        col_map = {
            "rda_number": 0, "commessa": 1, "desc_materiale": 2, "unita": 3,
            "qty": 4, "apf": 5, "data_rda": 6, "data_consegna": 7,
            "alert": 8, "richiedente": 9
        }
        
        idx = col_map.get(col, 0)
        reverse = getattr(self, f'_sort_reverse_{col}', False)
        
        def sort_key(row):
            val = row[idx]
            if val is None:
                return ""
            # Prova conversione numerica
            try:
                return float(str(val).replace(',', '.'))
            except:
                return str(val).lower()
        
        self.filtered_data.sort(key=sort_key, reverse=reverse)
        setattr(self, f'_sort_reverse_{col}', not reverse)
        
        self._refresh_table(self.filtered_data)
    
    def _show_context_menu(self, event):
        """Mostra menu contestuale"""
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)
    
    def _open_pdf(self):
        """Apre il PDF associato alla riga selezionata"""
        selected = self.tree.selection()
        if not selected:
            return
        
        pdf_path = self.path_map.get(selected[0], "")
        
        if pdf_path and os.path.exists(pdf_path):
            try:
                os.startfile(pdf_path)
            except Exception as e:
                messagebox.showerror("Errore", f"Impossibile aprire il file: {e}")
        else:
            messagebox.showwarning("Attenzione", f"File PDF non trovato:\n{pdf_path}")
    
    def _copy_row(self):
        """Copia la riga selezionata negli appunti"""
        selected = self.tree.selection()
        if not selected:
            return
        
        values = self.tree.item(selected[0])['values']
        text = "\t".join(str(v) for v in values)
        
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        
        self.status_var.set("Riga copiata negli appunti")
    
    def _filter_by_rda(self):
        """Filtra per il numero RDA selezionato"""
        selected = self.tree.selection()
        if not selected:
            return
        
        rda = self.tree.item(selected[0])['values'][0]
        self.search_var.set(str(rda))
        self._on_search()
    
    def _update_dashboard(self):
        """Aggiorna il tab dashboard"""
        # Pulisci frame esistente
        for widget in self.stats_cards_frame.winfo_children():
            widget.destroy()
        
        try:
            stats = self.db.get_statistics()
        except:
            return
        
        # Card statistiche principali
        cards_data = [
            ("📋", "RDA Totali", stats['total_rda'], ModernStyle.ACCENT_PRIMARY),
            ("📦", "Articoli Totali", stats['total_rows'], ModernStyle.ACCENT_SUCCESS),
            ("⚠️", "RDA Scadute", stats['overdue_rda'], ModernStyle.ACCENT_DANGER),
        ]
        
        for i, (icon, title, value, color) in enumerate(cards_data):
            card = ttk.Frame(self.stats_cards_frame, style="Card.TFrame")
            card.grid(row=0, column=i, padx=10, pady=10, sticky="nsew")
            
            # Bordo colorato
            card.configure(padding=20)
            
            ttk.Label(card, text=icon, font=(ModernStyle.FONT_FAMILY, 24)).pack()
            ttk.Label(card, text=str(value), style="Stat.TLabel").pack()
            ttk.Label(card, text=title, style="StatTitle.TLabel").pack()
        
        self.stats_cards_frame.grid_columnconfigure((0, 1, 2), weight=1)
        
        # Sezione richiedenti più attivi
        requesters_frame = ttk.LabelFrame(self.dashboard_frame, text="Top Richiedenti", padding=15)
        requesters_frame.pack(fill="x", padx=20, pady=10)
        
        if stats['by_requester']:
            for req, cnt in stats['by_requester'][:5]:
                row_frame = ttk.Frame(requesters_frame)
                row_frame.pack(fill="x", pady=2)
                ttk.Label(row_frame, text=f"👤 {req or 'N/D'}").pack(side="left")
                ttk.Label(row_frame, text=f"{cnt} RDA").pack(side="right")
    
    def _update_stats(self):
        """Aggiorna il tab statistiche"""
        # Pulisci contenuto esistente
        for widget in self.stats_content_frame.winfo_children():
            widget.destroy()
        
        try:
            stats = self.db.get_statistics()
        except:
            return
        
        # Distribuzione Alert
        alert_frame = ttk.LabelFrame(self.stats_content_frame, text="📊 Distribuzione Alert Level", padding=15)
        alert_frame.pack(fill="x", pady=10)
        
        if stats['by_alert']:
            max_cnt = max(cnt for _, cnt in stats['by_alert']) or 1
            for level, cnt in stats['by_alert']:
                row = ttk.Frame(alert_frame)
                row.pack(fill="x", pady=3)
                
                level_text = f"Level {level}" if level else "Nessun Alert"
                ttk.Label(row, text=level_text, width=15).pack(side="left")
                
                # Barra grafica testuale
                bar_width = int((cnt / max_cnt) * 30)
                bar = "█" * bar_width
                
                ttk.Label(row, text=bar, foreground=ModernStyle.ACCENT_PRIMARY).pack(side="left", padx=5)
                ttk.Label(row, text=str(cnt)).pack(side="left")
        
        # Top materiali
        materials_frame = ttk.LabelFrame(self.stats_content_frame, text="📦 Top 10 Materiali Richiesti", padding=15)
        materials_frame.pack(fill="x", pady=10)
        
        if stats['top_materials']:
            for i, (mat, qty, cnt) in enumerate(stats['top_materials'], 1):
                desc = mat[:60] + "..." if len(mat) > 60 else mat
                row = ttk.Frame(materials_frame)
                row.pack(fill="x", pady=2)
                ttk.Label(row, text=f"{i}.", width=3).pack(side="left")
                ttk.Label(row, text=desc).pack(side="left", padx=5)
                ttk.Label(row, text=f"(x{cnt})").pack(side="right")
        
        # APF Distribution
        apf_frame = ttk.LabelFrame(self.stats_content_frame, text="🏷️ Distribuzione APF", padding=15)
        apf_frame.pack(fill="x", pady=10)
        
        if stats['by_apf']:
            for apf, cnt in stats['by_apf'][:10]:
                row = ttk.Frame(apf_frame)
                row.pack(fill="x", pady=2)
                ttk.Label(row, text=f"APF: {apf or 'N/D'}").pack(side="left")
                ttk.Label(row, text=str(cnt)).pack(side="right")
    
    def _update_overdue(self):
        """Aggiorna il tab RDA scadute"""
        self.overdue_tree.delete(*self.overdue_tree.get_children())
        
        today = datetime.now()
        
        for row in self.all_data:
            alert_level = row[8] if len(row) > 8 else 0
            
            try:
                alert_level = int(alert_level) if alert_level else 0
            except:
                alert_level = 0
            
            if alert_level > 0:
                # Calcola giorni trascorsi dalla data RDA
                data_rda_str = row[6]
                giorni = 0
                try:
                    data_rda = datetime.strptime(data_rda_str, "%d/%m/%Y")
                    giorni = (today - data_rda).days
                except:
                    pass
                
                # Determina tag per colore
                if alert_level >= 10:
                    tag = "high"
                elif alert_level >= 5:
                    tag = "medium"
                else:
                    tag = "low"
                
                values = (row[0], row[2], row[6], alert_level, row[9], giorni)
                self.overdue_tree.insert("", "end", values=values, tags=(tag,))
    
    def _update_advanced_filters(self):
        """Aggiorna i filtri della ricerca avanzata"""
        # Raccogli valori unici
        requesters = set()
        apf_values = set()
        
        for row in self.all_data:
            if row[9]:
                requesters.add(str(row[9]))
            if row[5]:
                apf_values.add(str(row[5]))
        
        self.adv_requester_combo['values'] = [""] + sorted(requesters)
        self.adv_apf_combo['values'] = [""] + sorted(apf_values)
    
    def _advanced_search(self):
        """Esegue la ricerca avanzata"""
        results = []
        
        rda_filter = self.adv_rda_var.get().strip().lower()
        requester_filter = self.adv_requester_var.get().strip().lower()
        date_from = self.adv_date_from_var.get().strip()
        date_to = self.adv_date_to_var.get().strip()
        apf_filter = self.adv_apf_var.get().strip().lower()
        only_overdue = self.adv_overdue_var.get()
        
        # Parse date filters
        date_from_obj = None
        date_to_obj = None
        try:
            if date_from:
                date_from_obj = datetime.strptime(date_from, "%d/%m/%Y")
            if date_to:
                date_to_obj = datetime.strptime(date_to, "%d/%m/%Y")
        except ValueError:
            messagebox.showwarning("Attenzione", "Formato data non valido. Usa: gg/mm/aaaa")
            return
        
        for row in self.all_data:
            # Filtro RDA
            if rda_filter and rda_filter not in str(row[0]).lower():
                continue
            
            # Filtro Richiedente
            if requester_filter and requester_filter not in str(row[9]).lower():
                continue
            
            # Filtro APF
            if apf_filter and apf_filter not in str(row[5]).lower():
                continue
            
            # Filtro date
            try:
                row_date = datetime.strptime(row[6], "%d/%m/%Y") if row[6] else None
                if date_from_obj and row_date and row_date < date_from_obj:
                    continue
                if date_to_obj and row_date and row_date > date_to_obj:
                    continue
            except:
                pass
            
            # Filtro scadute
            if only_overdue:
                alert = int(row[8]) if row[8] else 0
                if alert <= 0:
                    continue
            
            results.append(row)
        
        # Aggiorna tabella risultati
        self.adv_tree.delete(*self.adv_tree.get_children())
        for row in results:
            values = (
                row[0], 
                format_number(row[1]), 
                row[2], 
                format_number(row[4]), 
                row[6], 
                row[9]
            )
            self.adv_tree.insert("", "end", values=values)
        
        self.status_var.set(f"Ricerca completata: {len(results)} risultati")
    
    def _reset_advanced_search(self):
        """Reset filtri ricerca avanzata"""
        self.adv_rda_var.set("")
        self.adv_requester_var.set("")
        self.adv_date_from_var.set("")
        self.adv_date_to_var.set("")
        self.adv_apf_var.set("")
        self.adv_overdue_var.set(False)
        
        self.adv_tree.delete(*self.adv_tree.get_children())
        self.status_var.set("Filtri reset")
    
    def _export_csv(self):
        """Esporta i dati in CSV"""
        from tkinter import filedialog
        
        filepath = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfilename=f"export_rda_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        
        if not filepath:
            return
        
        try:
            with open(filepath, 'w', encoding='utf-8-sig') as f:
                # Header
                f.write("N° RDA;Articolo;Descrizione;UM;Quantità;APF;Data RDA;Data Consegna;Alert;Richiedente\n")
                
                # Data
                data_to_export = self.filtered_data if self.notebook.index(self.notebook.select()) == 0 else self.all_data
                
                for row in data_to_export:
                    line = ";".join([
                        str(row[0]),
                        format_number(row[1]),
                        str(row[2]).replace(";", ","),
                        str(row[3]),
                        format_number(row[4]),
                        str(row[5]),
                        str(row[6]),
                        str(row[7]),
                        format_number(row[8]),
                        str(row[9])
                    ])
                    f.write(line + "\n")
            
            self.status_var.set(f"Esportato: {filepath}")
            messagebox.showinfo("Esportazione", f"File esportato con successo:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Errore", f"Errore durante l'esportazione: {e}")

    def _build_config_tab(self):
        """Costruisce il tab di configurazione"""
        from tkinter import filedialog

        # Container
        container = ttk.Frame(self.config_tab, padding=20)
        container.pack(fill="both", expand=True)

        ttk.Label(container, text="⚙️ Configurazione Applicazione", style="Title.TLabel").pack(anchor="w", pady=(0, 20))

        # Frame paths
        path_frame = ttk.LabelFrame(container, text="Percorsi File", padding=15)
        path_frame.pack(fill="x", pady=10)

        # Excel Path
        ttk.Label(path_frame, text="File Excel Database:").grid(row=0, column=0, sticky="w", pady=5)
        self.config_excel_var = tk.StringVar(value=config.EXCEL_DB_PATH)
        ttk.Entry(path_frame, textvariable=self.config_excel_var, width=70).grid(row=0, column=1, padx=10, pady=5)
        ttk.Button(path_frame, text="Sfoglia...", command=lambda: self._browse_file(self.config_excel_var, "Excel Files", "*.xlsm")).grid(row=0, column=2, pady=5)

        # PDF Folder
        ttk.Label(path_frame, text="Cartella PDF RDA:").grid(row=1, column=0, sticky="w", pady=5)
        self.config_pdf_var = tk.StringVar(value=config.PDF_SAVE_PATH)
        ttk.Entry(path_frame, textvariable=self.config_pdf_var, width=70).grid(row=1, column=1, padx=10, pady=5)
        ttk.Button(path_frame, text="Sfoglia...", command=lambda: self._browse_folder(self.config_pdf_var)).grid(row=1, column=2, pady=5)

        # Database Dir (Optional/Advanced)
        ttk.Label(path_frame, text="Cartella Database:").grid(row=2, column=0, sticky="w", pady=5)
        self.config_db_dir_var = tk.StringVar(value=config.DATABASE_DIR)
        ttk.Entry(path_frame, textvariable=self.config_db_dir_var, width=70).grid(row=2, column=1, padx=10, pady=5)
        ttk.Button(path_frame, text="Sfoglia...", command=lambda: self._browse_folder(self.config_db_dir_var)).grid(row=2, column=2, pady=5)

        # Buttons
        btn_frame = ttk.Frame(container, padding=20)
        btn_frame.pack(fill="x")
        ttk.Button(btn_frame, text="💾 Salva Configurazione", style="Success.TButton", command=self._save_configuration).pack(side="right")
        ttk.Button(btn_frame, text="🔄 Ripristina Default", command=self._reset_configuration).pack(side="right", padx=10)

        info_lbl = ttk.Label(container, text="Nota: Le modifiche richiedono il riavvio dell'applicazione.", foreground=ModernStyle.TEXT_MUTED)
        info_lbl.pack(pady=10)

    def _browse_file(self, var, file_desc, file_ext):
        from tkinter import filedialog
        path = filedialog.askopenfilename(filetypes=[(file_desc, file_ext), ("All Files", "*.*")])
        if path:
            var.set(path.replace('/', '\\'))

    def _browse_folder(self, var):
        from tkinter import filedialog
        path = filedialog.askdirectory()
        if path:
            var.set(path.replace('/', '\\'))

    def _save_configuration(self):
        new_config = {
            "excel_path": self.config_excel_var.get(),
            "pdf_folder": self.config_pdf_var.get(),
            "database_dir": self.config_db_dir_var.get()
        }

        if config_manager.save_config(new_config):
            messagebox.showinfo("Successo", "Configurazione salvata correttamente.\nRiavvia l'applicazione per applicare le modifiche.")
        else:
            messagebox.showerror("Errore", "Impossibile salvare la configurazione. Verifica i permessi.")

    def _reset_configuration(self):
        if messagebox.askyesno("Conferma", "Vuoi ripristinare la configurazione predefinita?"):
             defaults = config_manager.DEFAULT_CONFIG
             self.config_excel_var.set(defaults["excel_path"])
             self.config_pdf_var.set(defaults["pdf_folder"])
             self.config_db_dir_var.set(defaults["database_dir"])


def main():
    """Entry point dell'applicazione"""
    root = tk.Tk()
    
    # -------------------------------------------------------------------------
    # 1. CONTROLLO AGGIORNAMENTI APP
    # -------------------------------------------------------------------------
    try:
        app_updater.check_for_updates(silent=True)
    except Exception as e:
        print(f"[ERRORE] Check updates: {e}")

    # -------------------------------------------------------------------------
    # 2. AGGIORNAMENTO LICENZA
    # -------------------------------------------------------------------------
    try:
        license_updater.run_update()
    except Exception as e:
        print(f"[ERRORE] License update: {e}")

    # -------------------------------------------------------------------------
    # 3. VERIFICA LICENZA BLOCCANTE
    # -------------------------------------------------------------------------
    is_valid, message = license_validator.verify_license()

    if not is_valid:
        root.withdraw() # Nascondi finestra principale
        messagebox.showerror(
            "Licenza Non Valida",
            f"Impossibile avviare l'applicazione.\n\n{message}"
        )
        sys.exit(1)

    # -------------------------------------------------------------------------
    # AVVIO APP
    # -------------------------------------------------------------------------

    # Imposta icona se disponibile
    try:
        root.iconbitmap(default="")
    except:
        pass
    
    # Centra la finestra
    root.update_idletasks()
    width = 1400
    height = 800
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    app = RDAViewerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
