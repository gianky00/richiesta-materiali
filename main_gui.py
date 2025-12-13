import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import sqlite3
import os
import sys
import subprocess
from src.database import get_connection, init_db
from src.utils import logger

class RDAApp:
    def __init__(self, root):
        # Ensure DB table exists to avoid crash
        init_db()

        self.root = root
        self.root.title("RDA Viewer")
        self.root.geometry("1200x600")

        # --- Styles ---
        style = ttk.Style()
        style.configure("Treeview", rowheight=25)

        # --- Filter Frame ---
        filter_frame = ttk.LabelFrame(root, text="Filtri", padding=10)
        filter_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(filter_frame, text="Cerca:").pack(side="left", padx=5)
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(filter_frame, textvariable=self.search_var)
        self.search_entry.pack(side="left", fill="x", expand=True, padx=5)
        self.search_entry.bind("<KeyRelease>", self.filter_data)

        ttk.Button(filter_frame, text="Ricarica Dati", command=self.load_data).pack(side="right", padx=5)

        # --- Treeview (Table) ---
        table_frame = ttk.Frame(root)
        table_frame.pack(fill="both", expand=True, padx=10, pady=5)

        columns = (
            "rda_number", "commessa", "desc_materiale", "unita", "qty",
            "apf", "data_rda", "data_consegna", "alert", "richiedente"
        )
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")

        # Define Headings (Updated to match User's terms)
        self.tree.heading("rda_number", text="N° RDA")
        self.tree.heading("commessa", text="Articolo") # Was Commessa
        self.tree.heading("desc_materiale", text="Descrizione Materiale")
        self.tree.heading("unita", text="UM")
        self.tree.heading("qty", text="Quantita Richiesta")
        self.tree.heading("apf", text="APF")
        self.tree.heading("data_rda", text="Data RDA")
        self.tree.heading("data_consegna", text="Data di Consegna")
        self.tree.heading("alert", text="N° Alert")
        self.tree.heading("richiedente", text="Richiedente")

        # Define Column Widths
        self.tree.column("rda_number", width=100)
        self.tree.column("commessa", width=100)
        self.tree.column("desc_materiale", width=300)
        self.tree.column("unita", width=50)
        self.tree.column("qty", width=50)
        self.tree.column("apf", width=50)
        self.tree.column("data_rda", width=100)
        self.tree.column("data_consegna", width=100)
        self.tree.column("alert", width=50)
        self.tree.column("richiedente", width=150)

        # Scrollbar
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)

        # Context Menu
        self.context_menu = tk.Menu(self.tree, tearoff=0)
        self.context_menu.add_command(label="Apri PDF", command=self.open_pdf)
        self.tree.bind("<Button-3>", self.show_context_menu)

        # Load Data
        self.all_data = []
        self.load_data()

    def load_data(self):
        try:
            conn = get_connection()
            cursor = conn.cursor()
            # Select columns matching the treeview, plus pdf_path hidden
            cursor.execute("SELECT rda_number, commessa, descrizione_materiale, unita_misura, quantita, apf, data_rda, data_consegna, alert_level, richiedente, pdf_path FROM rda_data")
            self.all_data = cursor.fetchall()
            conn.close()
            self.refresh_table(self.all_data)
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile caricare i dati: {e}")

    def refresh_table(self, data):
        self.tree.delete(*self.tree.get_children())
        for row in data:
            # Row is (rda, commessa, desc, um, qty, apf, date, deliv, alert, req, PATH)
            values = list(row)[:10] # Exclude path from visible columns
            # Store path in tag or mapping? We can use the item ID map
            item_id = self.tree.insert("", "end", values=values)
            # We need to store the PDF path somewhere accessible.
            # We can use a dictionary mapping item_id -> path
            if not hasattr(self, 'path_map'): self.path_map = {}
            self.path_map[item_id] = row[10] # The last column selected

    def filter_data(self, event=None):
        query = self.search_var.get().lower()
        if not query:
            self.refresh_table(self.all_data)
            return

        filtered = []
        for row in self.all_data:
            # Check if query is in any of the visible fields
            match = False
            for field in row[:10]:
                if query in str(field).lower():
                    match = True
                    break
            if match:
                filtered.append(row)
        self.refresh_table(filtered)

    def show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    def open_pdf(self):
        selected_item = self.tree.selection()
        if not selected_item: return

        item_id = selected_item[0]
        pdf_path = self.path_map.get(item_id)

        if pdf_path and os.path.exists(pdf_path):
            try:
                if sys.platform == 'win32':
                    os.startfile(pdf_path)
                else:
                    subprocess.call(['xdg-open', pdf_path])
            except Exception as e:
                messagebox.showerror("Errore", f"Impossibile aprire il file: {e}")
        else:
            messagebox.showwarning("Attenzione", f"File PDF non trovato:\n{pdf_path}")

if __name__ == "__main__":
    root = tk.Tk()
    try:
        # Optional: Set icon if available
        # root.iconbitmap("icon.ico")
        pass
    except:
        pass

    app = RDAApp(root)
    root.mainloop()
