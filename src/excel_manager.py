"""
Modulo per la gestione del file Excel tramite COM automation
"""

import win32com.client
import logging
import re
from datetime import datetime
from .config import (
    EXCEL_DB_PATH, SHEET_PASSWORD, TABLE_NAME, 
    RDA_REFERENCE_COLUMN_LETTER, RDA_REFERENCE_COLUMN_NUMBER
)
from .utils import safe_str, safe_float, safe_int, format_date

logger = logging.getLogger("RDA_Bot")


class ExcelManager:
    """
    Gestisce le operazioni sul file Excel di registro RDA.
    Utilizza win32com per l'automazione COM.
    """
    
    def __init__(self):
        self.app = None
        self.workbook = None
        self.sheet = None
        self._is_open = False
    
    def open(self):
        """
        Apre il file Excel e sblocca il foglio.
        
        Returns:
            bool: True se apertura riuscita, False altrimenti
        """
        try:
            # Crea istanza Excel
            self.app = win32com.client.DispatchEx("Excel.Application")
            self.app.Visible = False
            self.app.DisplayAlerts = False
            
            # Apri workbook
            self.workbook = self.app.Workbooks.Open(EXCEL_DB_PATH)
            self.sheet = self.workbook.ActiveSheet
            
            # Sblocca foglio protetto
            try:
                self.sheet.Unprotect(Password=SHEET_PASSWORD)
            except Exception as e:
                logger.warning(f"Foglio non protetto o password errata: {e}")
            
            self._is_open = True
            return True
            
        except Exception as e:
            logger.error(f"Errore apertura Excel: {e}")
            self.close(save=False)
            return False
    
    def close(self, save=True):
        """
        Chiude il file Excel.
        
        Args:
            save: Se True, salva le modifiche prima di chiudere
        """
        try:
            if self.sheet and save:
                try:
                    self.sheet.Protect(Password=SHEET_PASSWORD)
                except:
                    pass
            
            if self.workbook:
                self.workbook.Close(SaveChanges=save)
            
            if self.app:
                self.app.Quit()
                
        except Exception as e:
            logger.error(f"Errore chiusura Excel: {e}")
        finally:
            self.workbook = None
            self.app = None
            self.sheet = None
            self._is_open = False
    
    def check_if_exists(self, rda_number):
        """
        Verifica se un numero RDA esiste già nel foglio.
        
        Args:
            rda_number: Numero RDA da cercare
        
        Returns:
            bool: True se esiste, False altrimenti
        """
        if not self._is_open:
            return True  # Fail safe
        
        try:
            last_row = self._get_last_row()
            
            # Leggi tutti i valori della colonna A per performance
            if last_row >= 2:
                rda_range = self.sheet.Range(f"A2:A{last_row}").Value
                
                # Normalizza struttura
                if rda_range:
                    if not isinstance(rda_range, tuple):
                        rda_range = ((rda_range,),)
                    elif not isinstance(rda_range[0], tuple):
                        rda_range = (rda_range,)
                    
                    for cell_tuple in rda_range:
                        cell_value = cell_tuple[0] if isinstance(cell_tuple, tuple) else cell_tuple
                        if cell_value and str(cell_value).strip() == rda_number:
                            return True
            
            return False
            
        except Exception as e:
            logger.error(f"Errore verifica esistenza RDA: {e}")
            return True  # Fail safe
    
    def append_data(self, rda_data):
        """
        Aggiunge i dati di una RDA al foglio Excel.
        
        Args:
            rda_data: Dictionary con i dati RDA estratti dal PDF
        """
        if not self._is_open:
            return
        
        try:
            first_empty_row = self._get_last_row() + 1
            
            # Filtra righe vuote dalla tabella PDF
            valid_rows = [
                row for row in rda_data['table'] 
                if any(cell is not None and str(cell).strip() != '' for cell in row)
            ]
            
            for row_data in valid_rows:
                # Pulisci dati
                cleaned_row = [item if item is not None else "" for item in row_data]
                
                # Gestisci data di consegna
                delivery_date_str = cleaned_row[8] if len(cleaned_row) > 8 else ""
                delivery_date_obj = None
                if delivery_date_str:
                    try:
                        delivery_date_obj = datetime.strptime(str(delivery_date_str), '%d/%m/%Y')
                    except (ValueError, TypeError):
                        delivery_date_obj = delivery_date_str
                
                # Gestisci quantità (converti formato italiano)
                quantity_val = cleaned_row[5] if len(cleaned_row) > 5 else 0
                try:
                    quantity_str = str(quantity_val).replace('.', '').replace(',', '.')
                    quantity_val = float(quantity_str)
                except (ValueError, TypeError):
                    pass
                
                # Costruisci riga da inserire
                # Colonne: A=RDA, B=Commessa, C=Desc1, D=DescMat, E=UM, F=Qty, 
                #          G=APF, H=Link, I=DataRDA, J=DataCons, K=Alert, L=Richiedente
                new_row_values = [
                    rda_data['rda_number_raw'],
                    cleaned_row[1] if len(cleaned_row) > 1 else "",  # Commessa
                    cleaned_row[2] if len(cleaned_row) > 2 else "",  # Descrizione 1
                    cleaned_row[3] if len(cleaned_row) > 3 else "",  # Descrizione Materiale
                    cleaned_row[4] if len(cleaned_row) > 4 else "",  # Unità Misura
                    quantity_val,
                    cleaned_row[7] if len(cleaned_row) > 7 else "",  # APF
                    f'=HYPERLINK("{rda_data["pdf_final_path"]}", "Apri PDF")',
                    rda_data['rda_date_obj'],
                    delivery_date_obj,
                    0,  # Alert level iniziale
                    rda_data.get('requester', '')
                ]
                
                # Inserisci valori
                for i, value in enumerate(new_row_values):
                    self.sheet.Cells(first_empty_row, i + 1).Value = value
                
                first_empty_row += 1
            
            logger.info(f"Aggiunti dati per RDA {rda_data['rda_number_raw']}")
            
        except Exception as e:
            logger.error(f"Errore inserimento dati RDA: {e}")
    
    def update_alerts_and_get_overdue(self):
        """
        Aggiorna i livelli di alert e restituisce le RDA scadute.
        
        Returns:
            list: Lista di dict con RDA scadute
        """
        if not self._is_open:
            return []
        
        overdue_items = []
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        
        try:
            last_row = self._get_last_row()
            
            for row in range(2, last_row + 1):
                # Leggi data RDA (colonna I = 9)
                rda_date = self._parse_cell_date(self.sheet.Cells(row, 9).Value)
                
                if not rda_date:
                    continue
                
                # Calcola giorni trascorsi
                days_diff = (today - rda_date).days
                
                # Calcola livello alert (1 per ogni settimana passata)
                alert_level = days_diff // 7 if days_diff >= 7 else 0
                
                # Aggiorna cella alert (colonna K = 11)
                self.sheet.Cells(row, 11).Value = alert_level
                
                # Se scaduta, aggiungi alla lista
                if alert_level > 0:
                    # Controlla se già consegnata
                    delivery_date = self._parse_cell_date(self.sheet.Cells(row, 10).Value)
                    if delivery_date and delivery_date <= today:
                        continue  # Già consegnata, salta
                    
                    # Raccogli dati per email
                    item_data = {
                        "N°RDA": self.sheet.Cells(row, 1).Value,
                        "Data RDA": rda_date.strftime('%d/%m/%Y'),
                        "Commessa": self.sheet.Cells(row, 2).Value,
                        "Descrizione Materiale": self.sheet.Cells(row, 4).Value,
                        "Unità di Misura": self.sheet.Cells(row, 5).Value,
                        "Quantità Richiesta": self.sheet.Cells(row, 6).Value,
                        "APF": self.sheet.Cells(row, 7).Value,
                        "richiesta da: (giorni)": days_diff,
                        "Richiedente": self.sheet.Cells(row, 12).Value
                    }
                    overdue_items.append(item_data)
            
            return overdue_items
            
        except Exception as e:
            logger.error(f"Errore aggiornamento alert: {e}")
            return []
    
    def delete_empty_rows(self):
        """
        Elimina le righe vuote dalla tabella Excel.
        """
        if not self._is_open:
            return
        
        try:
            if self.sheet.ListObjects.Count > 0:
                table = self.sheet.ListObjects(TABLE_NAME)
                
                # Itera al contrario per evitare problemi con gli indici
                for i in range(table.ListRows.Count, 0, -1):
                    list_row = table.ListRows(i)
                    if self.app.WorksheetFunction.CountA(list_row.Range) == 0:
                        list_row.Delete()
                        
        except Exception as e:
            logger.error(f"Errore eliminazione righe vuote: {e}")
    
    def get_all_data_for_sync(self):
        """
        Legge tutti i dati dal foglio per la sincronizzazione con SQLite.
        
        Returns:
            list: Lista di tuple con i dati
        """
        if not self._is_open:
            return []
        
        data_rows = []
        
        try:
            last_row = self._get_last_row()
            if last_row < 2:
                return []
            
            # Leggi valori e formule in batch
            raw_values = self.sheet.Range(f"A2:L{last_row}").Value
            formulas_col8 = self.sheet.Range(f"H2:H{last_row}").Formula
            
            # Normalizza struttura dati
            raw_values = self._normalize_range_data(raw_values)
            formulas_col8 = self._normalize_formula_data(formulas_col8, last_row - 1)
            
            for i, row_val in enumerate(raw_values):
                if row_val is None or all(v is None for v in row_val):
                    continue
                
                # Estrai path PDF dalla formula HYPERLINK
                pdf_path = self._extract_hyperlink_path(formulas_col8[i])
                
                # Costruisci tupla per database
                item = (
                    safe_str(row_val[0]),          # RDA Number
                    safe_str(row_val[1]),          # Commessa
                    safe_str(row_val[2]),          # Descrizione 1
                    safe_str(row_val[3]),          # Descrizione Materiale
                    safe_str(row_val[4]),          # Unità Misura
                    safe_float(row_val[5]),        # Quantità
                    safe_str(row_val[6]),          # APF
                    pdf_path,                      # PDF Path
                    self._format_cell_date(row_val[8]),   # Data RDA
                    self._format_cell_date(row_val[9]),   # Data Consegna
                    safe_int(row_val[10]),         # Alert Level
                    safe_str(row_val[11])          # Richiedente
                )
                data_rows.append(item)
            
        except Exception as e:
            logger.error(f"Errore lettura dati per sync: {e}")
        
        return data_rows
    
    def fit_columns(self):
        """Adatta la larghezza delle colonne al contenuto."""
        if not self._is_open:
            return
        
        try:
            self.sheet.Columns.AutoFit()
        except:
            pass
    
    def _get_last_row(self):
        """Restituisce il numero dell'ultima riga con dati."""
        return self.sheet.Cells(
            self.sheet.Rows.Count, 
            RDA_REFERENCE_COLUMN_LETTER
        ).End(-4162).Row
    
    def _parse_cell_date(self, cell_value):
        """Converte un valore cella in datetime."""
        if cell_value is None:
            return None
        
        # Se è già un datetime COM
        if hasattr(cell_value, 'timestamp'):
            return datetime.fromtimestamp(cell_value.timestamp())
        
        # Se è una stringa
        if isinstance(cell_value, str):
            try:
                return datetime.strptime(cell_value, '%d/%m/%Y')
            except ValueError:
                pass
        
        return None
    
    def _format_cell_date(self, cell_value):
        """Formatta un valore cella data come stringa dd/mm/yyyy."""
        dt = self._parse_cell_date(cell_value)
        if dt:
            return dt.strftime("%d/%m/%Y")
        return str(cell_value) if cell_value else ""
    
    def _normalize_range_data(self, data):
        """Normalizza i dati letti da un range Excel."""
        if not data:
            return []
        if not isinstance(data, tuple):
            return [[data]]
        if len(data) > 0 and not isinstance(data[0], (tuple, list)):
            return [data]
        return list(data)
    
    def _normalize_formula_data(self, formulas, expected_count):
        """Normalizza le formule lette da un range Excel."""
        if not formulas:
            return [("",)] * expected_count
        
        if isinstance(formulas, str):
            return [(formulas,)]
        
        if isinstance(formulas, tuple):
            if len(formulas) > 0 and not isinstance(formulas[0], (tuple, list)):
                return [formulas]
        
        return list(formulas)
    
    def _extract_hyperlink_path(self, formula_data):
        """Estrae il path da una formula HYPERLINK."""
        try:
            formula = formula_data[0] if isinstance(formula_data, tuple) else formula_data
            match = re.search(r'HYPERLINK\("([^"]+)"', str(formula))
            if match:
                return match.group(1)
        except:
            pass
        return ""
