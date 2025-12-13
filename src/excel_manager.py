import win32com.client
import pythoncom
import logging
import re
from datetime import datetime
from .config import EXCEL_DB_PATH, SHEET_PASSWORD, TABLE_NAME, RDA_REFERENCE_COLUMN_LETTER, RDA_REFERENCE_COLUMN_NUMBER

logger = logging.getLogger("RDA_Bot")

class ExcelManager:
    def __init__(self):
        self.app = None
        self.workbook = None
        self.sheet = None

    def open(self):
        try:
            self.app = win32com.client.DispatchEx("Excel.Application")
            self.app.Visible = False
            self.app.DisplayAlerts = False
            self.workbook = self.app.Workbooks.Open(EXCEL_DB_PATH)
            self.sheet = self.workbook.ActiveSheet
            self.sheet.Unprotect(Password=SHEET_PASSWORD)
            return True
        except Exception as e:
            logger.error(f"Error opening Excel: {e}")
            return False

    def close(self, save=True):
        try:
            if self.sheet:
                self.sheet.Protect(Password=SHEET_PASSWORD)
            if self.workbook:
                self.workbook.Close(SaveChanges=save)
            if self.app:
                self.app.Quit()
        except Exception as e:
            logger.error(f"Error closing Excel: {e}")
        finally:
            self.workbook = None
            self.app = None

    def check_if_exists(self, rda_number):
        try:
            last_row = self.sheet.Cells(self.sheet.Rows.Count, RDA_REFERENCE_COLUMN_LETTER).End(-4162).Row
            # A simple loop is fine for < 10k rows usually, but reading range is faster.
            # Keeping original logic for safety.
            for i in range(2, last_row + 1):
                cell_value = self.sheet.Cells(i, RDA_REFERENCE_COLUMN_NUMBER).Value
                if cell_value and str(cell_value).strip() == rda_number:
                    return True
            return False
        except Exception as e:
            logger.error(f"Error checking existence: {e}")
            return True # Fail safe

    def append_data(self, rda_data):
        try:
            first_empty_row = self.sheet.Cells(self.sheet.Rows.Count, RDA_REFERENCE_COLUMN_LETTER).End(-4162).Row + 1
            valid_rows = [row for row in rda_data['table'] if any(cell is not None and str(cell).strip() != '' for cell in row)]

            for row_data in valid_rows:
                cleaned_row = [item if item is not None else "" for item in row_data]

                # Logic copied from original script
                delivery_date_str = cleaned_row[8]
                delivery_date_obj = None
                if delivery_date_str:
                    try:
                        delivery_date_obj = datetime.strptime(str(delivery_date_str), '%d/%m/%Y')
                    except (ValueError, TypeError):
                        delivery_date_obj = delivery_date_str

                quantity_val = cleaned_row[5]
                try:
                    quantity_str = str(quantity_val).replace('.', '').replace(',', '.')
                    quantity_val = float(quantity_str)
                except (ValueError, TypeError):
                    pass

                new_row_values = [
                    rda_data['rda_number_raw'],
                    cleaned_row[1], # Commessa
                    cleaned_row[2], # Descrizione?
                    cleaned_row[3], # Descrizione Materiale
                    cleaned_row[4], # Unita Misura
                    quantity_val,
                    cleaned_row[7], # APF
                    f'=HYPERLINK("{rda_data["pdf_final_path"]}", "Apri PDF")',
                    rda_data['rda_date_obj'],
                    delivery_date_obj,
                    0,
                    rda_data.get('requester', '')
                ]

                for i, value in enumerate(new_row_values):
                    self.sheet.Cells(first_empty_row, i + 1).Value = value
                first_empty_row += 1
            logger.info(f"Added data for RDA {rda_data['rda_number_raw']}")
        except Exception as e:
            logger.error(f"Error appending data: {e}")

    def update_alerts_and_get_overdue(self):
        overdue_items = []
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        try:
            last_row = self.sheet.Cells(self.sheet.Rows.Count, RDA_REFERENCE_COLUMN_LETTER).End(-4162).Row

            # Read all data at once for speed improvement if possible?
            # Win32com cell access is slow. But logic requires updates.
            # We stick to iteration for safety with the specific update logic.

            for row in range(2, last_row + 1):
                rda_date = None
                rda_date_cell_val = self.sheet.Cells(row, 9).Value

                if hasattr(rda_date_cell_val, 'timestamp'):
                    rda_date = datetime.fromtimestamp(rda_date_cell_val.timestamp())
                elif isinstance(rda_date_cell_val, str):
                    try:
                        rda_date = datetime.strptime(rda_date_cell_val, '%d/%m/%Y')
                    except ValueError:
                        continue

                if rda_date:
                    days_diff = (today - rda_date).days
                    alert_level = days_diff // 7 if days_diff >= 7 else 0
                    self.sheet.Cells(row, 11).Value = alert_level

                    if alert_level > 0:
                        delivery_date = None
                        delivery_date_cell_val = self.sheet.Cells(row, 10).Value

                        if hasattr(delivery_date_cell_val, 'timestamp'):
                            delivery_date = datetime.fromtimestamp(delivery_date_cell_val.timestamp())
                        elif isinstance(delivery_date_cell_val, str):
                            try:
                                delivery_date = datetime.strptime(delivery_date_cell_val, '%d/%m/%Y')
                            except ValueError:
                                pass

                        if delivery_date and delivery_date.replace(hour=0, minute=0, second=0, microsecond=0) <= today:
                            continue

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
            logger.error(f"Error updating alerts: {e}")
            return []

    def delete_empty_rows(self):
        try:
            if self.sheet.ListObjects.Count > 0:
                table = self.sheet.ListObjects(TABLE_NAME)
                # Iterating backwards
                for i in range(table.ListRows.Count, 0, -1):
                    list_row = table.ListRows(i)
                    if self.app.WorksheetFunction.CountA(list_row.Range) == 0:
                        list_row.Delete()
        except Exception as e:
            logger.error(f"Error deleting empty rows: {e}")

    def get_all_data_for_sync(self):
        """Reads all data from the sheet to sync with DB."""
        data_rows = []
        try:
            last_row = self.sheet.Cells(self.sheet.Rows.Count, RDA_REFERENCE_COLUMN_LETTER).End(-4162).Row
            if last_row < 2:
                return []

            # Read range as values
            # Using specific columns 1 to 12
            # Warning: Using Range("A2:L" + last_row).Value returns a tuple of tuples.
            # But we need FORMULA for column 8.

            # Optimization: Read values for all, then read formulas for Col 8.
            raw_values = self.sheet.Range(f"A2:L{last_row}").Value
            formulas_col8 = self.sheet.Range(f"H2:H{last_row}").Formula

            # If there's only one row, raw_values might be a 1D tuple if accessed differently,
            # but Range(...).Value typically returns 2D tuple ((v1, ...),) for single row ranges.
            # We ensure it's iterable as rows.
            if not isinstance(raw_values, tuple) and not isinstance(raw_values, list):
                # Single cell case, unlikely here as we select A:L
                raw_values = ((raw_values,),)
            elif isinstance(raw_values, tuple) and len(raw_values) > 0 and not isinstance(raw_values[0], (tuple, list)):
                # 1D tuple case -> wrap to 2D
                raw_values = (raw_values,)

            if isinstance(formulas_col8, tuple) and len(formulas_col8) > 0 and not isinstance(formulas_col8[0], (tuple, list)):
                 formulas_col8 = (formulas_col8,)

            for i, row_val in enumerate(raw_values):
                # row_val is a tuple of 12 items
                # Col 8 (index 7) is the Link. We use the formula instead.
                formula = formulas_col8[i] # This might be the tuple if bulk read?
                # formulas_col8 is usually ((f1,), (f2,), ...)
                if isinstance(formulas_col8, tuple) and isinstance(formulas_col8[0], tuple):
                    formula = formulas_col8[i][0]

                pdf_path = ""
                # Parse formula: =HYPERLINK("path", "text")
                # Regex to extract path.
                match = re.search(r'HYPERLINK\("([^"]+)"', str(formula))
                if match:
                    pdf_path = match.group(1)
                else:
                    # Fallback if it's just a string or not a hyperlink formula
                    pdf_path = str(formula) if formula else ""

                # Construct the tuple for DB
                # Schema: rda_number, commessa, description_1, description_material, unit, qty, apf, pdf_path, date_rda, date_delivery, alert, requester

                # Handling dates in Value
                # pywin32 returns datetime objects for dates usually.

                def fmt_date(d):
                    if isinstance(d, datetime):
                        return d.strftime("%d/%m/%Y")
                    return str(d) if d else ""

                item = (
                    str(row_val[0]) if row_val[0] else "", # RDA Number
                    str(row_val[1]) if row_val[1] else "", # Commessa
                    str(row_val[2]) if row_val[2] else "", # Desc 1
                    str(row_val[3]) if row_val[3] else "", # Desc Material
                    str(row_val[4]) if row_val[4] else "", # Unit
                    row_val[5] if row_val[5] else 0.0,     # Qty
                    str(row_val[6]) if row_val[6] else "", # APF
                    pdf_path,                              # PDF Path
                    fmt_date(row_val[8]),                  # Date RDA
                    fmt_date(row_val[9]),                  # Date Delivery
                    int(row_val[10]) if row_val[10] else 0,# Alert Level
                    str(row_val[11]) if row_val[11] else ""# Requester
                )
                data_rows.append(item)

        except Exception as e:
            logger.error(f"Error reading data for sync: {e}")

        return data_rows

    def fit_columns(self):
        try:
            self.sheet.Columns.AutoFit()
        except:
            pass
