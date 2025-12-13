import os
import sys
import pytest
from unittest.mock import MagicMock

# The main scripts are in root.
import main_gui
import main_bot

class TestMainGui:
    def test_database_manager(self, tmp_path):
        db_path = tmp_path / "test.db"

        # Test Init
        db = main_gui.DatabaseManager(str(db_path))

        # Test Connection (creates DB)
        conn = db.get_connection()
        cursor = conn.cursor()

        # Create table for testing
        cursor.execute("CREATE TABLE rda_data (rda_number text, commessa text, descrizione_materiale text, unita_misura text, quantita real, apf text, data_rda text, data_consegna text, alert_level int, richiedente text, pdf_path text)")
        cursor.execute("INSERT INTO rda_data VALUES ('RDA001', 'C001', 'Item1', 'KG', 10, 'A1', '01/01/2023', '01/02/2023', 0, 'User1', 'path.pdf')")
        conn.commit()
        conn.close()

        # Test Fetch
        rows = db.fetch_all_data()
        assert len(rows) == 1
        assert rows[0]['rda_number'] == 'RDA001'

        # Test Stats
        stats = db.get_statistics()
        assert stats['total_rda'] == 1

    def test_formatting_utils(self):
        assert main_gui.format_number(10.0) == "10"
        assert main_gui.format_number(10.5) == "10,5"
        assert main_gui.format_number("25/039") == "25/039" # Should not change
        assert main_gui.format_number(None) == ""

class TestMainBot:
    def test_process_callback(self, mocker):
        # Mock ExcelManager
        mock_excel = MagicMock()
        mock_excel.check_if_exists.return_value = False

        # Mock PDF extraction
        mocker.patch("main_bot.extract_rda_data", return_value={
            'rda_number_raw': 'RDA999',
            'rda_date_str': '2023-01-01'
        })

        # Mock PDF saving
        mocker.patch("main_bot.save_pdf_to_archive", return_value="final/path/file.pdf")

        # Mock Logger
        mocker.patch("main_bot.logger")

        # Get callback
        cb = main_bot.process_pdf_callback(mock_excel)

        # Run callback
        cb("temp.pdf")

        # Verify calls
        mock_excel.check_if_exists.assert_called_with('RDA999')
        mock_excel.append_data.assert_called_once()
        args, _ = mock_excel.append_data.call_args
        assert args[0]['pdf_final_path'] == "final/path/file.pdf"
