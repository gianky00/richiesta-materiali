# RDA Automation Project

This project automates the extraction of Purchase Requests (RDA) from email attachments (PDFs) and maintains a central Excel registry and a synchronized SQLite database.

## Architecture

- **main_bot.py**: The backend automation service.
    - Scans Outlook for emails with specific criteria.
    - Extracts data from attached PDFs.
    - Updates `database_RDA.xlsm` using COM automation (preserving legacy compatibility).
    - Clones the updated Excel data into `database_RDA.db` (SQLite).
    - Sends email alerts for overdue requests.
- **main_gui.py**: The frontend application for users.
    - Connects to `database_RDA.db`.
    - Provides a searchable table view.
    - Allows direct opening of linked PDF files via right-click context menu.

## Configuration

All configuration (paths, credentials, logic thresholds) is located in `src/config.py`.

## Requirements

- Windows OS (required for Outlook COM automation).
- Microsoft Outlook and Excel installed.
- Python 3.x
- Python packages: `pywin32`, `pdfplumber`.

## Usage

1.  **Automation**: Schedule `main_bot.py` to run periodically.
2.  **User Access**: Distribute `main_gui.py` to users. Ensure they have read access to the network path containing the database and PDFs.
