import os
import sys
import json
import logging

# Default configuration
DEFAULT_CONFIG = {
    "excel_path": r"\\192.168.11.251\Condivisa\RICHIESTE MATERIALI\DATABASE\database_RDA.xlsm",
    "pdf_folder": r"\\192.168.11.251\Condivisa\RICHIESTE MATERIALI\RDA_PDF",
    "database_dir": r"\\192.168.11.251\Condivisa\RICHIESTE MATERIALI\DATABASE"
}

def get_base_path():
    """Returns the base path of the application (executable dir or script root)."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

CONFIG_FILE = os.path.join(get_base_path(), "config.json")

def load_config():
    """Loads configuration from config.json, or returns defaults if not found."""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                user_config = json.load(f)
                # Merge with defaults to ensure all keys exist
                config = DEFAULT_CONFIG.copy()
                config.update(user_config)
                return config
        except Exception as e:
            logging.error(f"Error loading config.json: {e}")
            return DEFAULT_CONFIG
    return DEFAULT_CONFIG

def save_config(config_data):
    """Saves configuration to config.json."""
    try:
        with open(CONFIG_FILE, 'w') as f:
            json.dump(config_data, f, indent=4)
        return True
    except Exception as e:
        logging.error(f"Error saving config.json: {e}")
        return False

# Initialize config on module load
current_config = load_config()
