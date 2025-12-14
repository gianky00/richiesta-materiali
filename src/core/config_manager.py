import os
import sys
import json
import logging
import platform

# Default configuration
DEFAULT_CONFIG = {
    "excel_path": r"\\192.168.11.251\Condivisa\RICHIESTE MATERIALI\DATABASE\database_RDA.xlsm",
    "pdf_folder": r"\\192.168.11.251\Condivisa\RICHIESTE MATERIALI\RDA_PDF",
    "database_dir": r"\\192.168.11.251\Condivisa\RICHIESTE MATERIALI\DATABASE"
}

def get_base_path():
    """Returns the base path of the application executable (read-only)."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

def get_data_path():
    """Returns the writable data path for the application (AppData)."""
    system = platform.system()

    if system == "Windows":
        base = os.getenv('LOCALAPPDATA')
        if not base:
             # Fallback
             base = os.path.expanduser("~")
        path = os.path.join(base, "Programs", "RDA Viewer")
    else:
        # Linux/Mac fallback
        path = os.path.join(os.path.expanduser("~"), ".local", "share", "RDA Viewer")

    # Ensure directory exists
    if not os.path.exists(path):
        try:
            os.makedirs(path)
        except OSError as e:
            logging.error(f"Error creating data directory {path}: {e}")

    return path

CONFIG_FILE = os.path.join(get_data_path(), "config.json")

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
        path = get_data_path()
        if not os.path.exists(path):
            os.makedirs(path)

        with open(CONFIG_FILE, 'w') as f:
            json.dump(config_data, f, indent=4)
        return True
    except Exception as e:
        logging.error(f"Error saving config.json: {e}")
        return False

# Initialize config on module load
current_config = load_config()
