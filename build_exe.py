"""
Script per la creazione dell'eseguibile standalone
Utilizza PyInstaller per creare un EXE che funziona senza Python installato

Uso:
    python build_exe.py

Requisiti:
    pip install pyinstaller
"""

import subprocess
import sys
import os
import shutil

# Configurazione
APP_NAME = "RDA_Viewer"
MAIN_SCRIPT = "main_gui.py"
ICON_FILE = None  # Opzionale: "icon.ico"

# Directory
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DIST_DIR = os.path.join(SCRIPT_DIR, "dist")
BUILD_DIR = os.path.join(SCRIPT_DIR, "build")


def check_pyinstaller():
    """Verifica che PyInstaller sia installato."""
    try:
        import PyInstaller
        print(f"✓ PyInstaller trovato (versione {PyInstaller.__version__})")
        return True
    except ImportError:
        print("✗ PyInstaller non trovato. Installazione in corso...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        return True


def clean_build():
    """Pulisce le directory di build precedenti."""
    for dir_path in [DIST_DIR, BUILD_DIR]:
        if os.path.exists(dir_path):
            print(f"Pulizia {dir_path}...")
            shutil.rmtree(dir_path)


def build_gui():
    """Compila l'applicazione GUI."""
    print("\n" + "="*50)
    print("Compilazione RDA Viewer GUI...")
    print("="*50 + "\n")
    
    # Opzioni PyInstaller
    options = [
        MAIN_SCRIPT,
        "--name=" + APP_NAME,
        "--onefile",           # Singolo file EXE
        "--windowed",          # Nessuna console (GUI app)
        "--noconfirm",         # Sovrascrive senza chiedere
        "--clean",             # Pulisce cache
        
        # Includi moduli necessari
        "--hidden-import=sqlite3",
        "--hidden-import=tkinter",
        "--hidden-import=tkinter.ttk",
        
        # Dati aggiuntivi (la cartella src)
        "--add-data=src;src",
    ]
    
    # Aggiungi icona se presente
    if ICON_FILE and os.path.exists(ICON_FILE):
        options.append(f"--icon={ICON_FILE}")
    
    # Esegui PyInstaller
    cmd = [sys.executable, "-m", "PyInstaller"] + options
    print("Comando:", " ".join(cmd))
    print()
    
    result = subprocess.run(cmd, cwd=SCRIPT_DIR)
    
    if result.returncode == 0:
        exe_path = os.path.join(DIST_DIR, f"{APP_NAME}.exe")
        print(f"\n✓ Compilazione completata!")
        print(f"  Eseguibile: {exe_path}")
        return True
    else:
        print("\n✗ Errore durante la compilazione")
        return False


def build_bot():
    """Compila il bot di automazione."""
    print("\n" + "="*50)
    print("Compilazione RDA Bot...")
    print("="*50 + "\n")
    
    options = [
        "main_bot.py",
        "--name=RDA_Bot",
        "--onefile",
        "--console",           # Con console per log
        "--noconfirm",
        "--clean",
        
        # Moduli necessari
        "--hidden-import=win32com.client",
        "--hidden-import=pythoncom",
        "--hidden-import=pdfplumber",
        "--hidden-import=sqlite3",
        
        # Dati aggiuntivi
        "--add-data=src;src",
    ]
    
    cmd = [sys.executable, "-m", "PyInstaller"] + options
    print("Comando:", " ".join(cmd))
    print()
    
    result = subprocess.run(cmd, cwd=SCRIPT_DIR)
    
    if result.returncode == 0:
        exe_path = os.path.join(DIST_DIR, "RDA_Bot.exe")
        print(f"\n✓ Compilazione completata!")
        print(f"  Eseguibile: {exe_path}")
        return True
    else:
        print("\n✗ Errore durante la compilazione")
        return False


def create_launcher_bat():
    """Crea file .bat per avviare l'applicazione."""
    bat_content = '''@echo off
REM Avvia RDA Viewer
start "" "%~dp0RDA_Viewer.exe"
'''
    bat_path = os.path.join(DIST_DIR, "Avvia_RDA_Viewer.bat")
    with open(bat_path, 'w') as f:
        f.write(bat_content)
    print(f"✓ Creato launcher: {bat_path}")


def main():
    print("="*60)
    print("   RDA Application Builder")
    print("   Creazione eseguibili standalone per Windows")
    print("="*60)
    
    # Verifica dipendenze
    if not check_pyinstaller():
        return 1
    
    # Pulisci build precedenti
    clean_build()
    
    # Compila GUI
    if not build_gui():
        return 1
    
    # Compila Bot
    if not build_bot():
        return 1
    
    # Crea launcher
    create_launcher_bat()
    
    # Riepilogo
    print("\n" + "="*60)
    print("   COMPILAZIONE COMPLETATA")
    print("="*60)
    print(f"\nFile creati nella cartella: {DIST_DIR}")
    print("  • RDA_Viewer.exe  - Applicazione GUI")
    print("  • RDA_Bot.exe     - Bot automazione (da schedulare)")
    print("  • Avvia_RDA_Viewer.bat - Launcher")
    print("\nIstruzioni:")
    print("1. Copia la cartella 'dist' sul PC target")
    print("2. Esegui RDA_Viewer.exe per la GUI")
    print("3. Schedula RDA_Bot.exe con Task Scheduler per l'automazione")
    
    return 0


if __name__ == "__main__":
    sys.exit(main())

