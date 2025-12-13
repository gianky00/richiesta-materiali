@echo off

REM Avvia lo script Python in background senza mostrare la finestra del terminale.
REM Utilizza pythonw.exe (versione windowless) per un'esecuzione completamente silenziosa.

start "Esecuzione Silenziosa Script RDA" /B pythonw.exe "\\192.168.11.251\Condivisa\RICHIESTE MATERIALI\RICHIESTE_MATERIALI.py"

exit