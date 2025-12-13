@echo off
REM ============================================
REM RDA Viewer - Launcher GUI
REM ============================================
REM Avvia l'interfaccia grafica RDA Viewer
REM ============================================

cd /d "%~dp0"

REM Verifica se esiste l'eseguibile compilato
if exist "dist\RDA_Viewer.exe" (
    start "" "dist\RDA_Viewer.exe"
    exit /b
)

REM Altrimenti usa Python
where pythonw >nul 2>nul
if %ERRORLEVEL% == 0 (
    start "" pythonw main_gui.py
) else (
    python main_gui.py
)

exit /b

