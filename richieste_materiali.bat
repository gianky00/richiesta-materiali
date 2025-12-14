@echo off
REM ============================================
REM RDA Automation - Launcher Script
REM ============================================
REM Avvia il bot RDA in background
REM Per uso con Task Scheduler o avvio manuale
REM ============================================

cd /d "%~dp0"

REM Verifica se esiste l'eseguibile compilato
if exist "dist\RDA_Bot.exe" (
    start "RDA Bot" /B "dist\RDA_Bot.exe"
    exit /b
)

REM Altrimenti usa Python
where pythonw >nul 2>nul
if %ERRORLEVEL% == 0 (
    start "RDA Bot" /B pythonw src/main_bot.py
) else (
    python src/main_bot.py
)

exit /b
