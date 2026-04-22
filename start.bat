@echo off
title IT-Docu-Maker
cd /d "%~dp0"

echo Starte IT-Docu-Maker...
python --version >nul 2>&1
if errorlevel 1 (
    echo FEHLER: Python nicht gefunden. Bitte Python 3.10+ installieren.
    pause
    exit /b 1
)

python -m pip install -r requirements.txt --quiet
python main.py
pause
