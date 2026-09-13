@echo off
title GST Reconciliation Studio
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Creating Python virtual environment...
    py -m venv .venv
)

echo Installing required packages...
.venv\Scripts\python.exe -m pip install -r requirements.txt

echo.
echo Starting GST Reconciliation Studio...
.venv\Scripts\python.exe -m streamlit run app.py

pause
