@echo off
cd /d "%~dp0"

if exist .venv\Scripts\activate.bat (
    call .venv\Scripts\activate.bat
) else (
    echo Virtual environment not found. Please create it first using:
    echo python -m venv .venv
    echo Then install requirements using:
    echo pip install -r requirements.txt
    pause
    exit /b 1
)

python app_muter.py
if errorlevel 1 (
    echo Error running app_muter.py
    pause
)

deactivate 