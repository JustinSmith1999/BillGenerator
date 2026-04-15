@echo off
REM Windows launcher — double-click to run.
REM First run creates a virtualenv and installs dependencies.

cd /d "%~dp0"

where python >nul 2>&1
if errorlevel 1 (
    echo Python is not on PATH. Install Python 3.10+ from python.org first.
    pause
    exit /b 1
)

if not exist .venv (
    echo First-time setup: creating virtualenv...
    python -m venv .venv
)

call .venv\Scripts\activate.bat

if not exist .venv\.deps_installed (
    echo Installing / updating dependencies...
    pip install --upgrade pip
    pip install -r requirements.txt
    if errorlevel 1 (
        echo Dependency install failed.
        pause
        exit /b 1
    )
    type nul > .venv\.deps_installed
)

echo Launching Bill Categorizer...
python app.py
