@echo off
REM ============================================================
REM  Build BillCategorizer.exe for Windows
REM  Run this on a Windows machine with Python 3.10+ installed.
REM ============================================================

where python >nul 2>&1
if errorlevel 1 (
    echo Python is not on PATH. Install Python 3.10+ from python.org first.
    exit /b 1
)

echo [1/3] Creating / activating virtualenv...
if not exist .venv (
    python -m venv .venv
)
call .venv\Scripts\activate.bat

echo [2/3] Installing dependencies...
pip install --upgrade pip
pip install -r requirements.txt
if errorlevel 1 (
    echo Dependency install failed.
    exit /b 1
)

echo [3/3] Building BillCategorizer.exe with PyInstaller...
pyinstaller ^
    --noconfirm ^
    --onefile ^
    --windowed ^
    --name BillCategorizer ^
    --icon assets\logo.png ^
    --add-data "config.json;." ^
    --add-data "category_map.json;." ^
    --add-data "data;data" ^
    --add-data "assets;assets" ^
    --hidden-import tkinterdnd2 ^
    app.py

if errorlevel 1 (
    echo Build failed.
    exit /b 1
)

echo.
echo ============================================================
echo Build complete. The executable is at:  dist\BillCategorizer.exe
echo ============================================================
