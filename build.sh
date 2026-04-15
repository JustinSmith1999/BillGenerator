#!/bin/bash
# ============================================================
#  Build a standalone Bill Categorizer binary on macOS or Linux.
#  - macOS produces: dist/BillCategorizer.app (double-clickable)
#  - Linux produces: dist/BillCategorizer (single-file binary)
# ============================================================
set -e
cd "$(dirname "$0")"

if ! command -v python3 >/dev/null 2>&1; then
    echo "Python 3.10+ required. Install from https://www.python.org/downloads/"
    exit 1
fi

echo "[1/3] Creating / activating virtualenv..."
if [ ! -d .venv ]; then
    python3 -m venv .venv
fi
# shellcheck disable=SC1091
source .venv/bin/activate

echo "[2/3] Installing dependencies..."
pip install --upgrade pip >/dev/null
pip install -r requirements.txt

echo "[3/3] Building with PyInstaller..."
# macOS / Linux PyInstaller uses ':' as --add-data separator (Windows uses ';').
pyinstaller \
    --noconfirm \
    --onefile \
    --windowed \
    --name BillCategorizer \
    --add-data "config.json:." \
    --add-data "category_map.json:." \
    --add-data "data:data" \
    --add-data "assets:assets" \
    --hidden-import tkinterdnd2 \
    app.py

echo
echo "============================================================"
if [ -d dist/BillCategorizer.app ]; then
    echo "Done. Double-click: dist/BillCategorizer.app"
else
    echo "Done. Run: ./dist/BillCategorizer"
fi
echo "============================================================"
