#!/bin/bash
# ============================================================
#  BillCategorizer — double-click to run on macOS
#  First run: creates a virtualenv + installs dependencies (~2 min).
#  Every run after: launches instantly.
# ============================================================
set -e

# cd into the folder this script lives in (handles spaces in path)
cd "$(dirname "$0")"

PYTHON_BIN=""
for candidate in python3 python; do
    if command -v "$candidate" >/dev/null 2>&1; then
        PYTHON_BIN="$candidate"
        break
    fi
done

if [ -z "$PYTHON_BIN" ]; then
    osascript -e 'display alert "Python not installed" message "Install Python 3.10+ from python.org, then double-click this file again."' || true
    echo "Python is not installed. Install Python 3.10+ from https://www.python.org/downloads/ and try again."
    read -n 1 -s -r -p "Press any key to exit..."
    exit 1
fi

if [ ! -d .venv ]; then
    echo "First-time setup: creating virtualenv..."
    "$PYTHON_BIN" -m venv .venv
fi

# shellcheck disable=SC1091
source .venv/bin/activate

# Only install deps if requirements.txt has changed since last install
DEPS_STAMP=".venv/.deps_installed"
if [ ! -f "$DEPS_STAMP" ] || [ "requirements.txt" -nt "$DEPS_STAMP" ]; then
    echo "Installing / updating dependencies..."
    pip install --upgrade pip >/dev/null
    pip install -r requirements.txt
    touch "$DEPS_STAMP"
fi

echo "Launching Bill Categorizer..."
python app.py
