#!/usr/bin/env bash
# Launcher script for FIRE Spending Dashboard
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

if [ ! -d ".venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv .venv
    .venv/bin/pip install -r requirements.txt
fi

echo "Launching Personal Finance & FIRE Dashboard..."
echo "Local database: data/spending.duckdb"
echo "Backup CSV: data/spending_ledger.csv"
echo ""
exec .venv/bin/streamlit run app.py
