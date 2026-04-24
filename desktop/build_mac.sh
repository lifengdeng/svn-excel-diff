#!/bin/bash
# Build macOS .app bundle
# Usage: ./build_mac.sh

set -e
cd "$(dirname "$0")"

echo "================================"
echo "  Building SVN Diff Viewer.app"
echo "================================"
echo ""

# Check / create venv
if [ ! -d ".venv" ]; then
    echo "[SETUP] Creating virtual environment..."
    python3 -m venv .venv
fi

source .venv/bin/activate

echo "[SETUP] Installing dependencies..."
pip install flask xlrd openpyxl pywebview pyinstaller Pillow -q

echo "[BUILD] Packaging with PyInstaller..."
pyinstaller \
    --name "SVN Diff Viewer" \
    --windowed \
    --onedir \
    --noconfirm \
    --clean \
    --icon "icon.png" \
    --add-data "server.py:." \
    --add-data "svn_excel_diff.py:." \
    --hidden-import webview \
    --hidden-import webview.platforms.cocoa \
    --hidden-import flask \
    --hidden-import xlrd \
    --hidden-import openpyxl \
    app.py

echo ""
echo "================================"
echo "  Build complete!"
echo "  Output: dist/SVN Diff Viewer.app"
echo "================================"
