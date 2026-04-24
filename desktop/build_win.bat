@echo off
chcp 65001 >nul 2>&1
title Build SVN Diff Viewer
cd /d "%~dp0"

echo ================================
echo   Building SVN Diff Viewer.exe
echo ================================
echo.

:: Check / create venv
if not exist ".venv" (
    echo [SETUP] Creating virtual environment...
    python -m venv .venv
)

call .venv\Scripts\activate.bat

echo [SETUP] Installing dependencies...
pip install flask xlrd openpyxl pywebview pyinstaller Pillow -q

echo [BUILD] Packaging with PyInstaller...
pyinstaller ^
    --name "SVN Diff Viewer" ^
    --windowed ^
    --onedir ^
    --noconfirm ^
    --clean ^
    --icon "icon.png" ^
    --add-data "server.py;." ^
    --add-data "svn_excel_diff.py;." ^
    --hidden-import webview ^
    --hidden-import webview.platforms.edgechromium ^
    --hidden-import flask ^
    --hidden-import xlrd ^
    --hidden-import openpyxl ^
    app.py

echo.
echo ================================
echo   Build complete!
echo   Output: dist\SVN Diff Viewer\
echo ================================
pause
