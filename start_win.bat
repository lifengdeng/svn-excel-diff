@echo off
chcp 65001 >nul 2>&1
title SVN Excel Diff Tool
cd /d "%~dp0"

echo ================================
echo   SVN Excel Diff Tool
echo ================================
echo.

:: 检查 Python
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] python not found.
    echo.
    echo Please install Python 3.8+ from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

:: 检查 SVN
where svn >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] svn not found.
    echo.
    echo Please install TortoiseSVN with command line tools:
    echo https://tortoisesvn.net/downloads.html
    echo (Check "command line client tools" during installation)
    echo.
    pause
    exit /b 1
)

:: 创建虚拟环境（如不存在）
if not exist ".venv" (
    echo [SETUP] Creating virtual environment...
    python -m venv .venv
    if %errorlevel% neq 0 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
)

:: 激活虚拟环境
call .venv\Scripts\activate.bat

:: 检查依赖
python -c "import flask" 2>nul
if %errorlevel% neq 0 goto :install_deps
python -c "import xlrd" 2>nul
if %errorlevel% neq 0 goto :install_deps
python -c "import openpyxl" 2>nul
if %errorlevel% neq 0 goto :install_deps
goto :start

:install_deps
echo [SETUP] Installing dependencies...
pip install flask xlrd openpyxl -q
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)
echo [SETUP] Dependencies installed.

:start
echo.
echo [OK] Starting server...
echo      Open http://localhost:5000 in your browser.
echo      Press Ctrl+C to stop.
echo.

python server.py
pause
