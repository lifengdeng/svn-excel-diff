#!/bin/bash
# SVN Diff Tool — macOS / Linux 启动脚本
# 双击运行即可，自动检查并安装依赖

cd "$(dirname "$0")"

echo "================================"
echo "  SVN Excel Diff Tool"
echo "================================"
echo ""

# 检查 Python3
if ! command -v python3 &>/dev/null; then
    echo "[ERROR] python3 not found."
    echo ""
    if [[ "$OSTYPE" == "darwin"* ]]; then
        echo "Install via Homebrew:  brew install python3"
    else
        echo "Install via:  sudo apt install python3 python3-venv  (Debian/Ubuntu)"
        echo "          or: sudo yum install python3              (CentOS/RHEL)"
    fi
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi

# 检查 SVN
if ! command -v svn &>/dev/null; then
    echo "[ERROR] svn not found."
    echo ""
    if [[ "$OSTYPE" == "darwin"* ]]; then
        echo "Install via Homebrew:  brew install svn"
    else
        echo "Install via:  sudo apt install subversion  (Debian/Ubuntu)"
        echo "          or: sudo yum install subversion   (CentOS/RHEL)"
    fi
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi

# 创建虚拟环境（如不存在）
if [ ! -d ".venv" ]; then
    echo "[SETUP] Creating virtual environment..."
    python3 -m venv .venv
    if [ $? -ne 0 ]; then
        echo "[ERROR] Failed to create virtual environment."
        read -p "Press Enter to exit..."
        exit 1
    fi
fi

# 激活虚拟环境
source .venv/bin/activate

# 检查并安装依赖
MISSING=0
python3 -c "import flask" 2>/dev/null || MISSING=1
python3 -c "import xlrd" 2>/dev/null || MISSING=1
python3 -c "import openpyxl" 2>/dev/null || MISSING=1

if [ $MISSING -eq 1 ]; then
    echo "[SETUP] Installing dependencies..."
    pip install flask xlrd openpyxl -q
    if [ $? -ne 0 ]; then
        echo "[ERROR] Failed to install dependencies."
        read -p "Press Enter to exit..."
        exit 1
    fi
    echo "[SETUP] Dependencies installed."
fi

echo ""
echo "[OK] Starting server..."
echo "     Open http://localhost:5000 in your browser."
echo "     Press Ctrl+C to stop."
echo ""

python3 server.py
