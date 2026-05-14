#!/bin/bash
cd "$(dirname "$0")"
echo "=== 北一區週報產生器 ==="
echo ""

# Check Python3
if ! command -v python3 &>/dev/null; then
  echo "[錯誤] 找不到 python3，請先安裝 Python 3。"
  read -p "按 Enter 關閉"
  exit 1
fi

# Check required packages
python3 -c "import openpyxl, zeep" 2>/dev/null
if [ $? -ne 0 ]; then
  echo "安裝必要套件中…"
  pip3 install openpyxl zeep --quiet
fi

echo "啟動伺服器（port 8782）…"
echo "請在瀏覽器開啟：http://127.0.0.1:8782/"
echo ""
echo "關閉此視窗即可停止伺服器。"
echo "-------------------------------------------"
python3 server.py
