#!/bin/bash
cd "$(dirname "$0")"
echo "=== 北一區週報產生器 ==="
echo ""

# ── 自動更新 ──────────────────────────────────────────────────
if command -v git &>/dev/null && [ -d ".git" ]; then
  echo "檢查更新中…"
  if git pull --quiet 2>/dev/null; then
    echo "已是最新版本。"
  else
    echo "（無法連線更新，繼續使用現有版本）"
  fi
  echo ""
fi

# ── 檢查 Python ───────────────────────────────────────────────
if ! command -v python3 &>/dev/null; then
  echo "[錯誤] 找不到 python3，請先安裝 Python 3。"
  read -p "按 Enter 關閉"
  exit 1
fi

# ── 安裝套件 ──────────────────────────────────────────────────
python3 -c "import openpyxl, pandas" 2>/dev/null
if [ $? -ne 0 ]; then
  echo "安裝必要套件中…"
  pip3 install openpyxl pandas --quiet
fi

# ── 清掉佔用 8782 的舊伺服器（避免「跑到舊版」）──────────────
OLD=$(lsof -ti :8782 2>/dev/null)
if [ -n "$OLD" ]; then
  echo "發現舊伺服器仍在執行，先關閉…"
  echo "$OLD" | xargs kill -9 2>/dev/null
  sleep 1
fi

echo "啟動伺服器（port 8782）…"
echo "請在瀏覽器開啟：http://127.0.0.1:8782/"
echo ""
echo "關閉此視窗即可停止伺服器。"
echo "-------------------------------------------"
python3 server.py
