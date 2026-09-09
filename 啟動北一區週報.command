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

# ── 選一個「真的能跑」的 Python ────────────────────────────────
# ⚠️ 絕對不要用裸 pip3：它可能屬於另一個 Python。實際踩過兩次 ——
#    (1) pip3 是 3.9 的、python3 是 3.14 的，裝了半天裝到沒在用的版本
#    (2) 同事的 python3 指到 Homebrew 3.14（空的），套件都在 3.9
#    所以一律用 "$PY" -m pip，跟執行的直譯器綁在一起。
REQ="openpyxl, pandas"                 # 缺這些沒得跑
OPT="gspread, google.oauth2"           # 缺這些只是「月進度」的目標欄空白

PY=""
for c in python3 python3.13 python3.12 python3.11 python3.10 /usr/bin/python3; do
  p=$(command -v "$c" 2>/dev/null) || continue
  if "$p" -c "import $REQ" 2>/dev/null; then PY="$p"; break; fi
done
if [ -z "$PY" ]; then
  PY=$(command -v python3 2>/dev/null)
fi
if [ -z "$PY" ]; then
  echo "[錯誤] 這台電腦找不到 Python 3。"
  echo "       請到 https://www.python.org/downloads/ 下載安裝後再試。"
  read -p "按 Enter 關閉"
  exit 1
fi
echo "使用 Python：$PY（$("$PY" -V 2>&1)）"

# ── 缺什麼就裝什麼 ────────────────────────────────────────────
if ! "$PY" -c "import $REQ, $OPT" 2>/dev/null; then
  echo "安裝必要套件中…"
  OK=0
  for extra in "" "--user" "--break-system-packages"; do
    if "$PY" -m pip install -r requirements.txt --quiet $extra 2>/dev/null; then OK=1; break; fi
  done
  if [ "$OK" -ne 1 ]; then
    echo "  （自動安裝失敗，繼續往下檢查實際缺哪些）"
  fi
fi

# ── 檢查結果：缺必要套件就停，缺選用套件只提醒 ────────────────
if ! "$PY" -c "import $REQ" 2>/dev/null; then
  echo ""
  echo "[錯誤] 缺少必要套件（openpyxl / pandas），而且自動安裝失敗。"
  echo "       請在終端機手動執行這行："
  echo ""
  echo "         \"$PY\" -m pip install --user -r \"$(pwd)/requirements.txt\""
  echo ""
  echo "       若出現 externally-managed-environment，改成："
  echo ""
  echo "         \"$PY\" -m pip install --break-system-packages -r \"$(pwd)/requirements.txt\""
  echo ""
  read -p "按 Enter 關閉"
  exit 1
fi
if ! "$PY" -c "import $OPT" 2>/dev/null; then
  echo "[提醒] 缺 gspread / google-auth，「月進度」的月目標與達成率會空白，"
  echo "       其餘功能正常。詳見 說明-月目標設定.md"
fi
if [ ! -f "$HOME/.config/studioa/forecast-service-account.json" ]; then
  echo "[提醒] 找不到目標表金鑰，「月進度」的月目標與達成率會空白，"
  echo "       其餘功能正常。詳見 說明-月目標設定.md"
fi

# ── 清掉佔用 8782 的舊伺服器（避免「跑到舊版」）──────────────
OLD=$(lsof -ti :8782 2>/dev/null)
if [ -n "$OLD" ]; then
  echo "發現舊伺服器仍在執行，先關閉…"
  echo "$OLD" | xargs kill -9 2>/dev/null
  sleep 1
fi

# ── 順便用 Chrome 打開 ShopperTrak（插件只在 Chrome 運作）+ 週報網頁 ──
echo "用 Chrome 開啟 ShopperTrak 與週報網頁…"
open -a "Google Chrome" "https://analytics.shoppertrak.com/" 2>/dev/null
( sleep 2 && open -a "Google Chrome" "http://127.0.0.1:8782/" ) &

echo "啟動伺服器（port 8782）…"
echo "（瀏覽器會自動開啟，若沒開請手動前往 http://127.0.0.1:8782/ ）"
echo ""
echo "關閉此視窗即可停止伺服器。"
echo "-------------------------------------------"
"$PY" server.py
