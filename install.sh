#!/bin/bash
# 北一區週報產生器 — 一鍵安裝腳本
set -e

REPO="https://github.com/samwang38/north1-weekly-report.git"
DEST="$HOME/Desktop/北一區週報-app"

echo "==================================="
echo "  北一區週報產生器 安裝程式"
echo "==================================="
echo ""

# ── 檢查 git ─────────────────────────────────────────────────
if ! command -v git &>/dev/null; then
  echo "正在安裝 git（需要 Xcode Command Line Tools）…"
  xcode-select --install 2>/dev/null || true
  echo ""
  echo "[提示] 安裝完成後，請重新執行此腳本。"
  exit 1
fi

# ── 檢查 python3 ──────────────────────────────────────────────
if ! command -v python3 &>/dev/null; then
  echo "[錯誤] 找不到 Python 3。"
  echo "請至 https://www.python.org/downloads/ 下載安裝後重試。"
  exit 1
fi

# ── Clone 或更新 ──────────────────────────────────────────────
if [ -d "$DEST/.git" ]; then
  echo "偵測到現有安裝，更新至最新版本…"
  cd "$DEST"
  git pull
else
  echo "下載工具中…"
  git clone "$REPO" "$DEST"
fi

# ── 安裝 Python 套件 ──────────────────────────────────────────
# ⚠️ 不要用裸 pip3：它可能屬於另一個 Python（實際踩過 —— pip3 是 3.9 的、
#    python3 是 3.14 的）。一律用 "$PY" -m pip，跟執行的直譯器綁在一起。
PY=""
for c in python3 python3.13 python3.12 python3.11 python3.10 /usr/bin/python3; do
  p=$(command -v "$c" 2>/dev/null) || continue
  if "$p" -c "import openpyxl, pandas" 2>/dev/null; then PY="$p"; break; fi
done
[ -z "$PY" ] && PY=$(command -v python3)
echo "使用 Python：$PY（$("$PY" -V 2>&1)）"

echo "安裝必要套件…"
OK=0
for extra in "" "--user" "--break-system-packages"; do
  if "$PY" -m pip install -r "$DEST/requirements.txt" --quiet $extra 2>/dev/null; then OK=1; break; fi
done
if [ "$OK" -ne 1 ] || ! "$PY" -c "import openpyxl, pandas" 2>/dev/null; then
  echo ""
  echo "[錯誤] 套件安裝失敗。請手動執行："
  echo ""
  echo "         \"$PY\" -m pip install --user -r \"$DEST/requirements.txt\""
  echo ""
  echo "       若出現 externally-managed-environment，把 --user 換成 --break-system-packages。"
  exit 1
fi

# ── 確保 .command 可執行 ──────────────────────────────────────
chmod +x "$DEST/啟動北一區週報.command"

echo ""
echo "==================================="
echo "  ✅ 安裝完成！"
echo "==================================="
echo ""
echo "桌面已建立「北一區週報-app」資料夾。"
echo "雙擊其中的「啟動北一區週報.command」即可使用。"
echo ""
echo "往後每次啟動會自動同步最新版本，無需重新安裝。"
