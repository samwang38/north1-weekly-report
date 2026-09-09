#!/bin/bash
# 把 Bob 傳來的目標表金鑰放到正確位置。雙擊執行即可。
cd "$(dirname "$0")"

DEST_DIR="$HOME/.config/studioa"
DEST="$DEST_DIR/forecast-service-account.json"
NAME="forecast-service-account.json"

echo "==================================="
echo "  目標表金鑰 安裝小工具"
echo "==================================="
echo ""

# ── 已經裝好就直接結束 ────────────────────────────────────────
if [ -f "$DEST" ]; then
  echo "✅ 金鑰已經在正確位置了："
  echo "   $DEST"
  echo ""
  read -p "按 Enter 關閉"
  exit 0
fi

# ── 建資料夾（預設不存在，本來就要自己建）──────────────────
mkdir -p "$DEST_DIR"
echo "已建立資料夾：$DEST_DIR"
echo ""

# ── 到常見位置找找 Bob 傳過來的檔案 ──────────────────────────
FOUND=""
for d in "$HOME/Downloads" "$HOME/Desktop" "$HOME/Documents"; do
  [ -f "$d/$NAME" ] && { FOUND="$d/$NAME"; break; }
done

if [ -n "$FOUND" ]; then
  echo "在這裡找到金鑰檔："
  echo "   $FOUND"
  echo ""
  read -p "要把它裝到正確位置嗎？(y/n) " ans
  if [ "$ans" = "y" ] || [ "$ans" = "Y" ]; then
    mv "$FOUND" "$DEST"
    chmod 600 "$DEST"
    echo ""
    echo "✅ 安裝完成。"
  else
    echo "已取消。"
    read -p "按 Enter 關閉"
    exit 0
  fi
else
  # ── 找不到就開資料夾請他拖進去 ──────────────────────────────
  echo "還沒收到金鑰檔（$NAME）。"
  echo ""
  echo "請先跟 Bob 要這個檔案（用 AirDrop，不要用群組或 email）。"
  echo "收到之後有兩個做法："
  echo ""
  echo "  做法一：把檔案放到「下載」資料夾，再雙擊這個工具一次，它會自動安裝。"
  echo "  做法二：把檔案直接拖進即將打開的資料夾視窗。"
  echo ""
  read -p "按 Enter 打開資料夾"
  open "$DEST_DIR"
  echo ""
  echo "把 $NAME 拖進那個視窗之後，再雙擊這個工具一次做確認。"
  echo ""
  read -p "按 Enter 關閉"
  exit 0
fi

# ── 驗證：檔案格式對不對、連不連得上 ─────────────────────────
echo ""
echo "檢查中…"

PY=""
for c in python3 python3.13 python3.12 python3.11 /usr/bin/python3; do
  p=$(command -v "$c" 2>/dev/null) || continue
  "$p" -c "import json" 2>/dev/null && { PY="$p"; break; }
done

if [ -z "$PY" ]; then
  echo "（找不到 Python，跳過檢查。檔案已就位，直接啟動週報試試看。）"
  read -p "按 Enter 關閉"
  exit 0
fi

"$PY" - "$DEST" <<'PYEOF'
import json, sys
path = sys.argv[1]
try:
    d = json.load(open(path))
except Exception as e:
    print(f'  ❌ 檔案讀不開或不是有效的金鑰：{e}')
    print('     可能傳輸過程壞了，請跟 Bob 重要一次。')
    sys.exit(1)
if d.get('type') != 'service_account' or 'client_email' not in d:
    print('  ❌ 這不像是服務帳號金鑰，請確認拿到的檔案正確。')
    sys.exit(1)
print(f'  ✅ 金鑰格式正確（帳號：{d["client_email"]}）')
try:
    import gspread  # noqa: F401
    from google.oauth2.service_account import Credentials  # noqa: F401
except ImportError:
    print('  ⚠️ 還缺 gspread / google-auth 套件。')
    print('     雙擊「啟動…週報.command」時會自動安裝，裝完就好了。')
    sys.exit(0)
print('  ✅ 需要的套件也都在')
PYEOF

echo ""
echo "==================================="
echo "  完成！"
echo "==================================="
echo ""
echo "現在雙擊「啟動…週報.command」，「月進度」分頁就會有月目標與達成率了。"
echo ""
read -p "按 Enter 關閉"
