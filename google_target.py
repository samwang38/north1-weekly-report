"""讀取線上「北一區門市月目標」Google 試算表。

只讀不寫。週報的「月進度」分頁用它取得各店的月營業目標與 3PP／SA Care 金額目標。

⚠️ 這支刻意不 import ~/hr_app/google_target.py —— 週報不該綁死在另一個 app 上。
   欄位定義與那支相同（來源是同一張線上表），對應關係見下方 _COLS。

容錯：Google 讀不到（沒網路／金鑰失效／分頁不存在）時回傳快取；
      快取也沒有就回 None，讓週報照常產出、目標欄留白。
"""
from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
KEY_FILE = Path.home() / '.config' / 'studioa' / 'forecast-service-account.json'
CACHE_FILE = ROOT / 'data' / 'target_cache.json'
SHEET_ID = '1SbGH_x5OkX16X5qEut8WkSoZcF9s_qkpNYlOl-Kt9L8'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# 線上表 r4~r9 六店的順序，對應 multistore_engine.STORES 的鍵。
# ⚠️ 表上是「高島屋門市」，engine 是「大葉高島屋門市」——不能直接字串相等，
#    用 _name_matches() 做寬鬆比對，對不上就丟例外，不要默默錯位。
ROW_ORDER = ['004', '005', '024', '046', '054', '057']
FIRST_ROW = 4
_EXPECT_NAME = {
    '004': '士林', '005': '微風', '024': '美麗華',
    '046': '阿波羅', '054': '高島屋', '057': '羅東',
}
# 1-based 欄號 → 回傳鍵（見 memory/google_target_sheet.md 的版面表）
_COLS = {'revenue': 2, 'tpp': 24, 'sac': 26}


class TargetSheetError(RuntimeError):
    pass


def _name_matches(code: str, cell: str) -> bool:
    return _EXPECT_NAME[code] in str(cell or '')


def _load_cache() -> dict:
    try:
        with open(CACHE_FILE, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def _save_cache(cache: dict):
    CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(CACHE_FILE, 'w', encoding='utf-8') as f:
        json.dump(cache, f, ensure_ascii=False, indent=1)


def _fetch(ym: str) -> dict:
    """實際連線抓一個月。ym 格式 'YYYY-MM'，分頁名為 'YYYY.MM'。"""
    import gspread
    from google.oauth2.service_account import Credentials

    gc = gspread.authorize(
        Credentials.from_service_account_file(str(KEY_FILE), scopes=SCOPES))
    try:
        ws = gc.open_by_key(SHEET_ID).worksheet(ym.replace('-', '.'))
    except Exception as e:
        raise TargetSheetError(f'找不到分頁 {ym.replace("-", ".")}: {e}') from e

    rows = ws.get(f'A{FIRST_ROW}:AF{FIRST_ROW + len(ROW_ORDER) - 1}',
                  value_render_option='UNFORMATTED_VALUE')
    if len(rows) < len(ROW_ORDER):
        raise TargetSheetError(f'{ym} 只讀到 {len(rows)} 列，預期 {len(ROW_ORDER)} 列')

    out = {}
    for i, code in enumerate(ROW_ORDER):
        row = rows[i]

        def cell(col_1based):
            v = row[col_1based - 1] if len(row) >= col_1based else 0
            return float(v) if isinstance(v, (int, float)) else 0.0

        if not _name_matches(code, row[0] if row else ''):
            raise TargetSheetError(
                f'{ym} 第 {FIRST_ROW + i} 列店名對不上：預期含「{_EXPECT_NAME[code]}」，'
                f'實得「{row[0] if row else ""}」')
        out[code] = {k: cell(c) for k, c in _COLS.items()}
    return out


def read_targets(ym: str, log=None) -> dict | None:
    """回傳 {store_code: {'revenue':…, 'tpp':…, 'sac':…}}；完全拿不到則 None。

    ym 格式 'YYYY-MM'。線上讀到就順手更新快取；讀不到就退回快取。
    """
    def _log(msg):
        if log:
            log(msg)

    cache = _load_cache()
    try:
        data = _fetch(ym)
    except Exception as e:
        if ym in cache:
            _log(f'  ⚠️ {ym} 目標改用快取（線上讀取失敗：{e}）')
            return cache[ym]
        _log(f'  ⚠️ {ym} 目標取不到，且無快取（{e}）')
        return None

    cache[ym] = data
    _save_cache(cache)
    return data


def read_targets_range(yms: list[str], log=None) -> dict[str, dict]:
    """一次拿多個月：{ym: {...}}，取不到的月份不會出現在結果裡。"""
    out = {}
    for ym in yms:
        got = read_targets(ym, log)
        if got:
            out[ym] = got
    if log:
        log(f'  月目標：{len(out)}/{len(yms)} 個月取得')
    return out
