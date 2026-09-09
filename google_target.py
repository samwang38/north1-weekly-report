"""讀取線上月目標 Google 試算表。只讀不寫。

主來源＝全國「YYYY年目標」表（北一北二共用）：
  分頁 '1月'~'12月'；第 1 列表頭，第 2 列起一店一列
  A=店碼(數字，補零成三位)  B=地點名稱  C=營業目標含稅  AC=3PP金額  AE=SA Care金額
  下方 小計／北一區／北二區／桃竹區… 等彙總列靠「A 欄是數字」濾掉
  ⭐ 用店碼查，不靠列位置 —— 這張表會加店，用列號遲早錯位

例外＝MONTH_OVERRIDES 列出的月份改讀舊的北一區目標表：
  分頁 'YYYY.MM'；r4~r9 依序為 士林/微風/美麗華/阿波羅/高島屋/羅東
  B=營業目標含稅  X=3PP金額  Z=SA Care金額
  ⚠️ 舊表只能靠列順序，所以會用 A 欄店名做寬鬆比對，對不上就丟例外

為什麼要有例外：全國表 2026「4月」分頁的北一 6 店忘了更新，留著 3 月的數字
（其他 27 店都正常）。使用者決定不動線上表，改由程式在該月讀舊表覆蓋。
舊表只有北一 6 店，所以北二不受影響（覆蓋不到就沿用全國表的值）。
兩份表其餘月份逐店比對完全相同，故只需覆蓋 4 月。

容錯：Google 讀不到（沒網路／金鑰失效／分頁不存在）時回傳快取；
      快取也沒有就回 None，讓週報照常產出、目標欄留白。
"""
from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
KEY_FILE = Path.home() / '.config' / 'studioa' / 'forecast-service-account.json'
CACHE_FILE = ROOT / 'data' / 'target_cache.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# 需要改讀舊北一表的月份 → {(年, 月): 原因}。
# 全國表 2026/4 分頁的北一 6 店忘了更新（留著 3 月的值），故該月改讀舊表。
# 線上表若之後補正，把這行刪掉即可。
MONTH_OVERRIDES = {(2026, 4): '全國表4月分頁的北一6店未更新，留著3月數字'}

# 全國表：一年一張。新的年度開檔後在這裡加一行即可。
YEAR_SHEETS = {
    2026: '1hcCPs-1gMaNYAD-wTGYepcG_JSsenQTMD_5Lrcz9UlE',   # 2026年目標
}
_COLS_NATIONAL = {'revenue': 3, 'tpp': 29, 'sac': 31}       # 1-based 欄號

# 舊北一表：單一張，分頁 YYYY.MM
NORTH1_SHEET = '1SbGH_x5OkX16X5qEut8WkSoZcF9s_qkpNYlOl-Kt9L8'
_COLS_NORTH1 = {'revenue': 2, 'tpp': 24, 'sac': 26}
NORTH1_ROWS = ['004', '005', '024', '046', '054', '057']    # 對應 r4~r9
NORTH1_FIRST_ROW = 4
# 表上是「高島屋門市」，engine 是「大葉高島屋門市」，故用「包含」而非相等
_NORTH1_NAME = {'004': '士林', '005': '微風', '024': '美麗華',
                '046': '阿波羅', '054': '高島屋', '057': '羅東'}


class TargetSheetError(RuntimeError):
    pass


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


def _open(sheet_id: str, tab: str):
    import gspread
    from google.oauth2.service_account import Credentials
    gc = gspread.authorize(
        Credentials.from_service_account_file(str(KEY_FILE), scopes=SCOPES))
    try:
        return gc.open_by_key(sheet_id).worksheet(tab)
    except Exception as e:
        raise TargetSheetError(f'找不到分頁「{tab}」: {e}') from e


def _cell(row, col_1based):
    v = row[col_1based - 1] if len(row) >= col_1based else 0
    return float(v) if isinstance(v, (int, float)) else 0.0


def _fetch_national(year: int, month: int) -> dict:
    sheet_id = YEAR_SHEETS.get(year)
    if not sheet_id:
        raise TargetSheetError(f'{year} 年的全國目標表未登錄（見 YEAR_SHEETS）')
    ws = _open(sheet_id, f'{month}月')
    out = {}
    for row in ws.get('A2:AF60', value_render_option='UNFORMATTED_VALUE'):
        if not row:
            continue
        code = row[0]
        # 只收「A 欄是數字」的門市列；小計／北一區／Total 等彙總列自動略過
        if not isinstance(code, (int, float)) or isinstance(code, bool):
            continue
        rec = {k: _cell(row, c) for k, c in _COLS_NATIONAL.items()}
        rec['name'] = str(row[1]).strip() if len(row) > 1 else ''
        out[f'{int(code):03d}'] = rec
    if not out:
        raise TargetSheetError(f'{year}/{month} 全國表讀不到任何門市列')
    return out


def _fetch_north1(year: int, month: int) -> dict:
    ws = _open(NORTH1_SHEET, f'{year}.{month:02d}')
    last = NORTH1_FIRST_ROW + len(NORTH1_ROWS) - 1
    rows = ws.get(f'A{NORTH1_FIRST_ROW}:AF{last}',
                  value_render_option='UNFORMATTED_VALUE')
    if len(rows) < len(NORTH1_ROWS):
        raise TargetSheetError(
            f'{year}.{month:02d} 只讀到 {len(rows)} 列，預期 {len(NORTH1_ROWS)} 列')
    out = {}
    for i, code in enumerate(NORTH1_ROWS):
        row = rows[i]
        name = str(row[0]) if row else ''
        if _NORTH1_NAME[code] not in name:
            raise TargetSheetError(
                f'{year}.{month:02d} 第 {NORTH1_FIRST_ROW + i} 列店名對不上：'
                f'預期含「{_NORTH1_NAME[code]}」，實得「{name}」')
        rec = {k: _cell(row, c) for k, c in _COLS_NORTH1.items()}
        rec['name'] = name.strip()
        out[code] = rec
    return out


def _fetch(year: int, month: int, log=None) -> dict:
    """回傳 {店碼: {'revenue','tpp','sac','name'}}。

    以全國表為準；MONTH_OVERRIDES 列到的月份再用舊北一表覆蓋該表有的門市。
    """
    data = _fetch_national(year, month)
    why = MONTH_OVERRIDES.get((year, month))
    if why:
        patch = _fetch_north1(year, month)
        data.update(patch)
        if log:
            log(f'  ℹ️ {year}-{month:02d} 北一 {len(patch)} 店改讀舊目標表（{why}）')
    return data


def read_targets(year: int, month: int, store_codes: list[str] | None = None,
                 log=None) -> dict | None:
    """回傳 {store_code: {'revenue','tpp','sac','name'}}；完全拿不到則 None。

    store_codes 給定時只回傳這些店（缺的店會在 log 提醒）。
    線上讀到就順手更新快取；讀不到就退回快取。
    """
    def _log(msg):
        if log:
            log(msg)

    key = f'{year}-{month:02d}'
    cache = _load_cache()
    try:
        data = _fetch(year, month, log)
        cache[key] = data
        _save_cache(cache)
    except Exception as e:
        if key in cache:
            _log(f'  ⚠️ {key} 目標改用快取（線上讀取失敗：{e}）')
            data = cache[key]
        else:
            _log(f'  ⚠️ {key} 目標取不到，且無快取（{e}）')
            return None

    if store_codes is None:
        return data
    missing = [c for c in store_codes if c not in data]
    if missing:
        _log(f'  ⚠️ {key} 目標表查無店碼 {missing}')
    return {c: data[c] for c in store_codes if c in data}


def duplicate_months(targets: dict[str, dict]) -> set[str]:
    """找出「整個區的目標與前一個月一字不差」的月份。

    正常情況下每月目標都會重訂，全區逐店完全相同幾乎只有一種可能：
    那個月的分頁忘了更新（實際發生過 —— 2026 年 4 月分頁的北一區 6 店
    留著 3 月的數字）。回傳可疑的月份鍵，讓報表把它標出來而不是默默算錯。
    """
    keys = sorted(targets)
    bad = set()
    for prev, cur in zip(keys, keys[1:]):
        a, b = targets[prev], targets[cur]
        if not b or set(a) != set(b):
            continue
        if all(a[c]['revenue'] == b[c]['revenue']
               and a[c]['tpp'] == b[c]['tpp']
               and a[c]['sac'] == b[c]['sac'] for c in b):
            bad.add(cur)
    return bad


def read_targets_range(periods: list[tuple[int, int]],
                       store_codes: list[str] | None = None,
                       log=None) -> tuple[dict[str, dict], set[str]]:
    """一次拿多個月。periods = [(year, month), ...]。

    回傳 ({'YYYY-MM': {...}}, 可疑月份集合)。
    """
    out = {}
    for year, month in periods:
        got = read_targets(year, month, store_codes, log)
        if got:
            out[f'{year}-{month:02d}'] = got
    dup = duplicate_months(out)
    if log:
        log(f'  月目標：{len(out)}/{len(periods)} 個月取得')
        for k in sorted(dup):
            log(f'  ⚠️ {k} 全區目標與前一個月完全相同 —— 線上表該月分頁可能忘了更新')
    return out, dup
