"""
北一區（多店合併）週報計算引擎 v0.1.0
─────────────────────────────────────────
核心計算邏輯沿用單店週報 v1.1.16 已驗證的規則：
  - NET = 銷售金額(含稅) + 銷退金額
  - SALE_TYPES = {'銷售', '尾款'}
  - SAcare 用價目表 × 數量
  - 主機台數：類別3=3001 或 認證機品牌（881/885/886/888），SALE_TYPES - 銷退
  - 成交筆數：所有交易類型各算 1 筆（直接數列數）

與單店版的差異：
  - 不做 by-店員統計（省去 員工代碼 / 員工名稱 / 等級代碼）
  - 多店用「地點代碼」切分
"""

from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path
import pandas as pd

# ─── 店別對應 ─────────────────────────────────────────────────────────
STORES = {
    '004': '士林門市',
    '005': '微風門市',
    '024': '美麗華門市',
    '046': '阿波羅門市',
    '054': '大葉高島屋門市',
    '057': '羅東門市',
}

# ─── 類別代碼 ─────────────────────────────────────────────────────────
# 主機（用於台數計算）
C4_IPHONE_HOST = {4004.0}
C4_IPAD_HOST   = {4005.0, 4006.0, 4041.0}
C4_WATCH_HOST  = {4038.0}
C4_MAC_HOST    = {4001.0, 4002.0}          # 台數用：4001 桌機 / 4002 筆電（涵蓋所有 Mac 新舊機型）
C6_CPU         = {6001.0, 6002.0, 6007.0, 6008.0, 6342.0}  # 細項拆分用（MacBook Air/Pro/iMac/mini/Studio）

# 認證機品牌（跳過類別3篩選，只要品牌符合都算主機）
CERT_BRANDS = {881.0, 885.0, 886.0, 888.0}

# SACare 對應的類別6（for 搭售率分母做類別拆分用，這版只用總主機台數）
C6_SA = {
    'cpu':     {6533.0},
    'ipad':    {6534.0},
    'iphone':  {6535.0},
    'watch':   {6536.0},
    'airpods': {6537.0},
}

# 配件 sheet 類別4 軸
C4_ACCESSORY = [
    (4007.0, 'CPU週邊配件'),
    (4009.0, 'iPhone週邊配件'),
    (4010.0, 'iPad週邊配件'),
    (4012.0, 'CPU/iOS通用週邊配件'),
    (4013.0, 'Speakers'),
    (4014.0, '耳機'),
    (4017.0, '其他週邊配件'),
    (4021.0, '其他收入'),
    (4022.0, 'iOS通用週邊配件'),
    (4026.0, 'SmartA週邊產品'),
    (4039.0, 'Apple Watch週邊配件'),
    (4050.0, '上網卡'),
    (4053.0, '家居'),
    (4069.0, 'AirPods配件'),
]

# 交易類型
SALE_TYPES = {'銷售', '尾款'}

# 抵用券 SKU
COUPON_GIVE   = '99901687'  # 贈出
COUPON_REDEEM = '99901689'  # 抵用


# ─── 會計年度 / 週次 ────────────────────────────────────────────────────
FISCAL_YEAR_START = date(2025, 9, 28)  # FY26 Q1W01 起始日（週日）

def fiscal_week(d: date) -> tuple[int, int]:
    """回傳 (quarter, week_in_quarter)，皆為 1-based"""
    days = (d - FISCAL_YEAR_START).days
    week_idx = days // 7  # 0-based
    return week_idx // 13 + 1, week_idx % 13 + 1


def week_range(week_end: date) -> tuple[date, date]:
    """輸入本週結束日（週六），回傳 (本週開始, 本週結束) 的日期區間"""
    return week_end - timedelta(days=6), week_end


def prev_week(week_end: date) -> tuple[date, date]:
    return week_range(week_end - timedelta(days=7))


def month_to_date(week_end: date) -> tuple[date, date]:
    """本月累積：月初 ~ 本週結束日"""
    return date(week_end.year, week_end.month, 1), week_end


def prev_month_same_span(week_end: date) -> tuple[date, date]:
    """上月同期：上月 1 日 ~ 上月同樣天數"""
    days_into_month = week_end.day
    if week_end.month == 1:
        prev_year, prev_month = week_end.year - 1, 12
    else:
        prev_year, prev_month = week_end.year, week_end.month - 1
    start = date(prev_year, prev_month, 1)
    # 上月同天數（若上月沒這天就取月底）
    try:
        end = date(prev_year, prev_month, days_into_month)
    except ValueError:
        # 上月天數較少（例如本月 31 號對上月 30）
        if prev_month == 12:
            next_m = date(prev_year + 1, 1, 1)
        else:
            next_m = date(prev_year, prev_month + 1, 1)
        end = next_m - timedelta(days=1)
    return start, end


def same_week_last_year(week_end: date) -> tuple[date, date]:
    """去年同週（會計年度口徑：回推 52 週）"""
    start, end = week_range(week_end)
    return start - timedelta(weeks=52), end - timedelta(weeks=52)


# ─── 資料載入 ─────────────────────────────────────────────────────────
REQUIRED_COLUMNS = [
    '單據日期', '單據代碼', '存貨代碼', '名稱', '數量',
    '銷售金額(含稅)', '銷退金額', '淨銷售金額(未稅)', '銷退金額(未稅)',
    '單位成本', '折扣',
    '類別1代碼', '類別3代碼', '類別4代碼', '類別6代碼',
    '交易類型', '品牌代碼', '地點代碼',
]


def find_header_row(filepath) -> int:
    for h in range(0, 20):
        try:
            df = pd.read_excel(filepath, header=h, nrows=1)
            cols = {str(c).strip().lstrip('\ufeff') for c in df.columns.tolist()}
            if '單據日期' in cols and '單據代碼' in cols:
                return h
        except Exception:
            pass
    return 8


def load_800ab(filepath: str | Path) -> pd.DataFrame:
    """載入 800AB 合併檔（含所有店資料）。

    大型 Excel（>20 MB）首次載入後會自動快取至 /tmp/<stem>_clean.csv，
    下次執行時直接讀 CSV，速度從 ~90 秒縮短至 ~2 秒。
    若 CSV 比 Excel 舊或不存在，自動重建。
    """
    filepath = Path(filepath)
    cache = Path(f'/tmp/{filepath.stem}_clean.csv')

    # ── CSV 快取路徑 ──
    use_cache = (cache.exists()
                 and cache.stat().st_mtime >= filepath.stat().st_mtime)
    if use_cache:
        print(f'  載入快取 {cache.name}...', flush=True)
        df = pd.read_csv(cache, dtype={'存貨代碼': str, '地點代碼': str},
                         parse_dates=['單據日期'])
        for col in ['品牌代碼', '類別1代碼', '類別3代碼', '類別4代碼', '類別6代碼']:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        print(f'    {len(df):,} 筆（來自快取）, 地點: {sorted(df["地點代碼"].unique().tolist())}',
              flush=True)
        return df

    # ── 從 Excel 讀取 ──
    print(f'  載入 {filepath.name}（首次，將建立快取）...', flush=True)
    hdr = find_header_row(filepath)
    try:
        df = pd.read_excel(filepath, sheet_name='Sheet', header=hdr)
    except Exception:
        df = pd.read_excel(filepath, header=hdr)
    df.columns = [str(c).strip().lstrip('\ufeff') for c in df.columns]

    # 欄位檢查
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f'800AB 檔案缺少必要欄位: {missing}')

    df['單據日期'] = pd.to_datetime(df['單據日期'], errors='coerce')
    for col in ['品牌代碼', '類別1代碼', '類別3代碼', '類別4代碼', '類別6代碼']:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    df['地點代碼'] = df['地點代碼'].astype(str).str.strip().str.zfill(3)
    # 存貨代碼：pandas 全量讀取時可能推斷為 float (e.g. 99902985.0)
    # fillna('') 避免 NaN→'nan' 汙染；再去掉 float 轉換產生的 .0 尾碼
    df['存貨代碼'] = (df['存貨代碼'].fillna('').astype(str).str.strip()
                      .str.replace(r'\.0$', '', regex=True))
    df['交易類型'] = df['交易類型'].astype(str).str.strip()

    # NET 欄位：銷售金額(含稅) + 銷退金額（銷退金額本身為負值）
    df['NET'] = df.get('銷售金額(含稅)', 0).fillna(0) + df.get('銷退金額', 0).fillna(0)

    # 寫入快取
    try:
        df.to_csv(cache, index=False)
        print(f'    → 快取已儲存至 {cache}', flush=True)
    except Exception as e:
        print(f'    ⚠ 快取寫入失敗（{e}），不影響結果', flush=True)

    print(f'    {len(df):,} 筆, 地點: {sorted(df["地點代碼"].unique().tolist())}', flush=True)
    return df


# ── EPB 即時查詢設定 ──────────────────────────────────────────────────────────
_EPB_APP_ROOT = Path(__file__).resolve().parent


def _find_java() -> tuple[str, str]:
    """Auto-detect java/javac. Returns (java_path, javac_path)."""
    import subprocess, glob
    # 1. Check PATH first
    for cmd in ('java', 'java8'):
        try:
            r = subprocess.run(['which', cmd], capture_output=True, text=True)
            if r.returncode == 0 and r.stdout.strip():
                java = r.stdout.strip()
                javac = java.replace('/jre/bin/java', '/bin/javac').replace('/bin/java', '/bin/javac')
                if Path(javac).exists():
                    return java, javac
        except Exception:
            pass
    # 2. Scan macOS JVM folders (prefer 1.8 if multiple found)
    jvm_root = '/Library/Java/JavaVirtualMachines'
    candidates = sorted(glob.glob(f'{jvm_root}/*/Contents/Home/bin/javac'))
    prefer = [c for c in candidates if '1.8' in c or 'jdk8' in c.lower()]
    ordered = prefer + [c for c in candidates if c not in prefer]
    for javac in ordered:
        java = javac.replace('/bin/javac', '/jre/bin/java')
        if not Path(java).exists():
            java = javac.replace('/bin/javac', '/bin/java')
        if Path(java).exists():
            return java, javac
    raise RuntimeError(
        'Java 未找到。請安裝 JDK 1.8（或確認 /Library/Java/JavaVirtualMachines/ 下有 JDK）。'
    )


def _find_epb_lib() -> str:
    """Auto-detect EPBrowser Shell lib path. Returns classpath fragment."""
    import glob
    search_roots = [
        '/Library/EPBrowser',
        str(Path.home() / 'Library' / 'EPBrowser'),
        '/Applications/EPBrowser.app/Contents/Resources',
    ]
    for root in search_roots:
        jars = glob.glob(f'{root}/**/shell.jar', recursive=True)
        if jars:
            lib_dir = str(Path(jars[0]).parent / 'lib' / '*')
            return f"{jars[0]}:{lib_dir}"
    raise RuntimeError(
        'EPBrowser lib 未找到。請確認 EPBrowser 已安裝（搜尋路徑：/Library/EPBrowser）。'
    )


_JAVA, _JAVAC = _find_java()
_EPB_LIB      = _find_epb_lib()
_JAVA_CP      = f"{_EPB_APP_ROOT}:{_EPB_LIB}"

_TRANS_TYPE_MAP = {'A': '銷售', 'E': '銷退', 'G': '訂金', 'H': '尾款', 'J': '退訂'}


def _run_epb_query(sql: str, timeout: int = 300, max_rows: int = 500_000) -> pd.DataFrame:
    import csv, subprocess
    source = _EPB_APP_ROOT / 'EPBReportQuery.java'
    target = _EPB_APP_ROOT / 'EPBReportQuery.class'
    if not target.exists() or target.stat().st_mtime < source.stat().st_mtime:
        proc = subprocess.run(
            [_JAVAC, '-cp', _JAVA_CP, str(source)],
            cwd=str(_EPB_APP_ROOT), text=True, capture_output=True, timeout=30,
        )
        if proc.returncode != 0:
            raise RuntimeError(f'javac: {proc.stderr.strip()}')
    proc = subprocess.run(
        [_JAVA, '-Dsun.net.client.defaultConnectTimeout=5000',
         '-Dsun.net.client.defaultReadTimeout=120000',
         '-cp', _JAVA_CP, 'EPBReportQuery', sql, str(max_rows)],
        cwd=str(_EPB_APP_ROOT), text=True, capture_output=True, timeout=timeout,
    )
    if proc.returncode != 0:
        raise RuntimeError(f'EPBReportQuery: {(proc.stderr or proc.stdout).strip()}')
    lines = [ln for ln in proc.stdout.splitlines() if ln.strip()]
    if not lines:
        return pd.DataFrame()
    reader = csv.reader(lines, delimiter='\t')
    rows = list(reader)
    headers = [h.upper() for h in rows[0]]
    records = []
    for row in rows[1:]:
        row = (row + [''] * len(headers))[:len(headers)]
        records.append(dict(zip(headers, row)))
    return pd.DataFrame(records)


def load_from_epb(
    start_date, end_date,
    store_codes: list[str] | None = None,
    org_id: str = '01',
) -> pd.DataFrame:
    """Query EPB WebService directly for multi-store data.

    Returns DataFrame compatible with load_800ab() — same REQUIRED_COLUMNS + NET.
    """
    if store_codes is None:
        store_codes = list(STORES.keys())
    end_excl = end_date + timedelta(days=1)
    shop_in  = ', '.join(f"'{s}'" for s in store_codes)
    sql = (
        f"select l.trans_type, l.doc_date, l.doc_id, l.shop_id,"
        f" l.emp_id1, coalesce(e.name, l.emp_id1) as emp_name,"
        f" l.stk_id, l.name as stk_name, l.stk_qty,"
        f" l.line_total_net, l.line_tax, l.trn_cost_price, l.cost_price,"
        f" l.brand_id, l.cat1_id, l.cat3_id, l.cat4_id, l.cat5_id, l.cat6_id, l.disc_num"
        f" from poslinev_bi l"
        f" left join (select emp_id, name from (select emp_id, name,"
        f" row_number() over (partition by emp_id order by lengthb(name) desc) rn"
        f" from ep_emp) where rn=1) e on e.emp_id = l.emp_id1"
        f" where l.org_id = '{org_id}'"
        f" and l.shop_id in ({shop_in})"
        f" and l.doc_date >= to_date('{start_date.isoformat()}', 'yyyy-mm-dd')"
        f" and l.doc_date < to_date('{end_excl.isoformat()}', 'yyyy-mm-dd')"
        f" order by l.shop_id, l.doc_date, l.doc_id"
    )
    print(f'  EPB 查詢 {start_date}~{end_date} ({len(store_codes)} 店)...', flush=True)
    raw = _run_epb_query(sql)
    if raw.empty:
        return pd.DataFrame(columns=REQUIRED_COLUMNS + ['NET'])

    def num(col):
        return pd.to_numeric(raw[col].replace('', '0'), errors='coerce').fillna(0)

    line_total_net = num('LINE_TOTAL_NET')
    line_tax       = num('LINE_TAX')
    trn_cost       = num('TRN_COST_PRICE')
    cost           = num('COST_PRICE')

    df = pd.DataFrame({
        '交易類型':        raw['TRANS_TYPE'].map(_TRANS_TYPE_MAP).fillna(raw['TRANS_TYPE']).astype(str),
        '單據日期':        pd.to_datetime(raw['DOC_DATE'].str[:19], errors='coerce'),
        '單據代碼':        raw['DOC_ID'].astype(str).str.strip(),
        '員工代碼':        raw['EMP_ID1'].astype(str).str.strip(),
        '員工名稱':        raw['EMP_NAME'].astype(str).str.strip(),
        '存貨代碼':        raw['STK_ID'].astype(str).str.strip(),
        '名稱':           raw['STK_NAME'].astype(str).str.strip(),
        '數量':           num('STK_QTY'),
        '銷售金額(含稅)':  line_total_net + line_tax,
        '銷退金額':        0.0,
        '淨銷售金額(未稅)': line_total_net,
        '銷退金額(未稅)':  0.0,
        '單位成本':        trn_cost.where(trn_cost != 0, cost),
        '折扣':           num('DISC_NUM'),
        '類別1代碼':       pd.to_numeric(raw['CAT1_ID'], errors='coerce'),
        '類別3代碼':       pd.to_numeric(raw['CAT3_ID'], errors='coerce'),
        '類別4代碼':       pd.to_numeric(raw['CAT4_ID'], errors='coerce'),
        '類別5代碼':       pd.to_numeric(raw['CAT5_ID'], errors='coerce'),
        '類別6代碼':       pd.to_numeric(raw['CAT6_ID'], errors='coerce'),
        '品牌代碼':        pd.to_numeric(raw['BRAND_ID'], errors='coerce'),
        '地點代碼':        raw['SHOP_ID'].astype(str).str.strip().str.zfill(3),
    })
    df['NET'] = df['銷售金額(含稅)'] + df['銷退金額']
    print(f'    {len(df):,} 筆, 地點: {sorted(df["地點代碼"].unique().tolist())}', flush=True)
    return df


def load_sacare_prices(filepath: str | Path) -> dict[str, float]:
    """載入 SAcare 對應價目表 → {存貨代碼: 價格}"""
    df = pd.read_excel(filepath, header=0)
    prices = {}
    for _, row in df.iterrows():
        if pd.notna(row.iloc[0]) and pd.notna(row.iloc[2]):
            prices[str(row.iloc[0]).strip()] = float(row.iloc[2])
    return prices


# ─── 期間篩選 ─────────────────────────────────────────────────────────
def filter_period(df: pd.DataFrame, start: date, end: date,
                  store_code: str | None = None) -> pd.DataFrame:
    s = pd.Timestamp(start)
    e = pd.Timestamp(f'{end} 23:59:59')
    m = (df['單據日期'] >= s) & (df['單據日期'] <= e)
    if store_code is not None:
        m &= (df['地點代碼'] == store_code)
    return df[m]


# ─── 計算函式 ─────────────────────────────────────────────────────────
def total_revenue_excl_sacare(df: pd.DataFrame, sa_codes: set[str]) -> int:
    """總業績（未加SA Care）= non_sa 的 NET 總和"""
    non_sa = df[~df['存貨代碼'].isin(sa_codes)]
    return int(non_sa['NET'].sum())


def threepp_revenue_excl_sacare(df: pd.DataFrame, sa_codes: set[str]) -> int:
    """3PP（未加SA Care）= 類別3=3003 的 NET（排除 SA SKU）"""
    sub = df[(~df['存貨代碼'].isin(sa_codes)) & (df['類別3代碼'] == 3003.0)]
    return int(sub['NET'].sum())


def sacare_revenue(df: pd.DataFrame, prices: dict[str, float]) -> int:
    """SA Care 金額 = Σ(數量 × 價目表單價)，只算 SALE_TYPES，扣除銷退"""
    sa_rows = df[df['存貨代碼'].isin(prices.keys())].copy()
    if sa_rows.empty:
        return 0
    sold = sa_rows[sa_rows['交易類型'].isin(SALE_TYPES)]
    ret  = sa_rows[sa_rows['交易類型'] == '銷退']
    sold_amt = (sold['存貨代碼'].map(prices) * sold['數量'].fillna(0)).sum()
    ret_amt  = (ret['存貨代碼'].map(prices) * ret['數量'].abs()).sum()
    return int(sold_amt - ret_amt)


def acpp_plus_revenue(df: pd.DataFrame) -> int:
    """AC+ 金額 = 名稱含「代收保費- AppleCare+」的含稅金額（NET）"""
    m = df['名稱'].astype(str).str.contains('代收保費', na=False) & \
        df['名稱'].astype(str).str.contains('AppleCare', case=False, na=False)
    return int(df.loc[m, 'NET'].sum())


def host_units(df: pd.DataFrame, c4_set: set[float] | None = None,
               c6_set: set[float] | None = None) -> int:
    """主機台數：類別3=3001 或 認證機品牌 (881/885/886/888)
    可用 c4_set (iPhone/iPad/Watch) 或 c6_set (Mac) 進一步篩選
    """
    base = (df['類別3代碼'] == 3001.0) | (df['品牌代碼'].isin(CERT_BRANDS))
    if c4_set is not None:
        m = base & df['類別4代碼'].isin(c4_set)
    elif c6_set is not None:
        m = base & df['類別6代碼'].isin(c6_set)
    else:
        m = base
    sale = df.loc[m & df['交易類型'].isin(SALE_TYPES), '數量'].sum()
    ret  = df.loc[m & (df['交易類型'] == '銷退'), '數量'].abs().sum()
    return int(sale - ret)


def txn_count(df: pd.DataFrame) -> int:
    """成交筆數 = distinct (單據代碼 × 交易類型) 組合數，對應門市單據分析報表邏輯"""
    return int(df[['單據代碼', '交易類型']].dropna(subset=['單據代碼']).drop_duplicates().shape[0])


def acpp_units_total(df: pd.DataFrame) -> int:
    """ACPP+ 台數 = 類別3=3032 的數量，扣除銷退"""
    base = (df['類別3代碼'] == 3032.0)
    sale = df.loc[base & df['交易類型'].isin(SALE_TYPES), '數量'].sum()
    ret  = df.loc[base & (df['交易類型'] == '銷退'), '數量'].abs().sum()
    return int(sale - ret)


def acpp_units_by_device(df: pd.DataFrame, keyword: str) -> int:
    """ACPP+ 台數（分類別）：類別3=3032 + 名稱含 keyword（mac/ipad/iphone/watch/airpods），扣除銷退"""
    base = (df['類別3代碼'] == 3032.0) & \
        df['名稱'].astype(str).str.lower().str.contains(keyword.lower(), na=False)
    sale = df.loc[base & df['交易類型'].isin(SALE_TYPES), '數量'].sum()
    ret  = df.loc[base & (df['交易類型'] == '銷退'), '數量'].abs().sum()
    return int(sale - ret)


def sacare_units(df: pd.DataFrame, c6_set: set[float], sa_codes: set[str]) -> int:
    """SAcare 台數：類別6 ∈ c6_set 且 存貨代碼 ∈ SA 價目表"""
    m = df['類別6代碼'].isin(c6_set) & df['存貨代碼'].isin(sa_codes)
    sale = df.loc[m & df['交易類型'].isin(SALE_TYPES), '數量'].sum()
    ret  = df.loc[m & (df['交易類型'] == '銷退'), '數量'].abs().sum()
    return int(sale - ret)


def coupon_stats(df: pd.DataFrame) -> tuple[int, int]:
    """抵用券：回傳 (贈出數量, 抵用數量)"""
    def _net(sku):
        sub = df.loc[df['存貨代碼'] == sku]
        sale = sub.loc[sub['交易類型'].isin(SALE_TYPES), '數量'].abs().sum()
        ret  = sub.loc[sub['交易類型'] == '銷退', '數量'].abs().sum()
        return int(sale - ret)
    return _net(COUPON_GIVE), _net(COUPON_REDEEM)


# 禮券（高島屋等百貨禮券折抵）
C6_VOUCHER       = {6884.0, 6888.0, 6889.0}  # 禮券類別6（負值折抵）
VOUCHER_EXCL_C3  = 3027.0                      # 排除類別3
VOUCHER_EXCL_EMP = 'SA999'                     # 排除員工（月底入帳發行，非顧客使用）


def voucher_revenue(df: pd.DataFrame) -> int:
    """禮券金額 = 類別6 ∈ {6884,6888,6889} 的 signed NET 加總後取絕對值。
    排除 emp_id1='SA999'（月底入帳發行）與 cat3=3027。
    """
    m = df['類別6代碼'].isin(C6_VOUCHER) & (df['類別3代碼'] != VOUCHER_EXCL_C3)
    if '員工代碼' in df.columns:
        m &= (df['員工代碼'] != VOUCHER_EXCL_EMP)
    return int(abs(df.loc[m, 'NET'].sum()))


def accessory_by_c4(df: pd.DataFrame, sa_codes: set[str]) -> dict[float, int]:
    """配件 sheet 按類別4 加總銷售金額（類別3=3003, 排除 SA SKU）"""
    sub = df[(~df['存貨代碼'].isin(sa_codes)) & (df['類別3代碼'] == 3003.0)]
    result = {}
    for c4, _name in C4_ACCESSORY:
        result[c4] = int(sub.loc[sub['類別4代碼'] == c4, 'NET'].sum())
    return result


# ─── 聚合：單一店 × 單一期間 全部指標 ─────────────────────────────────
def calc_store_metrics(df: pd.DataFrame, start: date, end: date,
                       store_code: str | None, sa_prices: dict) -> dict:
    """計算一店一區間的所有 KPI，回傳 dict"""
    sa_codes = set(sa_prices.keys())
    d = filter_period(df, start, end, store_code)

    total_excl_sa = total_revenue_excl_sacare(d, sa_codes)
    tpp_excl_sa   = threepp_revenue_excl_sacare(d, sa_codes)
    sa_rev        = sacare_revenue(d, sa_prices)

    return {
        'total_excl_sa': total_excl_sa,               # 總業績(未加SA Care)
        'total_rev':     total_excl_sa + sa_rev,       # 總業績
        'tpp_excl_sa':   tpp_excl_sa,                  # 3PP(未加SA Care)
        'tpp_rev':       tpp_excl_sa + sa_rev,         # 3PP
        'sa_rev':        sa_rev,                       # SA Care
        'acpp_plus':     acpp_plus_revenue(d),         # AC+ 金額
        'cpu_units':     host_units(d, c4_set=C4_MAC_HOST),
        'iphone_units':  host_units(d, c4_set=C4_IPHONE_HOST),
        'ipad_units':    host_units(d, c4_set=C4_IPAD_HOST),
        'watch_units':   host_units(d, c4_set=C4_WATCH_HOST),
        'txn_count':     txn_count(d),
        'acpp_total':    acpp_units_total(d),
        'acpp_mac':      acpp_units_by_device(d, 'mac'),
        'acpp_ipad':     acpp_units_by_device(d, 'ipad'),
        'acpp_iphone':   acpp_units_by_device(d, 'iphone'),
        'acpp_watch':    acpp_units_by_device(d, 'watch'),
        'acpp_airpods':  acpp_units_by_device(d, 'airpods'),
        'sa_cpu':        sacare_units(d, C6_SA['cpu'], sa_codes),
        'sa_ipad':       sacare_units(d, C6_SA['ipad'], sa_codes),
        'sa_iphone':     sacare_units(d, C6_SA['iphone'], sa_codes),
        'sa_watch':      sacare_units(d, C6_SA['watch'], sa_codes),
        'sa_airpods':    sacare_units(d, C6_SA['airpods'], sa_codes),
        'airpods_units': airpods_host_units(d),
        'coupon_rev':    voucher_revenue(d),   # 禮券金額（高島屋禮券：類別6 6884/6888/6889）
    }


def calc_accessory_by_c4(df: pd.DataFrame, start: date, end: date,
                          store_code: str | None, sa_codes: set[str]) -> dict[float, int]:
    d = filter_period(df, start, end, store_code)
    return accessory_by_c4(d, sa_codes)


# ─── 本週其他細項：類別代碼常數 ──────────────────────────────────────────
C6_SCREEN_PROT  = {6077.0, 6078.0, 6079.0}          # 保貼（iPhone/iPad 共用）
C6_CASE         = {6053.0, 6054.0, 6055.0, 6056.0}  # 保護殼（iPhone/iPad 共用）
C6_LENS_PROT    = {6081.0}                           # 鏡頭貼（iPhone only）
C6_PENCIL_FIRST = {6100.0}                           # 原廠筆（Apple Pencil）
C6_PENCIL_THIRD = {6093.0}                           # 副廠筆
C6_IPAD_KB      = {6088.0}                           # iPad 鍵盤
C6_WATCH_BAND   = {6506.0}                           # Watch 錶帶
C4_IPAD_ACC     = {4010.0, 4011.0}                  # iPad 配件 類別4


SACARE_CHECKNEW_CAT5 = 5807.0  # 090CaC SHOPPOSB 類別5：SACare 檢測新機


def sacare_checknew_units(df: pd.DataFrame, c6_set: set[float], sa_codes: set[str]) -> int:
    """SACare 檢測新機台數：cat5=5807 + 類別6 ∈ c6_set + 存貨代碼 ∈ SA 價目表"""
    if '類別5代碼' not in df.columns:
        # 無 cat5 欄位時（例如 800AB 路徑）退回舊邏輯
        return sacare_units(df, c6_set, sa_codes)
    m = ((df['類別5代碼'] == SACARE_CHECKNEW_CAT5) &
         df['類別6代碼'].isin(c6_set) &
         df['存貨代碼'].isin(sa_codes))
    sale = df.loc[m & df['交易類型'].isin(SALE_TYPES), '數量'].sum()
    ret  = df.loc[m & (df['交易類型'] == '銷退'), '數量'].abs().sum()
    return int(sale - ret)


def recycle_units(df: pd.DataFrame, sku: str) -> int:
    """回收件數：環保回收(99200234) / 線材回收(99200251) 等 SKU"""
    m = df['存貨代碼'] == sku
    sale = df.loc[m & df['交易類型'].isin(SALE_TYPES), '數量'].sum()
    ret  = df.loc[m & (df['交易類型'] == '銷退'), '數量'].abs().sum()
    return int(sale - ret)


def speaker_revenue(df: pd.DataFrame, include_momax: bool = True) -> int:
    """喇叭銷售金額：類別4=4013 & 類別3=3003
    include_momax=False 時排除 momax 品牌（品牌代碼=453 或名稱含 momax）
    """
    m = (df['類別4代碼'] == 4013.0) & (df['類別3代碼'] == 3003.0)
    if not include_momax:
        is_momax = (df['品牌代碼'] == 453.0) | \
                   df['名稱'].astype(str).str.lower().str.contains('momax', na=False)
        m = m & ~is_momax
    return int(df.loc[m, 'NET'].sum())


def accessory_units_revenue(df: pd.DataFrame,
                             c4: float | set,
                             c6_set: set[float]) -> tuple[int, int]:
    """配件件數與金額：SALE_TYPES 銷售扣銷退
    c4 可傳 float（單一類別4）或 set（多個類別4）
    回傳 (件數, 金額)
    """
    if isinstance(c4, (set, frozenset)):
        m = df['類別4代碼'].isin(c4) & df['類別6代碼'].isin(c6_set)
    else:
        m = (df['類別4代碼'] == c4) & df['類別6代碼'].isin(c6_set)
    sale_qty = df.loc[m & df['交易類型'].isin(SALE_TYPES), '數量'].sum()
    ret_qty  = df.loc[m & (df['交易類型'] == '銷退'), '數量'].abs().sum()
    units    = int(sale_qty - ret_qty)
    revenue  = int(df.loc[m, 'NET'].sum())
    return units, revenue


def ipad_pencil_first_units(df: pd.DataFrame) -> int:
    """iPad 原廠筆件數：品牌=073、類別6=6001、名稱含 Pencil、排除名稱含「筆尖」"""
    m = (df['品牌代碼'] == 73.0) & \
        df['類別6代碼'].isin(C6_PENCIL_FIRST) & \
        df['名稱'].astype(str).str.contains('Pencil', case=False, na=False) & \
        ~df['名稱'].astype(str).str.contains('筆尖', na=False)
    sale = df.loc[m & df['交易類型'].isin(SALE_TYPES), '數量'].sum()
    ret  = df.loc[m & (df['交易類型'] == '銷退'), '數量'].abs().sum()
    return int(sale - ret)


def ipad_pencil_third_units(df: pd.DataFrame) -> int:
    """iPad 副廠筆件數：類別3=3003、類別6=6093、排除名稱含「筆尖」"""
    m = (df['類別3代碼'] == 3003.0) & \
        df['類別6代碼'].isin(C6_PENCIL_THIRD) & \
        ~df['名稱'].astype(str).str.contains('筆尖', na=False)
    sale = df.loc[m & df['交易類型'].isin(SALE_TYPES), '數量'].sum()
    ret  = df.loc[m & (df['交易類型'] == '銷退'), '數量'].abs().sum()
    return int(sale - ret)


def ipad_keyboard_units(df: pd.DataFrame) -> int:
    """iPad 鍵盤件數：類別3=3003、類別6=6088"""
    m = (df['類別3代碼'] == 3003.0) & df['類別6代碼'].isin(C6_IPAD_KB)
    sale = df.loc[m & df['交易類型'].isin(SALE_TYPES), '數量'].sum()
    ret  = df.loc[m & (df['交易類型'] == '銷退'), '數量'].abs().sum()
    return int(sale - ret)


def watch_screen_prot_units(df: pd.DataFrame) -> int:
    """Watch 保貼件數：類別3=3003、類別4=4039、類別6=6077"""
    m = (df['類別3代碼'] == 3003.0) & (df['類別4代碼'] == 4039.0) & \
        (df['類別6代碼'] == 6077.0)
    sale = df.loc[m & df['交易類型'].isin(SALE_TYPES), '數量'].sum()
    ret  = df.loc[m & (df['交易類型'] == '銷退'), '數量'].abs().sum()
    return int(sale - ret)


def watch_band_units(df: pd.DataFrame) -> int:
    """Watch 錶帶件數：類別3=3003、類別4=4039、類別6=6506"""
    m = (df['類別3代碼'] == 3003.0) & (df['類別4代碼'] == 4039.0) & \
        df['類別6代碼'].isin(C6_WATCH_BAND)
    sale = df.loc[m & df['交易類型'].isin(SALE_TYPES), '數量'].sum()
    ret  = df.loc[m & (df['交易類型'] == '銷退'), '數量'].abs().sum()
    return int(sale - ret)


def calc_misc_metrics(df: pd.DataFrame, start: date, end: date,
                      store_code: str | None, sa_prices: dict) -> dict:
    """BY店 本週其他細項 所需的所有 KPI，回傳 dict"""
    sa_codes = set(sa_prices.keys())
    d = filter_period(df, start, end, store_code)

    # SAcare 檢測新機（按裝置別，使用 cat5=5807 篩選）
    sa_mac     = sacare_checknew_units(d, C6_SA['cpu'],     sa_codes)
    sa_iphone  = sacare_checknew_units(d, C6_SA['iphone'],  sa_codes)
    sa_ipad    = sacare_checknew_units(d, C6_SA['ipad'],    sa_codes)
    sa_watch   = sacare_checknew_units(d, C6_SA['watch'],   sa_codes)
    sa_airpods = sacare_checknew_units(d, C6_SA['airpods'], sa_codes)

    # 活動使用（抵用券贈出/兌換：99901687/99901689；環保/線材回收：99200234/99200251）
    coupon_give, coupon_redeem = coupon_stats(d)
    eco   = recycle_units(d, '99200234')
    cable = recycle_units(d, '99200251')

    # Mysetup（手動填入）：只提供主機總數做分母
    host_total = host_units(d)

    # 喇叭
    spk_with    = speaker_revenue(d, include_momax=True)
    spk_without = speaker_revenue(d, include_momax=False)

    # iPhone 配件
    iphone_host                          = host_units(d, c4_set=C4_IPHONE_HOST)
    iphone_prot_qty, iphone_prot_rev     = accessory_units_revenue(d, 4009.0, C6_SCREEN_PROT)
    iphone_case_qty, iphone_case_rev     = accessory_units_revenue(d, 4009.0, C6_CASE)
    iphone_lens_qty, iphone_lens_rev     = accessory_units_revenue(d, 4009.0, C6_LENS_PROT)

    # iPad 配件
    ipad_host              = host_units(d, c4_set=C4_IPAD_HOST)
    ipad_pencil1           = ipad_pencil_first_units(d)
    ipad_pencil3           = ipad_pencil_third_units(d)
    ipad_prot_qty, _       = accessory_units_revenue(d, C4_IPAD_ACC, C6_SCREEN_PROT)
    ipad_case_qty, _       = accessory_units_revenue(d, C4_IPAD_ACC, C6_CASE)
    ipad_kb                = ipad_keyboard_units(d)

    # Watch 配件
    watch_host = host_units(d, c4_set=C4_WATCH_HOST)
    watch_prot = watch_screen_prot_units(d)
    watch_band = watch_band_units(d)

    return {
        'sa_mac': sa_mac, 'sa_iphone': sa_iphone, 'sa_ipad': sa_ipad,
        'sa_watch': sa_watch, 'sa_airpods': sa_airpods,
        'sa_total': sa_mac + sa_iphone + sa_ipad + sa_watch + sa_airpods,
        'coupon_give': coupon_give, 'coupon_redeem': coupon_redeem,
        'eco': eco, 'cable': cable, 'host_total': host_total,
        'spk_with': spk_with, 'spk_without': spk_without,
        'iphone_host': iphone_host,
        'iphone_prot_qty': iphone_prot_qty, 'iphone_prot_rev': iphone_prot_rev,
        'iphone_case_qty': iphone_case_qty, 'iphone_case_rev': iphone_case_rev,
        'iphone_lens_qty': iphone_lens_qty, 'iphone_lens_rev': iphone_lens_rev,
        'ipad_host': ipad_host,
        'ipad_pencil1': ipad_pencil1, 'ipad_pencil3': ipad_pencil3,
        'ipad_prot_qty': ipad_prot_qty, 'ipad_case_qty': ipad_case_qty,
        'ipad_kb': ipad_kb,
        'watch_host': watch_host, 'watch_prot': watch_prot, 'watch_band': watch_band,
    }


# ─── 人員銷售：輔助函式 ────────────────────────────────────────────────
# AirPods 主機類別6（與 fill_multistore_excel.py 各 BY店 sheet 保持一致）
C6_AIRPODS_HOST = {6258.0, 6312.0, 6330.0}
C3_AIRPODS_HOST = 3002.0


def airpods_host_units(df: pd.DataFrame) -> int:
    """AirPods 主機台數：類別6 ∈ {6258/6312/6330} 且 類別3=3002（SALE_TYPES 扣銷退）
    與 fill_multistore_excel.py 各 BY店 sheet 的計算邏輯保持一致。
    """
    m = df['類別6代碼'].isin(C6_AIRPODS_HOST) & (df['類別3代碼'] == C3_AIRPODS_HOST)
    sale = df.loc[m & df['交易類型'].isin(SALE_TYPES), '數量'].sum()
    ret  = df.loc[m & (df['交易類型'] == '銷退'), '數量'].abs().sum()
    return int(sale - ret)


def airpods_acc_units(df: pd.DataFrame) -> int:
    """AirPods 配件件數：類別4=4069（SALE_TYPES 銷售扣銷退）"""
    m = df['類別4代碼'] == 4069.0
    sale = df.loc[m & df['交易類型'].isin(SALE_TYPES), '數量'].sum()
    ret  = df.loc[m & (df['交易類型'] == '銷退'), '數量'].abs().sum()
    return int(sale - ret)


def calc_person_metrics(df: pd.DataFrame, start: date, end: date,
                        store_code: str | None, sa_prices: dict) -> list[dict]:
    """BY人員銷售：依「員工名稱」分群，回傳每位員工的 KPI list。
    若資料中無「員工名稱」欄位或無資料則回傳空 list。
    每筆 dict 包含：name, iphone_host/acpp_iphone/sa_iphone, ipad_..., watch_...,
    mac_..., airpods_..., iPhone/iPad/Watch配件, airpods_acc, spk_with/without。
    """
    sa_codes = set(sa_prices.keys())
    d = filter_period(df, start, end, store_code)

    if '員工名稱' not in d.columns or d.empty:
        return []

    results = []
    # 依「員工代碼」升冪排序；如欄位不存在則退回依姓名排序
    if '員工代碼' in d.columns:
        pairs = (
            d[['員工代碼', '員工名稱']]
            .dropna(subset=['員工名稱'])
            .drop_duplicates(subset=['員工名稱'])
            .sort_values('員工代碼')
        )
        emp_pairs = list(zip(pairs['員工代碼'].tolist(), pairs['員工名稱'].tolist()))
    else:
        emp_pairs = [('', n) for n in sorted(d['員工名稱'].dropna().unique())]
    for emp_id, emp in emp_pairs:
        ed = d[d['員工名稱'] == emp]

        # ── 主機台數 ──
        iphone_host  = host_units(ed, c4_set=C4_IPHONE_HOST)
        ipad_host    = host_units(ed, c4_set=C4_IPAD_HOST)
        watch_host   = host_units(ed, c4_set=C4_WATCH_HOST)
        mac_host     = host_units(ed, c4_set=C4_MAC_HOST)
        apo_host     = airpods_host_units(ed)

        # ── ACPP+ ──
        acpp_iphone  = acpp_units_by_device(ed, 'iphone')
        acpp_ipad    = acpp_units_by_device(ed, 'ipad')
        acpp_watch   = acpp_units_by_device(ed, 'watch')
        acpp_mac     = acpp_units_by_device(ed, 'mac')
        acpp_airpods = acpp_units_by_device(ed, 'airpods')

        # ── SAcare ──
        sa_iphone  = sacare_units(ed, C6_SA['iphone'],  sa_codes)
        sa_ipad    = sacare_units(ed, C6_SA['ipad'],    sa_codes)
        sa_watch   = sacare_units(ed, C6_SA['watch'],   sa_codes)
        sa_mac     = sacare_units(ed, C6_SA['cpu'],     sa_codes)
        sa_airpods = sacare_units(ed, C6_SA['airpods'], sa_codes)

        # ── iPhone 配件 ──
        iphone_prot, _ = accessory_units_revenue(ed, 4009.0, C6_SCREEN_PROT)
        iphone_case, _ = accessory_units_revenue(ed, 4009.0, C6_CASE)
        iphone_lens, _ = accessory_units_revenue(ed, 4009.0, C6_LENS_PROT)

        # ── iPad 配件 ──
        ipad_pencil1 = ipad_pencil_first_units(ed)
        ipad_pencil3 = ipad_pencil_third_units(ed)
        ipad_prot, _ = accessory_units_revenue(ed, C4_IPAD_ACC, C6_SCREEN_PROT)
        ipad_case, _ = accessory_units_revenue(ed, C4_IPAD_ACC, C6_CASE)
        ipad_kb      = ipad_keyboard_units(ed)

        # ── Watch 配件 ──
        watch_prot = watch_screen_prot_units(ed)
        watch_band = watch_band_units(ed)

        # ── AirPods 配件 ──
        apo_acc = airpods_acc_units(ed)

        # ── 喇叭 ──
        spk_with    = speaker_revenue(ed, include_momax=True)
        spk_without = speaker_revenue(ed, include_momax=False)

        results.append({
            'emp_id':       emp_id,
            'name':         emp,
            'iphone_host':  iphone_host,  'acpp_iphone':  acpp_iphone,  'sa_iphone':  sa_iphone,
            'ipad_host':    ipad_host,    'acpp_ipad':    acpp_ipad,    'sa_ipad':    sa_ipad,
            'watch_host':   watch_host,   'acpp_watch':   acpp_watch,   'sa_watch':   sa_watch,
            'mac_host':     mac_host,     'acpp_mac':     acpp_mac,     'sa_mac':     sa_mac,
            'airpods_host': apo_host,     'acpp_airpods': acpp_airpods, 'sa_airpods': sa_airpods,
            'iphone_prot':  iphone_prot,  'iphone_case':  iphone_case,  'iphone_lens': iphone_lens,
            'ipad_pencil1': ipad_pencil1, 'ipad_pencil3': ipad_pencil3,
            'ipad_prot':    ipad_prot,    'ipad_case':    ipad_case,    'ipad_kb':    ipad_kb,
            'watch_prot':   watch_prot,   'watch_band':   watch_band,
            'airpods_acc':  apo_acc,
            'spk_with':     spk_with,     'spk_without':  spk_without,
        })
    return results
