#!/usr/bin/env python3
"""北一區週報產生器 — 本機 Web 工具"""
import json, os, sys, time, threading, traceback, urllib.parse, uuid
from calendar import monthrange
from datetime import date, timedelta
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO
from pathlib import Path

ROOT         = Path(__file__).resolve().parent
STATIC_ROOT  = ROOT / 'static'
TEMPLATE     = ROOT / 'template' / '北一區週報_優化.xlsx'
SA_FILE      = ROOT / 'data' / 'SAcare對應價目表.xlsx'

sys.path.insert(0, str(ROOT))
import multistore_engine as eng
from openpyxl import load_workbook

JOBS = {}
_LOCK = threading.Lock()


# ─── 預設本週結束日（最近一個週六）─────────────────────────────
def _last_saturday() -> date:
    today = date.today()
    days = (today.weekday() + 2) % 7   # Sat→0, Sun→1, Mon→2, …
    return today - timedelta(days=days if days else 7)


# ─── 核心填充邏輯 ─────────────────────────────────────────────
def _fill_workbook(wk_end: date, log) -> bytes:
    from openpyxl import load_workbook as _lw

    WK_START = wk_end - timedelta(days=6)
    PW_END   = WK_START - timedelta(days=1)
    PW_START = PW_END - timedelta(days=6)

    # MTD / PM_SAME（跨月邏輯）
    if WK_START.month != wk_end.month:
        MTD_START = date(WK_START.year, WK_START.month, 1)
        MTD_END   = date(WK_START.year, WK_START.month,
                         monthrange(WK_START.year, WK_START.month)[1])
    else:
        MTD_START = date(wk_end.year, wk_end.month, 1)
        MTD_END   = wk_end

    _pm_y  = MTD_START.year if MTD_START.month > 1 else MTD_START.year - 1
    _pm_m  = MTD_START.month - 1 if MTD_START.month > 1 else 12
    PM_SAME_START = date(_pm_y, _pm_m, 1)
    PM_SAME_END   = date(_pm_y, _pm_m, min(MTD_END.day, monthrange(_pm_y, _pm_m)[1]))
    PM_START = PM_SAME_START
    PM_END   = date(_pm_y, _pm_m, monthrange(_pm_y, _pm_m)[1])

    LYMO_START = date(MTD_START.year - 1, MTD_START.month, 1)
    LYMO_END   = date(LYMO_START.year, LYMO_START.month,
                      min(MTD_END.day, monthrange(LYMO_START.year, LYMO_START.month)[1]))

    LYW_START = WK_START - timedelta(weeks=52)
    LYW_END   = wk_end   - timedelta(weeks=52)
    YTD_S_CY  = date(wk_end.year, 1, 1);  YTD_E_CY = wk_end
    YTD_S_LY  = date(LYMO_END.year, 1, 1); YTD_E_LY = LYMO_END

    STORE_CODES = list(eng.STORES.keys())

    log(f'本週: {WK_START} ~ {wk_end}  |  上週: {PW_START} ~ {PW_END}')
    log(f'MTD: {MTD_START}~{MTD_END}  PM_SAME: {PM_SAME_START}~{PM_SAME_END}')
    log(f'LYMO: {LYMO_START}~{LYMO_END}')

    log('載入今年銷售資料（EPB）…')
    df_cy = eng.load_from_epb(YTD_S_CY, YTD_E_CY, store_codes=STORE_CODES)
    log(f'  今年資料：{len(df_cy):,} 筆')

    ly_end = max(LYW_END, LYMO_END, YTD_E_LY)
    log('載入去年銷售資料（EPB）…')
    df_ly = eng.load_from_epb(YTD_S_LY, ly_end, store_codes=STORE_CODES)
    log(f'  去年資料：{len(df_ly):,} 筆')

    sa_prices = eng.load_sacare_prices(str(SA_FILE))
    sa_codes  = set(sa_prices.keys())

    log('計算各期指標…')

    def rate(num, den):
        return round(num / den, 4) if den else 0.0

    def pct(new, old):
        return round((new - old) / abs(old), 4) if old else 0.0

    PERIODS = {
        'wk':      (df_cy, WK_START,      wk_end),
        'pw':      (df_cy, PW_START,      PW_END),
        'mtd':     (df_cy, MTD_START,     MTD_END),
        'pm_same': (df_cy, PM_SAME_START, PM_SAME_END),
        'pm':      (df_cy, PM_START,      PM_END),
        'lyw':     (df_ly, LYW_START,     LYW_END),
        'lymo':    (df_ly, LYMO_START,    LYMO_END),
        'ytd_cy':  (df_cy, YTD_S_CY,      YTD_E_CY),
        'ytd_ly':  (df_ly, YTD_S_LY,      YTD_E_LY),
    }
    M = {}; A = {}
    rows_stores = STORE_CODES + ['ALL']
    for period, (df, s, e) in PERIODS.items():
        M[period] = {}; A[period] = {}
        for code in rows_stores:
            sc = None if code == 'ALL' else code
            M[period][code] = eng.calc_store_metrics(df, s, e, sc, sa_prices)
            A[period][code] = eng.calc_accessory_by_c4(df, s, e, sc, sa_codes)

    MISC_WK = {}; MISC_PW = {}; MISC_MTD = {}
    for code in rows_stores:
        sc = None if code == 'ALL' else code
        MISC_WK[code]  = eng.calc_misc_metrics(df_cy, WK_START, wk_end,    sc, sa_prices)
        MISC_PW[code]  = eng.calc_misc_metrics(df_cy, PW_START, PW_END,    sc, sa_prices)
        MISC_MTD[code] = eng.calc_misc_metrics(df_cy, MTD_START, MTD_END,  sc, sa_prices)

    log('指標計算完成，載入範本…')
    wb = _lw(str(TEMPLATE))

    # ── helpers ──────────────────────────────────────────────────
    def fill_acc_section(ws, row_start, col_offset, code):
        for i, (c4, _) in enumerate(eng.C4_ACCESSORY):
            r = row_start + i
            pw_v     = A['pw'][code].get(c4, 0)
            wk_v     = A['wk'][code].get(c4, 0)
            mtd_v    = A['mtd'][code].get(c4, 0)
            pm_v     = A['pm'][code].get(c4, 0)
            lyw_v    = A['lyw'][code].get(c4, 0)
            ytd_ly_v = A['ytd_ly'][code].get(c4, 0)
            ytd_cy_v = A['ytd_cy'][code].get(c4, 0)
            ws.cell(r, col_offset+3).value  = pw_v
            ws.cell(r, col_offset+4).value  = wk_v
            ws.cell(r, col_offset+5).value  = wk_v - pw_v
            ws.cell(r, col_offset+6).value  = pct(wk_v, pw_v)
            ws.cell(r, col_offset+7).value  = mtd_v
            ws.cell(r, col_offset+8).value  = pm_v
            ws.cell(r, col_offset+9).value  = mtd_v - pm_v
            ws.cell(r, col_offset+10).value = pct(mtd_v, pm_v)
            ws.cell(r, col_offset+11).value = lyw_v
            ws.cell(r, col_offset+12).value = wk_v - lyw_v
            ws.cell(r, col_offset+13).value = pct(wk_v, lyw_v)
            ws.cell(r, col_offset+14).value = ytd_ly_v
            ws.cell(r, col_offset+15).value = ytd_cy_v
            ws.cell(r, col_offset+16).value = ytd_cy_v - ytd_ly_v
            ws.cell(r, col_offset+17).value = pct(ytd_cy_v, ytd_ly_v)
        total_r = row_start + len(eng.C4_ACCESSORY)
        for off, p in [(3,'pw'),(4,'wk'),(7,'mtd'),(8,'pm'),
                       (11,'lyw'),(14,'ytd_ly'),(15,'ytd_cy')]:
            ws.cell(total_r, col_offset+off).value = sum(A[p][code].values())
        pw_t = sum(A['pw'][code].values());  wk_t = sum(A['wk'][code].values())
        mtd_t= sum(A['mtd'][code].values()); pm_t = sum(A['pm'][code].values())
        lyw_t= sum(A['lyw'][code].values())
        ly_t = sum(A['ytd_ly'][code].values()); cy_t = sum(A['ytd_cy'][code].values())
        ws.cell(total_r, col_offset+5).value  = wk_t - pw_t
        ws.cell(total_r, col_offset+6).value  = pct(wk_t, pw_t)
        ws.cell(total_r, col_offset+9).value  = mtd_t - pm_t
        ws.cell(total_r, col_offset+10).value = pct(mtd_t, pm_t)
        ws.cell(total_r, col_offset+12).value = wk_t - lyw_t
        ws.cell(total_r, col_offset+13).value = pct(wk_t, lyw_t)
        ws.cell(total_r, col_offset+16).value = cy_t - ly_t
        ws.cell(total_r, col_offset+17).value = pct(cy_t, ly_t)

    def fill_acpp(ws, row_start, period):
        for ri, code in enumerate(rows_stores):
            r = row_start + ri
            m = M[period][code]
            ws.cell(r, 1).value = eng.STORES.get(code, 'Total')
            ws.cell(r, 2).value = m['cpu_units']
            ws.cell(r, 3).value = m['acpp_mac']
            ws.cell(r, 4).value = rate(m['acpp_mac'], m['cpu_units'])
            ws.cell(r, 5).value = m['sa_cpu']
            ws.cell(r, 6).value = rate(m['sa_cpu'], m['cpu_units'])
            ws.cell(r, 7).value = rate(m['acpp_mac']+m['sa_cpu'], m['cpu_units'])
            ws.cell(r, 8).value = m['watch_units']
            ws.cell(r, 9).value = m['acpp_watch']
            ws.cell(r,10).value = rate(m['acpp_watch'], m['watch_units'])
            ws.cell(r,11).value = m['sa_watch']
            ws.cell(r,12).value = rate(m['sa_watch'], m['watch_units'])
            ws.cell(r,13).value = rate(m['acpp_watch']+m['sa_watch'], m['watch_units'])
            ws.cell(r,14).value = m['ipad_units']
            ws.cell(r,15).value = m['acpp_ipad']
            ws.cell(r,16).value = rate(m['acpp_ipad'], m['ipad_units'])
            ws.cell(r,17).value = m['sa_ipad']
            ws.cell(r,18).value = rate(m['sa_ipad'], m['ipad_units'])
            ws.cell(r,19).value = rate(m['acpp_ipad']+m['sa_ipad'], m['ipad_units'])
            ws.cell(r,20).value = m['iphone_units']
            ws.cell(r,21).value = m['acpp_iphone']
            ws.cell(r,22).value = rate(m['acpp_iphone'], m['iphone_units'])
            ws.cell(r,23).value = m['sa_iphone']
            ws.cell(r,24).value = rate(m['sa_iphone'], m['iphone_units'])
            ws.cell(r,25).value = rate(m['acpp_iphone']+m['sa_iphone'], m['iphone_units'])
            ios_d = m['iphone_units']+m['ipad_units']
            ios_n = m['acpp_iphone']+m['acpp_ipad']+m['sa_iphone']+m['sa_ipad']
            ws.cell(r,26).value = rate(ios_n, ios_d)
            ws.cell(r,27).value = m['airpods_units']
            ws.cell(r,28).value = m['acpp_airpods']
            ws.cell(r,29).value = rate(m['acpp_airpods'], m['airpods_units'])
            ws.cell(r,30).value = m['sa_airpods']
            ws.cell(r,31).value = rate(m['sa_airpods'], m['airpods_units'])
            ws.cell(r,32).value = rate(m['acpp_airpods']+m['sa_airpods'], m['airpods_units'])

    def fill_units(ws, row_start, pa, pb):
        for ri, code in enumerate(rows_stores):
            r = row_start + ri
            ma = M[pa][code]; mb = M[pb][code]
            ws.cell(r, 1).value = eng.STORES.get(code, 'Total')
            ws.cell(r, 2).value = ma['cpu_units']
            ws.cell(r, 3).value = mb['cpu_units']
            ws.cell(r, 4).value = pct(mb['cpu_units'], ma['cpu_units'])
            ws.cell(r, 5).value = ma['ipad_units']
            ws.cell(r, 6).value = mb['ipad_units']
            ws.cell(r, 7).value = pct(mb['ipad_units'], ma['ipad_units'])
            ws.cell(r, 8).value = ma['iphone_units']
            ws.cell(r, 9).value = mb['iphone_units']
            ws.cell(r,10).value = pct(mb['iphone_units'], ma['iphone_units'])
            ws.cell(r,11).value = ma['watch_units']
            ws.cell(r,12).value = mb['watch_units']
            ws.cell(r,13).value = pct(mb['watch_units'], ma['watch_units'])
            ws.cell(r,14).value = ma['airpods_units']
            ws.cell(r,15).value = mb['airpods_units']
            ws.cell(r,16).value = pct(mb['airpods_units'], ma['airpods_units'])
            ws.cell(r,18).value = ma['txn_count']
            ws.cell(r,21).value = mb['txn_count']
            avg_a = rate(ma['total_excl_sa'], ma['txn_count'])
            avg_b = rate(mb['total_excl_sa'], mb['txn_count'])
            ws.cell(r,26).value = avg_a
            ws.cell(r,27).value = avg_b
            ws.cell(r,28).value = avg_b - avg_a
            ws.cell(r,29).value = pct(avg_b, avg_a)

    def fill_biz(ws, row_start, pa, pb):
        for ri, code in enumerate(rows_stores):
            r = row_start + ri
            a = M[pa][code]; b = M[pb][code]
            ws.cell(r, 1).value = eng.STORES.get(code, 'Total')
            ws.cell(r, 2).value = a['total_excl_sa']
            ws.cell(r, 3).value = b['total_excl_sa']
            ws.cell(r, 4).value = pct(b['total_excl_sa'], a['total_excl_sa'])
            ws.cell(r, 5).value = a['tpp_excl_sa']
            ws.cell(r, 6).value = b['tpp_excl_sa']
            ws.cell(r, 7).value = pct(b['tpp_excl_sa'], a['tpp_excl_sa'])
            ws.cell(r, 8).value = a['sa_rev']
            ws.cell(r, 9).value = b['sa_rev']
            ws.cell(r,10).value = pct(b['sa_rev'], a['sa_rev'])
            ws.cell(r,11).value = a['acpp_plus']
            ws.cell(r,12).value = b['acpp_plus']
            ws.cell(r,13).value = pct(b['acpp_plus'], a['acpp_plus'])
            ws.cell(r,14).value = a['total_rev']
            ws.cell(r,15).value = b['total_rev']
            ws.cell(r,16).value = pct(b['total_rev'], a['total_rev'])
            ws.cell(r,17).value = a['tpp_rev']
            ws.cell(r,18).value = b['tpp_rev']
            ws.cell(r,19).value = pct(b['tpp_rev'], a['tpp_rev'])
            ws.cell(r,20).value = b['total_excl_sa'] - a['total_excl_sa']
            ws.cell(r,21).value = b['tpp_excl_sa']   - a['tpp_excl_sa']
            ws.cell(r,22).value = b['sa_rev']         - a['sa_rev']
            ws.cell(r,23).value = b['acpp_plus']      - a['acpp_plus']
            ws.cell(r,24).value = rate(a['tpp_excl_sa'], a['total_excl_sa'])
            ws.cell(r,25).value = rate(b['tpp_excl_sa'], b['total_excl_sa'])
            ws.cell(r,26).value = rate(a['sa_rev'],      a['total_excl_sa'])
            ws.cell(r,27).value = rate(b['sa_rev'],      b['total_excl_sa'])
            ws.cell(r,28).value = a['coupon_rev']
            ws.cell(r,29).value = b['coupon_rev']
            ws.cell(r,30).value = pct(b['coupon_rev'], a['coupon_rev'])

    def fill_biz_mo(ws, row_start, pa, pb):
        """月累積版（多兩個手動欄 col2-3）"""
        for ri, code in enumerate(rows_stores):
            r = row_start + ri
            a = M[pa][code]; b = M[pb][code]
            ws.cell(r, 1).value = eng.STORES.get(code, 'Total')
            ws.cell(r, 4).value = a['total_excl_sa']
            ws.cell(r, 5).value = b['total_excl_sa']
            ws.cell(r, 6).value = pct(b['total_excl_sa'], a['total_excl_sa'])
            ws.cell(r, 7).value = a['tpp_excl_sa']
            ws.cell(r, 8).value = b['tpp_excl_sa']
            ws.cell(r, 9).value = pct(b['tpp_excl_sa'], a['tpp_excl_sa'])
            ws.cell(r,10).value = a['sa_rev']
            ws.cell(r,11).value = b['sa_rev']
            ws.cell(r,12).value = pct(b['sa_rev'], a['sa_rev'])
            ws.cell(r,13).value = a['acpp_plus']
            ws.cell(r,14).value = b['acpp_plus']
            ws.cell(r,15).value = pct(b['acpp_plus'], a['acpp_plus'])
            ws.cell(r,16).value = a['total_rev']
            ws.cell(r,17).value = b['total_rev']
            ws.cell(r,18).value = pct(b['total_rev'], a['total_rev'])
            ws.cell(r,19).value = a['tpp_rev']
            ws.cell(r,20).value = b['tpp_rev']
            ws.cell(r,21).value = pct(b['tpp_rev'], a['tpp_rev'])
            ws.cell(r,22).value = b['total_excl_sa'] - a['total_excl_sa']
            ws.cell(r,23).value = b['tpp_excl_sa']   - a['tpp_excl_sa']
            ws.cell(r,26).value = rate(a['tpp_excl_sa'], a['total_excl_sa'])
            ws.cell(r,27).value = rate(b['tpp_excl_sa'], b['total_excl_sa'])
            ws.cell(r,28).value = rate(a['sa_rev'],      a['total_excl_sa'])
            ws.cell(r,29).value = rate(b['sa_rev'],      b['total_excl_sa'])
            ws.cell(r,30).value = a['coupon_rev']
            ws.cell(r,31).value = b['coupon_rev']
            ws.cell(r,32).value = pct(b['coupon_rev'], a['coupon_rev'])

    # ── 配件 sheets ──────────────────────────────────────────────
    log('填入配件 sheets…')
    ws_sum = wb['配件-北一區 (匯總)']
    fill_acc_section(ws_sum, 4, 0, 'ALL')
    fill_acc_section(ws_sum, 23, 0,  '004'); fill_acc_section(ws_sum, 23, 17, '054')
    fill_acc_section(ws_sum, 42, 0,  '024'); fill_acc_section(ws_sum, 42, 17, '046')
    fill_acc_section(ws_sum, 61, 0,  '005'); fill_acc_section(ws_sum, 61, 17, '057')
    fill_acc_section(wb['配件-北一區'], 4, 0, 'ALL')
    for code, sname in [('004','配件 - 士林門市'), ('005','配件 - 微風門市'),
                        ('024','配件 - 美麗華門市'), ('046','配件 - 阿波羅門市'),
                        ('054','配件 - 大葉高島屋門市'), ('057','配件 - 羅東門市')]:
        fill_acc_section(wb[sname], 4, 0, code)

    # ── BY店 本週比較 ─────────────────────────────────────────────
    log('填入 BY店 本週比較…')
    ws_wk = wb['BY店 本週比較']
    fill_biz(ws_wk, 4, 'pw', 'wk')
    fill_units(ws_wk, 14, 'pw', 'wk')
    fill_acpp(ws_wk, 23, 'wk');  fill_acpp(ws_wk, 32, 'pw')

    # ── BY店 月累積 ───────────────────────────────────────────────
    log('填入 BY店 月累積…')
    ws_mo = wb['BY店 月累積']
    fill_biz_mo(ws_mo, 4, 'pm_same', 'mtd')
    fill_units(ws_mo, 14, 'pm_same', 'mtd')
    fill_acpp(ws_mo, 23, 'mtd');  fill_acpp(ws_mo, 32, 'pm_same')

    # ── BY店 去年同期 ─────────────────────────────────────────────
    log('填入 BY店 去年同期…')
    ws_ly = wb['BY店 去年同期']
    fill_biz(ws_ly, 4, 'lymo', 'mtd')
    fill_units(ws_ly, 14, 'lymo', 'mtd')
    fill_acpp(ws_ly, 23, 'mtd');  fill_acpp(ws_ly, 32, 'lymo')

    # ── BY店 整年同期 ─────────────────────────────────────────────
    log('填入 BY店 整年同期…')
    ws_yoy = wb['BY店 整年同期']
    fill_biz(ws_yoy, 4, 'ytd_ly', 'ytd_cy')
    fill_units(ws_yoy, 14, 'ytd_ly', 'ytd_cy')
    fill_acpp(ws_yoy, 23, 'ytd_cy');  fill_acpp(ws_yoy, 32, 'ytd_ly')

    # ── BY店 本週其他細項 ─────────────────────────────────────────
    log('填入 BY店 本週其他細項…')
    ws_misc = wb['BY店 本週其他細項']
    for ri, code in enumerate(rows_stores):
        r = 3 + ri; m = MISC_WK[code]
        ws_misc.cell(r,  2).value = m['sa_mac'];      ws_misc.cell(r,  3).value = m['sa_iphone']
        ws_misc.cell(r,  4).value = m['sa_ipad'];     ws_misc.cell(r,  5).value = m['sa_watch']
        ws_misc.cell(r,  6).value = m['sa_airpods'];  ws_misc.cell(r,  7).value = m['sa_total']
        ws_misc.cell(r, 10).value = m['coupon_give']; ws_misc.cell(r, 11).value = m['coupon_redeem']
        ws_misc.cell(r, 12).value = rate(m['coupon_redeem'], m['coupon_give'])
        ws_misc.cell(r, 13).value = m['eco'];         ws_misc.cell(r, 14).value = m['cable']
        ws_misc.cell(r, 17).value = m['host_total']
    for ri, code in enumerate(rows_stores):
        r = 13 + ri; pw_m = MISC_PW[code]; wk_m = MISC_WK[code]
        ws_misc.cell(r, 2).value = pw_m['spk_with'];   ws_misc.cell(r, 3).value = wk_m['spk_with']
        ws_misc.cell(r, 4).value = pct(wk_m['spk_with'], pw_m['spk_with'])
        ws_misc.cell(r, 5).value = wk_m['spk_with'] - pw_m['spk_with']
        ws_misc.cell(r, 6).value = pw_m['spk_without']; ws_misc.cell(r, 7).value = wk_m['spk_without']
        ws_misc.cell(r, 8).value = pct(wk_m['spk_without'], pw_m['spk_without'])
        ws_misc.cell(r, 9).value = wk_m['spk_without'] - pw_m['spk_without']
    for ri, code in enumerate(rows_stores):
        r = 23 + ri; pw_m = MISC_PW[code]; wk_m = MISC_WK[code]
        ws_misc.cell(r,  2).value = pw_m['iphone_host'];     ws_misc.cell(r, 12).value = wk_m['iphone_host']
        ws_misc.cell(r,  3).value = pw_m['iphone_prot_qty']; ws_misc.cell(r, 13).value = wk_m['iphone_prot_qty']
        ws_misc.cell(r,  4).value = pw_m['iphone_case_qty']; ws_misc.cell(r, 14).value = wk_m['iphone_case_qty']
        ws_misc.cell(r,  5).value = pw_m['iphone_lens_qty']; ws_misc.cell(r, 15).value = wk_m['iphone_lens_qty']
        ws_misc.cell(r,  6).value = pw_m['iphone_prot_rev']; ws_misc.cell(r, 16).value = wk_m['iphone_prot_rev']
        ws_misc.cell(r,  7).value = pw_m['iphone_case_rev']; ws_misc.cell(r, 17).value = wk_m['iphone_case_rev']
        ws_misc.cell(r,  8).value = pw_m['iphone_lens_rev']; ws_misc.cell(r, 18).value = wk_m['iphone_lens_rev']
        ws_misc.cell(r,  9).value = rate(pw_m['iphone_prot_qty'], pw_m['iphone_host'])
        ws_misc.cell(r, 10).value = rate(pw_m['iphone_case_qty'], pw_m['iphone_host'])
        ws_misc.cell(r, 11).value = rate(pw_m['iphone_lens_qty'], pw_m['iphone_host'])
        ws_misc.cell(r, 19).value = rate(wk_m['iphone_prot_qty'], wk_m['iphone_host'])
        ws_misc.cell(r, 20).value = rate(wk_m['iphone_case_qty'], wk_m['iphone_host'])
        ws_misc.cell(r, 21).value = rate(wk_m['iphone_lens_qty'], wk_m['iphone_host'])
        ws_misc.cell(r, 22).value = wk_m['iphone_prot_qty']-pw_m['iphone_prot_qty']
        ws_misc.cell(r, 23).value = wk_m['iphone_case_qty']-pw_m['iphone_case_qty']
        ws_misc.cell(r, 24).value = wk_m['iphone_lens_qty']-pw_m['iphone_lens_qty']
        ws_misc.cell(r, 25).value = wk_m['iphone_prot_rev']-pw_m['iphone_prot_rev']
        ws_misc.cell(r, 26).value = wk_m['iphone_case_rev']-pw_m['iphone_case_rev']
        ws_misc.cell(r, 27).value = wk_m['iphone_lens_rev']-pw_m['iphone_lens_rev']
    for ri, code in enumerate(rows_stores):
        r = 33 + ri; pw_m = MISC_PW[code]; wk_m = MISC_WK[code]
        ws_misc.cell(r,  2).value = pw_m['ipad_host'];  ws_misc.cell(r, 13).value = wk_m['ipad_host']
        ws_misc.cell(r,  3).value = pw_m['ipad_pencil1']; ws_misc.cell(r, 14).value = wk_m['ipad_pencil1']
        ws_misc.cell(r,  4).value = rate(pw_m['ipad_pencil1'], pw_m['ipad_host'])
        ws_misc.cell(r,  5).value = pw_m['ipad_pencil3']; ws_misc.cell(r, 16).value = wk_m['ipad_pencil3']
        ws_misc.cell(r,  6).value = rate(pw_m['ipad_pencil3'], pw_m['ipad_host'])
        ws_misc.cell(r,  7).value = pw_m['ipad_prot_qty']; ws_misc.cell(r, 18).value = wk_m['ipad_prot_qty']
        ws_misc.cell(r,  8).value = rate(pw_m['ipad_prot_qty'], pw_m['ipad_host'])
        ws_misc.cell(r,  9).value = pw_m['ipad_case_qty']; ws_misc.cell(r, 20).value = wk_m['ipad_case_qty']
        ws_misc.cell(r, 10).value = rate(pw_m['ipad_case_qty'], pw_m['ipad_host'])
        ws_misc.cell(r, 11).value = pw_m['ipad_kb'];    ws_misc.cell(r, 22).value = wk_m['ipad_kb']
        ws_misc.cell(r, 12).value = rate(pw_m['ipad_kb'], pw_m['ipad_host'])
        ws_misc.cell(r, 15).value = rate(wk_m['ipad_pencil1'], wk_m['ipad_host'])
        ws_misc.cell(r, 17).value = rate(wk_m['ipad_pencil3'], wk_m['ipad_host'])
        ws_misc.cell(r, 19).value = rate(wk_m['ipad_prot_qty'], wk_m['ipad_host'])
        ws_misc.cell(r, 21).value = rate(wk_m['ipad_case_qty'], wk_m['ipad_host'])
        ws_misc.cell(r, 23).value = rate(wk_m['ipad_kb'], wk_m['ipad_host'])
    for ri, code in enumerate(rows_stores):
        r = 43 + ri; pw_m = MISC_PW[code]; wk_m = MISC_WK[code]
        ws_misc.cell(r, 2).value = pw_m['watch_host'];  ws_misc.cell(r, 7).value = wk_m['watch_host']
        ws_misc.cell(r, 3).value = pw_m['watch_prot'];  ws_misc.cell(r, 8).value = wk_m['watch_prot']
        ws_misc.cell(r, 4).value = rate(pw_m['watch_prot'], pw_m['watch_host'])
        ws_misc.cell(r, 5).value = pw_m['watch_band'];  ws_misc.cell(r,10).value = wk_m['watch_band']
        ws_misc.cell(r, 6).value = rate(pw_m['watch_band'], pw_m['watch_host'])
        ws_misc.cell(r, 9).value = rate(wk_m['watch_prot'], wk_m['watch_host'])
        ws_misc.cell(r,11).value = rate(wk_m['watch_band'], wk_m['watch_host'])

    # ── BY店 人員銷售 ─────────────────────────────────────────────
    log('填入 BY店 人員銷售…')
    from copy import copy as _copy
    from openpyxl.utils import get_column_letter as _gcl
    ws_staff = wb['BY店 人員銷售']
    STAFF_BLOCKS = [('004',1),('005',23),('024',44),('046',67),('054',93),('057',116)]

    def _insert_rows_fix(ws, idx, amount):
        to_fix = []
        for mr in list(ws.merged_cells.ranges):
            if mr.min_row >= idx:
                to_fix.append((mr.min_row, mr.max_row, mr.min_col, mr.max_col))
        ws.insert_rows(idx, amount)
        for (min_r, max_r, min_c, max_c) in to_fix:
            old = f'{_gcl(min_c)}{min_r}:{_gcl(max_c)}{max_r}'
            try:
                ws.unmerge_cells(old)
            except Exception:
                pass
            ws.merge_cells(start_row=min_r+amount, start_column=min_c,
                           end_row=max_r+amount, end_column=max_c)

    _SKIP_LABELS = {
        '員工代碼', '員工名稱', '本週', '月累積', '小計',
        'iPhone台數', 'iPad台數', 'Watch台數', 'Mac台數', 'AirPods台數',
        'iPhone', 'iPad', 'Watch', 'Mac', 'AirPods', 'ACPP+', 'SAcare', '合計率',
    }

    def _find_blank_sec(ws, ds):
        slots, sub = [], None
        for r in range(ds, ds + 50):
            v1 = str(ws.cell(r, 1).value or '').strip()
            v2 = str(ws.cell(r, 2).value or '').strip()
            c3  = ws.cell(r, 3).value
            if v1 == '小計' or v2 == '小計':
                sub = r; break
            if v1 in _SKIP_LABELS or v2 in _SKIP_LABELS:
                continue
            if isinstance(c3, str) and c3.strip():   # 機型/欄位標頭行
                continue
            slots.append(r)
        return slots, sub

    def _copy_row_format(ws, src, dst):
        """將 src 列的格式複製到 dst 列（不複製值）。"""
        for c in range(1, ws.max_column + 1):
            sc_ = ws.cell(src, c); dc = ws.cell(dst, c)
            if sc_.has_style:
                dc.font        = _copy(sc_.font)
                dc.fill        = _copy(sc_.fill)
                dc.alignment   = _copy(sc_.alignment)
                dc.border      = _copy(sc_.border)
                dc.number_format = sc_.number_format
            dc.value = None
        src_h = ws.row_dimensions.get(src)
        if src_h and src_h.height:
            ws.row_dimensions[dst].height = src_h.height

    def _fill_person(ws, r, p):
        # col1=員工代碼  col2=員工名稱  col3~46=資料（新範本欄位佈局）
        ws.cell(r, 1).value = p.get('emp_id', '')
        ws.cell(r, 2).value = p.get('name', '')
        g = p.get
        # iPhone
        ws.cell(r, 3).value = g('iphone_host',0)
        ws.cell(r, 4).value = g('acpp_iphone',0)
        ws.cell(r, 5).value = g('sa_iphone',0)
        ws.cell(r, 6).value = rate(g('acpp_iphone',0)+g('sa_iphone',0), g('iphone_host',0))
        # iPad
        ws.cell(r, 7).value = g('ipad_host',0)
        ws.cell(r, 8).value = g('acpp_ipad',0)
        ws.cell(r, 9).value = g('sa_ipad',0)
        ws.cell(r,10).value = rate(g('acpp_ipad',0)+g('sa_ipad',0), g('ipad_host',0))
        # Watch
        ws.cell(r,11).value = g('watch_host',0)
        ws.cell(r,12).value = g('acpp_watch',0)
        ws.cell(r,13).value = g('sa_watch',0)
        ws.cell(r,14).value = rate(g('acpp_watch',0)+g('sa_watch',0), g('watch_host',0))
        # Mac
        ws.cell(r,15).value = g('mac_host',0)
        ws.cell(r,16).value = g('acpp_mac',0)
        ws.cell(r,17).value = g('sa_mac',0)
        ws.cell(r,18).value = rate(g('acpp_mac',0)+g('sa_mac',0), g('mac_host',0))
        # AirPods
        ws.cell(r,19).value = g('airpods_host',0)
        ws.cell(r,20).value = g('acpp_airpods',0)
        ws.cell(r,21).value = g('sa_airpods',0)
        ws.cell(r,22).value = rate(g('acpp_airpods',0)+g('sa_airpods',0), g('airpods_host',0))
        # iPhone 配件
        ws.cell(r,23).value = g('iphone_prot',0)
        ws.cell(r,24).value = g('iphone_case',0)
        ws.cell(r,25).value = g('iphone_lens',0)
        ws.cell(r,26).value = rate(g('iphone_prot',0), g('iphone_host',0))
        ws.cell(r,27).value = rate(g('iphone_case',0), g('iphone_host',0))
        ws.cell(r,28).value = rate(g('iphone_lens',0), g('iphone_host',0))
        # iPad 配件
        ws.cell(r,29).value = g('ipad_pencil1',0)
        ws.cell(r,30).value = g('ipad_pencil3',0)
        ws.cell(r,31).value = g('ipad_prot',0)
        ws.cell(r,32).value = g('ipad_case',0)
        ws.cell(r,33).value = g('ipad_kb',0)
        ws.cell(r,34).value = rate(g('ipad_pencil1',0), g('ipad_host',0))
        ws.cell(r,35).value = rate(g('ipad_pencil3',0), g('ipad_host',0))
        ws.cell(r,36).value = rate(g('ipad_prot',0),    g('ipad_host',0))
        ws.cell(r,37).value = rate(g('ipad_case',0),    g('ipad_host',0))
        ws.cell(r,38).value = rate(g('ipad_kb',0),      g('ipad_host',0))
        # Watch 配件
        ws.cell(r,39).value = g('watch_prot',0)
        ws.cell(r,40).value = g('watch_band',0)
        ws.cell(r,41).value = rate(g('watch_prot',0), g('watch_host',0))
        ws.cell(r,42).value = rate(g('watch_band',0), g('watch_host',0))
        # AirPods 配件
        ws.cell(r,43).value = g('airpods_acc',0)
        ws.cell(r,44).value = rate(g('airpods_acc',0), g('airpods_host',0))
        # 喇叭
        ws.cell(r,45).value = g('spk_with',0)
        ws.cell(r,46).value = g('spk_without',0)

    def _fill_sub(ws, r, m, misc):
        def g(k): return m.get(k) or 0
        ws.cell(r, 3).value = g('iphone_units')
        ws.cell(r, 4).value = g('acpp_iphone')
        ws.cell(r, 5).value = g('sa_iphone')
        ws.cell(r, 6).value = rate(g('acpp_iphone')+g('sa_iphone'), g('iphone_units'))
        ws.cell(r, 7).value = g('ipad_units')
        ws.cell(r, 8).value = g('acpp_ipad')
        ws.cell(r, 9).value = g('sa_ipad')
        ws.cell(r,10).value = rate(g('acpp_ipad')+g('sa_ipad'), g('ipad_units'))
        ws.cell(r,11).value = g('watch_units')
        ws.cell(r,12).value = g('acpp_watch')
        ws.cell(r,13).value = g('sa_watch')
        ws.cell(r,14).value = rate(g('acpp_watch')+g('sa_watch'), g('watch_units'))
        ws.cell(r,15).value = g('cpu_units')
        ws.cell(r,16).value = g('acpp_mac')
        ws.cell(r,17).value = g('sa_cpu')
        ws.cell(r,18).value = rate(g('acpp_mac')+g('sa_cpu'), g('cpu_units'))
        ws.cell(r,19).value = g('airpods_units')
        ws.cell(r,20).value = g('acpp_airpods')
        ws.cell(r,21).value = g('sa_airpods')
        ws.cell(r,22).value = rate(g('acpp_airpods')+g('sa_airpods'), g('airpods_units'))
        ws.cell(r,23).value = misc['iphone_prot_qty']
        ws.cell(r,24).value = misc['iphone_case_qty']
        ws.cell(r,25).value = misc['iphone_lens_qty']
        ws.cell(r,26).value = rate(misc['iphone_prot_qty'], m['iphone_units'])
        ws.cell(r,27).value = rate(misc['iphone_case_qty'], m['iphone_units'])
        ws.cell(r,28).value = rate(misc['iphone_lens_qty'], m['iphone_units'])
        ws.cell(r,29).value = misc['ipad_pencil1']
        ws.cell(r,30).value = misc['ipad_pencil3']
        ws.cell(r,31).value = misc['ipad_prot_qty']
        ws.cell(r,32).value = misc['ipad_case_qty']
        ws.cell(r,33).value = misc['ipad_kb']
        ws.cell(r,34).value = rate(misc['ipad_pencil1'],  m['ipad_units'])
        ws.cell(r,35).value = rate(misc['ipad_pencil3'],  m['ipad_units'])
        ws.cell(r,36).value = rate(misc['ipad_prot_qty'], m['ipad_units'])
        ws.cell(r,37).value = rate(misc['ipad_case_qty'], m['ipad_units'])
        ws.cell(r,38).value = rate(misc['ipad_kb'],       m['ipad_units'])
        ws.cell(r,39).value = misc['watch_prot']
        ws.cell(r,40).value = misc['watch_band']
        ws.cell(r,41).value = rate(misc['watch_prot'], m['watch_units'])
        ws.cell(r,42).value = rate(misc['watch_band'], m['watch_units'])
        ws.cell(r,45).value = misc.get('spk_with', 0)
        ws.cell(r,46).value = misc.get('spk_without', 0)

    def _fill_section(ws, ds, persons, m, misc):
        """填一個區段（本週或月累積）。回傳 (插入列數, 小計列號)。"""
        slots, sub = _find_blank_sec(ws, ds)
        if sub is None:
            return 0, ds
        inserted = 0
        if len(persons) > len(slots):
            n_extra = len(persons) - len(slots)
            ref = slots[-1] if slots else ds
            _insert_rows_fix(ws, sub, n_extra)
            for i in range(n_extra):
                _copy_row_format(ws, ref, sub + i)
                slots.append(sub + i)
            sub += n_extra
            inserted = n_extra
        for i, p in enumerate(persons):
            if i < len(slots):
                _fill_person(ws, slots[i], p)
        for r in slots[len(persons):]:
            for c in range(1, 47):
                ws.cell(r, c).value = None
        _fill_sub(ws, sub, m, misc)
        return inserted, sub

    row_offset = 0
    for sc, bs in STAFF_BLOCKS:
        actual_bs = bs + row_offset
        wk_persons  = eng.calc_person_metrics(df_cy, WK_START, wk_end,   sc, sa_prices)
        mtd_persons = eng.calc_person_metrics(df_cy, MTD_START, MTD_END, sc, sa_prices)
        log(f'  [{sc}] 本週 {len(wk_persons)} 人 / 月累積 {len(mtd_persons)} 人')
        ins_wk,  wk_sub  = _fill_section(ws_staff, actual_bs + 1, wk_persons,  M['wk'][sc],  MISC_WK[sc])
        ins_mtd, mtd_sub = _fill_section(ws_staff, wk_sub + 1,    mtd_persons, M['mtd'][sc], MISC_MTD[sc])
        row_offset += ins_wk + ins_mtd

    # ── 日期標題 ──────────────────────────────────────────────────
    log('更新日期標題…')

    def _d(d): return f'{d.month:02d}/{d.day:02d}'

    def _set_acc_dates(ws, off):
        ws.cell(3, off+3).value  = f'{_d(PW_START)}~{_d(PW_END)}'
        ws.cell(3, off+4).value  = f'{_d(WK_START)}~{_d(wk_end)}'
        ws.cell(3, off+7).value  = f'{_d(MTD_START)}~{_d(MTD_END)}'
        ws.cell(3, off+8).value  = f'{_d(PM_START)}~{_d(PM_END)}'
        ws.cell(3, off+11).value = f'{LYW_END.year}/{_d(LYW_START)}~{_d(LYW_END)}'
        ws.cell(3, off+14).value = f'{YTD_S_LY.year}/{_d(YTD_S_LY)}~{_d(YTD_E_LY)}'
        ws.cell(3, off+15).value = f'{YTD_S_CY.year}/{_d(YTD_S_CY)}~{_d(YTD_E_CY)}'

    for sname, label in [('配件-北一區 (匯總)','北一區'), ('配件-北一區','北一區'),
                          ('配件 - 士林門市','士林門市'), ('配件 - 微風門市','微風門市'),
                          ('配件 - 美麗華門市','美麗華門市'), ('配件 - 阿波羅門市','阿波羅門市'),
                          ('配件 - 大葉高島屋門市','大葉高島屋門市'), ('配件 - 羅東門市','羅東門市')]:
        if sname not in wb.sheetnames: continue
        ws_d = wb[sname]
        ws_d.cell(1,1).value = f'配件銷售分析  ·  {label}  ·  {WK_START.year}/{_d(WK_START)} ~ {_d(wk_end)}'
        _set_acc_dates(ws_d, 0)
        if sname == '配件-北一區 (匯總)': _set_acc_dates(ws_d, 17)

    ws_misc.cell(1,1).value = f'本週其他細項  ·  本週 {_d(WK_START)}~{_d(wk_end)}  |  對照 上週 {_d(PW_START)}~{_d(PW_END)}'

    def _rep_w(ws, row, la, lb):
        cnt = 0
        for c in range(1, ws.max_column+1):
            v = ws.cell(row, c).value
            if isinstance(v, str) and len(v) >= 2 and v[0]=='W' and v[1:].isdigit():
                cnt += 1
                ws.cell(row, c).value = la if cnt%2==1 else lb

    pw_lbl = f'{_d(PW_START)}~{_d(PW_END)}';  wk_lbl = f'{_d(WK_START)}~{_d(wk_end)}'
    _rep_w(ws_wk, 3, pw_lbl, wk_lbl);  _rep_w(ws_wk, 13, pw_lbl, wk_lbl)
    ws_wk.cell(22,1).value = wk_lbl;   ws_wk.cell(31,1).value = pw_lbl

    pm_s_lbl = f'{_d(PM_SAME_START)}~{_d(PM_SAME_END)}';  mtd_lbl = f'{_d(MTD_START)}~{_d(MTD_END)}'
    ws_mo.cell(1,1).value = f'月累積對照  ·  {MTD_START.year}/{mtd_lbl}  vs  {PM_SAME_START.year}/{pm_s_lbl}'
    for c in range(1, ws_mo.max_column+1):
        v = ws_mo.cell(3, c).value
        if not isinstance(v, str) or '\n' not in v: continue
        base = v.rsplit('\n', 1)[0]
        if c in {4,7,10,13,16,19}:   ws_mo.cell(3,c).value = base+'\n'+pm_s_lbl
        elif c in {5,8,11,14,17,20}: ws_mo.cell(3,c).value = base+'\n'+mtd_lbl
    pm_s_ymd=f'{PM_SAME_START.year}\n{pm_s_lbl}'; mtd_ymd=f'{MTD_START.year}\n{mtd_lbl}'
    for c in range(1, ws_mo.max_column+1):
        if ws_mo.cell(13,c).value is None: continue
        if c in {2,5,8,11,14}:  ws_mo.cell(13,c).value = pm_s_ymd
        elif c in {3,6,9,12,15}: ws_mo.cell(13,c).value = mtd_ymd
    ws_mo.cell(22,1).value = f'{MTD_START.year}\n{mtd_lbl}'
    ws_mo.cell(31,1).value = f'{PM_SAME_START.year}\n{pm_s_lbl}'

    ly_lbl = f'{_d(LYMO_START)}~{_d(LYMO_END)}'
    ws_ly.cell(1,1).value = f'去年同期對照  ·  本月 {MTD_START.year}/{mtd_lbl}  vs  去年同期 {LYMO_START.year}/{ly_lbl}'
    for c in range(1, ws_ly.max_column+1):
        v = ws_ly.cell(3,c).value
        if not isinstance(v, str) or '\n' not in v: continue
        parts = v.split('\n')
        if len(parts) >= 3:
            base = '\n'.join(parts[:-2])
            if c in {2,5,8,11,14,17}:   ws_ly.cell(3,c).value = f'{base}\n{LYMO_START.year}\n{ly_lbl}'
            elif c in {3,6,9,12,15,18}: ws_ly.cell(3,c).value = f'{base}\n{MTD_START.year}\n{mtd_lbl}'
    ly_ymd=f'{LYMO_START.year}\n{ly_lbl}'; cy_ymd=f'{MTD_START.year}\n{mtd_lbl}'
    for c in range(1, ws_ly.max_column+1):
        if ws_ly.cell(13,c).value is None: continue
        if c in {2,5,8,11,14}:  ws_ly.cell(13,c).value = ly_ymd
        elif c in {3,6,9,12,15}: ws_ly.cell(13,c).value = cy_ymd
    ws_ly.cell(22,1).value = f'{MTD_START.year}\n{mtd_lbl}'
    ws_ly.cell(31,1).value = f'{LYMO_START.year}\n{ly_lbl}'

    ytd_ly_lbl=f'{_d(YTD_S_LY)}~{_d(YTD_E_LY)}'; ytd_cy_lbl=f'{_d(YTD_S_CY)}~{_d(YTD_E_CY)}'
    for c in range(1, ws_yoy.max_column+1):
        v = ws_yoy.cell(3,c).value
        if not isinstance(v, str) or '\n' not in v: continue
        parts = v.split('\n')
        if len(parts) >= 3:
            base = '\n'.join(parts[:-2])
            if c in {2,5,8,11,14,17}:   ws_yoy.cell(3,c).value = f'{base}\n{YTD_S_LY.year}\n{ytd_ly_lbl}'
            elif c in {3,6,9,12,15,18}: ws_yoy.cell(3,c).value = f'{base}\n{YTD_S_CY.year}\n{ytd_cy_lbl}'
    ytd_ly_ymd=f'{YTD_S_LY.year}\n{ytd_ly_lbl}'; ytd_cy_ymd=f'{YTD_S_CY.year}\n{ytd_cy_lbl}'
    for c in range(1, ws_yoy.max_column+1):
        if ws_yoy.cell(13,c).value is None: continue
        if c in {2,5,8,11,14}:  ws_yoy.cell(13,c).value = ytd_ly_ymd
        elif c in {3,6,9,12,15}: ws_yoy.cell(13,c).value = ytd_cy_ymd
    ws_yoy.cell(22,1).value = f'{YTD_S_CY.year}\n{ytd_cy_lbl}'
    ws_yoy.cell(31,1).value = f'{YTD_S_LY.year}\n{ytd_ly_lbl}'

    log('儲存 Excel…')
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─── Job 管理 ─────────────────────────────────────────────────
def _run_job(job_id: str, wk_end_str: str):
    def log(msg):
        ts = time.strftime('%H:%M:%S')
        with _LOCK:
            JOBS[job_id]['messages'].append(f'[{ts}] {msg}')

    with _LOCK:
        JOBS[job_id]['status'] = 'running'

    try:
        wk_end = date.fromisoformat(wk_end_str)
        result = _fill_workbook(wk_end, log)
        filename = f'北一區週報_{wk_end}.xlsx'
        with _LOCK:
            JOBS[job_id]['status']   = 'done'
            JOBS[job_id]['result']   = result
            JOBS[job_id]['filename'] = filename
        log(f'✓ 完成！({len(result):,} bytes)')
    except Exception:
        tb = traceback.format_exc()
        with _LOCK:
            JOBS[job_id]['status'] = 'error'
            JOBS[job_id]['error']  = tb
        log(f'✗ 錯誤:\n{tb}')


# ─── HTTP Handler ─────────────────────────────────────────────
class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(STATIC_ROOT), **kwargs)

    def log_message(self, fmt, *args):
        pass  # suppress request logs

    def end_headers(self):
        self.send_header('Cache-Control', 'no-store')
        super().end_headers()

    def send_json(self, status, data):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_xlsx(self, filename, body):
        quoted = urllib.parse.quote(filename)
        self.send_response(200)
        self.send_header('Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        self.send_header('Content-Disposition',
            f"attachment; filename*=UTF-8''{quoted}")
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def read_body(self):
        n = int(self.headers.get('Content-Length', 0))
        return json.loads(self.rfile.read(n).decode('utf-8') if n else '{}')

    def do_GET(self):
        p = urllib.parse.urlparse(self.path)
        qs = urllib.parse.parse_qs(p.query)

        if p.path == '/api/default-date':
            self.send_json(200, {'date': _last_saturday().isoformat()})
            return

        if p.path == '/api/status':
            job_id = qs.get('jobId', [''])[0]
            with _LOCK:
                job = JOBS.get(job_id)
            if not job:
                self.send_json(404, {'error': '找不到工作'})
                return
            resp = {
                'status':   job['status'],
                'messages': list(job['messages']),
                'filename': job.get('filename', ''),
            }
            if job['status'] == 'error':
                resp['error'] = job.get('error', '未知錯誤')
            self.send_json(200, resp)
            return

        if p.path == '/api/download':
            job_id = qs.get('jobId', [''])[0]
            with _LOCK:
                job = JOBS.get(job_id)
            if not job or job['status'] != 'done':
                self.send_json(400, {'error': '檔案尚未就緒'})
                return
            self.send_xlsx(job['filename'], job['result'])
            return

        super().do_GET()

    def do_POST(self):
        p = urllib.parse.urlparse(self.path)
        if p.path == '/api/generate':
            try:
                payload  = self.read_body()
                wk_end_s = str(payload.get('week_end', payload.get('wkEnd', ''))).strip()
                date.fromisoformat(wk_end_s)  # validate
            except Exception as e:
                self.send_json(400, {'error': f'日期格式錯誤: {e}'})
                return
            job_id = str(uuid.uuid4())
            with _LOCK:
                JOBS[job_id] = {'status': 'pending', 'messages': [], 'result': None}
            threading.Thread(target=_run_job, args=(job_id, wk_end_s), daemon=True).start()
            self.send_json(200, {'jobId': job_id})
            return
        self.send_json(404, {'error': 'Not found'})


def main():
    port = int(os.environ.get('PORT', '8782'))
    server = ThreadingHTTPServer(('127.0.0.1', port), Handler)
    print(f'北一區週報產生器：http://127.0.0.1:{port}', flush=True)
    print('按 Ctrl+C 停止', flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('已停止。')


if __name__ == '__main__':
    main()
