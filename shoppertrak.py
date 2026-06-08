"""ShopperTrak 來客數抓取 — Playwright 無頭登入 + REST API。

ShopperTrak 無公開 API key，登入為 SSO（帳密 → sessionStorage 取 token）。
本模組用 Playwright 開無頭 Chromium 自動登入，取出 authToken/tenantId 後直接打 REST API。
任何一步失敗都丟 ShopperTrakError，由主程式 try/except 略過來客數、其餘照常產生。
"""
from datetime import date

ORG_ID    = 5536
API_BASE  = "https://rdc-api.shoppertrak.com"
LOGIN_URL = "https://analytics.shoppertrak.com/"

# 門市代碼 → ShopperTrak siteId（僅有計數器門市；無計數器門市不在此）
SITE_IDS = {
    '004': 82751,     # 士林
    '005': 80028316,  # 微風
    '024': 10094800,  # 美麗華
    '046': 80009128,  # 阿波羅
    '054': 80031194,  # 大葉高島屋
}

_ACCOUNT_KEYWORDS = ('email', 'user', 'account', 'login', '帳號')
_SUBMIT_KEYWORDS  = ('login', 'sign in', 'signin', 'next', 'continue',
                     'submit', '下一步', '繼續', '登入', '送出')

# token 程式內快取
_token = {"authToken": None, "tenantId": None}


class ShopperTrakError(Exception):
    pass


def _read_token(page):
    return page.evaluate(
        "() => ({a: sessionStorage.getItem('authToken'),"
        " t: sessionStorage.getItem('tenantId')})")


def _find_account_field(page):
    """掃所有 input，挑帳號欄（type 合適 + name/id/placeholder 等含關鍵字）。"""
    inputs = page.query_selector_all("input")
    fallback = None
    for el in inputs:
        try:
            if not el.is_visible():
                continue
            typ = (el.get_attribute("type") or "text").lower()
            if typ not in ("text", "email", "tel", "search"):
                continue
            attrs = " ".join(filter(None, [
                el.get_attribute("name"), el.get_attribute("id"),
                el.get_attribute("placeholder"), el.get_attribute("autocomplete"),
                el.get_attribute("aria-label")])).lower()
            if any(k in attrs for k in _ACCOUNT_KEYWORDS):
                return el
            if fallback is None:
                fallback = el
        except Exception:
            continue
    return fallback


def _password_field(page):
    return page.query_selector(
        "input[type=password]:not([disabled]):not([readonly])")


def _submit(page, field):
    """送出：先按 Enter，失敗再點含關鍵字的按鈕。"""
    try:
        field.press("Enter")
        return
    except Exception:
        pass
    for btn in page.query_selector_all("button, input[type=submit], a"):
        try:
            txt = ((btn.inner_text() or "") + " " +
                   (btn.get_attribute("value") or "")).strip().lower()
            if txt and any(k in txt for k in _SUBMIT_KEYWORDS):
                btn.click()
                return
        except Exception:
            continue


def _login(username, password, log):
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        raise ShopperTrakError(
            "未安裝 Playwright，請執行："
            "pip3 install playwright && python3 -m playwright install chromium")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            page = browser.new_page()
            page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=60000)
            page.wait_for_timeout(2000)

            # 已登入就直接用
            tok = _read_token(page)
            if tok.get("a") and tok.get("t"):
                return tok["a"], tok["t"]

            # 帳號欄
            acc = _find_account_field(page)
            if not acc:
                raise ShopperTrakError("找不到帳號輸入欄")
            acc.fill(username)

            # 密碼欄（可能兩步式）
            pwd = _password_field(page)
            if not pwd:
                _submit(page, acc)                 # 兩步式：先送帳號
                for _ in range(30):                # 輪詢等密碼欄出現
                    page.wait_for_timeout(1000)
                    pwd = _password_field(page)
                    if pwd:
                        break
            if not pwd:
                raise ShopperTrakError("送出帳號後仍找不到密碼欄")
            pwd.fill(password)
            _submit(page, pwd)

            # 輪詢 token（逾時 60 秒）
            for _ in range(60):
                page.wait_for_timeout(1000)
                tok = _read_token(page)
                if tok.get("a") and tok.get("t"):
                    log("ShopperTrak 登入成功")
                    return tok["a"], tok["t"]
            raise ShopperTrakError("登入逾時，未取得 token（帳密可能錯誤）")
        finally:
            browser.close()


def _ensure_token(username, password, log, force=False):
    if force or not (_token["authToken"] and _token["tenantId"]):
        a, t = _login(username, password, log)
        _token["authToken"], _token["tenantId"] = a, t
    return _token["authToken"], _token["tenantId"]


def _query_daily(site_id, start: date, end: date, auth, tenant):
    """打 REST API 取每日 traffic，回傳 {date: traffic}。"""
    import urllib.request, json as _json
    url = f"{API_BASE}/api/v1/kpis/organizations/{ORG_ID}/sites/{site_id}"
    body = _json.dumps({
        "groupBy": "day",
        "operatingHours": True,
        "reportStartDate": f"{start.isoformat()}T00:00:00.000Z",
        "reportEndDate":   f"{end.isoformat()}T00:00:00.000Z",
        "add_aggregated_data": True,
        "kpi": ["traffic"],
    }).encode("utf-8")
    req = urllib.request.Request(url, data=body, method="POST", headers={
        "Authorization": f"Bearer {auth}", "tenant": tenant,
        "Content-Type": "application/json", "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=120) as r:
        if r.status in (401, 419):
            raise PermissionError("token expired")
        data = _json.loads(r.read().decode("utf-8"))
    rows = (((data.get("result") or [{}])[0]).get("currentPeriod") or {}).get("data") or []
    return {d.get("day"): int(d.get("traffic") or 0) for d in rows if d.get("day")}


def fetch_all(store_codes, start: date, end: date, username, password, log):
    """登入並抓多店每日來客數，回傳 {storeCode: {date: traffic}}。失敗丟 ShopperTrakError。"""
    auth, tenant = _ensure_token(username, password, log)
    out = {}
    for code in store_codes:
        site = SITE_IDS.get(code)
        if not site:
            continue
        try:
            try:
                days = _query_daily(site, start, end, auth, tenant)
            except PermissionError:
                auth, tenant = _ensure_token(username, password, log, force=True)
                days = _query_daily(site, start, end, auth, tenant)
            if days:
                out[code] = days
                log(f"  來客數 {code}: {len(days)} 天")
        except Exception as e:
            log(f"  來客數 {code} 失敗：{e}")
    return out
