# =========================================================
#  네이버 시세 → prices.json 저장  (키움 키 불필요!)
#  - 보유종목 현재가/등락률 + 코스피/코스닥 지수 + (선택)환율을 받아 저장.
#  - 대시보드는 prices.json 을 읽어 화면을 채웁니다.
#  - 폴더의 '보유종목 CSV'(잔고수량 열이 있는 파일)에서 종목을 자동으로 읽습니다.
#
#  [사용법]  python fetch_prices.py   (또는 start_dashboard.bat 더블클릭)
# =========================================================

import requests, json, datetime, glob, os, csv as csvmod, io

EXIM_KEY = ""   # (선택) 수출입은행 환율 인증키. 비우면 환율은 건너뜀.
# (선택) 구글 시트 Apps Script 웹앱 URL(…/exec). 넣으면 폴더에 CSV가 없어도 시트의 모든 탭에서
#        보유종목을 읽어 시세를 받습니다. (대시보드 '시트연결'에 넣는 URL과 동일)
GAS_WEBAPP_URL = ""
FX_CURRENCIES = ["USD", "TWD"]
HDR = {"User-Agent": "Mozilla/5.0", "Referer": "https://m.stock.naver.com/"}

# ---------- 보유종목: 폴더의 '잔고/거래내역 CSV'에서 자동 추출 ----------
import re as _re

def _decode(path):
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp949", "euc-kr"):
        try: return raw.decode(enc)
        except UnicodeDecodeError: continue
    return raw.decode("utf-8", "ignore")

def _num(v):
    try: return float(str(v).replace(",", "").strip() or 0)
    except ValueError: return 0.0

_DATE = _re.compile(r"^\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2}")

def _is_tx(rows):
    """거래내역(헤더 여러 줄 + 2줄=1거래) 형식인지 판별. (헤더가 위쪽 메타행 아래 있어도 인식)"""
    hdr = "".join("".join(map(str, r)) for r in rows[:12]).replace(" ", "").replace("\n", "")
    if _re.search("거래종류|거래일자|거래번호|거래적요|주식매수|주식매도", hdr):
        return True
    dates = sum(1 for r in rows if any(_DATE.match(str(c).strip()) for c in r[:3]))
    has_jango = any("잔고수량" in "".join(map(str, r)).replace(" ", "") for r in rows)
    return dates >= 2 and not has_jango

def _from_balance(rows):
    """잔고 CSV → {code: name} (잔고수량>0)."""
    head = [str(h).replace(" ", "").replace("\n", "") for h in rows[0]]
    def fi(*ns):
        for n in ns:
            if n in head: return head.index(n)
        return -1
    iC, iN, iQ, iMkt = fi("종목코드"), fi("종목명"), fi("잔고수량", "보유수량", "수량"), fi("시장")
    out = {}
    if iC < 0 or iQ < 0: return out
    for c in rows[1:]:
        if len(c) <= max(iC, iQ): continue
        code = str(c[iC]).strip().lstrip("Aa")
        if not code or _num(c[iQ]) == 0: continue
        if iMkt >= 0 and len(c) > iMkt and c[iMkt] and "해외" in c[iMkt]: continue
        out[code] = (c[iN] if iN >= 0 and len(c) > iN else code).strip()
    return out

def _excel_or_date(v):
    s = str(v).strip()
    if _re.match(r"^\d+(\.\d+)?$", s):
        n = float(s)
        if 40000 <= n <= 80000: return True   # 엑셀 일련번호 날짜
    return bool(_re.match(r"^\d{4}[.\-/]\s?\d{1,2}[.\-/]\s?\d{1,2}", s))

def _from_tx(rows):
    """거래내역 CSV → 현재 순보유(매수-매도>0) {code: name}. 헤더 이름으로 컬럼을 찾음(형식 무관)."""
    norm = lambda s: _re.sub(r"\s+", "", str(s))
    hp = next((i for i, r in enumerate(rows)
               if any(norm(c) in ("거래일자", "거래종류", "거래적요") for c in r)), 0)
    HP = [norm(c) for c in rows[hp]] if hp < len(rows) else []
    HQ = [norm(c) for c in rows[hp + 1]] if hp + 1 < len(rows) else []
    def ih(h, *ns):
        for n in ns:
            if n in h: return h.index(n)
        return -1
    def both(*ns):
        i = ih(HP, *ns)
        if i >= 0: return (0, i)
        i = ih(HQ, *ns)
        return (1, i) if i >= 0 else (-1, -1)
    dC = ih(HP, "거래일자");  dC = dC if dC >= 0 else 0
    tC = ih(HP, "거래적요", "거래종류", "거래구분", "적요"); tC = tC if tC >= 0 else 2
    qC = ih(HP, "수량");      qC = qC if qC >= 0 else 4
    cR, cC = both("종목코드")
    if cC < 0: cR, cC = 0, 3
    nR, nC = both("종목명")
    if nC < 0: nR, nC = 1, 3
    pos, name = {}, {}
    for i, P in enumerate(rows):
        if i < hp + 1 or dC >= len(P) or not _excel_or_date(P[dC]): continue
        Q = rows[i + 1] if i + 1 < len(rows) else [""] * 30
        desc = norm(P[tC]) if tC < len(P) else ""
        sign = 1 if "매수" in desc else (-1 if "매도" in desc else 0)
        if not sign: continue
        row_c = Q if cR else P; row_n = Q if nR else P
        code = str(row_c[cC] if cC < len(row_c) else "").strip().lstrip("Aa")
        qty = _num(P[qC] if qC < len(P) else 0)
        if not code or not qty: continue
        pos[code] = pos.get(code, 0) + sign * qty
        nm = str(row_n[nC] if nC < len(row_n) else "").strip()
        if nm: name[code] = nm
    return {c: name.get(c, c) for c, q in pos.items() if q > 1e-6}

def _rows_of(text):
    return list(csvmod.reader(io.StringIO(text)))

def load_stocks():
    # 구글 시트 웹앱(설정 시) + 폴더의 모든 CSV에서 현재 보유 종목을 합칩니다(계좌 여러 개 OK).
    out = {}
    if GAS_WEBAPP_URL:
        try:
            data = requests.get(GAS_WEBAPP_URL, headers=HDR, timeout=20).json()
            for tab in data.get("tabs", []):
                rows = tab.get("rows", [])
                if _is_tx(rows):
                    got = _from_tx(rows)
                    print(f"  [시트:{tab.get('name','?')}] {len(got)}종목 읽음")
                    out.update(got)
        except Exception as e:
            print(f"  [시트] 읽기 실패: {type(e).__name__}")
    for path in glob.glob("*.csv"):
        try:
            rows = _rows_of(_decode(path))
            if len(rows) < 2: continue
            got = _from_tx(rows) if _is_tx(rows) else _from_balance(rows)
            if got:
                print(f"  '{path}' ({'거래내역' if _is_tx(rows) else '잔고'})에서 {len(got)}종목 읽음")
                out.update(got)
        except Exception:
            continue
    if not out:
        print("  보유 CSV/시트 없음 → 기본 종목 사용")
        return {"005930": "삼성전자", "000660": "SK하이닉스", "080220": "제주반도체"}
    return out

# ---------- 네이버 시세 ----------
def naver_stock(code):
    try:
        r = requests.get(f"https://polling.finance.naver.com/api/realtime/domestic/stock/{code}", headers=HDR, timeout=8)
        d = r.json()["datas"][0]
        price = int(str(d["closePrice"]).replace(",", ""))
        rate = float(str(d.get("fluctuationsRatio", "0")).replace(",", "") or 0)
        prev = price / (1 + rate / 100) if rate else price
        return {"price": price, "change": round(price - prev), "rate": rate}
    except Exception as e:
        print(f"    시세 실패 {code}: {type(e).__name__}")
        return None

def naver_index(name):  # name: 'KOSPI' or 'KOSDAQ'
    try:
        r = requests.get(f"https://polling.finance.naver.com/api/realtime/domestic/index/{name}", headers=HDR, timeout=8)
        d = r.json()["datas"][0]
        return float(str(d["closePrice"]).replace(",", ""))
    except Exception:
        return None

# ---------- 수출입은행 환율 (선택) ----------
def get_fx():
    if not EXIM_KEY: return {}
    import urllib3
    try: urllib3.disable_warnings()
    except Exception: pass
    day = datetime.date.today()
    for _ in range(5):
        ymd = day.strftime("%Y%m%d")
        for host in ("https://oapi.koreaexim.go.kr", "https://www.koreaexim.go.kr"):
            for verify in (True, False):
                try:
                    r = requests.get(host + "/site/program/financial/exchangeJSON",
                        params={"authkey": EXIM_KEY, "searchdate": ymd, "data": "AP01"}, timeout=20, verify=verify)
                    data = r.json()
                    if data:
                        out = {}
                        for d in data:
                            u = d.get("cur_unit", ""); base = u.split("(")[0]
                            rate = float(d.get("deal_bas_r", "0").replace(",", ""))
                            if base in FX_CURRENCIES and rate:
                                out[base] = round(rate / (100 if "(100)" in u else 1), 2)
                        if out: return out
                    break
                except Exception: pass
        day -= datetime.timedelta(days=1)
    return {}

# ---------- 과거 시세 재구성 엔진 (자산추이/월별요약용) ----------
def _date_str(v):
    s = str(v).strip()
    if _re.match(r"^\d+(\.\d+)?$", s):
        n = float(s)
        if 40000 <= n <= 80000:
            return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(n))).isoformat()
        return ""
    m = _re.match(r"^(\d{4})[.\-/]\s?(\d{1,2})[.\-/]\s?(\d{1,2})", s)
    return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}" if m else ""

def _parse_tx_full(rows):
    """거래내역 rows → (trades, cashflows). 헤더 이름으로 컬럼 탐지(형식 무관)."""
    norm = lambda s: _re.sub(r"\s+", "", str(s))
    hp = next((i for i, r in enumerate(rows) if any(norm(c) in ("거래일자", "거래종류", "거래적요") for c in r)), 0)
    HP = [norm(c) for c in rows[hp]] if hp < len(rows) else []
    HQ = [norm(c) for c in rows[hp + 1]] if hp + 1 < len(rows) else []
    def ih(h, *ns):
        for n in ns:
            if n in h: return h.index(n)
        return -1
    def both(*ns):
        i = ih(HP, *ns)
        if i >= 0: return (0, i)
        i = ih(HQ, *ns); return (1, i) if i >= 0 else (-1, -1)
    dC = ih(HP, "거래일자"); dC = dC if dC >= 0 else 0
    tC = ih(HP, "거래적요", "거래종류", "거래구분", "적요"); tC = tC if tC >= 0 else 2
    qC = ih(HP, "수량"); qC = qC if qC >= 0 else 4
    aC = ih(HP, "거래금액"); aC = aC if aC >= 0 else 7
    cR, cC = both("종목코드");  cC = cC if cC >= 0 else 3; cR = cR if cR >= 0 else 0
    pR, pC = both("단가");      pC = pC if pC >= 0 else 4; pR = pR if pR >= 0 else 1
    fR, fC = both("제세금", "세금")
    if fC < 0: fR, fC = pR, pC + 1
    trades, cash = [], []
    for i, P in enumerate(rows):
        if i < hp + 1: continue
        desc = norm(P[tC]) if tC < len(P) else ""
        if not any(k in desc for k in ("매수", "매도", "입금", "출금")): continue
        Q = rows[i + 1] if i + 1 < len(rows) else [""] * 40
        date = _date_str(P[dC] if dC < len(P) else "")
        if "매수" in desc or "매도" in desc:
            rc = Q if cR else P; rp = Q if pR else P; rf = Q if fR else P
            code = str(rc[cC] if cC < len(rc) else "").strip().lstrip("Aa")
            qty = _num(P[qC] if qC < len(P) else 0)
            if not code or not qty: continue
            trades.append({"date": date, "type": "매수" if "매수" in desc else "매도", "code": code,
                           "qty": qty, "price": _num(rp[pC] if pC < len(rp) else 0),
                           "fee": _num(rf[fC] if fC < len(rf) else 0)})
        else:
            amt = _num(P[aC] if aC < len(P) else 0) * (-1 if "출금" in desc else 1)
            if amt: cash.append({"date": date, "amount": amt})
    return trades, cash

def load_all_tx():
    trades, cash = [], []
    if GAS_WEBAPP_URL:
        try:
            data = requests.get(GAS_WEBAPP_URL, headers=HDR, timeout=20).json()
            for tab in data.get("tabs", []):
                rows = tab.get("rows", [])
                if _is_tx(rows):
                    t, c = _parse_tx_full(rows); trades += t; cash += c
        except Exception as e:
            print(f"  [시트] 거래내역 읽기 실패: {type(e).__name__}")
    for path in glob.glob("*.csv"):
        try:
            rows = _rows_of(_decode(path))
            if _is_tx(rows):
                t, c = _parse_tx_full(rows); trades += t; cash += c
        except Exception:
            continue
    return trades, cash

def naver_daily(code, start_ymd, end_ymd):
    """종목 일별 종가 {yyyy-mm-dd: close}."""
    try:
        url = (f"https://api.finance.naver.com/siseJson.naver?symbol={code}"
               f"&requestType=1&startTime={start_ymd}&endTime={end_ymd}&timeframe=day")
        txt = requests.get(url, headers=HDR, timeout=15).text.strip()
        arr = json.loads(txt.replace("'", '"'))
        out = {}
        for row in arr[1:]:
            ds = str(row[0])
            if len(ds) == 8 and ds.isdigit():
                out[f"{ds[:4]}-{ds[4:6]}-{ds[6:8]}"] = float(row[4])
        return out
    except Exception as e:
        print(f"    과거시세 실패 {code}: {type(e).__name__}")
        return {}

def _index_from_files(keys):
    """폴더의 파일명에 keys(예: kospi/코스피)가 든 CSV에서 일별 지수 읽기. (날짜 + 숫자 두 칸)"""
    out = {}
    for path in glob.glob("*.csv"):
        if not any(k in path.lower() for k in keys): continue
        try:
            for r in _rows_of(_decode(path)):
                d = next((_date_str(c) for c in r if _date_str(c)), "")
                if not d: continue
                v = next((_num(c) for c in r if not _date_str(c) and _num(c) > 0), 0)
                if v: out[d] = v
        except Exception:
            continue
    if out: print(f"    [지수파일] {keys[0]} {len(out)}일 읽음")
    return out

def index_daily(name):  # 'KOSPI' / 'KOSDAQ'
    keys = ["kospi", "코스피"] if name == "KOSPI" else ["kosdaq", "코스닥"]
    f = _index_from_files(keys)            # 1) 폴더의 kospi.csv / kosdaq.csv 우선
    if f: return f
    try:                                   # 2) 네이버 fchart 일봉(XML)
        txt = requests.get(f"https://fchart.stock.naver.com/sise.nhn?symbol={name}&timeframe=day&count=400&requestType=0",
                           headers=HDR, timeout=15).text
        out = {}
        for m in _re.finditer(r'data="([^"]+)"', txt):
            p = m.group(1).split("|")
            if len(p) >= 5 and len(p[0]) == 8 and p[0].isdigit():
                out[f"{p[0][:4]}-{p[0][4:6]}-{p[0][6:8]}"] = float(p[4])
        if out: return out
    except Exception as e:
        print(f"    지수 fchart 실패 {name}: {type(e).__name__}")
    try:                                   # 3) m.stock JSON (fallback)
        out = {}
        for page in range(1, 8):
            arr = requests.get(f"https://m.stock.naver.com/api/index/{name}/price?pageSize=100&page={page}",
                               headers=HDR, timeout=15).json()
            if not arr: break
            for d in arr:
                ds = _re.sub(r"[./]", "-", str(d.get("localTradedAt", ""))[:10])
                cp = float(str(d.get("closePrice", "0")).replace(",", "") or 0)
                if _re.match(r"\d{4}-\d{2}-\d{2}", ds) and cp: out[ds] = cp
        return out
    except Exception as e:
        print(f"    지수 과거 실패 {name}: {type(e).__name__}")
        return {}

def spx_daily(start, end):  # S&P500 {yyyy-mm-dd: close}
    # 1) 폴더의 snp.csv / sp500.csv 등 우선
    f = _index_from_files(["snp", "sp500", "s&p", "spx", "에스앤피"])
    if f: return {d: v for d, v in f.items() if start <= d <= end}
    # 2) Yahoo Finance (^GSPC)
    try:
        j = requests.get("https://query1.finance.yahoo.com/v8/finance/chart/%5EGSPC?range=2y&interval=1d",
                         headers={"User-Agent": "Mozilla/5.0"}, timeout=15).json()
        res = j["chart"]["result"][0]; ts = res["timestamp"]; cl = res["indicators"]["quote"][0]["close"]
        out = {}
        for t, c in zip(ts, cl):
            if c is None: continue
            d = datetime.datetime.utcfromtimestamp(t).strftime("%Y-%m-%d")
            if start <= d <= end: out[d] = float(c)
        if out: return out
    except Exception as e:
        print(f"    S&P Yahoo 실패: {type(e).__name__}")
    # 3) Stooq fallback
    for sym in ("^spx", "^gspc"):
        try:
            txt = requests.get(f"https://stooq.com/q/d/l/?s={sym}&i=d", headers={"User-Agent": "Mozilla/5.0"}, timeout=20).text.strip()
            out = {}
            for line in txt.splitlines()[1:]:
                p = line.split(",")
                if len(p) >= 5 and len(p[0]) == 10 and p[0][:4].isdigit() and start <= p[0] <= end:
                    try: out[p[0]] = float(p[4])
                    except ValueError: pass
            if out: return out
        except Exception as e:
            print(f"    S&P Stooq 실패({sym}): {type(e).__name__}")
    return {}

def get_fx_naver():  # 네이버 환율 {USD, TWD}
    out = {}
    for cur, sym in (("USD", "FX_USDKRW"), ("TWD", "FX_TWDKRW")):
        try:
            d = requests.get(f"https://api.stock.naver.com/marketindex/exchange/{sym}", headers=HDR, timeout=10).json()
            v = float(str(d.get("closePrice", "")).replace(",", "") or 0)
            if v: out[cur] = round(v, 2)
        except Exception as e:
            print(f"    환율 실패 {cur}: {type(e).__name__}")
    return out

def naver_sector(code):  # 종목 업종명(한글)
    # 1) finance.naver 종목 메인 페이지의 업종 링크
    try:
        html = requests.get(f"https://finance.naver.com/item/main.naver?code={code}",
                            headers=HDR, timeout=8).content.decode("euc-kr", "ignore")
        for pat in (r'type=upjong[^>]*>([^<]+)</a>', r'sise_group_detail[^>]*upjong[^>]*>([^<]+)</a>'):
            m = _re.search(pat, html)
            if m and _re.search(r'[가-힣]', m.group(1)):
                return _re.sub(r"\s+", " ", m.group(1)).strip()
    except Exception as e:
        print(f"      (섹터 html 실패 {code}: {type(e).__name__})")
    # 2) m.stock 통합 API에서 한글 업종 필드
    try:
        j = requests.get(f"https://m.stock.naver.com/api/stock/{code}/integration", headers=HDR, timeout=8).json()
        found = [""]
        def walk(o):
            if found[0]: return
            if isinstance(o, dict):
                for k, v in o.items():
                    if isinstance(v, str) and _re.search(r'[가-힣]', v) and ("industry" in k.lower() or "sector" in k.lower() or "업종" in k):
                        found[0] = v.strip(); return
                    walk(v)
            elif isinstance(o, list):
                for x in o: walk(x)
        walk(j)
        if found[0]: return found[0]
    except Exception:
        pass
    return ""

def build_series(trades, cash, today):
    import bisect
    alldates = [t["date"] for t in trades if t["date"]] + [c["date"] for c in cash if c["date"]]
    if not alldates: return []
    start = min(alldates); end = today.isoformat()
    codes = sorted({t["code"] for t in trades})
    print(f"  과거 일별 종가 받는 중... ({len(codes)}종목, {start}~{end})")
    phist = {code: naver_daily(code, start.replace("-", ""), end.replace("-", "")) for code in codes}
    kospi = index_daily("KOSPI"); kosdaq = index_daily("KOSDAQ"); spx = spx_daily(start, end)
    print(f"  지수/벤치마크: 코스피 {len(kospi)}일 · 코스닥 {len(kosdaq)}일 · S&P {len(spx)}일")
    daysset = set()
    for h in phist.values(): daysset |= set(h.keys())
    daysset |= set(kospi.keys())
    days = sorted(d for d in daysset if start <= d <= end)
    if not days: return []
    def ffill(hist):
        sk = sorted(hist.keys()); out = {}
        for d in days:
            idx = bisect.bisect_right(sk, d) - 1
            out[d] = hist[sk[idx]] if idx >= 0 else None
        return out
    pf = {code: ffill(h) for code, h in phist.items()}
    kf, qf, sf = ffill(kospi), ffill(kosdaq), ffill(spx)
    events = [(t["date"] or start, "t", t) for t in trades] + [(c["date"] or start, "c", c) for c in cash]
    events.sort(key=lambda x: x[0])
    holdings, cashbal, ei, series = {}, 0.0, 0, []
    for day in days:
        while ei < len(events) and events[ei][0] <= day:
            kind, ev = events[ei][1], events[ei][2]
            if kind == "t":
                if ev["type"] == "매수":
                    holdings[ev["code"]] = holdings.get(ev["code"], 0) + ev["qty"]; cashbal -= ev["qty"] * ev["price"] + ev["fee"]
                else:
                    holdings[ev["code"]] = holdings.get(ev["code"], 0) - ev["qty"]; cashbal += ev["qty"] * ev["price"] - ev["fee"]
            else:
                cashbal += ev["amount"]
            ei += 1
        stockval = sum(q * pf[code][day] for code, q in holdings.items() if q > 1e-6 and pf.get(code, {}).get(day))
        rec = {"d": day, "asset": round(stockval + cashbal)}
        if kf.get(day): rec["kospi"] = kf[day]
        if qf.get(day): rec["kosdaq"] = qf[day]
        if sf.get(day): rec["spx"] = sf[day]
        series.append(rec)
    return series

# ---------- 실행 ----------
print("보유종목 읽는 중...")
STOCKS = load_stocks()

result = {}
print("네이버 시세 받는 중...")
for code, name in STOCKS.items():
    s = naver_stock(code)
    if s:
        s["name"] = name; result[code] = s
        print(f"  {name}({code}): {s['price']:,}원 ({s['rate']:+.2f}%)")

print("지수 받는 중...")
ks = naver_index("KOSPI"); kq = naver_index("KOSDAQ")
if ks: result["_kospi"] = ks; print(f"  코스피: {ks:,.2f}")
if kq: result["_kosdaq"] = kq; print(f"  코스닥: {kq:,.2f}")

print("섹터 정보 받는 중...")
_secn = 0
for _code in [c for c in result if not c.startswith("_")]:
    if not _code.isdigit(): continue          # 국내 종목코드만
    sec = naver_sector(_code)
    print(f"  {result[_code].get('name', _code)}({_code}): {sec or '(못찾음)'}")
    if sec: result[_code]["sector"] = sec; _secn += 1
print(f"  섹터 {_secn}종목 확인")

print("환율 받는 중...")
fx = {**get_fx(), **get_fx_naver()}            # 네이버 우선, 없으면 수출입은행
if fx: result["_fx"] = fx; print(f"  환율: {fx}")

print("과거 시세로 자산추이 재구성 중...")
try:
    _trades, _cash = load_all_tx()
    _series = build_series(_trades, _cash, datetime.date.today())
    if _series:
        result["_series"] = _series
        print(f"  ✅ 자산추이 {len(_series)}일 재구성 (최근 총자산 약 {_series[-1]['asset']:,}원)")
    else:
        print("  거래내역이 없어 자산추이 생략 (시트연결/CSV 확인)")
except Exception as e:
    print(f"  자산추이 재구성 실패: {type(e).__name__}: {e}")

result["_updated"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
with open("prices.json", "w", encoding="utf-8") as f:
    json.dump(result, f, ensure_ascii=False, indent=2)
print("=========================================")
print("✅ prices.json 저장 완료!")
print("=========================================")
