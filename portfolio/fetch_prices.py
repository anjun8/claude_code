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
FX_CURRENCIES = ["USD", "TWD"]
HDR = {"User-Agent": "Mozilla/5.0", "Referer": "https://m.stock.naver.com/"}

# ---------- 보유종목: 폴더의 '잔고 CSV'에서 자동 추출 ----------
def _decode(path):
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp949", "euc-kr"):
        try: return raw.decode(enc)
        except UnicodeDecodeError: continue
    return raw.decode("utf-8", "ignore")

def load_stocks():
    # 폴더의 '잔고 CSV'를 모두 읽어 종목을 합칩니다(계좌가 여러 개여도 OK).
    files = []
    for path in glob.glob("*.csv"):
        try:
            head = next(csvmod.reader(io.StringIO(_decode(path))))
            head = [h.replace(" ", "").replace("\n", "") for h in head]
            if "잔고수량" in head or ("수량" in head and "매입단가" in head):  # 잔고/양식 파일만
                files.append((path, head))
        except Exception: continue
    if not files:
        print("  보유 CSV 없음 → 기본 종목 사용")
        return {"005930": "삼성전자", "000660": "SK하이닉스", "080220": "제주반도체"}
    out = {}
    for path, head in files:
        rows = list(csvmod.reader(io.StringIO(_decode(path))))
        def fi(*ns):
            for n in ns:
                if n in head: return head.index(n)
            return -1
        iC, iN, iQ, iMkt = fi("종목코드"), fi("종목명"), fi("잔고수량", "수량"), fi("시장")
        if iC < 0 or iQ < 0: continue
        cnt = 0
        for c in rows[1:]:
            if len(c) <= max(iC, iQ): continue
            code = str(c[iC]).strip().lstrip("Aa")
            if not code: continue
            try: q = float(str(c[iQ]).replace(",", ""))
            except ValueError: q = 0
            if not q: continue
            if iMkt >= 0 and len(c) > iMkt and c[iMkt] and "해외" in c[iMkt]: continue   # 해외는 네이버 국내시세 대상 아님
            out[code] = (c[iN] if iN >= 0 and len(c) > iN else code).strip()
            cnt += 1
        print(f"  '{path}' 에서 {cnt}종목 읽음")
    return out or {"005930": "삼성전자"}

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

fx = get_fx()
if fx: result["_fx"] = fx; print(f"  환율: {fx}")

result["_updated"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
with open("prices.json", "w", encoding="utf-8") as f:
    json.dump(result, f, ensure_ascii=False, indent=2)
print("=========================================")
print("✅ prices.json 저장 완료!")
print("=========================================")
