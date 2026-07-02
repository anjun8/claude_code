# -*- coding: utf-8 -*-
"""
재무제표 추출기 — 결과물 예시 엑셀 생성 스크립트

사업보고서/분기보고서를 넣으면 만들어질 최종 엑셀의 "모양"을 미리 보여주기 위한
샘플 데이터 기반 생성기.

시트 구성: IS / COST / BS / CF / Valuation (+ Check: 보고서 간 값이 달라진 항목 알림)

가상 시나리오
  입력 보고서: 사업보고서 FY2023 / FY2024 / FY2025,
              분기보고서 2024·2025 (1Q, 반기, 3Q)
  → FY2023 사업보고서의 전전기(제n-2기) 컬럼 덕분에 2021년 숫자까지 소급 추출
  → 같은 기간 숫자가 보고서마다 다르면 "최신 보고서" 값을 채택하고 Check 시트에 기록
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ──────────────────────────────────────────────────────────────
# 공통 스타일
# ──────────────────────────────────────────────────────────────
F_BASE = Font(name="맑은 고딕", size=10)
F_TITLE = Font(name="맑은 고딕", size=14, bold=True, color="1F4E79")
F_HDR = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
F_SRC = Font(name="맑은 고딕", size=8, italic=True, color="808080")
F_BOLD = Font(name="맑은 고딕", size=10, bold=True)
F_PCT = Font(name="맑은 고딕", size=9, color="606060")
F_NOTE = Font(name="맑은 고딕", size=9, color="808080")

FILL_HDR = PatternFill("solid", fgColor="1F4E79")
FILL_SUBTOTAL = PatternFill("solid", fgColor="DDEBF7")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
FILL_DERIVED = PatternFill("solid", fgColor="F2F2F2")

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

NUM = "#,##0.0;[Red](#,##0.0)"
PCT = "0.0%;[Red](0.0%)"
INT = "#,##0"

IDT = "　"  # 전각 공백 들여쓰기 (기존 양식과 동일)

# ──────────────────────────────────────────────────────────────
# 컬럼 레이아웃:  C~G = 연간 2021~2025,  H~O = 분기 1Q24~4Q25
# ──────────────────────────────────────────────────────────────
YEARS = [2021, 2022, 2023, 2024, 2025]
QTRS = ["1Q24", "2Q24", "3Q24", "4Q24", "1Q25", "2Q25", "3Q25", "4Q25"]
Y_COLS = {y: get_column_letter(3 + i) for i, y in enumerate(YEARS)}          # C..G
Q_COLS = {q: get_column_letter(8 + i) for i, q in enumerate(QTRS)}           # H..O
ALL_COLS = list(Y_COLS.values()) + list(Q_COLS.values())

SOURCES = {
    "2021": "FY2023 사업보고서 (전전기)",
    "2022": "FY2024 사업보고서 (전전기)",
    "2023": "FY2025 사업보고서 (전전기)",
    "2024": "FY2025 사업보고서 (전기)",
    "2025": "FY2025 사업보고서 (당기)",
    "1Q24": "2025.1Q 분기보고서 (전년비교)",
    "2Q24": "2025 반기보고서 (전년비교)",
    "3Q24": "2025.3Q 분기보고서 (전년비교)",
    "4Q24": "도출: FY2024 − (1Q+2Q+3Q)",
    "1Q25": "2025.1Q 분기보고서 (당기)",
    "2Q25": "2025 반기보고서 (당기)",
    "3Q25": "2025.3Q 분기보고서 (당기)",
    "4Q25": "도출: FY2025 − (1Q+2Q+3Q)",
}

# ──────────────────────────────────────────────────────────────
# 샘플 데이터 (단위: 억원) — 가상의 회사
# ──────────────────────────────────────────────────────────────
IS_Y = {  # 연간 2021~2025
    "매출액":     [812.4, 934.1, 1023.5, 1187.9, 1352.6],
    "매출원가":   [487.2, 549.8, 601.3, 688.5, 769.4],
    "판매비와관리비": [214.6, 241.3, 265.8, 301.2, 338.7],
    "기타수익":   [3.2, 4.1, 2.8, 5.6, 4.9],
    "기타비용":   [5.1, 3.8, 6.2, 4.3, 7.1],
    "금융수익":   [6.4, 8.9, 12.3, 15.1, 13.8],
    "금융비용":   [9.8, 11.2, 10.6, 9.4, 8.2],
    "법인세비용": [21.3, 28.9, 31.5, 42.8, 51.6],
}
IS_Q = {  # 1Q24, 2Q24, 3Q24, 1Q25, 2Q25, 3Q25  (4Q는 연간-3개분기 수식으로 도출)
    "매출액":     [275.3, 291.8, 298.4, 312.7, 334.2, 341.5],
    "매출원가":   [162.1, 169.5, 172.8, 180.3, 190.1, 194.6],
    "판매비와관리비": [71.2, 74.8, 76.1, 80.9, 83.6, 85.2],
    "기타수익":   [1.2, 1.5, 1.1, 1.0, 1.4, 1.2],
    "기타비용":   [0.9, 1.1, 1.3, 1.8, 1.6, 2.1],
    "금융수익":   [3.6, 3.9, 3.8, 3.4, 3.5, 3.3],
    "금융비용":   [2.5, 2.4, 2.3, 2.1, 2.0, 2.1],
    "법인세비용": [10.2, 11.5, 10.8, 12.4, 13.7, 12.9],
}

# COST: 비용의 성격별 분류 (합계 = 매출원가+판관비), '기타'가 plug
COST_NATURE_BASE = {  # 2025년 기준 구성비 산출용
    "제품과 재공품의 변동": [-12.3, -8.5, 3.1, -15.2, -10.8],
    "원재료와 소모품의 사용": [285.4, 322.1, 348.9, 401.5, 452.3],
    "종업원급여": [142.7, 158.3, 172.4, 195.8, 219.6],
    "감가상각비": [48.2, 54.6, 59.1, 66.3, 74.5],
    "무형자산상각비": [12.1, 13.5, 14.2, 15.8, 17.3],
    "지급수수료": [89.5, 101.2, 110.8, 128.4, 145.2],
    "임차료 및 관리비": [18.3, 20.1, 21.5, 24.2, 27.1],
    "광고선전비": [35.2, 41.8, 45.3, 52.6, 60.4],
    "운반비": [24.6, 28.3, 30.9, 35.7, 40.6],
    "소모품비": [15.8, 17.9, 19.4, 22.3, 25.4],
}

# 판관비 세부 (합계 = IS 판관비), '기타'가 plug
SGA_BASE = {
    "급여": [46.2, 52.8, 58.9, 68.3, 78.4],
    "퇴직급여": [5.4, 6.1, 6.8, 7.9, 8.9],
    "복리후생비": [7.8, 8.6, 9.7, 11.2, 12.6],
    "여비교통비": [2.1, 2.8, 3.3, 3.8, 4.2],
    "감가상각비": [13.5, 15.2, 17.1, 19.6, 22.1],
    "무형자산상각비": [4.2, 4.8, 5.3, 6.1, 6.8],
    "지급수수료": [41.8, 47.2, 52.6, 61.0, 68.3],
    "임차료 및 관리비": [9.6, 10.7, 11.9, 13.7, 15.4],
    "광고선전비": [35.2, 41.8, 45.3, 52.6, 60.4],
    "운반비": [11.3, 12.8, 14.2, 16.3, 18.2],
    "세금과공과": [3.2, 3.6, 4.0, 4.6, 5.1],
    "소모품비": [4.9, 5.5, 6.1, 7.0, 7.9],
}

# BS 연말 잔액 (2021~2025)
BS_Y = {
    "현금및현금성자산": [72.4, 85.1, 96.8, 118.3, 142.6],
    "단기금융상품": [45.0, 52.3, 60.1, 71.5, 85.2],
    "매출채권 및 기타채권": [98.6, 112.4, 121.9, 139.7, 158.3],
    "재고자산": [76.2, 88.9, 95.4, 108.2, 121.8],
    "기타유동자산": [8.3, 9.7, 10.5, 12.1, 13.9],
    "유형자산": [156.8, 172.3, 189.5, 212.6, 238.4],
    "사용권자산": [18.2, 20.5, 22.8, 25.4, 28.1],
    "무형자산": [34.5, 36.8, 39.2, 42.5, 45.9],
    "장기금융상품": [12.0, 14.5, 16.8, 19.2, 22.5],
    "이연법인세자산": [9.4, 10.2, 11.1, 12.3, 13.6],
    "매입채무 및 기타채무": [68.4, 76.2, 82.5, 93.8, 105.6],
    "단기차입금": [55.0, 50.0, 45.0, 40.0, 35.0],
    "유동성리스부채": [4.2, 4.7, 5.2, 5.8, 6.4],
    "당기법인세부채": [8.6, 10.4, 11.2, 14.8, 17.5],
    "기타유동부채": [6.8, 7.5, 8.2, 9.4, 10.7],
    "장기차입금": [80.0, 70.0, 60.0, 50.0, 40.0],
    "리스부채": [14.6, 16.2, 17.9, 19.8, 21.9],
    "확정급여부채": [22.3, 25.1, 28.2, 31.8, 35.7],
    "기타비유동부채": [3.5, 3.9, 4.3, 4.8, 5.3],
    "자본금": [50.0, 50.0, 50.0, 50.0, 50.0],
    "자본잉여금": [120.0, 120.0, 120.0, 120.0, 120.0],
    "기타자본구성요소": [-5.2, -5.2, -4.8, -4.8, -4.5],
}
BS_ASSET_ITEMS = ["현금및현금성자산", "단기금융상품", "매출채권 및 기타채권", "재고자산", "기타유동자산",
                  "유형자산", "사용권자산", "무형자산", "장기금융상품", "이연법인세자산"]
BS_CUR_ASSET = BS_ASSET_ITEMS[:5]
BS_NONCUR_ASSET = BS_ASSET_ITEMS[5:]
BS_CUR_LIAB = ["매입채무 및 기타채무", "단기차입금", "유동성리스부채", "당기법인세부채", "기타유동부채"]
BS_NONCUR_LIAB = ["장기차입금", "리스부채", "확정급여부채", "기타비유동부채"]
BS_EQUITY_FIXED = ["자본금", "자본잉여금", "기타자본구성요소"]

CASH_2020 = 58.3  # 2021 CF 기초현금용

# ──────────────────────────────────────────────────────────────
# 파생 데이터 계산
# ──────────────────────────────────────────────────────────────
def r1(x):
    return round(x + 1e-9, 1)

# 이익잉여금 = 자산총계 − 부채총계 − (자본금+자본잉여금+기타자본)  → 대차 자동 일치
def bs_totals(values_by_item):
    a = sum(values_by_item[k] for k in BS_ASSET_ITEMS)
    l = sum(values_by_item[k] for k in BS_CUR_LIAB + BS_NONCUR_LIAB)
    eq_fixed = sum(values_by_item[k] for k in BS_EQUITY_FIXED)
    re = r1(a - l - eq_fixed)
    return r1(a), r1(l), re

BS_Y_FULL = []  # 연도별 dict
for i, y in enumerate(YEARS):
    d = {k: BS_Y[k][i] for k in BS_Y}
    a, l, re = bs_totals(d)
    d["이익잉여금"] = re
    BS_Y_FULL.append(d)

# 분기 BS: 직전/당해 연말 사이 선형보간 + 소폭 변동 (4Q는 연간과 동일 → 수식 참조)
WIGGLE = [0.018, -0.011, 0.014]  # 1Q, 2Q, 3Q
BS_Q_FULL = {}  # "1Q24" → dict (1~3Q만)
for yi, ylabel in [(3, "24"), (4, "25")]:
    prev = BS_Y_FULL[yi - 1]
    cur = BS_Y_FULL[yi]
    for qn in range(1, 4):
        frac = qn / 4
        d = {}
        for k in BS_Y:
            base = prev[k] + (cur[k] - prev[k]) * frac
            d[k] = r1(base * (1 + WIGGLE[qn - 1]))
        a, l, re = bs_totals(d)
        d["이익잉여금"] = re
        BS_Q_FULL[f"{qn}Q{ylabel}"] = d

# CF: 운전자본 증감이 plug → 현금 잔액과 완전 일치
def build_cf(ni, dep_adj, int_in, int_out, tax_paid,
             stf, capex, disp, intg, borrow, repay, lease, div, fx,
             cash_begin, cash_end):
    inv = r1(stf + capex + disp + intg)
    fin = r1(borrow + repay + lease + div)
    delta = r1(cash_end - cash_begin)
    ops_needed = r1(delta - inv - fin - fx)
    wc = r1(ops_needed - ni - dep_adj - int_in - int_out - tax_paid)
    return {
        "당기순이익": ni, "비현금항목 조정": dep_adj, "영업활동 자산부채의 증감": wc,
        "이자수취": int_in, "이자지급": int_out, "법인세납부": tax_paid,
        "단기금융상품의 순증감": stf, "유형자산의 취득": capex, "유형자산의 처분": disp,
        "무형자산의 취득": intg,
        "차입금의 차입": borrow, "차입금의 상환": repay, "리스부채의 상환": lease,
        "배당금의 지급": div,
        "외화환산으로 인한 현금의 변동": fx,
        "기초현금": cash_begin, "기말현금": cash_end,
    }

def is_ni_y(i):  # 연간 당기순이익 (검산용)
    gp = IS_Y["매출액"][i] - IS_Y["매출원가"][i]
    op = gp - IS_Y["판매비와관리비"][i]
    pbt = op + IS_Y["기타수익"][i] - IS_Y["기타비용"][i] + IS_Y["금융수익"][i] - IS_Y["금융비용"][i]
    return r1(pbt - IS_Y["법인세비용"][i])

def is_ni_q(j):  # 분기(1~3Q) 당기순이익
    gp = IS_Q["매출액"][j] - IS_Q["매출원가"][j]
    op = gp - IS_Q["판매비와관리비"][j]
    pbt = op + IS_Q["기타수익"][j] - IS_Q["기타비용"][j] + IS_Q["금융수익"][j] - IS_Q["금융비용"][j]
    return r1(pbt - IS_Q["법인세비용"][j])

CF_Y = []
cf_params_y = [  # dep, int_in, int_out, tax, stf, capex, disp, intg, borrow, repay, lease, div, fx
    (62.4, 5.8, -9.2, -19.5, -6.2, -38.5, 2.1, -8.4, 20.0, -35.0, -4.5, -12.0, 0.8),
    (70.2, 8.1, -10.6, -26.4, -7.3, -42.8, 1.5, -9.2, 15.0, -30.0, -5.0, -14.0, -1.2),
    (75.6, 11.4, -10.1, -29.8, -7.8, -48.2, 3.2, -10.5, 10.0, -25.0, -5.5, -16.0, 0.5),
    (84.3, 14.2, -8.9, -38.6, -11.4, -55.4, 1.8, -11.8, 10.0, -25.0, -6.1, -18.0, -0.7),
    (94.8, 12.9, -7.8, -47.2, -13.7, -63.2, 2.4, -12.6, 5.0, -20.0, -6.7, -21.0, 1.1),
]
for i, y in enumerate(YEARS):
    begin = CASH_2020 if i == 0 else BS_Y_FULL[i - 1]["현금및현금성자산"]
    end = BS_Y_FULL[i]["현금및현금성자산"]
    p = cf_params_y[i]
    CF_Y.append(build_cf(is_ni_y(i), *p, cash_begin=begin, cash_end=end))

CF_Q = {}  # 1~3Q만 (4Q는 수식 도출)
cf_params_q = {
    "1Q24": (20.5, 3.4, -2.3, -9.1, -2.8, -13.2, 0.4, -2.9, 5.0, -6.0, -1.5, -18.0, -0.2),
    "2Q24": (21.1, 3.6, -2.2, -10.2, -2.9, -13.8, 0.5, -3.0, 2.0, -6.0, -1.5, 0.0, 0.3),
    "3Q24": (21.3, 3.5, -2.2, -9.8, -3.0, -14.1, 0.4, -2.9, 3.0, -6.5, -1.5, 0.0, -0.4),
    "1Q25": (23.2, 3.3, -2.0, -11.5, -3.3, -15.4, 0.6, -3.1, 2.0, -5.0, -1.6, -21.0, 0.3),
    "2Q25": (23.7, 3.2, -1.9, -12.1, -3.4, -15.9, 0.5, -3.2, 1.0, -5.0, -1.7, 0.0, -0.3),
    "3Q25": (23.9, 3.2, -2.0, -11.8, -3.5, -16.2, 0.7, -3.1, 2.0, -5.0, -1.7, 0.0, 0.4),
}
for j, q in enumerate(["1Q24", "2Q24", "3Q24", "1Q25", "2Q25", "3Q25"]):
    if q.endswith("24"):
        begin = BS_Y_FULL[2]["현금및현금성자산"] if q == "1Q24" else BS_Q_FULL[["1Q24", "2Q24"][int(q[0]) - 2]]["현금및현금성자산"]
    else:
        begin = BS_Y_FULL[3]["현금및현금성자산"] if q == "1Q25" else BS_Q_FULL[["1Q25", "2Q25"][int(q[0]) - 2]]["현금및현금성자산"]
    end = BS_Q_FULL[q]["현금및현금성자산"]
    CF_Q[q] = build_cf(is_ni_q(j), *cf_params_q[q], cash_begin=begin, cash_end=end)

# ──────────────────────────────────────────────────────────────
# 시트 작성 헬퍼
# ──────────────────────────────────────────────────────────────
def style_cell(c, font=F_BASE, fmt=None, fill=None, border=False, align=None):
    c.font = font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if border:
        c.border = BORDER
    if align:
        c.alignment = Alignment(horizontal=align)

def sheet_header(ws, title, unit, cols, labels, with_src=True):
    ws.sheet_view.showGridLines = False
    ws["B2"] = title
    ws["B2"].font = F_TITLE
    ws["B4"] = unit
    style_cell(ws["B4"], F_NOTE)
    for col, label in zip(cols, labels):
        c = ws[f"{col}4"]
        c.value = label
        style_cell(c, F_HDR, fill=FILL_HDR, border=True, align="center")
        if with_src:
            s = ws[f"{col}5"]
            s.value = SOURCES.get(str(label), "")
            style_cell(s, F_SRC, align="center")
            ws[f"{col}5"].alignment = Alignment(horizontal="center", wrap_text=True)
    hdr = ws["B4"]
    hdr.border = BORDER
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 34
    for col in cols:
        ws.column_dimensions[col].width = 12.5
    if with_src:
        ws.row_dimensions[5].height = 26
    ws.freeze_panes = "C7"

def put_row(ws, row, label, values_by_col, font=F_BASE, fmt=NUM, fill=None, bold_label=False):
    c = ws[f"B{row}"]
    c.value = label
    c.font = F_BOLD if bold_label else font
    if fill:
        c.fill = fill
    c.border = BORDER
    for col, v in values_by_col.items():
        cell = ws[f"{col}{row}"]
        cell.value = v
        style_cell(cell, font, fmt=fmt, fill=fill, border=True)

# ──────────────────────────────────────────────────────────────
# 통합 문서 생성
# ──────────────────────────────────────────────────────────────
wb = openpyxl.Workbook()
wb.remove(wb.active)

# ============================== IS ==============================
ws = wb.create_sheet("IS")
ws.sheet_properties.tabColor = "1F4E79"
sheet_header(ws, "IS — 손익계산서 (연결)", "(단위: 억원)",
             ALL_COLS, YEARS + QTRS)

Q4_MAP = {"4Q24": ("F", ["H", "I", "J"]), "4Q25": ("G", ["L", "M", "N"])}  # 연간컬럼, 1~3Q컬럼

def is_base_row(row, item):
    vals = {}
    for i, y in enumerate(YEARS):
        vals[Y_COLS[y]] = IS_Y[item][i]
    q123 = ["1Q24", "2Q24", "3Q24", "1Q25", "2Q25", "3Q25"]
    for j, q in enumerate(q123):
        vals[Q_COLS[q]] = IS_Q[item][j]
    for q, (ycol, qcols) in Q4_MAP.items():
        col = Q_COLS[q]
        vals[col] = f"={ycol}{row}-{qcols[0]}{row}-{qcols[1]}{row}-{qcols[2]}{row}"
    return vals

R = {}
row = 7
R["매출액"] = row; put_row(ws, row, "매출액", is_base_row(row, "매출액"), bold_label=True)
row += 1
yoy = {}
for i in range(1, 5):
    a, b = Y_COLS[YEARS[i]], Y_COLS[YEARS[i - 1]]
    yoy[a] = f"={a}{R['매출액']}/{b}{R['매출액']}-1"
for cur, prev in [("1Q25", "1Q24"), ("2Q25", "2Q24"), ("3Q25", "3Q24"), ("4Q25", "4Q24")]:
    yoy[Q_COLS[cur]] = f"={Q_COLS[cur]}{R['매출액']}/{Q_COLS[prev]}{R['매출액']}-1"
put_row(ws, row, f"{IDT}YoY (%)", yoy, font=F_PCT, fmt=PCT)
row += 1
R["매출원가"] = row; put_row(ws, row, "매출원가", is_base_row(row, "매출원가"))
row += 1
R["매출총이익"] = row
put_row(ws, row, "매출총이익",
        {c: f"={c}{R['매출액']}-{c}{R['매출원가']}" for c in ALL_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
row += 1
put_row(ws, row, f"{IDT}GPM (%)",
        {c: f"={c}{R['매출총이익']}/{c}{R['매출액']}" for c in ALL_COLS}, font=F_PCT, fmt=PCT)
row += 1
R["판관비"] = row; put_row(ws, row, "판매비와관리비", is_base_row(row, "판매비와관리비"))
row += 1
R["영업이익"] = row
put_row(ws, row, "영업이익",
        {c: f"={c}{R['매출총이익']}-{c}{R['판관비']}" for c in ALL_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
row += 1
put_row(ws, row, f"{IDT}OPM (%)",
        {c: f"={c}{R['영업이익']}/{c}{R['매출액']}" for c in ALL_COLS}, font=F_PCT, fmt=PCT)
row += 1
R["기타수익"] = row; put_row(ws, row, "기타수익", is_base_row(row, "기타수익"))
row += 1
R["기타비용"] = row; put_row(ws, row, "기타비용", is_base_row(row, "기타비용"))
row += 1
R["금융수익"] = row; put_row(ws, row, "금융수익", is_base_row(row, "금융수익"))
row += 1
R["금융비용"] = row; put_row(ws, row, "금융비용", is_base_row(row, "금융비용"))
row += 1
R["법차전"] = row
put_row(ws, row, "법인세비용차감전순이익",
        {c: f"={c}{R['영업이익']}+{c}{R['기타수익']}-{c}{R['기타비용']}+{c}{R['금융수익']}-{c}{R['금융비용']}" for c in ALL_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
row += 1
R["법인세"] = row; put_row(ws, row, "법인세비용", is_base_row(row, "법인세비용"))
row += 1
R["당기순이익"] = row
put_row(ws, row, "당기순이익",
        {c: f"={c}{R['법차전']}-{c}{R['법인세']}" for c in ALL_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
row += 1
put_row(ws, row, f"{IDT}NPM (%)",
        {c: f"={c}{R['당기순이익']}/{c}{R['매출액']}" for c in ALL_COLS}, font=F_PCT, fmt=PCT)
IS_NI_ROW = R["당기순이익"]
IS_COGS_ROW = R["매출원가"]
IS_SGA_ROW = R["판관비"]
note_row = row + 2
ws[f"B{note_row}"] = "※ 4Q 컬럼은 사업보고서 연간 실적 − (1Q+2Q+3Q)로 자동 도출. 음영(파랑) 행은 수식으로 계산되는 소계."
style_cell(ws[f"B{note_row}"], F_NOTE)

# ============================== COST ==============================
ws = wb.create_sheet("COST")
ws.sheet_properties.tabColor = "C55A11"
Y_ONLY = list(Y_COLS.values())
sheet_header(ws, "COST — 비용 구조 (주석)", "(단위: 억원)", Y_ONLY, YEARS)

row = 7
ws[f"B{row}"] = "① 비용의 성격별 분류 (매출원가 + 판관비)"
style_cell(ws[f"B{row}"], F_BOLD)
row += 1
nature_start = row
totals = [r1(IS_Y["매출원가"][i] + IS_Y["판매비와관리비"][i]) for i in range(5)]
others = list(totals)
for item, vals in COST_NATURE_BASE.items():
    put_row(ws, row, f"{IDT}{item}", {Y_COLS[YEARS[i]]: vals[i] for i in range(5)})
    others = [r1(others[i] - vals[i]) for i in range(5)]
    row += 1
put_row(ws, row, f"{IDT}기타", {Y_COLS[YEARS[i]]: others[i] for i in range(5)})
row += 1
nature_end = row - 1
put_row(ws, row, "합계 (매출원가+판관비)",
        {c: f"=SUM({c}{nature_start}:{c}{nature_end})" for c in Y_ONLY},
        bold_label=True, fill=FILL_SUBTOTAL)
total_row = row
row += 1
put_row(ws, row, f"{IDT}CHK (IS와 대사)",
        {c: f'=IF(ROUND({c}{total_row}-(IS!{c}{IS_COGS_ROW}+IS!{c}{IS_SGA_ROW}),1)=0,"OK","CHK")' for c in Y_ONLY},
        font=F_PCT, fmt="General")
row += 3

ws[f"B{row}"] = "② 판매비와관리비 세부"
style_cell(ws[f"B{row}"], F_BOLD)
row += 1
sga_start = row
others = [IS_Y["판매비와관리비"][i] for i in range(5)]
for item, vals in SGA_BASE.items():
    put_row(ws, row, f"{IDT}{item}", {Y_COLS[YEARS[i]]: vals[i] for i in range(5)})
    others = [r1(others[i] - vals[i]) for i in range(5)]
    row += 1
put_row(ws, row, f"{IDT}기타", {Y_COLS[YEARS[i]]: others[i] for i in range(5)})
row += 1
sga_end = row - 1
put_row(ws, row, "판관비 합계",
        {c: f"=SUM({c}{sga_start}:{c}{sga_end})" for c in Y_ONLY},
        bold_label=True, fill=FILL_SUBTOTAL)
sga_total_row = row
row += 1
put_row(ws, row, f"{IDT}CHK (IS와 대사)",
        {c: f'=IF(ROUND({c}{sga_total_row}-IS!{c}{IS_SGA_ROW},1)=0,"OK","CHK")' for c in Y_ONLY},
        font=F_PCT, fmt="General")
row += 2
ws[f"B{row}"] = "※ 비용의 성격별 분류·판관비 세부는 사업보고서(연간) 주석에서 추출. 분기 주석은 공시 범위가 제한적이라 연간만 제공."
style_cell(ws[f"B{row}"], F_NOTE)

# ============================== BS ==============================
ws = wb.create_sheet("BS")
ws.sheet_properties.tabColor = "2E7D32"
sheet_header(ws, "BS — 재무상태표 (연결)", "(단위: 억원)", ALL_COLS, YEARS + QTRS)

def bs_vals(item, row):
    vals = {}
    for i, y in enumerate(YEARS):
        vals[Y_COLS[y]] = BS_Y_FULL[i][item]
    for q in ["1Q24", "2Q24", "3Q24", "1Q25", "2Q25", "3Q25"]:
        vals[Q_COLS[q]] = BS_Q_FULL[q][item]
    vals[Q_COLS["4Q24"]] = f"={Y_COLS[2024]}{row}"   # 연간 기말과 동일
    vals[Q_COLS["4Q25"]] = f"={Y_COLS[2025]}{row}"
    return vals

row = 7
bs_R = {}
def bs_section(title, items, row):
    start = row + 1
    section_row = row
    r = row + 1
    for it in items:
        put_row(ws, r, f"{IDT}{IDT}{it}", bs_vals(it, r))
        r += 1
    put_row(ws, section_row, f"{IDT}{title}",
            {c: f"=SUM({c}{start}:{c}{r-1})" for c in ALL_COLS},
            bold_label=True, fill=FILL_SUBTOTAL)
    return section_row, r

ws[f"B{row}"] = "자산"; style_cell(ws[f"B{row}"], F_BOLD); row += 1
cur_a_row, row = bs_section("유동자산", BS_CUR_ASSET, row)
noncur_a_row, row = bs_section("비유동자산", BS_NONCUR_ASSET, row)
put_row(ws, row, "자산총계", {c: f"={c}{cur_a_row}+{c}{noncur_a_row}" for c in ALL_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
asset_total_row = row
row += 2

ws[f"B{row}"] = "부채"; style_cell(ws[f"B{row}"], F_BOLD); row += 1
cur_l_row, row = bs_section("유동부채", BS_CUR_LIAB, row)
noncur_l_row, row = bs_section("비유동부채", BS_NONCUR_LIAB, row)
put_row(ws, row, "부채총계", {c: f"={c}{cur_l_row}+{c}{noncur_l_row}" for c in ALL_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
liab_total_row = row
row += 2

ws[f"B{row}"] = "자본"; style_cell(ws[f"B{row}"], F_BOLD); row += 1
eq_start = row
for it in BS_EQUITY_FIXED + ["이익잉여금"]:
    put_row(ws, row, f"{IDT}{IDT}{it}", bs_vals(it, row))
    row += 1
put_row(ws, row, "자본총계", {c: f"=SUM({c}{eq_start}:{c}{row-1})" for c in ALL_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
eq_total_row = row
row += 1
put_row(ws, row, "부채와자본총계",
        {c: f"={c}{liab_total_row}+{c}{eq_total_row}" for c in ALL_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
le_total_row = row
row += 1
put_row(ws, row, f"{IDT}CHK (자산=부채+자본)",
        {c: f'=IF(ROUND({c}{asset_total_row}-{c}{le_total_row},1)=0,"OK","CHK")' for c in ALL_COLS},
        font=F_PCT, fmt="General")
BS_CASH_ROW = cur_a_row + 1  # 현금및현금성자산 = 유동자산 바로 아래 첫 항목
row += 2
ws[f"B{row}"] = "※ 분기 컬럼은 분기보고서의 기말 잔액. 4Q 컬럼은 사업보고서 연말 잔액과 동일(수식 참조)."
style_cell(ws[f"B{row}"], F_NOTE)

# ============================== CF ==============================
ws = wb.create_sheet("CF")
ws.sheet_properties.tabColor = "6A1B9A"
sheet_header(ws, "CF — 현금흐름표 (연결)", "(단위: 억원)", ALL_COLS, YEARS + QTRS)

def cf_vals(item, row, link_is=False):
    vals = {}
    for i, y in enumerate(YEARS):
        vals[Y_COLS[y]] = CF_Y[i][item]
    for q in ["1Q24", "2Q24", "3Q24", "1Q25", "2Q25", "3Q25"]:
        vals[Q_COLS[q]] = CF_Q[q][item]
    for q, (ycol, qcols) in Q4_MAP.items():
        col = Q_COLS[q]
        vals[col] = f"={ycol}{row}-{qcols[0]}{row}-{qcols[1]}{row}-{qcols[2]}{row}"
    if link_is:
        vals = {c: f"=IS!{c}{IS_NI_ROW}" for c in ALL_COLS}
    return vals

row = 7
ops_row = row
ops_items = ["당기순이익", "비현금항목 조정", "영업활동 자산부채의 증감", "이자수취", "이자지급", "법인세납부"]
r = row + 1
for it in ops_items:
    put_row(ws, r, f"{IDT}{it}", cf_vals(it, r, link_is=(it == "당기순이익")))
    r += 1
put_row(ws, ops_row, "영업활동현금흐름",
        {c: f"=SUM({c}{ops_row+1}:{c}{r-1})" for c in ALL_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
row = r + 1

inv_row = row
inv_items = ["단기금융상품의 순증감", "유형자산의 취득", "유형자산의 처분", "무형자산의 취득"]
r = row + 1
for it in inv_items:
    put_row(ws, r, f"{IDT}{it}", cf_vals(it, r))
    r += 1
put_row(ws, inv_row, "투자활동현금흐름",
        {c: f"=SUM({c}{inv_row+1}:{c}{r-1})" for c in ALL_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
row = r + 1

fin_row = row
fin_items = ["차입금의 차입", "차입금의 상환", "리스부채의 상환", "배당금의 지급"]
r = row + 1
for it in fin_items:
    put_row(ws, r, f"{IDT}{it}", cf_vals(it, r))
    r += 1
put_row(ws, fin_row, "재무활동현금흐름",
        {c: f"=SUM({c}{fin_row+1}:{c}{r-1})" for c in ALL_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
row = r + 1

fx_row = row
put_row(ws, row, "외화환산으로 인한 현금의 변동", cf_vals("외화환산으로 인한 현금의 변동", row))
row += 1
delta_row = row
put_row(ws, row, "현금및현금성자산의 증감",
        {c: f"={c}{ops_row}+{c}{inv_row}+{c}{fin_row}+{c}{fx_row}" for c in ALL_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
row += 1
begin_row = row
begin_vals = cf_vals("기초현금", row)
begin_vals[Q_COLS["4Q24"]] = f"={Q_COLS['3Q24']}{row+1}"
begin_vals[Q_COLS["4Q25"]] = f"={Q_COLS['3Q25']}{row+1}"
put_row(ws, row, "기초현금및현금성자산", begin_vals)
row += 1
end_row = row
put_row(ws, row, "기말현금및현금성자산",
        {c: f"={c}{begin_row}+{c}{delta_row}" for c in ALL_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
row += 1
put_row(ws, row, f"{IDT}CHK (기말현금=BS 현금)",
        {c: f'=IF(ROUND({c}{end_row}-BS!{c}{BS_CASH_ROW},1)=0,"OK","CHK")' for c in ALL_COLS},
        font=F_PCT, fmt="General")
row += 2
ws[f"B{row}"] = "※ 분기보고서 현금흐름표는 누적(YTD) 공시 → 직전 분기 누적을 차감해 3개월 값으로 환산. 4Q는 연간 − (1Q+2Q+3Q)."
style_cell(ws[f"B{row}"], F_NOTE)

# ============================== Valuation ==============================
ws = wb.create_sheet("Valuation")
ws.sheet_properties.tabColor = "B7950B"
ws.sheet_view.showGridLines = False
ws["B2"] = "Valuation"
ws["B2"].font = F_TITLE
ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 30
for col in ["C", "D", "E", "F", "G", "H"]:
    ws.column_dimensions[col].width = 13

ws["B4"] = "주주환원 — 배당 / 자사주"
style_cell(ws["B4"], F_BOLD)
val_years = [2022, 2023, 2024, 2025]
ws["B6"] = "구분"
style_cell(ws["B6"], F_HDR, fill=FILL_HDR, border=True, align="center")
for i, y in enumerate(val_years):
    c = ws.cell(row=6, column=3 + i)
    c.value = f"FY{y}"
    style_cell(c, F_HDR, fill=FILL_HDR, border=True, align="center")
put = lambda r, label, vals, fmt=NUM, fill=None, font=F_BASE: (
    ws.__setitem__(f"B{r}", label),
    style_cell(ws[f"B{r}"], F_BASE, border=True),
    [(
        ws.cell(row=r, column=3 + i).__setattr__("value", v),
        style_cell(ws.cell(row=r, column=3 + i), font, fmt=fmt, fill=fill, border=True),
    ) for i, v in enumerate(vals)],
)
put(7, "당기순이익 (억원)", [f"=IS!{Y_COLS[y]}{IS_NI_ROW}" for y in val_years])
put(8, "가중평균 유통주식수 (주)", [20_000_000] * 4, fmt=INT, fill=FILL_INPUT)
put(9, "EPS (원)", [f"={get_column_letter(3+i)}7*10^8/{get_column_letter(3+i)}8" for i in range(4)], fmt=INT)

ws["B11"] = "배당정책 (공시 내용 입력)"
style_cell(ws["B11"], F_BOLD)
for i, t in enumerate([
    "• 정책 적용기간: (입력)",
    "• 배당 기준: (입력)",
    "• 주당 최저배당액: (입력)",
    "• 자사주 소각 계획: (입력)",
]):
    ws[f"B{12+i}"] = t
    style_cell(ws[f"B{12+i}"], F_NOTE)

ws["B17"] = "PER 밴드 (FY2025 EPS 기준)"
style_cell(ws["B17"], F_BOLD)
ws["B18"] = "PER (배)"
ws["C18"] = "적정주가 (원)"
for cell in ["B18", "C18"]:
    style_cell(ws[cell], F_HDR, fill=FILL_HDR, border=True, align="center")
for i, per in enumerate([6, 7, 8, 9, 10, 11]):
    r = 19 + i
    ws[f"B{r}"] = per
    style_cell(ws[f"B{r}"], F_BASE, fmt="0", fill=FILL_INPUT, border=True, align="center")
    ws[f"C{r}"] = f"=$F$9*B{r}"
    style_cell(ws[f"C{r}"], F_BASE, fmt=INT, border=True)
ws["B26"] = "현재주가 (원)"
style_cell(ws["B26"], F_BASE, border=True)
ws["C26"] = 8500
style_cell(ws["C26"], F_BASE, fmt=INT, fill=FILL_INPUT, border=True)
ws["B27"] = "Upside vs PER 8x"
style_cell(ws["B27"], F_BASE, border=True)
ws["C27"] = "=C21/C26-1"
style_cell(ws["C27"], F_BASE, fmt=PCT, border=True)
ws["B29"] = "※ 노란색 셀은 직접 입력. 당기순이익은 IS 시트 자동 연결. (첫 번째 파일 Valuation 시트 양식 참고)"
style_cell(ws["B29"], F_NOTE)

# ============================== Check (차이내역) ==============================
ws = wb.create_sheet("Check")
ws.sheet_properties.tabColor = "C00000"
ws.sheet_view.showGridLines = False
ws["B2"] = "Check — 보고서 간 값이 달라진 항목"
ws["B2"].font = F_TITLE
ws["B4"] = "같은 기간·같은 항목의 숫자가 보고서마다 다른 경우, 최신 보고서 값을 채택하고 아래에 기록합니다. (재작성·재분류 등)"
style_cell(ws["B4"], F_NOTE)

headers = ["시트", "항목", "기간", "이전 값", "이전 값 출처", "채택 값 (최신)", "채택 값 출처", "차이", "차이율"]
ws.column_dimensions["A"].width = 2
widths = [8, 24, 8, 12, 26, 14, 26, 10, 9]
for i, (h, w) in enumerate(zip(headers, widths)):
    col = get_column_letter(2 + i)
    ws.column_dimensions[col].width = w
    c = ws[f"{col}6"]
    c.value = h
    style_cell(c, F_HDR, fill=FILL_HDR, border=True, align="center")

diff_rows = [
    ["IS", "매출원가", "FY2023", 603.1, "FY2023 사업보고서", 601.3, "FY2024 사업보고서 (재분류)", None, None],
    ["BS", "매출채권 및 기타채권", "FY2022", 110.8, "FY2022 사업보고서", 112.4, "FY2024 사업보고서 (재작성)", None, None],
    ["CF", "이자지급", "FY2023", -9.8, "FY2023 사업보고서", -10.1, "FY2024 사업보고서 (표시방법 변경)", None, None],
]
for j, dr in enumerate(diff_rows):
    r = 7 + j
    for i, v in enumerate(dr):
        col = get_column_letter(2 + i)
        cell = ws[f"{col}{r}"]
        if i == 7:
            cell.value = f"=G{r}-E{r}"
        elif i == 8:
            cell.value = f"=G{r}/E{r}-1"
        else:
            cell.value = v
        fmt = NUM if i in (3, 5, 7) else (PCT if i == 8 else "General")
        style_cell(cell, F_BASE, fmt=fmt, border=True)

ws["B12"] = "※ 예시 데이터입니다. 실제 추출기는 발견된 모든 차이를 여기에 나열합니다."
style_cell(ws["B12"], F_NOTE)

# 저장
import sys
out = sys.argv[1] if len(sys.argv) > 1 else "재무제표_추출_예시.xlsx"
wb.save(out)
print(f"saved: {out}")
