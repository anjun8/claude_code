# -*- coding: utf-8 -*-
"""
재무제표 추출기 — 결과물 예시 엑셀 생성 스크립트 (v3)

사업보고서/분기보고서를 넣으면 만들어질 최종 엑셀의 "모양"을 미리 보여주기 위한
샘플 데이터 기반 생성기.

시트 구성: IS / COST / BS / CF / Valuation (+ Check: 보고서 간 값이 달라진 항목 알림)

컬럼 레이아웃
  IS·COST·CF : 분기 13개(1Q23~1Q26, 4Q는 연간-3개분기 도출) → 한 칸 공백 → 연간 FY2021~FY2025
  BS         : 시간 순서 FY21, FY22, 1Q23, 2Q23, 3Q23, FY23, ... , FY25, 1Q26
  Valuation  : 추정 손익계산서 + 매출 추정 (SNUVALUE Financial 양식) + 주주환원 + 밸류에이션 밴드

디자인: 진한 빨강 타이틀 / 짙은 회색 기간 헤더 / 회색 구분 행 / 주황 주요 항목 /
        음수는 빨간 괄호 / 단위: 소수점 1자리 억원
"""

import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ──────────────────────────────────────────────────────────────
# 공통 스타일 (스크린샷 색감)
# ──────────────────────────────────────────────────────────────
C_RED = "C00000"       # 타이틀 블록
C_DARK = "404040"      # 기간 헤더
C_GRAY = "D9D9D9"      # 구분(섹션) 행
C_ORANGE = "FCE4D6"    # 주요 항목(소계/합계) 행
C_INPUT = "FFF2CC"     # 입력 셀

F_BASE = Font(name="맑은 고딕", size=10)
F_TITLE = Font(name="맑은 고딕", size=12, bold=True, color="FFFFFF")
F_HDR = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
F_SRC = Font(name="맑은 고딕", size=8, italic=True, color="808080")
F_BOLD = Font(name="맑은 고딕", size=10, bold=True)
F_PCT = Font(name="맑은 고딕", size=9, italic=True, color="808080")
F_NOTE = Font(name="맑은 고딕", size=9, color="808080")
F_SECTION = Font(name="맑은 고딕", size=10, bold=True, color="000000")

FILL_TITLE = PatternFill("solid", fgColor=C_RED)
FILL_HDR = PatternFill("solid", fgColor=C_DARK)
FILL_SECTION = PatternFill("solid", fgColor=C_GRAY)
FILL_KEY = PatternFill("solid", fgColor=C_ORANGE)
FILL_INPUT = PatternFill("solid", fgColor=C_INPUT)

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

NUM = "#,##0.0;[Red](#,##0.0)"      # 억원, 소수점 1자리, 음수 빨간 괄호
PCT = "0.0%;[Red](0.0%)"
INT = "#,##0;[Red](#,##0)"

IDT = "　"  # 전각 공백 들여쓰기

# ──────────────────────────────────────────────────────────────
# 기간 / 컬럼 레이아웃
# ──────────────────────────────────────────────────────────────
YEARS = [2021, 2022, 2023, 2024, 2025]
Q_LABELS = ["1Q23", "2Q23", "3Q23", "4Q23", "1Q24", "2Q24", "3Q24", "4Q24",
            "1Q25", "2Q25", "3Q25", "4Q25", "1Q26"]          # 13개 분기
EXTRACTED_Q = [q for q in Q_LABELS if not q.startswith("4Q")]  # 4Q는 도출

# IS·COST·CF: 분기 C~O(13개) / P 공백 / 연간 Q~U
ISCF_Q_COLS = {q: get_column_letter(3 + i) for i, q in enumerate(Q_LABELS)}
ISCF_Y_COLS = {y: get_column_letter(17 + i) for i, y in enumerate(YEARS)}   # Q..U
ISCF_COLS = list(ISCF_Q_COLS.values()) + list(ISCF_Y_COLS.values())
GAP_COL = get_column_letter(16)  # P
Q4_MAP = {
    "4Q23": (ISCF_Y_COLS[2023], [ISCF_Q_COLS["1Q23"], ISCF_Q_COLS["2Q23"], ISCF_Q_COLS["3Q23"]]),
    "4Q24": (ISCF_Y_COLS[2024], [ISCF_Q_COLS["1Q24"], ISCF_Q_COLS["2Q24"], ISCF_Q_COLS["3Q24"]]),
    "4Q25": (ISCF_Y_COLS[2025], [ISCF_Q_COLS["1Q25"], ISCF_Q_COLS["2Q25"], ISCF_Q_COLS["3Q25"]]),
}
Q_YEAR_IDX = {"23": 2, "24": 3, "25": 4, "26": 4}  # 분기 → 연도 인덱스(비중 산출용, 26은 25 비중)

# BS: 시간 순서
BS_ORDER = ["FY2021", "FY2022",
            "1Q23", "2Q23", "3Q23", "FY2023",
            "1Q24", "2Q24", "3Q24", "FY2024",
            "1Q25", "2Q25", "3Q25", "FY2025",
            "1Q26"]
BS_COLS = {p: get_column_letter(3 + i) for i, p in enumerate(BS_ORDER)}     # C..Q

SOURCES = {
    "FY2021": "FY2023 사업보고서 (전전기)",
    "FY2022": "FY2024 사업보고서 (전전기)",
    "FY2023": "FY2025 사업보고서 (전전기)",
    "FY2024": "FY2025 사업보고서 (전기)",
    "FY2025": "FY2025 사업보고서 (당기)",
    "1Q23": "2024.1Q 분기보고서 (전년비교)",
    "2Q23": "2024 반기보고서 (전년비교)",
    "3Q23": "2024.3Q 분기보고서 (전년비교)",
    "4Q23": "도출: FY2023 − (1Q+2Q+3Q)",
    "1Q24": "2025.1Q 분기보고서 (전년비교)",
    "2Q24": "2025 반기보고서 (전년비교)",
    "3Q24": "2025.3Q 분기보고서 (전년비교)",
    "4Q24": "도출: FY2024 − (1Q+2Q+3Q)",
    "1Q25": "2026.1Q 분기보고서 (전년비교)",
    "2Q25": "2025 반기보고서 (당기)",
    "3Q25": "2025.3Q 분기보고서 (당기)",
    "4Q25": "도출: FY2025 − (1Q+2Q+3Q)",
    "1Q26": "2026.1Q 분기보고서 (당기)",
}

SHARES = 20_000_000  # 주당이익 계산용 (예시)

# ──────────────────────────────────────────────────────────────
# 샘플 데이터 (단위: 억원) — 가상의 회사
# ──────────────────────────────────────────────────────────────
IS_Y = {
    "매출액":     [812.4, 934.1, 1023.5, 1187.9, 1352.6],
    "매출원가":   [487.2, 549.8, 601.3, 688.5, 769.4],
    "판매비와관리비": [214.6, 241.3, 265.8, 301.2, 338.7],
    "금융수익":   [6.4, 8.9, 12.3, 15.1, 13.8],
    "금융비용":   [9.8, 11.2, 10.6, 9.4, 8.2],
    "관계기업투자이익(손실)": [-0.5, 0.3, 0.2, 0.0, 0.4],
    "기타수익":   [3.2, 4.1, 2.8, 5.6, 4.9],
    "기타비용":   [5.1, 3.8, 6.2, 4.3, 7.1],
    "법인세비용": [21.3, 28.9, 31.5, 42.8, 51.6],
    "비지배지분순이익": [0.0, 0.0, 0.0, 0.0, 0.0],
    "OCI재분류불가": [0.2, -0.3, 0.1, 0.4, -0.2],
    "OCI재분류가능": [0.1, 0.0, -0.1, 0.2, 0.1],
}
# 추출 분기 순서: 1Q23,2Q23,3Q23, 1Q24,2Q24,3Q24, 1Q25,2Q25,3Q25, 1Q26
IS_Q_RAW = {
    "매출액":     [238.5, 252.1, 259.8, 275.3, 291.8, 298.4, 312.7, 334.2, 341.5, 348.2],
    "매출원가":   [141.2, 147.9, 152.4, 162.1, 169.5, 172.8, 180.3, 190.1, 194.6, 199.5],
    "판매비와관리비": [62.4, 65.3, 66.9, 71.2, 74.8, 76.1, 80.9, 83.6, 85.2, 89.4],
    "금융수익":   [2.9, 3.1, 3.0, 3.6, 3.9, 3.8, 3.4, 3.5, 3.3, 3.6],
    "금융비용":   [2.8, 2.7, 2.6, 2.5, 2.4, 2.3, 2.1, 2.0, 2.1, 1.9],
    "관계기업투자이익(손실)": [0.1, 0.0, 0.1, 0.0, 0.0, 0.0, 0.1, 0.2, 0.1, 0.1],
    "기타수익":   [0.6, 0.8, 0.7, 1.2, 1.5, 1.1, 1.0, 1.4, 1.2, 1.3],
    "기타비용":   [1.4, 1.6, 1.5, 0.9, 1.1, 1.3, 1.8, 1.6, 2.1, 1.5],
    "법인세비용": [7.4, 8.1, 7.8, 10.2, 11.5, 10.8, 12.4, 13.7, 12.9, 14.2],
    "비지배지분순이익": [0.0] * 10,
    "OCI재분류불가": [0.0, 0.1, -0.1, 0.1, 0.2, 0.0, -0.1, 0.0, 0.0, 0.1],
    "OCI재분류가능": [0.0, 0.0, 0.0, 0.1, 0.0, 0.1, 0.0, 0.1, 0.0, 0.0],
}
IS_Q = {item: dict(zip(EXTRACTED_Q, vals)) for item, vals in IS_Q_RAW.items()}

def r1(x):
    return round(x + 1e-9, 1)

def is_ni_from(vals):  # 당기순이익 (사업보고서 손익 구조)
    gp = vals["매출액"] - vals["매출원가"]
    op = gp - vals["판매비와관리비"]
    pbt = (op + vals["금융수익"] - vals["금융비용"] + vals["관계기업투자이익(손실)"]
           + vals["기타수익"] - vals["기타비용"])
    return r1(pbt - vals["법인세비용"])

NI_Y = [is_ni_from({k: IS_Y[k][i] for k in IS_Y}) for i in range(5)]
NI_Q = {q: is_ni_from({k: IS_Q[k][q] for k in IS_Q}) for q in EXTRACTED_Q}
EPS_Y = [round(NI_Y[i] * 1e8 / SHARES) for i in range(5)]
EPS_Q = {q: round(NI_Q[q] * 1e8 / SHARES) for q in EXTRACTED_Q}
IS_Y["기본주당이익(원)"] = EPS_Y
IS_Y["희석주당이익(원)"] = EPS_Y
IS_Q["기본주당이익(원)"] = EPS_Q
IS_Q["희석주당이익(원)"] = dict(EPS_Q)

# COST ①: 비용의 성격별 분류 (합계 = 매출원가+판관비), '기타'가 plug
COST_NATURE_BASE = {
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
# COST ②: 판관비 세부 (합계 = IS 판관비), '기타'가 plug
SGA_LABOR = {
    "급여": [46.2, 52.8, 58.9, 68.3, 78.4],
    "퇴직급여": [5.4, 6.1, 6.8, 7.9, 8.9],
    "복리후생비": [7.8, 8.6, 9.7, 11.2, 12.6],
}
SGA_OTHER = {
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
COGS_MATCH = {  # ③ = ① − ② 매칭 (종업원급여 ↔ ②인건비 소계)
    "종업원급여": "인건비",
    "감가상각비": "감가상각비",
    "무형자산상각비": "무형자산상각비",
    "지급수수료": "지급수수료",
    "임차료 및 관리비": "임차료 및 관리비",
    "광고선전비": "광고선전비",
    "운반비": "운반비",
    "소모품비": "소모품비",
}

# 분기 COST: 해당 분기 (매출원가+판관비)/판관비 총액을 그 해 연간 구성비로 배분, '기타'가 plug
COST_N_Q = {item: {} for item in COST_NATURE_BASE}
COST_N_Q["기타"] = {}
SGA_Q = {item: {} for item in list(SGA_LABOR) + list(SGA_OTHER)}
SGA_Q["기타"] = {}
NATURE_TOT_Y = [r1(IS_Y["매출원가"][i] + IS_Y["판매비와관리비"][i]) for i in range(5)]
for q in EXTRACTED_Q:
    yi = Q_YEAR_IDX[q[-2:]]
    q_tot = r1(IS_Q["매출원가"][q] + IS_Q["판매비와관리비"][q])
    ratio = q_tot / NATURE_TOT_Y[yi]
    rest = q_tot
    for item, vals in COST_NATURE_BASE.items():
        v = r1(vals[yi] * ratio)
        COST_N_Q[item][q] = v
        rest = r1(rest - v)
    COST_N_Q["기타"][q] = rest

    q_sga = IS_Q["판매비와관리비"][q]
    ratio = q_sga / IS_Y["판매비와관리비"][yi]
    rest = q_sga
    for item, vals in list(SGA_LABOR.items()) + list(SGA_OTHER.items()):
        v = r1(vals[yi] * ratio)
        SGA_Q[item][q] = v
        rest = r1(rest - v)
    SGA_Q["기타"][q] = rest

# BS 연말 잔액 — 실제 추출 시 사업보고서상 계정 전부를 가져옴
BS_Y = {
    "현금및현금성자산": [72.4, 85.1, 96.8, 118.3, 142.6],
    "단기금융상품": [45.0, 52.3, 60.1, 71.5, 85.2],
    "매출채권 및 기타채권": [98.6, 112.4, 121.9, 139.7, 158.3],
    "재고자산": [76.2, 88.9, 95.4, 108.2, 121.8],
    "당기법인세자산": [1.2, 1.5, 0.9, 1.8, 2.1],
    "기타유동자산": [8.3, 9.7, 10.5, 12.1, 13.9],
    "유형자산": [156.8, 172.3, 189.5, 212.6, 238.4],
    "사용권자산": [18.2, 20.5, 22.8, 25.4, 28.1],
    "무형자산": [34.5, 36.8, 39.2, 42.5, 45.9],
    "투자부동산": [10.0, 10.0, 9.5, 9.5, 9.0],
    "관계기업투자": [5.0, 5.5, 6.1, 6.8, 7.4],
    "기타포괄손익-공정가치측정금융자산": [3.2, 3.5, 4.1, 4.6, 5.2],
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
BS_CUR_ASSET = ["현금및현금성자산", "단기금융상품", "매출채권 및 기타채권", "재고자산",
                "당기법인세자산", "기타유동자산"]
BS_NONCUR_ASSET = ["유형자산", "사용권자산", "무형자산", "투자부동산", "관계기업투자",
                   "기타포괄손익-공정가치측정금융자산", "장기금융상품", "이연법인세자산"]
BS_ASSET_ITEMS = BS_CUR_ASSET + BS_NONCUR_ASSET
BS_CUR_LIAB = ["매입채무 및 기타채무", "단기차입금", "유동성리스부채", "당기법인세부채", "기타유동부채"]
BS_NONCUR_LIAB = ["장기차입금", "리스부채", "확정급여부채", "기타비유동부채"]
BS_EQUITY_FIXED = ["자본금", "자본잉여금", "기타자본구성요소"]

CASH_2020 = 58.3

def bs_totals(d):
    a = sum(d[k] for k in BS_ASSET_ITEMS)
    l = sum(d[k] for k in BS_CUR_LIAB + BS_NONCUR_LIAB)
    eq_fixed = sum(d[k] for k in BS_EQUITY_FIXED)
    return r1(a), r1(l), r1(a - l - eq_fixed)

BS_FY = {}
for i, y in enumerate(YEARS):
    d = {k: BS_Y[k][i] for k in BS_Y}
    _, _, re = bs_totals(d)
    d["이익잉여금"] = re
    BS_FY[f"FY{y}"] = d

WIGGLE = [0.018, -0.011, 0.014]
BS_Q = {}
for y_idx, tag in [(2, "23"), (3, "24"), (4, "25")]:
    prev, cur = BS_FY[f"FY{YEARS[y_idx-1]}"], BS_FY[f"FY{YEARS[y_idx]}"]
    for qn in range(1, 4):
        d = {}
        for k in BS_Y:
            base = prev[k] + (cur[k] - prev[k]) * qn / 4
            d[k] = r1(base * (1 + WIGGLE[qn - 1]))
        _, _, re = bs_totals(d)
        d["이익잉여금"] = re
        BS_Q[f"{qn}Q{tag}"] = d
d = {}
for k in BS_Y:
    base = BS_FY["FY2025"][k] + (BS_FY["FY2025"][k] - BS_FY["FY2024"][k]) * 0.25
    d[k] = r1(base * (1 + WIGGLE[0]))
_, _, re = bs_totals(d)
d["이익잉여금"] = re
BS_Q["1Q26"] = d

def bs_value(period, item):
    return (BS_FY.get(period) or BS_Q[period])[item]

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

cf_params_y = {
    2021: (62.4, 5.8, -9.2, -19.5, -6.2, -38.5, 2.1, -8.4, 20.0, -35.0, -4.5, -12.0, 0.8),
    2022: (70.2, 8.1, -10.6, -26.4, -7.3, -42.8, 1.5, -9.2, 15.0, -30.0, -5.0, -14.0, -1.2),
    2023: (75.6, 11.4, -10.1, -29.8, -7.8, -48.2, 3.2, -10.5, 10.0, -25.0, -5.5, -16.0, 0.5),
    2024: (84.3, 14.2, -8.9, -38.6, -11.4, -55.4, 1.8, -11.8, 10.0, -25.0, -6.1, -18.0, -0.7),
    2025: (94.8, 12.9, -7.8, -47.2, -13.7, -63.2, 2.4, -12.6, 5.0, -20.0, -6.7, -21.0, 1.1),
}
CF_FY = {}
for i, y in enumerate(YEARS):
    begin = CASH_2020 if i == 0 else BS_FY[f"FY{YEARS[i-1]}"]["현금및현금성자산"]
    end = BS_FY[f"FY{y}"]["현금및현금성자산"]
    CF_FY[f"FY{y}"] = build_cf(NI_Y[i], *cf_params_y[y], cash_begin=begin, cash_end=end)

cf_params_q = {
    "1Q23": (18.5, 2.8, -2.6, -7.2, -1.9, -11.8, 0.8, -2.5, 3.0, -6.0, -1.3, -14.0, 0.2),
    "2Q23": (19.0, 2.9, -2.5, -7.6, -2.0, -12.1, 0.7, -2.6, 2.0, -6.5, -1.4, 0.0, -0.3),
    "3Q23": (19.2, 2.9, -2.5, -7.4, -2.0, -12.3, 0.9, -2.6, 3.0, -6.0, -1.4, 0.0, 0.4),
    "1Q24": (20.5, 3.4, -2.3, -9.1, -2.8, -13.2, 0.4, -2.9, 5.0, -6.0, -1.5, -16.0, -0.2),
    "2Q24": (21.1, 3.6, -2.2, -10.2, -2.9, -13.8, 0.5, -3.0, 2.0, -6.0, -1.5, 0.0, 0.3),
    "3Q24": (21.3, 3.5, -2.2, -9.8, -3.0, -14.1, 0.4, -2.9, 3.0, -6.5, -1.5, 0.0, -0.4),
    "1Q25": (23.2, 3.3, -2.0, -11.5, -3.3, -15.4, 0.6, -3.1, 2.0, -5.0, -1.6, -18.0, 0.3),
    "2Q25": (23.7, 3.2, -1.9, -12.1, -3.4, -15.9, 0.5, -3.2, 1.0, -5.0, -1.7, 0.0, -0.3),
    "3Q25": (23.9, 3.2, -2.0, -11.8, -3.5, -16.2, 0.7, -3.1, 2.0, -5.0, -1.7, 0.0, 0.4),
    "1Q26": (25.0, 3.1, -1.8, -12.5, -3.6, -16.8, 0.5, -3.3, 2.0, -5.0, -1.8, -21.0, 0.2),
}
Q_BEGIN = {
    "1Q23": "FY2022", "2Q23": "1Q23", "3Q23": "2Q23",
    "1Q24": "FY2023", "2Q24": "1Q24", "3Q24": "2Q24",
    "1Q25": "FY2024", "2Q25": "1Q25", "3Q25": "2Q25",
    "1Q26": "FY2025",
}
CF_Q = {}
for q in EXTRACTED_Q:
    begin = bs_value(Q_BEGIN[q], "현금및현금성자산")
    end = bs_value(q, "현금및현금성자산")
    CF_Q[q] = build_cf(NI_Q[q], *cf_params_q[q], cash_begin=begin, cash_end=end)

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

def sheet_title(ws, title):
    ws.sheet_view.showGridLines = False
    ws.merge_cells("B2:C2")
    ws["B2"] = title
    for col in ("B", "C"):
        style_cell(ws[f"{col}2"], F_TITLE, fill=FILL_TITLE)
    ws["B2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 22

def sheet_header(ws, title, unit, col_label_pairs, gap_cols=(), with_src=True):
    sheet_title(ws, title)
    ws["B4"] = unit
    style_cell(ws["B4"], F_HDR, fill=FILL_HDR)
    for col, label in col_label_pairs:
        c = ws[f"{col}4"]
        c.value = label
        style_cell(c, F_HDR, fill=FILL_HDR, align="center")
        if with_src:
            s = ws[f"{col}5"]
            s.value = SOURCES.get(str(label), "")
            style_cell(s, F_SRC)
            s.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.column_dimensions[col].width = 12.5
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 34
    for g in gap_cols:
        ws.column_dimensions[g].width = 2.5
        ws[f"{g}4"].fill = FILL_HDR
    if with_src:
        ws.row_dimensions[5].height = 26
    ws.freeze_panes = "C7"

def put_row(ws, row, label, values_by_col, font=F_BASE, fmt=NUM, fill=None, bold_label=False):
    c = ws[f"B{row}"]
    c.value = label
    c.font = F_BOLD if bold_label else font
    if fill:
        c.fill = fill
    for col, v in values_by_col.items():
        cell = ws[f"{col}{row}"]
        cell.value = v
        style_cell(cell, font, fmt=fmt, fill=fill)

def section_row(ws, row, label, cols):
    ws[f"B{row}"] = label
    style_cell(ws[f"B{row}"], F_SECTION, fill=FILL_SECTION)
    for col in cols:
        ws[f"{col}{row}"].fill = FILL_SECTION

# ──────────────────────────────────────────────────────────────
# 통합 문서 생성
# ──────────────────────────────────────────────────────────────
wb = openpyxl.Workbook()
wb.remove(wb.active)

ISCF_HDR = [(ISCF_Q_COLS[q], q) for q in Q_LABELS] + [(ISCF_Y_COLS[y], f"FY{y}") for y in YEARS]

def base_vals(row, y_list, q_dict):
    """추출값(분기 1~3Q·1Q26 + 연간) + 4Q 도출 수식"""
    vals = {}
    for q in EXTRACTED_Q:
        vals[ISCF_Q_COLS[q]] = q_dict[q]
    for q, (ycol, qcols) in Q4_MAP.items():
        vals[ISCF_Q_COLS[q]] = f"={ycol}{row}-{qcols[0]}{row}-{qcols[1]}{row}-{qcols[2]}{row}"
    for i, y in enumerate(YEARS):
        vals[ISCF_Y_COLS[y]] = y_list[i]
    return vals

# ============================== IS ==============================
ws = wb.create_sheet("IS")
ws.sheet_properties.tabColor = C_RED
sheet_header(ws, "IS — 손익계산서 (연결)", "(단위: 억원)", ISCF_HDR, gap_cols=(GAP_COL,))

def is_row(row, item):
    return base_vals(row, IS_Y[item], IS_Q[item])

R = {}
row = 7
R["매출액"] = row; put_row(ws, row, "매출액", is_row(row, "매출액"), bold_label=True)
row += 1
yoy = {}
for i in range(4, len(Q_LABELS)):
    cur, prev = ISCF_Q_COLS[Q_LABELS[i]], ISCF_Q_COLS[Q_LABELS[i - 4]]
    yoy[cur] = f"={cur}{R['매출액']}/{prev}{R['매출액']}-1"
for i in range(1, 5):
    a, b = ISCF_Y_COLS[YEARS[i]], ISCF_Y_COLS[YEARS[i - 1]]
    yoy[a] = f"={a}{R['매출액']}/{b}{R['매출액']}-1"
put_row(ws, row, f"{IDT}YoY (%)", yoy, font=F_PCT, fmt=PCT)
row += 1
R["매출원가"] = row; put_row(ws, row, "매출원가", is_row(row, "매출원가"))
row += 1
R["매출총이익"] = row
put_row(ws, row, "매출총이익", {c: f"={c}{R['매출액']}-{c}{R['매출원가']}" for c in ISCF_COLS},
        bold_label=True, fill=FILL_KEY)
row += 1
put_row(ws, row, f"{IDT}GPM (%)", {c: f"={c}{R['매출총이익']}/{c}{R['매출액']}" for c in ISCF_COLS},
        font=F_PCT, fmt=PCT)
row += 1
R["판관비"] = row; put_row(ws, row, "판매비와관리비", is_row(row, "판매비와관리비"))
row += 1
R["영업이익"] = row
put_row(ws, row, "영업이익(손실)", {c: f"={c}{R['매출총이익']}-{c}{R['판관비']}" for c in ISCF_COLS},
        bold_label=True, fill=FILL_KEY)
row += 1
put_row(ws, row, f"{IDT}OPM (%)", {c: f"={c}{R['영업이익']}/{c}{R['매출액']}" for c in ISCF_COLS},
        font=F_PCT, fmt=PCT)
row += 1
R["금융수익"] = row; put_row(ws, row, "금융수익", is_row(row, "금융수익"))
row += 1
R["금융비용"] = row; put_row(ws, row, "금융비용", is_row(row, "금융비용"))
row += 1
R["관계기업"] = row; put_row(ws, row, "관계기업투자이익(손실)", is_row(row, "관계기업투자이익(손실)"))
row += 1
R["기타수익"] = row; put_row(ws, row, "기타수익", is_row(row, "기타수익"))
row += 1
R["기타비용"] = row; put_row(ws, row, "기타비용", is_row(row, "기타비용"))
row += 1
R["법차전"] = row
put_row(ws, row, "법인세비용차감전순이익(손실)",
        {c: f"={c}{R['영업이익']}+{c}{R['금융수익']}-{c}{R['금융비용']}+{c}{R['관계기업']}+{c}{R['기타수익']}-{c}{R['기타비용']}" for c in ISCF_COLS},
        bold_label=True, fill=FILL_KEY)
row += 1
R["법인세"] = row; put_row(ws, row, "법인세비용(수익)", is_row(row, "법인세비용"))
row += 1
R["당기순이익"] = row
put_row(ws, row, "당기순이익(손실)", {c: f"={c}{R['법차전']}-{c}{R['법인세']}" for c in ISCF_COLS},
        bold_label=True, fill=FILL_KEY)
row += 1
put_row(ws, row, f"{IDT}NPM (%)", {c: f"={c}{R['당기순이익']}/{c}{R['매출액']}" for c in ISCF_COLS},
        font=F_PCT, fmt=PCT)
row += 1
section_row(ws, row, "당기순이익(손실)의 귀속", ISCF_COLS)
row += 1
R["비지배"] = row + 1
put_row(ws, row, f"{IDT}지배기업 소유주",
        {c: f"={c}{R['당기순이익']}-{c}{R['비지배']}" for c in ISCF_COLS})
row += 1
put_row(ws, row, f"{IDT}비지배지분", is_row(row, "비지배지분순이익"))
row += 1
R["OCI"] = row
R["OCI불가"] = row + 1
R["OCI가능"] = row + 2
put_row(ws, row, "기타포괄손익",
        {c: f"={c}{R['OCI불가']}+{c}{R['OCI가능']}" for c in ISCF_COLS}, bold_label=True)
row += 1
put_row(ws, row, f"{IDT}당기손익으로 재분류되지 않는 항목", is_row(row, "OCI재분류불가"))
row += 1
put_row(ws, row, f"{IDT}당기손익으로 재분류될 수 있는 항목", is_row(row, "OCI재분류가능"))
row += 1
R["총포괄"] = row
put_row(ws, row, "총포괄손익",
        {c: f"={c}{R['당기순이익']}+{c}{R['OCI']}" for c in ISCF_COLS},
        bold_label=True, fill=FILL_KEY)
row += 1
section_row(ws, row, "총포괄손익의 귀속", ISCF_COLS)
row += 1
put_row(ws, row, f"{IDT}지배기업 소유주",
        {c: f"={c}{R['총포괄']}-{c}{R['비지배']}" for c in ISCF_COLS})
row += 1
put_row(ws, row, f"{IDT}비지배지분", {c: f"={c}{R['비지배']}" for c in ISCF_COLS})
row += 1
section_row(ws, row, "주당이익", ISCF_COLS)
row += 1
put_row(ws, row, f"{IDT}기본주당이익(손실) (원)", is_row(row, "기본주당이익(원)"), fmt=INT)
row += 1
put_row(ws, row, f"{IDT}희석주당이익(손실) (원)", is_row(row, "희석주당이익(원)"), fmt=INT)
row += 2
ws[f"B{row}"] = "※ 4Q 컬럼은 연간(FY) 실적 − (1Q+2Q+3Q)로 자동 도출. 주황 행은 수식으로 계산되는 주요 항목. 실제 추출 시 사업보고서 손익계산서의 모든 항목 포함."
style_cell(ws[f"B{row}"], F_NOTE)
IS_NI_ROW = R["당기순이익"]
IS_COGS_ROW = R["매출원가"]
IS_SGA_ROW = R["판관비"]
IS_REV_ROW = R["매출액"]
IS_OP_ROW = R["영업이익"]
IS_PBT_ROW = R["법차전"]
IS_TAX_ROW = R["법인세"]

# ============================== COST ==============================
ws = wb.create_sheet("COST")
ws.sheet_properties.tabColor = C_RED
ws.sheet_properties.outlinePr.summaryBelow = False
sheet_header(ws, "COST — 비용 구조 (주석)", "(단위: 억원)", ISCF_HDR, gap_cols=(GAP_COL,))

def cost_vals(row, y_map, q_map):
    return base_vals(row, y_map, q_map)

row = 7
section_row(ws, row, "① 비용의 성격별 분류 (매출원가 + 판관비)", ISCF_COLS)
row += 1
nature_rows = {}
nature_start = row
others_y = list(NATURE_TOT_Y)
for item, vals in COST_NATURE_BASE.items():
    nature_rows[item] = row
    put_row(ws, row, f"{IDT}{item}", cost_vals(row, vals, COST_N_Q[item]))
    others_y = [r1(others_y[i] - vals[i]) for i in range(5)]
    row += 1
nature_rows["기타"] = row
put_row(ws, row, f"{IDT}기타", cost_vals(row, others_y, COST_N_Q["기타"]))
row += 1
nature_end = row - 1
nature_total_row = row
put_row(ws, row, "합계 (매출원가+판관비)",
        {c: f"=SUM({c}{nature_start}:{c}{nature_end})" for c in ISCF_COLS},
        bold_label=True, fill=FILL_KEY)
row += 1
put_row(ws, row, f"{IDT}CHK (IS와 대사)",
        {c: f'=IF(ROUND({c}{nature_total_row}-(IS!{c}{IS_COGS_ROW}+IS!{c}{IS_SGA_ROW}),1)=0,"OK","CHK")' for c in ISCF_COLS},
        font=F_PCT, fmt="General")
row += 2

section_row(ws, row, "② 판매비와관리비 세부", ISCF_COLS)
row += 1
sga_rows = {}
labor_total_row = row
labor_detail_start = row + 1
labor_detail_end = row + len(SGA_LABOR)
put_row(ws, labor_total_row, f"{IDT}인건비 (급여+퇴직급여+복리후생비)",
        {c: f"=SUM({c}{labor_detail_start}:{c}{labor_detail_end})" for c in ISCF_COLS},
        bold_label=True)
sga_rows["인건비"] = labor_total_row
row += 1
for item, vals in SGA_LABOR.items():
    put_row(ws, row, f"{IDT}{IDT}{item}", cost_vals(row, vals, SGA_Q[item]), font=F_PCT)
    ws.row_dimensions[row].outline_level = 1
    ws.row_dimensions[row].hidden = True
    row += 1
others_y = [IS_Y["판매비와관리비"][i] for i in range(5)]
labor_sum = [r1(sum(SGA_LABOR[k][i] for k in SGA_LABOR)) for i in range(5)]
others_y = [r1(others_y[i] - labor_sum[i]) for i in range(5)]
for item, vals in SGA_OTHER.items():
    sga_rows[item] = row
    put_row(ws, row, f"{IDT}{item}", cost_vals(row, vals, SGA_Q[item]))
    others_y = [r1(others_y[i] - vals[i]) for i in range(5)]
    row += 1
sga_rows["기타"] = row
put_row(ws, row, f"{IDT}기타", cost_vals(row, others_y, SGA_Q["기타"]))
row += 1
sga_end = row - 1
sga_total_row = row
put_row(ws, row, "판관비 합계",
        {c: f"={c}{labor_total_row}+SUM({c}{labor_detail_end+1}:{c}{sga_end})" for c in ISCF_COLS},
        bold_label=True, fill=FILL_KEY)
row += 1
put_row(ws, row, f"{IDT}CHK (IS와 대사)",
        {c: f'=IF(ROUND({c}{sga_total_row}-IS!{c}{IS_SGA_ROW},1)=0,"OK","CHK")' for c in ISCF_COLS},
        font=F_PCT, fmt="General")
row += 2

section_row(ws, row, "③ 매출원가 세부 (① − ② 항목별 차감)", ISCF_COLS)
row += 1
cogs_start = row
for item in list(COST_NATURE_BASE) :
    if item in COGS_MATCH:
        m = sga_rows[COGS_MATCH[item]]
        vals = {c: f"={c}{nature_rows[item]}-{c}{m}" for c in ISCF_COLS}
    else:
        vals = {c: f"={c}{nature_rows[item]}" for c in ISCF_COLS}
    put_row(ws, row, f"{IDT}{item}", vals)
    row += 1
put_row(ws, row, f"{IDT}기타",
        {c: f"={c}{nature_rows['기타']}-{c}{sga_rows['여비교통비']}-{c}{sga_rows['세금과공과']}-{c}{sga_rows['기타']}" for c in ISCF_COLS})
row += 1
cogs_end = row - 1
cogs_total_row = row
put_row(ws, row, "매출원가 합계",
        {c: f"=SUM({c}{cogs_start}:{c}{cogs_end})" for c in ISCF_COLS},
        bold_label=True, fill=FILL_KEY)
row += 1
put_row(ws, row, f"{IDT}CHK (IS와 대사)",
        {c: f'=IF(ROUND({c}{cogs_total_row}-IS!{c}{IS_COGS_ROW},1)=0,"OK","CHK")' for c in ISCF_COLS},
        font=F_PCT, fmt="General")
row += 2
ws[f"B{row}"] = "※ ②의 '인건비'는 급여+퇴직급여+복리후생비 합산 — 행 왼쪽 [+]로 세부 펼침. 4Q는 연간 − (1Q+2Q+3Q) 도출."
style_cell(ws[f"B{row}"], F_NOTE)
row += 1
ws[f"B{row}"] = "※ 실제 추출 시 보고서 주석의 항목을 전부 가져오며, 연도별 항목명 차이는 LLM이 매칭·정규화."
style_cell(ws[f"B{row}"], F_NOTE)

# ============================== BS ==============================
ws = wb.create_sheet("BS")
ws.sheet_properties.tabColor = C_RED
BS_HDR = [(BS_COLS[p], p) for p in BS_ORDER]
sheet_header(ws, "BS — 재무상태표 (연결)", "(단위: 억원)", BS_HDR)
BS_ALL_COLS = [BS_COLS[p] for p in BS_ORDER]

def bs_vals(item):
    return {BS_COLS[p]: bs_value(p, item) for p in BS_ORDER}

row = 7
def bs_section(title, items, row):
    top = row
    r = row + 1
    for it in items:
        put_row(ws, r, f"{IDT}{IDT}{it}", bs_vals(it))
        r += 1
    put_row(ws, top, f"{IDT}{title}",
            {c: f"=SUM({c}{top+1}:{c}{r-1})" for c in BS_ALL_COLS}, bold_label=True)
    return top, r

section_row(ws, row, "자산", BS_ALL_COLS); row += 1
cur_a_row, row = bs_section("유동자산", BS_CUR_ASSET, row)
BS_CASH_ROW = cur_a_row + 1
noncur_a_row, row = bs_section("비유동자산", BS_NONCUR_ASSET, row)
asset_total_row = row
put_row(ws, row, "자산총계", {c: f"={c}{cur_a_row}+{c}{noncur_a_row}" for c in BS_ALL_COLS},
        bold_label=True, fill=FILL_KEY)
row += 2

section_row(ws, row, "부채", BS_ALL_COLS); row += 1
cur_l_row, row = bs_section("유동부채", BS_CUR_LIAB, row)
noncur_l_row, row = bs_section("비유동부채", BS_NONCUR_LIAB, row)
liab_total_row = row
put_row(ws, row, "부채총계", {c: f"={c}{cur_l_row}+{c}{noncur_l_row}" for c in BS_ALL_COLS},
        bold_label=True, fill=FILL_KEY)
row += 2

section_row(ws, row, "자본", BS_ALL_COLS); row += 1
eq_start = row
for it in BS_EQUITY_FIXED + ["이익잉여금"]:
    put_row(ws, row, f"{IDT}{IDT}{it}", bs_vals(it))
    row += 1
eq_total_row = row
put_row(ws, row, "자본총계", {c: f"=SUM({c}{eq_start}:{c}{row-1})" for c in BS_ALL_COLS},
        bold_label=True, fill=FILL_KEY)
row += 1
le_total_row = row
put_row(ws, row, "부채와자본총계",
        {c: f"={c}{liab_total_row}+{c}{eq_total_row}" for c in BS_ALL_COLS},
        bold_label=True, fill=FILL_KEY)
row += 1
put_row(ws, row, f"{IDT}CHK (자산=부채+자본)",
        {c: f'=IF(ROUND({c}{asset_total_row}-{c}{le_total_row},1)=0,"OK","CHK")' for c in BS_ALL_COLS},
        font=F_PCT, fmt="General")
row += 2
ws[f"B{row}"] = "※ 시간 순서 배열: FY 컬럼 = 해당 연말(=4Q) 잔액. 실제 추출 시 사업보고서상 계정과목 전부 포함."
style_cell(ws[f"B{row}"], F_NOTE)

# ============================== CF ==============================
ws = wb.create_sheet("CF")
ws.sheet_properties.tabColor = C_RED
sheet_header(ws, "CF — 현금흐름표 (연결)", "(단위: 억원)", ISCF_HDR, gap_cols=(GAP_COL,))

CF2BS = {}
for q in Q_LABELS:
    bs_p = f"FY20{q[-2:]}" if q.startswith("4Q") else q
    CF2BS[ISCF_Q_COLS[q]] = BS_COLS[bs_p]
for y in YEARS:
    CF2BS[ISCF_Y_COLS[y]] = BS_COLS[f"FY{y}"]

def cf_vals(item, row, link_is=False):
    if link_is:
        return {c: f"=IS!{c}{IS_NI_ROW}" for c in ISCF_COLS}
    return base_vals(row, [CF_FY[f"FY{y}"][item] for y in YEARS],
                     {q: CF_Q[q][item] for q in EXTRACTED_Q})

row = 7
ops_row = row
r = row + 1
for it in ["당기순이익", "비현금항목 조정", "영업활동 자산부채의 증감", "이자수취", "이자지급", "법인세납부"]:
    put_row(ws, r, f"{IDT}{it}", cf_vals(it, r, link_is=(it == "당기순이익")))
    r += 1
put_row(ws, ops_row, "영업활동현금흐름",
        {c: f"=SUM({c}{ops_row+1}:{c}{r-1})" for c in ISCF_COLS}, bold_label=True, fill=FILL_KEY)
row = r + 1

inv_row = row
r = row + 1
for it in ["단기금융상품의 순증감", "유형자산의 취득", "유형자산의 처분", "무형자산의 취득"]:
    put_row(ws, r, f"{IDT}{it}", cf_vals(it, r))
    r += 1
put_row(ws, inv_row, "투자활동현금흐름",
        {c: f"=SUM({c}{inv_row+1}:{c}{r-1})" for c in ISCF_COLS}, bold_label=True, fill=FILL_KEY)
row = r + 1

fin_row = row
r = row + 1
for it in ["차입금의 차입", "차입금의 상환", "리스부채의 상환", "배당금의 지급"]:
    put_row(ws, r, f"{IDT}{it}", cf_vals(it, r))
    r += 1
put_row(ws, fin_row, "재무활동현금흐름",
        {c: f"=SUM({c}{fin_row+1}:{c}{r-1})" for c in ISCF_COLS}, bold_label=True, fill=FILL_KEY)
row = r + 1

fx_row = row
put_row(ws, row, "외화환산으로 인한 현금의 변동", cf_vals("외화환산으로 인한 현금의 변동", row))
row += 1
delta_row = row
put_row(ws, row, "현금및현금성자산의 증감",
        {c: f"={c}{ops_row}+{c}{inv_row}+{c}{fin_row}+{c}{fx_row}" for c in ISCF_COLS},
        bold_label=True, fill=FILL_KEY)
row += 1
begin_row = row
end_row = row + 1
begin_vals = cf_vals("기초현금", row)
for q, (ycol, qcols) in Q4_MAP.items():
    begin_vals[ISCF_Q_COLS[q]] = f"={qcols[2]}{end_row}"
put_row(ws, row, "기초현금및현금성자산", begin_vals)
row += 1
put_row(ws, row, "기말현금및현금성자산",
        {c: f"={c}{begin_row}+{c}{delta_row}" for c in ISCF_COLS},
        bold_label=True, fill=FILL_KEY)
row += 1
put_row(ws, row, f"{IDT}CHK (기말현금=BS 현금)",
        {c: f'=IF(ROUND({c}{end_row}-BS!{CF2BS[c]}{BS_CASH_ROW},1)=0,"OK","CHK")' for c in ISCF_COLS},
        font=F_PCT, fmt="General")
row += 2
ws[f"B{row}"] = "※ 분기보고서 현금흐름표는 누적(YTD) 공시 → 직전 분기 누적을 차감해 3개월 값으로 환산. 4Q는 연간 − (1Q+2Q+3Q)."
style_cell(ws[f"B{row}"], F_NOTE)

# ============================== Valuation ==============================
ws = wb.create_sheet("Valuation")
ws.sheet_properties.tabColor = C_RED
sheet_title(ws, "Valuation")
ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 32
for i in range(3, 15):  # C..N
    ws.column_dimensions[get_column_letter(i)].width = 11.5

# ---- 블록 1: 추정 손익계산서 (C..J = FY21~FY25, 26E~28E) ----
V1_COLS = [get_column_letter(3 + i) for i in range(8)]
V1_LABELS = [f"FY{y}" for y in YEARS] + ["FY2026E", "FY2027E", "FY2028E"]
V1_HIST = V1_COLS[:5]
V1_EST = V1_COLS[5:]
IS_YC = {V1_COLS[i]: ISCF_Y_COLS[YEARS[i]] for i in range(5)}  # Valuation열 → IS연간열

row = 4
section_row(ws, row, "추정 손익계산서", V1_COLS)
row += 1
hdr_row = row
ws[f"B{row}"] = "(단위: 억원)"
style_cell(ws[f"B{row}"], F_HDR, fill=FILL_HDR)
for col, lab in zip(V1_COLS, V1_LABELS):
    c = ws[f"{col}{row}"]
    c.value = lab
    style_cell(c, F_HDR, fill=FILL_HDR, align="center")
row += 1

# 매출 추정 블록(블록2) 행 번호 선계산 — 블록1 추정연도 매출액이 참조
REV_BLOCK_TITLE = 26
REV_HDR = 27
REV_TOTAL = 28          # 매출액
V2_COLS = {"FY2021": "C", "FY2022": "D", "FY2023": "E", "FY2024": "F", "FY2025": "G",
           "1Q26": "H", "2Q26E": "I", "3Q26E": "J", "4Q26E": "K",
           "2026E": "L", "2027E": "M", "2028E": "N"}
V2_EST_MAP = {"FY2026E": "L", "FY2027E": "M", "FY2028E": "N"}

VR = {}
VR["매출액"] = row
rev_vals = {c: f"=IS!{IS_YC[c]}{IS_REV_ROW}" for c in V1_HIST}
for i, lab in enumerate(["FY2026E", "FY2027E", "FY2028E"]):
    rev_vals[V1_EST[i]] = f"={V2_EST_MAP[lab]}{REV_TOTAL}"
put_row(ws, row, "매출액", rev_vals, bold_label=True)
row += 1
yoy = {}
for i in range(1, 8):
    yoy[V1_COLS[i]] = f"={V1_COLS[i]}{VR['매출액']}/{V1_COLS[i-1]}{VR['매출액']}-1"
put_row(ws, row, f"{IDT}YoY (%)", yoy, font=F_PCT, fmt=PCT)
row += 1
VR["매출원가"] = row
VR["원가율"] = row + 1
cogs_vals = {c: f"=IS!{IS_YC[c]}{IS_COGS_ROW}" for c in V1_HIST}
for c in V1_EST:
    cogs_vals[c] = f"={c}{VR['매출액']}*{c}{VR['원가율']}"
put_row(ws, row, "매출원가", cogs_vals)
row += 1
ratio_vals = {c: f"={c}{VR['매출원가']}/{c}{VR['매출액']}" for c in V1_HIST}
put_row(ws, row, f"{IDT}매출원가율 가정 (%)", ratio_vals, font=F_PCT, fmt=PCT)
for c, v in zip(V1_EST, [0.565, 0.560, 0.555]):
    ws[f"{c}{row}"] = v
    style_cell(ws[f"{c}{row}"], F_PCT, fmt=PCT, fill=FILL_INPUT)
row += 1
VR["매출총이익"] = row
put_row(ws, row, "매출총이익",
        {c: f"={c}{VR['매출액']}-{c}{VR['매출원가']}" for c in V1_COLS}, bold_label=True, fill=FILL_KEY)
row += 1
put_row(ws, row, f"{IDT}GPM (%)",
        {c: f"={c}{VR['매출총이익']}/{c}{VR['매출액']}" for c in V1_COLS}, font=F_PCT, fmt=PCT)
row += 1
VR["판관비"] = row
VR["판관비율"] = row + 1
sga_vals = {c: f"=IS!{IS_YC[c]}{IS_SGA_ROW}" for c in V1_HIST}
for c in V1_EST:
    sga_vals[c] = f"={c}{VR['매출액']}*{c}{VR['판관비율']}"
put_row(ws, row, "판매비와관리비", sga_vals)
row += 1
ratio_vals = {c: f"={c}{VR['판관비']}/{c}{VR['매출액']}" for c in V1_HIST}
put_row(ws, row, f"{IDT}판관비율 가정 (%)", ratio_vals, font=F_PCT, fmt=PCT)
for c, v in zip(V1_EST, [0.250, 0.248, 0.246]):
    ws[f"{c}{row}"] = v
    style_cell(ws[f"{c}{row}"], F_PCT, fmt=PCT, fill=FILL_INPUT)
row += 1
VR["영업이익"] = row
put_row(ws, row, "영업이익",
        {c: f"={c}{VR['매출총이익']}-{c}{VR['판관비']}" for c in V1_COLS}, bold_label=True, fill=FILL_KEY)
row += 1
put_row(ws, row, f"{IDT}OPM (%)",
        {c: f"={c}{VR['영업이익']}/{c}{VR['매출액']}" for c in V1_COLS}, font=F_PCT, fmt=PCT)
row += 1
VR["영업외"] = row
nonop_vals = {c: f"=IS!{IS_YC[c]}{IS_PBT_ROW}-IS!{IS_YC[c]}{IS_OP_ROW}" for c in V1_HIST}
for c, v in zip(V1_EST, [3.0, 3.0, 3.0]):
    nonop_vals[c] = v
put_row(ws, row, "영업외손익 (추정연도 입력)", nonop_vals)
for c in V1_EST:
    ws[f"{c}{row}"].fill = FILL_INPUT
row += 1
VR["법차전"] = row
put_row(ws, row, "법인세비용차감전순이익",
        {c: f"={c}{VR['영업이익']}+{c}{VR['영업외']}" for c in V1_COLS}, bold_label=True, fill=FILL_KEY)
row += 1
VR["법인세"] = row
VR["세율"] = row + 1
tax_vals = {c: f"=IS!{IS_YC[c]}{IS_TAX_ROW}" for c in V1_HIST}
for c in V1_EST:
    tax_vals[c] = f"={c}{VR['법차전']}*{c}{VR['세율']}"
put_row(ws, row, "법인세비용", tax_vals)
row += 1
ratio_vals = {c: f"={c}{VR['법인세']}/{c}{VR['법차전']}" for c in V1_HIST}
put_row(ws, row, f"{IDT}실효세율 가정 (%)", ratio_vals, font=F_PCT, fmt=PCT)
for c in V1_EST:
    ws[f"{c}{row}"] = 0.21
    style_cell(ws[f"{c}{row}"], F_PCT, fmt=PCT, fill=FILL_INPUT)
row += 1
VR["당기순이익"] = row
put_row(ws, row, "당기순이익",
        {c: f"={c}{VR['법차전']}-{c}{VR['법인세']}" for c in V1_COLS}, bold_label=True, fill=FILL_KEY)
row += 1
put_row(ws, row, f"{IDT}NPM (%)",
        {c: f"={c}{VR['당기순이익']}/{c}{VR['매출액']}" for c in V1_COLS}, font=F_PCT, fmt=PCT)

# ---- 블록 2: 매출 추정 ----
V2_ALL = list(V2_COLS.values())
assert REV_BLOCK_TITLE > row + 1, f"블록1이 블록2 위치({REV_BLOCK_TITLE})를 침범: row={row}"
section_row(ws, REV_BLOCK_TITLE, "매출 추정 (사업부별 — 직접 입력)", V2_ALL)
ws[f"B{REV_HDR}"] = "(단위: 억원)"
style_cell(ws[f"B{REV_HDR}"], F_HDR, fill=FILL_HDR)
for lab, col in V2_COLS.items():
    c = ws[f"{col}{REV_HDR}"]
    c.value = lab
    style_cell(c, F_HDR, fill=FILL_HDR, align="center")

# 사업부 샘플 분해 (4번째 사업부 = plug) — 실제로는 회사 사업부 구조에 맞춰 입력
def seg_split(total):
    s1, s2, s3 = r1(total * 0.4), r1(total * 0.3), r1(total * 0.2)
    return [s1, s2, s3, r1(total - s1 - s2 - s3)]

SEG_HIST = {lab: seg_split(IS_Y["매출액"][i]) for i, lab in
            enumerate(["FY2021", "FY2022", "FY2023", "FY2024", "FY2025"])}
SEG_HIST["1Q26"] = seg_split(IS_Q["매출액"]["1Q26"])
SEG_EST = {  # 추정 분기·연간 (노란색 입력 예시)
    "2Q26E": [148.0, 111.0, 74.0, 39.0],
    "3Q26E": [151.0, 113.0, 75.5, 40.0],
    "4Q26E": [154.0, 115.5, 77.0, 40.5],
    "2027E": [660.0, 495.0, 330.0, 175.0],
    "2028E": [726.0, 544.5, 363.0, 192.5],
}

seg_row_start = REV_HDR + 1 + 1  # 매출액 행 다음
# 매출액 행
rev_total_vals = {}
for lab, col in V2_COLS.items():
    if lab.startswith("FY"):
        rev_total_vals[col] = f"=IS!{ISCF_Y_COLS[int(lab[2:])]}{IS_REV_ROW}"
    elif lab == "1Q26":
        rev_total_vals[col] = f"=IS!{ISCF_Q_COLS['1Q26']}{IS_REV_ROW}"
    elif lab == "2026E":
        rev_total_vals[col] = f"=SUM(H{REV_TOTAL}:K{REV_TOTAL})"
    else:  # 2Q~4Q26E, 2027E, 2028E = 사업부 합
        rev_total_vals[col] = f"=SUM({col}{seg_row_start}:{col}{seg_row_start+2*4-1})*0.5"  # 임시, 아래에서 교체
put_row(ws, REV_TOTAL, "매출액", rev_total_vals, bold_label=True, fill=FILL_KEY)

r = seg_row_start
seg_rows = []
for si in range(4):
    seg_rows.append(r)
    vals = {}
    for lab, col in V2_COLS.items():
        if lab in SEG_HIST:
            vals[col] = SEG_HIST[lab][si]
        elif lab in SEG_EST:
            vals[col] = SEG_EST[lab][si]
        elif lab == "2026E":
            vals[col] = f"=SUM(H{r}:K{r})"
    put_row(ws, r, f"{IDT}{si+1} 사업부", vals)
    for lab in list(SEG_HIST) + list(SEG_EST):
        ws[f"{V2_COLS[lab]}{r}"].fill = FILL_INPUT
    r += 1
    pct = {}
    for lab, col in V2_COLS.items():  # %비중
        pct[col] = f"={col}{r-1}/{col}{REV_TOTAL}"
    put_row(ws, r, f"{IDT}{IDT}비중 (%)", pct, font=F_PCT, fmt=PCT)
    r += 1
# 매출액 행의 추정 컬럼을 사업부 합으로 교체
for lab in ["2Q26E", "3Q26E", "4Q26E", "2027E", "2028E"]:
    col = V2_COLS[lab]
    parts = "+".join(f"{col}{sr}" for sr in seg_rows)
    ws[f"{col}{REV_TOTAL}"] = f"={parts}"
chk_row = r
put_row(ws, chk_row, f"{IDT}CHK (실적 = 사업부 합)",
        {V2_COLS[lab]: f'=IF(ROUND({V2_COLS[lab]}{REV_TOTAL}-({"+".join(f"{V2_COLS[lab]}{sr}" for sr in seg_rows)}),1)=0,"OK","CHK")'
         for lab in list(SEG_HIST)},
        font=F_PCT, fmt="General")

# ---- 블록 3: 주주환원 ----
row = chk_row + 3
V3_COLS = ["C", "D", "E", "F"]
V3_LABELS = ["FY2025", "FY2026E", "FY2027E", "FY2028E"]
V3_NI_SRC = {"C": V1_COLS[4], "D": V1_COLS[5], "E": V1_COLS[6], "F": V1_COLS[7]}
section_row(ws, row, "주주환원 — 배당 / 자사주 소각", V3_COLS)
row += 1
ws[f"B{row}"] = "구분"
style_cell(ws[f"B{row}"], F_HDR, fill=FILL_HDR)
for col, lab in zip(V3_COLS, V3_LABELS):
    c = ws[f"{col}{row}"]
    c.value = lab
    style_cell(c, F_HDR, fill=FILL_HDR, align="center")
row += 1
NI_R = row
put_row(ws, row, "당기순이익 (억원)",
        {c: f"={V3_NI_SRC[c]}{VR['당기순이익']}" for c in V3_COLS})
row += 1
PAYOUT_R = row
put_row(ws, row, "배당성향 가정 (%)", {c: 0.20 for c in V3_COLS}, fmt=PCT, fill=FILL_INPUT)
row += 1
BUYBACK_R = row
put_row(ws, row, "자사주 소각 주식수 (주)", {"C": 0, "D": 200000, "E": 200000, "F": 0},
        fmt=INT, fill=FILL_INPUT)
row += 1
SHARES_R = row
put_row(ws, row, "기말 유통주식수 (주)",
        {"C": SHARES, "D": f"=C{SHARES_R}-D{BUYBACK_R}", "E": f"=D{SHARES_R}-E{BUYBACK_R}",
         "F": f"=E{SHARES_R}-F{BUYBACK_R}"}, fmt=INT)
style_cell(ws[f"C{SHARES_R}"], F_BASE, fmt=INT, fill=FILL_INPUT)
row += 1
EPS_R = row
put_row(ws, row, "EPS (원)", {c: f"={c}{NI_R}*10^8/{c}{SHARES_R}" for c in V3_COLS},
        fmt=INT, bold_label=True, fill=FILL_KEY)
row += 1
DPS_R = row
put_row(ws, row, "DPS (원)", {c: f"={c}{NI_R}*{c}{PAYOUT_R}*10^8/{c}{SHARES_R}" for c in V3_COLS}, fmt=INT)
row += 1
put_row(ws, row, "총배당액 (억원)", {c: f"={c}{NI_R}*{c}{PAYOUT_R}" for c in V3_COLS})
row += 1
YIELD_R = row
row += 1
ws[f"B{row}"] = "• 배당정책 공시: (적용기간 / 배당 기준 / 최저배당액 입력)   • 자사주 소각 계획: (공시 내용 입력)"
style_cell(ws[f"B{row}"], F_NOTE)
row += 2

# ---- 블록 4: 밸류에이션 밴드 ----
section_row(ws, row, "밸류에이션 밴드 (PER × EPS)", V3_COLS)
row += 1
ws[f"B{row}"] = "PER (배)"
style_cell(ws[f"B{row}"], F_HDR, fill=FILL_HDR, align="center")
for col, lab in zip(V3_COLS, V3_LABELS):
    c = ws[f"{col}{row}"]
    c.value = lab
    style_cell(c, F_HDR, fill=FILL_HDR, align="center")
row += 1
PER_START = row
for i, per in enumerate([6, 7, 8, 9, 10, 11]):
    ws[f"B{row}"] = per
    style_cell(ws[f"B{row}"], F_BASE, fmt="0", fill=FILL_INPUT, align="center")
    for col in V3_COLS:
        ws[f"{col}{row}"] = f"={col}${EPS_R}*$B{row}"
        style_cell(ws[f"{col}{row}"], F_BASE, fmt=INT)
    row += 1
PER8_ROW = PER_START + 2
row += 1
CUR_PRICE = f"C{row}"
ws[f"B{row}"] = "현재주가 (원)"
style_cell(ws[f"B{row}"], F_BOLD)
ws[CUR_PRICE] = 8500
style_cell(ws[CUR_PRICE], F_BASE, fmt=INT, fill=FILL_INPUT)
row += 1
put_row(ws, row, "Upside vs PER 8x (%)",
        {c: f"={c}{PER8_ROW}/${CUR_PRICE}-1" for c in V3_COLS}, fmt=PCT)
# 배당수익률 (현재주가 확정 후 기입)
put_row(ws, YIELD_R, "배당수익률 (%)", {c: f"={c}{DPS_R}/${CUR_PRICE}" for c in V3_COLS}, fmt=PCT)
row += 2
ws[f"B{row}"] = "※ 노란색 셀 직접 입력. 과거 실적은 IS 시트 자동 연결, 추정 매출은 '매출 추정' 사업부 합계와 연동."
style_cell(ws[f"B{row}"], F_NOTE)

# ============================== Check ==============================
ws = wb.create_sheet("Check")
ws.sheet_properties.tabColor = C_DARK
sheet_title(ws, "Check — 보고서 간 값이 달라진 항목")
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
    style_cell(c, F_HDR, fill=FILL_HDR, align="center")

diff_rows = [
    ["IS", "매출원가", "FY2023", 603.1, "FY2023 사업보고서", 601.3, "FY2024 사업보고서 (재분류)"],
    ["BS", "매출채권 및 기타채권", "FY2022", 110.8, "FY2022 사업보고서", 112.4, "FY2024 사업보고서 (재작성)"],
    ["CF", "이자지급", "FY2023", -9.8, "FY2023 사업보고서", -10.1, "FY2024 사업보고서 (표시방법 변경)"],
]
for j, dr in enumerate(diff_rows):
    r = 7 + j
    for i in range(9):
        col = get_column_letter(2 + i)
        cell = ws[f"{col}{r}"]
        if i == 7:
            cell.value = f"=G{r}-E{r}"
        elif i == 8:
            cell.value = f"=G{r}/E{r}-1"
        else:
            cell.value = dr[i]
        fmt = NUM if i in (3, 5, 7) else (PCT if i == 8 else "General")
        style_cell(cell, F_BASE, fmt=fmt, border=True)

ws["B12"] = "※ 예시 데이터입니다. 실제 추출기는 발견된 모든 차이를 여기에 나열합니다."
style_cell(ws["B12"], F_NOTE)

# 저장
out = sys.argv[1] if len(sys.argv) > 1 else "재무제표_추출_예시.xlsx"
wb.save(out)
print(f"saved: {out}")
