# -*- coding: utf-8 -*-
"""
재무제표 추출기 — 결과물 예시 엑셀 생성 스크립트 (v2)

사업보고서/분기보고서를 넣으면 만들어질 최종 엑셀의 "모양"을 미리 보여주기 위한
샘플 데이터 기반 생성기.

시트 구성: IS / COST / BS / CF / Valuation (+ Check: 보고서 간 값이 달라진 항목 알림)

컬럼 레이아웃
  IS·CF : 분기 13개(1Q23~1Q26, 4Q는 연간-3개분기 도출) → 한 칸 공백 → 연간 FY2021~FY2025
  BS    : 시간 순서 FY21, FY22, 1Q23, 2Q23, 3Q23, FY23, ... , FY25, 1Q26
  COST  : 연간 FY2021~FY2025 (①비용의 성격별 분류 / ②판관비 세부(인건비 그룹) / ③매출원가 세부=①-②)
  Valuation : 전년도+3개년 (FY2025 실적 + FY2026E~FY2028E), 배당성향·자사주 소각 반영 PER 밴드

가상 시나리오
  입력 보고서: 사업보고서 FY2023 / FY2024 / FY2025,
              분기보고서 2023~2025 (1Q, 반기, 3Q) + 2026 1Q
  → FY2023 사업보고서의 전전기(제n-2기) 컬럼 덕분에 2021년 숫자까지 소급 추출
  → 같은 기간 숫자가 보고서마다 다르면 "최신 보고서" 값을 채택하고 Check 시트에 기록
"""

import sys
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

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

NUM = "#,##0.0;[Red](#,##0.0)"
PCT = "0.0%;[Red](0.0%)"
INT = "#,##0"

IDT = "　"  # 전각 공백 들여쓰기 (기존 양식과 동일)

# ──────────────────────────────────────────────────────────────
# 기간 / 컬럼 레이아웃
# ──────────────────────────────────────────────────────────────
YEARS = [2021, 2022, 2023, 2024, 2025]
FY_LABELS = [f"FY{y}" for y in YEARS]
Q_LABELS = ["1Q23", "2Q23", "3Q23", "4Q23", "1Q24", "2Q24", "3Q24", "4Q24",
            "1Q25", "2Q25", "3Q25", "4Q25", "1Q26"]          # 13개 분기
EXTRACTED_Q = [q for q in Q_LABELS if not q.startswith("4Q")]  # 4Q는 도출

# IS·CF: 분기 C~O(13개) / P 공백 / 연간 Q~U
ISCF_Q_COLS = {q: get_column_letter(3 + i) for i, q in enumerate(Q_LABELS)}
ISCF_Y_COLS = {y: get_column_letter(17 + i) for i, y in enumerate(YEARS)}   # Q..U
ISCF_COLS = list(ISCF_Q_COLS.values()) + list(ISCF_Y_COLS.values())
# 4Q 도출: (4Q라벨) → (연간컬럼, [1~3Q컬럼])
Q4_MAP = {
    "4Q23": (ISCF_Y_COLS[2023], [ISCF_Q_COLS["1Q23"], ISCF_Q_COLS["2Q23"], ISCF_Q_COLS["3Q23"]]),
    "4Q24": (ISCF_Y_COLS[2024], [ISCF_Q_COLS["1Q24"], ISCF_Q_COLS["2Q24"], ISCF_Q_COLS["3Q24"]]),
    "4Q25": (ISCF_Y_COLS[2025], [ISCF_Q_COLS["1Q25"], ISCF_Q_COLS["2Q25"], ISCF_Q_COLS["3Q25"]]),
}

# BS: 시간 순서 (FY 컬럼 = 해당 연말 잔액)
BS_ORDER = ["FY2021", "FY2022",
            "1Q23", "2Q23", "3Q23", "FY2023",
            "1Q24", "2Q24", "3Q24", "FY2024",
            "1Q25", "2Q25", "3Q25", "FY2025",
            "1Q26"]
BS_COLS = {p: get_column_letter(3 + i) for i, p in enumerate(BS_ORDER)}     # C..Q

# COST: 연간만
COST_Y_COLS = {y: get_column_letter(3 + i) for i, y in enumerate(YEARS)}    # C..G

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

# ──────────────────────────────────────────────────────────────
# 샘플 데이터 (단위: 억원) — 가상의 회사
# ──────────────────────────────────────────────────────────────
IS_Y = {  # 연간 FY2021~FY2025
    "매출액":     [812.4, 934.1, 1023.5, 1187.9, 1352.6],
    "매출원가":   [487.2, 549.8, 601.3, 688.5, 769.4],
    "판매비와관리비": [214.6, 241.3, 265.8, 301.2, 338.7],
    "기타수익":   [3.2, 4.1, 2.8, 5.6, 4.9],
    "기타비용":   [5.1, 3.8, 6.2, 4.3, 7.1],
    "금융수익":   [6.4, 8.9, 12.3, 15.1, 13.8],
    "금융비용":   [9.8, 11.2, 10.6, 9.4, 8.2],
    "법인세비용": [21.3, 28.9, 31.5, 42.8, 51.6],
}
# 추출 분기(1~3Q + 1Q26). 순서: 1Q23,2Q23,3Q23, 1Q24,2Q24,3Q24, 1Q25,2Q25,3Q25, 1Q26
IS_Q_RAW = {
    "매출액":     [238.5, 252.1, 259.8, 275.3, 291.8, 298.4, 312.7, 334.2, 341.5, 348.2],
    "매출원가":   [141.2, 147.9, 152.4, 162.1, 169.5, 172.8, 180.3, 190.1, 194.6, 199.5],
    "판매비와관리비": [62.4, 65.3, 66.9, 71.2, 74.8, 76.1, 80.9, 83.6, 85.2, 89.4],
    "기타수익":   [0.6, 0.8, 0.7, 1.2, 1.5, 1.1, 1.0, 1.4, 1.2, 1.3],
    "기타비용":   [1.4, 1.6, 1.5, 0.9, 1.1, 1.3, 1.8, 1.6, 2.1, 1.5],
    "금융수익":   [2.9, 3.1, 3.0, 3.6, 3.9, 3.8, 3.4, 3.5, 3.3, 3.6],
    "금융비용":   [2.8, 2.7, 2.6, 2.5, 2.4, 2.3, 2.1, 2.0, 2.1, 1.9],
    "법인세비용": [7.4, 8.1, 7.8, 10.2, 11.5, 10.8, 12.4, 13.7, 12.9, 14.2],
}
IS_Q = {item: dict(zip(EXTRACTED_Q, vals)) for item, vals in IS_Q_RAW.items()}

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
SGA_LABOR = {  # 인건비로 묶이는 항목 (그룹 접기)
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

# ③ 매출원가 세부 = ① − ② 매칭 (①항목명 → ②항목명; 종업원급여는 ②인건비 소계와 매칭)
COGS_MATCH = {
    "종업원급여": "인건비",
    "감가상각비": "감가상각비",
    "무형자산상각비": "무형자산상각비",
    "지급수수료": "지급수수료",
    "임차료 및 관리비": "임차료 및 관리비",
    "광고선전비": "광고선전비",
    "운반비": "운반비",
    "소모품비": "소모품비",
    # 기타: ①기타 − (②여비교통비+②세금과공과+②기타) → 코드에서 별도 처리
}

# BS 연말 잔액 (FY2021~FY2025) — 실제 추출 시 사업보고서상 계정 전부를 가져옴
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

CASH_2020 = 58.3  # FY2021 CF 기초현금용

# ──────────────────────────────────────────────────────────────
# 파생 데이터 계산
# ──────────────────────────────────────────────────────────────
def r1(x):
    return round(x + 1e-9, 1)

def bs_totals(d):
    a = sum(d[k] for k in BS_ASSET_ITEMS)
    l = sum(d[k] for k in BS_CUR_LIAB + BS_NONCUR_LIAB)
    eq_fixed = sum(d[k] for k in BS_EQUITY_FIXED)
    return r1(a), r1(l), r1(a - l - eq_fixed)   # 이익잉여금 = plug → 대차 자동 일치

BS_FY = {}   # "FY2021" → dict
for i, y in enumerate(YEARS):
    d = {k: BS_Y[k][i] for k in BS_Y}
    _, _, re = bs_totals(d)
    d["이익잉여금"] = re
    BS_FY[f"FY{y}"] = d

# 분기 BS: 연말 사이 선형보간 + 소폭 변동. 1Q26은 FY25 추세 연장.
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

def is_ni(vals):  # 당기순이익 검산 (매출액~법인세 dict/list)
    gp = vals["매출액"] - vals["매출원가"]
    op = gp - vals["판매비와관리비"]
    pbt = op + vals["기타수익"] - vals["기타비용"] + vals["금융수익"] - vals["금융비용"]
    return r1(pbt - vals["법인세비용"])

CF_FY = {}
cf_params_y = {  # dep, int_in, int_out, tax, stf, capex, disp, intg, borrow, repay, lease, div, fx
    2021: (62.4, 5.8, -9.2, -19.5, -6.2, -38.5, 2.1, -8.4, 20.0, -35.0, -4.5, -12.0, 0.8),
    2022: (70.2, 8.1, -10.6, -26.4, -7.3, -42.8, 1.5, -9.2, 15.0, -30.0, -5.0, -14.0, -1.2),
    2023: (75.6, 11.4, -10.1, -29.8, -7.8, -48.2, 3.2, -10.5, 10.0, -25.0, -5.5, -16.0, 0.5),
    2024: (84.3, 14.2, -8.9, -38.6, -11.4, -55.4, 1.8, -11.8, 10.0, -25.0, -6.1, -18.0, -0.7),
    2025: (94.8, 12.9, -7.8, -47.2, -13.7, -63.2, 2.4, -12.6, 5.0, -20.0, -6.7, -21.0, 1.1),
}
for i, y in enumerate(YEARS):
    begin = CASH_2020 if i == 0 else BS_FY[f"FY{YEARS[i-1]}"]["현금및현금성자산"]
    end = BS_FY[f"FY{y}"]["현금및현금성자산"]
    ni = is_ni({k: IS_Y[k][i] for k in IS_Y})
    CF_FY[f"FY{y}"] = build_cf(ni, *cf_params_y[y], cash_begin=begin, cash_end=end)

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
Q_BEGIN = {  # 각 분기의 기초현금 출처
    "1Q23": "FY2022", "2Q23": "1Q23", "3Q23": "2Q23",
    "1Q24": "FY2023", "2Q24": "1Q24", "3Q24": "2Q24",
    "1Q25": "FY2024", "2Q25": "1Q25", "3Q25": "2Q25",
    "1Q26": "FY2025",
}
CF_Q = {}
for q in EXTRACTED_Q:
    begin = bs_value(Q_BEGIN[q], "현금및현금성자산")
    end = bs_value(q, "현금및현금성자산")
    ni = is_ni({k: IS_Q[k][q] for k in IS_Q})
    CF_Q[q] = build_cf(ni, *cf_params_q[q], cash_begin=begin, cash_end=end)

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

def sheet_header(ws, title, unit, col_label_pairs, gap_cols=()):
    ws.sheet_view.showGridLines = False
    ws["B2"] = title
    ws["B2"].font = F_TITLE
    ws["B4"] = unit
    style_cell(ws["B4"], F_NOTE)
    ws["B4"].border = BORDER
    for col, label in col_label_pairs:
        c = ws[f"{col}4"]
        c.value = label
        style_cell(c, F_HDR, fill=FILL_HDR, border=True, align="center")
        s = ws[f"{col}5"]
        s.value = SOURCES.get(str(label), "")
        style_cell(s, F_SRC)
        s.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.column_dimensions[col].width = 12.5
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 34
    for g in gap_cols:
        ws.column_dimensions[g].width = 2.5
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

ISCF_HDR = [(ISCF_Q_COLS[q], q) for q in Q_LABELS] + [(ISCF_Y_COLS[y], f"FY{y}") for y in YEARS]
GAP_COL = get_column_letter(16)  # P

# ============================== IS ==============================
ws = wb.create_sheet("IS")
ws.sheet_properties.tabColor = "1F4E79"
sheet_header(ws, "IS — 손익계산서 (연결)", "(단위: 억원)", ISCF_HDR, gap_cols=(GAP_COL,))

def is_base_row(row, item):
    vals = {}
    for q in EXTRACTED_Q:
        vals[ISCF_Q_COLS[q]] = IS_Q[item][q]
    for q, (ycol, qcols) in Q4_MAP.items():
        vals[ISCF_Q_COLS[q]] = f"={ycol}{row}-{qcols[0]}{row}-{qcols[1]}{row}-{qcols[2]}{row}"
    for i, y in enumerate(YEARS):
        vals[ISCF_Y_COLS[y]] = IS_Y[item][i]
    return vals

R = {}
row = 7
R["매출액"] = row; put_row(ws, row, "매출액", is_base_row(row, "매출액"), bold_label=True)
row += 1
yoy = {}
for i in range(4, len(Q_LABELS)):  # 1Q24부터 전년동기 대비
    cur, prev = ISCF_Q_COLS[Q_LABELS[i]], ISCF_Q_COLS[Q_LABELS[i - 4]]
    yoy[cur] = f"={cur}{R['매출액']}/{prev}{R['매출액']}-1"
for i in range(1, 5):
    a, b = ISCF_Y_COLS[YEARS[i]], ISCF_Y_COLS[YEARS[i - 1]]
    yoy[a] = f"={a}{R['매출액']}/{b}{R['매출액']}-1"
put_row(ws, row, f"{IDT}YoY (%)", yoy, font=F_PCT, fmt=PCT)
row += 1
R["매출원가"] = row; put_row(ws, row, "매출원가", is_base_row(row, "매출원가"))
row += 1
R["매출총이익"] = row
put_row(ws, row, "매출총이익",
        {c: f"={c}{R['매출액']}-{c}{R['매출원가']}" for c in ISCF_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
row += 1
put_row(ws, row, f"{IDT}GPM (%)",
        {c: f"={c}{R['매출총이익']}/{c}{R['매출액']}" for c in ISCF_COLS}, font=F_PCT, fmt=PCT)
row += 1
R["판관비"] = row; put_row(ws, row, "판매비와관리비", is_base_row(row, "판매비와관리비"))
row += 1
R["영업이익"] = row
put_row(ws, row, "영업이익",
        {c: f"={c}{R['매출총이익']}-{c}{R['판관비']}" for c in ISCF_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
row += 1
put_row(ws, row, f"{IDT}OPM (%)",
        {c: f"={c}{R['영업이익']}/{c}{R['매출액']}" for c in ISCF_COLS}, font=F_PCT, fmt=PCT)
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
        {c: f"={c}{R['영업이익']}+{c}{R['기타수익']}-{c}{R['기타비용']}+{c}{R['금융수익']}-{c}{R['금융비용']}" for c in ISCF_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
row += 1
R["법인세"] = row; put_row(ws, row, "법인세비용", is_base_row(row, "법인세비용"))
row += 1
R["당기순이익"] = row
put_row(ws, row, "당기순이익",
        {c: f"={c}{R['법차전']}-{c}{R['법인세']}" for c in ISCF_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
row += 1
put_row(ws, row, f"{IDT}NPM (%)",
        {c: f"={c}{R['당기순이익']}/{c}{R['매출액']}" for c in ISCF_COLS}, font=F_PCT, fmt=PCT)
IS_NI_ROW = R["당기순이익"]
IS_COGS_ROW = R["매출원가"]
IS_SGA_ROW = R["판관비"]
note_row = row + 2
ws[f"B{note_row}"] = "※ 4Q 컬럼은 연간(FY) 실적 − (1Q+2Q+3Q)로 자동 도출. 음영(파랑) 행은 수식으로 계산되는 소계."
style_cell(ws[f"B{note_row}"], F_NOTE)

# ============================== COST ==============================
ws = wb.create_sheet("COST")
ws.sheet_properties.tabColor = "C55A11"
ws.sheet_properties.outlinePr.summaryBelow = False   # 그룹 소계가 위, 세부가 아래
COST_HDR = [(COST_Y_COLS[y], f"FY{y}") for y in YEARS]
sheet_header(ws, "COST — 비용 구조 (주석)", "(단위: 억원)", COST_HDR)
Y_ONLY = list(COST_Y_COLS.values())
cost_col = {y: COST_Y_COLS[y] for y in YEARS}
IS_YCOL = {y: ISCF_Y_COLS[y] for y in YEARS}

row = 7
ws[f"B{row}"] = "① 비용의 성격별 분류 (매출원가 + 판관비)"
style_cell(ws[f"B{row}"], F_BOLD)
row += 1
nature_rows = {}          # 항목명 → 행번호
nature_start = row
totals = [r1(IS_Y["매출원가"][i] + IS_Y["판매비와관리비"][i]) for i in range(5)]
others = list(totals)
for item, vals in COST_NATURE_BASE.items():
    nature_rows[item] = row
    put_row(ws, row, f"{IDT}{item}", {cost_col[YEARS[i]]: vals[i] for i in range(5)})
    others = [r1(others[i] - vals[i]) for i in range(5)]
    row += 1
nature_rows["기타"] = row
put_row(ws, row, f"{IDT}기타", {cost_col[YEARS[i]]: others[i] for i in range(5)})
row += 1
nature_end = row - 1
nature_total_row = row
put_row(ws, row, "합계 (매출원가+판관비)",
        {c: f"=SUM({c}{nature_start}:{c}{nature_end})" for c in Y_ONLY},
        bold_label=True, fill=FILL_SUBTOTAL)
row += 1
put_row(ws, row, f"{IDT}CHK (IS와 대사)",
        {cost_col[y]: f'=IF(ROUND({cost_col[y]}{nature_total_row}-(IS!{IS_YCOL[y]}{IS_COGS_ROW}+IS!{IS_YCOL[y]}{IS_SGA_ROW}),1)=0,"OK","CHK")' for y in YEARS},
        font=F_PCT, fmt="General")
row += 3

ws[f"B{row}"] = "② 판매비와관리비 세부"
style_cell(ws[f"B{row}"], F_BOLD)
row += 1
sga_start = row
sga_rows = {}
# 인건비: 소계 행 + 세부(급여/퇴직급여/복리후생비)는 그룹으로 접어둠 (+ 버튼으로 펼침)
labor_total_row = row
labor_detail_start = row + 1
labor_detail_end = row + len(SGA_LABOR)
put_row(ws, labor_total_row, f"{IDT}인건비 (급여+퇴직급여+복리후생비)",
        {c: f"=SUM({c}{labor_detail_start}:{c}{labor_detail_end})" for c in Y_ONLY},
        bold_label=True)
sga_rows["인건비"] = labor_total_row
row += 1
for item, vals in SGA_LABOR.items():
    put_row(ws, row, f"{IDT}{IDT}{item}", {cost_col[YEARS[i]]: vals[i] for i in range(5)}, font=F_PCT)
    ws.row_dimensions[row].outline_level = 1
    ws.row_dimensions[row].hidden = True
    row += 1
others = [IS_Y["판매비와관리비"][i] for i in range(5)]
labor_sum = [r1(sum(SGA_LABOR[k][i] for k in SGA_LABOR)) for i in range(5)]
others = [r1(others[i] - labor_sum[i]) for i in range(5)]
for item, vals in SGA_OTHER.items():
    sga_rows[item] = row
    put_row(ws, row, f"{IDT}{item}", {cost_col[YEARS[i]]: vals[i] for i in range(5)})
    others = [r1(others[i] - vals[i]) for i in range(5)]
    row += 1
sga_rows["기타"] = row
put_row(ws, row, f"{IDT}기타", {cost_col[YEARS[i]]: others[i] for i in range(5)})
row += 1
sga_end = row - 1
sga_total_row = row
put_row(ws, row, "판관비 합계",
        {c: f"={c}{labor_total_row}+SUM({c}{labor_detail_end+1}:{c}{sga_end})" for c in Y_ONLY},
        bold_label=True, fill=FILL_SUBTOTAL)
row += 1
put_row(ws, row, f"{IDT}CHK (IS와 대사)",
        {cost_col[y]: f'=IF(ROUND({cost_col[y]}{sga_total_row}-IS!{IS_YCOL[y]}{IS_SGA_ROW},1)=0,"OK","CHK")' for y in YEARS},
        font=F_PCT, fmt="General")
row += 3

ws[f"B{row}"] = "③ 매출원가 세부 (① − ② 항목별 차감)"
style_cell(ws[f"B{row}"], F_BOLD)
row += 1
cogs_start = row
for item in COST_NATURE_BASE:
    if item in COGS_MATCH:
        m = sga_rows[COGS_MATCH[item]]
        vals = {c: f"={c}{nature_rows[item]}-{c}{m}" for c in Y_ONLY}
    else:  # 판관비에 대응 항목 없음 → 전액 매출원가
        vals = {c: f"={c}{nature_rows[item]}" for c in Y_ONLY}
    put_row(ws, row, f"{IDT}{item}", vals)
    row += 1
# 기타 = ①기타 − (②여비교통비 + ②세금과공과 + ②기타)
put_row(ws, row, f"{IDT}기타",
        {c: f"={c}{nature_rows['기타']}-{c}{sga_rows['여비교통비']}-{c}{sga_rows['세금과공과']}-{c}{sga_rows['기타']}" for c in Y_ONLY})
row += 1
cogs_end = row - 1
cogs_total_row = row
put_row(ws, row, "매출원가 합계",
        {c: f"=SUM({c}{cogs_start}:{c}{cogs_end})" for c in Y_ONLY},
        bold_label=True, fill=FILL_SUBTOTAL)
row += 1
put_row(ws, row, f"{IDT}CHK (IS와 대사)",
        {cost_col[y]: f'=IF(ROUND({cost_col[y]}{cogs_total_row}-IS!{IS_YCOL[y]}{IS_COGS_ROW},1)=0,"OK","CHK")' for y in YEARS},
        font=F_PCT, fmt="General")
row += 2
ws[f"B{row}"] = "※ ②의 '인건비'는 급여+퇴직급여+복리후생비 합산. 행 왼쪽 [+] 버튼으로 세부 항목 펼침."
style_cell(ws[f"B{row}"], F_NOTE)
row += 1
ws[f"B{row}"] = "※ 실제 추출 시 보고서 주석에 있는 항목을 전부 가져오며, 연도별 항목명 차이는 LLM이 매칭·정규화."
style_cell(ws[f"B{row}"], F_NOTE)

# ============================== BS ==============================
ws = wb.create_sheet("BS")
ws.sheet_properties.tabColor = "2E7D32"
BS_HDR = [(BS_COLS[p], p) for p in BS_ORDER]
sheet_header(ws, "BS — 재무상태표 (연결)", "(단위: 억원)", BS_HDR)
BS_ALL_COLS = [BS_COLS[p] for p in BS_ORDER]

def bs_vals(item):
    return {BS_COLS[p]: bs_value(p, item) for p in BS_ORDER}

row = 7
def bs_section(title, items, row):
    section_row = row
    r = row + 1
    for it in items:
        put_row(ws, r, f"{IDT}{IDT}{it}", bs_vals(it))
        r += 1
    put_row(ws, section_row, f"{IDT}{title}",
            {c: f"=SUM({c}{section_row+1}:{c}{r-1})" for c in BS_ALL_COLS},
            bold_label=True, fill=FILL_SUBTOTAL)
    return section_row, r

ws[f"B{row}"] = "자산"; style_cell(ws[f"B{row}"], F_BOLD); row += 1
cur_a_row, row = bs_section("유동자산", BS_CUR_ASSET, row)
BS_CASH_ROW = cur_a_row + 1
noncur_a_row, row = bs_section("비유동자산", BS_NONCUR_ASSET, row)
asset_total_row = row
put_row(ws, row, "자산총계", {c: f"={c}{cur_a_row}+{c}{noncur_a_row}" for c in BS_ALL_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
row += 2

ws[f"B{row}"] = "부채"; style_cell(ws[f"B{row}"], F_BOLD); row += 1
cur_l_row, row = bs_section("유동부채", BS_CUR_LIAB, row)
noncur_l_row, row = bs_section("비유동부채", BS_NONCUR_LIAB, row)
liab_total_row = row
put_row(ws, row, "부채총계", {c: f"={c}{cur_l_row}+{c}{noncur_l_row}" for c in BS_ALL_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
row += 2

ws[f"B{row}"] = "자본"; style_cell(ws[f"B{row}"], F_BOLD); row += 1
eq_start = row
for it in BS_EQUITY_FIXED + ["이익잉여금"]:
    put_row(ws, row, f"{IDT}{IDT}{it}", bs_vals(it))
    row += 1
eq_total_row = row
put_row(ws, row, "자본총계", {c: f"=SUM({c}{eq_start}:{c}{row-1})" for c in BS_ALL_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
row += 1
le_total_row = row
put_row(ws, row, "부채와자본총계",
        {c: f"={c}{liab_total_row}+{c}{eq_total_row}" for c in BS_ALL_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
row += 1
put_row(ws, row, f"{IDT}CHK (자산=부채+자본)",
        {c: f'=IF(ROUND({c}{asset_total_row}-{c}{le_total_row},1)=0,"OK","CHK")' for c in BS_ALL_COLS},
        font=F_PCT, fmt="General")
row += 2
ws[f"B{row}"] = "※ 시간 순서 배열: FY 컬럼 = 해당 연말(=4Q) 잔액. 실제 추출 시 사업보고서상 계정과목 전부 포함."
style_cell(ws[f"B{row}"], F_NOTE)

# ============================== CF ==============================
ws = wb.create_sheet("CF")
ws.sheet_properties.tabColor = "6A1B9A"
sheet_header(ws, "CF — 현금흐름표 (연결)", "(단위: 억원)", ISCF_HDR, gap_cols=(GAP_COL,))

# CF 컬럼 → BS 컬럼 매핑 (기말현금 대사용)
CF2BS = {}
for q in Q_LABELS:
    bs_p = f"FY20{q[-2:]}" if q.startswith("4Q") else q
    CF2BS[ISCF_Q_COLS[q]] = BS_COLS[bs_p]
for y in YEARS:
    CF2BS[ISCF_Y_COLS[y]] = BS_COLS[f"FY{y}"]

def cf_vals(item, row, link_is=False):
    if link_is:
        return {c: f"=IS!{c}{IS_NI_ROW}" for c in ISCF_COLS}
    vals = {}
    for q in EXTRACTED_Q:
        vals[ISCF_Q_COLS[q]] = CF_Q[q][item]
    for q, (ycol, qcols) in Q4_MAP.items():
        vals[ISCF_Q_COLS[q]] = f"={ycol}{row}-{qcols[0]}{row}-{qcols[1]}{row}-{qcols[2]}{row}"
    for y in YEARS:
        vals[ISCF_Y_COLS[y]] = CF_FY[f"FY{y}"][item]
    return vals

row = 7
ops_row = row
ops_items = ["당기순이익", "비현금항목 조정", "영업활동 자산부채의 증감", "이자수취", "이자지급", "법인세납부"]
r = row + 1
for it in ops_items:
    put_row(ws, r, f"{IDT}{it}", cf_vals(it, r, link_is=(it == "당기순이익")))
    r += 1
put_row(ws, ops_row, "영업활동현금흐름",
        {c: f"=SUM({c}{ops_row+1}:{c}{r-1})" for c in ISCF_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
row = r + 1

inv_row = row
inv_items = ["단기금융상품의 순증감", "유형자산의 취득", "유형자산의 처분", "무형자산의 취득"]
r = row + 1
for it in inv_items:
    put_row(ws, r, f"{IDT}{it}", cf_vals(it, r))
    r += 1
put_row(ws, inv_row, "투자활동현금흐름",
        {c: f"=SUM({c}{inv_row+1}:{c}{r-1})" for c in ISCF_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
row = r + 1

fin_row = row
fin_items = ["차입금의 차입", "차입금의 상환", "리스부채의 상환", "배당금의 지급"]
r = row + 1
for it in fin_items:
    put_row(ws, r, f"{IDT}{it}", cf_vals(it, r))
    r += 1
put_row(ws, fin_row, "재무활동현금흐름",
        {c: f"=SUM({c}{fin_row+1}:{c}{r-1})" for c in ISCF_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
row = r + 1

fx_row = row
put_row(ws, row, "외화환산으로 인한 현금의 변동", cf_vals("외화환산으로 인한 현금의 변동", row))
row += 1
delta_row = row
put_row(ws, row, "현금및현금성자산의 증감",
        {c: f"={c}{ops_row}+{c}{inv_row}+{c}{fin_row}+{c}{fx_row}" for c in ISCF_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
row += 1
begin_row = row
end_row = row + 1
begin_vals = cf_vals("기초현금", row)
for q, (ycol, qcols) in Q4_MAP.items():
    begin_vals[ISCF_Q_COLS[q]] = f"={qcols[2]}{end_row}"   # 4Q 기초 = 3Q 기말
put_row(ws, row, "기초현금및현금성자산", begin_vals)
row += 1
put_row(ws, row, "기말현금및현금성자산",
        {c: f"={c}{begin_row}+{c}{delta_row}" for c in ISCF_COLS},
        bold_label=True, fill=FILL_SUBTOTAL)
row += 1
put_row(ws, row, f"{IDT}CHK (기말현금=BS 현금)",
        {c: f'=IF(ROUND({c}{end_row}-BS!{CF2BS[c]}{BS_CASH_ROW},1)=0,"OK","CHK")' for c in ISCF_COLS},
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
ws.column_dimensions["B"].width = 32
for col in ["C", "D", "E", "F"]:
    ws.column_dimensions[col].width = 13

VAL_COLS = ["C", "D", "E", "F"]                       # FY2025(실적) + 2026E~2028E
VAL_LABELS = ["FY2025", "FY2026E", "FY2027E", "FY2028E"]

ws["B4"] = "실적 · 추정 (전년도 + 3개년)"
style_cell(ws["B4"], F_BOLD)
ws["B6"] = "구분"
style_cell(ws["B6"], F_HDR, fill=FILL_HDR, border=True, align="center")
for col, lab in zip(VAL_COLS, VAL_LABELS):
    c = ws[f"{col}6"]
    c.value = lab
    style_cell(c, F_HDR, fill=FILL_HDR, border=True, align="center")

def vput(r, label, vals_by_col, fmt=NUM, fill=None, font=F_BASE):
    ws[f"B{r}"] = label
    style_cell(ws[f"B{r}"], F_BASE, border=True)
    for col, v in vals_by_col.items():
        cell = ws[f"{col}{r}"]
        cell.value = v
        style_cell(cell, font, fmt=fmt, fill=fill, border=True)

NI_R, GROW_R, PAYOUT_R, BUYBACK_R, SHARES_R, EPS_R, DPS_R, TOTDIV_R, YIELD_R = range(7, 16)
vput(NI_R, "당기순이익 (억원)",
     {"C": f"=IS!{ISCF_Y_COLS[2025]}{IS_NI_ROW}",
      "D": f"=C{NI_R}*(1+D{GROW_R})", "E": f"=D{NI_R}*(1+E{GROW_R})", "F": f"=E{NI_R}*(1+F{GROW_R})"})
vput(GROW_R, f"{IDT}순이익 성장률 가정 (%)", {"C": "-", "D": 0.12, "E": 0.10, "F": 0.10},
     fmt=PCT, fill=FILL_INPUT)
vput(PAYOUT_R, "배당성향 가정 (%)", {c: 0.20 for c in VAL_COLS}, fmt=PCT, fill=FILL_INPUT)
vput(BUYBACK_R, "자사주 소각 주식수 (주)", {"C": 0, "D": 200000, "E": 200000, "F": 0},
     fmt=INT, fill=FILL_INPUT)
vput(SHARES_R, "기말 유통주식수 (주)",
     {"C": 20000000, "D": f"=C{SHARES_R}-D{BUYBACK_R}", "E": f"=D{SHARES_R}-E{BUYBACK_R}",
      "F": f"=E{SHARES_R}-F{BUYBACK_R}"}, fmt=INT)
style_cell(ws[f"C{SHARES_R}"], F_BASE, fmt=INT, fill=FILL_INPUT, border=True)
vput(EPS_R, "EPS (원)", {c: f"={c}{NI_R}*10^8/{c}{SHARES_R}" for c in VAL_COLS}, fmt=INT)
vput(DPS_R, "DPS (원)", {c: f"={c}{NI_R}*{c}{PAYOUT_R}*10^8/{c}{SHARES_R}" for c in VAL_COLS}, fmt=INT)
vput(TOTDIV_R, "총배당액 (억원)", {c: f"={c}{NI_R}*{c}{PAYOUT_R}" for c in VAL_COLS})
CUR_PRICE_CELL = "C30"
vput(YIELD_R, "배당수익률 (%)", {c: f"={c}{DPS_R}/${CUR_PRICE_CELL}" for c in VAL_COLS}, fmt=PCT)

ws["B17"] = "배당정책 / 자사주 (공시 내용 입력)"
style_cell(ws["B17"], F_BOLD)
for i, t in enumerate([
    "• 정책 적용기간: (입력)",
    "• 배당 기준: (입력)  — 위 배당성향 가정에 반영",
    "• 자사주 소각 계획: (입력)  — 위 소각 주식수에 반영",
]):
    ws[f"B{18+i}"] = t
    style_cell(ws[f"B{18+i}"], F_NOTE)

ws["B22"] = "PER 밴드 — 연도별 적정주가 (원)"
style_cell(ws["B22"], F_BOLD)
ws["B23"] = "PER (배)"
style_cell(ws["B23"], F_HDR, fill=FILL_HDR, border=True, align="center")
for col, lab in zip(VAL_COLS, VAL_LABELS):
    c = ws[f"{col}23"]
    c.value = lab
    style_cell(c, F_HDR, fill=FILL_HDR, border=True, align="center")
PER_START = 24
for i, per in enumerate([6, 7, 8, 9, 10, 11]):
    r = PER_START + i
    ws[f"B{r}"] = per
    style_cell(ws[f"B{r}"], F_BASE, fmt="0", fill=FILL_INPUT, border=True, align="center")
    for col in VAL_COLS:
        ws[f"{col}{r}"] = f"={col}${EPS_R}*$B{r}"
        style_cell(ws[f"{col}{r}"], F_BASE, fmt=INT, border=True)

ws["B30"] = "현재주가 (원)"
style_cell(ws["B30"], F_BASE, border=True)
ws[CUR_PRICE_CELL] = 8500
style_cell(ws[CUR_PRICE_CELL], F_BASE, fmt=INT, fill=FILL_INPUT, border=True)
PER8_ROW = PER_START + 2
ws["B31"] = "Upside vs PER 8x (%)"
style_cell(ws["B31"], F_BASE, border=True)
for col in VAL_COLS:
    ws[f"{col}31"] = f"={col}{PER8_ROW}/${CUR_PRICE_CELL}-1"
    style_cell(ws[f"{col}31"], F_BASE, fmt=PCT, border=True)

ws["B33"] = "※ 노란색 셀은 직접 입력(성장률·배당성향·소각주식수·주식수·PER·현재주가). 당기순이익 실적은 IS 시트 자동 연결."
style_cell(ws["B33"], F_NOTE)
ws["B34"] = "※ EPS·DPS는 기말 유통주식수 기준 단순 계산. 자사주 소각 시 주식수 감소가 다음 연도 EPS에 반영됨."
style_cell(ws["B34"], F_NOTE)

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
