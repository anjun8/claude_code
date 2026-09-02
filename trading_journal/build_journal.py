"""주식 매매일지 워크북 생성기.

전략 하나당 '거래장부_<전략>' + '대시보드_<전략>' 시트 한 쌍을 만든다.
STRATEGIES 리스트에 전략명을 추가하면 시트 쌍이 그대로 늘어난다.

    python trading_journal/build_journal.py
"""

from __future__ import annotations

import datetime as dt
import os

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------------------
# 설정
# ---------------------------------------------------------------------------

STRATEGIES = ["스윙트레이딩"]

FIRST_ROW = 6            # 거래장부 첫 데이터 행
LAST_ROW = 505           # 거래장부 마지막 데이터 행
HEADER_ROW = 5
MONTH_COUNT = 24         # 대시보드 월별 표 행 수
TICKER_ROWS = 15         # 대시보드 종목별 표 행 수

FONT = "Arial"

# 색
NAVY = "1F3864"
BLUE_HDR = "2F5597"
LIGHT = "D9E2F3"
INPUT_FILL = "FFF2CC"    # 사용자가 입력하는 칸
CALC_FILL = "F2F2F2"     # 자동 계산 칸
HELPER_FILL = "EDEDED"
EXAMPLE_FILL = "FCE4D6"
GREEN = "006100"
RED = "9C0006"

INPUT_FONT_COLOR = "0000FF"   # 파랑 = 직접 입력
CALC_FONT_COLOR = "000000"    # 검정 = 수식
LINK_FONT_COLOR = "008000"    # 초록 = 다른 시트 참조

# 표시 형식
MONEY = '#,##0;[Red](#,##0);"-"'
MONEY_PLAIN = '#,##0;(#,##0);"-"'
PRICE = '#,##0.##'
PCT = '0.00%;[Red](0.00%);"-"'
DAYS = '#,##0;;"-"'
DATE_FMT = 'yyyy-mm-dd'
MONTH_FMT = 'yyyy-mm'
RATIO = '0.00'

THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# ---------------------------------------------------------------------------
# 공통 헬퍼
# ---------------------------------------------------------------------------

def style_title(ws, cell_range: str, text: str) -> None:
    first = cell_range.split(":")[0]
    ws.merge_cells(cell_range)
    c = ws[first]
    c.value = text
    c.font = Font(name=FONT, size=14, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=NAVY)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[c.row].height = 26


def section(ws, row: int, col: int, text: str, width: int) -> None:
    """굵은 구분 머리말 한 줄."""
    ws.cell(row=row, column=col, value=text).font = Font(
        name=FONT, size=11, bold=True, color="FFFFFF")
    for i in range(width):
        cell = ws.cell(row=row, column=col + i)
        cell.fill = PatternFill("solid", fgColor=BLUE_HDR)
        cell.border = BOX


def kpi_block(ws, row: int, col: int, title: str, items: list[tuple[str, str, str]]) -> None:
    """라벨/값 2열 블록. items = [(라벨, 수식, 표시형식), ...]"""
    section(ws, row, col, title, 2)
    for i, (label, formula, fmt) in enumerate(items, start=1):
        lc = ws.cell(row=row + i, column=col, value=label)
        lc.font = Font(name=FONT, size=10)
        lc.border = BOX
        lc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        vc = ws.cell(row=row + i, column=col + 1, value=formula)
        vc.font = Font(name=FONT, size=10, bold=True, color=LINK_FONT_COLOR)
        vc.number_format = fmt
        vc.border = BOX
        vc.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        vc.fill = PatternFill("solid", fgColor=CALC_FILL)


def table_header(ws, row: int, col: int, headers: list[str]) -> None:
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=col + i, value=h)
        c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=BLUE_HDR)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX


def set_widths(ws, widths: dict[str, float]) -> None:
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


# ---------------------------------------------------------------------------
# 1. 사용법 시트
# ---------------------------------------------------------------------------

def build_guide(wb: Workbook) -> None:
    ws = wb.create_sheet("사용법")
    ws.sheet_view.showGridLines = False
    set_widths(ws, {"A": 3, "B": 20, "C": 96})

    style_title(ws, "B2:C2", "주식 매매일지 · 사용법")

    rows: list[tuple[str, str]] = [
        ("", ""),
        ("■ 시트 구성", ""),
        ("사용법", "이 시트. 입력 규칙과 색 규칙 설명."),
        ("설정", "전략명·초기자본·수수료율·대시보드 시작월 등 공통 파라미터. 노란 칸만 고치면 됨."),
        ("거래장부_<전략>", "매매 1건(매수~매도 한 사이클)이 1행. 실제로 입력하는 곳."),
        ("대시보드_<전략>", "해당 전략의 성과 요약. 전부 수식이므로 직접 고치지 말 것."),
        ("원본거래내역", "증권사 HTS/MTS에서 뽑은 거래내역 원본을 그대로 붙여넣는 보관용 시트."),
        ("", ""),
        ("■ 색 규칙", ""),
        ("노란 배경 / 파란 글씨", "직접 입력하는 칸."),
        ("회색 배경 / 검은 글씨", "자동 계산되는 칸. 덮어쓰면 수식이 깨짐."),
        ("초록 글씨", "다른 시트를 참조하는 값."),
        ("주황 배경", "예시 행. 실제 사용 전에 삭제할 것."),
        ("", ""),
        ("■ 입력 순서", ""),
        ("1", "매수 체결 직후: 종목코드·종목명·매수일자·매수단가·수량·매수수수료·매매유형·메모를 입력한다."),
        ("2", "이 시점에는 매도 관련 칸이 비어 있으므로 상태가 '보유중'으로 잡히고, 손익 계산에서 제외된다."),
        ("3", "매도 체결 후: 매도일자·매도단가·매도수수료·제세금을 채운다. 상태가 '청산'으로 바뀌며 손익·수익률·보유일수가 계산된다."),
        ("4", "분할매수/분할매도는 평균단가와 총수량으로 합쳐 한 행에 적거나, 매수 묶음별로 행을 나눠 적는다(둘 중 하나로 일관되게)."),
        ("5", "대시보드는 파일을 열 때 자동 재계산된다. 값이 안 바뀌면 F9(수동 계산 모드일 때)를 누른다."),
        ("", ""),
        ("■ 지표 정의", ""),
        ("순손익", "총매도금액 − 총매수금액 − 총비용(매수수수료+매도수수료+제세금). 세후·수수료후 실현손익."),
        ("수익률", "순손익 ÷ 총매수금액. 투입원금 대비 실현 수익률이며 보유기간으로 연환산하지 않은 값."),
        ("승률", "수익 거래 수 ÷ 청산 거래 수. 보유중 거래는 분모·분자 모두에서 제외."),
        ("손익비(Profit Factor)", "총수익 합계 ÷ |총손실 합계|. 1보다 크면 전략이 플러스."),
        ("평균 손익비(R/R)", "평균 수익 ÷ |평균 손실|. 한 번 이길 때 버는 돈이 한 번 질 때 잃는 돈의 몇 배인지."),
        ("기대값", "총 실현손익 ÷ 청산 거래 수. 거래 1건당 평균적으로 기대되는 금액."),
        ("보유일수", "청산 거래는 매도일자−매수일자. 보유중 거래는 오늘−매수일자(파일 열 때마다 갱신)."),
        ("최대 연속 수익/손실", "장부의 행 순서 기준. 매수일자 오름차순으로 정렬해 두어야 의미가 맞음."),
        ("", ""),
        ("■ 전략 추가 방법", ""),
        ("코드로 추가", "trading_journal/build_journal.py 의 STRATEGIES 리스트에 전략명을 추가하고 다시 실행하면 시트 쌍이 생성된다."),
        ("엑셀에서 추가", "거래장부/대시보드 시트를 각각 우클릭 → 이동/복사 → 복사본 만들기 후 이름을 바꾸고, "
                        "대시보드 수식의 시트 이름을 새 장부 이름으로 일괄 치환한다."),
        ("주의", "전략을 나누는 기준은 보유기간이 아니라 '매매 근거'다. 같은 종목이라도 근거가 다르면 다른 장부에 기록해야 통계가 섞이지 않는다."),
    ]

    r = 3
    for label, desc in rows:
        if label.startswith("■"):
            c = ws.cell(row=r, column=2, value=label)
            c.font = Font(name=FONT, size=11, bold=True, color=NAVY)
            ws.row_dimensions[r].height = 22
        elif label or desc:
            lc = ws.cell(row=r, column=2, value=label)
            lc.font = Font(name=FONT, size=10, bold=True)
            lc.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            dc = ws.cell(row=r, column=3, value=desc)
            dc.font = Font(name=FONT, size=10)
            dc.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        r += 1


# ---------------------------------------------------------------------------
# 2. 설정 시트
# ---------------------------------------------------------------------------

def build_settings(wb: Workbook, strategies: list[str]) -> None:
    ws = wb.create_sheet("설정")
    ws.sheet_view.showGridLines = False
    set_widths(ws, {"A": 3, "B": 24, "C": 18, "D": 62})

    style_title(ws, "B2:D2", "공통 설정")

    table_header(ws, 4, 2, ["항목", "값", "설명"])

    params: list[tuple[str, object, str, str]] = [
        ("계좌명", "하나증권 / 안현준", "@", "여러 계좌를 쓰면 계좌별로 파일을 나누는 편이 낫다."),
        ("기준 통화", "KRW", "@", "원화 기준. 해외주식은 별도 파일 권장."),
        ("초기 투자원금", 4000000, MONEY_PLAIN, "원본거래내역의 입금 합계. 계좌 전체 수익률 계산의 기준."),
        ("대시보드 시작월", dt.date(2025, 10, 1), MONTH_FMT,
         f"월별 성과 표의 첫 달. 여기서부터 {MONTH_COUNT}개월이 표시된다."),
        ("매수 수수료율", 0.00015, '0.000%', "참고값. 장부에는 실제 체결 수수료를 원 단위로 직접 입력한다."),
        ("매도 수수료율", 0.00015, '0.000%', "참고값."),
        ("증권거래세율", 0.0018, '0.000%', "2025년 코스피/코스닥 기준 참고값. 실제 부과액을 장부에 입력한다."),
        ("1거래 최대 손실 한도", 0.02, '0.0%', "총자본 대비 손절 한도. 리스크 관리 기준선(수식에는 쓰이지 않음)."),
    ]

    r = 5
    for name, value, fmt, desc in params:
        nc = ws.cell(row=r, column=2, value=name)
        nc.font = Font(name=FONT, size=10, bold=True)
        nc.border = BOX
        nc.alignment = Alignment(horizontal="left", vertical="center", indent=1)

        vc = ws.cell(row=r, column=3, value=value)
        vc.font = Font(name=FONT, size=10, bold=True, color=INPUT_FONT_COLOR)
        vc.fill = PatternFill("solid", fgColor=INPUT_FILL)
        vc.number_format = fmt
        vc.border = BOX
        vc.alignment = Alignment(horizontal="center", vertical="center")

        dc = ws.cell(row=r, column=4, value=desc)
        dc.font = Font(name=FONT, size=9, color="595959")
        dc.border = BOX
        dc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        r += 1

    # 전략 목록
    r += 1
    section(ws, r, 2, "전략 목록", 3)
    r += 1
    table_header(ws, r, 2, ["전략명", "장부 시트", "설명"])
    desc_map = {
        "가치투자": "저평가 판단 후 장기 보유. 보유기간 수개월~수년.",
        "스윙트레이딩": "수일~수주 보유. 추세/수급 기반 진입, 목표가·손절가 사전 설정.",
        "데이트레이딩": "당일 청산. 보유일수 0일.",
    }
    for s in strategies:
        r += 1
        ws.cell(row=r, column=2, value=s).font = Font(name=FONT, size=10, bold=True)
        ws.cell(row=r, column=3, value=f"거래장부_{s}").font = Font(name=FONT, size=10)
        ws.cell(row=r, column=4, value=desc_map.get(s, "")).font = Font(name=FONT, size=9, color="595959")
        for col in (2, 3, 4):
            ws.cell(row=r, column=col).border = BOX

    # 매매유형 코드값 (데이터 유효성 목록의 근거를 눈에 보이게 남겨둠)
    r += 2
    section(ws, r, 2, "매매유형 선택지", 3)
    r += 1
    ws.cell(row=r, column=2, value="목록").font = Font(name=FONT, size=10, bold=True)
    ws.cell(row=r, column=3, value=TRADE_TAGS_TEXT).font = Font(name=FONT, size=10)
    ws.cell(row=r, column=4,
            value="거래장부의 '매매유형' 열 드롭다운 값. 바꾸려면 각 장부 시트의 데이터 유효성 목록도 함께 수정.").font = \
        Font(name=FONT, size=9, color="595959")

    r += 2
    ws.cell(row=r, column=2,
            value="※ 수수료율·거래세율은 참고용 상수이며 손익 계산에 직접 쓰이지 않는다. "
                  "장부에는 증권사가 실제로 부과한 금액을 원 단위로 입력할 것.").font = \
        Font(name=FONT, size=9, italic=True, color="833C00")


TRADE_TAGS = ["돌파", "눌림목", "추세추종", "실적", "배당", "테마", "분할매수", "역추세", "기타"]
TRADE_TAGS_TEXT = ", ".join(TRADE_TAGS)


# ---------------------------------------------------------------------------
# 3. 거래장부 시트
# ---------------------------------------------------------------------------

LEDGER_COLS = [
    # (열, 헤더, 폭, 종류, 표시형식)
    ("A", "No",          5,  "calc",  '0'),
    ("B", "종목코드",     10, "input", '@'),
    ("C", "종목명",       14, "input", '@'),
    ("D", "매수일자",     12, "input", DATE_FMT),
    ("E", "매수단가",     11, "input", PRICE),
    ("F", "수량",         9,  "input", '#,##0'),
    ("G", "매수\n수수료",  9,  "input", MONEY_PLAIN),
    ("H", "매도일자",     12, "input", DATE_FMT),
    ("I", "매도단가",     11, "input", PRICE),
    ("J", "매도\n수수료",  9,  "input", MONEY_PLAIN),
    ("K", "제세금",       10, "input", MONEY_PLAIN),
    ("L", "상태",         9,  "calc",  '@'),
    ("M", "총매수금액",   13, "calc",  MONEY_PLAIN),
    ("N", "총매도금액",   13, "calc",  MONEY_PLAIN),
    ("O", "총비용",       10, "calc",  MONEY_PLAIN),
    ("P", "순손익",       13, "calc",  MONEY),
    ("Q", "수익률",       10, "calc",  PCT),
    ("R", "보유일수",     9,  "calc",  DAYS),
    ("S", "청산월",       9,  "calc",  MONTH_FMT),
    ("T", "매매유형",     11, "input", '@'),
    ("U", "메모 / 매매근거", 34, "input", '@'),
    ("V", "연속승\n(보조)", 8, "helper", '0;;"-"'),
    ("W", "연속패\n(보조)", 8, "helper", '0;;"-"'),
]

EXAMPLE_ROWS = [
    # 코드, 종목명, 매수일, 매수단가, 수량, 매수수수료, 매도일, 매도단가, 매도수수료, 제세금, 유형, 메모
    ("005930", "삼성전자", dt.date(2025, 10, 20), 68000, 30, 306,
     dt.date(2025, 11, 14), 74500, 335, 4023, "돌파",
     "예시) 전고점 돌파 + 거래대금 급증. 목표 75,000 / 손절 64,500"),
    ("035420", "NAVER", dt.date(2025, 11, 3), 215000, 8, 258,
     dt.date(2025, 11, 10), 203000, 243, 2924, "눌림목",
     "예시) 20일선 지지 실패로 손절. 진입 근거 조기 소멸"),
    ("000660", "SK하이닉스", dt.date(2025, 12, 1), 178000, 10, 267,
     None, None, None, None, "추세추종",
     "예시) 보유중 행. 매도 칸이 비면 손익 계산에서 제외됨"),
]


def build_ledger(wb: Workbook, strategy: str) -> str:
    name = f"거래장부_{strategy}"
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False

    last_col = LEDGER_COLS[-1][0]
    style_title(ws, f"A1:{last_col}1", f"{strategy} 거래장부")

    ws["A2"] = "전략"
    ws["A2"].font = Font(name=FONT, size=10, bold=True)
    ws["B2"] = strategy
    ws["B2"].font = Font(name=FONT, size=10, bold=True, color=NAVY)
    ws["D2"] = "청산 건수"
    ws["D2"].font = Font(name=FONT, size=10, bold=True)
    ws["E2"] = f'=COUNTIF($L${FIRST_ROW}:$L${LAST_ROW},"청산")'
    ws["E2"].font = Font(name=FONT, size=10, bold=True, color=NAVY)
    ws["F2"] = "보유중"
    ws["F2"].font = Font(name=FONT, size=10, bold=True)
    ws["G2"] = f'=COUNTIF($L${FIRST_ROW}:$L${LAST_ROW},"보유중")'
    ws["G2"].font = Font(name=FONT, size=10, bold=True, color=NAVY)
    ws["H2"] = "실현손익 합계"
    ws["H2"].font = Font(name=FONT, size=10, bold=True)
    ws["I2"] = f'=SUM($P${FIRST_ROW}:$P${LAST_ROW})'
    ws["I2"].font = Font(name=FONT, size=10, bold=True, color=NAVY)
    ws["I2"].number_format = MONEY

    ws["A3"] = ("입력: 노란 칸(파란 글씨)만 채운다.  ·  자동: 회색 칸은 수식이므로 덮어쓰지 말 것.  ·  "
                "매수 직후엔 매수 칸만 채우고(상태=보유중), 매도 후 매도 칸을 채우면 상태가 '청산'으로 바뀌며 손익이 계산된다.  ·  "
                "주황색 예시 3행은 삭제 후 사용.")
    ws["A3"].font = Font(name=FONT, size=9, italic=True, color="833C00")
    ws.merge_cells(f"A3:{last_col}3")

    # 헤더
    for col, header, width, kind, fmt in LEDGER_COLS:
        ws.column_dimensions[col].width = width
        c = ws[f"{col}{HEADER_ROW}"]
        c.value = header
        c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=BLUE_HDR if kind != "helper" else "808080")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    ws.row_dimensions[HEADER_ROW].height = 32

    def f(row: int) -> dict[str, str]:
        p = row - 1  # 직전 행 (연속 카운트용)
        return {
            "A": f'=IF($B{row}="","",ROW()-{HEADER_ROW})',
            "L": f'=IF($B{row}="","",IF($H{row}="","보유중","청산"))',
            "M": f'=IF($B{row}="","",$E{row}*$F{row})',
            "N": f'=IF($L{row}="청산",$I{row}*$F{row},"")',
            "O": f'=IF($B{row}="","",$G{row}+$J{row}+$K{row})',
            "P": f'=IF($L{row}="청산",$N{row}-$M{row}-$O{row},"")',
            "Q": f'=IF($L{row}="청산",IF($M{row}=0,"",$P{row}/$M{row}),"")',
            "R": (f'=IF($B{row}="","",IF($L{row}="청산",$H{row}-$D{row},'
                  f'IF($D{row}="","",TODAY()-$D{row})))'),
            "S": f'=IF($L{row}="청산",DATE(YEAR($H{row}),MONTH($H{row}),1),"")',
            "V": (f'=IF($B{row}="","",IF($L{row}<>"청산",N(V{p}),'
                  f'IF($P{row}>0,N(V{p})+1,0)))'),
            "W": (f'=IF($B{row}="","",IF($L{row}<>"청산",N(W{p}),'
                  f'IF($P{row}<0,N(W{p})+1,0)))'),
        }

    example_last = FIRST_ROW + len(EXAMPLE_ROWS) - 1

    for row in range(FIRST_ROW, LAST_ROW + 1):
        formulas = f(row)
        is_example = row <= example_last
        for col, header, width, kind, fmt in LEDGER_COLS:
            c = ws[f"{col}{row}"]
            c.number_format = fmt
            c.border = BOX
            c.alignment = Alignment(
                horizontal="left" if col in ("C", "U") else "center",
                vertical="center",
                indent=1 if col in ("C", "U") else 0,
            )
            if kind == "input":
                c.font = Font(name=FONT, size=10, color=INPUT_FONT_COLOR)
                c.fill = PatternFill("solid", fgColor=EXAMPLE_FILL if is_example else INPUT_FILL)
            elif kind == "helper":
                c.value = formulas[col]
                c.font = Font(name=FONT, size=9, color="808080")
                c.fill = PatternFill("solid", fgColor=HELPER_FILL)
            else:
                c.value = formulas[col]
                c.font = Font(name=FONT, size=10, color=CALC_FONT_COLOR,
                              bold=col in ("P", "Q"))
                c.fill = PatternFill("solid", fgColor=CALC_FILL)

    # 예시 데이터
    for i, ex in enumerate(EXAMPLE_ROWS):
        row = FIRST_ROW + i
        code, nm, bd, bp, qty, bf, sd, sp, sf, tax, tag, memo = ex
        for col, val in zip("BCDEFGHIJKTU",
                            [code, nm, bd, bp, qty, bf, sd, sp, sf, tax, tag, memo]):
            if val is not None:
                ws[f"{col}{row}"] = val

    # 합계 행
    total = LAST_ROW + 1
    ws[f"C{total}"] = "합계 / 평균"
    ws[f"C{total}"].font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
    for col, header, width, kind, fmt in LEDGER_COLS:
        c = ws[f"{col}{total}"]
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center")
    for col in ("M", "N", "O", "P"):
        ws[f"{col}{total}"] = f'=SUM({col}${FIRST_ROW}:{col}${LAST_ROW})'
        ws[f"{col}{total}"].number_format = MONEY
        ws[f"{col}{total}"].font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
    # 총 수익률은 '청산된' 매수금액만 분모로 쓴다 (보유중 원금을 섞으면 과소 계상됨)
    ws[f"Q{total}"] = (f'=IFERROR(P{total}/SUMIF($L${FIRST_ROW}:$L${LAST_ROW},"청산",'
                       f'$M${FIRST_ROW}:$M${LAST_ROW}),"")')
    ws[f"Q{total}"].number_format = PCT
    ws[f"Q{total}"].font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
    ws[f"R{total}"] = (f'=IFERROR(AVERAGEIFS($R${FIRST_ROW}:$R${LAST_ROW},'
                       f'$L${FIRST_ROW}:$L${LAST_ROW},"청산"),"")')
    ws[f"R{total}"].number_format = '#,##0.0'
    ws[f"R{total}"].font = Font(name=FONT, size=10, bold=True, color="FFFFFF")

    # 매매유형 드롭다운
    dv = DataValidation(type="list", formula1='"' + ",".join(TRADE_TAGS) + '"',
                        allow_blank=True, showDropDown=False)
    dv.error = "설정 시트의 '매매유형 선택지'에 있는 값 중에서 고를 것."
    dv.errorTitle = "매매유형 값 오류"
    ws.add_data_validation(dv)
    dv.add(f"T{FIRST_ROW}:T{LAST_ROW}")

    # 조건부 서식
    rng_pnl = f"P{FIRST_ROW}:P{LAST_ROW}"
    rng_ret = f"Q{FIRST_ROW}:Q{LAST_ROW}"
    for rng in (rng_pnl, rng_ret):
        ws.conditional_formatting.add(rng, CellIsRule(
            operator="greaterThan", formula=["0"],
            font=Font(color=GREEN, bold=True),
            fill=PatternFill("solid", bgColor="C6EFCE")))
        ws.conditional_formatting.add(rng, CellIsRule(
            operator="lessThan", formula=["0"],
            font=Font(color=RED, bold=True),
            fill=PatternFill("solid", bgColor="FFC7CE")))
    ws.conditional_formatting.add(f"L{FIRST_ROW}:L{LAST_ROW}", CellIsRule(
        operator="equal", formula=['"보유중"'],
        font=Font(color="9C6500", bold=True),
        fill=PatternFill("solid", bgColor="FFEB9C")))

    ws.freeze_panes = f"D{FIRST_ROW}"
    ws.auto_filter.ref = f"A{HEADER_ROW}:{last_col}{LAST_ROW}"
    return name


# ---------------------------------------------------------------------------
# 4. 대시보드 시트
# ---------------------------------------------------------------------------

def build_dashboard(wb: Workbook, strategy: str, ledger: str) -> None:
    ws = wb.create_sheet(f"대시보드_{strategy}")
    ws.sheet_view.showGridLines = False
    # B~I는 KPI 블록(라벨/값 3쌍)과 아래쪽 표(월별·TOP5·종목별)가 함께 쓰는 열이다.
    # 홀수 열은 라벨, 짝수 열은 값 폭으로 잡되 표의 금액 칸이 잘리지 않을 만큼은 넓혀 둔다.
    set_widths(ws, {
        "A": 2.5, "B": 22, "C": 15, "D": 15, "E": 22, "F": 15, "G": 15,
        "H": 22, "I": 15, "J": 3,
        "K": 12, "L": 11, "M": 14, "N": 14, "O": 10, "P": 13, "Q": 13, "R": 11,
    })

    L = f"'{ledger}'!"
    rng = lambda col: f"{L}${col}${FIRST_ROW}:${col}${LAST_ROW}"  # noqa: E731
    CODE, NAME, BUYD, BUYP = rng("B"), rng("C"), rng("D"), rng("E")
    QTY, SELLD = rng("F"), rng("H")
    STATUS, BUYAMT, COST = rng("L"), rng("M"), rng("O")
    PNL, RET, HOLD, MON = rng("P"), rng("Q"), rng("R"), rng("S")
    WSTREAK, LSTREAK = rng("V"), rng("W")

    style_title(ws, "B2:R2", f"{strategy} 성과 대시보드")
    ws["B3"] = "기준일"
    ws["B3"].font = Font(name=FONT, size=9, bold=True, color="595959")
    ws["C3"] = "=TODAY()"
    ws["C3"].number_format = DATE_FMT
    ws["C3"].font = Font(name=FONT, size=9, color="595959")
    ws["E3"] = f"장부 시트: {ledger}  ·  모든 값은 수식으로 자동 계산됨 (직접 수정 금지)"
    ws["E3"].font = Font(name=FONT, size=9, italic=True, color="595959")

    # ---------------- KPI 3블록 ----------------
    top = 5
    kpi_block(ws, top, 2, "종합 성과", [
        ("청산 거래 수", f'=COUNTIF({STATUS},"청산")', '#,##0'),
        ("보유중 거래 수", f'=COUNTIF({STATUS},"보유중")', '#,##0'),
        ("수익 거래 수", f'=COUNTIF({PNL},">0")', '#,##0'),
        ("손실 거래 수", f'=COUNTIF({PNL},"<0")', '#,##0'),
        ("승률", '=IFERROR($C$8/$C$6,0)', '0.0%'),
        ("총 실현손익", f'=SUM({PNL})', MONEY),
        ("총 수익 합계", f'=SUMIF({PNL},">0")', MONEY_PLAIN),
        ("총 손실 합계", f'=SUMIF({PNL},"<0")', MONEY),
        ("손익비 (Profit Factor)", '=IFERROR($C$12/ABS($C$13),0)', RATIO),
        ("기대값 (거래당)", '=IFERROR($C$11/$C$6,0)', MONEY),
        ("총 거래비용", f'=SUM({COST})', MONEY_PLAIN),
    ])

    kpi_block(ws, top, 5, "평균 지표", [
        ("평균 수익", f'=IFERROR(AVERAGEIF({PNL},">0"),0)', MONEY_PLAIN),
        ("평균 손실", f'=IFERROR(AVERAGEIF({PNL},"<0"),0)', MONEY),
        ("평균 손익비 (R/R)", '=IFERROR($F$6/ABS($F$7),0)', RATIO),
        ("전체 평균 수익률", f'=IFERROR(AVERAGE({RET}),0)', PCT),
        ("수익 거래 평균 수익률", f'=IFERROR(AVERAGEIFS({RET},{PNL},">0"),0)', PCT),
        ("손실 거래 평균 수익률", f'=IFERROR(AVERAGEIFS({RET},{PNL},"<0"),0)', PCT),
        ("평균 보유일수", f'=IFERROR(AVERAGEIFS({HOLD},{STATUS},"청산"),0)', '#,##0.0'),
        ("수익 거래 평균 보유일수", f'=IFERROR(AVERAGEIFS({HOLD},{PNL},">0"),0)', '#,##0.0'),
        ("손실 거래 평균 보유일수", f'=IFERROR(AVERAGEIFS({HOLD},{PNL},"<0"),0)', '#,##0.0'),
        ("최대 연속 수익", f'=IFERROR(MAX({WSTREAK}),0)', '#,##0'),
        ("최대 연속 손실", f'=IFERROR(MAX({LSTREAK}),0)', '#,##0'),
    ])

    kpi_block(ws, top, 8, "최고 / 최악", [
        ("최대 수익 거래", f'=IFERROR(MAX({PNL}),0)', MONEY_PLAIN),
        ("  └ 종목", f'=IFERROR(INDEX({NAME},MATCH(MAX({PNL}),{PNL},0)),"-")', '@'),
        ("최대 손실 거래", f'=IFERROR(MIN({PNL}),0)', MONEY),
        ("  └ 종목", f'=IFERROR(INDEX({NAME},MATCH(MIN({PNL}),{PNL},0)),"-")', '@'),
        ("최고 수익률", f'=IFERROR(MAX({RET}),0)', PCT),
        ("최저 수익률", f'=IFERROR(MIN({RET}),0)', PCT),
        ("최장 보유일수 (청산)", f'=IFERROR(_xlfn.MAXIFS({HOLD},{STATUS},"청산"),0)', '#,##0'),
        ("총 매수금액 (누적)", f'=SUM({BUYAMT})', MONEY_PLAIN),
        ("현재 보유중 투입원금", f'=SUMIF({STATUS},"보유중",{BUYAMT})', MONEY_PLAIN),
        ("초기자본 대비 수익률", "=IFERROR($C$11/'설정'!$C$7,0)", PCT),
        ("비용 / 총수익 비중", '=IFERROR($C$16/$C$12,0)', '0.0%'),
    ])

    # ---------------- 월별 성과 ----------------
    mrow = top + 14           # 19
    section(ws, mrow, 2, "월별 성과", 8)
    ws.cell(row=mrow, column=2).value = "월별 성과"
    hrow = mrow + 1
    table_header(ws, hrow, 2, ["월", "청산\n건수", "실현손익", "누적 실현손익",
                               "승률", "최대 수익", "최대 손실", "평균 수익률"])
    ws.row_dimensions[hrow].height = 30

    first_m = hrow + 1
    last_m = first_m + MONTH_COUNT - 1
    for i in range(MONTH_COUNT):
        r = first_m + i
        if i == 0:
            month_f = "=DATE(YEAR('설정'!$C$8),MONTH('설정'!$C$8),1)"
        else:
            month_f = f"=EDATE($B{r - 1},1)"
        vals = [
            (2, month_f, MONTH_FMT),
            (3, f'=COUNTIFS({MON},$B{r})', '#,##0;;"-"'),
            (4, f'=SUMIFS({PNL},{MON},$B{r})', MONEY),
            (5, f'=SUM($D${first_m}:$D{r})', MONEY),
            (6, f'=IFERROR(COUNTIFS({MON},$B{r},{PNL},">0")/$C{r},"")', '0.0%;;"-"'),
            (7, f'=IFERROR(_xlfn.MAXIFS({PNL},{MON},$B{r}),0)', MONEY_PLAIN),
            (8, f'=IFERROR(_xlfn.MINIFS({PNL},{MON},$B{r}),0)', MONEY),
            (9, f'=IFERROR(AVERAGEIFS({RET},{MON},$B{r}),"")', PCT),
        ]
        for col, formula, fmt in vals:
            c = ws.cell(row=r, column=col, value=formula)
            c.number_format = fmt
            c.font = Font(name=FONT, size=10, bold=(col == 2))
            c.border = BOX
            c.alignment = Alignment(horizontal="center", vertical="center")
            if col != 2:
                c.fill = PatternFill("solid", fgColor=CALC_FILL)

    # 월별 합계
    tr = last_m + 1
    ws.cell(row=tr, column=2, value="합계")
    for col in range(2, 10):
        c = ws.cell(row=tr, column=col)
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=tr, column=3, value=f'=SUM($C${first_m}:$C${last_m})').number_format = '#,##0'
    ws.cell(row=tr, column=4, value=f'=SUM($D${first_m}:$D${last_m})').number_format = MONEY
    ws.cell(row=tr, column=5, value=f'=$E${last_m}').number_format = MONEY
    ws.cell(row=tr, column=7, value=f'=IFERROR(MAX($G${first_m}:$G${last_m}),0)').number_format = MONEY_PLAIN
    ws.cell(row=tr, column=8, value=f'=IFERROR(MIN($H${first_m}:$H${last_m}),0)').number_format = MONEY
    for col in (3, 4, 5, 7, 8):
        ws.cell(row=tr, column=col).font = Font(name=FONT, size=10, bold=True, color="FFFFFF")

    for rng_cf in (f"D{first_m}:D{last_m}", f"E{first_m}:E{last_m}"):
        ws.conditional_formatting.add(rng_cf, CellIsRule(
            operator="greaterThan", formula=["0"], font=Font(color=GREEN, bold=True)))
        ws.conditional_formatting.add(rng_cf, CellIsRule(
            operator="lessThan", formula=["0"], font=Font(color=RED, bold=True)))

    # ---------------- 차트 ----------------
    bar = BarChart()
    bar.type = "col"
    bar.title = "월별 실현손익"
    bar.y_axis.title = "원"
    bar.height, bar.width = 7.5, 17
    bar.add_data(Reference(ws, min_col=4, min_row=hrow, max_row=last_m), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=2, min_row=first_m, max_row=last_m))
    bar.gapWidth = 40
    ws.add_chart(bar, f"K{mrow}")

    line = LineChart()
    line.title = "누적 실현손익"
    line.y_axis.title = "원"
    line.height, line.width = 7.5, 17
    line.add_data(Reference(ws, min_col=5, min_row=hrow, max_row=last_m), titles_from_data=True)
    line.set_categories(Reference(ws, min_col=2, min_row=first_m, max_row=last_m))
    # 스무딩을 켜두면 보간 곡선이 실제 최대·최소를 넘어서 그려져 손익 곡선을 왜곡한다.
    for s in line.series:
        s.smooth = False
    ws.add_chart(line, f"K{mrow + 16}")

    # ---------------- TOP 5 ----------------
    def top_table(start: int, title: str, fn: str) -> None:
        """순손익 상/하위 5건. 순손익 칸(F)에서 한 번만 뽑고 나머지 열은 그 값을 되참조한다.

        LARGE/SMALL은 값이 5개에 못 미치면 나머지 순위를 반대 부호 거래로 채우므로,
        수익 표는 양수만 / 손실 표는 음수만 남기도록 부호로 한 번 더 거른다.
        """
        section(ws, start, 2, title, 7)
        headers = ["순위", "종목명", "매수일자", "매도일자", "순손익", "수익률", "보유일수"]
        table_header(ws, start + 1, 2, headers)
        sign = ">" if fn == "LARGE" else "<"
        for k in range(1, 6):
            r = start + 1 + k
            pick = f'{fn}({PNL},$B{r})'
            hit = f'MATCH($F{r},{PNL},0)'
            cells = [
                (2, k, '0"위"'),
                (3, f'=IF($F{r}="","-",IFERROR(INDEX({NAME},{hit}),"-"))', '@'),
                (4, f'=IF($F{r}="","",IFERROR(INDEX({BUYD},{hit}),""))', DATE_FMT),
                (5, f'=IF($F{r}="","",IFERROR(INDEX({SELLD},{hit}),""))', DATE_FMT),
                (6, f'=IFERROR(IF({pick}{sign}0,{pick},""),"")', MONEY),
                (7, f'=IF($F{r}="","",IFERROR(INDEX({RET},{hit}),""))', PCT),
                (8, f'=IF($F{r}="","",IFERROR(INDEX({HOLD},{hit}),""))', '#,##0'),
            ]
            for col, val, fmt in cells:
                c = ws.cell(row=r, column=col, value=val)
                c.number_format = fmt
                c.font = Font(name=FONT, size=10, bold=(col == 6))
                c.border = BOX
                c.alignment = Alignment(horizontal="center", vertical="center")
                if col != 2:
                    c.fill = PatternFill("solid", fgColor=CALC_FILL)
        ws.conditional_formatting.add(f"F{start + 2}:G{start + 6}", CellIsRule(
            operator="greaterThan", formula=["0"], font=Font(color=GREEN, bold=True)))
        ws.conditional_formatting.add(f"F{start + 2}:G{start + 6}", CellIsRule(
            operator="lessThan", formula=["0"], font=Font(color=RED, bold=True)))

    best = tr + 2
    top_table(best, "수익 TOP 5", "LARGE")
    worst = best + 8
    top_table(worst, "손실 TOP 5", "SMALL")

    # ---------------- 종목별 요약 ----------------
    tick = worst + 9
    section(ws, tick, 2, "종목별 요약  (종목명 칸에 직접 입력)", 7)
    table_header(ws, tick + 1, 2,
                 ["종목명\n(입력)", "청산\n건수", "실현손익", "승률",
                  "평균 수익률", "평균 보유일수", "매수금액\n(청산분)"])
    ws.row_dimensions[tick + 1].height = 30
    for i in range(TICKER_ROWS):
        r = tick + 2 + i
        nc = ws.cell(row=r, column=2)
        nc.fill = PatternFill("solid", fgColor=INPUT_FILL)
        nc.font = Font(name=FONT, size=10, color=INPUT_FONT_COLOR)
        nc.border = BOX
        nc.alignment = Alignment(horizontal="center", vertical="center")
        cells = [
            (3, f'=IF($B{r}="","",COUNTIFS({NAME},$B{r},{STATUS},"청산"))', '#,##0;;"-"'),
            (4, f'=IF($B{r}="","",SUMIFS({PNL},{NAME},$B{r}))', MONEY),
            (5, f'=IFERROR(COUNTIFS({NAME},$B{r},{PNL},">0")/$C{r},"")', '0.0%;;"-"'),
            (6, f'=IFERROR(AVERAGEIFS({RET},{NAME},$B{r}),"")', PCT),
            (7, f'=IFERROR(AVERAGEIFS({HOLD},{NAME},$B{r},{STATUS},"청산"),"")', '#,##0.0'),
            (8, f'=IF($B{r}="","",SUMIFS({BUYAMT},{NAME},$B{r},{STATUS},"청산"))', MONEY_PLAIN),
        ]
        for col, val, fmt in cells:
            c = ws.cell(row=r, column=col, value=val)
            c.number_format = fmt
            c.font = Font(name=FONT, size=10)
            c.border = BOX
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.fill = PatternFill("solid", fgColor=CALC_FILL)
    ws.conditional_formatting.add(f"D{tick + 2}:D{tick + 1 + TICKER_ROWS}", CellIsRule(
        operator="greaterThan", formula=["0"], font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(f"D{tick + 2}:D{tick + 1 + TICKER_ROWS}", CellIsRule(
        operator="lessThan", formula=["0"], font=Font(color=RED, bold=True)))

    # ---------------- 매매유형별 요약 ----------------
    tag = tick + TICKER_ROWS + 3
    section(ws, tag, 2, "매매유형별 요약", 7)
    table_header(ws, tag + 1, 2,
                 ["매매유형", "청산\n건수", "실현손익", "승률",
                  "평균 수익률", "평균 보유일수", "매수금액\n(청산분)"])
    ws.row_dimensions[tag + 1].height = 30
    TAG_RNG = rng("T")
    for i, t in enumerate(TRADE_TAGS):
        r = tag + 2 + i
        nc = ws.cell(row=r, column=2, value=t)
        nc.font = Font(name=FONT, size=10, bold=True)
        nc.border = BOX
        nc.alignment = Alignment(horizontal="center", vertical="center")
        cells = [
            (3, f'=COUNTIFS({TAG_RNG},$B{r},{STATUS},"청산")', '#,##0;;"-"'),
            (4, f'=SUMIFS({PNL},{TAG_RNG},$B{r})', MONEY),
            (5, f'=IFERROR(COUNTIFS({TAG_RNG},$B{r},{PNL},">0")/$C{r},"")', '0.0%;;"-"'),
            (6, f'=IFERROR(AVERAGEIFS({RET},{TAG_RNG},$B{r}),"")', PCT),
            (7, f'=IFERROR(AVERAGEIFS({HOLD},{TAG_RNG},$B{r},{STATUS},"청산"),"")', '#,##0.0'),
            (8, f'=SUMIFS({BUYAMT},{TAG_RNG},$B{r},{STATUS},"청산")', MONEY_PLAIN),
        ]
        for col, val, fmt in cells:
            c = ws.cell(row=r, column=col, value=val)
            c.number_format = fmt
            c.font = Font(name=FONT, size=10)
            c.border = BOX
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.fill = PatternFill("solid", fgColor=CALC_FILL)
    ws.conditional_formatting.add(f"D{tag + 2}:D{tag + 1 + len(TRADE_TAGS)}", CellIsRule(
        operator="greaterThan", formula=["0"], font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(f"D{tag + 2}:D{tag + 1 + len(TRADE_TAGS)}", CellIsRule(
        operator="lessThan", formula=["0"], font=Font(color=RED, bold=True)))

    note = tag + len(TRADE_TAGS) + 3
    ws.cell(row=note, column=2,
            value="※ TOP 5 표는 순손익이 완전히 동일한 거래가 둘 이상이면 첫 번째 거래만 반복 표시된다. "
                  "※ 최대 연속 수익/손실은 장부의 행 순서 기준이므로 매수일자 오름차순 정렬을 유지할 것. "
                  "※ '초기자본 대비 수익률'은 설정 시트의 초기 투자원금을 분모로 쓴다.").font = \
        Font(name=FONT, size=9, italic=True, color="833C00")

    ws.freeze_panes = "A5"


# ---------------------------------------------------------------------------
# 5. 원본거래내역 시트 (증권사 출력 원본 보관)
# ---------------------------------------------------------------------------

RAW_HEADERS = [
    ("거래일자", "처리지점", 13),
    ("거래번호", "거래원번", 10),
    ("거래적요", "관련계좌", 18),
    ("종목코드", "종목명", 14),
    ("수량", "단가", 11),
    ("수수료", "제세금", 11),
    ("세전이자", "신용이자", 11),
    ("거래금액", "반영금액", 14),
    ("미수발생/변제금", "융자/대주금액", 14),
    ("질권자", "질권설정수량", 12),
    ("예수금잔고", "유가증권잔고", 14),
    ("거래원금", "원금잔액", 13),
    ("매수일자(질권설정일/출고일)", "만기일자", 15),
    ("경과일수", "적용금리", 10),
    ("통화구분", "외화예수금잔고", 13),
]

RAW_SAMPLE = [
    # (윗줄, 아랫줄) — 사용자가 제공한 캡처의 실제 값
    (["2025-12-09", 1, "전자계좌입금", "", "", 0, 0, 2000000, 0, "", 2000000, 0, "", 0, ""],
     ["디지털센터", 0, "하나은행/안현준", "", "", 0, 0, 2000000, 0, "", 0, 0, "", 0, 0]),
    (["2025-10-17", 8, "전자계좌입금", "", "", 0, 0, 2000000, 0, "", 4000297, 0, "", 0, ""],
     ["디지털센터", 0, "하나은행/안현준", "", "", 0, 0, 2000000, 0, "", 0, 0, "", 0, 0]),
    (["2025-10-19", 10, "전자계좌입금", "", "", 0, 0, 1000000, 0, "", 1000000, 0, "", 0, ""],
     ["디지털센터", 0, "하나은행/안현준", "", "", 0, 0, 1000000, 0, "", 0, 0, "", 0, 0]),
]

RAW_ROWS = 300


def build_raw(wb: Workbook) -> None:
    ws = wb.create_sheet("원본거래내역")
    ws.sheet_view.showGridLines = False
    ncol = len(RAW_HEADERS)
    last_col = get_column_letter(ncol)

    style_title(ws, f"A1:{last_col}1", "증권사 거래내역 원본 (보관용)")
    ws["A2"] = ("HTS/MTS에서 내려받은 거래내역을 3행부터 그대로 붙여넣는다. "
                "한 거래가 두 줄(윗줄/아랫줄)로 구성된 증권사 양식 그대로 보관하는 시트이며, "
                "손익 계산에는 사용되지 않는다. 매매 손익은 거래장부 시트에 정리해서 입력할 것.")
    ws["A2"].font = Font(name=FONT, size=9, italic=True, color="833C00")
    ws.merge_cells(f"A2:{last_col}2")

    for i, (h1, h2, width) in enumerate(RAW_HEADERS, start=1):
        col = get_column_letter(i)
        ws.column_dimensions[col].width = width
        for r, text in ((3, h1), (4, h2)):
            c = ws.cell(row=r, column=i, value=text)
            c.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=BLUE_HDR if r == 3 else "5B7DB1")
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = BOX
    ws.row_dimensions[3].height = 28
    ws.row_dimensions[4].height = 28

    r = 5
    for upper, lower in RAW_SAMPLE:
        for offset, line in ((0, upper), (1, lower)):
            for i, val in enumerate(line, start=1):
                c = ws.cell(row=r + offset, column=i, value=val if val != "" else None)
                c.font = Font(name=FONT, size=9, color=INPUT_FONT_COLOR)
                c.fill = PatternFill("solid", fgColor=EXAMPLE_FILL)
                c.border = BOX
                c.alignment = Alignment(horizontal="center", vertical="center")
                if isinstance(val, (int, float)):
                    c.number_format = MONEY_PLAIN
        r += 2

    for row in range(r, 5 + RAW_ROWS):
        for i in range(1, ncol + 1):
            c = ws.cell(row=row, column=i)
            c.font = Font(name=FONT, size=9, color=INPUT_FONT_COLOR)
            c.border = BOX
            c.alignment = Alignment(horizontal="center", vertical="center")
            if (row - r) % 2 == 0:
                c.fill = PatternFill("solid", fgColor="FFFFFF")
            else:
                c.fill = PatternFill("solid", fgColor="FAFAFA")

    ws.freeze_panes = "A5"


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def build(path: str, strategies: list[str]) -> None:
    wb = Workbook()
    wb.remove(wb.active)

    build_guide(wb)
    build_settings(wb, strategies)
    for s in strategies:
        ledger = build_ledger(wb, s)
        build_dashboard(wb, s, ledger)
    build_raw(wb)

    wb.active = 0
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    wb.save(path)
    print(f"saved: {path}")


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    build(os.path.join(here, "주식_매매일지.xlsx"), STRATEGIES)
