"""테스트용 합성 universe (예시 메시지를 축약 재현)."""
from __future__ import annotations

import pandas as pd

from screener.config import EOK, JO
from screener.screen import UNIVERSE_COLUMNS


def _row(name, market, sector, close, change, value_eok, cap_jo):
    return {
        "종목명": name,
        "시장": market,
        "섹터": sector,
        "종가": float(close),
        "등락률": float(change),
        "거래대금": value_eok * EOK,
        "시가총액": cap_jo * JO,
    }


def sample_universe() -> pd.DataFrame:
    data = {
        # 건설: 통과 2개(대우건설, GS건설) + 동반상승 다수
        "047040": _row("대우건설", "KOSPI", "건설", 22550, 30.0, 19205, 9.4),
        "006360": _row("GS건설", "KOSPI", "건설", 37400, 29.9, 5523, 3.2),
        "037440": _row("희림", "KOSPI", "건설", 10000, 30.0, 50, 0.3),
        "002995": _row("금호건설우", "KOSPI", "건설", 8000, 30.0, 40, 0.1),
        "042940": _row("상지건설", "KOSPI", "건설", 5000, 29.9, 30, 0.2),
        "009410": _row("태영건설우", "KOSPI", "건설", 6000, 29.9, 25, 0.15),
        "000999": _row("건설하락주", "KOSPI", "건설", 3000, -2.0, 10, 0.5),
        # 통신장비: 통과 1개(대한광통신) + 동반상승
        "010170": _row("대한광통신", "KOSDAQ", "통신장비", 14690, 16.2, 9770, 2.3),
        "037760": _row("코위버", "KOSDAQ", "통신장비", 9000, 30.0, 100, 0.2),
        "073640": _row("에프알텍", "KOSDAQ", "통신장비", 7000, 29.9, 80, 0.15),
        # 반도체: 통과 1개(한미반도체, 시총 27조 < 30조)
        "042700": _row("한미반도체", "KOSPI", "반도체", 280500, 10.7, 2820, 27.0),
        "033640": _row("엠케이전자", "KOSDAQ", "반도체", 20000, 26.8, 150, 0.4),
        # 제외 대상: 시총 30조 초과
        "005930": _row("삼성전자", "KOSPI", "반도체", 80000, 12.0, 50000, 480.0),
        # 제외 대상: 거래대금 미달
        "111111": _row("소형급등주", "KOSDAQ", "기타", 1000, 25.0, 100, 0.1),
        # 제외 대상: 등락률 미달
        "222222": _row("대형거래주", "KOSPI", "기타", 50000, 3.0, 8000, 5.0),
    }
    df = pd.DataFrame.from_dict(data, orient="index")
    return df.reindex(columns=UNIVERSE_COLUMNS)


# 3개월 추세 provider (테스트용 고정값)
TREND = {
    "047040": (518.7, 24),
    "006360": (102.7, 2),
    "010170": (398.8, 19),
    "042700": (51.1, 44),
}


def fake_trend_provider(ticker: str):
    return TREND.get(ticker, (None, None))
