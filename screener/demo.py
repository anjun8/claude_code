"""오프라인 데모 데이터.

KRX/텔레그램 자격증명이나 네트워크 없이도 파이프라인 전체와 메시지 포맷을
즉시 확인할 수 있도록, 합성 universe와 가짜 3개월 추세를 제공한다.
``python -m screener.cli --demo`` 에서 사용한다.
"""
from __future__ import annotations

import pandas as pd

from .config import EOK, JO
from .screen import UNIVERSE_COLUMNS


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


def demo_universe() -> pd.DataFrame:
    """블로그 예시(2026-04-08)를 축약 재현한 합성 universe."""
    data = {
        # 건설: 통과 다수 + 동반상승
        "047040": _row("대우건설", "KOSPI", "건설", 22550, 30.0, 19205, 9.4),
        "000720": _row("현대건설", "KOSPI", "건설", 188700, 21.0, 10844, 21.0),
        "006360": _row("GS건설", "KOSPI", "건설", 37400, 29.9, 5523, 3.2),
        "375500": _row("DL이앤씨", "KOSPI", "건설", 95200, 25.9, 5485, 3.7),
        "037440": _row("희림", "KOSPI", "건설", 10000, 30.0, 50, 0.3),
        "002995": _row("금호건설우", "KOSPI", "건설", 8000, 30.0, 40, 0.1),
        "042940": _row("상지건설", "KOSPI", "건설", 5000, 29.9, 30, 0.2),
        "009410": _row("태영건설우", "KOSPI", "건설", 6000, 29.9, 25, 0.15),
        # 통신장비
        "010170": _row("대한광통신", "KOSDAQ", "통신장비", 14690, 16.2, 9770, 2.3),
        "037760": _row("코위버", "KOSDAQ", "통신장비", 9000, 30.0, 100, 0.2),
        "073640": _row("에프알텍", "KOSDAQ", "통신장비", 7000, 29.9, 80, 0.15),
        # 반도체 (시총 30조 이하만 통과)
        "042700": _row("한미반도체", "KOSPI", "반도체", 280500, 10.7, 2820, 27.0),
        "033640": _row("엠케이전자", "KOSDAQ", "반도체", 20000, 26.8, 150, 0.4),
        # 제외: 시총 초과
        "005930": _row("삼성전자", "KOSPI", "반도체", 80000, 12.0, 50000, 480.0),
        # 제외: 거래대금 미달 / 등락률 미달 / 하락
        "111111": _row("소형급등주", "KOSDAQ", "기타", 1000, 25.0, 100, 0.1),
        "222222": _row("대형거래주", "KOSPI", "기타", 50000, 3.0, 8000, 5.0),
    }
    df = pd.DataFrame.from_dict(data, orient="index")
    return df.reindex(columns=UNIVERSE_COLUMNS)


_TREND = {
    "047040": (518.7, 24),
    "000720": (151.6, 31),
    "006360": (102.7, 2),
    "375500": (141.6, 4),
    "010170": (398.8, 19),
    "042700": (51.1, 44),
}

_NEWS = {
    "대우건설": "상한가 종목대우건설-아이티엠반도체 이어 머큐리-GS건설 등 마감",
    "현대건설": "현대건설 '힐스테이트 안양펠루스', 전 세대 1순위 마감",
    "GS건설": "하나금융, GS건설과 생산적금융 대전환 추진협약",
    "DL이앤씨": "DL이앤씨, 이제 시작인 SMR 모멘텀 - 키움증권",
    "대한광통신": "대한광통신, 투자경고종목 지정예고",
    "한미반도체": "외국인이 韓증시에 돌아왔다…'무슨 종목' 많이 샀나",
}


def demo_trend_provider(ticker: str):
    return _TREND.get(ticker, (None, None))


def demo_news(name: str):
    return _NEWS.get(name)
