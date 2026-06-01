"""pykrx 기반 실데이터 수집 계층.

이 모듈만 네트워크/외부 의존성을 가진다. 순수 로직(screen.py)과 분리해
테스트는 이 계층을 모킹/대체하여 수행한다.

주의: pykrx 1.2.x는 KRX 로그인이 필요해 환경변수 ``KRX_ID`` / ``KRX_PW``를
설정해야 한다. 미설정 시 빈 DataFrame이 반환된다.
"""
from __future__ import annotations

import datetime as dt

import pandas as pd

from .config import ScreenConfig
from .screen import assemble_universe

MARKETS = ("KOSPI", "KOSDAQ")


def _stock():
    # import를 함수 안으로 넣어, 테스트에서 pykrx 없이도 모듈 로드가 가능하게 한다.
    from pykrx import stock

    return stock


def resolve_date(date: str | None = None) -> str:
    """YYYYMMDD 영업일을 반환. None이면 가장 가까운 직전 영업일."""
    stock = _stock()
    if date:
        return date.replace("-", "")
    return stock.get_nearest_business_day_in_a_week()


def fetch_market(date: str, market: str) -> pd.DataFrame:
    """한 시장(KOSPI/KOSDAQ)의 universe 프레임을 만든다."""
    stock = _stock()
    ohlcv = stock.get_market_ohlcv_by_ticker(date, market=market)
    cap = stock.get_market_cap_by_ticker(date, market=market)
    sector = stock.get_market_sector_classifications(date, market=market)
    if sector is None or sector.empty:
        raise RuntimeError(
            f"{market} 업종 데이터가 비어 있습니다. KRX_ID/KRX_PW 설정 또는 "
            f"네트워크/영업일({date})을 확인하세요."
        )
    sector = sector.set_index("종목코드")
    return assemble_universe(ohlcv, cap, sector, market)


def fetch_universe(date: str) -> pd.DataFrame:
    """KOSPI + KOSDAQ 전체 universe를 결합해 반환."""
    frames = [fetch_market(date, m) for m in MARKETS]
    return pd.concat(frames)


def make_trend_provider(date: str, cfg: ScreenConfig):
    """3개월 수익률/대량거래일수를 계산하는 provider를 만든다.

    반환 함수: (ticker) -> (return_3m_pct | None, big_value_days | None)
    """
    stock = _stock()
    todate = dt.datetime.strptime(date, "%Y%m%d")
    fromdate = todate - dt.timedelta(days=cfg.lookback_days)
    from_s = fromdate.strftime("%Y%m%d")

    def provider(ticker: str) -> tuple[float | None, int | None]:
        try:
            df = stock.get_market_ohlcv_by_date(from_s, date, ticker)
        except Exception:
            return None, None
        if df is None or df.empty or "종가" not in df:
            return None, None
        closes = df["종가"].dropna()
        if len(closes) < 2 or closes.iloc[0] == 0:
            ret = None
        else:
            ret = (closes.iloc[-1] / closes.iloc[0] - 1.0) * 100.0
        days = None
        if "거래대금" in df:
            days = int((df["거래대금"] >= cfg.value_threshold_won).sum())
        return ret, days

    return provider
