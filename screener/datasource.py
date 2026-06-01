"""pykrx 기반 실데이터 수집 계층.

이 모듈만 네트워크/외부 의존성을 가진다. 순수 로직(screen.py)과 분리해
테스트는 이 계층을 모킹/대체하여 수행한다.

주의: pykrx 1.2.x는 KRX 로그인이 필요해 환경변수 ``KRX_ID`` / ``KRX_PW``를
설정해야 한다. 미설정 시 빈 DataFrame이 반환된다.
"""
from __future__ import annotations

import datetime as dt
import socket
import time

import pandas as pd

from .config import ScreenConfig
from .screen import assemble_universe

MARKETS = ("KOSPI", "KOSDAQ")
NETWORK_TIMEOUT = 30  # 초. 와이파이 끊김 시 무한 대기 방지


def _stock():
    # import를 함수 안으로 넣어, 테스트에서 pykrx 없이도 모듈 로드가 가능하게 한다.
    from pykrx import stock

    return stock


def _retry(fn, label: str, tries: int = 3):
    """일시적 네트워크 오류 시 백오프 재시도."""
    last = None
    for i in range(tries):
        try:
            return fn()
        except Exception as exc:  # noqa: BLE001
            last = exc
            if i < tries - 1:
                wait = 2 * (i + 1)
                print(f"[screener]   {label} 실패, {wait}초 후 재시도 ({i + 1}/{tries - 1}) — {exc}")
                time.sleep(wait)
    raise RuntimeError(f"{label} 반복 실패: {last}")


def resolve_date(date: str | None = None) -> str:
    """YYYYMMDD 영업일을 반환. None이면 가장 가까운 직전 영업일."""
    socket.setdefaulttimeout(NETWORK_TIMEOUT)
    stock = _stock()
    if date:
        return date.replace("-", "")
    return stock.get_nearest_business_day_in_a_week()


def fetch_market(date: str, market: str) -> pd.DataFrame:
    """한 시장(KOSPI/KOSDAQ)의 universe 프레임을 만든다."""
    socket.setdefaulttimeout(NETWORK_TIMEOUT)
    stock = _stock()
    print(f"[screener]   {market} 시세 받는 중...")
    ohlcv = _retry(lambda: stock.get_market_ohlcv_by_ticker(date, market=market), f"{market} 시세")
    print(f"[screener]   {market} 시가총액 받는 중...")
    cap = _retry(lambda: stock.get_market_cap_by_ticker(date, market=market), f"{market} 시총")
    print(f"[screener]   {market} 업종 받는 중...")
    sector = _retry(
        lambda: stock.get_market_sector_classifications(date, market=market), f"{market} 업종"
    )
    if sector is None or sector.empty:
        raise RuntimeError(
            f"{market} 업종 데이터가 비어 있습니다. KRX_ID/KRX_PW 설정 또는 "
            f"네트워크/영업일({date})을 확인하세요."
        )
    # pykrx 버전에 따라 종목코드가 컬럼일 수도, 이미 인덱스일 수도 있다.
    if "종목코드" in sector.columns:
        sector = sector.set_index("종목코드")
    sector.index = sector.index.astype(str)
    return assemble_universe(ohlcv, cap, sector, market)


def fetch_universe(date: str) -> pd.DataFrame:
    """KOSPI + KOSDAQ 전체 universe를 결합해 반환."""
    frames = []
    for m in MARKETS:
        print(f"[screener] {m} 데이터 수집 시작")
        frames.append(fetch_market(date, m))
    return pd.concat(frames)


def make_trend_provider(date: str, cfg: ScreenConfig):
    """3개월 수익률/대량거래일수를 계산하는 provider를 만든다.

    반환 함수: (ticker) -> (return_3m_pct | None, big_value_days | None)
    """
    socket.setdefaulttimeout(NETWORK_TIMEOUT)
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
