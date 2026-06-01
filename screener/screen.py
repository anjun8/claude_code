"""스크리닝 핵심 로직.

모든 함수는 pandas DataFrame을 입력으로 받는 순수 함수라 네트워크 없이
단위 테스트가 가능하다. 데이터 수집은 :mod:`screener.datasource`가 담당한다.

universe DataFrame 스키마 (index = 6자리 티커 문자열):
    종목명   : str
    시장     : "KOSPI" | "KOSDAQ"
    섹터     : str (KRX 업종명)
    종가     : float (원)
    등락률   : float (%, 예: 30.0)
    거래대금 : float (원)
    시가총액 : float (원)
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable

import pandas as pd

from .config import ScreenConfig

UNIVERSE_COLUMNS = ["종목명", "시장", "섹터", "종가", "등락률", "거래대금", "시가총액"]


@dataclass
class Companion:
    name: str
    change_pct: float


@dataclass
class Pick:
    """필터를 통과한 한 종목과 그에 딸린 분석 정보."""

    ticker: str
    name: str
    market: str
    sector: str
    close: float
    change_pct: float
    value_won: float
    marketcap_won: float
    # 3개월 추세 (datasource가 채움; 없으면 None)
    return_3m_pct: float | None = None
    big_value_days_3m: int | None = None
    # 부가 정보
    news_headline: str | None = None
    companions: list[Companion] = field(default_factory=list)


@dataclass
class SectorStrength:
    sector: str
    representative: str          # 대표 종목명(필터 통과 종목 중 거래대금 최대)
    companion_count: int         # 대표 종목 외 동반 상승 종목 수
    is_filtered_theme: bool      # 통과 종목이 2개 이상인 "테마 집중" 여부
    filtered_names: list[str] = field(default_factory=list)


@dataclass
class Highlights:
    themes: list[SectorStrength] = field(default_factory=list)        # 테마 집중(통과 2+)
    broad_sectors: list[SectorStrength] = field(default_factory=list)  # 섹터 전반 강세
    limit_ups: list[Pick] = field(default_factory=list)               # 상한가
    persistent_surges: list[Pick] = field(default_factory=list)       # 단기 급등 지속


def assemble_universe(
    ohlcv: pd.DataFrame,
    cap: pd.DataFrame,
    sector: pd.DataFrame,
    market: str,
) -> pd.DataFrame:
    """pykrx 원본 프레임들을 표준 universe 스키마로 병합한다.

    기대 컬럼:
        ohlcv  : 종가, 등락률, 거래대금  (index=티커)
        cap    : 시가총액                (index=티커)
        sector : 종목명, 업종명          (index=티커)
    """
    df = ohlcv[["종가", "등락률", "거래대금"]].copy()
    df["시가총액"] = cap["시가총액"]
    df["섹터"] = sector["업종명"]
    df["종목명"] = sector["종목명"]
    df["시장"] = market
    df = df.reindex(columns=UNIVERSE_COLUMNS)
    df = df.dropna(subset=["종가", "등락률", "거래대금", "시가총액"])
    df.index = df.index.astype(str)
    return df


def filter_universe(universe: pd.DataFrame, cfg: ScreenConfig) -> pd.DataFrame:
    """필터 조건을 통과한 종목을 거래대금 내림차순으로 반환한다."""
    mask = (
        (universe["거래대금"] >= cfg.value_threshold_won)
        & (universe["등락률"] >= cfg.change_threshold_pct)
        & (universe["시가총액"] <= cfg.marketcap_max_won)
    )
    return universe[mask].sort_values("거래대금", ascending=False)


def sector_companions(
    universe: pd.DataFrame, ticker: str, sector: str, cfg: ScreenConfig
) -> list[Companion]:
    """같은 섹터에서 상승한 종목을 등락률 내림차순 top-N으로(자기 자신 제외)."""
    same = universe[(universe["섹터"] == sector) & (universe["등락률"] > cfg.companion_rise_pct)]
    same = same[same.index != ticker].sort_values("등락률", ascending=False)
    out: list[Companion] = []
    for _, row in same.head(cfg.companion_top_n).iterrows():
        out.append(Companion(name=row["종목명"], change_pct=float(row["등락률"])))
    return out


def build_picks(
    universe: pd.DataFrame, filtered: pd.DataFrame, cfg: ScreenConfig
) -> list[Pick]:
    """필터 통과 종목을 Pick 객체로 변환(동반상승 포함, 3개월/뉴스는 이후 채움)."""
    picks: list[Pick] = []
    for ticker, row in filtered.iterrows():
        picks.append(
            Pick(
                ticker=str(ticker),
                name=row["종목명"],
                market=row["시장"],
                sector=row["섹터"],
                close=float(row["종가"]),
                change_pct=float(row["등락률"]),
                value_won=float(row["거래대금"]),
                marketcap_won=float(row["시가총액"]),
                companions=sector_companions(universe, str(ticker), row["섹터"], cfg),
            )
        )
    return picks


def analyze(
    universe: pd.DataFrame, picks: list[Pick], cfg: ScreenConfig
) -> Highlights:
    """'오늘의 주목 포인트'를 계산한다."""
    hl = Highlights()

    # 섹터별 통과 종목 그룹화 (거래대금 순서 유지)
    by_sector: dict[str, list[Pick]] = {}
    for p in picks:
        by_sector.setdefault(p.sector, []).append(p)

    for sector, members in by_sector.items():
        members_sorted = sorted(members, key=lambda x: x.value_won, reverse=True)
        rep = members_sorted[0]
        risers = universe[
            (universe["섹터"] == sector) & (universe["등락률"] > cfg.companion_rise_pct)
        ]
        companion_count = max(int(len(risers)) - 1, 0)
        strength = SectorStrength(
            sector=sector,
            representative=rep.name,
            companion_count=companion_count,
            is_filtered_theme=len(members) >= cfg.theme_min_count,
            filtered_names=[m.name for m in members_sorted],
        )
        if strength.is_filtered_theme:
            hl.themes.append(strength)
        hl.broad_sectors.append(strength)

    # 거래대금 큰 섹터부터 정렬
    hl.themes.sort(key=lambda s: s.companion_count, reverse=True)
    hl.broad_sectors.sort(key=lambda s: s.companion_count, reverse=True)

    # 상한가 / 단기 급등 지속
    hl.limit_ups = [p for p in picks if p.change_pct >= cfg.limit_up_pct]
    hl.persistent_surges = [
        p
        for p in picks
        if p.return_3m_pct is not None and p.return_3m_pct >= cfg.surge_return_pct
    ]
    hl.persistent_surges.sort(key=lambda p: p.return_3m_pct or 0, reverse=True)
    return hl


# 3개월 추세를 채워주는 provider 시그니처:
#   (ticker) -> (return_3m_pct | None, big_value_days_3m | None)
TrendProvider = Callable[[str], tuple[float | None, int | None]]


def enrich_trends(picks: list[Pick], provider: TrendProvider) -> None:
    """각 Pick에 3개월 수익률/대량거래일수를 채운다(in-place)."""
    for p in picks:
        ret, days = provider(p.ticker)
        p.return_3m_pct = ret
        p.big_value_days_3m = days
