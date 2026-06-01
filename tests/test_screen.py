import pandas as pd

from screener.config import EOK, JO, ScreenConfig
from screener.screen import (
    analyze,
    assemble_universe,
    build_picks,
    enrich_trends,
    filter_universe,
    sector_companions,
)
from tests.fixtures import fake_trend_provider, sample_universe

CFG = ScreenConfig()


def test_filter_conditions_and_ordering():
    uni = sample_universe()
    filtered = filter_universe(uni, CFG)
    # 통과: 대우건설, GS건설, 대한광통신, 한미반도체
    assert list(filtered["종목명"]) == ["대우건설", "대한광통신", "GS건설", "한미반도체"]
    # 거래대금 내림차순 정렬 확인
    vals = list(filtered["거래대금"])
    assert vals == sorted(vals, reverse=True)


def test_filter_excludes_boundaries():
    uni = sample_universe()
    filtered = filter_universe(uni, CFG)
    names = set(filtered["종목명"])
    assert "삼성전자" not in names      # 시총 초과
    assert "소형급등주" not in names    # 거래대금 미달
    assert "대형거래주" not in names    # 등락률 미달
    assert "건설하락주" not in names    # 등락률 음수


def test_sector_companions_excludes_self_and_sorts():
    uni = sample_universe()
    comps = sector_companions(uni, "047040", "건설", CFG)
    names = [c.name for c in comps]
    assert "대우건설" not in names                 # 자기 자신 제외
    assert "건설하락주" not in names               # 하락 제외
    assert names[0] in {"희림", "금호건설우"}      # 등락률 30.0 최상위
    assert len(comps) <= CFG.companion_top_n


def test_analyze_theme_and_highlights():
    uni = sample_universe()
    filtered = filter_universe(uni, CFG)
    picks = build_picks(uni, filtered, CFG)
    enrich_trends(picks, fake_trend_provider)
    hl = analyze(uni, picks, CFG)

    # 건설은 통과 2개 → 테마 집중
    theme_sectors = {t.sector for t in hl.themes}
    assert "건설" in theme_sectors
    constr = next(t for t in hl.themes if t.sector == "건설")
    assert set(constr.filtered_names) == {"대우건설", "GS건설"}

    # 상한가: 등락률 >= 29.5 (대우건설 30.0, GS건설 29.9)
    limit_names = {p.name for p in hl.limit_ups}
    assert limit_names == {"대우건설", "GS건설"}

    # 단기 급등 지속: 3개월 수익률 >= 100 (대우건설 518.7, GS건설 102.7, 대한광통신 398.8)
    surge_names = {p.name for p in hl.persistent_surges}
    assert surge_names == {"대우건설", "GS건설", "대한광통신"}
    assert "한미반도체" not in surge_names          # 51.1% < 100


def test_assemble_universe_merges_frames():
    ohlcv = pd.DataFrame(
        {"종가": [100.0], "등락률": [12.0], "거래대금": [3000 * EOK]},
        index=["123456"],
    )
    cap = pd.DataFrame({"시가총액": [5 * JO]}, index=["123456"])
    sector = pd.DataFrame({"업종명": ["건설"], "종목명": ["테스트건설"]}, index=["123456"])
    uni = assemble_universe(ohlcv, cap, sector, "KOSPI")
    row = uni.loc["123456"]
    assert row["섹터"] == "건설"
    assert row["시장"] == "KOSPI"
    assert row["시가총액"] == 5 * JO
