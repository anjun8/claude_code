from screener.config import ScreenConfig
from screener.formatter import (
    _won_to_eok,
    _won_to_jo,
    format_message,
    format_pick,
    split_for_telegram,
)
from screener.config import EOK, JO
from screener.screen import analyze, build_picks, enrich_trends, filter_universe
from tests.fixtures import fake_trend_provider, sample_universe

CFG = ScreenConfig()


def test_won_to_eok():
    assert _won_to_eok(19205 * EOK) == "19205"
    assert _won_to_eok(2820 * EOK) == "2820"


def test_won_to_jo_drops_trailing_zero():
    assert _won_to_jo(21 * JO) == "21"        # 정수면 소수점 생략
    assert _won_to_jo(9.4 * JO) == "9.4"      # 소수 유지
    assert _won_to_jo(27 * JO) == "27"


def test_format_pick_layout():
    uni = sample_universe()
    filtered = filter_universe(uni, CFG)
    picks = build_picks(uni, filtered, CFG)
    enrich_trends(picks, fake_trend_provider)
    daewoo = next(p for p in picks if p.name == "대우건설")
    text = format_pick(daewoo)
    assert "대우건설 047040 (KOSPI)" in text
    assert "섹터: 건설" in text
    assert "종가 22,550원  등락률 +30.0%" in text
    assert "거래대금 19205억  시총 9.4조" in text
    assert "3개월 수익률: +518.7%" in text
    assert "🔴 3개월 거래대금 2천억↑: 24회" in text
    assert "동일섹터 동반상승:" in text


def test_full_message_structure():
    uni = sample_universe()
    filtered = filter_universe(uni, CFG)
    picks = build_picks(uni, filtered, CFG)
    enrich_trends(picks, fake_trend_provider)
    hl = analyze(uni, picks, CFG)
    msg = format_message("2026-04-08", picks, hl, CFG)

    assert "📈 KOSPI/KOSDAQ 급등 스크리닝 (2026-04-08)" in msg
    assert "조건: 거래대금 2천억↑ | 등락률 +10%↑ | 시총 30조↓" in msg
    assert "총 4개 종목 해당" in msg
    assert "📌 오늘의 주목 포인트" in msg
    assert "건설 테마 집중 — 필터 통과 종목 2개" in msg
    assert "상한가 종목:" in msg
    assert "단기 급등 지속 종목 (추격 주의):" in msg


def test_empty_message():
    msg = format_message("2026-04-08", [], analyze(sample_universe(), [], CFG), CFG)
    assert "총 0개 종목 해당" in msg
    assert "조건을 만족하는 종목이 없습니다." in msg


def test_split_for_telegram():
    big = "\n\n".join([f"block-{i} " + "x" * 500 for i in range(20)])
    chunks = split_for_telegram(big, limit=1000)
    assert len(chunks) > 1
    assert all(len(c) <= 1000 for c in chunks)
