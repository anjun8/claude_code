"""스크리닝 결과를 텔레그램 메시지 텍스트로 변환한다."""
from __future__ import annotations

from .config import EOK, JO, ScreenConfig
from .screen import Highlights, Pick

DIVIDER = "──────────────────────────────"


def _won_to_eok(won: float) -> str:
    """원 → '억' 정수 문자열 (예: 1,920,500,000,000 → '19205')."""
    return f"{int(round(won / EOK))}"


def _won_to_jo(won: float) -> str:
    """원 → '조' 문자열. 정수면 소수점 생략 (21.0→'21', 9.4→'9.4')."""
    jo = won / JO
    rounded = round(jo, 1)
    if rounded == int(rounded):
        return f"{int(rounded)}"
    return f"{rounded:.1f}"


def _pct(value: float) -> str:
    return f"{value:+.1f}%"


def _won(value: float) -> str:
    return f"{int(round(value)):,}원"


def format_pick(pick: Pick) -> str:
    lines = [
        f"{pick.name} {pick.ticker} ({pick.market})",
        f"섹터: {pick.sector}",
        f"종가 {_won(pick.close)}  등락률 {_pct(pick.change_pct)}",
        f"거래대금 {_won_to_eok(pick.value_won)}억  시총 {_won_to_jo(pick.marketcap_won)}조",
    ]

    trend = ""
    if pick.return_3m_pct is not None:
        trend = f"3개월 수익률: {_pct(pick.return_3m_pct)}"
    if pick.big_value_days_3m is not None:
        seg = f"🔴 3개월 거래대금 2천억↑: {pick.big_value_days_3m}회"
        trend = f"{trend}  |  {seg}" if trend else seg
    if trend:
        lines.append(trend)

    if pick.news_headline:
        lines.append(f"📰 {pick.news_headline}")

    if pick.companions:
        comp = ", ".join(f"{c.name}({_pct(c.change_pct)})" for c in pick.companions)
        lines.append(f"동일섹터 동반상승: {comp}")

    return "\n".join(lines)


def format_highlights(hl: Highlights, cfg: ScreenConfig) -> str:
    bullets: list[str] = []

    for theme in hl.themes:
        names = ", ".join(theme.filtered_names)
        bullets.append(
            f"• {theme.sector} 테마 집중 — 필터 통과 종목 "
            f"{len(theme.filtered_names)}개 ({names})"
        )

    for s in hl.broad_sectors:
        if s.companion_count <= 0:
            continue
        bullets.append(
            f"• {s.sector} 섹터 전반 강세 — {s.representative} 외 "
            f"{s.companion_count}개 종목 동반 상승"
        )

    if hl.limit_ups:
        names = ", ".join(f"{p.name}({_pct(p.change_pct)})" for p in hl.limit_ups)
        bullets.append(f"• 상한가 종목: {names}")

    if hl.persistent_surges:
        items = ", ".join(
            f"{p.name} (3개월 {p.return_3m_pct:+.0f}%, 오늘 {_pct(p.change_pct)})"
            for p in hl.persistent_surges
        )
        bullets.append(f"• 단기 급등 지속 종목 (추격 주의): {items}")

    if not bullets:
        return ""
    return "📌 오늘의 주목 포인트\n\n" + "\n".join(bullets)


def format_message(date: str, picks: list[Pick], hl: Highlights, cfg: ScreenConfig) -> str:
    """전체 메시지 텍스트 생성."""
    value_eok = int(round(cfg.value_threshold_won / EOK))
    if value_eok >= 10000 and value_eok % 10000 == 0:
        value_label = f"{value_eok // 10000}조"
    elif value_eok >= 1000 and value_eok % 1000 == 0:
        value_label = f"{value_eok // 1000}천억"
    else:
        value_label = f"{value_eok:,}억"
    cond = (
        f"조건: 거래대금 {value_label}↑ | "
        f"등락률 +{cfg.change_threshold_pct:.0f}%↑ | "
        f"시총 {int(cfg.marketcap_max_won // JO)}조↓"
    )

    parts = [
        f"📈 KOSPI/KOSDAQ 급등 스크리닝 ({date})",
        cond,
        f"총 {len(picks)}개 종목 해당",
    ]
    header = "\n".join(parts)

    if not picks:
        return header + "\n\n조건을 만족하는 종목이 없습니다."

    body = "\n\n".join(format_pick(p) for p in picks)
    footer = format_highlights(hl, cfg)

    sections = [header, body]
    if footer:
        sections.append(DIVIDER + "\n\n" + footer)
    return "\n\n".join(sections)


def split_for_telegram(text: str, limit: int = 4000) -> list[str]:
    """텔레그램 4096자 제한에 맞춰 빈 줄 경계로 분할한다."""
    if len(text) <= limit:
        return [text]
    chunks: list[str] = []
    current = ""
    for block in text.split("\n\n"):
        candidate = block if not current else current + "\n\n" + block
        if len(candidate) > limit and current:
            chunks.append(current)
            current = block
        else:
            current = candidate
    if current:
        chunks.append(current)
    return chunks
