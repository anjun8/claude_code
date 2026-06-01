"""스크리너 실행 엔트리포인트.

사용 예:
    python -m screener.cli                 # 직전 영업일, 텔레그램 전송
    python -m screener.cli --date 2026-04-08
    python -m screener.cli --dry-run       # 전송 없이 콘솔 출력
    python -m screener.cli --no-trend --no-news   # 3개월/뉴스 생략(빠름)
"""
from __future__ import annotations

import argparse
import sys

from .config import ScreenConfig, TelegramConfig
from .formatter import format_message, split_for_telegram
from .news import fetch_headline
from .screen import analyze, build_picks, enrich_trends, filter_universe


def run(args: argparse.Namespace) -> int:
    cfg = ScreenConfig.from_env()

    if args.demo:
        # 네트워크/자격증명 없이 합성 데이터로 전체 파이프라인 확인
        from . import demo

        date = "20260408"
        pretty_date = "2026-04-08"
        print("[screener] --demo: 합성 데이터로 실행 (네트워크 미사용)")
        universe = demo.demo_universe()
        trend_provider = demo.demo_trend_provider
        news_fn = demo.demo_news
    else:
        # datasource는 pykrx 의존이므로 실제 실행 시점에만 import
        from . import datasource

        date = datasource.resolve_date(args.date)
        pretty_date = f"{date[:4]}-{date[4:6]}-{date[6:]}"
        print(f"[screener] 대상 영업일: {pretty_date}")
        universe = datasource.fetch_universe(date)
        trend_provider = datasource.make_trend_provider(date, cfg)
        news_fn = fetch_headline

    print(f"[screener] universe 종목 수: {len(universe)}")

    filtered = filter_universe(universe, cfg)
    print(f"[screener] 필터 통과: {len(filtered)}개")

    picks = build_picks(universe, filtered, cfg)

    if not args.no_trend and picks:
        enrich_trends(picks, trend_provider)

    if not args.no_news and picks:
        for p in picks:
            p.news_headline = news_fn(p.name)

    hl = analyze(universe, picks, cfg)
    message = format_message(pretty_date, picks, hl, cfg)

    if args.dry_run:
        print("\n" + message)
        return 0

    tg = TelegramConfig.from_env()
    if not tg.enabled:
        print("[screener] TELEGRAM_BOT_TOKEN/CHAT_ID 미설정 — 콘솔 출력으로 대체")
        print("\n" + message)
        return 0

    from .notify import send_chunks

    ok = send_chunks(tg, split_for_telegram(message))
    print("[screener] 텔레그램 전송:", "성공" if ok else "실패")
    return 0 if ok else 1


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="KOSPI/KOSDAQ 급등 거래대금 스크리너")
    parser.add_argument("--date", help="대상일 YYYY-MM-DD (기본: 직전 영업일)")
    parser.add_argument(
        "--demo",
        action="store_true",
        help="네트워크/자격증명 없이 합성 데이터로 실행 (스모크 테스트)",
    )
    parser.add_argument("--dry-run", action="store_true", help="전송 없이 콘솔 출력")
    parser.add_argument("--no-trend", action="store_true", help="3개월 추세 계산 생략")
    parser.add_argument("--no-news", action="store_true", help="뉴스 헤드라인 생략")
    args = parser.parse_args(argv)
    try:
        return run(args)
    except Exception as exc:  # noqa: BLE001
        print(f"[screener] 오류: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
