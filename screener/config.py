"""스크리너 설정.

임계값은 환경변수로 덮어쓸 수 있고, 기본값은 예시 메시지의 조건
(거래대금 2천억↑ | 등락률 +10%↑ | 시총 30조↓)과 동일하다.
"""
from __future__ import annotations

import os
from dataclasses import dataclass

# 단위 상수
EOK = 100_000_000          # 1억 원
JO = 1_000_000_000_000     # 1조 원


def _load_dotenv() -> None:
    """`.env`가 있으면 best-effort로 로드한다(python-dotenv 없어도 동작)."""
    try:
        from dotenv import load_dotenv  # type: ignore

        load_dotenv()
        return
    except Exception:
        pass
    path = os.path.join(os.getcwd(), ".env")
    if not os.path.exists(path):
        return
    with open(path, encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def _f(name: str, default: float) -> float:
    raw = os.environ.get(name)
    if raw is None or raw == "":
        return default
    try:
        return float(raw)
    except ValueError:
        return default


@dataclass(frozen=True)
class ScreenConfig:
    """스크리닝 임계값과 표시 옵션."""

    # 필터 조건
    value_threshold_won: float = 2_000 * EOK      # 거래대금 2천억 원
    change_threshold_pct: float = 10.0            # 등락률 +10%
    marketcap_max_won: float = 30 * JO            # 시총 30조 원

    # 3개월 추세 분석
    lookback_days: int = 90                       # "3개월" 윈도우(달력일)
    surge_return_pct: float = 100.0               # 단기 급등 지속(추격 주의) 기준 수익률

    # 섹터 분석
    companion_rise_pct: float = 0.0               # 동반 상승으로 볼 최소 등락률(초과)
    theme_min_count: int = 2                      # "테마 집중" 최소 통과 종목 수
    companion_top_n: int = 4                      # 종목별 동반상승 표시 개수
    limit_up_pct: float = 29.5                    # 상한가로 간주할 등락률

    @classmethod
    def from_env(cls) -> "ScreenConfig":
        _load_dotenv()
        return cls(
            value_threshold_won=_f("SCREEN_VALUE_EOK", 2_000) * EOK,
            change_threshold_pct=_f("SCREEN_CHANGE_PCT", 10.0),
            marketcap_max_won=_f("SCREEN_MARKETCAP_JO", 30) * JO,
            lookback_days=int(_f("SCREEN_LOOKBACK_DAYS", 90)),
            surge_return_pct=_f("SCREEN_SURGE_RETURN_PCT", 100.0),
            companion_rise_pct=_f("SCREEN_COMPANION_RISE_PCT", 0.0),
            theme_min_count=int(_f("SCREEN_THEME_MIN_COUNT", 2)),
            companion_top_n=int(_f("SCREEN_COMPANION_TOP_N", 4)),
            limit_up_pct=_f("SCREEN_LIMIT_UP_PCT", 29.5),
        )


@dataclass(frozen=True)
class TelegramConfig:
    bot_token: str | None = None
    chat_id: str | None = None

    @classmethod
    def from_env(cls) -> "TelegramConfig":
        _load_dotenv()
        return cls(
            bot_token=os.environ.get("TELEGRAM_BOT_TOKEN"),
            chat_id=os.environ.get("TELEGRAM_CHAT_ID"),
        )

    @property
    def enabled(self) -> bool:
        return bool(self.bot_token and self.chat_id)
