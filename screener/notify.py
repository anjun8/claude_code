"""텔레그램 전송(외부 의존성 없이 urllib 사용)."""
from __future__ import annotations

import json
import urllib.parse
import urllib.request

from .config import TelegramConfig

API = "https://api.telegram.org/bot{token}/sendMessage"


def send_message(cfg: TelegramConfig, text: str, timeout: int = 15) -> bool:
    """단일 메시지를 전송한다. 성공 시 True."""
    if not cfg.enabled:
        return False
    data = urllib.parse.urlencode(
        {
            "chat_id": cfg.chat_id,
            "text": text,
            "disable_web_page_preview": "true",
        }
    ).encode("utf-8")
    req = urllib.request.Request(API.format(token=cfg.bot_token), data=data)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            body = json.loads(resp.read().decode("utf-8"))
        return bool(body.get("ok"))
    except Exception as exc:  # noqa: BLE001
        print(f"[telegram] 전송 실패: {exc}")
        return False


def send_chunks(cfg: TelegramConfig, chunks: list[str]) -> bool:
    ok = True
    for chunk in chunks:
        ok = send_message(cfg, chunk) and ok
    return ok
