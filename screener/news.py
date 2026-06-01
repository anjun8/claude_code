"""종목별 뉴스 헤드라인(선택 기능).

best-effort로 동작한다. 실패하거나 키가 없으면 None을 반환해 메시지에서
📰 줄을 생략한다. 네트워크가 막힌 환경에서도 전체 파이프라인이 죽지 않도록 한다.

지원 소스:
    1) 네이버 검색 OpenAPI (NAVER_CLIENT_ID / NAVER_CLIENT_SECRET) — 권장
       (없으면) None
"""
from __future__ import annotations

import json
import os
import re
import urllib.parse
import urllib.request

_TAG_RE = re.compile(r"<[^>]+>")


def _strip(text: str) -> str:
    return _TAG_RE.sub("", text).replace("&quot;", '"').replace("&amp;", "&").strip()


def naver_headline(query: str, timeout: int = 6) -> str | None:
    cid = os.environ.get("NAVER_CLIENT_ID")
    secret = os.environ.get("NAVER_CLIENT_SECRET")
    if not (cid and secret):
        return None
    url = (
        "https://openapi.naver.com/v1/search/news.json?display=1&sort=sim&query="
        + urllib.parse.quote(query)
    )
    req = urllib.request.Request(url)
    req.add_header("X-Naver-Client-Id", cid)
    req.add_header("X-Naver-Client-Secret", secret)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        items = data.get("items") or []
        if not items:
            return None
        return _strip(items[0].get("title", "")) or None
    except Exception:
        return None


def fetch_headline(stock_name: str) -> str | None:
    """종목명으로 대표 헤드라인 1건을 가져온다(없으면 None)."""
    return naver_headline(stock_name)
