#!/usr/bin/env bash
# 포트폴리오 대시보드 실행 스크립트 (Mac / Linux)
# 1) 시세 받아오기 → prices.json 생성  2) 로컬 웹서버로 대시보드 띄우기
set -e
cd "$(dirname "$0")"

echo "▶ 시세 받아오는 중... (prices.json 생성)"
python3 fetch_prices.py || echo "  (시세 받기 실패 — 인터넷/CSV 확인. 대시보드는 그대로 열립니다)"

PORT=8000
URL="http://localhost:${PORT}/portfolio_dashboard.html"
echo "▶ 대시보드 주소: ${URL}"

# 브라우저 자동 열기 (가능한 경우)
( sleep 1; (command -v open >/dev/null && open "$URL") || (command -v xdg-open >/dev/null && xdg-open "$URL") ) >/dev/null 2>&1 &

echo "▶ 웹서버 시작 (종료: Ctrl+C)"
python3 -m http.server "$PORT"
