@echo off
REM 포트폴리오 대시보드 실행 스크립트 (Windows)
cd /d "%~dp0"

echo [1/2] 시세 받아오는 중... (prices.json 생성)
python fetch_prices.py

echo [2/2] 대시보드 주소: http://localhost:8000/portfolio_dashboard.html
start "" "http://localhost:8000/portfolio_dashboard.html"

echo 웹서버 시작 (종료: Ctrl+C)
python -m http.server 8000
