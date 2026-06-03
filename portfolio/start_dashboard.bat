@echo off
chcp 65001 >nul
cd /d "%~dp0"

REM Pick python command (python or py)
where python >nul 2>nul && (set PY=python) || (set PY=py)

echo [1/2] Fetching prices...  (prices.json)
%PY% fetch_prices.py

echo [2/2] Starting local web server on http://localhost:8000
start "portfolio-server" %PY% -m http.server 8000

REM give the server a moment, then open the dashboard in the browser
timeout /t 2 >nul
start "" "http://localhost:8000/portfolio_dashboard.html"

echo.
echo Dashboard opened in your browser:
echo   http://localhost:8000/portfolio_dashboard.html
echo.
echo To STOP: close the "portfolio-server" window.
pause
