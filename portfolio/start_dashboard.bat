@echo off
chcp 65001 >nul
cd /d "%~dp0"

REM Pick python command (python or py)
where python >nul 2>nul && (set PY=python) || (set PY=py)

echo [1/2] Fetching prices...  (prices.json)  [log: fetch_log.txt]
%PY% fetch_prices.py > fetch_log.txt 2>&1
type fetch_log.txt

echo [2/2] Starting local web server on http://localhost:8000
start "portfolio-server" %PY% -m http.server 8000

REM give the server a moment, then open the dashboard in Chrome (fallback: default browser)
timeout /t 2 >nul
set "URL=http://localhost:8000/portfolio_dashboard.html"
set "CHROME="
for %%P in ("%ProgramFiles%\Google\Chrome\Application\chrome.exe" "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" "%LocalAppData%\Google\Chrome\Application\chrome.exe") do if exist "%%~P" set "CHROME=%%~P"
if defined CHROME (
  start "" "%CHROME%" "%URL%"
) else (
  start "" "%URL%"
)

echo.
echo Dashboard opened in your browser:
echo   http://localhost:8000/portfolio_dashboard.html
echo.
echo To STOP: close the "portfolio-server" window.
pause
