@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo Creating virtual environment...
  py -m venv .venv
)

echo Installing or checking dependencies...
".venv\Scripts\python.exe" -m pip install -r requirements.txt

echo.
echo Starting Eastmoney daily service...
echo Keep this window open. Close it to stop the service.
echo.
".venv\Scripts\python.exe" eastmoney_daily.py service

pause
