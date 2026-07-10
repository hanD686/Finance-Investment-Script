@echo off
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" eastmoney_daily.py update
) else (
  py eastmoney_daily.py update
)
pause
