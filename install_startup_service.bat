@echo off
setlocal
set "TASK_NAME=Eastmoney Daily Service"
set "SERVICE_BAT=%~dp0start_eastmoney_service.bat"

echo Installing Windows startup task: %TASK_NAME%
schtasks /Create /TN "%TASK_NAME%" /SC ONLOGON /TR "\"%SERVICE_BAT%\"" /F

echo.
echo Done. The service will start automatically after you log in to Windows.
echo To remove it later, run:
echo schtasks /Delete /TN "%TASK_NAME%" /F
pause
