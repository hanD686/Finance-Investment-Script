@echo off
setlocal

echo Allowing inbound TCP port 8000 for Eastmoney Daily Service...
netsh advfirewall firewall add rule name="Eastmoney Daily Service 8000" dir=in action=allow protocol=TCP localport=8000

echo.
echo Done. If Windows asked for administrator permission, approve it.
echo Phone URL should look like: http://YOUR-PC-IP:8000/
pause
