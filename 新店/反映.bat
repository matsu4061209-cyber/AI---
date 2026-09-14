@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "push_gas.ps1"
echo.
pause
