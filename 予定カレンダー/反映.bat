@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "push_calendar.ps1"
echo.
pause
