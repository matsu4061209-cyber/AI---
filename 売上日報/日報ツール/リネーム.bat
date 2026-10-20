@echo off
chcp 65001 >nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0rename.ps1"
echo.
echo ----------------------------------------
echo Enter key de tojimasu.
pause >nul
