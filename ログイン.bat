@echo off
set "PATH=%PATH%;%AppData%\npm"
echo.
echo ==============================================
echo   Google login for clasp
echo ==============================================
echo.
echo A browser will open automatically.
echo Log in with matsu4061209@gmail.com and click Allow.
echo.
call clasp login
echo.
echo ----------------------------------------------
echo   If you see "Authorization successful", done.
echo   You can close this window.
echo ----------------------------------------------
pause
