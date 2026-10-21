@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo.
echo  ============================================================
echo    Watch mode: check for new mail every 15 seconds
echo    Stops automatically once a verification code is found.
echo    Close this window or press Ctrl+C to stop earlier.
echo  ============================================================
echo.

where node >nul 2>nul
if not errorlevel 1 goto use_node
where python >nul 2>nul
if not errorlevel 1 goto use_python
echo  [ERROR] Neither node nor python found. Install Node.js: https://nodejs.org
pause
endlocal & exit /b 1

:use_node
node "%~dp0outlook_mail.js" -w 15 --exit-on-code
goto after

:use_python
python "%~dp0outlook_mail.py" -w 15 --exit-on-code
goto after

:after
echo.
pause
endlocal
