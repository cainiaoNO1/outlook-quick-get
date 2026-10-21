@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

where node >nul 2>nul
if not errorlevel 1 goto use_node

where python >nul 2>nul
if not errorlevel 1 goto use_python

echo.
echo  [ERROR] Neither node nor python found on this machine.
echo          Install Node.js: https://nodejs.org
echo.
pause
endlocal & exit /b 1

:use_node
node "%~dp0outlook_mail.js" %*
set "RC=%ERRORLEVEL%"
goto after

:use_python
python "%~dp0outlook_mail.py" %*
set "RC=%ERRORLEVEL%"
goto after

:after
if "%~1"=="" (
  echo.
  pause
)
endlocal & exit /b %RC%
