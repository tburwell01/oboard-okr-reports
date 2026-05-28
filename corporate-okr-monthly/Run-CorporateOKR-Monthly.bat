@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Invoke-CorporateOKRMonthly.ps1"
set EXIT=%ERRORLEVEL%
if not "%EXIT%"=="0" exit /b %EXIT%
echo.
echo Open the HTML file under output\ when the script prints OK.
endlocal
