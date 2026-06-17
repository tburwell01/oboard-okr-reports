@echo off
REM One-time setup for a teammate: creates config.json (if missing) and a Desktop shortcut.
setlocal
set "HERE=%~dp0"

if not exist "%HERE%config.json" (
  copy "%HERE%config.example.json" "%HERE%config.json" >nul
  echo Created config.json from the example.
  echo   -^> Open it and paste the shared Oboard API token. Leave outputFolder as AUTO.
) else (
  echo config.json already exists - leaving it as is.
)

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ws=New-Object -ComObject WScript.Shell; $lnk=$ws.CreateShortcut((Join-Path ([Environment]::GetFolderPath('Desktop')) 'Corporate OKR Board.lnk')); $lnk.TargetPath=(Join-Path '%HERE%' 'Run-OKR-WebApp.bat'); $lnk.WorkingDirectory='%HERE%'; $lnk.WindowStyle=7; $lnk.Description='Generate the FY27 Corporate OKR Board'; $ppt='C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE'; if(Test-Path $ppt){$lnk.IconLocation=\"$ppt,0\"}; $lnk.Save(); Write-Host 'Desktop shortcut created: Corporate OKR Board'"

echo.
echo Setup complete. After adding your token to config.json, double-click the
echo "Corporate OKR Board" shortcut on your Desktop to start.
pause
