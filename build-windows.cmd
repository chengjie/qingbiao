@echo off
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0build-windows.ps1"
if errorlevel 1 (
  echo Build failed. See the message above.
) else (
  echo Build completed. Open dist\QingBiao.exe
)
pause
