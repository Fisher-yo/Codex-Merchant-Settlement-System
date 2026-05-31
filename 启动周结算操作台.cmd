@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0run-settlement-console.ps1"
if errorlevel 1 (
  echo.
  echo Settlement console failed to start.
  echo Please check startup-error-log.txt.
  echo.
  pause
)
