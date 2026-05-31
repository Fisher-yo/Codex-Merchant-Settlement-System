@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0create-desktop-shortcut.ps1"
if errorlevel 1 (
  echo.
  echo Failed to create desktop shortcut.
  echo.
  pause
) else (
  echo.
  echo Desktop shortcut created.
  echo.
  pause
)
