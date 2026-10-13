@echo off
rem Installs and starts the Claude AI add-in for Excel on Windows.
rem Needs only Node.js - no npm install, about 1 MB of disk space.
rem Run this file again whenever you want to use the add-in.
cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
  echo Node.js is not installed. Install the LTS version from https://nodejs.org and run this file again.
  pause
  exit /b 1
)

tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>nul | find /I "EXCEL.EXE" >nul
if not errorlevel 1 (
  echo Excel is open. Please SAVE your work and CLOSE Excel, then press any key.
  pause >nul
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0windows\install.ps1"
if errorlevel 1 (
  echo.
  echo Setup failed. Scroll up to see the error.
  pause
  exit /b 1
)

echo Opening Excel...
start "" excel.exe
node "%~dp0windows\server.js" "%LOCALAPPDATA%\ClaudeExcelAddin\localhost.pfx" "claude-excel-localhost"
pause
