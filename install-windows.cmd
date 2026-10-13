@echo off
rem One-click setup for Excel on Windows: installs dependencies, trusts the
rem localhost certificate, starts the add-in server and sideloads it into Excel.
cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
  echo Node.js is not installed. Install the LTS version from https://nodejs.org and run this file again.
  pause
  exit /b 1
)

echo Closing Excel so the add-in can be registered...
taskkill /IM EXCEL.EXE /F >nul 2>nul

echo Installing dependencies (first run takes a minute)...
call npm install
if errorlevel 1 goto :fail

echo Installing the localhost HTTPS certificate (click Yes if Windows asks)...
call npx office-addin-dev-certs install
if errorlevel 1 goto :fail

echo Starting the add-in and opening Excel. Keep this window open while you use the add-in.
call npm start
goto :eof

:fail
echo Setup failed. Scroll up to see the error.
pause
exit /b 1
