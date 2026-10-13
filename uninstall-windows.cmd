@echo off
rem Removes the Claude AI add-in and its localhost certificate from this PC.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0windows\uninstall.ps1"
pause
