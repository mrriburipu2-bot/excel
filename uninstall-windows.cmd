@echo off
rem Stops the add-in server and removes the sideloaded add-in from Excel.
cd /d "%~dp0"
call npm stop
pause
