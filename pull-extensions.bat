@echo off
rem Seconds to keep this window open after the run (press any key to close sooner).
set CLOSE_SECS=10
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0pull-extensions.ps1"
echo.
timeout /t %CLOSE_SECS%
