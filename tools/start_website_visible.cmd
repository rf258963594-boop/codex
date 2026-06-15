@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\start.ps1"
echo.
echo Website stopped. Press any key to close.
pause >nul
