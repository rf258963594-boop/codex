@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0website_control.ps1" start
pause
