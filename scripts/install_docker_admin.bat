@echo off
set SCRIPT_DIR=%~dp0
powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process PowerShell -Verb RunAs -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%SCRIPT_DIR%setup_local_docker_windows.ps1"" -Install'"
echo If a Windows permission prompt appears, choose Yes.
echo After installation, restart Windows if requested.
pause
