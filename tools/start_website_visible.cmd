@echo off
setlocal
set PYTHONWARNINGS=ignore::DeprecationWarning
cd /d "%~dp0.."
echo Starting Secretary File Generator at http://127.0.0.1:8088/
echo Keep this window open while using the website.
start "" powershell -NoProfile -WindowStyle Hidden -Command "Start-Sleep -Seconds 2; Start-Process 'http://127.0.0.1:8088/'"
"C:\Users\25896\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" "app\server.py" 8088
echo.
echo Website stopped. Press any key to close.
pause >nul
