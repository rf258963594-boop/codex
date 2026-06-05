@echo off
set PYTHONWARNINGS=ignore::DeprecationWarning
cd /d "%~dp0.."
"C:\Users\25896\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" "C:\Users\25896\Documents\Codex\2026-05-25\new-chat\app\server.py" >> "C:\Users\25896\Documents\Codex\2026-05-25\new-chat\server_8088.log" 2>&1
