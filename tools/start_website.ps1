$ErrorActionPreference = "Continue"
$env:PYTHONWARNINGS = "ignore::DeprecationWarning"

$projectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$appDir = Join-Path $projectRoot "app"
$python = "C:\Users\25896\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$log = Join-Path $projectRoot "server_8088.log"

Set-Location $appDir
& $python server.py *>> $log
