$ErrorActionPreference = "Stop"

$Port = 8088
$Url = "http://127.0.0.1:$Port"
$Python = "C:\Users\25896\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$App = Join-Path $PSScriptRoot "app\server.py"

if (-not (Test-Path $Python)) {
  Write-Host "Cannot find Codex Python runtime:"
  Write-Host $Python
  exit 1
}

$Existing = Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue
if ($Existing) {
  Write-Host "The local website is already running:"
  Write-Host $Url
  Start-Process $Url
  exit 0
}

Write-Host "Starting RSIN document generator..."
Write-Host "Open this address in your browser:"
Write-Host $Url
Write-Host ""
Write-Host "Keep this window open while testing. Closing it stops the website."
Start-Process $Url

& $Python $App $Port
