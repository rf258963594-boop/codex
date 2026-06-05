param(
  [ValidateSet("start", "stop", "restart", "status")]
  [string]$Action = "status"
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$Python = "C:\Users\25896\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$App = Join-Path $ProjectRoot "app\server.py"
$HiddenRunner = Join-Path $ProjectRoot "tools\run_hidden.vbs"
$DataDir = Join-Path $ProjectRoot "app\data"
$PidFile = Join-Path $DataDir "server_8088.pid"
$OutLog = Join-Path $ProjectRoot "server_8088.out.log"
$ErrLog = Join-Path $ProjectRoot "server_8088.err.log"
$Port = 8088
$Url = "http://127.0.0.1:$Port/"

function Get-WebsiteProcess {
  if (Test-Path $PidFile) {
    $pidText = (Get-Content $PidFile -ErrorAction SilentlyContinue | Select-Object -First 1)
    if ($pidText -match '^\d+$') {
      return Get-Process -Id ([int]$pidText) -ErrorAction SilentlyContinue
    }
  }
  return $null
}

function Test-WebsiteHttp {
  try {
    $response = Invoke-WebRequest -UseBasicParsing -Uri ($Url + "login") -TimeoutSec 3
    return $response.StatusCode -eq 200
  } catch {
    return $false
  }
}

function Start-Website {
  New-Item -ItemType Directory -Force -Path $DataDir | Out-Null
  if (-not (Test-Path $Python)) {
    Write-Host "Cannot find Python runtime:"
    Write-Host $Python
    exit 1
  }
  if (Test-WebsiteHttp) {
    Write-Host "Website is already running: $Url"
    Start-Process $Url
    return
  }
  $psi = [System.Diagnostics.ProcessStartInfo]::new()
  $psi.FileName = (Join-Path $env:WINDIR "System32\wscript.exe")
  $psi.Arguments = "`"$HiddenRunner`" `"$Python`" `"$App`" $Port `"$OutLog`" `"$ErrLog`""
  $psi.WorkingDirectory = $ProjectRoot
  $psi.UseShellExecute = $true
  $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
  $proc = [System.Diagnostics.Process]::Start($psi)
  Start-Sleep -Seconds 2
  $conn = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($conn) {
    Set-Content -Path $PidFile -Value $conn.OwningProcess
  } else {
    Set-Content -Path $PidFile -Value $proc.Id
  }
  if (Test-WebsiteHttp) {
    Write-Host "Website started: $Url"
    Start-Process $Url
  } else {
    Write-Host "Website did not respond yet. Check logs:"
    Write-Host $OutLog
    Write-Host $ErrLog
  }
}

function Stop-Website {
  $proc = Get-WebsiteProcess
  if ($proc) {
    Stop-Process -Id $proc.Id -Force
    Remove-Item -Path $PidFile -Force -ErrorAction SilentlyContinue
    Write-Host "Website stopped."
    return
  }
  $conn = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($conn) {
    Stop-Process -Id $conn.OwningProcess -Force
    Remove-Item -Path $PidFile -Force -ErrorAction SilentlyContinue
    Write-Host "Website stopped by port owner."
    return
  }
  Write-Host "Website is not running."
}

function Show-Status {
  if (Test-WebsiteHttp) {
    Write-Host "Website is running: $Url"
    return
  }
  Write-Host "Website is not running."
}

switch ($Action) {
  "start" { Start-Website }
  "stop" { Stop-Website }
  "restart" { Stop-Website; Start-Sleep -Seconds 1; Start-Website }
  "status" { Show-Status }
}
