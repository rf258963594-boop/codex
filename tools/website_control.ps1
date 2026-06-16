param(
  [ValidateSet("start", "stop", "restart", "status")]
  [string]$Action = "status"
)

$ErrorActionPreference = "Stop"

function Set-Utf8Process {
  [Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
  [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
  $global:OutputEncoding = [System.Text.UTF8Encoding]::new($false)
  $env:PYTHONUTF8 = "1"
  $env:PYTHONIOENCODING = "utf-8"
}

Set-Utf8Process

$ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)

function Import-DotEnv {
  param([string]$Path)
  if (-not (Test-Path $Path)) {
    return
  }
  Get-Content $Path | ForEach-Object {
    $line = $_.Trim()
    if ($line -and -not $line.StartsWith("#") -and $line -match "=") {
      $parts = $line.Split("=", 2)
      $key = $parts[0].Trim()
      $value = $parts[1].Trim().Trim('"').Trim("'")
      if ($key -and -not [Environment]::GetEnvironmentVariable($key, "Process")) {
        [Environment]::SetEnvironmentVariable($key, $value, "Process")
      }
    }
  }
}

Import-DotEnv (Join-Path $ProjectRoot ".env")

$App = Join-Path $ProjectRoot "app\server.py"
$HiddenRunner = Join-Path $ProjectRoot "tools\run_hidden.vbs"
$DataDir = if ($env:DATA_DIR) { $env:DATA_DIR } else { Join-Path $ProjectRoot "app\data" }
$PidFile = Join-Path $DataDir "server_8088.pid"
$OutLog = Join-Path $ProjectRoot "server_8088.out.log"
$ErrLog = Join-Path $ProjectRoot "server_8088.err.log"
$Port = if ($env:APP_PORT) { [int]$env:APP_PORT } else { 8088 }
$Url = "http://127.0.0.1:$Port/"

function Resolve-PythonCommand {
  $candidates = @()
  if ($env:PYTHON_EXE) { $candidates += $env:PYTHON_EXE }
  if ($env:PYTHON_LAUNCHER_DIR) { $candidates += (Join-Path $env:PYTHON_LAUNCHER_DIR "py.exe") }
  $candidates += "C:\Users\Administrator\AppData\Local\Programs\Python\Launcher\py.exe"
  $py = Get-Command py -ErrorAction SilentlyContinue
  if ($py) { $candidates += $py.Source }
  $python = Get-Command python -ErrorAction SilentlyContinue
  if ($python -and $python.Source -notlike "*WindowsApps*") { $candidates += $python.Source }
  foreach ($candidate in $candidates) {
    if ($candidate -and (Test-Path $candidate)) { return $candidate }
  }
  throw "Cannot find Python. Set PYTHON_EXE or PYTHON_LAUNCHER_DIR."
}

function Python-Arguments {
  param([string]$PythonPath)
  $args = @()
  if ((Split-Path $PythonPath -Leaf).ToLowerInvariant() -eq "py.exe") {
    $args += "-3"
  }
  $args += $App
  $args += "$Port"
  return $args
}

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
  $Python = Resolve-PythonCommand
  $PythonArgs = Python-Arguments $Python
  if (Test-WebsiteHttp) {
    Write-Host "Website is already running: $Url"
    Start-Process $Url -ErrorAction SilentlyContinue
    return
  }
  $psi = [System.Diagnostics.ProcessStartInfo]::new()
  $psi.FileName = (Join-Path $env:WINDIR "System32\wscript.exe")
  $allArgs = ($PythonArgs | ForEach-Object { '"' + ($_ -replace '"', '\"') + '"' }) -join " "
  $psi.Arguments = "`"$HiddenRunner`" `"$Python`" `"$allArgs`" `"$OutLog`" `"$ErrLog`""
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
    Start-Process $Url -ErrorAction SilentlyContinue
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
