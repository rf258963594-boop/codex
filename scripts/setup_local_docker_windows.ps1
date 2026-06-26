param(
  [switch]$Install,
  [switch]$Start,
  [switch]$StopLocalPython,
  [int]$Port = 8088
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Resolve-Path (Join-Path $ScriptDir "..")
Set-Location $ProjectRoot

function Write-Step($Message) {
  Write-Host ""
  Write-Host "== $Message =="
}

function Test-Admin {
  $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
  $principal = New-Object Security.Principal.WindowsPrincipal($identity)
  return $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

function Require-Admin($Action) {
  if (-not (Test-Admin)) {
    throw "$Action requires Administrator PowerShell. Right-click PowerShell and choose 'Run as administrator', then run this script again."
  }
}

function Ensure-LocalEnv {
  $envPath = Join-Path $ProjectRoot ".env"
  if (Test-Path $envPath) {
    Write-Host ".env already exists; keeping it."
    return
  }
  @"
APP_PORT=$Port
DEFAULT_ADMIN_USERNAME=admin
DEFAULT_ADMIN_PASSWORD=admin123
PYTHONUTF8=1
PYTHONIOENCODING=utf-8
"@ | Set-Content -Encoding UTF8 $envPath
  Write-Host "Created local .env for Docker."
}

function Enable-DockerPrerequisites {
  Require-Admin "Enabling WSL2 / virtual machine platform"
  Write-Step "Enable Windows features"
  $features = @(
    "Microsoft-Windows-Subsystem-Linux",
    "VirtualMachinePlatform"
  )
  foreach ($feature in $features) {
    $state = (Get-WindowsOptionalFeature -Online -FeatureName $feature).State
    Write-Host "${feature}: $state"
    if ($state -ne "Enabled") {
      Enable-WindowsOptionalFeature -Online -FeatureName $feature -All -NoRestart | Out-Null
      Write-Host "Enabled $feature. A Windows restart may be required."
    }
  }

  Write-Step "Install or update WSL"
  try {
    wsl --install --no-distribution
  } catch {
    Write-Host "WSL install command returned a non-blocking error. Continue if WSL is already installed."
  }
  try {
    wsl --set-default-version 2
  } catch {
    Write-Host "Could not set WSL default version yet. If Windows asks for restart, restart first and rerun -Start."
  }
}

function Install-DockerDesktop {
  Require-Admin "Installing Docker Desktop"
  Write-Step "Install Docker Desktop"
  $dockerDesktop = "C:\Program Files\Docker\Docker\Docker Desktop.exe"
  if (Test-Path $dockerDesktop) {
    Write-Host "Docker Desktop already installed."
    return
  }
  if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    throw "winget is not available. Install Docker Desktop manually from https://www.docker.com/products/docker-desktop/"
  }
  winget install -e --id Docker.DockerDesktop --accept-source-agreements --accept-package-agreements
  Write-Host "Docker Desktop installer finished. If Windows asks for restart, restart before running -Start."
}

function Stop-LocalPythonServer {
  Write-Step "Check local Python server on port $Port"
  $conn = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
  if (-not $conn) {
    Write-Host "Port $Port is free."
    return
  }
  $processId = $conn.OwningProcess
  $proc = Get-CimInstance Win32_Process -Filter "ProcessId=$processId" -ErrorAction SilentlyContinue
  $cmd = if ($proc) { $proc.CommandLine } else { "" }
  Write-Host "Port $Port is used by PID $processId"
  if ($cmd -match "app[\\/]+server\.py|server\.py\s+$Port") {
    Stop-Process -Id $processId -Force
    Write-Host "Stopped local Python website process."
    return
  }
  throw "Port $Port is already in use by another process. Close it or change APP_PORT in .env."
}

function Wait-DockerReady {
  Write-Step "Start Docker Desktop"
  $dockerDesktop = "C:\Program Files\Docker\Docker\Docker Desktop.exe"
  if (-not (Get-Command docker -ErrorAction SilentlyContinue)) {
    if (Test-Path $dockerDesktop) {
      Start-Process -FilePath $dockerDesktop | Out-Null
    } else {
      throw "Docker command is not available and Docker Desktop is not installed."
    }
  }
  for ($i = 1; $i -le 60; $i++) {
    try {
      docker info *> $null
      Write-Host "Docker is ready."
      return
    } catch {
      Start-Sleep -Seconds 3
    }
  }
  throw "Docker did not become ready. Open Docker Desktop manually and check WSL2 engine status."
}

function Start-Project {
  Ensure-LocalEnv
  if ($StopLocalPython) {
    Stop-LocalPythonServer
  }
  Wait-DockerReady

  Write-Step "Build and start website container"
  docker compose -f docker-compose.yml -f docker-compose.local.yml up -d --build

  Write-Step "Verify website"
  $url = "http://127.0.0.1:$Port/login"
  for ($i = 1; $i -le 30; $i++) {
    try {
      Invoke-WebRequest -UseBasicParsing -Uri $url -TimeoutSec 3 | Out-Null
      Write-Host "Website is ready: $url"
      return
    } catch {
      Start-Sleep -Seconds 2
    }
  }
  throw "Container started, but website health check failed. Run: docker compose logs --tail=100"
}

if ($Install) {
  Enable-DockerPrerequisites
  Install-DockerDesktop
  Write-Host ""
  Write-Host "Install step complete. Restart Windows if prompted, then run:"
  Write-Host ".\scripts\setup_local_docker_windows.ps1 -Start -StopLocalPython"
  exit 0
}

if ($Start) {
  Start-Project
  exit 0
}

Write-Host "Usage:"
Write-Host "  Admin once: .\scripts\setup_local_docker_windows.ps1 -Install"
Write-Host "  Start site: .\scripts\setup_local_docker_windows.ps1 -Start -StopLocalPython"
