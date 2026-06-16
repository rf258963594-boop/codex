$ErrorActionPreference = "Continue"

function Write-Check {
  param(
    [string]$Name,
    [bool]$Ok,
    [string]$Detail
  )
  $status = if ($Ok) { "OK" } else { "CHECK" }
  Write-Host "[$status] $Name - $Detail"
}

. (Join-Path $PSScriptRoot "set_utf8_console.ps1") | Out-Null

$docker = Get-Command docker -ErrorAction SilentlyContinue
Write-Check "Docker CLI" ([bool]$docker) ($(if ($docker) { $docker.Source } else { "not found; install Docker Desktop" }))

if ($docker) {
  try {
    $dockerVersion = docker --version
    Write-Check "Docker version" $true $dockerVersion
  } catch {
    Write-Check "Docker version" $false $_.Exception.Message
  }
  try {
    $composeVersion = docker compose version
    Write-Check "Docker Compose" $true $composeVersion
  } catch {
    Write-Check "Docker Compose" $false $_.Exception.Message
  }
}

$py = Get-Command py -ErrorAction SilentlyContinue
$python = Get-Command python -ErrorAction SilentlyContinue
Write-Check "Python launcher" ([bool]$py) ($(if ($py) { $py.Source } else { "py.exe not found" }))
Write-Check "Python command" ([bool]$python -and $python.Source -notlike "*WindowsApps*") ($(if ($python) { $python.Source } else { "python.exe not found" }))

$sofficeCandidates = @(
  "C:\Program Files\LibreOffice\program\soffice.exe",
  "C:\Program Files\LibreOffice\program\soffice.com",
  "C:\Program Files (x86)\LibreOffice\program\soffice.exe",
  "C:\Program Files (x86)\LibreOffice\program\soffice.com",
  "D:\Program Files\LibreOffice\program\soffice.exe",
  "D:\Program Files\LibreOffice\program\soffice.com",
  "D:\Program Files\program\soffice.com"
)
$soffice = $sofficeCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
Write-Check "LibreOffice" ([bool]$soffice) ($(if ($soffice) { $soffice } else { "not found; install LibreOffice or set SOFFICE_PATH" }))

$outputEncodingName = $OutputEncoding.WebName
$consoleEncodingName = [Console]::OutputEncoding.WebName
Write-Check "PowerShell OutputEncoding" ($outputEncodingName -eq "utf-8") $outputEncodingName
Write-Check "Console OutputEncoding" ($consoleEncodingName -eq "utf-8") $consoleEncodingName
Write-Check "PYTHONUTF8" ($env:PYTHONUTF8 -eq "1") ($(if ($env:PYTHONUTF8) { $env:PYTHONUTF8 } else { "not set" }))
Write-Check "PYTHONIOENCODING" ($env:PYTHONIOENCODING -eq "utf-8") ($(if ($env:PYTHONIOENCODING) { $env:PYTHONIOENCODING } else { "not set" }))

try {
  $wslStatus = wsl --status 2>&1
  $wslOk = $LASTEXITCODE -eq 0 -and -not (($wslStatus -join " ") -match "wsl\.exe --install")
  Write-Check "WSL" $wslOk ($(if ($wslOk) { ($wslStatus | Select-Object -First 1) -join " " } else { "not installed or not available; run wsl --install before Docker Desktop" }))
} catch {
  Write-Check "WSL" $false "not installed or not available"
}
