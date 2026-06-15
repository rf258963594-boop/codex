$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

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

$DefaultPort = if ($env:APP_PORT) { [int]$env:APP_PORT } else { 8088 }
$Port = if ($args.Count -ge 1) { [int]$args[0] } else { $DefaultPort }
$Url = "http://127.0.0.1:$Port"
$App = Join-Path $ProjectRoot "app\server.py"

function Resolve-PythonCommand {
    $candidates = @()
    if ($env:PYTHON_EXE) {
        $candidates += $env:PYTHON_EXE
    }
    if ($env:PYTHON_LAUNCHER_DIR) {
        $candidates += (Join-Path $env:PYTHON_LAUNCHER_DIR "py.exe")
    }
    $candidates += "C:\Users\Administrator\AppData\Local\Programs\Python\Launcher\py.exe"

    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) { $candidates += $py.Source }

    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python -and $python.Source -notlike "*WindowsApps*") {
        $candidates += $python.Source
    }

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path $candidate)) {
            return $candidate
        }
    }
    throw "Cannot find Python. Set PYTHON_EXE to python.exe or PYTHON_LAUNCHER_DIR to the folder containing py.exe."
}

$Python = Resolve-PythonCommand
$PythonArgs = @()
if ((Split-Path $Python -Leaf).ToLowerInvariant() -eq "py.exe") {
    $PythonArgs += "-3"
}
$PythonArgs += $App
$PythonArgs += "$Port"

Set-Location $ProjectRoot
$env:PYTHONWARNINGS = "ignore::DeprecationWarning"

$Existing = Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue
if ($Existing) {
    Write-Host "The website is already running:"
    Write-Host $Url
    Start-Process $Url -ErrorAction SilentlyContinue
    exit 0
}

Write-Host "Starting RSIN document generator..."
Write-Host "Project root: $ProjectRoot"
Write-Host "Python: $Python"
Write-Host "URL: $Url"
Write-Host ""
Write-Host "Keep this window open while testing. Closing it stops the website."
Start-Process $Url -ErrorAction SilentlyContinue

& $Python @PythonArgs
