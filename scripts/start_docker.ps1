$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $ProjectRoot

if (-not (Test-Path ".env")) {
    Copy-Item ".env.example" ".env"
    Write-Host "Created .env from .env.example."
    Write-Host "Please edit DEFAULT_ADMIN_PASSWORD before first production start."
    exit 1
}

docker compose up -d --build
docker compose ps
