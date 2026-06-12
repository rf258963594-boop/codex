$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $ProjectRoot

$BackupDir = if ($env:BACKUP_DIR) { $env:BACKUP_DIR } else { ".\backups" }
$Stamp = Get-Date -Format "yyyyMMdd-HHmmss"
New-Item -ItemType Directory -Force -Path $BackupDir | Out-Null

$BackupPath = Join-Path $BackupDir "secretary-files-runtime-$Stamp.tar.gz"
docker compose exec -T secretary-files sh -c "tar -czf /tmp/secretary-files-runtime.tar.gz -C /app/app data uploads generated"
docker cp "secretary-files:/tmp/secretary-files-runtime.tar.gz" $BackupPath
docker compose exec -T secretary-files rm -f /tmp/secretary-files-runtime.tar.gz

Write-Host "Backup created: $BackupPath"
