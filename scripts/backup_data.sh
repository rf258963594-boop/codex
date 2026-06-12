#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")/.."

BACKUP_DIR="${BACKUP_DIR:-./backups}"
STAMP="$(date +%Y%m%d-%H%M%S)"
mkdir -p "$BACKUP_DIR"

docker compose exec -T secretary-files tar -czf - -C /app/app data uploads generated > "$BACKUP_DIR/secretary-files-runtime-$STAMP.tar.gz"

echo "Backup created: $BACKUP_DIR/secretary-files-runtime-$STAMP.tar.gz"
