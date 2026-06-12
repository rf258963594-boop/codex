#!/usr/bin/env bash
set -euo pipefail

APP_DIR="${APP_DIR:-$(pwd)}"
DEPLOY_BRANCH="${DEPLOY_BRANCH:-main}"
BACKUP_DIR="${BACKUP_DIR:-$APP_DIR/backups}"
SERVICE_NAME="${SERVICE_NAME:-secretary-files}"

cd "$APP_DIR"

if [ ! -d ".git" ]; then
  echo "This folder is not a Git checkout: $APP_DIR"
  echo "Clone the repository here before enabling automatic deployment."
  exit 1
fi

if [ ! -f ".env" ]; then
  cp .env.example .env
  echo "Created .env from .env.example."
  echo "Edit .env and set DEFAULT_ADMIN_PASSWORD before deploying."
  exit 1
fi

if ! command -v docker >/dev/null 2>&1; then
  echo "Docker is not installed."
  exit 1
fi

compose() {
  if docker compose version >/dev/null 2>&1; then
    docker compose "$@"
  elif command -v docker-compose >/dev/null 2>&1; then
    docker-compose "$@"
  else
    echo "Docker Compose is not installed."
    exit 1
  fi
}

mkdir -p "$BACKUP_DIR"

if compose ps --services --filter "status=running" | grep -qx "$SERVICE_NAME"; then
  stamp="$(date +%Y%m%d-%H%M%S)"
  backup_path="$BACKUP_DIR/pre-deploy-runtime-$stamp.tar.gz"
  echo "Creating pre-deploy backup: $backup_path"
  if ! compose exec -T "$SERVICE_NAME" tar -czf - -C /app/app data uploads generated > "$backup_path"; then
    rm -f "$backup_path"
    echo "Backup failed. Deployment stopped."
    exit 1
  fi
else
  echo "No running container found. Skipping pre-deploy backup."
fi

echo "Fetching branch: $DEPLOY_BRANCH"
git fetch origin "$DEPLOY_BRANCH"
git checkout "$DEPLOY_BRANCH"
git reset --hard "origin/$DEPLOY_BRANCH"

echo "Building and starting containers..."
compose up -d --build --remove-orphans

echo "Checking application health..."
for attempt in $(seq 1 30); do
  if compose exec -T "$SERVICE_NAME" python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8088/login', timeout=3).read()" >/dev/null 2>&1; then
    echo "Deployment completed successfully."
    compose ps
    exit 0
  fi
  sleep 2
done

echo "Deployment finished but health check failed."
compose logs --tail=120 "$SERVICE_NAME"
exit 1
