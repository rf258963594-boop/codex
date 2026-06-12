#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")/.."

if [ ! -f ".env" ]; then
  cp .env.example .env
  echo "Created .env from .env.example."
  echo "Please edit DEFAULT_ADMIN_PASSWORD before first production start."
  exit 1
fi

docker compose up -d --build
docker compose ps
