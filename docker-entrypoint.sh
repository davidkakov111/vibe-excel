#!/bin/sh
set -e

wait_for_database() {
  if [ -z "$DATABASE_URL" ]; then
    return 0
  fi

  echo "Waiting for database..."
  attempt=0
  until getent hosts db >/dev/null 2>&1 && nc -z db 5432 >/dev/null 2>&1; do
    attempt=$((attempt + 1))
    if [ "$attempt" -ge 30 ]; then
      echo "Cannot reach db:5432 (DNS or Docker network). Run: docker compose down --remove-orphans && docker compose up --build"
      exit 1
    fi
    echo "db:5432 not reachable yet (attempt $attempt), retrying in 2s..."
    sleep 2
  done

  npx prisma migrate deploy
}

wait_for_database
exec "$@"