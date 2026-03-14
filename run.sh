#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

REBUILD=0

for arg in "$@"; do
  case "$arg" in
    -rb|--rebuild) REBUILD=1 ;;
    *)
      echo "Unknown flag: $arg"
      echo "Usage: $0 [-rb|--rebuild]"
      exit 1
      ;;
  esac
done

if ! command -v docker >/dev/null 2>&1; then
  echo "Error: docker is not installed or not in PATH."
  exit 1
fi

if ! docker compose version >/dev/null 2>&1; then
  echo "Error: docker compose (v2) is not available."
  exit 1
fi

cd "$ROOT_DIR"

DB_PATH="$ROOT_DIR/server/app.db"
if [[ -d "$DB_PATH" ]]; then
  echo "Error: $DB_PATH is a directory, but it must be a file."
  echo "Fix: remove that directory and create an empty file named app.db."
  exit 1
fi
if [[ ! -f "$DB_PATH" ]]; then
  touch "$DB_PATH"
fi

if [[ "$REBUILD" -eq 1 ]]; then
  echo "Rebuilding images and starting containers..."
  docker compose up --build
else
  echo "Starting containers (using cached images)..."
  docker compose up
fi
