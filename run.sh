#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SERVER_DIR="$ROOT_DIR/server"
CLIENT_DIR="$ROOT_DIR/client"

if ! command -v uv >/dev/null 2>&1; then
  echo "Error: uv is not installed or not in PATH."
  exit 1
fi

if ! command -v npm >/dev/null 2>&1; then
  echo "Error: npm is not installed or not in PATH."
  exit 1
fi

cleaned_up=0
cleanup() {
  if [[ "$cleaned_up" -eq 1 ]]; then
    return
  fi
  cleaned_up=1

  echo ""
  echo "Stopping server and client..."
  if [[ -n "${SERVER_PID:-}" ]] && kill -0 "$SERVER_PID" 2>/dev/null; then
    kill "$SERVER_PID" 2>/dev/null || true
  fi
  if [[ -n "${CLIENT_PID:-}" ]] && kill -0 "$CLIENT_PID" 2>/dev/null; then
    kill "$CLIENT_PID" 2>/dev/null || true
  fi
}

trap cleanup INT TERM EXIT

echo "Starting backend on http://127.0.0.1:8009 ..."
(
  cd "$SERVER_DIR"
  uv run uvicorn main:app --host 127.0.0.1 --port 8009
) &
SERVER_PID=$!

echo "Starting frontend on http://localhost:3002 ..."
(
  cd "$CLIENT_DIR"
  npm run dev
) &
CLIENT_PID=$!

echo "Both processes started."
echo "Backend PID: $SERVER_PID | Frontend PID: $CLIENT_PID"

wait -n "$SERVER_PID" "$CLIENT_PID"
