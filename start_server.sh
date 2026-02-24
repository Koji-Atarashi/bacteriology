#!/usr/bin/env bash
# start_server.sh — install deps and launch the search web UI
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
DB="${SEARCH_DB:-/mnt/hdd1/.search_index.sqlite3}"
HOST="${SEARCH_HOST:-0.0.0.0}"
PORT="${SEARCH_PORT:-8090}"

# Install Python dependencies if needed
if ! python3 -c "import flask" 2>/dev/null; then
    echo "[setup] Installing dependencies..."
    pip install -r "$SCRIPT_DIR/requirements.txt"
fi

echo "[info] Starting search server on http://${HOST}:${PORT}"
echo "[info] DB: ${DB}"
exec python3 "$SCRIPT_DIR/search_indexer.py" serve \
    --db "$DB" \
    --host "$HOST" \
    --port "$PORT"
