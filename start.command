#!/usr/bin/env bash
# Notater - Start (macOS / Linux)
NOTATER_DIR="$(cd "$(dirname "$0")" && pwd)"
VENV_PY="$NOTATER_DIR/venv/bin/python"

if [ ! -x "$VENV_PY" ]; then
    echo "Notater er ikke installert. Kjør install.sh først."
    exit 1
fi

pkill -f "backend/main.py" >/dev/null 2>&1 || true
sleep 1
"$VENV_PY" "$NOTATER_DIR/backend/main.py" &
