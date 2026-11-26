#!/usr/bin/env bash
# SP Shah - Web Dashboard Runner
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" >/dev/null 2>&1 && pwd )"
cd "$DIR"

PORT=${PORT:-8055}
echo "Starting SP Shah Job Register Auditor Web Dashboard on port $PORT..."
echo "Open your browser at: http://localhost:$PORT"
python3 app.py
