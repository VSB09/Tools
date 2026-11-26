#!/usr/bin/env bash
# SP Shah - CLI Audit Tool Runner
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" >/dev/null 2>&1 && pwd )"
cd "$DIR"

echo "Running SP Shah Job Register Audit Engine..."
python3 run_audit.py "$@"
