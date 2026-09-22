#!/usr/bin/env bash
# Avvio del programma su Linux / macOS:  ./avvia.sh   (oppure: bash avvia.sh)
cd "$(dirname "$0")" || exit 1
if command -v python3 >/dev/null 2>&1; then PY=python3
elif command -v python >/dev/null 2>&1; then PY=python
else
    echo "Python 3 non è installato. Su Ubuntu/Debian:  sudo apt install python3 python3-venv python3-tk"
    exit 1
fi
exec "$PY" avvia.py "$@"
