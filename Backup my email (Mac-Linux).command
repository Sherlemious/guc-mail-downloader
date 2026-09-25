#!/usr/bin/env bash
# Double-click on macOS, or run:  bash "Backup my email (Mac-Linux).command"
cd "$(dirname "$0")" || exit 1

PY=""
for c in python3 python; do
    if command -v "$c" >/dev/null 2>&1 &&
       "$c" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)' 2>/dev/null; then
        PY="$c"; break
    fi
done
if [ -z "$PY" ]; then
    echo "Python 3 is not installed. Get it from https://www.python.org/downloads/ and run this again."
    read -r -p "Press Enter to close."
    exit 1
fi

if [ ! -x .venv/bin/python ]; then
    echo "First run: setting things up, this takes a minute..."
    "$PY" -m venv .venv || { echo "Could not create the Python environment."; read -r -p "Press Enter to close."; exit 1; }
fi
if [ ! -f .venv/installed.txt ]; then
    .venv/bin/python -m pip install --disable-pip-version-check -q -r requirements.txt ||
        { echo "Could not install the requirements. Check your internet connection."; read -r -p "Press Enter to close."; exit 1; }
    echo ok > .venv/installed.txt
fi

exec .venv/bin/python guc_mail_backup.py "$@"
