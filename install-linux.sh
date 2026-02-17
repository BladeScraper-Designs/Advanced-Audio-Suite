#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

if command -v python3 >/dev/null 2>&1; then
  PYTHON_BIN="python3"
elif command -v python >/dev/null 2>&1; then
  PYTHON_BIN="python"
else
  echo "Python 3 is not installed or not in PATH. Install Python 3.11+ and try again." >&2
  exit 1
fi

if [[ ! -d ".venv" ]]; then
  echo "Creating virtual environment..."
  "$PYTHON_BIN" -m venv .venv
fi

VENV_PY=".venv/bin/python"
if [[ ! -x "$VENV_PY" ]]; then
  echo "Virtual environment Python not found at $VENV_PY" >&2
  exit 1
fi

echo "Installing dependencies..."
"$VENV_PY" -m pip install --upgrade pip
"$VENV_PY" -m pip install -r requirements.txt

echo
echo "Setup complete. Run the app with:"
echo "./.venv/bin/python AAS.py"
