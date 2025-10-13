#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"
PYTHON_BIN="${PYTHON_BIN:-python3}"

if ! command -v "$PYTHON_BIN" >/dev/null 2>&1; then
  echo "ERROR: python3 not found on PATH. On macOS, install with:  brew install python"
  exit 1
fi

$PYTHON_BIN -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip wheel
pip install pandas numpy matplotlib openpyxl pillow
echo "\u2713 Virtual env ready at $(pwd)/.venv"