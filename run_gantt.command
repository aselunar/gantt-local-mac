#!/bin/zsh
set -euo pipefail

APPDIR="$(cd "$(dirname "$0")" && pwd)"
VENV="$APPDIR/.venv"
PY="$VENV/bin/python3"

# Ensure venv + deps on first run
if [ ! -x "$PY" ]; then
  echo "Setting up Python environment (first run)..."
  "$APPDIR/install.sh"
fi

# --- File picker (works on your Mac; single-line osascript) ---
FILE="$(
/usr/bin/osascript \
  -e 'set theTypes to {"org.openxmlformats.spreadsheetml.sheet","com.microsoft.excel.xlsx","xlsx"}' \
  -e 'set f to choose file with prompt "Choose the Excel workbook (.xlsx)" of type theTypes' \
  -e 'POSIX path of f'
)" || true

# Fallback: manual prompt (supports drag-drop)
if [ -z "${FILE:-}" ]; then
  echo "Dialog canceled. Paste full path to your .xlsx (or drag the file here) and press Enter:"
  read -r FILE
fi

# Trim quotes if pasted
FILE="${FILE%\"}"; FILE="${FILE#\"}"; FILE="${FILE%\'}"; FILE="${FILE#\'}"

if [ ! -f "$FILE" ]; then
  echo "The path does not point to a file: $FILE"
  exit 2
fi

# Outputs next to input
DIR="$(dirname "$FILE")"
BASENAME="$(basename "$FILE")"
STEM="${BASENAME%.*}"
OUT_XLSX="$DIR/${STEM} - Gantt (Daily Top Labels, Preserve Input).xlsx"
OUT_PNG="$DIR/Parent_Gantt_Daily_Top_preserve.png"

# Run renderer
"$PY" "$APPDIR/python/run_gantt.py" \
  --input "$FILE" \
  --output-xlsx "$OUT_XLSX" \
  --output-png "$OUT_PNG"

# Reveal PNG
open -R "$OUT_PNG"
