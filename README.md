# Local Gantt PNG Builder (macOS)

This folder runs your original **Matplotlib + openpyxl** pipeline locally—no cloud—so you can iterate on Python until the PNG looks perfect.

## Quick start
1. Double-click `run_gantt.command`.
2. Pick your **.xlsx**.
3. Get:
   - `Parent_Gantt_Daily_Top_preserve.png` (next to the input), and
   - a new workbook `… - Gantt (Daily Top Labels, Preserve Input).xlsx` with the image embedded.

> First run sets up a Python **virtual environment** under `.venv` and installs `pandas`, `numpy`, `matplotlib`, `openpyxl`, `pillow`.

## Customize the rendering
- Edit `python/generate_gantt.py` (colors, fonts, label sizes, figure width/height, etc.).
- Re-run `run_gantt.command`.

## Manual CLI (optional)
```bash
# From this folder:
./install.sh
./.venv/bin/python3 python/run_gantt.py \
  --input "/path/to/Full Backlog Iteration 2.xlsx" \
  --output-xlsx "/path/to/Full Backlog Iteration 2 - Gantt (Daily Top Labels, Preserve Input).xlsx" \
  --output-png "/path/to/Parent_Gantt_Daily_Top_preserve.png"
```

## Notes
- Works on macOS; Power Automate Desktop isn’t needed.
- If you store files in OneDrive locally, choose the synced path when prompted.
