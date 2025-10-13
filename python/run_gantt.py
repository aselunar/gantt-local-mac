import argparse, os, sys
from generate_gantt import run_gantt

def main():
    ap = argparse.ArgumentParser(description="Render Parent Gantt PNG and updated workbook from Excel input.")
    ap.add_argument("--input", "-i", required=True, help="Path to source Excel workbook (.xlsx).")
    ap.add_argument("--output-xlsx", "-o", required=True, help="Path to NEW output workbook (.xlsx).")
    ap.add_argument("--output-png", "-p", required=True, help="Path to PNG to write.")
    args = ap.parse_args()

    src = os.path.abspath(args.input)
    out_xlsx = os.path.abspath(args.output_xlsx)
    out_png = os.path.abspath(args.output_png)

    if not os.path.isfile(src):
        print(f"Input workbook not found: {src}", file=sys.stderr)
        sys.exit(2)

    os.makedirs(os.path.dirname(out_xlsx), exist_ok=True)
    os.makedirs(os.path.dirname(out_png), exist_ok=True)

    run_gantt(src, out_xlsx, out_png)
    print("\u2713 PNG:", out_png)
    print("\u2713 Workbook:", out_xlsx)

if __name__ == "__main__":
    main()
