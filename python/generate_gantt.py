import matplotlib
matplotlib.use("Agg")  # headless PNG rendering

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font

try:
    from PIL import Image as PILImage
    PILImage.MAX_IMAGE_PIXELS = None  # allow large images
except Exception:
    pass


# -----------------------------
# Configuration (edit if needed)
# -----------------------------
# Scheduling precedence (first non-empty wins)
SCHEDULE_COLS = [
    "StartProject",            # Excel auto-scheduled start (authoritative if you compute it)
    "ChartStart",              # priority-driven start
    "ChartStart_Recon_Date",   # recon/manual as fallback
    "FinalStart",
    "Start",
    "GanttStart",
]

# "Done" states for parent filter (immediate parent only)
CLOSED_STATES = {"Closed", "Done", "Removed", "Resolved"}


def run_gantt(in_path: str, out_path: str, chart_png: str):
    # --- Load first sheet & detect header row ---
    xl = pd.ExcelFile(in_path)
    first_sheet = xl.sheet_names[0]
    df_raw = xl.parse(first_sheet, header=None, dtype=str)

    header_idx = None
    for i in range(min(500, len(df_raw))):
        vals = df_raw.iloc[i].astype(str).tolist()
        if "ID" in vals and "Title" in vals:
            header_idx = i
            break
    if header_idx is None:
        raise RuntimeError("Couldn't find a header row containing 'ID' and 'Title'.")

    headers = [
        h if isinstance(h, str) and h.strip() != "nan" else f"col_{j}"
        for j, h in enumerate(df_raw.iloc[header_idx].tolist())
    ]
    df = df_raw.iloc[header_idx + 1 :].copy()
    df.columns = headers

    # Ensure fields exist
    needed = [
        "ID","Parent","Title","SP","Story Points","Priority","State",
        "StartProject","ChartStart","ChartStart_Recon_Date",
        "FinalStart","Start","GanttStart","ChartEnd","FinalEnd","End","IsParent"
    ]
    for c in needed:
        if c not in df.columns:
            df[c] = np.nan

    # Types
    for c in ["ID","Parent","SP","Story Points","Priority"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    excel_origin = datetime(1899, 12, 30)

    def to_date(v):
        if v is None:
            return None
        if isinstance(v, (datetime, pd.Timestamp)):
            try:
                return pd.to_datetime(v).to_pydatetime()
            except Exception:
                pass
        s = str(v).strip()
        if s == "" or s.lower() in ("nan","none"):
            return None
        # Excel serial?
        try:
            f = float(s)
            if f > 20000:
                return excel_origin + timedelta(days=int(f))
        except Exception:
            pass
        # Trim trailing time if present and parse
        if (" " in s) and s.split(" ")[0].count("-") == 2:
            s = s.split(" ")[0]
        dt = pd.to_datetime(s, errors="coerce")
        return None if pd.isna(dt) else dt.to_pydatetime()

    # Choose SP column
    sp_col = "SP" if df["SP"].notna().any() else ("Story Points" if df["Story Points"].notna().any() else None)
    if sp_col is None:
        df["__SP__"] = np.nan
        sp_col = "__SP__"

    end_pri = ["ChartEnd","FinalEnd","End"]  # preserve original end precedence

    def pick_start_and_source(row):
        for c in SCHEDULE_COLS:
            if c in row and pd.notna(row[c]):
                d = to_date(row[c])
                if d is not None:
                    return d, c
        # Fallbacks if all above empty
        for c in ["FinalStart","Start","GanttStart","ChartStart_Recon_Date"]:
            if c in row and pd.notna(row[c]):
                d = to_date(row[c])
                if d is not None:
                    return d, c
        return None, None

    def pick_end(row, sdt):
        for c in end_pri:
            if c in row and pd.notna(row[c]):
                d = to_date(row[c])
                if d is not None:
                    return d
        # SP fallback
        if sdt is not None:
            sp = pd.to_numeric(row.get(sp_col, np.nan), errors="coerce")
            if pd.notna(sp) and sp > 0:
                return sdt + timedelta(days=int(round(sp)) - 1)
        return None

    # Maps
    id_title = {}
    id_priority = {}
    id_state = {}
    parent_of = {}

    for _, r in df.iterrows():
        rid = pd.to_numeric(r.get("ID"), errors="coerce")
        if pd.isna(rid):
            continue
        rid = int(rid)
        if pd.notna(r.get("Title")):
            id_title[rid] = str(r["Title"])
        if pd.notna(r.get("Priority")):
            try: id_priority[rid] = int(r["Priority"])
            except Exception: pass
        id_state[rid] = str(r.get("State") or "").strip()
        pval = pd.to_numeric(r.get("Parent"), errors="coerce")
        if pd.notna(pval):
            parent_of[rid] = int(pval)

    parent_ids = set(parent_of.values())
    has_is_parent = "IsParent" in df.columns

    def is_parent_row(row) -> bool:
        rid = pd.to_numeric(row.get("ID"), errors="coerce")
        if pd.isna(rid): return False
        rid = int(rid)
        if has_is_parent:
            v = str(row.get("IsParent")).strip().lower()
            if v in ("1","true","yes"):
                return True
        return rid in parent_ids

    # -----------------------------
    # Build child segments
    # -----------------------------
    seg_rows = []
    for _, row in df.iterrows():
        sdt, src = pick_start_and_source(row)
        edt = pick_end(row, sdt)
        if sdt is None or edt is None:
            continue
        if edt < sdt:  # guard negative durations
            edt = sdt

        cid = pd.to_numeric(row.get("ID"), errors="coerce")
        if pd.isna(cid): continue
        cid = int(cid)

        ctitle = (str(row.get("Title")) if pd.notna(row.get("Title"))
                  else (f"ID {cid}" if cid else "(untitled)"))

        parent = pd.to_numeric(row.get("Parent"), errors="coerce")

        # Skip pure parent rows (we only draw children; parents are group headers)
        if pd.isna(parent) and is_parent_row(row):
            continue

        if pd.notna(parent):
            pid = int(parent)
            # Exclude children whose direct parent is closed
            if id_state.get(pid, "") in CLOSED_STATES:
                continue
            gid = pid
            gtitle = id_title.get(gid, f"ID {gid}")
            orphan = False
        else:
            # true orphan child
            gid = -cid
            gtitle = f"[Orphan] {ctitle}"
            orphan = True

        seg_rows.append({
            "gid": gid,
            "gtitle": gtitle,
            "orphan": orphan,
            "cid": cid,
            "ctitle": ctitle,
            "sdt": sdt,
            "edt": edt,
            "start_src": src
        })

    seg_df = pd.DataFrame(seg_rows)
    if seg_df.empty:
        raise RuntimeError("No segments to plot after filtering and scheduling.")

    # -----------------------------
    # Parent ordering: Priority first, then earliest child start
    # -----------------------------
    earliest_by_gid = seg_df.groupby(["gid","gtitle"])["sdt"].min().reset_index()

    def gid_priority(gid):
        return id_priority.get(int(gid), None) if gid > 0 else None

    earliest_by_gid["Priority"] = earliest_by_gid["gid"].apply(
        lambda g: gid_priority(g) if pd.notna(g) else None
    )
    earliest_by_gid["PrioritySort"] = earliest_by_gid["Priority"].fillna(999999).astype(int)
    order_df = earliest_by_gid.sort_values(["PrioritySort","sdt","gid"])
    order = list(order_df[["gid","gtitle"]].itertuples(index=False, name=None))

    # -----------------------------
    # Layout calculations
    # -----------------------------
    min_date = seg_df["sdt"].min()
    max_date = seg_df["edt"].max()
    range_days = (max_date - min_date).days + 1

    dpi = 110
    fig_w = max(24, range_days / 3.0)
    BAR_H = 0.68
    LABEL_GAP = 0.20
    LEVEL_STEP = 0.36
    TOP_PAD = 0.10
    ROW_BOTTOM_GAP = 0.35

    seg_df["x0"] = seg_df["sdt"].apply(mdates.date2num)
    seg_df["w_days"] = seg_df.apply(lambda r: (r["edt"] - r["sdt"]).days + 1, axis=1)
    seg_df["xc"] = seg_df["x0"] + seg_df["w_days"] / 2.0

    xlim_left  = mdates.date2num(min_date - timedelta(days=1))
    xlim_right = mdates.date2num(max_date + timedelta(days=1))
    total_days_shown = xlim_right - xlim_left

    # Measure label widths for stacking
    fig_m, ax_m = plt.subplots(figsize=(fig_w, 2.0), dpi=dpi)
    ax_m.set_xlim(xlim_left, xlim_right); ax_m.set_ylim(0, 1)
    tmp_texts = []
    for idx, s in seg_df.iterrows():
        t = ax_m.text(s["xc"], 0.5, s["ctitle"], ha="center", va="bottom", fontsize=8.2, rotation=0)
        tmp_texts.append((idx, t))
    fig_m.canvas.draw()
    renderer = fig_m.canvas.get_renderer()
    width_px_map = {}
    for idx, t in tmp_texts:
        bb = t.get_window_extent(renderer=renderer)
        width_px_map[idx] = bb.width
        t.remove()
    plt.close(fig_m)

    px_per_day = (dpi * fig_w) / max(total_days_shown, 1e-6)
    seg_df["label_w_days"] = [width_px_map[i] / max(px_per_day, 1e-6) for i in seg_df.index]

    # Assign caption levels per parent row
    seg_df["label_level"] = 0
    max_level_by_gid = {}
    EPS_DAYS = 0.0

    for (gid, gtitle), grp in seg_df.groupby(["gid","gtitle"], sort=False):
        items = []
        for idx, s in grp.iterrows():
            xc, hw = s["xc"], s["label_w_days"]/2.0
            items.append((idx, xc - hw, xc + hw))
        items.sort(key=lambda x: (x[1], x[2]))
        levels_right = []
        for idx, left, right in items:
            assigned = None
            for lev, rgt in enumerate(levels_right):
                if left >= rgt + EPS_DAYS:
                    assigned = lev
                    levels_right[lev] = right
                    break
            if assigned is None:
                levels_right.append(right)
                assigned = len(levels_right) - 1
            seg_df.at[idx, "label_level"] = assigned
        max_level_by_gid[gid] = len(levels_right) if levels_right else 1

    # Compute Y centers considering stacked labels
    y_center = {}
    tick_ys, tick_labels = [], []
    y_cursor = 0.0
    for gid, gtitle in order:
        L = max_level_by_gid.get(gid, 1)
        top_extra = LABEL_GAP + L * LEVEL_STEP + TOP_PAD
        y0 = y_cursor + top_extra + (BAR_H / 2.0)
        y_center[gid] = y0
        tick_ys.append(y0)
        tick_labels.append(gtitle)
        y_cursor = y0 + (BAR_H / 2.0) + ROW_BOTTOM_GAP

    fig_h = max(4.5, y_cursor + 0.6)

    # -----------------------------
    # Draw chart
    # -----------------------------
    fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=dpi)
    colors = {}
    for (gid, gtitle), grp in seg_df.groupby(["gid","gtitle"], sort=False):
        y0 = y_center[gid]
        col = colors.setdefault(gid, plt.cm.tab20((len(colors) % 20) / 20))
        for _, s in grp.sort_values("sdt").iterrows():
            # bar
            ax.broken_barh([(s["x0"], s["w_days"])], (y0 - BAR_H/2.0, BAR_H),
                           facecolors=col, edgecolors="black", linewidth=0.6, zorder=2)
            # caption
            xc = s["xc"]; lvl = int(s["label_level"])
            y_lab = y0 - (BAR_H/2.0) - LABEL_GAP - lvl * LEVEL_STEP
            ax.text(xc, y_lab, s["ctitle"], ha="center", va="bottom",
                    fontsize=8.2, color="black", zorder=5, clip_on=False)
            # leader
            ax.plot([xc, xc], [y0 - BAR_H/2.0, y_lab],
                    color=col, linewidth=0.8, alpha=0.9, zorder=4)

    # Axes & save
    ax.set_yticks(tick_ys); ax.set_yticklabels(tick_labels)
    ax.set_ylim(-0.3, y_cursor + 0.3)
    ax.invert_yaxis()
    ax.set_xlim(xlim_left, xlim_right)
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=1))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b %Y"))
    for t in ax.get_xticklabels():
        t.set_rotation(90); t.set_fontsize(7)
    ax.set_xlabel("Date (daily)")
    ax.set_title("Parent Gantt — Priority-Ordered Rows, Stacked Captions")
    ax.grid(axis="x", linestyle=":", alpha=0.35)

    plt.tight_layout()
    fig.savefig(chart_png, dpi=dpi, bbox_inches="tight")
    plt.close(fig)

    # -----------------------------
    # Export data back to workbook
    # -----------------------------
    wb = load_workbook(in_path)

    data_ws = "Parent Calendar Gantt Data"
    if data_ws in wb.sheetnames:
        wb.remove(wb[data_ws])
    wsd = wb.create_sheet(data_ws)
    wsd.append([
        "RowKey","RowTitle","GroupPriority","ChildID","ChildTitle",
        "Start","End","DurationDays","IsOrphan","ChosenStartSource"
    ])

    # Add rows (sorted for readability)
    for r in seg_df.sort_values(["gid","sdt","cid"]).itertuples(index=False):
        wsd.append([
            int(r.gid),
            order_df.loc[order_df["gid"]==r.gid, "gtitle"].values[0],
            id_priority.get(int(r.gid), None) if r.gid > 0 else None,
            int(r.cid) if r.cid is not None else "",
            r.ctitle,
            r.sdt.strftime("%Y-%m-%d"),
            r.edt.strftime("%Y-%m-%d"),
            (r.edt - r.sdt).days + 1,
            "Yes" if r.orphan else "No",
            r.start_src or ""
        ])

    chart_ws = "Parent Gantt (Daily)"
    if chart_ws in wb.sheetnames:
        wb.remove(wb[chart_ws])
    wsc = wb.create_sheet(chart_ws)
    wsc["A1"] = "Parent Gantt — Daily Axis with Top Labels (Priority-Ordered)"
    wsc["A1"].font = Font(size=14, bold=True)
    img = XLImage(chart_png); img.anchor = "A3"
    wsc.add_image(img)

    wb.save(out_path)