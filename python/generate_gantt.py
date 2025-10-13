import matplotlib
matplotlib.use("Agg")  # headless PNG rendering on macOS
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
    PILImage.MAX_IMAGE_PIXELS = 300_000_000  # or None to disable the limit
except Exception:
    pass

def run_gantt(in_path: str, out_path: str, chart_png: str):
    # --- Read FIRST sheet as the authoritative input (do not modify it) ---
    xl = pd.ExcelFile(in_path)
    first_sheet = xl.sheet_names[0]
    df_raw = xl.parse(first_sheet, header=None, dtype=str)

    # Detect header row by presence of ID and Title
    header_idx = None
    for i in range(min(500, len(df_raw))):
        vals = df_raw.iloc[i].astype(str).tolist()
        if 'ID' in vals and 'Title' in vals:
            header_idx = i
            break
    if header_idx is None:
        raise RuntimeError("Couldn't locate header row with 'ID' and 'Title' on the first sheet")

    headers = [h if isinstance(h,str) and h.strip()!='nan' else f'col_{j}' for j,h in enumerate(df_raw.iloc[header_idx].tolist())]
    df = df_raw.iloc[header_idx+1:].copy(); df.columns = headers

    # Ensure expected columns
    needed = ['ID','Parent','Title','SP','Story Points','ChartStart_Recon_Date','ChartStart','ChartEnd','FinalStart','FinalEnd','Start','End','GanttStart']
    for c in needed:
        if c not in df.columns:
            df[c] = np.nan

    # Numeric coercion
    for c in ['ID','Parent','SP','Story Points','GanttStart']:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')

    excel_origin = datetime(1899,12,30)

    def to_date(v):
        if v is None: return None
        if isinstance(v,(datetime,pd.Timestamp)):
            try:
                return pd.to_datetime(v).to_pydatetime()
            except Exception:
                pass
        s = str(v).strip()
        if s=='' or s.lower() in ('nan','none'): return None
        try:
            f = float(s)
            if f > 20000:
                return excel_origin + timedelta(days=int(f))
        except Exception:
            pass
        if (' ' in s) and s.split(' ')[0].count('-')==2:
            s = s.split(' ')[0]
        dt = pd.to_datetime(s, errors='coerce', dayfirst=False)
        return None if pd.isna(dt) else dt.to_pydatetime()

    # Choose SP column
    sp_col = 'SP' if 'SP' in df.columns and df['SP'].notna().any() else ('Story Points' if 'Story Points' in df.columns and df['Story Points'].notna().any() else None)
    if sp_col is None:
        df['__SP__'] = np.nan
        sp_col = '__SP__'

    start_pri = ['ChartStart_Recon_Date','ChartStart','FinalStart','Start','GanttStart']
    end_pri = ['ChartEnd','FinalEnd','End']

    def pick_start(row):
        for c in start_pri:
            if c in row and pd.notna(row[c]):
                d = to_date(row[c])
                if d is not None:
                    return d
        return None

    def pick_end(row, sdt):
        for c in end_pri:
            if c in row and pd.notna(row[c]):
                d = to_date(row[c])
                if d is not None:
                    return d
        if sdt is not None:
            sp = pd.to_numeric(row.get(sp_col, np.nan), errors='coerce')
            if pd.notna(sp) and sp > 0:
                return sdt + timedelta(days=int(round(sp))-1)
        return None

    # Map ID->Title for parent titles
    id_title = {}
    for _, r in df.iterrows():
        rid = pd.to_numeric(r.get('ID'), errors='coerce')
        if pd.notna(rid) and pd.notna(r.get('Title')):
            id_title[int(rid)] = str(r['Title'])

    # Build segments
    rows = []
    for _, row in df.iterrows():
        sdt = pick_start(row); edt = pick_end(row, sdt)
        if sdt is None or edt is None: continue
        cid = pd.to_numeric(row.get('ID'), errors='coerce')
        ctitle = str(row.get('Title')) if pd.notna(row.get('Title')) else (f'ID {int(cid)}' if pd.notna(cid) else '(untitled)')
        parent = pd.to_numeric(row.get('Parent'), errors='coerce') if 'Parent' in row else np.nan
        if pd.notna(parent):
            gid = int(parent); gtitle = id_title.get(gid, f'ID {gid}'); orphan=False
        else:
            gid = -int(cid) if pd.notna(cid) else int(-1e9 - len(rows))
            gtitle = f'[Orphan] {ctitle}'; orphan=True
        rows.append({'gid':gid,'gtitle':gtitle,'orphan':orphan,'cid':int(cid) if pd.notna(cid) else None,'ctitle':ctitle,'sdt':sdt,'edt':edt})

    seg_df = pd.DataFrame(rows)
    if seg_df.empty:
        raise RuntimeError('No segments to plot based on available Start/End')

    min_date = seg_df['sdt'].min(); max_date = seg_df['edt'].max()
    order_df = seg_df.groupby(['gid','gtitle'])['sdt'].min().reset_index().sort_values('sdt')
    order = list(order_df[['gid','gtitle']].itertuples(index=False, name=None))
    ypos = {gid:i for i,(gid,_) in enumerate(order)}

        # --- Render chart: daily axis with above-the-bar captions & leader lines ---
    range_days = (max_date - min_date).days + 1

    # Assign a "label level" per child so captions stack above the bar when time spans overlap
    BAR_H = 0.68
    LABEL_GAP = 0.18       # gap from top of bar to first caption line
    LEVEL_STEP = 0.36      # vertical distance between caption lines
    TOP_PAD = 0.08         # padding above the highest caption line
    BOTTOM_GAP = 0.35      # gap below bar before next parent row

    seg_df['label_level'] = 0
    max_level_by_gid = {}

    for (gid, gtitle), grp in seg_df.groupby(['gid', 'gtitle'], sort=False):
        # Greedy interval stacking: smallest number of vertical levels so overlapping bars get separate lines
        levels_right = []  # rightmost x (date2num) occupied per level
        for idx, s in grp.sort_values('sdt').iterrows():
            left = mdates.date2num(s['sdt'])
            w = (s['edt'] - s['sdt']).days + 1
            right = left + w
            assigned = None
            for l, rgt in enumerate(levels_right):
                if left >= rgt:  # no overlap with that level
                    assigned = l
                    levels_right[l] = right
                    break
            if assigned is None:
                levels_right.append(right)
                assigned = len(levels_right) - 1
            seg_df.at[idx, 'label_level'] = assigned
        max_level_by_gid[gid] = len(levels_right) if levels_right else 1

    # Compute per-parent Y centers with extra headroom for stacked captions
    order_df = seg_df.groupby(['gid','gtitle'])['sdt'].min().reset_index().sort_values('sdt')
    order = list(order_df[['gid','gtitle']].itertuples(index=False, name=None))

    y_center = {}
    tick_ys = []
    tick_labels = []
    y_cursor = 0.0
    for gid, gtitle in order:
        L = max_level_by_gid.get(gid, 1)
        top_extra = LABEL_GAP + L * LEVEL_STEP + TOP_PAD
        y0 = y_cursor + top_extra + (BAR_H / 2.0)
        y_center[gid] = y0
        tick_ys.append(y0)
        tick_labels.append(gtitle)
        # Move cursor to start of next parent row space
        y_cursor = y0 + (BAR_H / 2.0) + BOTTOM_GAP

    # Stretch width so dates remain readable (you already raised MAX_IMAGE_PIXELS; this can get large)
    fig_w = max(24, range_days / 3.0)
    # Height proportional to total stacked space used
    fig_h = max(4.5, 0.46 * y_cursor + 1.0)

    fig, ax = plt.subplots(figsize=(fig_w, fig_h))

    colors = {}
    for (gid, gtitle), grp in seg_df.groupby(['gid','gtitle'], sort=False):
        y0 = y_center[gid]
        col = colors.setdefault(gid, plt.cm.tab20((len(colors)%20)/20))
        for _, s in grp.sort_values('sdt').iterrows():
            # Bar
            x0 = mdates.date2num(s['sdt'])
            w = (s['edt'] - s['sdt']).days + 1
            ax.broken_barh([(x0, w)], (y0 - BAR_H/2.0, BAR_H),
                           facecolors=col, edgecolors='black', linewidth=0.6, zorder=2)
            # Caption above bar
            xc = mdates.date2num(s['sdt'] + timedelta(days=w/2.0))
            lvl = int(s['label_level'])
            y_label = y0 - (BAR_H/2.0) - LABEL_GAP - lvl * LEVEL_STEP
            ax.text(xc, y_label, s['ctitle'],
                    ha='center', va='bottom', fontsize=8.2, color='black', zorder=5, clip_on=False)
            # Leader line from caption down to bar top
            ax.plot([xc, xc], [y0 - BAR_H/2.0, y_label],
                    color=col, linewidth=0.8, alpha=0.9, zorder=4)

    # Y axis: ticks at each parent center
    ax.set_yticks(tick_ys)
    ax.set_yticklabels(tick_labels)
    ax.set_ylim(-0.3, y_cursor + 0.3)
    ax.invert_yaxis()  # earliest parent at the top

    # X axis: daily labels (you can keep your previous adaptive logic if desired)
    ax.set_xlim(mdates.date2num(min_date - timedelta(days=1)), mdates.date2num(max_date + timedelta(days=1)))
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=1))
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%d %b %Y'))
    for t in ax.get_xticklabels():
        t.set_rotation(90); t.set_fontsize(7)

    ax.set_xlabel('Date (daily)')
    ax.set_title('Parent (single row) Gantt — Captions Above with Leader Lines (Earliest at Top)')
    ax.grid(axis='x', linestyle=':', alpha=0.35)

    plt.tight_layout()
    fig.savefig(chart_png, dpi=110, bbox_inches='tight')
    plt.close(fig)

    # --- Open the ORIGINAL workbook; add outputs only ---
    wb = load_workbook(in_path)

    # Add/replace data sheet safely
    data_ws = 'Parent Calendar Gantt Data'
    if data_ws in wb.sheetnames:
        wb.remove(wb[data_ws])
    wsd = wb.create_sheet(data_ws)
    wsd.append(['RowKey','RowTitle','ChildID','ChildTitle','Start','End','DurationDays','IsOrphan'])
    for r in seg_df.sort_values(['gid','sdt','cid']).itertuples(index=False):
        wsd.append([int(r.gid), r.gtitle, '' if r.cid is None else int(r.cid), r.ctitle,
                    r.sdt.strftime('%Y-%m-%d'), r.edt.strftime('%Y-%m-%d'), (r.edt-r.sdt).days+1,
                    'Yes' if r.orphan else 'No'])

    # Add/replace chart sheet safely
    chart_ws = 'Parent Gantt (Daily)'
    if chart_ws in wb.sheetnames:
        wb.remove(wb[chart_ws])
    wsc = wb.create_sheet(chart_ws)
    wsc['A1'] = 'Parent Gantt — Daily Axis with Top Labels (Earliest at Top)'
    wsc['A1'].font = Font(size=14, bold=True)
    img = XLImage(chart_png)
    img.anchor = 'A3'
    wsc.add_image(img)

    wb.save(out_path)
