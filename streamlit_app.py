"""
D2 vs Competitor — Multi-Destination Pricing Dashboard
Streamlit app · upload .xlsm/.xlsx/.csv · persistent SQLite DB
"""
import streamlit as st
import pandas as pd
import numpy as np
import sqlite3
import re
import os
import json
from datetime import datetime, date
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

try:
    from rapidfuzz import fuzz
    HAS_FUZZ = True
except ImportError:
    HAS_FUZZ = False

st.set_page_config(page_title="D2 Pricing Intelligence", page_icon="✈",
                   layout="wide", initial_sidebar_state="expanded")

DB_PATH = "pricing_data.db"

# Nuclear DB reset on startup if flag file exists
if os.path.exists("reset_flag.txt"):
    try:
        os.remove(DB_PATH)
        os.remove("reset_flag.txt")
    except:
        pass

MARGIN_FLOOR    = 7.0
MARGIN_CEILING  = 15.0
MARGIN_NEAR     = 8.5
RAISE_THRESHOLD = 15.0
HOLD_UPPER      = 8.0
# ── Markup-suggestion tuning ────────────────────────────────────────────────────
NOISE_GAP       = 0.5   # |diff%| within this of zero = effectively even (pricing/refresh
                        #   noise) — never chase such a "loss" with a margin cut
WR_HIGH         = 70.0  # at/above this win rate, bias to RAISE (capture margin) not cut
WR_KEEP         = 70.0  # when raising, don't let win rate fall below this
FLIP_FRAC       = 0.10  # when raising, tolerate flipping at most ~10% of solid wins
# ── Product-gap detection ───────────────────────────────────────────────────────
PRODUCT_GAP_SHARE = 0.5  # if >= this share of a hotel's losses can't be matched without
                         #   breaching the margin floor, it's a product/contracting gap
PRODUCT_GAP_MIN   = 2    # need at least this many below-floor losses to flag a hotel

st.markdown("""
<style>
[data-testid="stSidebar"]{background:#1B1464}
[data-testid="stSidebar"] *{color:#fff!important}
</style>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# DATABASE
# ══════════════════════════════════════════════════════════════════════════════
def get_conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)

def init_db():
    conn = get_conn()
    conn.execute("""CREATE TABLE IF NOT EXISTS pricing(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        file_name TEXT, file_date TEXT, destination TEXT, competitor TEXT,
        giata TEXT, hotel_name TEXT, board TEXT,
        dep_date TEXT, dep_month TEXT, dep_window TEXT, nights INTEGER,
        pp_price REAL, current_margin REAL, op_room TEXT,
        comp_price REAL, diff_gbp REAL, diff_pct REAL,
        result TEXT, margin_after REAL, margin_range TEXT,
        margin_flag TEXT, booking_tier TEXT, priority_score REAL,
        uploaded_at TEXT)""")
    conn.execute("""CREATE TABLE IF NOT EXISTS bookings(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        hotel_name TEXT, destination TEXT,
        bkgs_4wk INTEGER, bkgs_py INTEGER, uploaded_at TEXT)""")
    # Giata-keyed booking volumes (one row per giata; aggregated from the export)
    conn.execute("""CREATE TABLE IF NOT EXISTS bkg(
        giata TEXT PRIMARY KEY, hotel_name TEXT, dest_name TEXT,
        bookings INTEGER, revenue REAL, uploaded_at TEXT)""")
    # Booking time-series: bookings aggregated by giata × booked-month × departure-month
    conn.execute("""CREATE TABLE IF NOT EXISTS bkg_tx(
        giata TEXT, booked_month TEXT, dep_month TEXT,
        bookings INTEGER, revenue REAL)""")
    # Markup rule-engine export (flattened, one row per rule)
    conn.execute("""CREATE TABLE IF NOT EXISTS rules(
        rid TEXT, enabled INTEGER, markup_type TEXT, applies_to TEXT,
        layer TEXT, verb TEXT, action_value REAL, action_unit TEXT,
        ispkg TEXT, user_csv TEXT, giatas TEXT, giata_n INTEGER, n_conds INTEGER,
        generated_text TEXT, uploaded_at TEXT)""")
    conn.commit(); conn.close()

init_db()

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def parse_date(val):
    if val is None: return None
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.date() if hasattr(val,'date') else val
    s = str(val).strip()
    if not s or s in ("","nan","None","NaT"): return None
    if " 00:00:00" in s: s = s.replace(" 00:00:00","").strip()
    for fmt in ("%Y-%m-%d","%d/%m/%Y","%d-%m-%Y","%m/%d/%Y"):
        try: return datetime.strptime(s, fmt).date()
        except: pass
    return None

def dep_window(dep_dt, run_dt):
    if dep_dt is None: return "241+"
    days = (dep_dt - run_dt).days
    if days <= 60:  return "0-60"
    if days <= 120: return "61-120"
    if days <= 240: return "121-240"
    return "241+"

def margin_flag(v):
    try: v = float(v)
    except: return "unknown"
    if v < MARGIN_FLOOR:   return "floor"
    if v < MARGIN_NEAR:    return "near"
    if v > MARGIN_CEILING: return "ceiling"
    return "ok"

def clean_num(val):
    if val is None: return np.nan
    s = str(val).replace("£","").replace(",","").replace(" ","").strip()
    try: return float(s)
    except: return np.nan

def extract_meta(fname):
    base = re.sub(r'\.(xlsm|xlsx|csv)$','',fname,flags=re.I)
    parts = base.split("_")
    dest = parts[2] if len(parts)>2 else "UNK"
    comp = parts[3] if len(parts)>3 else "LH"
    try:
        d,m,y = int(parts[-3]),int(parts[-2]),int(parts[-1])
        fd = date(y,m,d)
    except:
        fd = date.today()
    return dest, comp, fd

def derive_stance(wr, avg_lose, avg_win):
    if wr < 0.45 or avg_lose < -10: return "Aggressively Reduce"
    if wr < 0.55 or avg_lose < -5:  return "Selective Adjust"
    if wr >= 0.65 and avg_win > 15: return "Increase Pricing"
    if wr >= 0.55:                  return "Hold — Optimise Margin"
    return "Product-Led Fix"

def margin_cell(cur, after):
    try:
        delta = float(after)-float(cur)
        flag  = margin_flag(after)
        lbl   = {"floor":"⚠ FLOOR","near":"NEAR","ceiling":"↑ CEILING","ok":"","unknown":""}[flag]
        return f"{float(cur):.1f}%→{float(after):.1f}% {lbl} ({delta:+.1f}pp)"
    except: return "—"

def calc_priority(row):
    tw = {"high":3,"medium":2,"low":1}.get(row.get("booking_tier","low"),1)
    gap = abs(row.get("diff_pct",0))
    gs  = 1 if gap<5 else 2 if gap<10 else 3 if gap<20 else 4
    mf  = row.get("margin_flag","ok")
    ms  = 0.1 if mf=="floor" else 0.5 if mf=="near" else 1.5 if mf=="ceiling" else 1.0
    dw  = row.get("dep_window","241+")
    di  = 1.0 if dw=="0-60" else 1.1 if dw=="61-120" else 1.3 if dw=="121-240" else 1.2
    return round(tw*gs*ms*di,3)

def suggest_markup_change(hrows, step=0.25, band=5.0):
    """High-level suggested change to Total Markup (margin, in pp) for ONE hotel,
    computed across its real comparison rows (Win/Lose) over all departure dates.
    Win - No Comp rows are excluded entirely and never influence the result.

    Idea: markup is the lever. Cutting markup lowers our price → more wins but
    less margin; raising markup does the reverse. We sweep candidate markup
    deltas and pick the one that maximises EXPECTED BOOKED MARGIN, defined as

        projected_win_rate  ×  projected_avg_markup

    i.e. how often we win × how much margin we keep. This naturally balances the
    two asks: over-cutting kills margin (term 2 falls), and raising too far
    converts wins to losses (term 1 falls), so the optimum sits in between.

    Two refinements:
      • Noise band — a "loss" within NOISE_GAP of break-even (pennies) is treated as
        effectively even and never chased with a margin cut (likely refresh noise).
      • High win rate (>= WR_HIGH) flips the goal to RAISING margin: take the largest
        raise that keeps win rate >= WR_KEEP and flips at most ~FLIP_FRAC of solid wins,
        rather than shaving margin across many wins to convert one marginal loss.

    Guardrails (never violated):
      • no individual row is ever pushed below the 7% floor
      • no individual row is ever pushed above the 15% ceiling
      • among near-optimal deltas we pick the SMALLEST move (give away least margin)

    Price sensitivity: a 1pp markup change moves price by ~100/(100+markup) %, so
    a row's competitive gap shifts by that much. diff_pct>0 = D2 winning (cheaper),
    diff_pct<0 = D2 losing (dearer); raising markup (delta>0) makes diff_pct smaller.

    Returns dict: delta, wr0, wr1, m0, m1, note.
    """
    h = hrows.copy()
    # Win - No Comp rows never feed a markup decision: drop them BEFORE anything else,
    # so a hotel that has real comparisons on OTHER dates is decided on those — even if
    # its most recent run happened to return no competitor for some departures.
    comp = h[h["result"].isin(["Win","Lose"])].copy()
    base_m = pd.to_numeric(h.get("current_margin"), errors="coerce").mean()
    base_m = float(base_m) if pd.notna(base_m) else 0.0

    if comp.empty:
        # Genuinely no comparison on ANY date for this hotel → handled in No Comp tab
        return {"delta":0.0,"wr0":np.nan,"wr1":np.nan,"m0":base_m,"m1":base_m,
                "note":"No direct comp — raise selectively (dest-guided)"}

    # Among the real comparisons only, use the latest run (dedupes re-compared departures)
    if "file_date" in comp.columns and comp["file_date"].nunique() > 1:
        comp = comp[comp["file_date"] == comp["file_date"].max()]

    m_rows = pd.to_numeric(comp["current_margin"], errors="coerce").fillna(0.0).values
    gaps   = pd.to_numeric(comp["diff_pct"],       errors="coerce").fillna(0.0).values
    m0  = float(np.mean(m_rows))
    # Noise-tolerant win test: a "loss" within NOISE_GAP of break-even (e.g. a few
    # pennies) is treated as effectively even — almost always pricing/refresh noise,
    # never worth chasing with a margin cut.
    wr0 = float((gaps > -NOISE_GAP).mean()*100)

    # Safe delta band: keep every row inside [floor, ceiling]
    d_min = -max(0.0, float(np.min(m_rows)) - MARGIN_FLOOR)
    d_max =  max(0.0, MARGIN_CEILING - float(np.max(m_rows)))
    d_min, d_max = max(d_min, -band), min(d_max, band)
    if d_max - d_min < step:
        return {"delta":0.0,"wr0":wr0,"wr1":wr0,"m0":m0,"m1":m0,
                "note":"At margin guardrail — hold"}

    factors = 100.0/(100.0+np.clip(m_rows, 0, None))   # %price move per pp markup

    # ── HIGH WIN RATE → raise to capture margin (don't chase the last win) ──────
    if wr0 >= WR_HIGH:
        solid = int((gaps > NOISE_GAP).sum())          # wins with genuine headroom
        cap   = int(np.floor(FLIP_FRAC * solid))       # how many we'll let flip
        delta, wr1, m1 = 0.0, wr0, m0
        d = step
        while d <= d_max + 1e-9:
            new   = gaps - d*factors
            wr1c  = float((new > -NOISE_GAP).mean()*100)
            flips = int(((gaps > NOISE_GAP) & (new <= -NOISE_GAP)).sum())  # solid win → real loss
            if wr1c >= WR_KEEP and flips <= cap:
                delta, wr1, m1 = round(d,2), wr1c, m0+d        # keep largest feasible raise
            d += step
        delta = round(delta, 1)
        if delta >= 0.05:
            note = "High win rate — raise to capture margin"
        elif solid == 0:
            note = "High win rate but near parity — hold"
        else:
            note = "High win rate — no safe room to raise (ceiling/win threshold)"
        return {"delta":delta,"wr0":wr0,"wr1":wr1,"m0":m0,"m1":m1,"note":note}

    # ── CONTESTED (win rate < WR_HIGH) → balanced optimisation ─────────────────
    cands = []
    d = d_min
    while d <= d_max + 1e-9:
        m1 = m0 + d
        if m1 >= MARGIN_FLOOR - 1e-9:
            wr1 = float(((gaps - d*factors) > -NOISE_GAP).mean()*100)  # noise-tolerant
            obj = (wr1/100.0) * m1                              # expected booked margin
            cands.append((round(d,2), wr1, m1, obj))
        d += step

    if not cands:
        return {"delta":0.0,"wr0":wr0,"wr1":wr0,"m0":m0,"m1":m0,
                "note":"At margin guardrail — hold"}

    best_obj = max(c[3] for c in cands)
    near = [c for c in cands if c[3] >= best_obj - 0.02]   # treat ~ties as equal
    near.sort(key=lambda c:(abs(c[0]), -c[2]))             # smallest move, then higher margin
    delta, wr1, m1, _ = near[0]
    delta = round(delta, 1)

    if delta <= -0.05:
        flipped = int(round((wr1-wr0)/100.0 * len(gaps)))
        note = f"Cut markup → ~{max(flipped,0)} loss(es) flip to win"
    elif delta >= 0.05:
        note = "Headroom to raise markup (wins safe)"
    elif wr0 < 50:
        note = "Product-led — losses too wide for markup to fix"
    else:
        note = "Maintain — already optimal"
    return {"delta":delta,"wr0":wr0,"wr1":wr1,"m0":m0,"m1":m1,"note":note}

# ══════════════════════════════════════════════════════════════════════════════
# PARSER
# ══════════════════════════════════════════════════════════════════════════════
def find_header_row(df_raw):
    for i in range(min(10, len(df_raw))):
        try:
            vals = [str(v).lower() for v in df_raw.iloc[i].tolist()]
            row_str = " ".join(vals)
            if "giata" in row_str or ("hotel" in row_str and "price" in row_str):
                return i
        except: continue
    return 0

def find_data_sheets(xl):
    results = []
    priority = ["d2 - data","data","mle","dxb","mru","tfs","ace"]
    skip = ["board codes","namegiata","summary pivot","exclusiv","direct",
            "lh only","tr only","otb only","tb - data","tui - data",
            "lh - data","tr - data","otb - data","first search",
            "mle trans","sheet1","sheet2","sheet3","sheet4","sheet5",
            "exclusive offers","price range","directproduct"]
    sheet_names_lower = {sn.lower(): sn for sn in xl.sheet_names}
    for p in priority:
        if p in sheet_names_lower:
            sn = sheet_names_lower[p]
            try:
                raw = xl.parse(sn, header=None, dtype=str)
                if raw.shape[1] >= 18 and raw.shape[0] >= 5:
                    results.append((sn, raw))
            except: pass
    if not results:
        for sn in xl.sheet_names:
            if any(s in sn.lower() for s in skip): continue
            try:
                raw = xl.parse(sn, header=None, dtype=str)
                if raw.shape[1] >= 18 and raw.shape[0] >= 5:
                    results.append((sn, raw))
            except: pass
    return results

def find_price_col(data_df):
    for ci in range(4, min(9, data_df.shape[1])):
        col = pd.to_numeric(
            data_df.iloc[:10, ci].astype(str).str.replace(r"[£,]","",regex=True),
            errors="coerce").dropna()
        if len(col) > 0 and col.gt(500).any():
            return ci
    return 5

def parse_pricing_file(uploaded_file, fname):
    dest, comp, file_dt = extract_meta(fname)
    all_rows = []; debug_info = []
    try:
        if fname.lower().endswith(".csv"):
            raw = pd.read_csv(uploaded_file, header=None, dtype=str, encoding="utf-8-sig")
            sheets = [("csv", raw)]; debug_info.append("csv")
        else:
            xl = pd.ExcelFile(uploaded_file, engine="openpyxl")
            debug_info = xl.sheet_names
            sheets = find_data_sheets(xl)
            if not sheets:
                return [], [f"No data sheets found in: {xl.sheet_names}"]
    except Exception as e:
        return [], [f"ERROR: {e}"]

    for sname, raw in sheets:
        hrow = find_header_row(raw)
        data = raw.iloc[hrow+1:].reset_index(drop=True).copy()
        data.columns = range(data.shape[1])
        price_col = find_price_col(data)
        offset = price_col - 5
        debug_info.append(f"Sheet='{sname}' hrow={hrow} price_col={price_col} offset={offset} cols={data.shape[1]} rows={data.shape[0]}")
        if len(data) > 0:
            r0 = data.iloc[0]
            debug_info.append(f"Row0: col0={r0.get(0)} col1={r0.get(1)} col3={r0.get(3)} col5={r0.get(5)} col6={r0.get(6)} col21={r0.get(21)}")

        def gc(base):
            idx = base + offset
            return idx if 0 <= idx < data.shape[1] else base

        for _, row in data.iterrows():
            try:
                giata = str(row.get(gc(0),"")).strip()
                if not giata or giata in ("nan","-","","None"): continue
                try: float(giata)
                except: continue
                hotel = str(row.get(gc(1),"")).strip()
                if not hotel or hotel in ("nan","None",""): continue
                board   = str(row.get(gc(2),"")).strip()
                dep_raw = row.get(gc(3),"")
                dep_dt  = parse_date(dep_raw)
                nights  = int(clean_num(row.get(gc(4),7)) or 7)
                pp  = clean_num(row.get(gc(5)))
                cm  = clean_num(row.get(gc(6)))
                oproom = str(row.get(gc(7),"")).strip()
                cp = clean_num(row.get(gc(18)))
                if pd.isna(cp) or cp == 0:
                    for ci in [15,16,17]:
                        v = clean_num(row.get(gc(ci)))
                        if not pd.isna(v) and v > 0: cp = v; break
                diff_g = clean_num(row.get(gc(19)))
                diff_p = clean_num(row.get(gc(20)))
                result = str(row.get(gc(21),"")).strip()
                ma     = clean_num(row.get(gc(22)))
                mr     = str(row.get(gc(23),"")).strip()
                if pd.isna(pp) or pp < 100: continue
                rl = result.lower()
                if "no comp" in rl:               result = "Win - No Comp"
                elif "aft change" in rl:           result = "Win Aft Change"
                elif "win" in rl:                  result = "Win"
                elif "lose" in rl or "loss" in rl: result = "Lose"
                else: continue
                dep_month = dep_dt.strftime("%Y-%m") if dep_dt else ""
                dw = dep_window(dep_dt, file_dt) if dep_dt else "241+"
                mf = margin_flag(ma if not pd.isna(ma) else 0)
                all_rows.append({
                    "file_name":fname,"file_date":str(file_dt),"destination":dest,
                    "competitor":comp,"giata":giata,"hotel_name":hotel,"board":board,
                    "dep_date":str(dep_dt) if dep_dt else "","dep_month":dep_month,
                    "dep_window":dw,"nights":nights,"pp_price":float(pp),
                    "current_margin":float(cm) if not pd.isna(cm) else 0.0,
                    "op_room":oproom,"comp_price":float(cp) if not pd.isna(cp) else 0.0,
                    "diff_gbp":float(diff_g) if not pd.isna(diff_g) else 0.0,
                    "diff_pct":float(diff_p) if not pd.isna(diff_p) else 0.0,
                    "result":result,"margin_after":float(ma) if not pd.isna(ma) else 0.0,
                    "margin_range":mr,"margin_flag":mf,"booking_tier":"medium",
                    "priority_score":0.0,"uploaded_at":datetime.now().isoformat(),
                })
            except: continue
    return all_rows, debug_info

def save_rows(rows):
    if not rows: return 0, 0
    conn = get_conn(); c = conn.cursor(); n = 0; skipped = 0
    for r in rows:
        r["priority_score"] = calc_priority(r)
        c.execute("""SELECT id FROM pricing
            WHERE file_name=? AND giata=? AND dep_date=? AND dep_window=?""",
            (r["file_name"],r["giata"],r["dep_date"],r["dep_window"]))
        if c.fetchone() is None:
            c.execute("""INSERT INTO pricing(
                file_name,file_date,destination,competitor,giata,hotel_name,board,
                dep_date,dep_month,dep_window,nights,pp_price,current_margin,op_room,
                comp_price,diff_gbp,diff_pct,result,margin_after,margin_range,
                margin_flag,booking_tier,priority_score,uploaded_at)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                tuple(r[k] for k in [
                    "file_name","file_date","destination","competitor","giata","hotel_name","board",
                    "dep_date","dep_month","dep_window","nights","pp_price","current_margin","op_room",
                    "comp_price","diff_gbp","diff_pct","result","margin_after","margin_range",
                    "margin_flag","booking_tier","priority_score","uploaded_at"]))
            n += 1
        else: skipped += 1
    conn.commit(); conn.close()
    return n, skipped

def load_data(dest=None, comp=None, window=None, files=None):
    conn = get_conn()
    q = "SELECT * FROM pricing WHERE 1=1"; p = []
    if dest:   q+=f" AND destination IN ({','.join(['?']*len(dest))})";  p+=dest
    if comp:   q+=f" AND competitor IN ({','.join(['?']*len(comp))})";   p+=comp
    if window: q+=f" AND dep_window IN ({','.join(['?']*len(window))})"; p+=window
    if files:  q+=f" AND file_name IN ({','.join(['?']*len(files))})";   p+=files
    df = pd.read_sql_query(q, conn, params=p)
    conn.close()
    if df.empty: return df
    for col in ["pp_price","comp_price","current_margin","margin_after","diff_pct","diff_gbp","priority_score"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    df["dep_date_dt"] = pd.to_datetime(df["dep_date"].astype(str), errors="coerce")
    df["dep_month"] = df["dep_date_dt"].apply(lambda d: d.strftime("%Y-%m") if pd.notna(d) else "")
    df = df[df["pp_price"] > 0]
    return df

def load_bookings():
    conn = get_conn()
    df = pd.read_sql_query("SELECT * FROM bookings", conn)
    conn.close()
    return df

def norm_giata(x):
    """Normalise a giata id to a plain digit string for matching across files."""
    s = str(x).strip()
    if s.endswith(".0"): s = s[:-2]
    return s

def load_bkg_map():
    """Return (dict giata->bookings, dict giata->revenue, dataframe)."""
    conn = get_conn()
    try:
        b = pd.read_sql_query("SELECT * FROM bkg", conn)
    except Exception:
        b = pd.DataFrame()
    conn.close()
    if b.empty:
        return {}, {}, b
    b["giata"] = b["giata"].map(norm_giata)
    bk_map  = dict(zip(b["giata"], pd.to_numeric(b["bookings"], errors="coerce").fillna(0).astype(int)))
    rev_map = dict(zip(b["giata"], pd.to_numeric(b["revenue"],  errors="coerce").fillna(0.0)))
    return bk_map, rev_map, b

def dest_bookings(frame, bk_map):
    """Total bookings per destination, summing each giata ONCE (not per pricing row)."""
    if frame.empty or not bk_map: return {}
    uniq = frame.drop_duplicates("giata")[["giata","destination"]].copy()
    uniq["b"] = uniq["giata"].map(lambda g: bk_map.get(norm_giata(g), 0))
    return uniq.groupby("destination")["b"].sum().to_dict()

def load_bkg_tx():
    """Booking time-series: giata × booked_month × dep_month → bookings, revenue."""
    conn = get_conn()
    try: t = pd.read_sql_query("SELECT * FROM bkg_tx", conn)
    except Exception: t = pd.DataFrame()
    conn.close()
    if not t.empty:
        t["giata"] = t["giata"].map(norm_giata)
    return t

def bookings_by_month(tx, giatas, dim):
    """Bookings+revenue by month for a set of giatas. dim = 'dep_month' or 'booked_month'."""
    if tx.empty: return pd.DataFrame(columns=["Month","Bookings","Revenue"])
    sub = tx[tx["giata"].isin({norm_giata(g) for g in giatas})] if giatas is not None else tx
    sub = sub[sub[dim].notna() & (sub[dim].astype(str).str.len()==7)]
    if sub.empty: return pd.DataFrame(columns=["Month","Bookings","Revenue"])
    out = (sub.groupby(dim)
              .agg(Bookings=("bookings","sum"), Revenue=("revenue","sum"))
              .reset_index().rename(columns={dim:"Month"}).sort_values("Month"))
    out["Label"] = out["Month"].apply(lambda m: pd.to_datetime(m+"-01").strftime("%b %Y"))
    return out

def dest_month_grouped(v, tx, bar_metric, line_metric, max_dests=6):
    """Grouped columns by destination (one colour per destination, several columns per
    travel month) for `bar_metric`, with an optional dotted line per destination for
    `line_metric` on a secondary axis. Mirrors the dual-axis bars+line layout.
    Returns (figure, destinations_shown, total_destinations) or None."""
    comp = v[v["result"].isin(["Win","Lose"])].copy()
    comp["current_margin"] = pd.to_numeric(comp["current_margin"], errors="coerce")
    comp = comp[comp["dep_month"].astype(str).str.len()==7]
    if comp.empty: return None
    g = comp.groupby(["destination","dep_month"]).agg(
        wins=("result", lambda s:(s=="Win").sum()), n=("result","size"),
        markup=("current_margin","mean")).reset_index()
    g["winrate"] = g["wins"]/g["n"]*100
    base = g.set_index(["destination","dep_month"])

    bk = None
    if tx is not None and not tx.empty:
        g2d = dict(zip(v["giata"].map(norm_giata), v["destination"]))
        t = tx.copy(); t["destination"] = t["giata"].map(lambda x: g2d.get(norm_giata(x)))
        t = t[t["destination"].notna() & (t["dep_month"].astype(str).str.len()==7)]
        if not t.empty:
            bk = t.groupby(["destination","dep_month"])["bookings"].sum()

    order  = comp["destination"].value_counts().index.tolist()
    dests  = order[:max_dests]
    months = sorted(set(g["dep_month"]) | (set(i[1] for i in bk.index) if bk is not None else set()))
    if not months: return None
    labels = [pd.to_datetime(m+"-01").strftime("%b %Y") for m in months]
    palette = ["#1B6FD4","#0A7C4E","#D4AF37","#C0392B","#7D3C98","#16A085","#E67E22","#2C3E50"]

    def series(metric, d):
        col = {"Win Rate %":"winrate","Avg Markup %":"markup"}.get(metric)
        out = []
        for m in months:
            if metric=="Bookings":
                out.append(bk.get((d,m)) if bk is not None else None)
            else:
                out.append(base[col].get((d,m)))
        return out

    AX = {"Bookings":(None,"Bookings"), "Win Rate %":([0,100],"Win Rate %"),
          "Avg Markup %":(None,"Markup %")}
    fig = go.Figure()
    for i,d in enumerate(dests):
        c = palette[i % len(palette)]
        fig.add_trace(go.Bar(x=labels, y=series(bar_metric,d), name=d, legendgroup=d,
                             marker_color=c, offsetgroup=str(i)))
    has_line = line_metric and line_metric not in ("None", bar_metric)
    if has_line:
        for i,d in enumerate(dests):
            c = palette[i % len(palette)]
            fig.add_trace(go.Scatter(x=labels, y=series(line_metric,d), name=d, legendgroup=d,
                showlegend=False, mode="lines+markers", yaxis="y2",
                line=dict(color=c, width=2, dash="dot"), marker=dict(size=5), connectgaps=False))
    if bar_metric=="Win Rate %":
        fig.add_hline(y=50, line_dash="dot", line_color="#E04A3F", annotation_text="50%")
    if bar_metric=="Avg Markup %":
        fig.add_hline(y=MARGIN_FLOOR,   line_dash="dot", line_color="#C0392B", annotation_text=f"{MARGIN_FLOOR:.0f}%")
        fig.add_hline(y=MARGIN_CEILING, line_dash="dot", line_color="#0A7C4E", annotation_text=f"{MARGIN_CEILING:.0f}%")
    layout = dict(barmode="group", height=470, hovermode="x unified",
        xaxis_title="Departure (Travel) Month",
        yaxis=dict(title=AX[bar_metric][1], range=AX[bar_metric][0]),
        legend=dict(orientation="h", yanchor="bottom", y=1.03))
    if has_line:
        layout["yaxis2"] = dict(title=AX[line_metric][1], range=AX[line_metric][0],
            overlaying="y", side="right", showgrid=False)
    fig.update_layout(**layout)
    return fig, dests, len(order)

def dest_month_panels(v, tx, max_dests=8):
    """One figure, travel-month x-axis, with stacked panels for Bookings, Win Rate %
    and Avg Markup % — one coloured line per destination (shown separately), sharing a
    legend. Returns (figure, destinations_shown, total_destinations) or None."""
    comp = v[v["result"].isin(["Win","Lose"])].copy()
    comp["current_margin"] = pd.to_numeric(comp["current_margin"], errors="coerce")
    comp = comp[comp["dep_month"].astype(str).str.len()==7]
    if comp.empty: return None
    grp = comp.groupby(["destination","dep_month"]).agg(
        wins=("result", lambda s:(s=="Win").sum()), n=("result","size"),
        markup=("current_margin","mean")).reset_index()
    grp["winrate"] = grp["wins"]/grp["n"]*100

    bk = None
    if tx is not None and not tx.empty:
        g2d = dict(zip(v["giata"].map(norm_giata), v["destination"]))
        t = tx.copy(); t["destination"] = t["giata"].map(lambda g: g2d.get(norm_giata(g)))
        t = t[t["destination"].notna() & (t["dep_month"].astype(str).str.len()==7)]
        if not t.empty:
            bk = t.groupby(["destination","dep_month"])["bookings"].sum().reset_index()

    order  = comp["destination"].value_counts().index.tolist()
    dests  = order[:max_dests]
    months = sorted(set(grp["dep_month"]) | (set(bk["dep_month"]) if bk is not None else set()))
    if not months: return None
    labels = [pd.to_datetime(m+"-01").strftime("%b %Y") for m in months]
    palette = ["#1B6FD4","#0A7C4E","#D4AF37","#C0392B","#7D3C98","#16A085","#E67E22","#2C3E50"]

    metrics = (["bookings"] if bk is not None else []) + ["winrate","markup"]
    titles  = {"bookings":"📊 Bookings","winrate":"Win Rate %","markup":"Avg Markup %"}
    fig = make_subplots(rows=len(metrics), cols=1, shared_xaxes=True, vertical_spacing=0.08,
        subplot_titles=[titles[m] for m in metrics])

    for i, d in enumerate(dests):
        c  = palette[i % len(palette)]
        gd = grp[grp["destination"]==d].set_index("dep_month")
        bd = bk[bk["destination"]==d].set_index("dep_month")["bookings"] if bk is not None else None
        first = True
        for mi, metric in enumerate(metrics):
            if metric=="bookings":
                y = [bd.get(m) if bd is not None else None for m in months]
            elif metric=="winrate":
                y = [gd["winrate"].get(m) for m in months]
            else:
                y = [gd["markup"].get(m) for m in months]
            fig.add_trace(go.Scatter(x=labels, y=y, name=d, legendgroup=d, showlegend=first,
                mode="lines+markers", line=dict(color=c), marker=dict(size=5),
                connectgaps=False), row=mi+1, col=1)
            first = False

    rr = {m:i+1 for i,m in enumerate(metrics)}
    fig.add_hline(y=50,             row=rr["winrate"], col=1, line_dash="dot", line_color="#888")
    fig.add_hline(y=MARGIN_FLOOR,   row=rr["markup"],  col=1, line_dash="dot", line_color="#C0392B")
    fig.add_hline(y=MARGIN_CEILING, row=rr["markup"],  col=1, line_dash="dot", line_color="#0A7C4E")
    if "bookings" in rr: fig.update_yaxes(title_text="Bookings", row=rr["bookings"], col=1)
    fig.update_yaxes(title_text="Win %", range=[0,100], row=rr["winrate"], col=1)
    fig.update_yaxes(title_text="Markup %", row=rr["markup"], col=1)
    fig.update_layout(height=200*len(metrics)+120, hovermode="x unified",
        margin=dict(t=70,b=10), legend=dict(orientation="h", yanchor="bottom", y=1.03))
    return fig, dests, len(order)

# ── MARKUP RULE ENGINE ──────────────────────────────────────────────────────────
RULE_LAYER = {"0":"System cost","5":"System cost",
              "1":"Consumer (Manual)","3":"Consumer (Exclusive)"}

def parse_rule_row(rid, rulejson, enabled, markup_type, applies_to, generated_text):
    """Flatten one rule-engine row into queryable fields."""
    try: j = json.loads(rulejson)
    except Exception: j = {}
    props = j.get("RuleProperties", []) or []
    a = j.get("Action", {}) or {}
    giatas, users, ispkg = [], [], None
    for p in props:
        t = p.get("RulePropertyType")
        if t == "GIata" and p.get("ConditionOperator") in ("Equal","IsOneOf"):
            giatas += [norm_giata(x) for x in str(p.get("ConditionValue","")).split(",") if x.strip()]
        if t == "User": users.append(str(p.get("ConditionValue","")))
        if t == "IsPackage": ispkg = str(p.get("ConditionValue","")).lower()
    verb_m = re.search(r"Then\s+(\w+)", str(generated_text))
    op = a.get("RuleActionPropertyOperator")
    unit = "%" if str(op)=="1" else ("£" if str(op)=="2" else "")
    val = a.get("RuleActionPropertyValue")
    return dict(rid=str(rid), enabled=1 if str(enabled)=="1" else 0,
        markup_type=str(markup_type), applies_to=str(applies_to),
        layer=RULE_LAYER.get(str(markup_type),"Other"),
        verb=verb_m.group(1) if verb_m else "",
        action_value=float(val) if val not in (None,"") else None,
        action_unit=unit, ispkg=ispkg, user_csv=",".join(users),
        giatas=",".join(giatas), giata_n=len(giatas), n_conds=len(props),
        generated_text=str(generated_text))

def load_rules():
    conn = get_conn()
    try: r = pd.read_sql_query("SELECT * FROM rules", conn)
    except Exception: r = pd.DataFrame()
    conn.close()
    return r

def rules_for_giata(rules_df, giata):
    """All enabled rules whose giata list includes this hotel."""
    if rules_df.empty: return rules_df
    g = norm_giata(giata)
    m = rules_df["giatas"].fillna("").apply(lambda s: g in s.split(",") if s else False)
    return rules_df[m & (rules_df["enabled"]==1)]

def recommend_lever(rules_df, giata, delta, is_package=True):
    """Given a suggested Total-Markup change (delta, pp), pick the rule to edit.

    Levers = consumer-layer markups (Manual/Exclusive). Only rules that target THIS
    HOTEL ALONE are safe to edit — i.e. a giata condition listing exactly one hotel.
    A rule whose giata condition is an IsOneOf list of several hotels is SHARED:
    editing it would move every hotel in that list, so it is never proposed as the
    edit. When only shared rules exist, recommend CREATING a single-hotel rule that
    carries the delta (consumer markups stack additively, so a +delta hotel rule
    adjusts just this hotel on top of the shared one). System-cost rules are context
    only. Cleanest single-hotel lever = fewest extra conditions, matching package vs
    hotel-only, package-level preferred for packages, Exclusive over Manual."""
    hits = rules_for_giata(rules_df, giata)
    if hits.empty:
        return {"status":"no_rules"}
    mk = hits[hits["verb"].isin(["MarkupNew","MarkupPerPerson"])].copy()
    # package vs hotel-only compatibility
    def compatible(row):
        if pd.isna(row["ispkg"]) or row["ispkg"] in ("none",""): return True
        if is_package and row["ispkg"]=="false": return False
        if (not is_package) and row["ispkg"]=="true": return False
        return True
    if not mk.empty:
        mk = mk[mk.apply(compatible, axis=1)]
    # No markup rules left for this hotel (e.g. only Delete/board-change rules) → create.
    # NB: indexing an empty frame by an empty mask drops columns, so guard before using them.
    if mk.empty:
        empty = hits.iloc[0:0]
        return {"status":"create","reason":"none","shared_context":empty,"system_context":empty}
    consumer = mk[mk["layer"].str.startswith("Consumer")].copy()
    system   = mk[mk["layer"].str.startswith("System")].copy()
    # Split by reach: single-hotel rules (safe to edit) vs shared multi-hotel rules
    gn = pd.to_numeric(consumer.get("giata_n"), errors="coerce").fillna(0)
    single = consumer[gn == 1].copy()
    shared = consumer[gn > 1].copy()

    if single.empty:
        # Only shared (or no) consumer rules → editing would hit other hotels.
        return {"status":"create",
                "reason":"shared" if not shared.empty else "none",
                "shared_context":shared, "system_context":system}

    # Prefer Exclusive, then fewest conditions, then package-level for packages
    single["excl"] = (single["layer"]=="Consumer (Exclusive)").astype(int)
    single["lvl_pref"] = single["applies_to"].map(
        lambda a: 1 if (is_package and a=="15") or ((not is_package) and a=="12") else 0)
    single = single.sort_values(["excl","lvl_pref","n_conds"], ascending=[False,False,True])
    best = single.iloc[0]
    new_val = None
    if best["action_unit"]=="%" and best["action_value"] is not None:
        new_val = round(best["action_value"] + delta, 2)
    return {"status":"edit","rule":best,"alternatives":single.iloc[1:],
            "shared_context":shared,"system_context":system,"new_value":new_val}

def get_distinct(col):
    conn = get_conn()
    rows = conn.execute(f"SELECT DISTINCT {col} FROM pricing ORDER BY {col}").fetchall()
    conn.close()
    return [r[0] for r in rows if r[0]]

def get_file_list():
    conn = get_conn()
    rows = conn.execute("""SELECT file_name,file_date,destination,competitor,COUNT(*) as rows
        FROM pricing GROUP BY file_name ORDER BY file_date DESC""").fetchall()
    conn.close()
    return rows

def apply_tiers(df, bk_map=None):
    if df.empty: return df
    # Prefer REAL booking volumes (per giata) when a bookings file is loaded;
    # otherwise fall back to comparison-row count as a volume proxy.
    vol = None
    if bk_map:
        v = df["giata"].map(lambda g: bk_map.get(norm_giata(g), 0))
        if v.sum() > 0: vol = v
    if vol is None:
        vol = df.groupby(["hotel_name","destination"])["hotel_name"].transform("count")
    p75, p25 = vol.quantile(0.75), vol.quantile(0.25)
    df["booking_tier"] = np.where(vol>=p75,"high",np.where(vol>=p25,"medium","low"))
    return df

# ══════════════════════════════════════════════════════════════════════════════
# REUSABLE FILTER BAR  — used at top of every tab
# ══════════════════════════════════════════════════════════════════════════════
def tab_filters(df, key_prefix, show_hotel=False):
    """Returns filtered dataframe. Renders dest/comp/window/hotel/board filters."""
    all_dest  = sorted(df["destination"].unique().tolist())
    all_comp  = sorted(df["competitor"].unique().tolist())
    all_win   = ["0-60","61-120","121-240","241+"]

    if show_hotel:
        c1,c2,c3,c4,c5 = st.columns([2,2,2,3,2])
    else:
        c1,c2,c3 = st.columns([2,2,3])

    sel_dest = c1.multiselect("Destination", all_dest, default=all_dest, key=f"{key_prefix}_dest")
    sel_comp = c2.multiselect("Competitor",  all_comp, default=all_comp, key=f"{key_prefix}_comp")
    sel_win  = c3.multiselect("Dep Window",  all_win,  default=all_win,  key=f"{key_prefix}_win")

    filtered = df.copy()
    if sel_dest: filtered = filtered[filtered["destination"].isin(sel_dest)]
    if sel_comp: filtered = filtered[filtered["competitor"].isin(sel_comp)]
    if sel_win:  filtered = filtered[filtered["dep_window"].isin(sel_win)]

    if show_hotel:
        hotel_list = ["All"] + sorted(filtered["hotel_name"].unique().tolist())
        sel_hotel  = c4.selectbox("Hotel", hotel_list, key=f"{key_prefix}_hotel")
        board_list = ["All"] + sorted(filtered["board"].unique().tolist())
        sel_board  = c5.selectbox("Board", board_list, key=f"{key_prefix}_board")
        if sel_hotel != "All": filtered = filtered[filtered["hotel_name"]==sel_hotel]
        if sel_board != "All": filtered = filtered[filtered["board"]==sel_board]

    return filtered

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## ✈ D2 Pricing Intelligence")
    st.markdown("---")
    st.markdown("### 📂 Upload Pricing Files")
    uploaded = st.file_uploader("Drop .xlsm / .xlsx / .csv",
        type=["xlsm","xlsx","csv"], accept_multiple_files=True, key="pu")
    if uploaded:
        for uf in uploaded:
            with st.spinner(f"Parsing {uf.name}…"):
                rows, debug_info = parse_pricing_file(uf, uf.name)
            if not rows:
                st.error(f"❌ {uf.name}: 0 rows parsed")
                continue
            n, skipped = save_rows(rows)
            if n > 0: st.success(f"✅ {uf.name}: {n} rows added ({skipped} skipped)")
            else:     st.info(f"ℹ {uf.name}: already in database ({skipped} rows)")

    st.markdown("### 📊 Upload Bookings")
    bfile = st.file_uploader("Bookings export with a 'Giata' column (one row per booking)",
        type=["csv","xlsx"], key="bu")
    if bfile:
        try:
            bdf = pd.read_csv(bfile) if bfile.name.endswith(".csv") else pd.read_excel(bfile)
            cols = {c.lower().strip(): c for c in bdf.columns}
            gcol = next((cols[k] for k in cols if k == "giata"), None)
            if gcol is not None:
                # Transactional export → aggregate to per-giata volume + revenue
                hcol = next((cols[k] for k in cols if k in ("hotels","hotel","hotel_name")), None)
                dcol = next((cols[k] for k in cols if k == "destination"), None)
                rcol = next((cols[k] for k in cols if k == "revenue"), None)
                bdf["_g"] = bdf[gcol].map(norm_giata)
                bdf = bdf[bdf["_g"].str.len() > 0]
                agg = bdf.groupby("_g").agg(
                    bookings=("_g","size"),
                    revenue=(rcol,"sum") if rcol else ("_g","size"),
                    hotel_name=(hcol,"first") if hcol else ("_g","first"),
                    dest_name=(dcol,"first") if dcol else ("_g","first"),
                ).reset_index()
                conn = get_conn()
                for _, r in agg.iterrows():
                    conn.execute("""INSERT INTO bkg(giata,hotel_name,dest_name,bookings,revenue,uploaded_at)
                        VALUES(?,?,?,?,?,?)
                        ON CONFLICT(giata) DO UPDATE SET
                            hotel_name=excluded.hotel_name, dest_name=excluded.dest_name,
                            bookings=excluded.bookings, revenue=excluded.revenue,
                            uploaded_at=excluded.uploaded_at""",
                        (str(r["_g"]), str(r["hotel_name"]), str(r["dest_name"]),
                         int(r["bookings"]), float(r["revenue"] or 0), datetime.now().isoformat()))
                conn.commit(); conn.close()
                # Time-series: bookings by giata × booked-month × departure-month
                bkcol = next((cols[k] for k in cols if k in ("booked date","booked_date","booked")), None)
                dpcol = next((cols[k] for k in cols if k in ("departure date","departure_date","departure","travel date")), None)
                if bkcol or dpcol:
                    tx = bdf.copy()
                    tx["bm"] = (pd.to_datetime(tx[bkcol], errors="coerce").dt.strftime("%Y-%m")
                                if bkcol else "")
                    tx["dm"] = (pd.to_datetime(tx[dpcol], errors="coerce").dt.strftime("%Y-%m")
                                if dpcol else "")
                    tx["rev"] = pd.to_numeric(tx[rcol], errors="coerce").fillna(0.0) if rcol else 0.0
                    txg = (tx.groupby(["_g","bm","dm"])
                             .agg(bookings=("_g","size"), revenue=("rev","sum")).reset_index())
                    conn = get_conn()
                    txg.rename(columns={"_g":"giata","bm":"booked_month","dm":"dep_month"})\
                       .to_sql("bkg_tx", conn, if_exists="replace", index=False)
                    conn.commit(); conn.close()
                st.success(f"✅ {len(agg):,} hotels · {int(agg['bookings'].sum()):,} bookings matched by giata")
            else:
                # Legacy simple format: hotel_name, destination, bkgs_4wk
                bdf.columns = [c.lower().strip().replace(" ","_") for c in bdf.columns]
                conn = get_conn()
                for _, r in bdf.iterrows():
                    conn.execute("INSERT INTO bookings(hotel_name,destination,bkgs_4wk,bkgs_py,uploaded_at) VALUES(?,?,?,?,?)",
                        (str(r.get("hotel_name","")),str(r.get("destination","")),
                         int(r.get("bkgs_4wk",0) or 0),int(r.get("bkgs_py",0) or 0),
                         datetime.now().isoformat()))
                conn.commit(); conn.close()
                st.success(f"✅ {len(bdf)} hotels loaded (legacy format)")
        except Exception as e: st.error(f"Bookings error: {e}")

    st.markdown("### 🎚 Upload Markup Rules")
    rfile = st.file_uploader("Rule-engine export (RuleJson + GeneratedText)",
        type=["csv","xlsx"], key="ru")
    if rfile:
        try:
            rdf = pd.read_csv(rfile) if rfile.name.endswith(".csv") else pd.read_excel(rfile, dtype=str)
            cmap = {c.lower().strip(): c for c in rdf.columns}
            need = ["rulejson","enabled","markuptype","ruleappliesto","generatedtext"]
            if all(any(k==cc for cc in cmap) for k in need) or "rulejson" in cmap:
                idc = next((cmap[k] for k in cmap if k=="id"), None)
                parsed = [parse_rule_row(
                            rdf.iloc[i][idc] if idc else i,
                            rdf.iloc[i].get(cmap.get("rulejson"),"{}"),
                            rdf.iloc[i].get(cmap.get("enabled"),"1"),
                            rdf.iloc[i].get(cmap.get("markuptype"),""),
                            rdf.iloc[i].get(cmap.get("ruleappliesto"),""),
                            rdf.iloc[i].get(cmap.get("generatedtext"),""))
                          for i in range(len(rdf))]
                pr = pd.DataFrame(parsed); pr["uploaded_at"] = datetime.now().isoformat()
                conn = get_conn()
                pr.to_sql("rules", conn, if_exists="replace", index=False)
                conn.commit(); conn.close()
                nlev = int((pr["enabled"]==1).sum())
                st.success(f"✅ {len(pr):,} rules parsed · {nlev:,} enabled · "
                           f"{pr['giatas'].apply(lambda s: bool(s)).sum():,} hotel-specific")
            else:
                st.error("Need a 'RuleJson' column (and ideally Enabled, MarkupType, RuleAppliesTo, GeneratedText)")
        except Exception as e: st.error(f"Rules error: {e}")

    st.markdown("---")
    st.markdown("### 🔽 Global Filters")
    all_d = get_distinct("destination")
    all_c = get_distinct("competitor")
    all_f = [r[0] for r in get_file_list()]
    sel_d = st.multiselect("Destination", all_d, default=all_d)
    sel_c = st.multiselect("Competitor",  all_c, default=all_c)
    sel_w = st.multiselect("Dep Window",  ["0-60","61-120","121-240","241+"],
                            default=["0-60","61-120","121-240","241+"])
    sel_f = st.multiselect("Files", all_f, default=all_f)
    st.markdown("---")
    if st.button("🗑 Clear ALL data"):
        conn = get_conn()
        conn.execute("DELETE FROM pricing"); conn.execute("DELETE FROM bookings")
        conn.execute("DELETE FROM bkg")
        conn.execute("DELETE FROM bkg_tx")
        conn.execute("DELETE FROM rules")
        conn.commit(); conn.close(); st.success("Cleared"); st.rerun()
    if st.button("🔴 FORCE RESET (drop & recreate DB)"):
        try:
            conn = get_conn()
            conn.execute("DROP TABLE IF EXISTS pricing")
            conn.execute("DROP TABLE IF EXISTS bookings")
            conn.commit(); conn.close()
        except: pass
        try: os.remove(DB_PATH)
        except: pass
        with open("reset_flag.txt","w") as f: f.write("reset")
        init_db()
        st.success("✅ Reset complete — re-upload files"); st.rerun()
    st.markdown("---")
    st.markdown("**🔍 DB Status**")
    try:
        conn = get_conn()
        rc   = conn.execute("SELECT COUNT(*) FROM pricing").fetchone()[0]
        pps  = conn.execute("SELECT pp_price FROM pricing LIMIT 3").fetchall()
        dps  = conn.execute("SELECT dep_date FROM pricing LIMIT 3").fetchall()
        dsts = conn.execute("SELECT DISTINCT destination FROM pricing").fetchall()
        conn.close()
        st.write(f"Rows: {rc}")
        st.write(f"pp_price: {[round(r[0],0) for r in pps]}")
        st.write(f"dep_date: {[r[0] for r in dps]}")
        st.write(f"Destinations: {[r[0] for r in dsts]}")
    except Exception as e: st.write(f"DB error: {e}")

# ══════════════════════════════════════════════════════════════════════════════
# LOAD DATA
# ══════════════════════════════════════════════════════════════════════════════
df = load_data(dest=sel_d or None, comp=sel_c or None,
               window=sel_w or None, files=sel_f or None)
bk = load_bookings()
bk_map, rev_map, bk_df = load_bkg_map()
HAS_BKG = bool(bk_map)
bk_tx = load_bkg_tx()
HAS_TX = not bk_tx.empty
rules_df = load_rules()
HAS_RULES = not rules_df.empty
if not df.empty:
    df = apply_tiers(df, bk_map)
    df["priority_score"] = df.apply(lambda r: calc_priority(r.to_dict()), axis=1)
    # Per-row hotel booking volume + revenue (constant per giata) for tables
    df["bookings"] = df["giata"].map(lambda g: bk_map.get(norm_giata(g), 0)).astype(int)
    df["revenue"]  = df["giata"].map(lambda g: rev_map.get(norm_giata(g), 0.0))

comp_df  = df[df["result"].isin(["Win","Lose"])]            if not df.empty else pd.DataFrame()
wins_df  = df[df["result"]=="Win"]                           if not df.empty else pd.DataFrame()
loses_df = df[df["result"]=="Lose"]                          if not df.empty else pd.DataFrame()
nc_df    = df[df["result"].str.contains("No Comp",na=False)] if not df.empty else pd.DataFrame()

# Hotels that have a real comparison (Win/Lose) on ANY date — used to scope the No Comp
# tab so it lists ONLY hotels never compared across all dates in the data.
compared_giatas = set(comp_df["giata"].unique()) if not comp_df.empty else set()

st.title("✈ D2 Pricing Intelligence Dashboard")
if not df.empty:
    st.caption(f"Destinations: **{', '.join(sel_d)}** · Competitor: **{', '.join(sel_c or [])}** · {len(sel_f)} files · {len(df):,} rows")

if df.empty:
    st.info("👆 Upload pricing files using the sidebar to get started.")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
tabs = st.tabs(["⚡ Master Queue","Overview","Hotel Actions",
                "↓ Reduce","↑ Raise","Maintain",
                "Product Gaps","No Comp","Dest Actions",
                "⚑ Outliers","Missing","Advantage",
                "Insights","↗ Price Trends","📂 Files","🎚 Markup Levers",
                "📖 How It Works"])

# ── SHARED COLS helper ─────────────────────────────────────────────────────────
HOTEL_COLS     = ["giata","hotel_name","destination","competitor"]
HOTEL_COL_LBLS = ["Giata","Hotel","Dest","Comp"]

def with_bkg(show_df, loc=2):
    """Insert a per-hotel '📊 Bkgs' column into any table that has a 'Giata' column.
    No-op when no booking file is loaded."""
    if HAS_BKG and "Giata" in show_df.columns and "📊 Bkgs" not in show_df.columns:
        b = show_df["Giata"].map(lambda g: bk_map.get(norm_giata(g), 0))
        show_df.insert(min(loc, show_df.shape[1]), "📊 Bkgs", b.map("{:,.0f}".format))
    return show_df

# ── 1. MASTER QUEUE ────────────────────────────────────────────────────────────
with tabs[0]:
    st.subheader("Master Action Queue — Top 30")
    v = tab_filters(df, "maq")
    v_loses = v[v["result"]=="Lose"]
    v_wins  = v[v["result"]=="Win"]
    acts = []
    for _, r in v_loses.iterrows():
        gap = abs(r["diff_pct"])
        act = ("⚠ SUPPRESS — floor breach" if r["margin_flag"]=="floor"
               else f"↓ Reduce ~£{abs(r['comp_price']-r['pp_price']):.0f}pp" if gap>5
               else "↓ Monitor")
        acts.append({**r.to_dict(),"action":act})
    for _, r in v_wins[v_wins["diff_pct"]>RAISE_THRESHOLD].iterrows():
        acts.append({**r.to_dict(),"action":"↑ Margin opportunity"})
    if acts:
        adf = pd.DataFrame(acts)
        adf["_tw"] = adf["booking_tier"].map({"high":3,"medium":2,"low":1}).fillna(1)
        adf = adf.sort_values(["_tw","priority_score","diff_pct"],ascending=[False,False,True]).head(30)
        adf["Margin"] = adf.apply(lambda r: margin_cell(r["current_margin"],r["margin_after"]),axis=1)
        show = adf[["giata","hotel_name","destination","competitor","dep_window","dep_month",
                     "action","pp_price","comp_price","diff_pct","Margin","booking_tier"]].copy()
        show.columns = ["Giata","Hotel","Dest","Comp","Window","Month","Action",
                        "D2 £pp","Comp £pp","Diff %","Margin","Tier"]
        st.dataframe(with_bkg(show).style.format({"D2 £pp":"{:.0f}","Comp £pp":"{:.0f}","Diff %":"{:.1f}%"}),
                     use_container_width=True, height=550)
    else:
        st.info("No actions with current filters.")

# ── 2. OVERVIEW ────────────────────────────────────────────────────────────────
with tabs[1]:
    v = tab_filters(df, "ovw")
    vc = v[v["result"].isin(["Win","Lose"])]
    wn,ln,ncn = (v["result"]=="Win").sum(),(v["result"]=="Lose").sum(),(v["result"].str.contains("No Comp",na=False)).sum()
    total = wn+ln
    wr  = wn/total*100 if total>0 else 0
    alg = v[v["result"]=="Lose"]["diff_pct"].mean() if ln>0 else 0
    fln = (v[v["result"]=="Lose"]["margin_flag"]=="floor").sum()
    tot_bkg = int(v.drop_duplicates("giata")["bookings"].sum()) if HAS_BKG else 0
    if HAS_BKG:
        c1,c2,c3,c4,c5,c6,c7 = st.columns(7)
    else:
        c1,c2,c3,c4,c5,c6 = st.columns(6)
    c1.metric("Hotels",    v["hotel_name"].nunique())
    c2.metric("Win Rate",  f"{wr:.0f}%", f"{wn}W/{ln}L")
    c3.metric("Losses",    ln, f"Avg {alg:.1f}%")
    c4.metric("No Comp",   ncn)
    c5.metric("⚠ Floor",   fln)
    c6.metric("Avg Margin",f"{v['current_margin'].mean():.1f}%")
    if HAS_BKG:
        c7.metric("📊 Bookings", f"{tot_bkg:,}")
    col1,col2 = st.columns(2)
    with col1:
        wl = v.groupby(["dep_window","result"]).size().reset_index(name="n")
        fig = px.bar(wl,x="dep_window",y="n",color="result",
            color_discrete_map={"Win":"#0A7C4E","Lose":"#C0392B","Win - No Comp":"#D4AF37"},
            title="Results by Departure Window",
            category_orders={"dep_window":["0-60","61-120","121-240","241+"]})
        fig.update_layout(height=300,margin=dict(t=40,b=10))
        st.plotly_chart(fig,use_container_width=True)
    with col2:
        if not vc.empty:
            ds = vc.groupby("destination").apply(
                lambda g:(g["result"]=="Win").sum()/len(g)*100).reset_index()
            ds.columns = ["Dest","WinRate"]
            db = dest_bookings(v, bk_map)
            ds["Bookings"] = ds["Dest"].map(lambda d: db.get(d,0))
            fig2 = px.bar(ds,x="Dest",y="WinRate",title="Win Rate by Destination (%)",
                color="WinRate",color_continuous_scale=["#C0392B","#D4AF37","#0A7C4E"],range_color=[30,80],
                custom_data=["Bookings"] if HAS_BKG else None)
            if HAS_BKG:
                fig2.update_traces(hovertemplate="<b>%{x}</b><br>Win Rate %{y:.0f}%<br>Bookings %{customdata[0]:,}<extra></extra>")
            fig2.update_layout(height=300,margin=dict(t=40,b=10))
            st.plotly_chart(fig2,use_container_width=True)

    # Bookings by Destination — volume context for the win-rate picture above
    if HAS_BKG:
        db = dest_bookings(v, bk_map)
        if db:
            bdest = pd.DataFrame({"Dest":list(db.keys()),"Bookings":list(db.values())})
            bdest = bdest[bdest["Bookings"]>0].sort_values("Bookings",ascending=False)
            if not bdest.empty:
                figb = px.bar(bdest,x="Dest",y="Bookings",title="📊 Bookings by Destination",
                              color="Bookings",color_continuous_scale=["#1B6FD4","#1B1464"],
                              text="Bookings")
                figb.update_traces(textposition="outside")
                figb.update_layout(height=300,margin=dict(t=40,b=10),coloraxis_showscale=False)
                st.plotly_chart(figb,use_container_width=True)

    # By travel month & destination — columns per destination, bars + optional line
    comp_any = v[v["result"].isin(["Win","Lose"])]
    if not comp_any.empty and comp_any["dep_month"].astype(str).str.len().eq(7).any():
        st.markdown("#### 📈 By Travel Month & Destination")
        bar_opts  = (["Bookings"] if HAS_TX else []) + ["Win Rate %","Avg Markup %"]
        line_opts = ["None"] + (["Bookings"] if HAS_TX else []) + ["Win Rate %","Avg Markup %"]
        cc1, cc2 = st.columns(2)
        bar_metric  = cc1.radio("Columns show", bar_opts, index=0, horizontal=True, key="dmg_bar")
        line_metric = cc2.radio("Overlay line", line_opts,
                                 index=line_opts.index("Win Rate %"), horizontal=True, key="dmg_line")
        res = dest_month_grouped(v, bk_tx if HAS_TX else None, bar_metric, line_metric)
        if res:
            figdm, dshown, ntot = res
            cap = (f"Each coloured column is a destination, grouped by travel (departure) month. "
                   f"Bars = **{bar_metric}**"
                   + (f"; dotted line = **{line_metric}** (right axis)."
                      if line_metric not in ("None", bar_metric) else "."))
            if ntot > len(dshown):
                cap += f"  Showing the top {len(dshown)} of {ntot} destinations by data volume."
            st.caption(cap)
            st.plotly_chart(figdm, use_container_width=True)

    st.markdown("#### 🏆 Top Sellers")
    st.caption("**Sugg. Markup Δ** = high-level change to Total Markup (in pp) per hotel, "
               "across the latest comparisons over all departure dates. It maximises expected "
               "booked margin (projected win rate × projected markup) while keeping every row "
               "inside the 7% floor / 15% ceiling and giving away the least margin possible. "
               "Negative = cut to convert losses; positive = headroom to raise where we already win.")
    ts = v.groupby(["giata","hotel_name"]).agg(
        Dest=("destination","first"),Comp=("competitor","first"),
        Rows=("hotel_name","count"),AvgMargin=("current_margin","mean"),
        Bookings=("bookings","max"),Revenue=("revenue","max"),
        HasLoss=("result",lambda x:"⚠ Loss" if (x=="Lose").any() else "✓ OK")
    ).reset_index()
    # Rank by REAL bookings when available, else by comparison rows
    ts = ts.sort_values("Bookings" if HAS_BKG else "Rows",ascending=False).head(15).reset_index(drop=True)

    # Suggested markup change per hotel (across all its dates, latest run)
    sug = [suggest_markup_change(
              v[(v["giata"]==g) & (v["hotel_name"]==hn)])
           for g, hn in zip(ts["giata"], ts["hotel_name"])]
    sdf = pd.DataFrame(sug)
    ts["SuggDelta"] = sdf["delta"].apply(lambda d: f"{d:+.1f}pp" if abs(d)>=0.05 else "0.0pp")
    ts["ProjWR"]    = sdf.apply(lambda r: f"{r['wr0']:.0f}%→{r['wr1']:.0f}%"
                                if pd.notna(r["wr0"]) else "—", axis=1)
    ts["ProjMargin"]= sdf.apply(lambda r: f"{r['m0']:.1f}%→{r['m1']:.1f}%", axis=1)
    ts["Note"]      = sdf["note"]

    ts["AvgMargin"] = ts["AvgMargin"].map("{:.1f}%".format)
    if HAS_BKG:
        ts["Bookings"] = ts["Bookings"].map("{:,.0f}".format)
        ts["Revenue"]  = ts["Revenue"].map("£{:,.0f}".format)
        ts = ts[["giata","hotel_name","Dest","Comp","Bookings","Revenue","Rows","AvgMargin","HasLoss",
                 "SuggDelta","ProjWR","ProjMargin","Note"]]
        ts.columns = ["Giata","Hotel","Dest","Comp","📊 Bookings","Revenue","Rows","Avg Margin","Loss?",
                      "Sugg. Markup Δ","Proj. Win Rate","Proj. Markup","Note"]
    else:
        ts = ts[["giata","hotel_name","Dest","Comp","Rows","AvgMargin","HasLoss",
                 "SuggDelta","ProjWR","ProjMargin","Note"]]
        ts.columns = ["Giata","Hotel","Dest","Comp","Rows","Avg Margin","Loss?",
                      "Sugg. Markup Δ","Proj. Win Rate","Proj. Markup","Note"]
    st.dataframe(ts,use_container_width=True)

# ── 3. HOTEL ACTIONS ───────────────────────────────────────────────────────────
with tabs[2]:
    st.subheader("Individual Hotel Actions")
    v = tab_filters(df, "ha")
    c1,c2 = st.columns([3,2])
    srch = c1.text_input("Search hotel",placeholder="e.g. Fushifaru…",key="s3")
    rf   = c2.selectbox("Result",["All","Win","Lose","Win - No Comp"],key="r3")
    if srch: v = v[v["hotel_name"].str.contains(srch,case=False,na=False)]
    if rf!="All": v = v[v["result"]==rf]
    v["_tw"] = v["booking_tier"].map({"high":3,"medium":2,"low":1}).fillna(1)
    v = v.sort_values(["_tw","priority_score"],ascending=False)
    def make_action(r):
        if r["result"]=="Lose": return "↓ Reduce" if abs(r["diff_pct"])>5 else "Monitor"
        if r["result"]=="Win" and r["diff_pct"]>RAISE_THRESHOLD: return "↑ Raise"
        if r["result"]=="Win": return "Maintain"
        return "Dest-guided"
    v["Action"] = v.apply(make_action,axis=1)
    v["Margin"] = v.apply(lambda r: margin_cell(r["current_margin"],r["margin_after"]),axis=1)
    show3 = v[["giata","hotel_name","destination","competitor","dep_window","dep_month","board",
               "pp_price","comp_price","diff_pct","result","Margin","Action","booking_tier"]].copy()
    show3.columns = ["Giata","Hotel","Dest","Comp","Window","Month","Board",
                     "D2 £pp","Comp £pp","Diff %","Result","Margin","Action","Tier"]
    st.dataframe(with_bkg(show3).style.format({"D2 £pp":"{:.0f}","Comp £pp":"{:.0f}","Diff %":"{:.1f}%"}),
                 use_container_width=True,height=600)

# ── 4. REDUCE ──────────────────────────────────────────────────────────────────
with tabs[3]:
    st.subheader("↓ Price Reduction Candidates")
    v = tab_filters(df, "red")
    cands = v[v["result"]=="Lose"].copy()
    fln2 = (cands["margin_flag"]=="floor").sum()
    nrn  = (cands["margin_flag"]=="near").sum()
    st.error(f"⚠ {fln2} of {len(cands)} reductions breach 7% floor. {nrn} enter 7–8.5% caution band.")
    srch4 = st.text_input("Search hotel",key="s4")
    if srch4: cands=cands[cands["hotel_name"].str.contains(srch4,case=False,na=False)]
    cands = cands.sort_values(["booking_tier","diff_pct"],ascending=[False,True])
    cands["Margin"] = cands.apply(lambda r: margin_cell(r["current_margin"],r["margin_after"]),axis=1)
    cands["Flag"]   = cands["margin_flag"].map({"floor":"🔴 SUPPRESS","near":"🟡 Near Floor",
                                                 "ok":"✅ OK","ceiling":"↑ Ceiling","unknown":"?"})
    cands["Rec"]    = cands.apply(lambda r:
        "SUPPRESS — floor breach" if r["margin_flag"]=="floor"
        else f"↓ Reduce ~£{abs(r['comp_price']-r['pp_price']):.0f}pp",axis=1)
    show4 = cands[["giata","hotel_name","destination","competitor","dep_window","dep_month",
                    "pp_price","comp_price","diff_pct","Margin","Flag","Rec","booking_tier"]].copy()
    show4.columns = ["Giata","Hotel","Dest","Comp","Window","Month",
                     "D2 £pp","Comp £pp","Diff %","Margin","Flag","Action","Tier"]
    st.dataframe(with_bkg(show4).style.format({"D2 £pp":"{:.0f}","Comp £pp":"{:.0f}","Diff %":"{:.1f}%"}),
                 use_container_width=True,height=550)

# ── 5. RAISE ───────────────────────────────────────────────────────────────────
with tabs[4]:
    st.subheader("↑ Raise / Margin Opportunities")
    v = tab_filters(df, "rse")
    rdf = v[v["result"]=="Win"][v[v["result"]=="Win"]["diff_pct"]>RAISE_THRESHOLD].copy()
    st.success(f"✅ {len(rdf)} hotels winning by >15%")
    rdf = rdf.sort_values(["booking_tier","diff_pct"],ascending=[False,False])
    rdf["Margin"]  = rdf.apply(lambda r: margin_cell(r["current_margin"],r["margin_after"]),axis=1)
    rdf["Ceiling"] = rdf["margin_flag"].map({"ceiling":"🔵 ABOVE","ok":"✅ OK",
                                              "near":"🟡 Near","floor":"🔴 Floor","unknown":"?"})
    show5 = rdf[["giata","hotel_name","destination","competitor","dep_window","dep_month",
                  "pp_price","comp_price","diff_pct","Margin","Ceiling","booking_tier"]].copy()
    show5.columns = ["Giata","Hotel","Dest","Comp","Window","Month",
                     "D2 £pp","Comp £pp","Win %","Margin","Ceiling","Tier"]
    st.dataframe(with_bkg(show5).style.format({"D2 £pp":"{:.0f}","Comp £pp":"{:.0f}","Win %":"{:.1f}%"}),
                 use_container_width=True,height=500)

# ── 6. MAINTAIN ────────────────────────────────────────────────────────────────
with tabs[5]:
    st.subheader("Maintain — Hold Band (−5% to +8%)")
    v = tab_filters(df, "mtn")
    mtn = v[(v["result"]=="Win")&(v["diff_pct"]>=-5)&(v["diff_pct"]<=HOLD_UPPER)].copy()
    mtn["MFlag"] = mtn["margin_flag"].map({"floor":"🔴 FLOOR","near":"🟡 Near",
                                            "ok":"✅ OK","ceiling":"↑ Ceiling","unknown":"?"})
    show6 = mtn[["giata","hotel_name","destination","competitor","dep_window","dep_month",
                  "pp_price","comp_price","diff_pct","current_margin","MFlag","booking_tier"]].copy()
    show6.columns = ["Giata","Hotel","Dest","Comp","Window","Month",
                     "D2 £pp","Comp £pp","Diff %","Margin %","Flag","Tier"]
    st.dataframe(with_bkg(show6).style.format({"D2 £pp":"{:.0f}","Comp £pp":"{:.0f}",
                                      "Diff %":"{:.1f}%","Margin %":"{:.1f}%"}),
                 use_container_width=True,height=500)

# ── 7. PRODUCT GAPS ────────────────────────────────────────────────────────────
with tabs[6]:
    st.subheader("Product Gaps")
    st.warning("⚠ NEVER price cut — product/contracting fix required.")
    st.caption(f"A hotel is flagged a **product gap** when a high share of its losses can't be "
               f"matched without breaching the {MARGIN_FLOOR:.0f}% margin floor — i.e. to beat the "
               f"competitor we'd have to price below our minimum margin. That's unreachable on price, "
               f"so it's a product or contracting issue, not a markup one. "
               f"(Threshold: ≥{PRODUCT_GAP_SHARE*100:.0f}% of losses below floor, min {PRODUCT_GAP_MIN}.)")
    v = tab_filters(df, "pg")
    losses = v[v["result"]=="Lose"].copy()
    if losses.empty:
        st.info("No losses in the current selection.")
    else:
        # A loss is "below floor" when matching the competitor would need a markup under the
        # floor (margin_after < floor). Guard against missing margin_after (stored as 0).
        losses["below_floor"] = (losses["margin_after"] > 0) & (losses["margin_after"] < MARGIN_FLOOR)
        g = losses.groupby(["giata","hotel_name","destination","competitor"])
        summ = g.agg(Loses=("result","size"),
                     BelowFloor=("below_floor","sum"),
                     AvgGap=("diff_pct","mean"),
                     AvgMatchMarkup=("margin_after","mean")).reset_index()
        summ["Share"] = summ["BelowFloor"] / summ["Loses"]
        summ["ProductGap"] = (summ["Share"] >= PRODUCT_GAP_SHARE) & (summ["BelowFloor"] >= PRODUCT_GAP_MIN)
        if HAS_BKG:
            summ["Bookings"] = summ["giata"].map(lambda x: bk_map.get(norm_giata(x), 0))
        summ = summ.sort_values(["ProductGap","BelowFloor"], ascending=[False,False])

        n_gap = int(summ["ProductGap"].sum())
        if n_gap:
            st.error(f"🚫 {n_gap} hotel(s) flagged as product gaps — competitor undercuts beyond "
                     f"what margin can reach.")
        else:
            st.success("No hotels meet the product-gap threshold in this selection.")

        out = pd.DataFrame({
            "Giata":summ["giata"], "Hotel":summ["hotel_name"],
            "Dest":summ["destination"], "Comp":summ["competitor"],
        })
        if HAS_BKG: out["📊 Bkgs"] = summ["Bookings"].map("{:,.0f}".format)
        out["Loses"]          = summ["Loses"]
        out["Below Floor"]    = summ["BelowFloor"].astype(int)
        out["% Below Floor"]  = (summ["Share"]*100).map("{:.0f}%".format)
        out["Avg Gap %"]      = summ["AvgGap"].map("{:.1f}%".format)
        out["Match Markup"]   = summ["AvgMatchMarkup"].map(lambda x: f"{x:.1f}%" if x>0 else "—")
        out["Verdict"]        = np.where(summ["ProductGap"], "🚫 Product Gap", "Reachable on price")
        st.dataframe(out, use_container_width=True, height=460)
        st.caption(f"{len(out)} hotels with losses · {n_gap} product gaps · "
                   f"“Match Markup” is the average markup we’d need to match the competitor "
                   f"(below {MARGIN_FLOOR:.0f}% = unreachable).")

# ── 8. NO COMP ─────────────────────────────────────────────────────────────────
with tabs[7]:
    st.subheader("No Direct Comparison — Destination-Guided")
    st.info("📌 Never 'price freely' — always destination-guided.")
    v = tab_filters(df, "nc")
    vc2 = v[v["result"].isin(["Win","Lose"])]
    stances = {}
    for (d,w),grp in vc2.groupby(["destination","dep_window"]):
        wc=(grp["result"]=="Win").sum(); lc=(grp["result"]=="Lose").sum()
        wr2=wc/(wc+lc) if (wc+lc)>0 else 0
        al2=grp[grp["result"]=="Lose"]["diff_pct"].mean() if lc>0 else 0
        aw2=grp[grp["result"]=="Win"]["diff_pct"].mean() if wc>0 else 0
        stances[(d,w)] = derive_stance(wr2,al2,aw2)
    nc2 = v[v["result"].str.contains("No Comp",na=False)].copy()
    # Only hotels with NO comparison on ANY date across the data. Hotels that DO have a
    # comparison on other dates are excluded here and handled by the markup suggestions.
    nc2 = nc2[~nc2["giata"].isin(compared_giatas)]
    if nc2.empty:
        st.caption("No hotels are uncompared across all dates in the current selection.")
    else:
        nc2["Stance"] = nc2.apply(lambda r:stances.get((r["destination"],r["dep_window"]),"Hold — Optimise Margin"),axis=1)
        if HAS_BKG:
            nc2["bkg"] = nc2["giata"].map(lambda x: bk_map.get(norm_giata(x), 0))
            nc2["rev"] = nc2["giata"].map(lambda x: rev_map.get(norm_giata(x), 0.0))
            # Opportunity: no competitor anywhere, but we ARE selling it → room to raise margin
            booked = nc2[nc2["bkg"] > 0]
            if not booked.empty:
                st.success(f"💡 {booked['giata'].nunique()} no-comp hotel(s) have bookings "
                           f"({int(booked.drop_duplicates('giata')['bkg'].sum()):,} bookings, "
                           f"£{booked.drop_duplicates('giata')['rev'].sum():,.0f} revenue) — no competitor "
                           f"on any date, so candidates to **increase margin**.")
            nc2 = nc2.sort_values("bkg", ascending=False)
            nc2["Opportunity"] = np.where(nc2["bkg"] > 0, "💡 Booked — review margin", "")
            show8 = nc2[["giata","hotel_name","destination","dep_window","dep_month",
                         "pp_price","current_margin","bkg","rev","Opportunity","Stance"]].copy()
            show8["bkg"] = show8["bkg"].map("{:,.0f}".format)
            show8["rev"] = show8["rev"].map(lambda x: f"£{x:,.0f}" if x>0 else "—")
            show8.columns = ["Giata","Hotel","Dest","Window","Month",
                             "D2 £pp","Margin","📊 Bkgs","Revenue","Opportunity","Dest Stance"]
        else:
            show8 = nc2[["giata","hotel_name","destination","competitor","dep_window","dep_month",
                          "pp_price","current_margin","Stance","booking_tier"]].copy()
            show8.columns = ["Giata","Hotel","Dest","Comp","Window","Month",
                             "D2 £pp","Margin","Dest Stance","Tier"]
        st.dataframe(show8,use_container_width=True,height=500)
        st.caption(f"{nc2['giata'].nunique()} hotel(s) with no comparison on any date "
                   f"({len(nc2)} rows). Hotels compared on other dates are excluded and "
                   f"driven by the markup suggestions instead."
                   + ("" if HAS_BKG else "  Upload bookings to highlight which of these are selling."))

# ── 9. DEST ACTIONS ────────────────────────────────────────────────────────────
with tabs[8]:
    st.subheader("Destination × Window — Pricing Stances")
    v = tab_filters(df, "dst")
    vc3 = v[v["result"].isin(["Win","Lose"])]
    db9 = dest_bookings(v, bk_map)
    rows9 = []
    for (d,w),grp in vc3.groupby(["destination","dep_window"]):
        wc=(grp["result"]=="Win").sum(); lc=(grp["result"]=="Lose").sum()
        ncw=v[(v["result"].str.contains("No Comp",na=False))&(v["destination"]==d)&(v["dep_window"]==w)].shape[0]
        wr3=wc/(wc+lc) if (wc+lc)>0 else 0
        al3=grp[grp["result"]=="Lose"]["diff_pct"].mean() if lc>0 else 0
        aw3=grp[grp["result"]=="Win"]["diff_pct"].mean() if wc>0 else 0
        row = {"Dest":d,"Window":w,"Win Rate":f"{wr3*100:.0f}%",
                       "Avg Lose":f"{al3:.1f}%","Avg Win":f"{aw3:.1f}%",
                       "Stance":derive_stance(wr3,al3,aw3),
                       "Wins":wc,"Losses":lc,"No Comp":ncw}
        if HAS_BKG: row["📊 Bookings"] = f"{db9.get(d,0):,}"
        rows9.append(row)
    if rows9: st.dataframe(pd.DataFrame(rows9),use_container_width=True,height=400)

# ── 10. OUTLIERS ───────────────────────────────────────────────────────────────
with tabs[9]:
    st.subheader("⚑ Outliers")
    v = tab_filters(df, "out")
    vc4 = v[v["result"].isin(["Win","Lose"])]
    cl2,bw2 = [],[]
    for (h,d,w),grp in vc4.groupby(["hotel_name","destination","dep_window"]):
        wc=(grp["result"]=="Win").sum(); lc=(grp["result"]=="Lose").sum()
        gi = grp["giata"].iloc[0] if len(grp)>0 else ""
        if lc>=2 and wc==0:
            cl2.append({"Giata":gi,"Hotel":h,"Dest":d,"Window":w,"Losses":lc,
                         "Avg Gap":f"{grp[grp['result']=='Lose']['diff_pct'].mean():.1f}%",
                         "Floor":"⚠ YES" if grp["margin_flag"].eq("floor").any() else "No"})
        if wc>=2:
            aw4=grp[grp["result"]=="Win"]["diff_pct"].mean()
            if aw4>30:
                bw2.append({"Giata":gi,"Hotel":h,"Dest":d,"Window":w,"Wins":wc,"Avg Win":f"+{aw4:.1f}%"})
    if cl2:
        st.error(f"🔴 {len(cl2)} consistent loser situations")
        st.dataframe(with_bkg(pd.DataFrame(cl2)),use_container_width=True)
    if bw2:
        st.success(f"🔵 {len(bw2)} big win opportunities")
        st.dataframe(with_bkg(pd.DataFrame(bw2)),use_container_width=True)
    if not cl2 and not bw2: st.info("No severe outliers in current filter.")

# ── 11. MISSING ────────────────────────────────────────────────────────────────
with tabs[10]:
    st.subheader("Missing Properties")
    v = tab_filters(df, "mis")
    v_files = v["file_name"].nunique()
    if v_files > 1:
        hf = v.groupby(["giata","hotel_name"])["file_name"].nunique().reset_index()
        miss = hf[hf["file_name"]<v_files].copy()
        miss["Files Missing"] = v_files - miss["file_name"]
        miss.columns = ["Giata","Hotel","Files Present","Files Missing"]
        st.dataframe(miss.sort_values("Files Missing",ascending=False),use_container_width=True)
    else:
        st.info("Load 2+ files to detect missing properties across runs.")

# ── 12. ADVANTAGE ──────────────────────────────────────────────────────────────
with tabs[11]:
    st.subheader("Competitive Advantage — D2 Strong Win Positions")
    v = tab_filters(df, "adv")
    adv = v[v["result"]=="Win"][v[v["result"]=="Win"]["diff_pct"]>HOLD_UPPER].copy()
    adv2 = adv.groupby(["giata","hotel_name","destination","competitor"]).agg(
        Count=("diff_pct","count"),AvgWin=("diff_pct","mean"),AvgMargin=("current_margin","mean")
    ).reset_index().sort_values("AvgWin",ascending=False)
    adv2["AvgWin"]    = adv2["AvgWin"].map("{:+.1f}%".format)
    adv2["AvgMargin"] = adv2["AvgMargin"].map("{:.1f}%".format)
    adv2.columns = ["Giata","Hotel","Dest","Comp","Records","Avg Win Gap","Avg Margin"]
    adv2["Recommendation"] = "Consider margin increase"
    st.dataframe(with_bkg(adv2),use_container_width=True,height=500)

# ── 13. INSIGHTS ───────────────────────────────────────────────────────────────
with tabs[12]:
    st.subheader("Strategic Insights")
    v = tab_filters(df, "ins")
    v_loses = v[v["result"]=="Lose"]
    v_wins  = v[v["result"]=="Win"]
    fln3 = (v_loses["margin_flag"]=="floor").sum() if not v_loses.empty else 0
    st.error("### 🔴 Critical Losses")
    if not v_loses.empty:
        tl = v_loses.nsmallest(8,"diff_pct")[["giata","hotel_name","destination","competitor",
              "dep_window","pp_price","comp_price","diff_pct","margin_after","margin_flag"]].copy()
        tl.columns = ["Giata","Hotel","Dest","Comp","Window","D2 £pp","Comp £pp","Gap %","Margin After","Flag"]
        st.dataframe(with_bkg(tl).style.format({"D2 £pp":"{:.0f}","Comp £pp":"{:.0f}",
                                       "Gap %":"{:.1f}%","Margin After":"{:.1f}%"}),
                     use_container_width=True)
    st.warning(f"### ⚠ Margin Guardrails\n- **{fln3}** hotels breach 7% floor\n- Avg D2 margin: **{v['current_margin'].mean():.1f}%**")
    if not v_wins.empty:
        st.success("### ✅ Margin Opportunities")
        bw3 = v_wins[v_wins["diff_pct"]>20].nlargest(5,"diff_pct")[
            ["giata","hotel_name","destination","competitor","dep_window","diff_pct","current_margin"]].copy()
        bw3.columns = ["Giata","Hotel","Dest","Comp","Window","Win %","Margin"]
        st.dataframe(with_bkg(bw3).style.format({"Win %":"{:.1f}%","Margin":"{:.1f}%"}),use_container_width=True)
    st.info("### 🗓 Window Split")
    vc5 = v[v["result"].isin(["Win","Lose"])]
    for w in ["0-60","61-120","121-240","241+"]:
        wg = vc5[vc5["dep_window"]==w]
        if wg.empty: continue
        wr5 = (wg["result"]=="Win").sum()/len(wg)*100
        st.write(f"**{w}d** — {wr5:.0f}% ({(wg['result']=='Win').sum()}W / {(wg['result']=='Lose').sum()}L)")

# ── 14. PRICE TRENDS ───────────────────────────────────────────────────────────
with tabs[13]:
    st.subheader("↗ Price Trends")
    v = tab_filters(df, "trd", show_hotel=True)

    if HAS_BKG and not v.empty:
        uniq = v.drop_duplicates("giata")
        b1,b2,b3 = st.columns(3)
        b1.metric("📊 Bookings (selection)", f"{int(uniq['bookings'].sum()):,}")
        b2.metric("£ Revenue (selection)",   f"£{uniq['revenue'].sum():,.0f}")
        b3.metric("Hotels with bookings",    f"{int((uniq['bookings']>0).sum()):,}")

    st.markdown("---")

    # Seasonal curve
    if not v.empty:
        sc = v.copy()
        sc["pp_price"]   = pd.to_numeric(sc["pp_price"],   errors="coerce")
        sc["comp_price"] = pd.to_numeric(sc["comp_price"], errors="coerce")
        sc["dep_date_dt"] = pd.to_datetime(sc["dep_date"].astype(str), errors="coerce")
        sc = sc[sc["pp_price"]>100].dropna(subset=["dep_date_dt"]).sort_values("dep_date_dt")

        if sc.empty:
            st.warning("No valid price/date data for this selection.")
        else:
            sc_grp = sc.groupby("dep_date_dt",as_index=False).agg(
                D2_pp  =("pp_price","mean"),
                Comp_pp=("comp_price",lambda x: x[x>0].mean() if (x>0).any() else np.nan)
            )
            dest_lbl = ", ".join(v["destination"].unique().tolist())
            chart_title = f"Avg D2 vs Competitor Price — {dest_lbl}"
            fig_sc = go.Figure()
            fig_sc.add_trace(go.Scatter(x=sc_grp["dep_date_dt"],y=sc_grp["D2_pp"],
                name="D2 Price £",mode="lines+markers",
                line=dict(color="#1B6FD4",width=3),marker=dict(size=6)))
            cv = sc_grp[sc_grp["Comp_pp"].notna()&(sc_grp["Comp_pp"]>100)]
            if not cv.empty:
                fig_sc.add_trace(go.Scatter(x=cv["dep_date_dt"],y=cv["Comp_pp"],
                    name="Comp Price £",mode="lines+markers",
                    line=dict(color="#E04A3F",width=2,dash="dot"),marker=dict(size=6)))
            fig_sc.update_layout(title=chart_title,
                xaxis=dict(title="Departure Date",type="date"),
                yaxis_title="Price per person £",height=380,hovermode="x unified",
                legend=dict(orientation="h",yanchor="bottom",y=1.02))
            st.plotly_chart(fig_sc,use_container_width=True)

    # Avg markup & bookings by departure month
    if not v.empty:
        dest_lbl_m = ", ".join(v["destination"].unique().tolist())
        mk_src = v.copy()
        mk_src["current_margin"] = pd.to_numeric(mk_src["current_margin"], errors="coerce")
        mk_m = (mk_src[mk_src["dep_month"].ne("")].groupby("dep_month")
                  .agg(Markup=("current_margin","mean")).reset_index().sort_values("dep_month"))
        if not mk_m.empty:
            mk_m["Label"] = mk_m["dep_month"].apply(
                lambda m: pd.to_datetime(m+"-01").strftime("%b %Y") if len(m)==7 else m)
            figm = go.Figure()
            if HAS_TX:
                bk_dep = bookings_by_month(bk_tx, v["giata"].unique(), "dep_month")
                if not bk_dep.empty:
                    bmap = dict(zip(bk_dep["Month"], bk_dep["Bookings"]))
                    figm.add_trace(go.Bar(x=mk_m["Label"],
                        y=[bmap.get(m,0) for m in mk_m["dep_month"]],
                        name="Bookings", marker_color="rgba(27,111,212,0.35)", yaxis="y2"))
            figm.add_trace(go.Scatter(x=mk_m["Label"], y=mk_m["Markup"], name="Avg Markup %",
                mode="lines+markers", line=dict(color="#D4AF37", width=3), marker=dict(size=6)))
            figm.add_hline(y=MARGIN_FLOOR, line_dash="dot", line_color="#C0392B",
                            annotation_text=f"floor {MARGIN_FLOOR:.0f}%", annotation_position="right")
            figm.update_layout(title=f"Avg Markup &amp; Bookings by Departure Month — {dest_lbl_m}",
                yaxis=dict(title="Avg Markup %"),
                yaxis2=dict(title="Bookings", overlaying="y", side="right", showgrid=False),
                height=340, hovermode="x unified", barmode="overlay",
                legend=dict(orientation="h", yanchor="bottom", y=1.02))
            st.plotly_chart(figm, use_container_width=True)

    st.markdown("---")
    col3,col4 = st.columns(2)

    with col3:
        wr_src = v[v["result"].isin(["Win","Lose"])].copy()
        if not wr_src.empty and wr_src["dep_month"].ne("").any():
            wr_m = wr_src.groupby("dep_month").apply(
                lambda g:(g["result"]=="Win").sum()/len(g)*100).reset_index()
            wr_m.columns = ["Month","WinRate"]
            wr_m = wr_m[wr_m["Month"]!=""].sort_values("Month")
            wr_m["Label"]  = wr_m["Month"].apply(
                lambda m: pd.to_datetime(m+"-01").strftime("%b %Y") if len(m)==7 else m)
            wr_m["Colour"] = wr_m["WinRate"].apply(lambda v2:"#C0392B" if v2<50 else "#27AE60")
            fig_wr = go.Figure()
            fig_wr.add_trace(go.Bar(x=wr_m["Label"],y=wr_m["WinRate"],marker_color=wr_m["Colour"],
                                    name="Win Rate %"))
            if HAS_TX:
                bk_dep_wr = bookings_by_month(bk_tx, wr_src["giata"].unique(), "dep_month")
                if not bk_dep_wr.empty:
                    wbmap = dict(zip(bk_dep_wr["Label"], bk_dep_wr["Bookings"]))
                    fig_wr.add_trace(go.Scatter(x=wr_m["Label"],
                        y=[wbmap.get(l,0) for l in wr_m["Label"]],
                        name="Bookings", mode="lines+markers",
                        line=dict(color="#1B6FD4",width=2), yaxis="y2"))
            fig_wr.add_hline(y=50,line_dash="dot",line_color="#E04A3F",
                              annotation_text="50%",annotation_position="right")
            dest_lbl2 = ", ".join(v["destination"].unique().tolist())
            fig_wr.update_layout(title=f"Win Rate &amp; Bookings by Departure Month — {dest_lbl2}",
                yaxis=dict(range=[0,100],title="Win Rate %"),
                yaxis2=dict(title="Bookings", overlaying="y", side="right", showgrid=False),
                xaxis_title="Departure Month",height=320,
                legend=dict(orientation="h",yanchor="bottom",y=1.02))
            st.plotly_chart(fig_wr,use_container_width=True)

    with col4:
        wr_run_src = v[v["result"].isin(["Win","Lose"])].copy()
        if not wr_run_src.empty and wr_run_src["file_date"].nunique()>1:
            wr_run = wr_run_src.groupby("file_date").apply(
                lambda g:(g["result"]=="Win").sum()/len(g)*100).reset_index()
            wr_run.columns = ["RunDate","WinRate"]
            wr_run = wr_run.sort_values("RunDate")
            wr_run["Label"] = pd.to_datetime(wr_run["RunDate"]).dt.strftime("%d %b")
            fig_run = go.Figure()
            fig_run.add_trace(go.Scatter(x=wr_run["Label"],y=wr_run["WinRate"],
                mode="lines+markers",line=dict(color="#27AE60",width=2.5),
                marker=dict(size=8,color="#27AE60"),
                fill="tozeroy",fillcolor="rgba(39,174,96,0.1)"))
            fig_run.update_layout(
                title=f"Win Rate Trend ({wr_run['Label'].iloc[0]}–{wr_run['Label'].iloc[-1]})",
                yaxis_title="Win Rate %",xaxis_title="Run Date",height=320,showlegend=False)
            st.plotly_chart(fig_run,use_container_width=True)
        elif not wr_run_src.empty:
            wr_now=(wr_run_src["result"]=="Win").sum()/len(wr_run_src)*100
            col4.metric("Current Win Rate",f"{wr_now:.1f}%","Upload more files to track trend")

    # Booking pace by booked month (when bookings carry booked dates)
    if HAS_TX and not v.empty:
        bk_booked = bookings_by_month(bk_tx, v["giata"].unique(), "booked_month")
        if not bk_booked.empty:
            figp = go.Figure()
            figp.add_trace(go.Bar(x=bk_booked["Label"], y=bk_booked["Bookings"],
                name="Bookings", marker_color="#1B6FD4"))
            figp.add_trace(go.Scatter(x=bk_booked["Label"], y=bk_booked["Revenue"],
                name="Revenue £", mode="lines+markers",
                line=dict(color="#D4AF37", width=2), yaxis="y2"))
            figp.update_layout(title="Booking Pace — by Booked Month (when the bookings were made)",
                yaxis=dict(title="Bookings"),
                yaxis2=dict(title="Revenue £", overlaying="y", side="right", showgrid=False),
                height=320, hovermode="x unified",
                legend=dict(orientation="h", yanchor="bottom", y=1.02))
            st.plotly_chart(figp, use_container_width=True)

    if v["file_date"].nunique()>1:
        st.markdown("#### Price Drift Across Run Dates")
        drift_src = v.copy()
        drift_src["pp_price"]   = pd.to_numeric(drift_src["pp_price"],   errors="coerce")
        drift_src["comp_price"] = pd.to_numeric(drift_src["comp_price"], errors="coerce")
        drift = drift_src[drift_src["pp_price"]>100].groupby("file_date").agg(
            D2  =("pp_price","mean"),
            Comp=("comp_price",lambda x: x[x>0].mean() if (x>0).any() else np.nan)
        ).reset_index().sort_values("file_date")
        drift["Label"] = pd.to_datetime(drift["file_date"]).dt.strftime("%d %b")
        dest_lbl3 = ", ".join(v["destination"].unique().tolist())
        fig_dr = go.Figure()
        fig_dr.add_trace(go.Scatter(x=drift["Label"],y=drift["D2"],
            name="D2 avg £pp",line=dict(color="#1B6FD4",width=2.5),mode="lines+markers"))
        if drift["Comp"].notna().any():
            fig_dr.add_trace(go.Scatter(x=drift["Label"],y=drift["Comp"],
                name="Comp avg £pp",line=dict(color="#E04A3F",width=2,dash="dash"),mode="lines+markers"))
        fig_dr.update_layout(title=f"Price Drift — {dest_lbl3}",
                              height=260,xaxis_title="Run Date",yaxis_title="£ pp")
        st.plotly_chart(fig_dr,use_container_width=True)

# ── 15. FILES ──────────────────────────────────────────────────────────────────
with tabs[14]:
    st.subheader("📂 Uploaded Files")
    fl = get_file_list()
    if fl:
        fdf = pd.DataFrame(fl,columns=["File","Run Date","Destination","Competitor","Rows"])
        st.dataframe(fdf,use_container_width=True)
        st.caption(f"{len(fdf)} files · {fdf['Rows'].sum():,} total rows")
        st.download_button("⬇ Download filtered data as CSV",
            data=df.to_csv(index=False),file_name="d2_pricing_export.csv",mime="text/csv")
    else:
        st.info("No files uploaded yet.")

# ── 16. MARKUP LEVERS ───────────────────────────────────────────────────────────
with tabs[15]:
    st.subheader("🎚 Markup Levers — which rule to change")
    if not HAS_RULES:
        st.info("Upload the markup rule-engine export in the sidebar to enable this. "
                "It links each suggested margin change to the actual rule you'd edit.")
    elif df.empty:
        st.info("Upload pricing/comparison data to generate suggestions first.")
    else:
        st.caption("For each hotel where the engine suggests a markup change, this finds the "
                   "**consumer-layer rule that targets that hotel ALONE** to edit — preferring the "
                   "cleanest lever (fewest conditions, package-level for packages, Exclusive over Manual). "
                   "Rules that list several hotels (GIata IsOneOf …) are **shared**: editing one would "
                   "move every hotel in the list, so they're never proposed — instead the app recommends "
                   "**adding a single-hotel rule** carrying the change (consumer markups stack additively). "
                   "System-cost rules are context only. The new value is an approximate starting point — "
                   "a human approves before it goes live.")

        # Scope: hotels in the filtered comparison set, with a suggested change
        v = tab_filters(df, "lev")
        hotels = (v[v["result"].isin(["Win","Lose"])]
                    .drop_duplicates("giata")[["giata","hotel_name","destination"]])
        is_pkg = st.radio("Search type for lever matching", ["Package","Hotel-only"],
                          horizontal=True) == "Package"
        only_changes = st.checkbox("Only show hotels needing a change", value=True)

        rows_out, detail = [], {}
        for _, h in hotels.iterrows():
            g = h["giata"]
            s = suggest_markup_change(v[(v["giata"]==g) & (v["hotel_name"]==h["hotel_name"])])
            delta = s["delta"]
            if only_changes and abs(delta) < 0.05: continue
            rec = recommend_lever(rules_df, g, delta, is_package=is_pkg)
            bvol = bk_map.get(norm_giata(g), 0) if HAS_BKG else None
            if rec["status"]=="edit":
                r = rec["rule"]
                nv = rec["new_value"]
                cur = f"{r['action_value']:g}{r['action_unit']}" if pd.notna(r["action_value"]) else "—"
                newv = (f"{nv:g}%" if nv is not None else "set £ manually")
                action = f"EDIT rule {r['rid']} (this hotel only)"
                lever_text = r["generated_text"]
            elif rec["status"]=="create":
                cur, newv = "—", f"{delta:+.1f}% (new)"
                lvl = "HotelFlightPackage" if is_pkg else "HotelRoom"
                if rec.get("reason")=="shared":
                    ns = int(rec["shared_context"]["giata_n"].max()) if not rec["shared_context"].empty else 0
                    action = f"ADD hotel rule (existing shared across {ns} hotels)"
                else:
                    action = "ADD hotel rule (no hotel-specific rule exists)"
                lever_text = (f"(suggested) If (GIata Equal {g}) And (User Equal Destination2) "
                              f"Then MarkupNew {lvl} By {delta:+.1f}% Total")
            else:
                cur, newv, action, lever_text = "—","—","No rules for giata","(none found in export)"
            row = {"Giata":g,"Hotel":h["hotel_name"],"Dest":h["destination"],
                   "Sugg. Δ":f"{delta:+.1f}pp","Action":action,
                   "Current":cur,"→ New":newv,"Rule (GeneratedText)":lever_text}
            if HAS_BKG: row["📊 Bkgs"] = f"{bvol:,}"
            rows_out.append(row); detail[g]=(s,rec)

        if not rows_out:
            st.success("No markup changes suggested for the current selection.")
        else:
            order = ["Giata","Hotel","Dest"] + (["📊 Bkgs"] if HAS_BKG else []) + \
                    ["Sugg. Δ","Action","Current","→ New","Rule (GeneratedText)"]
            out = pd.DataFrame(rows_out)[order]
            st.dataframe(out, use_container_width=True, height=460)
            st.caption(f"{len(out)} hotel(s) with a suggested lever. "
                       f"{(out['Action'].str.startswith('EDIT')).sum()} edit an existing single-hotel rule · "
                       f"{(out['Action'].str.startswith('ADD')).sum()} need a new hotel-specific rule.")

            # Drill-down: full applicable stack for one hotel
            pick = st.selectbox("Inspect a hotel's full applicable rule stack",
                                [f"{r['Giata']} · {r['Hotel']}" for r in rows_out])
            if pick:
                gsel = pick.split(" · ")[0]
                s, rec = detail[gsel]
                c1,c2,c3 = st.columns(3)
                c1.metric("Suggested Δ", f"{s['delta']:+.1f}pp")
                c2.metric("Proj. Win Rate", f"{s['wr0']:.0f}%→{s['wr1']:.0f}%" if pd.notna(s['wr0']) else "—")
                c3.metric("Proj. Markup", f"{s['m0']:.1f}%→{s['m1']:.1f}%")
                if rec["status"]=="edit":
                    r = rec["rule"]
                    st.markdown(f"**Recommended lever — edit rule `{r['rid']}` "
                                f"({r['layer']}, applies to {'package' if r['applies_to']=='15' else 'hotel'}):**")
                    st.code(r["generated_text"], language=None)
                    if rec["new_value"] is not None:
                        st.markdown(f"Change **{r['action_value']:g}% → {rec['new_value']:g}%** "
                                    f"(approx; confirm against the hotel:flight split).")
                    alts = rec.get("alternatives")
                    if alts is not None and not alts.empty:
                        st.markdown("**Other levers on this hotel:**")
                        st.dataframe(alts[["rid","layer","applies_to","action_value","action_unit","generated_text"]]
                                     .rename(columns={"rid":"Id","action_value":"Val","action_unit":"Unit",
                                                      "generated_text":"GeneratedText","applies_to":"AppliesTo"}),
                                     use_container_width=True)
                elif rec["status"]=="create":
                    lvl = "HotelFlightPackage" if is_pkg else "HotelRoom"
                    if rec.get("reason")=="shared":
                        st.warning("The markup for this hotel comes from a **shared rule listing several "
                                   "hotels** (GIata IsOneOf …). Editing it would change every hotel in that "
                                   "list. Instead, add a single-hotel rule carrying the change — consumer "
                                   "markups stack, so it adjusts only this hotel:")
                    else:
                        st.warning("No single-hotel consumer rule exists (markup comes from blanket/provider "
                                   "rules that can't be edited for one hotel). Add a hotel-specific rule:")
                    st.code(f"If (GIata Equal {gsel}) And (User Equal Destination2) "
                            f"Then MarkupNew {lvl} By {s['delta']:+.1f}% Total", language=None)
                    shc = rec.get("shared_context")
                    if shc is not None and not shc.empty:
                        st.markdown("**Shared rules touching this hotel (do NOT edit — affect multiple hotels):**")
                        st.dataframe(shc[["rid","giata_n","layer","action_value","action_unit","generated_text"]]
                                     .rename(columns={"rid":"Id","giata_n":"# Hotels","action_value":"Val",
                                                      "action_unit":"Unit","generated_text":"GeneratedText"}),
                                     use_container_width=True)
                sc = rec.get("system_context")
                if sc is not None and not sc.empty:
                    st.markdown("**System-cost rules on this hotel (context only — not margin levers):**")
                    st.dataframe(sc[["rid","action_value","action_unit","generated_text"]]
                                 .rename(columns={"rid":"Id","action_value":"Val","action_unit":"Unit",
                                                  "generated_text":"GeneratedText"}),
                                 use_container_width=True)

        # ── Rules that may be driving our price-vs-competitor gaps ──────────────
        st.markdown("---")
        st.markdown("#### 🔎 Markup rules that may be driving our losses")
        st.caption("Across the **losing** hotels in this selection, which markup rules recur most? "
                   "A rule that sits on many losers — especially with a wide average gap — is a "
                   "candidate for systematically pricing us above the competitor. Consumer rules are "
                   "levers you can adjust; system-cost rules are shown for context.")
        losers = v[v["result"]=="Lose"].drop_duplicates("giata")[["giata","diff_pct"]]
        if losers.empty:
            st.success("No losses in the current selection.")
        else:
            # Build a giata → rule-rows index once (markups only), then tally over losers
            rmk = rules_df[(rules_df["enabled"]==1) &
                           (rules_df["verb"].isin(["MarkupNew","MarkupPerPerson"]))].copy()
            idx = {}
            for _, r in rmk.iterrows():
                for gg in (r["giatas"].split(",") if r["giatas"] else []):
                    idx.setdefault(gg, []).append(r)
            tally = {}
            for _, lr in losers.iterrows():
                for r in idx.get(norm_giata(lr["giata"]), []):
                    t = tally.setdefault(r["rid"], {"rid":r["rid"],"layer":r["layer"],
                        "val":r["action_value"],"unit":r["action_unit"],"giata_n":r["giata_n"],
                        "gt":r["generated_text"],"hotels":0,"gapsum":0.0})
                    t["hotels"] += 1; t["gapsum"] += abs(lr["diff_pct"])
            if not tally:
                st.info("None of the losing hotels are touched by markup rules in the export.")
            else:
                imp = pd.DataFrame(tally.values())
                imp["Avg Loss Gap"] = (imp["gapsum"]/imp["hotels"]).map("{:.1f}%".format)
                imp["Reach"] = np.where(imp["giata_n"]>1, imp["giata_n"].map(lambda n:f"shared ×{int(n)}"), "single hotel")
                imp = imp.sort_values(["hotels","gapsum"], ascending=False).head(20)
                out2 = pd.DataFrame({
                    "Rule Id":imp["rid"], "Layer":imp["layer"],
                    "Losing Hotels":imp["hotels"].astype(int),
                    "Avg Loss Gap":imp["Avg Loss Gap"],
                    "Markup":imp.apply(lambda r:f"{r['val']:g}{r['unit']}" if pd.notna(r['val']) else "—",axis=1),
                    "Reach":imp["Reach"], "GeneratedText":imp["gt"]})
                st.dataframe(out2, use_container_width=True, height=420)
                st.caption("Ranked by how many losing hotels each rule touches. Higher up = more likely "
                           "to be contributing to the price gap. Consumer-layer rules can be tuned; "
                           "system-cost rules reflect real cost and usually can't.")

# ── 17. HOW IT WORKS (reference guide for new users) ────────────────────────────
# MAINTAINER NOTE: this tab is the plain-English guide to the app's logic. If you
# change how decisions are made (the suggestion engine, thresholds, layers, tabs),
# update the relevant text below at the same time so the guide stays accurate. The
# key numbers are pulled from the live constants, so tuning them updates the guide
# automatically.
with tabs[16]:
    st.subheader("📖 How It Works — a plain-English guide")
    st.caption("Reference for anyone seeing the dashboard for the first time. Mirrors the "
               "current logic; the key thresholds below read from the live settings.")

    st.markdown(f"""
**What this is.** This tool takes our pricing-versus-competitor data, our booking data,
and our markup rule settings, and tells us — for each hotel — whether we should change
our margin, by how much, and exactly which rule to edit to do it. Its job is to turn
thousands of rows of price comparisons into a short, prioritised list of sensible
commercial actions, instead of working them out by hand.

**The one-line version:** where we lose, it works out if a small margin cut would win the
business back; where we win comfortably, it works out how much more margin we can safely
take — always staying inside agreed limits.
""")

    st.markdown("### 1.  What you feed it")
    st.markdown("""
The app combines three uploads. It works with just the first, and gets smarter as you add the other two.

| Upload | What it is | What the app does with it |
|---|---|---|
| **Pricing / comparison data** (the "D2 Live" sheets) | One row per hotel, per departure date, per board: our price, our Total Markup, the competitor price, and the result. | The backbone — every win/loss, gap and margin suggestion comes from here. |
| **Booking data** | One row per actual booking, tagged with the hotel's Giata ID. | Shows real booking volume and revenue per hotel, and prioritises the busiest hotels. |
| **Markup rule export** | The rules from the markup system that build each hotel's price. | Lets the app name the exact rule you'd edit to apply a suggested change. |

Everything joins on the hotel's **Giata ID** — a unique number per hotel — so the three files line up automatically.
""")

    st.markdown("### 2.  A few words you'll see everywhere")
    st.markdown(f"""
| Term | What it means |
|---|---|
| **Total Markup** | The margin we add on top of cost to get the selling price. This is the lever the app suggests changing. |
| **Win** | Our price is cheaper than the competitor's — we'd expect to take the booking. |
| **Lose** | Our price is dearer than the competitor's — we'd likely lose the booking. |
| **Win – No Comp** | There was no competitor to compare against on that date (handled separately — see §6). |
| **The gap** | How far our price sits from the competitor's, as a %. Positive = we're cheaper (winning); negative = dearer (losing). |
| **Floor & Ceiling** | Margin is never pushed below **{MARGIN_FLOOR:.0f}%** (floor) or above **{MARGIN_CEILING:.0f}%** (ceiling). These limits are never broken. |
""")

    st.markdown("### 3.  The core decision: should we change this hotel's margin?")
    st.markdown(f"""
This is the heart of the tool. For each hotel it looks at **all its departure dates together**
(using the most recent data), counts how often we win versus lose, then follows one of two paths.

**Important:** it only looks at dates where there was a real competitor. *Win – No Comp* dates
are ignored completely, so they can't distort the decision.

**Path A — winning most of the time ({WR_HIGH:.0f}%+ win rate).**
The hotel is clearly competitive, so the goal switches to making more money rather than chasing
the last few wins. The app finds the biggest margin **increase** it can take while keeping the win
rate at {WR_KEEP:.0f}%+ and turning no more than ~{FLIP_FRAC*100:.0f}% of current wins into losses.
In short: *raise margin to capture profit, without giving away the strong position we already have.*

**Path B — genuinely competitive (below {WR_HIGH:.0f}% win rate).**
It weighs the trade-off both ways, testing every small margin move and picking the one with the best
balance of **how often we'd win** (more bookings) and **how much margin we'd keep** (more profit per
booking). A modest cut that flips several losses → suggests the cut; already best → hold; losing by
too much for any sensible cut → flagged as a product/positioning problem, not pricing.

**The "don't chase pennies" rule.** A loss of a few pennies is almost always noise (prices move
constantly, refreshes lag). Any "loss" within **{NOISE_GAP}% of breaking even** is treated as a draw,
so we never shave margin across a whole hotel to flip one near-zero loss.

**The safety rails (always on).** No date is pushed outside the {MARGIN_FLOOR:.0f}%–{MARGIN_CEILING:.0f}%
band, and if two suggestions are about as good, it picks the one that changes margin the least.
""")
    st.markdown("""
| Hotel | What the data shows | What the app says |
|---|---|---|
| **Grand Hyatt Dubai** | 11 wins of 12 dates, the only "loss" being 23p (a draw). Headroom on the wins. | **Raise.** Winning comfortably, so take more margin — the penny loss is ignored. |
| **A contested hotel** | Winning ~38% of dates, with several real losses of 2–4%. | **Cut.** A modest cut flips several real losses into wins. |
""")

    st.markdown("### 4.  From a suggestion to the actual rule to change")
    st.markdown("""
A suggestion is only useful if you know which setting to change. A price is built from many stacked
rules, so the **Markup Levers** tab pinpoints the right one:

- It only suggests rules you can **safely change for one hotel** (a hotel-specific, customer-facing rule) — never a blanket rule that would move every hotel.
- If no hotel-specific rule exists, it tells you to **create** one and writes out exactly what it should say.
- **Cost-recovery rules** (supplier/system costs) are shown for context only — they aren't a margin lever.

**How a price is built** (the order the markup system applies things, which the app mirrors):
1. **System / cost markups first** — added to the hotel and flight separately, to cover supplier/system costs.
2. **Customer (consumer) markups next** — our actual margin, on top. (An "Exclusive" rate, where we have one, *replaces* the standard margin rather than adding to it.)
3. **Package adjustment last** — a final tweak on the combined hotel-plus-flight price.
""")

    st.markdown("### 5.  How booking data is used")
    st.markdown("""
When the booking file is loaded, real booking volume and revenue show beside each hotel, so you can
see which suggestions matter most. A hotel selling 48 holidays deserves more attention than one selling
two — and volume also pushes the busiest hotels up the action lists.
""")

    st.markdown("### 6.  Hotels with no competitor")
    st.markdown("""
The **No Comp** tab lists only hotels with **no competitor comparison on any date at all**. If a hotel
has a comparison on even one date, it's treated as a normal hotel and handled by the margin suggestions —
its no-comp dates simply don't affect the decision. When booking data is loaded, any of these hotels that
are **actually selling** are highlighted: no competitor + real bookings is a prime candidate to **raise
margin**. For genuinely uncompared hotels with no bookings, the app falls back to a destination-level steer.
""")

    st.markdown("### 7.  A quick tour of the tabs")
    st.markdown("""
| Tab | What it's for |
|---|---|
| **Master Queue** | The single prioritised to-do list — highest-impact actions across all hotels, busiest first. |
| **Overview** | Headline picture: win rate, losses, margin, bookings, and Top Sellers with each hotel's suggested margin change. |
| **Hotel Actions** | Every hotel/date with its recommended action, filterable. |
| **Reduce / Raise / Maintain** | The same actions split into three buckets — where to cut, raise, or hold. |
| **Product Gaps** | Hotels where most losses can't be matched without breaking the margin floor — unreachable on price, so a product/contracting issue. |
| **No Comp** | Hotels with no competitor on any date (see §6). |
| **Dest Actions** | The competitive picture rolled up by destination and departure window. |
| **Outliers** | Consistent losers to fix and big-win opportunities to capitalise on. |
| **Missing** | Gaps in the data. |
| **Advantage** | Hotels where we win by a wide margin — clear room to raise. |
| **Insights** | Critical losses and margin opportunities at a glance. |
| **Price Trends** | How pricing, **markup and bookings** move over time — by departure month and by booked month, with win-rate and price-drift trends. |
| **Files** | What's been uploaded; export the data. |
| **Markup Levers** | Links each suggested change to the exact rule to edit (§4), and flags the markup rules that recur across losing hotels. |
| **How It Works** | This guide. |
""")

    st.markdown("### 8.  What to keep in mind")
    st.markdown(f"""
The app is a decision-support tool, not an autopilot:

- **Suggestions are starting points** — the recommended change and new rule value are well-reasoned estimates for a person to review and approve, not applied automatically.
- **It's only as current as the data** — a suggestion reflects the latest upload, which is why tiny losses are treated as noise.
- **The exact new value is approximate** — because a price is built from several stacked layers, translating a target margin into one rule's percentage is close but not to-the-penny exact; worth a sense-check on big moves.
- **Limits are firm** — whatever the suggestion, margin stays within the {MARGIN_FLOOR:.0f}%–{MARGIN_CEILING:.0f}% band on every date.
""")
    st.markdown("### 9.  The core idea behind the numbers")
    st.markdown("""
Every decision rests on a small set of dials. They encode our commercial guardrails and how
cautious or bold the suggestions are. They live as named settings at the top of the app, so they
can be changed and redeployed without touching the logic — and this guide updates to match. Here's
what each one does and its current value.
""")
    st.markdown(f"""
**Margin limits — the hard guardrails**

| Setting | Now | What it controls |
|---|---|---|
| Margin floor | **{MARGIN_FLOOR:.0f}%** | The lowest markup we'll ever price at. No suggestion goes below it. |
| Margin ceiling | **{MARGIN_CEILING:.0f}%** | The highest markup we'll ever price at. No suggestion goes above it. |
| Near band | **{MARGIN_NEAR:.1f}%** | A "getting close to the floor" marker, used to flag margins that are running tight. |
| Raise threshold | **{RAISE_THRESHOLD:.0f}%** | A win gap above which a hotel is clearly winning with room to raise (used by the Raise / Advantage views). |
| Hold upper | **{HOLD_UPPER:.0f}%** | The top of the "leave it alone" zone for win gaps (used by the Maintain view). |
""")
    st.markdown(f"""
**How the suggestion decides — markup-suggestion tuning**

| Setting | Now | What it controls |
|---|---|---|
| Noise gap | **{NOISE_GAP}%** | How close to break-even a "loss" must be to count as noise (a draw) and never be chased with a cut. |
| Win-rate high | **{WR_HIGH:.0f}%** | At or above this win rate, the goal switches from balancing to **raising** margin. |
| Win-rate keep | **{WR_KEEP:.0f}%** | The win rate the app won't let a hotel fall below when it raises. |
| Flip tolerance | **{FLIP_FRAC*100:.0f}%** | When raising, the most of a hotel's solid wins it will tolerate tipping into losses to capture more margin. |
""")
    st.markdown(f"""
**Spotting product gaps**

| Setting | Now | What it controls |
|---|---|---|
| Below-floor share | **{PRODUCT_GAP_SHARE*100:.0f}%** | If at least this share of a hotel's losses can't be matched without breaking the floor, it's flagged a product gap. |
| Minimum count | **{PRODUCT_GAP_MIN}** | The fewest below-floor losses needed before a hotel is flagged, so one-off losses don't trip it. |
""")
    st.markdown(f"""
Turning a dial up or down makes the app more cautious or more aggressive in that one dimension —
e.g. lowering the **win-rate high** ({WR_HIGH:.0f}%) treats more hotels as raise candidates; raising
the **noise gap** ({NOISE_GAP}%) lets slightly bigger near-misses be ignored. Best practice: change
one at a time and watch the Top Sellers suggestions after redeploy.
""")

    st.info("**In a sentence:** the app reads our prices against competitors, factors in how much "
            "each hotel actually sells, and produces a short, prioritised list of margin changes — "
            "raise where we win, cut where a small move wins business back, hold where we're already "
            "best — then points to the exact rule to change to make it happen.")
