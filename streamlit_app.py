"""
D2 vs Competitor — Multi-Destination Pricing Dashboard
Streamlit app · upload .xlsm/.xlsx/.csv files · persistent SQLite DB
"""

import streamlit as st
import pandas as pd
import numpy as np
import sqlite3
import os
import io
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
from rapidfuzz import fuzz

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="D2 Pricing Intelligence",
    page_icon="✈",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Constants ──────────────────────────────────────────────────────────────────
DB_PATH = "pricing_data.db"
RUN_DATE = date.today()

MARGIN_FLOOR    = 7.0
MARGIN_CEILING  = 15.0
MARGIN_NEAR     = 8.5
LOSE_THRESHOLD  = -5.0
RAISE_THRESHOLD = 15.0
HOLD_UPPER      = 8.0

RESULT_VALS = ["Win", "Lose", "Win - No Comp", "Win Aft Change"]

# ── Styling ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stSidebar"] { background:#1B1464; }
[data-testid="stSidebar"] * { color:#fff !important; }
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label { color:#D4AF37 !important; }
.metric-card {
    background:#fff; border-radius:8px; padding:14px 18px;
    border-left:4px solid #1B1464; box-shadow:0 1px 4px rgba(0,0,0,.08);
    margin-bottom:8px;
}
.metric-card.red  { border-left-color:#C0392B; }
.metric-card.green{ border-left-color:#0A7C4E; }
.metric-card.amber{ border-left-color:#C17A00; }
.badge-floor  { background:#FEE2E2; color:#991B1B; padding:2px 6px; border-radius:3px; font-size:11px; font-weight:700; }
.badge-near   { background:#FEF3C7; color:#92400E; padding:2px 6px; border-radius:3px; font-size:11px; font-weight:700; }
.badge-ok     { background:#D1FAE5; color:#065F46; padding:2px 6px; border-radius:3px; font-size:11px; font-weight:700; }
.badge-ceiling{ background:#DBEAFE; color:#1E40AF; padding:2px 6px; border-radius:3px; font-size:11px; font-weight:700; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# DATABASE
# ══════════════════════════════════════════════════════════════════════════════
def get_conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)

def init_db():
    conn = get_conn()
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS pricing (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name       TEXT,
            file_date       TEXT,
            destination     TEXT,
            competitor      TEXT,
            giata           TEXT,
            hotel_name      TEXT,
            board           TEXT,
            dep_date        TEXT,
            dep_month       TEXT,
            dep_window      TEXT,
            nights          INTEGER,
            pp_price        REAL,
            current_margin  REAL,
            op_room         TEXT,
            comp_price      REAL,
            diff_gbp        REAL,
            diff_pct        REAL,
            result          TEXT,
            margin_after    REAL,
            margin_range    TEXT,
            margin_flag     TEXT,
            booking_tier    TEXT,
            priority_score  REAL,
            uploaded_at     TEXT
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS bookings (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            hotel_name  TEXT,
            destination TEXT,
            bkgs_4wk    INTEGER,
            bkgs_py     INTEGER,
            uploaded_at TEXT
        )
    """)
    conn.commit()
    conn.close()

init_db()

# ══════════════════════════════════════════════════════════════════════════════
# PARSING
# ══════════════════════════════════════════════════════════════════════════════
COL_MAP = {
    "giata": 0, "hotel_name": 1, "board": 2, "dep_date": 3,
    "nights": 4, "pp_price": 5, "current_margin": 6, "op_room": 7,
    "supplier": 8, "airline": 9, "stopovers": 10, "free_transfers": 11,
    "dep": 12, "arr": 13, "rtn": 14,
    "comp1": 15, "comp2": 16, "comp3": 17, "cheapest_comp": 18,
    "diff_gbp": 19, "diff_pct": 20, "result": 21,
    "margin_after": 22, "margin_range": 23,
}

def parse_date(val):
    if pd.isna(val): return None
    s = str(val).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try: return datetime.strptime(s, fmt).date()
        except: pass
    return None

def dep_window(dep_dt, run_dt=None):
    if dep_dt is None: return "241+"
    rd = run_dt or RUN_DATE
    if isinstance(dep_dt, str):
        dep_dt = parse_date(dep_dt)
    if dep_dt is None: return "241+"
    days = (dep_dt - rd).days
    if days <= 60:  return "0-60"
    if days <= 120: return "61-120"
    if days <= 240: return "121-240"
    return "241+"

def margin_flag(v):
    if pd.isna(v) or v is None: return "unknown"
    v = float(v)
    if v < MARGIN_FLOOR:    return "floor"
    if v < MARGIN_NEAR:     return "near"
    if v > MARGIN_CEILING:  return "ceiling"
    return "ok"

def clean_num(val):
    if pd.isna(val): return np.nan
    s = str(val).replace("£","").replace(",","").strip()
    try: return float(s)
    except: return np.nan

def extract_meta_from_filename(fname):
    """D2_Live_MLE_LH_12_06_2026 → dest=MLE, comp=LH, date=2026-06-12"""
    parts = fname.replace(".xlsm","").replace(".xlsx","").replace(".csv","").split("_")
    dest = parts[2] if len(parts) > 2 else "UNK"
    comp = parts[3] if len(parts) > 3 else "LH"
    try:
        d, m, y = parts[-3], parts[-2], parts[-1]
        file_dt = date(int(y), int(m), int(d))
    except:
        file_dt = date.today()
    return dest, comp, file_dt

def find_data_sheet(xl):
    """Find the sheet containing hotel pricing data."""
    for sname in xl.sheet_names:
        try:
            df = xl.parse(sname, header=None)
            if df.shape[1] < 20: continue
            # Check col 0 has numeric-ish giata values
            col0 = df.iloc[1:6, 0].dropna().astype(str)
            if col0.str.match(r"^\d{4,}").any():
                return sname, df
        except: pass
    return None, None

def parse_pricing_file(uploaded_file, fname):
    dest, comp, file_dt = extract_meta_from_filename(fname)
    rows = []

    try:
        if fname.endswith(".csv"):
            raw = pd.read_csv(uploaded_file, header=None, dtype=str)
            sheets = {"data": raw}
        else:
            xl = pd.ExcelFile(uploaded_file)
            sname, raw = find_data_sheet(xl)
            if raw is None:
                # fallback: first sheet
                raw = xl.parse(xl.sheet_names[0], header=None)
            sheets = {"data": raw}
    except Exception as e:
        st.error(f"Could not read {fname}: {e}")
        return []

    df = sheets["data"]

    # Find header row (contains 'Giata' or numeric IDs starting)
    header_row = 0
    for i, row in df.iterrows():
        vals = row.astype(str).str.lower()
        if "giata" in vals.values or "hotel" in vals.values:
            header_row = i
            break

    df.columns = range(df.shape[1])
    data = df.iloc[header_row+1:].reset_index(drop=True)

    run_dt = file_dt  # use file date as run date for window calc

    for _, row in data.iterrows():
        try:
            giata = str(row[COL_MAP["giata"]]).strip()
            if not giata or giata in ("nan", "-", ""): continue
            try: float(giata)
            except: continue

            hotel = str(row[COL_MAP["hotel_name"]]).strip()
            if not hotel or hotel == "nan": continue

            dep_raw = row[COL_MAP["dep_date"]]
            dep_dt  = parse_date(str(dep_raw)) if not pd.isna(dep_raw) else None
            dep_month = dep_dt.strftime("%Y-%m") if dep_dt else ""
            dw = dep_window(dep_dt, run_dt)

            pp   = clean_num(row[COL_MAP["pp_price"]])
            cm   = clean_num(row[COL_MAP["current_margin"]])
            ma   = clean_num(row[COL_MAP["margin_after"]])
            cp   = clean_num(row[COL_MAP["cheapest_comp"]])
            dg   = clean_num(row[COL_MAP["diff_gbp"]])
            dp   = clean_num(row[COL_MAP["diff_pct"]])
            res  = str(row[COL_MAP["result"]]).strip()
            mr   = str(row[COL_MAP["margin_range"]]).strip()
            oproom = str(row[COL_MAP["op_room"]]).strip()
            board  = str(row[COL_MAP["board"]]).strip()
            nights = int(clean_num(row[COL_MAP["nights"]]) or 7)

            if res not in RESULT_VALS and "Win" not in res and "Lose" not in res: continue
            if pd.isna(pp) or pp == 0: continue

            mf = margin_flag(ma)

            rows.append({
                "file_name": fname,
                "file_date": str(file_dt),
                "destination": dest,
                "competitor": comp,
                "giata": giata,
                "hotel_name": hotel,
                "board": board,
                "dep_date": str(dep_dt) if dep_dt else "",
                "dep_month": dep_month,
                "dep_window": dw,
                "nights": nights,
                "pp_price": pp,
                "current_margin": cm if not pd.isna(cm) else 0,
                "op_room": oproom,
                "comp_price": cp if not pd.isna(cp) else 0,
                "diff_gbp": dg if not pd.isna(dg) else 0,
                "diff_pct": dp if not pd.isna(dp) else 0,
                "result": res,
                "margin_after": ma if not pd.isna(ma) else 0,
                "margin_range": mr,
                "margin_flag": mf,
                "booking_tier": "medium",   # updated after bookings loaded
                "priority_score": 0,        # calculated after insert
                "uploaded_at": datetime.now().isoformat(),
            })
        except Exception:
            continue

    return rows

def calc_priority(row, tier_weights={"high":3,"medium":2,"low":1}):
    tw  = tier_weights.get(row.get("booking_tier","low"), 1)
    gap = abs(row.get("diff_pct", 0))
    gs  = 1 if gap < 5 else 2 if gap < 10 else 3 if gap < 20 else 4
    mf  = row.get("margin_flag","ok")
    ms  = 0.1 if mf=="floor" else 0.5 if mf=="near" else 1.5 if mf=="ceiling" else 1.0
    dw  = row.get("dep_window","241+")
    di  = 1.0 if dw=="0-60" else 1.1 if dw=="61-120" else 1.3 if dw=="121-240" else 1.2
    return round(tw * gs * ms * di, 3)

def save_rows(rows):
    if not rows: return 0
    conn = get_conn()
    c = conn.cursor()
    inserted = 0
    for r in rows:
        r["priority_score"] = calc_priority(r)
        # De-duplicate: same file + hotel + dep_date + dep_window
        c.execute("""
            SELECT id FROM pricing
            WHERE file_name=? AND hotel_name=? AND dep_date=? AND dep_window=?
        """, (r["file_name"], r["hotel_name"], r["dep_date"], r["dep_window"]))
        if c.fetchone() is None:
            c.execute("""
                INSERT INTO pricing (
                    file_name,file_date,destination,competitor,giata,hotel_name,board,
                    dep_date,dep_month,dep_window,nights,pp_price,current_margin,op_room,
                    comp_price,diff_gbp,diff_pct,result,margin_after,margin_range,
                    margin_flag,booking_tier,priority_score,uploaded_at
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, tuple(r[k] for k in [
                "file_name","file_date","destination","competitor","giata","hotel_name","board",
                "dep_date","dep_month","dep_window","nights","pp_price","current_margin","op_room",
                "comp_price","diff_gbp","diff_pct","result","margin_after","margin_range",
                "margin_flag","booking_tier","priority_score","uploaded_at"
            ]))
            inserted += 1
    conn.commit()
    conn.close()
    return inserted

def load_data(dest_filter=None, comp_filter=None, window_filter=None,
              file_filter=None, date_from=None, date_to=None) -> pd.DataFrame:
    conn = get_conn()
    q = "SELECT * FROM pricing WHERE 1=1"
    params = []
    if dest_filter:
        q += f" AND destination IN ({','.join(['?']*len(dest_filter))})"
        params += dest_filter
    if comp_filter:
        q += f" AND competitor IN ({','.join(['?']*len(comp_filter))})"
        params += comp_filter
    if window_filter:
        q += f" AND dep_window IN ({','.join(['?']*len(window_filter))})"
        params += window_filter
    if file_filter:
        q += f" AND file_name IN ({','.join(['?']*len(file_filter))})"
        params += file_filter
    if date_from:
        q += " AND file_date >= ?"
        params.append(str(date_from))
    if date_to:
        q += " AND file_date <= ?"
        params.append(str(date_to))
    df = pd.read_sql_query(q, conn, params=params)
    conn.close()
    return df

def load_bookings() -> pd.DataFrame:
    conn = get_conn()
    df = pd.read_sql_query("SELECT * FROM bookings", conn)
    conn.close()
    return df

def get_distinct(col):
    conn = get_conn()
    rows = conn.execute(f"SELECT DISTINCT {col} FROM pricing ORDER BY {col}").fetchall()
    conn.close()
    return [r[0] for r in rows if r[0]]

def get_file_list():
    conn = get_conn()
    rows = conn.execute("SELECT DISTINCT file_name, file_date, destination, competitor, COUNT(*) as rows FROM pricing GROUP BY file_name ORDER BY file_date DESC").fetchall()
    conn.close()
    return rows

def apply_bookings_tiers(df: pd.DataFrame, bookings: pd.DataFrame) -> pd.DataFrame:
    if bookings.empty or df.empty:
        # volume proxy
        counts = df.groupby(["hotel_name","destination"])["hotel_name"].transform("count")
        p75 = counts.quantile(0.75)
        p25 = counts.quantile(0.25)
        df["booking_tier"] = np.where(counts >= p75, "high", np.where(counts >= p25, "medium", "low"))
        return df

    def match_hotel(name, bdf, thresh=80):
        best_score, best_bk = 0, None
        for _, brow in bdf.iterrows():
            score = fuzz.token_sort_ratio(name.lower(), str(brow["hotel_name"]).lower())
            if score > best_score:
                best_score, best_bk = score, brow["bkgs_4wk"]
        return best_bk if best_score >= thresh else None

    dest_groups = {}
    for dest, grp in df.groupby("destination"):
        b_dest = bookings[bookings["destination"]==dest] if "destination" in bookings.columns else bookings
        hotels = grp["hotel_name"].unique()
        hotel_bkgs = {}
        for h in hotels:
            bk = match_hotel(h, b_dest)
            hotel_bkgs[h] = bk if bk is not None else 0
        bkgs_series = grp["hotel_name"].map(hotel_bkgs).fillna(0)
        p75 = bkgs_series.quantile(0.75)
        p25 = bkgs_series.quantile(0.25)
        tier = np.where(bkgs_series >= p75, "high", np.where(bkgs_series >= p25, "medium", "low"))
        df.loc[grp.index, "booking_tier"] = tier
        dest_groups[dest] = hotel_bkgs
    return df

# ══════════════════════════════════════════════════════════════════════════════
# ANALYSIS HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def derive_stance(win_rate, avg_lose, avg_win):
    if win_rate < 0.45 or avg_lose < -10: return "Aggressively Reduce"
    if win_rate < 0.55 or avg_lose < -5:  return "Selective Adjust"
    if win_rate >= 0.65 and avg_win > 15: return "Increase Pricing"
    if win_rate >= 0.55:                  return "Hold — Optimise Margin"
    return "Product-Led Fix"

def margin_cell(cur, after):
    delta = after - cur
    flag  = margin_flag(after)
    label = {"floor":"⚠ BELOW FLOOR","near":"NEAR FLOOR","ceiling":"↑ ABOVE CEILING","ok":"","unknown":""}[flag]
    return f"{cur:.1f}% → {after:.1f}% {label} ({delta:+.1f}pp)"

def colour_diff(val):
    if val < -5: return "background-color:#FEE2E2"
    if val > 15: return "background-color:#D1FAE5"
    if val < 0:  return "background-color:#FEF3C7"
    return ""

def colour_margin(val):
    if val < MARGIN_FLOOR:  return "background-color:#FEE2E2"
    if val < MARGIN_NEAR:   return "background-color:#FEF3C7"
    if val > MARGIN_CEILING:return "background-color:#DBEAFE"
    return ""

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## ✈ D2 Pricing Intelligence")
    st.markdown("---")

    # ── Upload pricing files ─────────────────────────────────────────────────
    st.markdown("### 📂 Upload Pricing Files")
    uploaded_files = st.file_uploader(
        "Drop .xlsm / .xlsx / .csv files",
        type=["xlsm","xlsx","csv"],
        accept_multiple_files=True,
        key="pricing_upload"
    )
    if uploaded_files:
        for uf in uploaded_files:
            with st.spinner(f"Parsing {uf.name}…"):
                rows = parse_pricing_file(uf, uf.name)
                n = save_rows(rows)
            if n > 0:
                st.success(f"✅ {uf.name}: {n} new rows added")
            else:
                st.info(f"ℹ {uf.name}: already in database (no duplicates)")

    # ── Upload bookings ──────────────────────────────────────────────────────
    st.markdown("### 📊 Upload Bookings File")
    bookings_file = st.file_uploader(
        "CSV: hotel_name, destination, bkgs_4wk, bkgs_py (optional)",
        type=["csv","xlsx"],
        key="bookings_upload"
    )
    if bookings_file:
        try:
            if bookings_file.name.endswith(".csv"):
                bdf = pd.read_csv(bookings_file)
            else:
                bdf = pd.read_excel(bookings_file)
            bdf.columns = [c.lower().strip().replace(" ","_") for c in bdf.columns]
            conn = get_conn()
            for _, row in bdf.iterrows():
                conn.execute("""
                    INSERT INTO bookings (hotel_name,destination,bkgs_4wk,bkgs_py,uploaded_at)
                    VALUES (?,?,?,?,?)
                """, (
                    str(row.get("hotel_name","")),
                    str(row.get("destination","")),
                    int(row.get("bkgs_4wk",0) or 0),
                    int(row.get("bkgs_py",0) or 0),
                    datetime.now().isoformat()
                ))
            conn.commit()
            conn.close()
            st.success(f"✅ Bookings loaded: {len(bdf)} hotels")
        except Exception as e:
            st.error(f"Bookings error: {e}")

    st.markdown("---")
    # ── Filters ──────────────────────────────────────────────────────────────
    st.markdown("### 🔽 Filters")
    all_dests = get_distinct("destination")
    all_comps = get_distinct("competitor")
    all_files = [r[0] for r in get_file_list()]

    sel_dest   = st.multiselect("Destination", all_dests, default=all_dests)
    sel_comp   = st.multiselect("Competitor",  all_comps, default=all_comps)
    sel_window = st.multiselect("Dep Window",  ["0-60","61-120","121-240","241+"],
                                default=["0-60","61-120","121-240","241+"])
    sel_files  = st.multiselect("Files (runs)", all_files, default=all_files)
    st.markdown("---")
    if st.button("🗑 Clear ALL data", type="secondary"):
        conn = get_conn()
        conn.execute("DELETE FROM pricing")
        conn.execute("DELETE FROM bookings")
        conn.commit()
        conn.close()
        st.success("Database cleared")
        st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# LOAD & PREPARE DATA
# ══════════════════════════════════════════════════════════════════════════════
df_raw = load_data(
    dest_filter   = sel_dest   or None,
    comp_filter   = sel_comp   or None,
    window_filter = sel_window or None,
    file_filter   = sel_files  or None,
)

bookings_df = load_bookings()

if not df_raw.empty:
    df_raw = apply_bookings_tiers(df_raw, bookings_df)
    df_raw["priority_score"] = df_raw.apply(lambda r: calc_priority(r.to_dict()), axis=1)

comparable = df_raw[df_raw["result"].isin(["Win","Lose"])] if not df_raw.empty else pd.DataFrame()
wins_df    = df_raw[df_raw["result"]=="Win"]  if not df_raw.empty else pd.DataFrame()
loses_df   = df_raw[df_raw["result"]=="Lose"] if not df_raw.empty else pd.DataFrame()
nocomp_df  = df_raw[df_raw["result"].str.contains("No Comp", na=False)] if not df_raw.empty else pd.DataFrame()

# ══════════════════════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("# ✈ D2 Pricing Intelligence Dashboard")
dest_str = ", ".join(sel_dest) if sel_dest else "All destinations"
st.caption(f"Destinations: **{dest_str}** · Competitor: **{', '.join(sel_comp) if sel_comp else 'All'}** · Runs: {len(sel_files)} files loaded")

if df_raw.empty:
    st.info("👆 Upload pricing files using the sidebar to get started.")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
tabs = st.tabs([
    "1 Master Queue","2 Overview","3 Hotel Actions",
    "4 ↓ Reduce","5 ↑ Raise","6 Maintain",
    "7 Product Gaps","8 No Comp","9 Dest Actions",
    "10 ⚑ Outliers","11 Missing","12 Advantage",
    "13 Insights","14 ↗ Trends","15 📂 Files"
])

# ── TAB 1: Master Action Queue ────────────────────────────────────────────────
with tabs[0]:
    st.subheader("Master Action Queue — Top 30")
    actions = []
    for _, r in loses_df.iterrows():
        gap = abs(r["diff_pct"])
        act = ("⚠ SUPPRESS — floor breach" if r["margin_flag"]=="floor"
               else f"↓ Reduce ~£{abs(r['comp_price']-r['pp_price']):.0f}pp" if gap>5
               else "↓ Monitor")
        actions.append({**r.to_dict(), "action": act, "type": "reduce"})
    for _, r in wins_df[wins_df["diff_pct"]>RAISE_THRESHOLD].iterrows():
        actions.append({**r.to_dict(), "action": "↑ Margin opportunity", "type": "raise"})

    if actions:
        adf = pd.DataFrame(actions)
        tier_w = {"high":3,"medium":2,"low":1}
        adf["_tw"] = adf["booking_tier"].map(tier_w).fillna(1)
        adf = adf.sort_values(["_tw","priority_score","diff_pct"], ascending=[False,False,True]).head(30)
        adf["margin_display"] = adf.apply(lambda r: margin_cell(r["current_margin"],r["margin_after"]), axis=1)
        show = adf[["hotel_name","destination","dep_window","dep_month","action",
                     "pp_price","comp_price","diff_pct","margin_display","booking_tier"]].copy()
        show.columns = ["Hotel","Dest","Window","Month","Action","D2 £pp","Comp £pp","Diff %","Margin","Tier"]
        def style_row(row):
            if row["Diff %"] < -10: return ["background-color:#FEE2E2"]*len(row)
            if row["Diff %"] > 20:  return ["background-color:#D1FAE5"]*len(row)
            return [""]*len(row)
        st.dataframe(
            show.style.apply(style_row, axis=1).format({"D2 £pp":"{:.0f}","Comp £pp":"{:.0f}","Diff %":"{:.1f}%"}),
            use_container_width=True, height=550
        )
    else:
        st.info("No actions — upload data first.")

# ── TAB 2: Overview ───────────────────────────────────────────────────────────
with tabs[1]:
    total = len(comparable)
    win_n = len(wins_df); lose_n = len(loses_df); nc_n = len(nocomp_df)
    wr = win_n/total*100 if total>0 else 0
    avg_lose = loses_df["diff_pct"].mean() if not loses_df.empty else 0
    floor_n  = len(loses_df[loses_df["margin_flag"]=="floor"])

    c1,c2,c3,c4,c5,c6 = st.columns(6)
    c1.metric("Hotels", df_raw["hotel_name"].nunique())
    c2.metric("Win Rate", f"{wr:.0f}%", f"{win_n}W / {lose_n}L")
    c3.metric("Losses", lose_n, f"Avg {avg_lose:.1f}%")
    c4.metric("No Comp", nc_n)
    c5.metric("⚠ Floor Breaches", floor_n)
    c6.metric("Avg D2 Margin", f"{df_raw['current_margin'].mean():.1f}%")

    st.markdown("---")
    col1, col2 = st.columns(2)

    with col1:
        # Win/Lose by window
        wl = df_raw.groupby(["dep_window","result"]).size().reset_index(name="count")
        fig = px.bar(wl, x="dep_window", y="count", color="result",
                     color_discrete_map={"Win":"#0A7C4E","Lose":"#C0392B","Win - No Comp":"#D4AF37"},
                     title="Win / Lose / No-Comp by Departure Window",
                     category_orders={"dep_window":["0-60","61-120","121-240","241+"]})
        fig.update_layout(height=300, margin=dict(t=40,b=10))
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        # Win rate by destination
        dest_stats = []
        for dest, grp in comparable.groupby("destination"):
            w = (grp["result"]=="Win").sum()
            l = (grp["result"]=="Lose").sum()
            dest_stats.append({"Dest":dest,"WinRate":w/(w+l)*100 if (w+l)>0 else 0,"Wins":w,"Losses":l})
        if dest_stats:
            ddf = pd.DataFrame(dest_stats)
            fig2 = px.bar(ddf, x="Dest", y="WinRate", title="Win Rate by Destination (%)",
                          color="WinRate", color_continuous_scale=["#C0392B","#D4AF37","#0A7C4E"],
                          range_color=[30,80])
            fig2.update_layout(height=300, margin=dict(t=40,b=10))
            st.plotly_chart(fig2, use_container_width=True)

    # Top sellers table
    st.markdown("#### 🏆 Top Sellers")
    ts = df_raw.groupby("hotel_name").agg(
        Dest=("destination","first"),
        Rows=("hotel_name","count"),
        WinRate=("result", lambda x: f"{(x=='Win').sum()/max(len(x[x.isin(['Win','Lose'])]),1)*100:.0f}%"),
        AvgMargin=("current_margin","mean"),
        HasLoss=("result", lambda x: "⚠ Loss" if (x=="Lose").any() else "✓ OK")
    ).sort_values("Rows", ascending=False).head(15)
    ts["AvgMargin"] = ts["AvgMargin"].map("{:.1f}%".format)
    st.dataframe(ts, use_container_width=True)

# ── TAB 3: Hotel Actions ──────────────────────────────────────────────────────
with tabs[2]:
    st.subheader("Individual Hotel Actions")
    col1, col2, col3 = st.columns([3,2,2])
    search = col1.text_input("Search hotel", placeholder="e.g. Fushifaru…")
    res_filter = col2.selectbox("Result filter", ["All","Win","Lose","Win - No Comp"])
    win_filter = col3.multiselect("Window", ["0-60","61-120","121-240","241+"],
                                  default=["0-60","61-120","121-240","241+"])

    view = df_raw.copy()
    if search: view = view[view["hotel_name"].str.contains(search, case=False, na=False)]
    if res_filter != "All": view = view[view["result"]==res_filter]
    if win_filter: view = view[view["dep_window"].isin(win_filter)]
    tier_w = {"high":3,"medium":2,"low":1}
    view["_tw"] = view["booking_tier"].map(tier_w).fillna(1)
    view = view.sort_values(["_tw","priority_score"], ascending=False)

    def make_action(row):
        if row["result"]=="Lose":
            return "↓ Reduce" if abs(row["diff_pct"])>5 else "Monitor"
        if row["result"]=="Win" and row["diff_pct"]>RAISE_THRESHOLD: return "↑ Raise"
        if row["result"]=="Win": return "Maintain"
        return "Dest-guided"

    view["Action"] = view.apply(make_action, axis=1)
    view["Margin"] = view.apply(lambda r: margin_cell(r["current_margin"],r["margin_after"]), axis=1)
    show = view[["hotel_name","destination","dep_window","dep_month","board",
                 "pp_price","comp_price","diff_pct","result","op_room","Margin","Action","booking_tier"]].copy()
    show.columns = ["Hotel","Dest","Window","Month","Board","D2 £pp","Comp £pp","Diff %","Result","D2 Room","Margin","Action","Tier"]
    st.dataframe(
        show.style.map(lambda v: "color:#C0392B;font-weight:bold" if isinstance(v,str) and "Lose" in v else
                                 "color:#0A7C4E;font-weight:bold" if isinstance(v,str) and "Win" in v and "No" not in v else ""),
        use_container_width=True, height=600
    )

# ── TAB 4: Reduce ─────────────────────────────────────────────────────────────
with tabs[3]:
    st.subheader("↓ Price Reduction Candidates")
    candidates = loses_df[loses_df["diff_pct"] < 0].copy()
    floor_n = (candidates["margin_flag"]=="floor").sum()
    near_n  = (candidates["margin_flag"]=="near").sum()
    st.error(f"⚠ {floor_n} of {len(candidates)} proposed reductions would breach the 7% margin floor. {near_n} would enter the 7–8.5% caution band.")

    col1,col2 = st.columns([3,2])
    search4 = col1.text_input("Search hotel", key="s4")
    win4    = col2.multiselect("Window", ["0-60","61-120","121-240","241+"],
                                default=["0-60","61-120","121-240","241+"], key="w4")
    if search4: candidates = candidates[candidates["hotel_name"].str.contains(search4, case=False)]
    if win4: candidates = candidates[candidates["dep_window"].isin(win4)]
    candidates = candidates.sort_values(["booking_tier","diff_pct"], ascending=[False,True])
    candidates["Margin"] = candidates.apply(lambda r: margin_cell(r["current_margin"],r["margin_after"]), axis=1)
    candidates["Flag"]   = candidates["margin_flag"].map({"floor":"🔴 FLOOR — SUPPRESS","near":"🟡 Near Floor","ok":"✅ OK","ceiling":"↑ Ceiling","unknown":"?"})
    candidates["Suggested"] = candidates.apply(lambda r:
        "SUPPRESS — floor breach" if r["margin_flag"]=="floor"
        else f"↓ Reduce ~£{abs(r['comp_price']-r['pp_price']):.0f}pp", axis=1)
    show4 = candidates[["hotel_name","destination","dep_window","dep_month","pp_price","comp_price","diff_pct","Margin","Flag","Suggested","booking_tier"]].copy()
    show4.columns = ["Hotel","Dest","Window","Month","D2 £pp","Comp £pp","Diff %","Margin","Flag","Action","Tier"]
    st.dataframe(
        show4.style.map(lambda v: "background-color:#FEE2E2" if isinstance(v,str) and "FLOOR" in v else ""),
        use_container_width=True, height=550
    )

# ── TAB 5: Raise ──────────────────────────────────────────────────────────────
with tabs[4]:
    st.subheader("↑ Price Raise / Margin Opportunities")
    raise_df = wins_df[wins_df["diff_pct"] > RAISE_THRESHOLD].copy()
    above_ceil = (raise_df["margin_flag"]=="ceiling").sum()
    st.success(f"✅ {len(raise_df)} hotels winning by >15% — potential margin uplift. {above_ceil} already above 15% ceiling.")
    raise_df = raise_df.sort_values(["booking_tier","diff_pct"], ascending=[False,False])
    raise_df["Margin"] = raise_df.apply(lambda r: margin_cell(r["current_margin"],r["margin_after"]), axis=1)
    raise_df["Ceiling"] = raise_df["margin_flag"].map({"ceiling":"🔵 ABOVE CEILING","ok":"✅ Within ceiling","near":"🟡 Near floor","floor":"🔴 Below floor","unknown":"?"})
    show5 = raise_df[["hotel_name","destination","dep_window","dep_month","pp_price","comp_price","diff_pct","Margin","Ceiling","booking_tier"]].copy()
    show5.columns = ["Hotel","Dest","Window","Month","D2 £pp","Comp £pp","Win %","Margin","Ceiling","Tier"]
    st.dataframe(show5.style.format({"Win %":"{:.1f}%","D2 £pp":"{:.0f}","Comp £pp":"{:.0f}"}),
                 use_container_width=True, height=500)

# ── TAB 6: Maintain ───────────────────────────────────────────────────────────
with tabs[5]:
    st.subheader("Maintain — Competitive Hold Band (−5% to +8%)")
    mtn = wins_df[(wins_df["diff_pct"]>=-5)&(wins_df["diff_pct"]<=HOLD_UPPER)].copy()
    mtn["MarginFlag"] = mtn["margin_flag"].map({"floor":"🔴 BELOW FLOOR","near":"🟡 Near floor","ok":"✅ OK","ceiling":"↑ Ceiling","unknown":"?"})
    show6 = mtn[["hotel_name","destination","dep_window","dep_month","pp_price","comp_price","diff_pct","current_margin","MarginFlag","booking_tier"]].copy()
    show6.columns = ["Hotel","Dest","Window","Month","D2 £pp","Comp £pp","Diff %","Cur Margin %","Flag","Tier"]
    st.dataframe(show6.style.format({"Diff %":"{:.1f}%","Cur Margin %":"{:.1f}%","D2 £pp":"{:.0f}","Comp £pp":"{:.0f}"}),
                 use_container_width=True, height=500)

# ── TAB 7: Product Gaps ───────────────────────────────────────────────────────
with tabs[6]:
    st.subheader("Product Gaps")
    st.warning("⚠ Product Gap = competitor offering a different room tier. These are NEVER price cut candidates — require product/contracting action.")
    pg = loses_df[loses_df["op_room"].str.lower().str.contains(
        "suite|villa|overwater|pool villa|penthouse|bungalow suite", na=False, regex=True)].copy()
    if pg.empty:
        pg = loses_df.copy()
    show7 = pg[["hotel_name","destination","dep_window","op_room","pp_price","comp_price","diff_pct","current_margin","booking_tier"]].copy()
    show7.columns = ["Hotel","Dest","Window","D2 Room","D2 £pp","Comp £pp","Diff %","Margin","Tier"]
    show7["Required Action"] = "Product Fix — not price cut"
    st.dataframe(show7, use_container_width=True, height=500)

# ── TAB 8: No Comp ────────────────────────────────────────────────────────────
with tabs[7]:
    st.subheader("No Direct Comparison — Destination-Guided")
    st.info("📌 These hotels have no competitor comparable. They are guided by destination stance — never 'price freely'.")
    # Derive stance per dest/window
    stances = {}
    for (dest, win), grp in comparable.groupby(["destination","dep_window"]):
        wc = (grp["result"]=="Win").sum()
        lc = (grp["result"]=="Lose").sum()
        wr = wc/(wc+lc) if (wc+lc)>0 else 0
        al = grp[grp["result"]=="Lose"]["diff_pct"].mean() if lc>0 else 0
        aw = grp[grp["result"]=="Win"]["diff_pct"].mean() if wc>0 else 0
        stances[(dest,win)] = derive_stance(wr, al, aw)

    nc = nocomp_df.copy()
    nc["Dest Stance"] = nc.apply(lambda r: stances.get((r["destination"],r["dep_window"]),"Hold — Optimise Margin"), axis=1)
    show8 = nc[["hotel_name","destination","dep_window","dep_month","pp_price","current_margin","Dest Stance","booking_tier"]].copy()
    show8.columns = ["Hotel","Dest","Window","Month","D2 £pp","Margin","Dest Stance","Tier"]
    st.dataframe(show8, use_container_width=True, height=500)

# ── TAB 9: Destination Actions ────────────────────────────────────────────────
with tabs[8]:
    st.subheader("Destination × Departure Window — Pricing Stances")
    rows9 = []
    for (dest, win), grp in comparable.groupby(["destination","dep_window"]):
        wc = (grp["result"]=="Win").sum()
        lc = (grp["result"]=="Lose").sum()
        nc_c = nocomp_df[(nocomp_df["destination"]==dest)&(nocomp_df["dep_window"]==win)].shape[0]
        wr = wc/(wc+lc) if (wc+lc)>0 else 0
        al = grp[grp["result"]=="Lose"]["diff_pct"].mean() if lc>0 else 0
        aw = grp[grp["result"]=="Win"]["diff_pct"].mean() if wc>0 else 0
        rows9.append({
            "Destination":dest,"Window":win,
            "Win Rate":f"{wr*100:.0f}%","Avg Lose":f"{al:.1f}%","Avg Win":f"{aw:.1f}%",
            "Stance":derive_stance(wr,al,aw),
            "Wins":wc,"Losses":lc,"No Comp":nc_c
        })
    if rows9:
        d9 = pd.DataFrame(rows9)
        st.dataframe(d9, use_container_width=True, height=500)

# ── TAB 10: Outliers ─────────────────────────────────────────────────────────
with tabs[9]:
    st.subheader("⚑ Outliers")
    by_hotel = comparable.groupby(["hotel_name","destination","dep_window"])
    cl, bw = [], []
    for (h,d,w), grp in by_hotel:
        wc = (grp["result"]=="Win").sum()
        lc = (grp["result"]=="Lose").sum()
        if lc >= 2 and wc == 0:
            cl.append({"Hotel":h,"Dest":d,"Window":w,"Losses":lc,"Avg Gap":f"{grp[grp['result']=='Lose']['diff_pct'].mean():.1f}%","Floor":grp["margin_flag"].eq("floor").any()})
        if wc >= 2:
            avg_win = grp[grp["result"]=="Win"]["diff_pct"].mean()
            if avg_win > 30:
                bw.append({"Hotel":h,"Dest":d,"Window":w,"Wins":wc,"Avg Win":f"+{avg_win:.1f}%"})

    if cl:
        st.error(f"🔴 {len(cl)} consistent loser situations identified")
        cldf = pd.DataFrame(cl)
        cldf["Floor Breach"] = cldf["Floor"].map({True:"⚠ YES", False:"No"})
        st.dataframe(cldf[["Hotel","Dest","Window","Losses","Avg Gap","Floor Breach"]], use_container_width=True)
    if bw:
        st.success(f"🔵 {len(bw)} big win margin opportunities")
        st.dataframe(pd.DataFrame(bw), use_container_width=True)
    if not cl and not bw:
        st.info("No severe outliers detected in current filter.")

# ── TAB 11: Missing ──────────────────────────────────────────────────────────
with tabs[10]:
    st.subheader("Missing Properties")
    st.info("Upload multiple files for cross-run missing property detection. Below shows hotels present in some runs but not others.")
    if len(sel_files) > 1:
        hotel_files = df_raw.groupby("hotel_name")["file_name"].nunique()
        missing = hotel_files[hotel_files < len(sel_files)].reset_index()
        missing.columns = ["Hotel","Files Present"]
        missing["Files Missing"] = len(sel_files) - missing["Files Present"]
        st.dataframe(missing.sort_values("Files Missing", ascending=False), use_container_width=True)
    else:
        st.write("Load 2+ files to detect missing properties across runs.")

# ── TAB 12: Competitive Advantage ────────────────────────────────────────────
with tabs[11]:
    st.subheader("Competitive Advantage — D2 Strong Win Positions")
    adv = wins_df[wins_df["diff_pct"] > HOLD_UPPER].groupby(["hotel_name","destination"]).agg(
        Count=("diff_pct","count"),
        AvgWin=("diff_pct","mean"),
        AvgMargin=("current_margin","mean")
    ).reset_index().sort_values("AvgWin", ascending=False)
    adv["AvgWin"]    = adv["AvgWin"].map("{:+.1f}%".format)
    adv["AvgMargin"] = adv["AvgMargin"].map("{:.1f}%".format)
    adv.columns = ["Hotel","Destination","Winning Records","Avg Win Gap","Avg Margin"]
    adv["Recommendation"] = "Consider selective margin increase"
    st.dataframe(adv, use_container_width=True, height=500)

# ── TAB 13: Strategic Insights ───────────────────────────────────────────────
with tabs[12]:
    st.subheader("Strategic Insights")
    floor_n  = loses_df[loses_df["margin_flag"]=="floor"].shape[0]
    avg_marg = df_raw["current_margin"].mean()

    # Critical losses
    top_losers = loses_df.nsmallest(8,"diff_pct")[["hotel_name","destination","dep_window","pp_price","comp_price","diff_pct","margin_after","margin_flag"]].copy()
    top_losers.columns = ["Hotel","Dest","Window","D2 £pp","Comp £pp","Gap %","Margin After","Flag"]

    st.error("### 🔴 Critical Losses")
    st.dataframe(top_losers.style.format({"Gap %":"{:.1f}%","D2 £pp":"{:.0f}","Comp £pp":"{:.0f}","Margin After":"{:.1f}%"}), use_container_width=True)

    st.warning(f"### ⚠ Margin Guardrails\n- **{floor_n} hotels** breach 7% floor if prices matched\n- Average D2 margin: **{avg_marg:.1f}%**\n- Margin floor: 7% · Ceiling: 15% · Caution: 7–8.5%")

    big_wins = wins_df[wins_df["diff_pct"]>20].nlargest(5,"diff_pct")[["hotel_name","destination","dep_window","diff_pct","current_margin"]].copy()
    big_wins.columns = ["Hotel","Dest","Window","Win %","Margin"]
    st.success("### ✅ Top Margin Opportunities")
    st.dataframe(big_wins.style.format({"Win %":"{:.1f}%","Margin":"{:.1f}%"}), use_container_width=True)

    # Window split
    st.info("### 🗓 Departure Window Split")
    for w in ["0-60","61-120","121-240","241+"]:
        wg = comparable[comparable["dep_window"]==w]
        if wg.empty: continue
        wr = (wg["result"]=="Win").sum()/len(wg)*100
        st.write(f"**{w}d** — Win rate: {wr:.0f}% ({(wg['result']=='Win').sum()}W / {(wg['result']=='Lose').sum()}L)")

    # Dynamic pricing rules
    with st.expander("📋 Dynamic Pricing Rules"):
        st.markdown("""
- **Target band:** −5% to +8% on like-for-like rooms
- **Room match first** — verify room tier before any price action
- **No-comp hotels** — destination-guided, not 'price freely'
- **Margin floor 7%** — suppress or flag; never recommend without sign-off
- **Margin ceiling 15%** — flag raise actions above ceiling
- **Booking volume** — top sellers treated first in every table
- **Product gaps** — never convert to price cut; product fix required
        """)

# ── TAB 14: Price Trends ─────────────────────────────────────────────────────
with tabs[13]:
    st.subheader("↗ Price Trends")
    col1, col2 = st.columns([3,1])
    hotel_list = sorted(df_raw["hotel_name"].unique().tolist())
    # Default to worst loser
    worst = loses_df.nsmallest(1,"diff_pct")["hotel_name"].values[0] if not loses_df.empty else hotel_list[0]
    sel_hotel = col1.selectbox("Select hotel (type to search)", hotel_list,
                                index=hotel_list.index(worst) if worst in hotel_list else 0)
    sel_board = col2.selectbox("Board", ["All"] + sorted(df_raw[df_raw["hotel_name"]==sel_hotel]["board"].unique().tolist()))

    h_data = df_raw[df_raw["hotel_name"]==sel_hotel].copy()
    if sel_board != "All":
        h_data = h_data[h_data["board"]==sel_board]

    if not h_data.empty:
        # Seasonal curve
        monthly = h_data.groupby("dep_month").agg(
            D2_pp=("pp_price","mean"),
            Comp_pp=("comp_price", lambda x: x[x>0].mean() if (x>0).any() else np.nan)
        ).reset_index().sort_values("dep_month")
        monthly["Month"] = monthly["dep_month"].apply(lambda m: pd.to_datetime(m).strftime("%b %y") if m else m)

        fig = go.Figure()
        fig.add_trace(go.Scatter(x=monthly["Month"], y=monthly["D2_pp"],
                                  name="D2 £pp", line=dict(color="#1B1464",width=3)))
        if monthly["Comp_pp"].notna().any():
            fig.add_trace(go.Scatter(x=monthly["Month"], y=monthly["Comp_pp"],
                                      name="Competitor £pp", line=dict(color="#C17A00",width=2,dash="dash")))
        fig.update_layout(title=f"Seasonal Curve — {sel_hotel}", height=320,
                          yaxis_title="£ per person", legend=dict(orientation="h"))
        st.plotly_chart(fig, use_container_width=True)

    # Destination-level win rate by month
    st.markdown("#### Win Rate by Departure Month — All Hotels")
    dest_filter2 = st.selectbox("Destination", sorted(comparable["destination"].unique()) if not comparable.empty else ["MLE"], key="tr_dest")
    wr_data = comparable[comparable["destination"]==dest_filter2].copy()
    if not wr_data.empty:
        wr_monthly = wr_data.groupby("dep_month").apply(
            lambda g: (g["result"]=="Win").sum() / len(g) * 100
        ).reset_index()
        wr_monthly.columns = ["Month","Win Rate %"]
        wr_monthly["Month_label"] = wr_monthly["Month"].apply(lambda m: pd.to_datetime(m).strftime("%b %y") if m else m)
        wr_monthly = wr_monthly.sort_values("Month")
        fig2 = go.Figure()
        # Red shading below 50%
        for _, row in wr_monthly.iterrows():
            if row["Win Rate %"] < 50:
                fig2.add_shape(type="rect", x0=row["Month_label"], x1=row["Month_label"],
                               y0=0, y1=50, fillcolor="rgba(192,57,43,0.08)", line_width=0)
        fig2.add_trace(go.Scatter(
            x=wr_monthly["Month_label"], y=wr_monthly["Win Rate %"],
            mode="lines+markers", name="Win Rate %",
            line=dict(color="#1B1464",width=2.5),
            marker=dict(color=wr_monthly["Win Rate %"].apply(lambda v: "#C0392B" if v<50 else "#1B1464"),size=8)
        ))
        fig2.add_hline(y=50, line_dash="dot", line_color="rgba(192,57,43,0.5)", annotation_text="50%")
        fig2.update_layout(height=280, yaxis=dict(range=[0,100],title="Win Rate %"),
                           xaxis_title="Departure Month")
        st.plotly_chart(fig2, use_container_width=True)

    # Price drift (multi-run)
    if df_raw["file_date"].nunique() > 1:
        st.markdown("#### Price Drift — Multi-Run Comparison")
        drift = df_raw[df_raw["hotel_name"]==sel_hotel].groupby("file_date").agg(
            D2_avg=("pp_price","mean"),
            Comp_avg=("comp_price", lambda x: x[x>0].mean() if (x>0).any() else np.nan)
        ).reset_index().sort_values("file_date")
        fig3 = go.Figure()
        fig3.add_trace(go.Scatter(x=drift["file_date"],y=drift["D2_avg"],name="D2 avg £pp",line=dict(color="#1B1464",width=2)))
        if drift["Comp_avg"].notna().any():
            fig3.add_trace(go.Scatter(x=drift["file_date"],y=drift["Comp_avg"],name="Comp avg £pp",line=dict(color="#C17A00",width=2,dash="dash")))
        fig3.update_layout(title=f"Price Drift — {sel_hotel}",height=260,xaxis_title="Run Date",yaxis_title="£ pp")
        st.plotly_chart(fig3, use_container_width=True)

# ── TAB 15: Files ─────────────────────────────────────────────────────────────
with tabs[14]:
    st.subheader("📂 Uploaded Files Database")
    files = get_file_list()
    if files:
        fdf = pd.DataFrame(files, columns=["File Name","Run Date","Destination","Competitor","Rows"])
        st.dataframe(fdf, use_container_width=True)
        total_rows = fdf["Rows"].sum()
        st.caption(f"Total: {len(fdf)} files · {total_rows:,} pricing rows in database")

        # Download all data
        csv_data = df_raw.to_csv(index=False)
        st.download_button("⬇ Download all filtered data as CSV",
                           data=csv_data, file_name="d2_pricing_export.csv", mime="text/csv")
    else:
        st.info("No files uploaded yet.")
