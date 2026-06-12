"""
D2 vs Competitor — Multi-Destination Pricing Dashboard
Streamlit app · upload .xlsm/.xlsx/.csv files · persistent SQLite DB
"""

import streamlit as st
import pandas as pd
import numpy as np
import sqlite3
import io
import re
from datetime import datetime, date
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

try:
    from rapidfuzz import fuzz
    HAS_FUZZ = True
except ImportError:
    HAS_FUZZ = False

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="D2 Pricing Intelligence",
    page_icon="✈",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Constants ──────────────────────────────────────────────────────────────────
DB_PATH            = "pricing_data.db"
MARGIN_FLOOR       = 7.0
MARGIN_CEILING     = 15.0
MARGIN_NEAR        = 8.5
LOSE_THRESHOLD     = -5.0
RAISE_THRESHOLD    = 15.0
HOLD_UPPER         = 8.0
RESULT_VALS        = ["Win","Lose","Win - No Comp","Win Aft Change"]

# ── CSS ────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stSidebar"]{background:#1B1464}
[data-testid="stSidebar"] *{color:#fff!important}
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label{color:#D4AF37!important}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# DATABASE
# ══════════════════════════════════════════════════════════════════════════════
def get_conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)

def init_db():
    c = get_conn()
    c.execute("""CREATE TABLE IF NOT EXISTS pricing(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        file_name TEXT, file_date TEXT, destination TEXT, competitor TEXT,
        giata TEXT, hotel_name TEXT, board TEXT,
        dep_date TEXT, dep_month TEXT, dep_window TEXT, nights INTEGER,
        pp_price REAL, current_margin REAL, op_room TEXT,
        comp_price REAL, diff_gbp REAL, diff_pct REAL,
        result TEXT, margin_after REAL, margin_range TEXT,
        margin_flag TEXT, booking_tier TEXT, priority_score REAL,
        uploaded_at TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS bookings(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        hotel_name TEXT, destination TEXT,
        bkgs_4wk INTEGER, bkgs_py INTEGER, uploaded_at TEXT)""")
    c.commit(); c.close()

init_db()

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def parse_date(val):
    if pd.isna(val): return None
    s = str(val).strip()
    for fmt in ("%d/%m/%Y","%Y-%m-%d","%d-%m-%Y","%m/%d/%Y"):
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
    if val is None or (isinstance(val, float) and np.isnan(val)): return np.nan
    s = str(val).replace("£","").replace(",","").replace(" ","").strip()
    try: return float(s)
    except: return np.nan

def extract_meta(fname):
    """D2_Live_MLE_LH_08_06_2026 → MLE, LH, 2026-06-08"""
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

# ══════════════════════════════════════════════════════════════════════════════
# PARSER  (robust — handles xlsm with multiple sheets)
# ══════════════════════════════════════════════════════════════════════════════
def find_header_row(df_raw):
    """Find the row index that contains column headers."""
    for i in range(min(10, len(df_raw))):
        try:
            vals = [str(v).lower() for v in df_raw.iloc[i].tolist()]
            row_str = " ".join(vals)
            if "giata" in row_str or ("hotel" in row_str and "price" in row_str):
                return i
        except Exception:
            continue
    return 0

def find_data_sheets(xl):
    """Return list of (sheet_name, df) that look like pricing sheets."""
    results = []
    for sname in xl.sheet_names:
        # Skip admin sheets
        skip_keywords = ["board","namegiata","summary","pivot","sheet","comp data",
                         "exclusiv","direct","lh only","tr only","otb only"]
        if any(k in sname.lower() for k in skip_keywords):
            continue
        try:
            raw = xl.parse(sname, header=None, dtype=str)
            if raw.shape[1] < 18 or raw.shape[0] < 5:
                continue
            # Check if any cell in first 10 rows contains a numeric giata-like value
            sample = raw.iloc[:10, :5].values.flatten()
            has_numeric = any(
                str(v).strip().isdigit() and len(str(v).strip()) >= 4
                for v in sample if v and str(v) != "nan"
            )
            if has_numeric:
                results.append((sname, raw))
        except Exception:
            continue
    return results

def parse_pricing_file(uploaded_file, fname):
    dest, comp, file_dt = extract_meta(fname)
    all_rows = []

    try:
        if fname.lower().endswith(".csv"):
            raw = pd.read_csv(uploaded_file, header=None, dtype=str, encoding="utf-8-sig")
            sheets = [("csv", raw)]
        else:
            xl = pd.ExcelFile(uploaded_file, engine="openpyxl")
            sheets = find_data_sheets(xl)
            if not sheets:
                # fallback: try all sheets
                sheets = []
                for sn in xl.sheet_names:
                    try:
                        raw = xl.parse(sn, header=None, dtype=str)
                        if raw.shape[1] >= 18:
                            sheets.append((sn, raw))
                    except: pass
    except Exception as e:
        st.error(f"Cannot open {fname}: {e}")
        return []

    for sname, raw in sheets:
        hrow = find_header_row(raw)
        # Use row hrow as header, data starts hrow+1
        data = raw.iloc[hrow+1:].reset_index(drop=True).copy()
        data.columns = range(data.shape[1])

        for _, row in data.iterrows():
            try:
                giata = str(row.get(0,"")).strip()
                if not giata or giata in ("nan","-","","None"): continue
                try: float(giata)
                except: continue

                hotel = str(row.get(1,"")).strip()
                if not hotel or hotel in ("nan","None",""): continue

                board  = str(row.get(2,"")).strip()
                dep_raw= row.get(3,"")
                dep_dt = parse_date(dep_raw)
                nights = int(clean_num(row.get(4,7)) or 7)

                pp  = clean_num(row.get(5))
                cm  = clean_num(row.get(6))
                oproom = str(row.get(7,"")).strip()

                # Competitor price — try col 18 first (cheapest comp), fallback to 15/16/17
                cp = clean_num(row.get(18))
                if pd.isna(cp) or cp == 0:
                    for ci in [15,16,17]:
                        v = clean_num(row.get(ci))
                        if not pd.isna(v) and v > 0:
                            cp = v; break

                diff_g = clean_num(row.get(19))
                diff_p = clean_num(row.get(20))
                result = str(row.get(21,"")).strip()
                ma     = clean_num(row.get(22))
                mr     = str(row.get(23,"")).strip()

                if pd.isna(pp) or pp == 0: continue
                # Normalise result
                if "no comp" in result.lower(): result = "Win - No Comp"
                elif "aft change" in result.lower(): result = "Win Aft Change"
                elif "win" in result.lower(): result = "Win"
                elif "lose" in result.lower() or "loss" in result.lower(): result = "Lose"
                else: continue

                dep_month = dep_dt.strftime("%Y-%m") if dep_dt else ""
                dw = dep_window(dep_dt, file_dt) if dep_dt else "241+"
                mf = margin_flag(ma if not pd.isna(ma) else 0)

                all_rows.append({
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
                    "pp_price": float(pp),
                    "current_margin": float(cm) if not pd.isna(cm) else 0.0,
                    "op_room": oproom,
                    "comp_price": float(cp) if not pd.isna(cp) else 0.0,
                    "diff_gbp": float(diff_g) if not pd.isna(diff_g) else 0.0,
                    "diff_pct": float(diff_p) if not pd.isna(diff_p) else 0.0,
                    "result": result,
                    "margin_after": float(ma) if not pd.isna(ma) else 0.0,
                    "margin_range": mr,
                    "margin_flag": mf,
                    "booking_tier": "medium",
                    "priority_score": 0.0,
                    "uploaded_at": datetime.now().isoformat(),
                })
            except Exception:
                continue

    return all_rows

def calc_priority(row):
    tw = {"high":3,"medium":2,"low":1}.get(row.get("booking_tier","low"),1)
    gap= abs(row.get("diff_pct",0))
    gs = 1 if gap<5 else 2 if gap<10 else 3 if gap<20 else 4
    mf = row.get("margin_flag","ok")
    ms = 0.1 if mf=="floor" else 0.5 if mf=="near" else 1.5 if mf=="ceiling" else 1.0
    dw = row.get("dep_window","241+")
    di = 1.0 if dw=="0-60" else 1.1 if dw=="61-120" else 1.3 if dw=="121-240" else 1.2
    return round(tw*gs*ms*di,3)

def save_rows(rows):
    if not rows: return 0
    conn = get_conn(); c = conn.cursor(); inserted = 0
    for r in rows:
        r["priority_score"] = calc_priority(r)
        c.execute("""SELECT id FROM pricing
            WHERE file_name=? AND hotel_name=? AND dep_date=? AND dep_window=?""",
            (r["file_name"],r["hotel_name"],r["dep_date"],r["dep_window"]))
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
            inserted += 1
    conn.commit(); conn.close()
    return inserted

def load_data(dest=None, comp=None, window=None, files=None):
    conn = get_conn()
    q = "SELECT * FROM pricing WHERE 1=1"; p = []
    if dest:   q+=f" AND destination IN ({','.join(['?']*len(dest))})";   p+=dest
    if comp:   q+=f" AND competitor IN ({','.join(['?']*len(comp))})";    p+=comp
    if window: q+=f" AND dep_window IN ({','.join(['?']*len(window))})";  p+=window
    if files:  q+=f" AND file_name IN ({','.join(['?']*len(files))})";    p+=files
    df = pd.read_sql_query(q, conn, params=p)
    conn.close()

    if df.empty:
        return df

    # Ensure numeric columns are properly typed after SQLite load
    num_cols = ["pp_price","comp_price","current_margin","margin_after",
                "diff_pct","diff_gbp","priority_score"]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    # Parse dep_date back to proper date strings, and build dep_date_dt
    df["dep_date"] = df["dep_date"].astype(str).str.strip()

    # Remove any rows where dep_date looks invalid
    df = df[df["dep_date"].str.match(r"\d{4}-\d{2}-\d{2}", na=False) |
            df["dep_date"].str.match(r"\d{2}/\d{2}/\d{4}", na=False)]

    # Normalise dep_month — recalculate from dep_date to be safe
    def safe_dep_month(d):
        try:
            return pd.to_datetime(d).strftime("%Y-%m")
        except:
            return ""
    df["dep_month"] = df["dep_date"].apply(safe_dep_month)

    # Filter out rows with zero/missing prices
    df = df[df["pp_price"] > 0]

    return df

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

def load_bookings():
    conn = get_conn()
    df = pd.read_sql_query("SELECT * FROM bookings", conn)
    conn.close()
    return df

def apply_tiers(df, bookings):
    if df.empty: return df
    if not bookings.empty and HAS_FUZZ:
        def get_bk(name, bdf):
            best, val = 0, 0
            for _, br in bdf.iterrows():
                sc = fuzz.token_sort_ratio(name.lower(), str(br["hotel_name"]).lower())
                if sc > best: best, val = sc, int(br.get("bkgs_4wk",0) or 0)
            return val if best >= 80 else 0
        for dest, grp in df.groupby("destination"):
            bd = bookings[bookings["destination"]==dest] if "destination" in bookings.columns else bookings
            bkgs = grp["hotel_name"].apply(lambda h: get_bk(h, bd))
            p75, p25 = bkgs.quantile(0.75), bkgs.quantile(0.25)
            tier = np.where(bkgs>=p75,"high",np.where(bkgs>=p25,"medium","low"))
            df.loc[grp.index,"booking_tier"] = tier
    else:
        counts = df.groupby(["hotel_name","destination"])["hotel_name"].transform("count")
        p75, p25 = counts.quantile(0.75), counts.quantile(0.25)
        df["booking_tier"] = np.where(counts>=p75,"high",np.where(counts>=p25,"medium","low"))
    return df

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
                rows = parse_pricing_file(uf, uf.name)
            n = save_rows(rows)
            if n > 0:   st.success(f"✅ {uf.name}: {n} rows added")
            else:        st.info(f"ℹ {uf.name}: already in database")

    st.markdown("### 📊 Upload Bookings")
    bfile = st.file_uploader("CSV: hotel_name, destination, bkgs_4wk",
        type=["csv","xlsx"], key="bu")
    if bfile:
        try:
            bdf = pd.read_csv(bfile) if bfile.name.endswith(".csv") else pd.read_excel(bfile)
            bdf.columns = [c.lower().strip().replace(" ","_") for c in bdf.columns]
            conn = get_conn()
            for _, r in bdf.iterrows():
                conn.execute("INSERT INTO bookings(hotel_name,destination,bkgs_4wk,bkgs_py,uploaded_at) VALUES(?,?,?,?,?)",
                    (str(r.get("hotel_name","")), str(r.get("destination","")),
                     int(r.get("bkgs_4wk",0) or 0), int(r.get("bkgs_py",0) or 0),
                     datetime.now().isoformat()))
            conn.commit(); conn.close()
            st.success(f"✅ {len(bdf)} hotels loaded")
        except Exception as e: st.error(f"Bookings error: {e}")

    st.markdown("---")
    st.markdown("### 🔽 Filters")
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
        conn.commit(); conn.close()
        st.success("Cleared"); st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# LOAD DATA
# ══════════════════════════════════════════════════════════════════════════════
df = load_data(dest=sel_d or None, comp=sel_c or None,
               window=sel_w or None, files=sel_f or None)
bk = load_bookings()
if not df.empty:
    df = apply_tiers(df, bk)
    df["priority_score"] = df.apply(lambda r: calc_priority(r.to_dict()), axis=1)

comp_df  = df[df["result"].isin(["Win","Lose"])]           if not df.empty else pd.DataFrame()
wins_df  = df[df["result"]=="Win"]                          if not df.empty else pd.DataFrame()
loses_df = df[df["result"]=="Lose"]                         if not df.empty else pd.DataFrame()
nc_df    = df[df["result"].str.contains("No Comp",na=False)]if not df.empty else pd.DataFrame()

st.title("✈ D2 Pricing Intelligence Dashboard")
if sel_d: st.caption(f"Destinations: **{', '.join(sel_d)}** · Competitor: **{', '.join(sel_c or [])}** · {len(sel_f)} files")

if df.empty:
    st.info("👆 Upload pricing files using the sidebar to get started.")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
tab_names = ["⚡ Master Queue","Overview","Hotel Actions",
             "↓ Reduce","↑ Raise","Maintain",
             "Product Gaps","No Comp","Dest Actions",
             "⚑ Outliers","Missing","Advantage",
             "Insights","↗ Price Trends","📂 Files"]
tabs = st.tabs(tab_names)

# ── 1. MASTER QUEUE ────────────────────────────────────────────────────────────
with tabs[0]:
    st.subheader("Master Action Queue")
    acts = []
    for _, r in loses_df.iterrows():
        gap = abs(r["diff_pct"])
        act = ("⚠ SUPPRESS — floor breach" if r["margin_flag"]=="floor"
               else f"↓ Reduce ~£{abs(r['comp_price']-r['pp_price']):.0f}pp" if gap>5
               else "↓ Monitor")
        acts.append({**r.to_dict(),"action":act})
    for _, r in wins_df[wins_df["diff_pct"]>RAISE_THRESHOLD].iterrows():
        acts.append({**r.to_dict(),"action":"↑ Margin opportunity"})

    if acts:
        adf = pd.DataFrame(acts)
        tw  = {"high":3,"medium":2,"low":1}
        adf["_tw"] = adf["booking_tier"].map(tw).fillna(1)
        adf = adf.sort_values(["_tw","priority_score","diff_pct"],
                               ascending=[False,False,True]).head(30)
        adf["Margin"] = adf.apply(lambda r: margin_cell(r["current_margin"],r["margin_after"]),axis=1)
        show = adf[["hotel_name","destination","dep_window","dep_month",
                     "action","pp_price","comp_price","diff_pct","Margin","booking_tier"]].copy()
        show.columns = ["Hotel","Dest","Window","Month","Action","D2 £pp","Comp £pp","Diff %","Margin","Tier"]
        st.dataframe(show.style.format({"D2 £pp":"{:.0f}","Comp £pp":"{:.0f}","Diff %":"{:.1f}%"}),
                     use_container_width=True, height=550)

# ── 2. OVERVIEW ────────────────────────────────────────────────────────────────
with tabs[1]:
    total = len(comp_df)
    wn, ln, ncn = len(wins_df), len(loses_df), len(nc_df)
    wr  = wn/total*100 if total>0 else 0
    alg = loses_df["diff_pct"].mean() if not loses_df.empty else 0
    fln = (loses_df["margin_flag"]=="floor").sum() if not loses_df.empty else 0

    c1,c2,c3,c4,c5,c6 = st.columns(6)
    c1.metric("Hotels",   df["hotel_name"].nunique())
    c2.metric("Win Rate", f"{wr:.0f}%", f"{wn}W/{ln}L")
    c3.metric("Losses",   ln, f"Avg {alg:.1f}%")
    c4.metric("No Comp",  ncn)
    c5.metric("⚠ Floor",  fln)
    c6.metric("Avg Margin",f"{df['current_margin'].mean():.1f}%")

    col1, col2 = st.columns(2)
    with col1:
        wl = df.groupby(["dep_window","result"]).size().reset_index(name="n")
        fig = px.bar(wl, x="dep_window", y="n", color="result",
            color_discrete_map={"Win":"#0A7C4E","Lose":"#C0392B","Win - No Comp":"#D4AF37"},
            title="Results by Departure Window",
            category_orders={"dep_window":["0-60","61-120","121-240","241+"]})
        fig.update_layout(height=300,margin=dict(t=40,b=10))
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        if not comp_df.empty:
            ds = comp_df.groupby("destination").apply(
                lambda g: (g["result"]=="Win").sum()/len(g)*100).reset_index()
            ds.columns = ["Dest","WinRate"]
            fig2 = px.bar(ds, x="Dest", y="WinRate", title="Win Rate by Destination (%)",
                color="WinRate",color_continuous_scale=["#C0392B","#D4AF37","#0A7C4E"],
                range_color=[30,80])
            fig2.update_layout(height=300,margin=dict(t=40,b=10))
            st.plotly_chart(fig2, use_container_width=True)

    st.markdown("#### 🏆 Top Sellers")
    ts = df.groupby("hotel_name").agg(
        Dest=("destination","first"),
        Rows=("hotel_name","count"),
        AvgMargin=("current_margin","mean"),
        HasLoss=("result",lambda x:"⚠ Loss" if (x=="Lose").any() else "✓ OK")
    ).sort_values("Rows",ascending=False).head(15)
    ts["WinRate"] = df.groupby("hotel_name").apply(
        lambda g: f"{(g['result']=='Win').sum()/max(len(g[g['result'].isin(['Win','Lose'])]),1)*100:.0f}%"
    ).values[:15]
    ts["AvgMargin"] = ts["AvgMargin"].map("{:.1f}%".format)
    st.dataframe(ts[["Dest","Rows","WinRate","AvgMargin","HasLoss"]], use_container_width=True)

# ── 3. HOTEL ACTIONS ───────────────────────────────────────────────────────────
with tabs[2]:
    st.subheader("Individual Hotel Actions")
    c1,c2,c3 = st.columns([3,2,2])
    srch = c1.text_input("Search hotel",placeholder="e.g. Fushifaru…",key="srch3")
    rf   = c2.selectbox("Result",["All","Win","Lose","Win - No Comp"],key="rf3")
    wf   = c3.multiselect("Window",["0-60","61-120","121-240","241+"],
                           default=["0-60","61-120","121-240","241+"],key="wf3")
    v = df.copy()
    if srch: v = v[v["hotel_name"].str.contains(srch,case=False,na=False)]
    if rf!="All": v = v[v["result"]==rf]
    if wf: v = v[v["dep_window"].isin(wf)]
    v["_tw"] = v["booking_tier"].map({"high":3,"medium":2,"low":1}).fillna(1)
    v = v.sort_values(["_tw","priority_score"],ascending=False)
    def make_action(r):
        if r["result"]=="Lose": return "↓ Reduce" if abs(r["diff_pct"])>5 else "Monitor"
        if r["result"]=="Win" and r["diff_pct"]>RAISE_THRESHOLD: return "↑ Raise"
        if r["result"]=="Win": return "Maintain"
        return "Dest-guided"
    v["Action"]= v.apply(make_action,axis=1)
    v["Margin"]= v.apply(lambda r: margin_cell(r["current_margin"],r["margin_after"]),axis=1)
    show3 = v[["giata","hotel_name","destination","dep_window","dep_month","board",
               "pp_price","comp_price","diff_pct","result","Margin","Action","booking_tier"]].copy()
    show3.columns=["Giata","Hotel","Dest","Window","Month","Board","D2 £pp","Comp £pp","Diff %","Result","Margin","Action","Tier"]
    st.dataframe(show3.style.format({"D2 £pp":"{:.0f}","Comp £pp":"{:.0f}","Diff %":"{:.1f}%"}),
                 use_container_width=True,height=600)

# ── 4. REDUCE ──────────────────────────────────────────────────────────────────
with tabs[3]:
    st.subheader("↓ Price Reduction Candidates")
    cands = loses_df.copy()
    fln2  = (cands["margin_flag"]=="floor").sum()
    nrn   = (cands["margin_flag"]=="near").sum()
    st.error(f"⚠ {fln2} of {len(cands)} reductions breach 7% floor. {nrn} enter 7–8.5% caution band.")
    c1,c2 = st.columns([3,2])
    s4=c1.text_input("Search",key="s4"); w4=c2.multiselect("Window",["0-60","61-120","121-240","241+"],
        default=["0-60","61-120","121-240","241+"],key="w4")
    if s4: cands=cands[cands["hotel_name"].str.contains(s4,case=False,na=False)]
    if w4: cands=cands[cands["dep_window"].isin(w4)]
    cands=cands.sort_values(["booking_tier","diff_pct"],ascending=[False,True])
    cands["Margin"] = cands.apply(lambda r: margin_cell(r["current_margin"],r["margin_after"]),axis=1)
    cands["Flag"]   = cands["margin_flag"].map({"floor":"🔴 SUPPRESS","near":"🟡 Near Floor",
                                                 "ok":"✅ OK","ceiling":"↑ Ceiling","unknown":"?"})
    cands["Action"] = cands.apply(lambda r:
        "SUPPRESS — floor breach" if r["margin_flag"]=="floor"
        else f"↓ Reduce ~£{abs(r['comp_price']-r['pp_price']):.0f}pp",axis=1)
    show4=cands[["giata","hotel_name","destination","dep_window","dep_month",
                  "pp_price","comp_price","diff_pct","Margin","Flag","Action","booking_tier"]].copy()
    show4.columns=["Giata","Hotel","Dest","Window","Month","D2 £pp","Comp £pp","Diff %","Margin","Flag","Action","Tier"]
    st.dataframe(show4.style.format({"D2 £pp":"{:.0f}","Comp £pp":"{:.0f}","Diff %":"{:.1f}%"}),
                 use_container_width=True,height=550)

# ── 5. RAISE ───────────────────────────────────────────────────────────────────
with tabs[4]:
    st.subheader("↑ Raise / Margin Opportunities")
    rdf = wins_df[wins_df["diff_pct"]>RAISE_THRESHOLD].copy()
    st.success(f"✅ {len(rdf)} hotels winning by >15%")
    rdf=rdf.sort_values(["booking_tier","diff_pct"],ascending=[False,False])
    rdf["Margin"]=rdf.apply(lambda r: margin_cell(r["current_margin"],r["margin_after"]),axis=1)
    rdf["Ceiling"]=rdf["margin_flag"].map({"ceiling":"🔵 ABOVE","ok":"✅ OK","near":"🟡 Near","floor":"🔴 Floor","unknown":"?"})
    show5=rdf[["hotel_name","destination","dep_window","dep_month",
               "pp_price","comp_price","diff_pct","Margin","Ceiling","booking_tier"]].copy()
    show5.columns=["Hotel","Dest","Window","Month","D2 £pp","Comp £pp","Win %","Margin","Ceiling","Tier"]
    st.dataframe(show5.style.format({"D2 £pp":"{:.0f}","Comp £pp":"{:.0f}","Win %":"{:.1f}%"}),
                 use_container_width=True,height=500)

# ── 6. MAINTAIN ────────────────────────────────────────────────────────────────
with tabs[5]:
    mtn=wins_df[(wins_df["diff_pct"]>=-5)&(wins_df["diff_pct"]<=HOLD_UPPER)].copy()
    mtn["MFlag"]=mtn["margin_flag"].map({"floor":"🔴 FLOOR","near":"🟡 Near","ok":"✅ OK",
                                          "ceiling":"↑ Ceiling","unknown":"?"})
    show6=mtn[["hotel_name","destination","dep_window","dep_month",
               "pp_price","comp_price","diff_pct","current_margin","MFlag","booking_tier"]].copy()
    show6.columns=["Hotel","Dest","Window","Month","D2 £pp","Comp £pp","Diff %","Margin %","Flag","Tier"]
    st.dataframe(show6.style.format({"D2 £pp":"{:.0f}","Comp £pp":"{:.0f}","Diff %":"{:.1f}%","Margin %":"{:.1f}%"}),
                 use_container_width=True,height=500)

# ── 7. PRODUCT GAPS ────────────────────────────────────────────────────────────
with tabs[6]:
    st.warning("⚠ Product Gap rows — competitor offers a different tier. NEVER price cut — product fix required.")
    pg=loses_df[loses_df["op_room"].str.lower().str.contains(
        "suite|villa|overwater|pool villa|penthouse",na=False,regex=True)].copy()
    if pg.empty: pg=loses_df.copy()
    show7=pg[["hotel_name","destination","dep_window","op_room",
              "pp_price","comp_price","diff_pct","current_margin","booking_tier"]].copy()
    show7.columns=["Hotel","Dest","Window","D2 Room","D2 £pp","Comp £pp","Gap %","Margin","Tier"]
    show7["Action"]="Product Fix — not price cut"
    st.dataframe(show7,use_container_width=True,height=500)

# ── 8. NO COMP ─────────────────────────────────────────────────────────────────
with tabs[7]:
    st.info("📌 No LH comparable — destination-guided pricing. Never 'price freely'.")
    stances={}
    for (d,w),grp in comp_df.groupby(["destination","dep_window"]):
        wc=(grp["result"]=="Win").sum(); lc=(grp["result"]=="Lose").sum()
        wr2=wc/(wc+lc) if (wc+lc)>0 else 0
        al2=grp[grp["result"]=="Lose"]["diff_pct"].mean() if lc>0 else 0
        aw2=grp[grp["result"]=="Win"]["diff_pct"].mean() if wc>0 else 0
        stances[(d,w)]=derive_stance(wr2,al2,aw2)
    nc2=nc_df.copy()
    nc2["Stance"]=nc2.apply(lambda r:stances.get((r["destination"],r["dep_window"]),"Hold — Optimise Margin"),axis=1)
    show8=nc2[["hotel_name","destination","dep_window","dep_month","pp_price","current_margin","Stance","booking_tier"]].copy()
    show8.columns=["Hotel","Dest","Window","Month","D2 £pp","Margin","Dest Stance","Tier"]
    st.dataframe(show8,use_container_width=True,height=500)

# ── 9. DEST ACTIONS ────────────────────────────────────────────────────────────
with tabs[8]:
    rows9=[]
    for (d,w),grp in comp_df.groupby(["destination","dep_window"]):
        wc=(grp["result"]=="Win").sum(); lc=(grp["result"]=="Lose").sum()
        ncw=nc_df[(nc_df["destination"]==d)&(nc_df["dep_window"]==w)].shape[0]
        wr3=wc/(wc+lc) if (wc+lc)>0 else 0
        al3=grp[grp["result"]=="Lose"]["diff_pct"].mean() if lc>0 else 0
        aw3=grp[grp["result"]=="Win"]["diff_pct"].mean() if wc>0 else 0
        rows9.append({"Dest":d,"Window":w,"Win Rate":f"{wr3*100:.0f}%",
                       "Avg Lose":f"{al3:.1f}%","Avg Win":f"{aw3:.1f}%",
                       "Stance":derive_stance(wr3,al3,aw3),
                       "Wins":wc,"Losses":lc,"No Comp":ncw})
    if rows9: st.dataframe(pd.DataFrame(rows9),use_container_width=True,height=400)

# ── 10. OUTLIERS ───────────────────────────────────────────────────────────────
with tabs[9]:
    cl2,bw2=[],[]
    for (h,d,w),grp in comp_df.groupby(["hotel_name","destination","dep_window"]):
        wc=(grp["result"]=="Win").sum(); lc=(grp["result"]=="Lose").sum()
        if lc>=2 and wc==0:
            cl2.append({"Hotel":h,"Dest":d,"Window":w,"Losses":lc,
                         "Avg Gap":f"{grp[grp['result']=='Lose']['diff_pct'].mean():.1f}%",
                         "Floor Breach":"⚠ YES" if grp["margin_flag"].eq("floor").any() else "No"})
        if wc>=2:
            aw4=grp[grp["result"]=="Win"]["diff_pct"].mean()
            if aw4>30:
                bw2.append({"Hotel":h,"Dest":d,"Window":w,"Wins":wc,"Avg Win":f"+{aw4:.1f}%"})
    if cl2:
        st.error(f"🔴 {len(cl2)} consistent loser situations")
        st.dataframe(pd.DataFrame(cl2),use_container_width=True)
    if bw2:
        st.success(f"🔵 {len(bw2)} big win opportunities")
        st.dataframe(pd.DataFrame(bw2),use_container_width=True)
    if not cl2 and not bw2: st.info("No severe outliers in current filter.")

# ── 11. MISSING ────────────────────────────────────────────────────────────────
with tabs[10]:
    if len(sel_f)>1:
        hf=df.groupby("hotel_name")["file_name"].nunique()
        miss=hf[hf<len(sel_f)].reset_index()
        miss.columns=["Hotel","Files Present"]
        miss["Files Missing"]=len(sel_f)-miss["Files Present"]
        st.dataframe(miss.sort_values("Files Missing",ascending=False),use_container_width=True)
    else:
        st.info("Load 2+ files to detect missing properties across runs.")

# ── 12. ADVANTAGE ──────────────────────────────────────────────────────────────
with tabs[11]:
    adv=wins_df[wins_df["diff_pct"]>HOLD_UPPER].groupby(["hotel_name","destination"]).agg(
        Count=("diff_pct","count"),AvgWin=("diff_pct","mean"),AvgMargin=("current_margin","mean")
    ).reset_index().sort_values("AvgWin",ascending=False)
    adv["AvgWin"]=adv["AvgWin"].map("{:+.1f}%".format)
    adv["AvgMargin"]=adv["AvgMargin"].map("{:.1f}%".format)
    adv.columns=["Hotel","Destination","Records","Avg Win Gap","Avg Margin"]
    adv["Recommendation"]="Consider margin increase"
    st.dataframe(adv,use_container_width=True,height=500)

# ── 13. INSIGHTS ───────────────────────────────────────────────────────────────
with tabs[12]:
    fln3=(loses_df["margin_flag"]=="floor").sum() if not loses_df.empty else 0
    st.error("### 🔴 Critical Losses")
    if not loses_df.empty:
        tl=loses_df.nsmallest(8,"diff_pct")[["hotel_name","destination","dep_window",
            "pp_price","comp_price","diff_pct","margin_after","margin_flag"]].copy()
        tl.columns=["Hotel","Dest","Window","D2 £pp","Comp £pp","Gap %","Margin After","Flag"]
        st.dataframe(tl.style.format({"D2 £pp":"{:.0f}","Comp £pp":"{:.0f}","Gap %":"{:.1f}%","Margin After":"{:.1f}%"}),
                     use_container_width=True)
    st.warning(f"### ⚠ Margin Guardrails\n- **{fln3}** hotels breach 7% floor if matched\n- Avg D2 margin: **{df['current_margin'].mean():.1f}%**")
    if not wins_df.empty:
        st.success("### ✅ Margin Opportunities")
        bw3=wins_df[wins_df["diff_pct"]>20].nlargest(5,"diff_pct")[
            ["hotel_name","destination","dep_window","diff_pct","current_margin"]].copy()
        bw3.columns=["Hotel","Dest","Window","Win %","Margin"]
        st.dataframe(bw3.style.format({"Win %":"{:.1f}%","Margin":"{:.1f}%"}),use_container_width=True)
    st.info("### 🗓 Window Split")
    for w in ["0-60","61-120","121-240","241+"]:
        wg=comp_df[comp_df["dep_window"]==w]
        if wg.empty: continue
        wr5=(wg["result"]=="Win").sum()/len(wg)*100
        st.write(f"**{w}d** — Win rate: {wr5:.0f}% ({(wg['result']=='Win').sum()}W / {(wg['result']=='Lose').sum()}L)")

# ── 14. PRICE TRENDS ───────────────────────────────────────────────────────────
with tabs[13]:
    st.subheader("↗ Price Trends")

    # ── Filters row ──────────────────────────────────────────────────────────
    fc1, fc2, fc3, fc4 = st.columns([2,2,2,2])

    all_trend_dests = sorted(df["destination"].unique().tolist())
    sel_trend_dest  = fc1.multiselect("Destination", all_trend_dests,
                                       default=all_trend_dests, key="td_dest")

    all_trend_comps = sorted(df["competitor"].unique().tolist())
    sel_trend_comp  = fc2.multiselect("Competitor", all_trend_comps,
                                       default=all_trend_comps, key="td_comp")

    trend_df = df.copy()
    if sel_trend_dest: trend_df = trend_df[trend_df["destination"].isin(sel_trend_dest)]
    if sel_trend_comp: trend_df = trend_df[trend_df["competitor"].isin(sel_trend_comp)]

    hotel_list = ["All Hotels"] + sorted(trend_df["hotel_name"].unique().tolist())
    worst_hotel = (loses_df.nsmallest(1,"diff_pct")["hotel_name"].values[0]
                   if not loses_df.empty else hotel_list[1] if len(hotel_list)>1 else "All Hotels")
    default_idx = hotel_list.index(worst_hotel) if worst_hotel in hotel_list else 0

    sel_h = fc3.selectbox("Hotel (type to search)", hotel_list,
                           index=default_idx, key="td_hotel")

    board_opts = ["All"]
    if sel_h != "All Hotels":
        board_opts += sorted(trend_df[trend_df["hotel_name"]==sel_h]["board"].unique().tolist())
    else:
        board_opts += sorted(trend_df["board"].unique().tolist())
    sel_b = fc4.selectbox("Board", board_opts, key="td_board")

    # Apply hotel/board filter
    h = trend_df.copy()
    if sel_h != "All Hotels": h = h[h["hotel_name"]==sel_h]
    if sel_b != "All":        h = h[h["board"]==sel_b]

    st.markdown("---")

    # ── Seasonal curve ────────────────────────────────────────────────────────
    if not h.empty:
        h = h.copy()
        h["dep_date_dt"] = pd.to_datetime(h["dep_date"], errors="coerce")
        sc = h.dropna(subset=["dep_date_dt"])
        sc = sc[sc["pp_price"] > 0].sort_values("dep_date_dt")

        if sel_h == "All Hotels":
            # Aggregate all hotels by departure date
            sc_grp = sc.groupby("dep_date_dt").agg(
                D2_pp  =("pp_price","mean"),
                Comp_pp=("comp_price", lambda x: x[x>0].mean() if (x>0).any() else np.nan)
            ).reset_index()
            chart_title = f"Avg D2 vs LH Price by Departure Date — {', '.join(sel_trend_dest) if sel_trend_dest else 'All'}"
        else:
            sc_grp = sc.groupby("dep_date_dt").agg(
                D2_pp  =("pp_price","mean"),
                Comp_pp=("comp_price", lambda x: x[x>0].mean() if (x>0).any() else np.nan),
                Result =("result","first")
            ).reset_index()
            chart_title = f"{sel_h} — D2 vs LoveHolidays Price by Departure Date"

        fig_sc = go.Figure()
        fig_sc.add_trace(go.Scatter(
            x=sc_grp["dep_date_dt"], y=sc_grp["D2_pp"],
            name="D2 Price £", mode="lines+markers",
            line=dict(color="#1B6FD4", width=3), marker=dict(size=6)
        ))
        comp_valid = sc_grp[sc_grp["Comp_pp"].notna() & (sc_grp["Comp_pp"]>0)]
        if not comp_valid.empty:
            fig_sc.add_trace(go.Scatter(
                x=comp_valid["dep_date_dt"], y=comp_valid["Comp_pp"],
                name="LH Price £", mode="lines+markers",
                line=dict(color="#E04A3F", width=2, dash="dot"), marker=dict(size=6)
            ))
        fig_sc.update_layout(
            title=chart_title,
            xaxis_title="Departure Date", yaxis_title="Price per person £",
            height=380, hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02)
        )
        st.plotly_chart(fig_sc, use_container_width=True)

    st.markdown("---")
    col3, col4 = st.columns(2)

    # ── Win Rate by departure month ───────────────────────────────────────────
    with col3:
        wr_src = trend_df[trend_df["result"].isin(["Win","Lose"])]
        if sel_h != "All Hotels":
            wr_src = wr_src[wr_src["hotel_name"]==sel_h]

        if not wr_src.empty:
            wr_m = wr_src.groupby("dep_month").apply(
                lambda g: (g["result"]=="Win").sum()/len(g)*100
            ).reset_index()
            wr_m.columns = ["Month","WinRate"]
            wr_m = wr_m.sort_values("Month")
            wr_m["Label"] = wr_m["Month"].apply(
                lambda m: pd.to_datetime(m+"-01").strftime("%b %Y") if m and len(m)==7 else "")
            wr_m["Colour"] = wr_m["WinRate"].apply(
                lambda v: "#C0392B" if v < 50 else "#27AE60")

            fig_wr = go.Figure()
            fig_wr.add_trace(go.Bar(
                x=wr_m["Label"], y=wr_m["WinRate"],
                marker_color=wr_m["Colour"], name="Win Rate %"
            ))
            fig_wr.add_hline(y=50, line_dash="dot", line_color="#E04A3F",
                              annotation_text="50%", annotation_position="right")
            lbl = sel_h if sel_h != "All Hotels" else f"{', '.join(sel_trend_dest) if sel_trend_dest else 'All'}"
            fig_wr.update_layout(
                title=f"Win Rate by Departure Month — {lbl}",
                yaxis=dict(range=[0,100], title="Win Rate %"),
                xaxis_title="Departure Month",
                height=320, showlegend=False
            )
            st.plotly_chart(fig_wr, use_container_width=True)

    # ── Overall Win Rate Trend across run dates ───────────────────────────────
    with col4:
        wr_run_src = trend_df[trend_df["result"].isin(["Win","Lose"])]
        if sel_h != "All Hotels":
            wr_run_src = wr_run_src[wr_run_src["hotel_name"]==sel_h]

        if wr_run_src["file_date"].nunique() > 1:
            wr_run = wr_run_src.groupby("file_date").apply(
                lambda g: (g["result"]=="Win").sum()/len(g)*100
            ).reset_index()
            wr_run.columns = ["RunDate","WinRate"]
            wr_run = wr_run.sort_values("RunDate")
            wr_run["Label"] = pd.to_datetime(wr_run["RunDate"]).dt.strftime("%d %b")

            fig_run = go.Figure()
            fig_run.add_trace(go.Scatter(
                x=wr_run["Label"], y=wr_run["WinRate"],
                mode="lines+markers",
                line=dict(color="#27AE60", width=2.5),
                marker=dict(size=8, color="#27AE60"),
                fill="tozeroy", fillcolor="rgba(39,174,96,0.1)"
            ))
            fig_run.update_layout(
                title=f"Win Rate Trend ({wr_run['Label'].iloc[0]}–{wr_run['Label'].iloc[-1]})",
                yaxis=dict(title="Win Rate %"),
                xaxis_title="Run Date",
                height=320, showlegend=False
            )
            st.plotly_chart(fig_run, use_container_width=True)
        else:
            if not wr_run_src.empty:
                wr_now = (wr_run_src["result"]=="Win").sum()/len(wr_run_src)*100
                st.metric("Current Win Rate", f"{wr_now:.1f}%",
                           "Upload more run files to track trend over time")

    # ── Price drift across run dates ──────────────────────────────────────────
    if trend_df["file_date"].nunique() > 1:
        st.markdown("#### Price Drift Across Run Dates")
        drift_src = h if sel_h != "All Hotels" else trend_df
        drift = drift_src.groupby("file_date").agg(
            D2  =("pp_price","mean"),
            Comp=("comp_price", lambda x: x[x>0].mean() if (x>0).any() else np.nan)
        ).reset_index().sort_values("file_date")
        drift["Label"] = pd.to_datetime(drift["file_date"]).dt.strftime("%d %b")
        fig_dr = go.Figure()
        fig_dr.add_trace(go.Scatter(x=drift["Label"],y=drift["D2"],
            name="D2 avg £pp",line=dict(color="#1B6FD4",width=2.5),mode="lines+markers"))
        if drift["Comp"].notna().any():
            fig_dr.add_trace(go.Scatter(x=drift["Label"],y=drift["Comp"],
                name="LH avg £pp",line=dict(color="#E04A3F",width=2,dash="dash"),
                mode="lines+markers"))
        drift_title = sel_h if sel_h != "All Hotels" else f"All Hotels — {', '.join(sel_trend_dest) if sel_trend_dest else 'All'}"
        fig_dr.update_layout(title=f"Price Drift — {drift_title}",
                              height=260,xaxis_title="Run Date",yaxis_title="£ pp")
        st.plotly_chart(fig_dr, use_container_width=True)

# ── 15. FILES ──────────────────────────────────────────────────────────────────
with tabs[14]:
    st.subheader("📂 Uploaded Files")
    fl = get_file_list()
    if fl:
        fdf=pd.DataFrame(fl,columns=["File","Run Date","Destination","Competitor","Rows"])
        st.dataframe(fdf, use_container_width=True)
        st.caption(f"{len(fdf)} files · {fdf['Rows'].sum():,} total rows")
        csv = df.to_csv(index=False)
        st.download_button("⬇ Download filtered data as CSV",
                           data=csv,file_name="d2_pricing_export.csv",mime="text/csv")
    else:
        st.info("No files uploaded yet.")
