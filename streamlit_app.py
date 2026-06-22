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
    # Markup rule-engine export (flattened, one row per rule)
    conn.execute("""CREATE TABLE IF NOT EXISTS rules(
        rid TEXT, enabled INTEGER, markup_type TEXT, applies_to TEXT,
        layer TEXT, verb TEXT, action_value REAL, action_unit TEXT,
        ispkg TEXT, user_csv TEXT, giatas TEXT, n_conds INTEGER,
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
    computed across its LATEST comparison rows (all departure dates / boards).

    Idea: markup is the lever. Cutting markup lowers our price → more wins but
    less margin; raising markup does the reverse. We sweep candidate markup
    deltas and pick the one that maximises EXPECTED BOOKED MARGIN, defined as

        projected_win_rate  ×  projected_avg_markup

    i.e. how often we win × how much margin we keep. This naturally balances the
    two asks: over-cutting kills margin (term 2 falls), and raising too far
    converts wins to losses (term 1 falls), so the optimum sits in between.

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
    # Latest run only — "latest comparisons across all dates for the hotel"
    if "file_date" in h.columns and h["file_date"].nunique() > 1:
        h = h[h["file_date"] == h["file_date"].max()]

    comp = h[h["result"].isin(["Win","Lose"])].copy()
    base_m = pd.to_numeric(h.get("current_margin"), errors="coerce").mean()
    base_m = float(base_m) if pd.notna(base_m) else 0.0

    if comp.empty:
        return {"delta":0.0,"wr0":np.nan,"wr1":np.nan,"m0":base_m,"m1":base_m,
                "note":"No direct comp — raise selectively (dest-guided)"}

    m_rows = pd.to_numeric(comp["current_margin"], errors="coerce").fillna(0.0).values
    gaps   = pd.to_numeric(comp["diff_pct"],       errors="coerce").fillna(0.0).values
    m0  = float(np.mean(m_rows))
    wr0 = float((gaps > 0).mean()*100)

    # Safe delta band: keep every row inside [floor, ceiling]
    d_min = -max(0.0, float(np.min(m_rows)) - MARGIN_FLOOR)
    d_max =  max(0.0, MARGIN_CEILING - float(np.max(m_rows)))
    d_min, d_max = max(d_min, -band), min(d_max, band)
    if d_max - d_min < step:
        return {"delta":0.0,"wr0":wr0,"wr1":wr0,"m0":m0,"m1":m0,
                "note":"At margin guardrail — hold"}

    factors = 100.0/(100.0+np.clip(m_rows, 0, None))   # %price move per pp markup
    cands = []
    d = d_min
    while d <= d_max + 1e-9:
        m1 = m0 + d
        if m1 >= MARGIN_FLOOR - 1e-9:
            wr1 = float(((gaps - d*factors) > 0).mean()*100)   # win if still cheaper
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
        giatas=",".join(giatas), n_conds=len(props),
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

    Levers = consumer-layer markups (Manual/Exclusive) specific to this hotel.
    Cleanest lever = fewest extra conditions (so it reliably applies), matching
    package vs hotel-only, package-level preferred for package searches.
    Exclusive (overrides Manual) wins if present. If no hotel-specific consumer
    rule exists, recommend CREATING one. System-cost rules are shown as context
    only, never proposed as a margin lever."""
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
    mk = mk[mk.apply(compatible, axis=1)]
    consumer = mk[mk["layer"].str.startswith("Consumer")].copy()
    system   = mk[mk["layer"].str.startswith("System")].copy()
    if consumer.empty:
        return {"status":"create", "system_context":system}
    # Prefer Exclusive, then fewest conditions, then package-level for packages
    consumer["excl"] = (consumer["layer"]=="Consumer (Exclusive)").astype(int)
    consumer["lvl_pref"] = consumer["applies_to"].map(
        lambda a: 1 if (is_package and a=="15") or ((not is_package) and a=="12") else 0)
    consumer = consumer.sort_values(["excl","lvl_pref","n_conds"],
                                     ascending=[False,False,True])
    best = consumer.iloc[0]
    # Proposed new value (approximate; % levers only)
    new_val = None
    if best["action_unit"]=="%" and best["action_value"] is not None:
        new_val = round(best["action_value"] + delta, 2)
    return {"status":"edit","rule":best,"alternatives":consumer.iloc[1:],
            "system_context":system,"new_value":new_val}

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
                conn.execute("DELETE FROM rules")
                pr.to_sql("rules", conn, if_exists="append", index=False)
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
                "Insights","↗ Price Trends","📂 Files","🎚 Markup Levers"])

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
    v = tab_filters(df, "pg")
    pg = v[v["result"]=="Lose"].copy()
    pg = pg[pg["op_room"].str.lower().str.contains(
        "suite|villa|overwater|pool villa|penthouse",na=False,regex=True)]
    if pg.empty: pg = v[v["result"]=="Lose"].copy()
    show7 = pg[["giata","hotel_name","destination","competitor","dep_window",
                 "op_room","pp_price","comp_price","diff_pct","current_margin","booking_tier"]].copy()
    show7.columns = ["Giata","Hotel","Dest","Comp","Window","D2 Room",
                     "D2 £pp","Comp £pp","Gap %","Margin","Tier"]
    show7["Action"] = "Product Fix — not price cut"
    st.dataframe(with_bkg(show7),use_container_width=True,height=500)

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
    nc2["Stance"] = nc2.apply(lambda r:stances.get((r["destination"],r["dep_window"]),"Hold — Optimise Margin"),axis=1)
    show8 = nc2[["giata","hotel_name","destination","competitor","dep_window","dep_month",
                  "pp_price","current_margin","Stance","booking_tier"]].copy()
    show8.columns = ["Giata","Hotel","Dest","Comp","Window","Month",
                     "D2 £pp","Margin","Dest Stance","Tier"]
    st.dataframe(with_bkg(show8),use_container_width=True,height=500)

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
            fig_wr.add_trace(go.Bar(x=wr_m["Label"],y=wr_m["WinRate"],marker_color=wr_m["Colour"]))
            fig_wr.add_hline(y=50,line_dash="dot",line_color="#E04A3F",
                              annotation_text="50%",annotation_position="right")
            dest_lbl2 = ", ".join(v["destination"].unique().tolist())
            fig_wr.update_layout(title=f"Win Rate by Departure Month — {dest_lbl2}",
                yaxis=dict(range=[0,100],title="Win Rate %"),
                xaxis_title="Departure Month",height=320,showlegend=False)
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
                   "**consumer-layer rule specific to that hotel** to edit — preferring the cleanest "
                   "lever (fewest conditions, package-level for packages, Exclusive over Manual). "
                   "System-cost rules are shown as context only. Where no hotel-specific consumer "
                   "rule exists, it recommends **creating** one. The new value is an approximate "
                   "starting point — a human approves before it goes live.")

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
                action = f"EDIT rule {r['rid']}"
                lever_text = r["generated_text"]
            elif rec["status"]=="create":
                cur, newv = "—", f"{delta:+.1f}% (new)"
                action = "CREATE giata rule"
                lvl = "HotelFlightPackage" if is_pkg else "HotelRoom"
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
                       f"{(out['Action'].str.startswith('EDIT')).sum()} edit an existing rule · "
                       f"{(out['Action'].str.startswith('CREATE')).sum()} need a new rule.")

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
                    st.warning("No hotel-specific consumer rule exists — the markup here comes from "
                               "blanket/provider rules, which can't be edited for one hotel. Create a "
                               "giata-specific rule:")
                    lvl = "HotelFlightPackage" if is_pkg else "HotelRoom"
                    st.code(f"If (GIata Equal {gsel}) And (User Equal Destination2) "
                            f"Then MarkupNew {lvl} By {s['delta']:+.1f}% Total", language=None)
                sc = rec.get("system_context")
                if sc is not None and not sc.empty:
                    st.markdown("**System-cost rules on this hotel (context only — not margin levers):**")
                    st.dataframe(sc[["rid","action_value","action_unit","generated_text"]]
                                 .rename(columns={"rid":"Id","action_value":"Val","action_unit":"Unit",
                                                  "generated_text":"GeneratedText"}),
                                 use_container_width=True)
