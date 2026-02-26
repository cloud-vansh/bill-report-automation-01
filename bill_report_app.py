import streamlit as st
from openpyxl import load_workbook
from collections import defaultdict
from datetime import datetime
import pandas as pd

st.set_page_config(
    page_title="BillFlow · Report Automation",
    page_icon="✦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ── session state ──────────────────────────────────────────────────────────────
for k, v in {
    'selected_month': max(0, datetime.now().month - 2),
    'processing_history': [],
    'show_preview': False,
    'show_settings': False,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── helpers ────────────────────────────────────────────────────────────────────
MONTHS = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"]
BILL_HEADER_ROW, BILL_DATA_START = 3, 4
SRC_HEADER_ROW,  SRC_DATA_START  = 7, 8

def validate_file(file, file_type):
    try:
        wb = load_workbook(file)
        ws = wb.active
        hr = 7 if file_type in ["job","einv"] else 3
        return {
            "valid": True,
            "rows": ws.max_row,
            "cols": ws.max_column,
            "sheets": wb.sheetnames if file_type == "bill" else None,
        }
    except Exception as e:
        return {"valid": False, "error": str(e)}

def norm(x): return str(x).strip().upper() if x else ""
def clean_job(j): return norm(j).split("/")[0]

def parse_date(d):
    if not d: return None
    if isinstance(d, datetime): return d.date()
    s = str(d).strip()
    if not s or s.upper() == "NONE": return None
    for fmt in ["%Y-%m-%d","%d-%m-%Y","%d/%m/%Y","%m/%d/%Y","%d.%m.%Y",
                "%Y/%m/%d","%d-%b-%Y","%d %b %Y","%d-%m-%y","%d/%m/%y"]:
        try: return datetime.strptime(s, fmt).date()
        except: pass
    return None

def fmt_date(d):
    if not d: return None
    if isinstance(d, datetime): return d.strftime("%d/%m/%y")
    try: return d.strftime("%d/%m/%y")
    except: return None

def col_map(ws, hr):
    m = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(hr, c).value
        if v: m[norm(v)] = c
    return m

def get_existing_jobs(ws, bc, start):
    ex = set()
    for r in range(start, ws.max_row + 1):
        j = ws.cell(r, bc["JOB NO"]).value
        if j: ex.add(clean_job(j))
    return ex

def next_empty(ws, start):
    for r in range(start, ws.max_row + 100):
        if not ws.cell(r,1).value and not ws.cell(r,2).value: return r
    return ws.max_row + 1

# ══════════════════════════════════════════════════════════════════════════════
# DESIGN SYSTEM
# ══════════════════════════════════════════════════════════════════════════════
CSS = """
@import url('https://fonts.googleapis.com/css2?family=Abril+Fatface&family=Dancing+Script:wght@600;700&family=Jost:wght@300;400;500;600&family=Playfair+Display:ital,wght@0,700;0,900;1,700&display=swap');

:root {
    --cream:      #F5EDD8;
    --cream-lt:   #FAF5E9;
    --cream-dk:   #EBE0C4;
    --espresso:   #2C1503;
    --brown:      #5C3015;
    --tan:        #A07040;
    --sienna:     #B8451E;
    --sienna-lt:  #D4632E;
    --sienna-pale:#F5E8E0;
    --green:      #3D7A55;
    --green-pale: #E6F2EB;
    --orange:     #C45E1A;
    --orange-pale:#F7EDE0;
    --gold:       #A07820;
    --gold-pale:  #F5EED8;
    --teal:       #2A7070;
    --teal-pale:  #E0F0F0;
    --rule:       rgba(44,21,3,.12);
    --rule-md:    rgba(44,21,3,.22);
    --shadow-sm:  0 2px 8px rgba(44,21,3,.08);
    --shadow-md:  0 4px 20px rgba(44,21,3,.12);
    --shadow-lg:  0 10px 40px rgba(44,21,3,.16);
    --ff-display: 'Abril Fatface', Georgia, serif;
    --ff-script:  'Dancing Script', cursive;
    --ff-serif:   'Playfair Display', Georgia, serif;
    --ff-body:    'Jost', sans-serif;
    --r-sm: 8px; --r-md: 14px; --r-lg: 20px; --r-pill: 100px;
}

/* ── reset & globals ── */
*, *::before, *::after { box-sizing: border-box; }
html, body,
[data-testid="stAppViewContainer"],
[data-testid="stAppViewBlockContainer"] {
    font-family: var(--ff-body) !important;
    background: var(--cream) !important;
    color: var(--espresso) !important;
}
.main { background: var(--cream) !important; }
.block-container {
    max-width: 1200px !important;
    padding: 0 2rem 4rem !important;
}
#MainMenu, footer, .stDeployButton,
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"] { display: none !important; }
::-webkit-scrollbar { width: 5px; }
::-webkit-scrollbar-track { background: var(--cream-dk); }
::-webkit-scrollbar-thumb { background: var(--tan); border-radius: 3px; }

/* ── sticky nav ── */
.bf-nav {
    position: sticky; top: 0; z-index: 200;
    background: var(--cream);
    border-bottom: 1px solid var(--rule);
    display: flex; align-items: center; justify-content: space-between;
    padding: 0 2rem;
    height: 62px;
    margin: 0 -2rem 0;
}
.nav-brand {
    font-family: var(--ff-body);
    font-weight: 700;
    font-size: 0.82rem;
    letter-spacing: .22em;
    text-transform: uppercase;
    color: var(--espresso);
    display: flex; align-items: center; gap: .5rem;
}
.nav-brand-dot { color: var(--sienna); font-size: 1rem; }
.nav-links { display: flex; align-items: center; gap: 2rem; }
.nav-link {
    font-family: var(--ff-body);
    font-size: .82rem; font-weight: 500;
    color: var(--brown);
    letter-spacing: .02em;
    transition: color .15s;
}
.nav-link:hover { color: var(--espresso); }

/* ── hero ── */
.bf-hero {
    text-align: center;
    padding: 3.5rem 1rem 3rem;
    background: linear-gradient(180deg, var(--cream-lt) 0%, var(--cream) 100%);
    margin: 0 -2rem 0;
    border-bottom: 1px solid var(--rule);
    position: relative;
    overflow: hidden;
}
.bf-hero::before {
    content: '';
    position: absolute; inset: 0;
    background: radial-gradient(ellipse 60% 50% at 50% 110%,
        rgba(184,69,30,.07) 0%, transparent 70%);
    pointer-events: none;
}
.hero-eyebrow {
    font-family: var(--ff-script);
    font-size: 1.5rem; color: var(--sienna);
    display: block; margin-bottom: .15rem;
    animation: fadeUp .5s ease both;
}
.hero-h1 {
    font-family: var(--ff-display);
    font-size: clamp(2.6rem, 5.5vw, 4.4rem);
    color: var(--espresso);
    line-height: .96; letter-spacing: -.01em;
    margin-bottom: .5rem;
    animation: fadeUp .55s .08s ease both;
}
.hero-h1 em {
    font-family: var(--ff-serif);
    font-style: italic; font-size: .88em;
    color: var(--brown);
}
.hero-sub {
    font-family: var(--ff-body);
    font-size: .9rem; font-weight: 400;
    color: var(--tan);
    letter-spacing: .1em; text-transform: uppercase;
    margin-bottom: 2rem;
    animation: fadeUp .6s .15s ease both;
}
.hero-pill {
    display: inline-flex; align-items: center; gap: .6rem;
    background: var(--espresso); color: var(--cream-lt);
    font-family: var(--ff-body); font-size: .78rem; font-weight: 600;
    letter-spacing: .1em; text-transform: uppercase;
    padding: .75rem 2rem;
    border-radius: var(--r-pill);
    box-shadow: var(--shadow-md);
    animation: fadeUp .65s .2s ease both;
}
.hero-pill-dot {
    width: 20px; height: 20px; border-radius: 50%;
    background: var(--sienna);
    display: flex; align-items: center; justify-content: center;
    font-size: .7rem;
}

/* ── section header ── */
.sec-hdr {
    display: flex; align-items: center; gap: 1rem;
    margin: 2.5rem 0 1.25rem;
}
.sec-hdr-title {
    font-family: var(--ff-serif);
    font-style: italic; font-size: 1.65rem; font-weight: 700;
    color: var(--espresso); white-space: nowrap;
}
.sec-hdr-rule { flex: 1; height: 1px; background: var(--rule-md); }
.sec-hdr-tag {
    font-family: var(--ff-body);
    font-size: .68rem; font-weight: 600;
    letter-spacing: .16em; text-transform: uppercase;
    color: var(--tan);
    background: var(--cream-dk);
    border: 1px solid var(--rule-md);
    padding: .28rem .85rem; border-radius: var(--r-pill);
    white-space: nowrap;
}

/* ── upload card (purely decorative header) ── */
.uc-header {
    background: var(--cream-lt);
    border: 1px solid var(--rule-md);
    border-radius: var(--r-md);
    padding: 1.4rem 1.4rem .9rem;
    margin-bottom: .5rem;
    position: relative; overflow: hidden;
    transition: border-color .2s, box-shadow .2s;
}
.uc-header:hover {
    border-color: var(--rule-md);
    box-shadow: var(--shadow-md);
}
.uc-stripe {
    position: absolute; top: 0; left: 0; right: 0; height: 5px;
}
.uc-badge {
    display: inline-flex; align-items: center; gap: .35rem;
    font-family: var(--ff-body); font-size: .64rem; font-weight: 600;
    letter-spacing: .14em; text-transform: uppercase;
    padding: .25rem .7rem;
    border-radius: var(--r-pill);
    margin-bottom: .85rem;
}
.uc-title {
    font-family: var(--ff-serif); font-size: 1.25rem; font-weight: 700;
    color: var(--espresso); margin-bottom: .3rem; line-height: 1.2;
}
.uc-desc {
    font-family: var(--ff-body); font-size: .82rem;
    color: var(--tan); line-height: 1.5;
}
.uc-loaded {
    display: inline-flex; align-items: center; gap: .35rem;
    font-family: var(--ff-body); font-size: .68rem; font-weight: 600;
    letter-spacing: .1em; text-transform: uppercase;
    padding: .25rem .75rem; border-radius: var(--r-pill);
    margin-top: .1rem;
}

/* colour tokens per card */
.uc-green  .uc-stripe  { background: var(--green); }
.uc-green  .uc-badge   { color: var(--green);  background: var(--green-pale);  border: 1px solid rgba(61,122,85,.2); }
.uc-green  .uc-loaded  { color: var(--green);  background: var(--green-pale);  border: 1px solid rgba(61,122,85,.2); }

.uc-orange .uc-stripe  { background: var(--orange); }
.uc-orange .uc-badge   { color: var(--orange); background: var(--orange-pale); border: 1px solid rgba(196,94,26,.2); }
.uc-orange .uc-loaded  { color: var(--orange); background: var(--orange-pale); border: 1px solid rgba(196,94,26,.2); }

.uc-gold   .uc-stripe  { background: var(--gold); }
.uc-gold   .uc-badge   { color: var(--gold);   background: var(--gold-pale);   border: 1px solid rgba(160,120,32,.2); }
.uc-gold   .uc-loaded  { color: var(--gold);   background: var(--gold-pale);   border: 1px solid rgba(160,120,32,.2); }

/* preview table */
.pv-table {
    background: var(--cream-dk); border-radius: var(--r-sm);
    padding: .6rem .85rem; margin-top: .4rem;
}
.pv-row {
    display: flex; justify-content: space-between;
    font-family: 'Courier Prime', monospace; font-size: .7rem;
    color: var(--brown);
    padding: .18rem 0; border-bottom: 1px solid var(--rule);
}
.pv-row:last-child { border-bottom: none; }
.pv-val { font-weight: 700; color: var(--sienna); }

/* ── file uploader overrides ── */
[data-testid="stFileUploader"] section {
    background: var(--cream-lt) !important;
    border: 1.5px dashed rgba(44,21,3,.2) !important;
    border-radius: var(--r-md) !important;
    padding: .7rem 1rem !important;
    transition: border-color .2s !important;
}
[data-testid="stFileUploader"] section:hover {
    border-color: var(--sienna) !important;
    background: #fff9f0 !important;
}
[data-testid="stFileUploader"] section p,
[data-testid="stFileUploader"] section span,
[data-testid="stFileUploader"] section small {
    font-family: var(--ff-body) !important;
    font-size: .8rem !important;
    color: var(--tan) !important;
}
[data-testid="stFileUploader"] section button {
    background: var(--cream-dk) !important;
    border: 1px solid var(--rule-md) !important;
    color: var(--brown) !important;
    font-family: var(--ff-body) !important;
    font-size: .78rem !important; font-weight: 500 !important;
    border-radius: var(--r-sm) !important;
    transition: background .15s !important;
}
[data-testid="stFileUploader"] section button:hover {
    background: var(--cream) !important;
}

/* ── config panel ── */
.cfg-card {
    background: var(--cream-lt);
    border: 1px solid var(--rule-md);
    border-radius: var(--r-lg);
    padding: 2rem;
    box-shadow: var(--shadow-md);
    position: relative; overflow: hidden;
}
.cfg-card::before {
    content: '';
    position: absolute; top: 0; left: 0; right: 0; height: 4px;
    background: linear-gradient(90deg, var(--green), var(--gold), var(--orange));
}
.cfg-eyebrow {
    font-family: var(--ff-script);
    font-size: 1.2rem; color: var(--sienna);
    display: block; margin-bottom: .05rem;
}
.cfg-title {
    font-family: var(--ff-display);
    font-size: 1.7rem; color: var(--espresso);
    line-height: 1.1; margin-bottom: 1.5rem;
}
.cfg-label {
    font-family: var(--ff-body);
    font-size: .68rem; font-weight: 600;
    letter-spacing: .16em; text-transform: uppercase;
    color: var(--tan); display: block;
    margin-bottom: .45rem;
}
.cfg-rule { height: 1px; background: var(--rule); margin: 1.5rem 0; }

/* ── selectbox ── */
[data-testid="stSelectbox"] > div > div {
    background: var(--cream) !important;
    border: 1px solid var(--rule-md) !important;
    border-radius: var(--r-sm) !important;
    color: var(--espresso) !important;
    font-family: 'Courier Prime', monospace !important;
    font-size: .9rem !important; letter-spacing: .06em !important;
}
[data-testid="stSelectbox"] > div > div:hover { border-color: var(--sienna) !important; }
[data-testid="stSelectbox"] > div > div > div { color: var(--espresso) !important; }

/* ── radio (tab style) ── */
[data-testid="stRadio"] > div {
    display: flex !important; gap: .5rem !important; flex-direction: row !important;
}
[data-testid="stRadio"] label {
    background: var(--cream) !important;
    border: 1px solid var(--rule-md) !important;
    border-radius: var(--r-sm) !important;
    padding: .6rem 1rem !important;
    font-family: var(--ff-body) !important;
    font-size: .85rem !important; color: var(--brown) !important;
    cursor: pointer !important; transition: all .18s !important;
    flex: 1 !important; text-align: center !important;
}
[data-testid="stRadio"] label:hover {
    border-color: var(--sienna) !important;
    color: var(--espresso) !important; background: var(--sienna-pale) !important;
}
[data-testid="stRadio"] label:has(input:checked) {
    background: var(--espresso) !important;
    border-color: var(--espresso) !important;
    color: var(--cream-lt) !important; font-weight: 600 !important;
}

/* ── mode note ── */
.mode-note {
    font-family: var(--ff-body); font-style: italic;
    font-size: .83rem; line-height: 1.5;
    padding: .6rem .9rem; border-radius: var(--r-sm);
    margin-top: .6rem;
}
.mode-ok   { color: var(--green);  background: var(--green-pale);  border: 1px solid rgba(61,122,85,.2); border-left: 3px solid var(--green); }
.mode-warn { color: var(--sienna); background: var(--sienna-pale); border: 1px solid rgba(184,69,30,.2); border-left: 3px solid var(--sienna); }

/* ── execute button ── */
.stButton > button {
    background: var(--cream-dk) !important;
    border: 1px solid var(--rule-md) !important;
    color: var(--brown) !important;
    font-family: var(--ff-body) !important;
    font-size: .8rem !important; font-weight: 600 !important;
    letter-spacing: .06em !important; text-transform: uppercase !important;
    border-radius: var(--r-pill) !important;
    padding: .65rem 1.5rem !important;
    transition: all .2s !important;
}
.stButton > button:hover {
    background: var(--espresso) !important;
    border-color: var(--espresso) !important;
    color: var(--cream-lt) !important;
    transform: translateY(-1px) !important;
    box-shadow: var(--shadow-md) !important;
}
.exec-btn > div > button, .exec-btn .stButton > button {
    background: var(--espresso) !important;
    border-color: var(--espresso) !important;
    color: var(--cream-lt) !important;
    font-size: .82rem !important;
    padding: .85rem 2rem !important;
    box-shadow: var(--shadow-md) !important;
    width: 100% !important;
    letter-spacing: .1em !important;
}
.exec-btn .stButton > button:hover {
    background: var(--sienna) !important;
    border-color: var(--sienna) !important;
    box-shadow: 0 6px 24px rgba(184,69,30,.35) !important;
}

/* ── download button ── */
.stDownloadButton > button {
    background: var(--green) !important;
    border-color: var(--green) !important;
    color: #fff !important;
    font-family: var(--ff-body) !important;
    font-size: .78rem !important; font-weight: 600 !important;
    letter-spacing: .08em !important; text-transform: uppercase !important;
    border-radius: var(--r-pill) !important;
    padding: .7rem 1.5rem !important;
    transition: all .2s !important;
    width: 100% !important;
}
.stDownloadButton > button:hover {
    background: #326646 !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 16px rgba(61,122,85,.3) !important;
}

/* ── checkbox ── */
[data-testid="stCheckbox"] label,
[data-testid="stCheckbox"] span {
    font-family: var(--ff-body) !important;
    font-size: .85rem !important;
    color: var(--brown) !important;
}

/* ── progress ── */
[data-testid="stProgress"] > div {
    background: var(--cream-dk) !important;
    border-radius: 3px !important; height: 4px !important;
}
[data-testid="stProgress"] > div > div {
    background: linear-gradient(90deg, var(--sienna), var(--orange), var(--gold)) !important;
    border-radius: 3px !important;
    box-shadow: 0 0 8px rgba(184,69,30,.35) !important;
}

/* ── native alerts ── */
[data-testid="stAlert"] {
    font-family: var(--ff-body) !important;
    font-size: .85rem !important;
    border-radius: var(--r-md) !important;
}

/* ── history sidebar card ── */
.hist-card {
    background: var(--espresso);
    border-radius: var(--r-lg);
    padding: 1.5rem 1.25rem;
    box-shadow: var(--shadow-lg);
}
.hist-eyebrow {
    font-family: var(--ff-script);
    font-size: 1.1rem; color: rgba(245,237,216,.6);
    display: block; margin-bottom: .05rem;
}
.hist-card-title {
    font-family: var(--ff-display);
    font-size: 1.3rem; color: var(--cream-lt);
    margin-bottom: 1rem;
}
.hist-item {
    display: flex; align-items: center; justify-content: space-between;
    background: rgba(255,255,255,.06);
    border: 1px solid rgba(255,255,255,.09);
    border-radius: var(--r-pill);
    padding: .5rem .9rem;
    margin-bottom: .45rem;
    transition: background .15s;
}
.hist-item:last-child { margin-bottom: 0; }
.hist-item:hover { background: rgba(255,255,255,.09); }
.hi-left { display: flex; flex-direction: column; gap: 1px; }
.hi-time {
    font-family: 'Courier Prime', monospace;
    font-size: .7rem; font-weight: 700; color: var(--cream-lt);
}
.hi-detail {
    font-family: var(--ff-body); font-size: .68rem; color: rgba(245,237,216,.5);
}
.hi-badge {
    font-family: var(--ff-body); font-size: .6rem; font-weight: 700;
    letter-spacing: .12em; text-transform: uppercase;
    padding: .18rem .55rem; border-radius: var(--r-pill);
}
.hi-ok  { color: #8EC49A; background: rgba(61,122,85,.3); }
.hi-err { color: #E8A898; background: rgba(184,69,30,.3); }

/* ── results ── */
.results-card {
    background: var(--cream-lt);
    border: 1px solid var(--rule-md);
    border-radius: var(--r-lg);
    padding: 2rem 2.25rem 1.5rem;
    box-shadow: var(--shadow-md);
    position: relative; overflow: hidden;
    animation: fadeUp .45s ease both;
}
.results-card::before {
    content: '';
    position: absolute; top: 0; left: 0; right: 0; height: 4px;
    background: linear-gradient(90deg, var(--green), var(--gold), var(--orange));
}
.res-eyebrow {
    font-family: var(--ff-script);
    font-size: 1.3rem; color: var(--sienna);
    display: block; margin-bottom: .05rem;
}
.res-title {
    font-family: var(--ff-display);
    font-size: 2rem; color: var(--espresso);
    line-height: 1.05; margin-bottom: 1.5rem;
}
.metrics-row {
    display: grid; grid-template-columns: repeat(3, 1fr); gap: 1rem;
    margin-bottom: 1.5rem;
}
.metric-tile {
    background: var(--cream);
    border: 1px solid var(--rule-md);
    border-radius: var(--r-md);
    padding: 1.2rem 1rem; text-align: center;
    box-shadow: var(--shadow-sm);
}
.mt-num {
    font-family: var(--ff-display);
    font-size: 2.8rem; color: var(--sienna);
    line-height: 1; margin-bottom: .3rem;
}
.mt-lbl {
    font-family: var(--ff-body); font-size: .68rem; font-weight: 600;
    letter-spacing: .12em; text-transform: uppercase; color: var(--tan);
}

/* ── progress status ── */
.prog-txt {
    font-family: 'Courier Prime', monospace;
    font-size: .76rem; color: var(--sienna);
    text-align: center; padding: .35rem;
    letter-spacing: .06em;
}
.prog-detail {
    display: flex; justify-content: space-between;
    font-family: 'Courier Prime', monospace;
    font-size: .66rem; color: var(--tan); margin-top: .2rem;
}

/* ── scroll indicator ── */
.scroll-indicator {
    position: fixed; bottom: 1.5rem; left: 50%;
    transform: translateX(-50%);
    display: flex; flex-direction: column; align-items: center; gap: .3rem;
    animation: fadeIn 1s 1.5s ease both;
    pointer-events: none; z-index: 50;
    opacity: 0;
}
.scroll-indicator.visible { opacity: 1; }
.si-label {
    font-family: var(--ff-body); font-size: .62rem; font-weight: 600;
    letter-spacing: .18em; text-transform: uppercase;
    color: var(--tan);
}
.si-arrow {
    width: 32px; height: 32px;
    background: var(--cream-dk);
    border: 1px solid var(--rule-md);
    border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    color: var(--brown); font-size: .8rem;
    animation: bounce 1.6s ease-in-out infinite;
}
@keyframes bounce {
    0%, 100% { transform: translateY(0); }
    50%       { transform: translateY(5px); }
}
@keyframes fadeIn  { from { opacity:0; } to { opacity:1; } }
@keyframes fadeUp  {
    from { opacity:0; transform:translateY(14px); }
    to   { opacity:1; transform:translateY(0); }
}

/* ── settings panel ── */
.settings-panel {
    background: var(--cream-lt);
    border: 1px solid var(--rule-md);
    border-radius: var(--r-lg);
    padding: 1.5rem 2rem;
    margin-bottom: 1.5rem;
    box-shadow: var(--shadow-sm);
}
.sp-title {
    font-family: var(--ff-serif); font-style: italic;
    font-size: 1.1rem; color: var(--espresso);
    margin-bottom: 1rem;
    display: flex; align-items: center; gap: .75rem;
}
.sp-title::after { content:''; flex:1; height:1px; background:var(--rule); }

/* ── expander ── */
.streamlit-expanderHeader {
    font-family: var(--ff-body) !important;
    font-size: .82rem !important;
    color: var(--brown) !important;
    background: var(--cream-dk) !important;
    border-radius: var(--r-sm) !important;
}

/* ── responsive ── */
@media (max-width: 900px) {
    .block-container { padding: 0 1rem 4rem !important; }
    .bf-nav { margin: 0 -1rem; padding: 0 1rem; }
    .bf-hero { margin: 0 -1rem; }
    .hero-h1 { font-size: 2.5rem; }
    .metrics-row { grid-template-columns: 1fr 1fr 1fr; }
}
@media (max-width: 640px) {
    .hero-h1 { font-size: 2rem; }
    .metrics-row { grid-template-columns: 1fr; }
    .mt-num { font-size: 2.2rem; }
}
"""
st.markdown(f"<style>{CSS}</style>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SCROLL INDICATOR (JS-injected, shows when page is scrollable)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div id="scroll-ind" class="scroll-indicator">
    <div class="si-label">Scroll</div>
    <div class="si-arrow">↓</div>
</div>
<script>
    (function() {
        function check() {
            var el = document.getElementById('scroll-ind');
            if (!el) return;
            if (document.documentElement.scrollHeight > window.innerHeight + 60) {
                el.style.opacity = '1';
            } else {
                el.style.opacity = '0';
            }
            // hide once user scrolls near bottom
            if ((window.innerHeight + window.scrollY) >= document.body.offsetHeight - 100) {
                el.style.opacity = '0';
            }
        }
        window.addEventListener('scroll', check);
        setTimeout(check, 800);
        setTimeout(check, 2000);
    })();
</script>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# NAV
# ══════════════════════════════════════════════════════════════════════════════
nav_l, nav_m, nav_r = st.columns([2, 1, 2])
with nav_l:
    st.markdown("""
    <div style="display:flex;align-items:center;height:62px;gap:2rem;">
        <span style="font-family:var(--ff-body);font-size:.82rem;font-weight:500;color:var(--brown);letter-spacing:.02em;cursor:pointer;">Documents</span>
        <span style="font-family:var(--ff-body);font-size:.82rem;font-weight:500;color:var(--brown);letter-spacing:.02em;cursor:pointer;">History</span>
    </div>
    """, unsafe_allow_html=True)
with nav_m:
    st.markdown("""
    <div style="display:flex;align-items:center;justify-content:center;height:62px;border-left:1px solid var(--rule);border-right:1px solid var(--rule);">
        <span style="font-family:var(--ff-body);font-weight:700;font-size:.82rem;letter-spacing:.22em;text-transform:uppercase;color:var(--espresso);display:flex;align-items:center;gap:.45rem;">
            <span style="color:var(--sienna);">✦</span> BILLFLOW
        </span>
    </div>
    """, unsafe_allow_html=True)
with nav_r:
    st.markdown("""
    <div style="display:flex;align-items:center;justify-content:flex-end;height:62px;gap:2rem;">
        <span style="font-family:var(--ff-body);font-size:.82rem;font-weight:500;color:var(--brown);letter-spacing:.02em;cursor:pointer;">Reports</span>
        <span style="font-family:var(--ff-body);font-size:.82rem;font-weight:500;color:var(--brown);letter-spacing:.02em;cursor:pointer;">Export</span>
    </div>
    """, unsafe_allow_html=True)
st.markdown('<div style="height:1px;background:var(--rule);margin:0 -2rem;"></div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# HERO
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="bf-hero">
    <span class="hero-eyebrow">Seamless &amp; Smart</span>
    <h1 class="hero-h1">AUTOMATE YOUR<br><em>Bill Reports</em></h1>
    <p class="hero-sub">Upload · Reconcile · Export — in seconds</p>
    <span class="hero-pill">
        Get started below
        <span class="hero-pill-dot">↓</span>
    </span>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SETTINGS TOGGLE
# ══════════════════════════════════════════════════════════════════════════════
_, stog_col = st.columns([5, 1])
with stog_col:
    lbl = "⚙ Settings" if not st.session_state.show_settings else "✕ Close"
    if st.button(lbl, key="stog"):
        st.session_state.show_settings = not st.session_state.show_settings

if st.session_state.show_settings:
    st.markdown('<div class="settings-panel">', unsafe_allow_html=True)
    st.markdown('<div class="sp-title">Preferences</div>', unsafe_allow_html=True)
    sp1, sp2 = st.columns(2, gap="large")
    with sp1:
        sp = st.checkbox("Show file statistics after upload",
                         value=st.session_state.show_preview, key="prev_chk")
        st.session_state.show_preview = sp
    with sp2:
        if st.button("Clear processing history", key="clr_hist"):
            st.session_state.processing_history = []
            st.success("History cleared.")
    st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# STEP 1 — DOCUMENT INTAKE
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="sec-hdr">
    <span class="sec-hdr-title">Documents &amp; intake</span>
    <div class="sec-hdr-rule"></div>
    <span class="sec-hdr-tag">Step 01</span>
</div>
""", unsafe_allow_html=True)

c1, c2, c3 = st.columns(3, gap="medium")

# ── Card 1: Bill Report ──────────────────────────────────────────────────────
with c1:
    st.markdown("""
    <div class="uc-header uc-green">
        <div class="uc-stripe"></div>
        <div class="uc-badge">01 · Primary Register</div>
        <div class="uc-title">Bill Report</div>
        <div class="uc-desc">Master billing workbook with monthly tabs — data is written into this file.</div>
    </div>
    """, unsafe_allow_html=True)
    bill_file = st.file_uploader("Bill Report", type="xlsx",
                                  key="bill", label_visibility="collapsed")
    if bill_file:
        st.markdown('<div class="uc-loaded uc-green">✓ Loaded</div>', unsafe_allow_html=True)
        if st.session_state.show_preview:
            v = validate_file(bill_file, "bill")
            if v["valid"]:
                sheets_n = len(v["sheets"]) if v["sheets"] else "—"
                st.markdown(f"""
                <div class="pv-table">
                    <div class="pv-row"><span>Rows</span><span class="pv-val">{v['rows']}</span></div>
                    <div class="pv-row"><span>Columns</span><span class="pv-val">{v['cols']}</span></div>
                    <div class="pv-row"><span>Sheets</span><span class="pv-val">{sheets_n}</span></div>
                </div>
                """, unsafe_allow_html=True)
            bill_file.seek(0)

# ── Card 2: Job Report ───────────────────────────────────────────────────────
with c2:
    st.markdown("""
    <div class="uc-header uc-orange">
        <div class="uc-stripe"></div>
        <div class="uc-badge">02 · Job Register</div>
        <div class="uc-title">Job Report</div>
        <div class="uc-desc">Shipper details, container numbers, invoice references and SB/BE data.</div>
    </div>
    """, unsafe_allow_html=True)
    job_file = st.file_uploader("Job Report", type="xlsx",
                                 key="job", label_visibility="collapsed")
    if job_file:
        st.markdown('<div class="uc-loaded uc-orange">✓ Loaded</div>', unsafe_allow_html=True)
        if st.session_state.show_preview:
            v = validate_file(job_file, "job")
            if v["valid"]:
                st.markdown(f"""
                <div class="pv-table">
                    <div class="pv-row"><span>Rows</span><span class="pv-val">{v['rows']}</span></div>
                    <div class="pv-row"><span>Columns</span><span class="pv-val">{v['cols']}</span></div>
                </div>
                """, unsafe_allow_html=True)
            job_file.seek(0)

# ── Card 3: E-Invoice ────────────────────────────────────────────────────────
with c3:
    st.markdown("""
    <div class="uc-header uc-gold">
        <div class="uc-stripe"></div>
        <div class="uc-badge">03 · Invoice Data</div>
        <div class="uc-title">E-Invoice Report</div>
        <div class="uc-desc">Electronic invoice records — bill numbers and issuance dates per job.</div>
    </div>
    """, unsafe_allow_html=True)
    einv_file = st.file_uploader("E-Invoice", type="xlsx",
                                  key="einv", label_visibility="collapsed")
    if einv_file:
        st.markdown('<div class="uc-loaded uc-gold">✓ Loaded</div>', unsafe_allow_html=True)
        if st.session_state.show_preview:
            v = validate_file(einv_file, "einv")
            if v["valid"]:
                st.markdown(f"""
                <div class="pv-table">
                    <div class="pv-row"><span>Rows</span><span class="pv-val">{v['rows']}</span></div>
                    <div class="pv-row"><span>Columns</span><span class="pv-val">{v['cols']}</span></div>
                </div>
                """, unsafe_allow_html=True)
            einv_file.seek(0)

# ══════════════════════════════════════════════════════════════════════════════
# STEP 2 — CONFIGURE & RUN
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="sec-hdr" style="margin-top:2.5rem;">
    <span class="sec-hdr-title">Configure &amp; run</span>
    <div class="sec-hdr-rule"></div>
    <span class="sec-hdr-tag">Step 02</span>
</div>
""", unsafe_allow_html=True)

left_cfg, right_cfg = st.columns([3, 1], gap="large")

with left_cfg:
    st.markdown('<div class="cfg-card">', unsafe_allow_html=True)
    st.markdown('<span class="cfg-eyebrow">Ready to process</span>', unsafe_allow_html=True)
    st.markdown('<div class="cfg-title">Set your parameters</div>', unsafe_allow_html=True)

    param_l, param_r = st.columns(2, gap="large")

    with param_l:
        st.markdown('<span class="cfg-label">Billing Month</span>', unsafe_allow_html=True)
        month = st.selectbox(
            "Billing Month", MONTHS,
            index=st.session_state.selected_month,
            label_visibility="collapsed", key="month_sel"
        )
        st.session_state.selected_month = MONTHS.index(month)

    with param_r:
        st.markdown('<span class="cfg-label">Update Mode</span>', unsafe_allow_html=True)
        update_mode = st.radio(
            "Update Mode", ["Append", "Overwrite"],
            label_visibility="collapsed", key="umode"
        )

    if update_mode == "Append":
        st.markdown('<div class="mode-note mode-ok">New job entries will be appended — all existing records remain untouched.</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="mode-note mode-warn">All existing entries for the selected month will be cleared and replaced.</div>', unsafe_allow_html=True)

    st.markdown('<div class="cfg-rule"></div>', unsafe_allow_html=True)

    # Upload status summary
    files_ok = bill_file and job_file and einv_file
    status_html = ""
    for label, f, color in [("Bill Report", bill_file, "var(--green)"),
                              ("Job Report",  job_file,  "var(--orange)"),
                              ("E-Invoice",   einv_file, "var(--gold)")]:
        icon = "✓" if f else "○"
        clr  = color if f else "var(--tan)"
        status_html += f'<span style="font-family:var(--ff-body);font-size:.75rem;font-weight:600;color:{clr};margin-right:1.2rem;">{icon} {label}</span>'
    st.markdown(f'<div style="margin-bottom:.75rem;">{status_html}</div>', unsafe_allow_html=True)

    st.markdown('<div class="exec-btn">', unsafe_allow_html=True)
    process_clicked = st.button("✦  Execute Report Generation", key="proc_main",
                                 use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)  # close cfg-card

with right_cfg:
    if st.session_state.processing_history:
        st.markdown('<div class="hist-card">', unsafe_allow_html=True)
        st.markdown('<span class="hist-eyebrow">Recent</span>', unsafe_allow_html=True)
        st.markdown('<div class="hist-card-title">Activity</div>', unsafe_allow_html=True)
        for entry in reversed(st.session_state.processing_history[-5:]):
            bc = "hi-ok" if entry["status"] == "success" else "hi-err"
            bm = "OK" if entry["status"] == "success" else "ERR"
            st.markdown(f"""
            <div class="hist-item">
                <div class="hi-left">
                    <span class="hi-time">{entry["time"]} · {entry["month"]}</span>
                    <span class="hi-detail">+{entry["added"]} added · {entry.get("duration","—")}</span>
                </div>
                <span class="hi-badge {bc}">{bm}</span>
            </div>
            """, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PROCESSING LOGIC
# ══════════════════════════════════════════════════════════════════════════════
if process_clicked:
    if not (bill_file and job_file and einv_file):
        st.error("Please upload all three documents before running the report.")
    else:
        pb = st.progress(0)
        se = st.empty()
        de = st.empty()
        t0 = datetime.now()

        def upd(p, msg, det=""):
            pb.progress(p)
            se.markdown(f'<div class="prog-txt">{msg}</div>', unsafe_allow_html=True)
            if det:
                de.markdown(f'<div class="prog-detail"><span>{det}</span><span>{p}%</span></div>',
                            unsafe_allow_html=True)

        try:
            upd(10, "Loading workbooks…", "Reading Excel files")
            bill_wb = load_workbook(bill_file)
            job_wb  = load_workbook(job_file)
            einv_wb = load_workbook(einv_file)

            upd(22, "Locating sheets…", f"Looking for tab '{month}'")
            if month not in bill_wb.sheetnames:
                st.error(f"Sheet '{month}' was not found in the Bill Report. Available: {', '.join(bill_wb.sheetnames)}")
                st.stop()

            bill_ws = bill_wb[month]
            job_ws  = job_wb.active
            einv_ws = einv_wb.active

            upd(35, "Mapping columns…", "Analysing header rows")
            bill_cols = col_map(bill_ws, BILL_HEADER_ROW)
            job_cols  = col_map(job_ws,  SRC_HEADER_ROW)
            einv_cols = col_map(einv_ws, SRC_HEADER_ROW)

            upd(45, "Validating structure…", "Checking required fields")
            for col in ["JOB NO","JOB DATE","SHIPPER NAME","INVOICE NO","SB / BE NO","CONTAINER NO"]:
                if col not in job_cols:
                    st.error(f"Missing column in Job Report: '{col}'"); st.stop()
            for col in ["JOB NO","BILL NO","BILL DATE"]:
                if col not in einv_cols:
                    st.error(f"Missing column in E-Invoice: '{col}'"); st.stop()
            for col in ["JOB NO","JOB DATE","PARTY NAME","INVOICE NO","SB / BE NO","CONTAINER NO","BILL NO","BILL DATE"]:
                if col not in bill_cols:
                    st.error(f"Missing column in Bill Report: '{col}'"); st.stop()

            append_mode   = (update_mode == "Append")
            existing_jobs = set()

            if append_mode:
                upd(55, "Scanning existing entries…", "Detecting duplicates")
                existing_jobs = get_existing_jobs(bill_ws, bill_cols, BILL_DATA_START)
                row = next_empty(bill_ws, BILL_DATA_START)
                last_sr = 0
                for r in range(BILL_DATA_START, row):
                    sv = bill_ws.cell(r, 1).value
                    if sv and isinstance(sv, (int, float)):
                        last_sr = max(last_sr, int(sv))
                sr = last_sr + 1
            else:
                upd(55, "Clearing sheet…", "Removing existing entries")
                for r in range(BILL_DATA_START, bill_ws.max_row + 1):
                    for c in range(1, bill_ws.max_column + 1):
                        bill_ws.cell(r, c).value = None
                row = BILL_DATA_START; sr = 1

            upd(65, "Reading job register…", "Parsing job entries")
            jobs = {}
            for r in range(SRC_DATA_START, job_ws.max_row + 1):
                raw = job_ws.cell(r, job_cols["JOB NO"]).value
                if not raw: continue
                jn = clean_job(raw)
                if jn in jobs: continue
                jobs[jn] = {
                    "JOB DATE":    parse_date(job_ws.cell(r, job_cols["JOB DATE"]).value),
                    "PARTY NAME":  job_ws.cell(r, job_cols["SHIPPER NAME"]).value,
                    "INVOICE NO":  job_ws.cell(r, job_cols["INVOICE NO"]).value,
                    "SB / BE NO":  job_ws.cell(r, job_cols["SB / BE NO"]).value,
                    "CONTAINER NO":job_ws.cell(r, job_cols["CONTAINER NO"]).value,
                }

            upd(78, "Reading invoice register…", "Parsing bill entries")
            bill_map = defaultdict(list)
            for r in range(SRC_DATA_START, einv_ws.max_row + 1):
                raw = einv_ws.cell(r, einv_cols["JOB NO"]).value
                bno = einv_ws.cell(r, einv_cols["BILL NO"]).value
                bd  = parse_date(einv_ws.cell(r, einv_cols["BILL DATE"]).value)
                if raw and bno:
                    bill_map[clean_job(raw)].append((str(bno), bd))

            upd(88, "Writing consolidated report…", "Generating final output")
            added = 0; skipped = 0

            for jn, data in jobs.items():
                if append_mode and jn in existing_jobs:
                    skipped += 1; continue
                bill_ws.cell(row, 1).value = sr
                bill_ws.cell(row, bill_cols["JOB NO"]).value = jn
                for field in ["JOB DATE","PARTY NAME","INVOICE NO","SB / BE NO","CONTAINER NO"]:
                    val = data[field]
                    if field == "JOB DATE": val = fmt_date(val)
                    bill_ws.cell(row, bill_cols[field]).value = val
                if jn in bill_map:
                    bills = bill_map[jn]
                    if len(bills) == 1:
                        merged = bills[0][0]
                    else:
                        fb = str(bills[0][0])
                        if "/" in fb:
                            main, suf = fb.rsplit("/", 1); suf = "/" + suf
                        else:
                            main, suf = fb, ""
                        extra = []
                        for i in range(1, len(bills)):
                            bn = str(bills[i][0])
                            if "/" in bn: bn = bn.split("/")[0]
                            extra.append(bn[-4:] if len(bn) >= 4 else bn)
                        merged = main + "/" + "/".join(extra) + suf
                    vd = [b[1] for b in bills if b[1]]
                    ld = max(vd) if vd else None
                    bill_ws.cell(row, bill_cols["BILL NO"]).value   = merged
                    bill_ws.cell(row, bill_cols["BILL DATE"]).value = fmt_date(ld)
                row += 1; sr += 1; added += 1

            upd(96, "Saving workbook…", "Writing to disk")
            out = f"UPDATED_BILL_REPORT_{month}.xlsx"
            bill_wb.save(out)
            pb.progress(100)

            t1  = datetime.now()
            dur = (t1 - t0).total_seconds()
            se.empty(); de.empty(); pb.empty()

            st.session_state.processing_history.append({
                "time": t1.strftime("%H:%M:%S"), "month": month,
                "status": "success", "added": added, "skipped": skipped,
                "duration": f"{dur:.1f}s",
            })

            # ── Results layout ─────────────────────────────────────────────
            res_main, res_side = st.columns([3, 1], gap="large")

            with res_main:
                st.markdown(f"""
                <div class="results-card">
                    <span class="res-eyebrow">All done!</span>
                    <div class="res-title">REPORT GENERATED</div>
                    <div class="metrics-row">
                        <div class="metric-tile">
                            <div class="mt-num">{added}</div>
                            <div class="mt-lbl">Entries Added</div>
                        </div>
                        <div class="metric-tile">
                            <div class="mt-num">{skipped}</div>
                            <div class="mt-lbl">Skipped</div>
                        </div>
                        <div class="metric-tile">
                            <div class="mt-num">{dur:.1f}s</div>
                            <div class="mt-lbl">Duration</div>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

            with res_side:
                st.markdown("""
                <div style="background:var(--cream-lt);border:1px solid var(--rule-md);
                            border-radius:var(--r-lg);padding:1.5rem;box-shadow:var(--shadow-sm);">
                    <div style="font-family:var(--ff-script);font-size:1.1rem;color:var(--sienna);
                                margin-bottom:.1rem;">Export</div>
                    <div style="font-family:var(--ff-display);font-size:1.2rem;color:var(--espresso);
                                margin-bottom:1.1rem;">Download Files</div>
                """, unsafe_allow_html=True)
                with open(out, "rb") as f:
                    st.download_button(
                        "↓ Download Excel", data=f, file_name=out,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True, key="dl_xl"
                    )
                try:
                    df  = pd.read_excel(out, sheet_name=month, header=BILL_HEADER_ROW - 1)
                    csv = df.to_csv(index=False)
                    st.download_button(
                        "↓ Download CSV", data=csv,
                        file_name=f"UPDATED_BILL_REPORT_{month}.csv",
                        mime="text/csv", use_container_width=True, key="dl_csv"
                    )
                except:
                    pass
                st.markdown('</div>', unsafe_allow_html=True)

        except Exception as e:
            pb.empty(); se.empty(); de.empty()
            st.session_state.processing_history.append({
                "time": datetime.now().strftime("%H:%M:%S"), "month": month,
                "status": "error", "added": 0, "skipped": 0, "duration": "—",
            })
            st.error(f"Processing error: {str(e)}")
            with st.expander("View full error details"):
                st.exception(e)