import streamlit as st
from openpyxl import load_workbook
from collections import defaultdict
from datetime import datetime
import pandas as pd

st.set_page_config(
    page_title="BillFlow · Automation",
    page_icon="✦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ── session state ──
for k, v in {
    'selected_month': max(0, datetime.now().month - 2),
    'processing_history': [],
    'show_preview': False,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

MONTHS = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"]
BILL_HEADER_ROW, BILL_DATA_START = 3, 4
SRC_HEADER_ROW,  SRC_DATA_START  = 7, 8

def validate_file(file, file_type):
    try:
        wb = load_workbook(file)
        ws = wb.active
        return {"valid": True, "rows": ws.max_row, "cols": ws.max_column,
                "sheets": wb.sheetnames if file_type == "bill" else None}
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
CSS = """
@import url('https://fonts.googleapis.com/css2?family=Cormorant:ital,wght@0,300;0,400;0,500;0,700;1,300;1,400;1,700&family=Libre+Baskerville:ital,wght@0,400;0,700;1,400&family=Jost:wght@300;400;500;600;700&family=Caveat:wght@500;600;700&display=swap');

:root {
    --paper:     #F4ECD6;
    --paper-lt:  #FAF4E4;
    --paper-dk:  #E8DABC;
    --paper-dkr: #D9C89E;
    --ink:       #1E0F02;
    --ink-mid:   #4A2810;
    --ink-light: #8A6040;
    --rust:      #A63418;
    --rust-lt:   #C4501E;
    --rust-pale: #F5E4DC;
    --sage:      #3A6B4A;
    --sage-pale: #E4EFE7;
    --gold:      #8A6A18;
    --gold-pale: #F4EBD0;
    --border:    rgba(30,15,2,.14);
    --border-md: rgba(30,15,2,.24);
    --border-hv: rgba(30,15,2,.40);
    --ff-display:'Cormorant', Georgia, serif;
    --ff-serif:  'Libre Baskerville', Georgia, serif;
    --ff-body:   'Jost', sans-serif;
    --ff-hand:   'Caveat', cursive;
    --shadow-sm: 0 2px 8px rgba(30,15,2,.07);
    --shadow-md: 0 5px 24px rgba(30,15,2,.11);
    --shadow-lg: 0 12px 48px rgba(30,15,2,.16);
    --shadow-xl: 0 24px 80px rgba(30,15,2,.20);
}

/* ── KILL ALL STREAMLIT CHROME ── */
header[data-testid="stHeader"],
[data-testid="stHeader"],
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"],
[data-testid="stMainBlockContainer"] > div:first-child > div:first-child iframe,
#MainMenu, footer, .stDeployButton {
    display: none !important;
    height: 0 !important;
    visibility: hidden !important;
}

/* ── GLOBAL ── */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

html {
    scroll-behavior: smooth;
    overflow-x: hidden;
}

/* Custom scrollbar — thin, doesn't overlap content */
html { scrollbar-gutter: stable; }
::-webkit-scrollbar { width: 6px; }
::-webkit-scrollbar-track { background: var(--paper-dk); }
::-webkit-scrollbar-thumb { background: var(--paper-dkr); border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: var(--ink-light); }

body,
[data-testid="stAppViewContainer"],
.main,
[data-testid="stAppViewBlockContainer"] {
    font-family: var(--ff-body) !important;
    background: var(--paper) !important;
    color: var(--ink) !important;
    overflow-x: hidden !important;
}

/* Remove ALL default streamlit padding/margin */
.main .block-container {
    padding: 0 !important;
    max-width: 100% !important;
    margin: 0 !important;
}
[data-testid="stVerticalBlock"] { gap: 0 !important; }
[data-testid="stVerticalBlockBorderWrapper"] { padding: 0 !important; }

/* ── CONTENT WRAPPER — prevents scrollbar overlap ── */
.page-wrap {
    width: 100%;
    max-width: 1160px;
    margin: 0 auto;
    padding: 0 2.5rem;
}

/* ══════════════════════════════
   STICKY NAV
══════════════════════════════ */
.v-nav {
    position: sticky;
    top: 0;
    z-index: 500;
    background: var(--paper);
    border-bottom: 1px solid var(--border);
    padding: 0 2.5rem;
}
.v-nav-inner {
    max-width: 1160px;
    margin: 0 auto;
    height: 58px;
    display: flex;
    align-items: center;
    justify-content: space-between;
}
.nav-logo {
    font-family: var(--ff-display);
    font-size: 1.5rem;
    font-weight: 700;
    color: var(--ink);
    letter-spacing: .02em;
    display: flex; align-items: center; gap: .4rem;
}
.nav-logo sup {
    font-family: var(--ff-hand);
    font-size: .85rem;
    color: var(--rust);
    font-weight: 600;
    letter-spacing: 0;
    vertical-align: super;
}
.nav-links {
    display: flex; align-items: center; gap: 2.5rem;
}
.nav-a {
    font-family: var(--ff-body);
    font-size: .8rem; font-weight: 500;
    letter-spacing: .08em; text-transform: uppercase;
    color: var(--ink-light);
    cursor: pointer;
    transition: color .15s;
    text-decoration: none;
}
.nav-a:hover { color: var(--ink); }
.nav-badge {
    font-family: var(--ff-body);
    font-size: .72rem; font-weight: 600;
    letter-spacing: .1em; text-transform: uppercase;
    color: var(--paper-lt);
    background: var(--ink);
    padding: .4rem 1.2rem;
    border-radius: 2px;
    transition: background .15s;
    cursor: pointer;
}
.nav-badge:hover { background: var(--rust); }

/* ══════════════════════════════
   HERO — full viewport height
══════════════════════════════ */
.v-hero {
    min-height: calc(100vh - 58px);
    display: flex; flex-direction: column;
    align-items: center; justify-content: center;
    text-align: center;
    padding: 5rem 2.5rem 4rem;
    position: relative; overflow: hidden;
    background: var(--paper);
}

/* Decorative horizontal rules */
.v-hero::before {
    content: '';
    position: absolute;
    top: 3.5rem; left: 8%; right: 8%;
    height: 1px;
    background: linear-gradient(90deg, transparent, var(--border-md) 20%, var(--border-md) 80%, transparent);
}
.v-hero::after {
    content: '';
    position: absolute;
    top: 3.75rem; left: 8%; right: 8%;
    height: 1px;
    background: linear-gradient(90deg, transparent, var(--border) 20%, var(--border) 80%, transparent);
}

.hero-vol {
    font-family: var(--ff-body);
    font-size: .68rem; font-weight: 600;
    letter-spacing: .28em; text-transform: uppercase;
    color: var(--ink-light);
    margin-bottom: 1.5rem;
    opacity: 0; animation: revealUp .6s .1s ease forwards;
}
.hero-vol span { color: var(--rust); }

.hero-h {
    font-family: var(--ff-display);
    font-size: clamp(4.5rem, 10vw, 9rem);
    font-weight: 700;
    color: var(--ink);
    line-height: .9;
    letter-spacing: -.02em;
    margin-bottom: .5rem;
    opacity: 0; animation: revealUp .65s .18s ease forwards;
}
.hero-h em {
    font-style: italic;
    font-weight: 300;
    color: var(--ink-mid);
}

.hero-script {
    font-family: var(--ff-hand);
    font-size: clamp(2.2rem, 5vw, 4rem);
    font-weight: 600;
    color: var(--rust);
    line-height: 1;
    margin-bottom: 1.75rem;
    opacity: 0; animation: revealUp .65s .28s ease forwards;
    display: block;
}

.hero-rule {
    width: 120px; height: 1px;
    background: var(--border-md);
    margin: 0 auto 1.75rem;
    opacity: 0; animation: revealFade .8s .38s ease forwards;
}

.hero-sub {
    font-family: var(--ff-serif);
    font-style: italic;
    font-size: clamp(.95rem, 1.8vw, 1.15rem);
    color: var(--ink-light);
    max-width: 520px;
    line-height: 1.65;
    margin: 0 auto 2.5rem;
    opacity: 0; animation: revealUp .65s .42s ease forwards;
}

.hero-actions {
    display: flex; align-items: center; gap: 1rem;
    opacity: 0; animation: revealUp .65s .52s ease forwards;
}
.hero-btn-primary {
    font-family: var(--ff-body);
    font-size: .8rem; font-weight: 700;
    letter-spacing: .12em; text-transform: uppercase;
    color: var(--paper-lt);
    background: var(--ink);
    padding: .9rem 2.5rem;
    border: 2px solid var(--ink);
    cursor: pointer;
    transition: all .2s;
}
.hero-btn-primary:hover {
    background: var(--rust);
    border-color: var(--rust);
    transform: translateY(-2px);
    box-shadow: var(--shadow-md);
}
.hero-btn-ghost {
    font-family: var(--ff-body);
    font-size: .8rem; font-weight: 600;
    letter-spacing: .1em; text-transform: uppercase;
    color: var(--ink-mid);
    background: transparent;
    padding: .9rem 2rem;
    border: 1px solid var(--border-md);
    cursor: pointer;
    transition: all .2s;
}
.hero-btn-ghost:hover {
    border-color: var(--ink);
    color: var(--ink);
    background: rgba(30,15,2,.04);
}

/* Scroll hint */
.hero-scroll {
    position: absolute;
    bottom: 2rem; left: 50%;
    transform: translateX(-50%);
    display: flex; flex-direction: column; align-items: center; gap: .4rem;
    opacity: 0; animation: revealFade 1s 1s ease forwards;
}
.hero-scroll-line {
    width: 1px; height: 40px;
    background: linear-gradient(180deg, transparent, var(--border-hv));
    animation: scrollPulse 2s ease-in-out infinite;
}
.hero-scroll-label {
    font-family: var(--ff-body);
    font-size: .6rem; font-weight: 600;
    letter-spacing: .22em; text-transform: uppercase;
    color: var(--ink-light);
}

/* ══════════════════════════════
   BAND / MARQUEE
══════════════════════════════ */
.v-band {
    background: var(--ink);
    padding: .85rem 0;
    overflow: hidden;
    white-space: nowrap;
}
.band-track {
    display: inline-flex;
    animation: marquee 24s linear infinite;
}
.band-item {
    font-family: var(--ff-display);
    font-size: 1rem; font-weight: 700;
    letter-spacing: .08em; text-transform: uppercase;
    color: var(--paper-dk);
    padding: 0 2rem;
}
.band-item span { color: var(--rust); margin-right: 2rem; }

/* ══════════════════════════════
   SECTION DIVIDER TITLE
══════════════════════════════ */
.v-section {
    padding: 5rem 2.5rem;
    max-width: 1160px;
    margin: 0 auto;
}
.v-section-narrow {
    padding: 5rem 2.5rem;
    max-width: 900px;
    margin: 0 auto;
}

.sec-kicker {
    font-family: var(--ff-body);
    font-size: .65rem; font-weight: 700;
    letter-spacing: .28em; text-transform: uppercase;
    color: var(--rust);
    display: flex; align-items: center; gap: .75rem;
    margin-bottom: 1.2rem;
}
.sec-kicker::before {
    content: ''; width: 32px; height: 1px;
    background: var(--rust); flex-shrink: 0;
}
.sec-title {
    font-family: var(--ff-display);
    font-size: clamp(2.8rem, 5.5vw, 5rem);
    font-weight: 700;
    color: var(--ink);
    line-height: .95;
    letter-spacing: -.02em;
    margin-bottom: .6rem;
}
.sec-title em {
    font-style: italic;
    font-weight: 300;
    color: var(--ink-mid);
}
.sec-sub {
    font-family: var(--ff-serif);
    font-style: italic;
    font-size: 1.05rem;
    color: var(--ink-light);
    line-height: 1.65;
    max-width: 580px;
    margin-bottom: 3rem;
}

/* ══════════════════════════════
   UPLOAD CARDS (editorial tiles)
══════════════════════════════ */
.upload-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 1.5rem;
    margin-bottom: 1rem;
}
.upload-tile {
    background: var(--paper-lt);
    border: 1px solid var(--border);
    padding: 0;
    position: relative;
    overflow: hidden;
    transition: border-color .25s, box-shadow .25s, transform .25s;
}
.upload-tile:hover {
    border-color: var(--border-hv);
    box-shadow: var(--shadow-lg);
    transform: translateY(-4px);
}
.tile-accent { height: 4px; width: 100%; }
.tile-body { padding: 1.75rem 1.6rem 1.4rem; }
.tile-num {
    font-family: var(--ff-display);
    font-size: 3.5rem; font-weight: 700;
    line-height: 1;
    color: var(--paper-dk);
    float: right;
    margin: -.5rem -.2rem 0 0;
    letter-spacing: -.03em;
}
.tile-kicker {
    font-family: var(--ff-body);
    font-size: .62rem; font-weight: 700;
    letter-spacing: .2em; text-transform: uppercase;
    color: var(--ink-light);
    margin-bottom: .65rem;
    display: block;
}
.tile-title {
    font-family: var(--ff-display);
    font-size: 1.5rem; font-weight: 700;
    color: var(--ink);
    line-height: 1.1;
    margin-bottom: .45rem;
    letter-spacing: -.01em;
}
.tile-desc {
    font-family: var(--ff-serif);
    font-style: italic;
    font-size: .85rem;
    color: var(--ink-light);
    line-height: 1.55;
    margin-bottom: 1.1rem;
}
.tile-rule { height: 1px; background: var(--border); margin-bottom: 1.1rem; }
.tile-status {
    display: inline-flex; align-items: center; gap: .4rem;
    font-family: var(--ff-body);
    font-size: .68rem; font-weight: 700;
    letter-spacing: .1em; text-transform: uppercase;
    padding: .3rem .8rem;
    border: 1px solid var(--border-md);
    color: var(--ink-light);
}
.tile-status.loaded {
    color: var(--sage);
    background: var(--sage-pale);
    border-color: rgba(58,107,74,.25);
}

/* preview */
.tile-preview {
    background: var(--paper-dk);
    padding: .6rem .85rem;
    margin-top: .6rem;
    font-family: 'Courier New', monospace;
    font-size: .68rem;
}
.tp-row {
    display: flex; justify-content: space-between;
    color: var(--ink-mid); padding: .15rem 0;
    border-bottom: 1px solid var(--border);
}
.tp-row:last-child { border-bottom: none; }
.tp-val { font-weight: 700; color: var(--rust); }

/* ══════════════════════════════
   FULL-BLEED DARK SECTION
══════════════════════════════ */
.v-dark {
    background: var(--ink);
    padding: 5rem 2.5rem;
}
.v-dark-inner {
    max-width: 1160px;
    margin: 0 auto;
}
.dark-kicker {
    font-family: var(--ff-body);
    font-size: .65rem; font-weight: 700;
    letter-spacing: .28em; text-transform: uppercase;
    color: var(--rust-lt);
    display: flex; align-items: center; gap: .75rem;
    margin-bottom: 1.2rem;
}
.dark-kicker::before {
    content: ''; width: 32px; height: 1px;
    background: var(--rust-lt); flex-shrink: 0;
}
.dark-title {
    font-family: var(--ff-display);
    font-size: clamp(2.4rem, 4.5vw, 4rem);
    font-weight: 700;
    color: var(--paper-lt);
    line-height: .95;
    letter-spacing: -.02em;
    margin-bottom: 2.5rem;
}
.dark-title em {
    font-style: italic;
    font-weight: 300;
    color: rgba(244,236,214,.5);
}

/* ── CONFIG GRID on dark bg ── */
.cfg-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 2rem;
    margin-bottom: 2rem;
}
.cfg-block { }
.cfg-label {
    font-family: var(--ff-body);
    font-size: .65rem; font-weight: 700;
    letter-spacing: .2em; text-transform: uppercase;
    color: rgba(244,236,214,.4);
    display: block; margin-bottom: .6rem;
}

/* ── file upload on dark bg ── */
[data-testid="stFileUploader"] section {
    background: rgba(255,255,255,.05) !important;
    border: 1px solid rgba(255,255,255,.12) !important;
    border-radius: 2px !important;
    padding: .75rem 1rem !important;
    transition: border-color .2s !important;
}
[data-testid="stFileUploader"] section:hover {
    border-color: var(--rust-lt) !important;
    background: rgba(255,255,255,.08) !important;
}
[data-testid="stFileUploader"] section p,
[data-testid="stFileUploader"] section span,
[data-testid="stFileUploader"] section small {
    font-family: var(--ff-body) !important;
    font-size: .78rem !important;
    color: rgba(244,236,214,.45) !important;
}
[data-testid="stFileUploader"] section button {
    background: rgba(255,255,255,.08) !important;
    border: 1px solid rgba(255,255,255,.18) !important;
    color: rgba(244,236,214,.7) !important;
    font-family: var(--ff-body) !important;
    font-size: .75rem !important;
    border-radius: 1px !important;
    transition: background .15s !important;
}
[data-testid="stFileUploader"] section button:hover {
    background: rgba(255,255,255,.14) !important;
}

/* ── selectbox on dark ── */
[data-testid="stSelectbox"] > div > div {
    background: rgba(255,255,255,.06) !important;
    border: 1px solid rgba(255,255,255,.15) !important;
    border-radius: 2px !important;
    color: var(--paper-lt) !important;
    font-family: 'Courier New', monospace !important;
    font-size: .9rem !important;
    letter-spacing: .06em !important;
}
[data-testid="stSelectbox"] > div > div:hover { border-color: var(--rust-lt) !important; }
[data-testid="stSelectbox"] > div > div > div { color: var(--paper-lt) !important; }
[data-testid="stSelectbox"] svg { fill: rgba(244,236,214,.45) !important; }

/* ── radio on dark — segmented pill ── */
[data-testid="stRadio"] > div {
    display: flex !important;
    gap: 0 !important;
    flex-direction: row !important;
    background: rgba(255,255,255,.06) !important;
    border: 1px solid rgba(255,255,255,.15) !important;
    border-radius: 2px !important;
    padding: 3px !important;
}
[data-testid="stRadio"] label > div:first-child,
[data-testid="stRadio"] input[type="radio"] {
    display: none !important;
}
[data-testid="stRadio"] label {
    background: transparent !important;
    border: none !important;
    border-radius: 1px !important;
    padding: .55rem 1.4rem !important;
    font-family: var(--ff-body) !important;
    font-size: .82rem !important; font-weight: 500 !important;
    color: rgba(244,236,214,.5) !important;
    cursor: pointer !important;
    transition: all .18s !important;
    flex: 1 !important; text-align: center !important;
    letter-spacing: .04em !important;
}
[data-testid="stRadio"] label:hover {
    color: var(--paper-lt) !important;
    background: rgba(255,255,255,.06) !important;
}
[data-testid="stRadio"] label:has(input:checked) {
    background: var(--rust) !important;
    color: #fff !important; font-weight: 700 !important;
    letter-spacing: .06em !important;
    box-shadow: 0 2px 10px rgba(164,52,24,.4) !important;
}

/* ── checkbox on dark ── */
[data-testid="stCheckbox"] label,
[data-testid="stCheckbox"] span {
    font-family: var(--ff-body) !important;
    font-size: .85rem !important;
    color: rgba(244,236,214,.6) !important;
}

/* ── mode note ── */
.mode-note {
    font-family: var(--ff-serif);
    font-style: italic;
    font-size: .85rem; line-height: 1.6;
    padding: .7rem 1rem;
    margin-top: .65rem;
    border-left: 2px solid;
}
.mode-ok   { color: #7EC49A; background: rgba(58,107,74,.15); border-color: var(--sage); }
.mode-warn { color: #E09888; background: rgba(166,52,24,.15); border-color: var(--rust-lt); }

/* ── file status row ── */
.file-status-row {
    display: flex; align-items: center; gap: 1.5rem;
    margin-bottom: 1.5rem;
    padding: .75rem 1rem;
    background: rgba(255,255,255,.04);
    border: 1px solid rgba(255,255,255,.08);
}
.fs-item {
    display: flex; align-items: center; gap: .45rem;
    font-family: var(--ff-body);
    font-size: .72rem; font-weight: 600;
    letter-spacing: .06em; text-transform: uppercase;
}

/* ── BUTTONS on dark ── */
.stButton > button {
    background: transparent !important;
    border: 1px solid rgba(255,255,255,.25) !important;
    color: rgba(244,236,214,.7) !important;
    font-family: var(--ff-body) !important;
    font-size: .75rem !important; font-weight: 600 !important;
    letter-spacing: .1em !important; text-transform: uppercase !important;
    border-radius: 1px !important;
    padding: .65rem 1.5rem !important;
    transition: all .2s !important;
    width: 100% !important;
}
.stButton > button:hover {
    background: rgba(255,255,255,.08) !important;
    border-color: rgba(255,255,255,.45) !important;
    color: var(--paper-lt) !important;
}
.exec-wrap .stButton > button {
    background: var(--rust) !important;
    border: 2px solid var(--rust) !important;
    color: #fff !important;
    font-size: .82rem !important;
    padding: 1rem 2rem !important;
    letter-spacing: .14em !important;
    box-shadow: 0 4px 20px rgba(164,52,24,.35) !important;
    width: 100% !important;
}
.exec-wrap .stButton > button:hover {
    background: #C4501E !important;
    border-color: #C4501E !important;
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 32px rgba(164,52,24,.45) !important;
}

/* ── download button ── */
.stDownloadButton > button {
    background: var(--sage) !important;
    border: 2px solid var(--sage) !important;
    color: #fff !important;
    font-family: var(--ff-body) !important;
    font-size: .75rem !important; font-weight: 700 !important;
    letter-spacing: .1em !important; text-transform: uppercase !important;
    border-radius: 1px !important;
    padding: .75rem 1.5rem !important;
    width: 100% !important;
    transition: all .2s !important;
}
.stDownloadButton > button:hover {
    background: #2E5A3C !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 16px rgba(58,107,74,.35) !important;
}

/* ── progress ── */
[data-testid="stProgress"] > div {
    background: rgba(255,255,255,.1) !important;
    border-radius: 1px !important; height: 3px !important;
}
[data-testid="stProgress"] > div > div {
    background: linear-gradient(90deg, var(--rust), #D4501E, var(--gold)) !important;
    border-radius: 1px !important;
}

/* ── native alerts ── */
[data-testid="stAlert"] {
    font-family: var(--ff-body) !important;
    font-size: .85rem !important;
    border-radius: 0 !important;
}

/* ══════════════════════════════
   RESULTS (light section below dark)
══════════════════════════════ */
.v-results {
    background: var(--paper-lt);
    border-top: 1px solid var(--border-md);
    padding: 4rem 2.5rem;
}
.v-results-inner { max-width: 1160px; margin: 0 auto; }
.results-title {
    font-family: var(--ff-display);
    font-size: clamp(3rem, 6vw, 5.5rem);
    font-weight: 700;
    color: var(--ink);
    line-height: .9; letter-spacing: -.02em;
    margin-bottom: 2.5rem;
}
.results-title em { font-style: italic; font-weight: 300; color: var(--sage); }
.metrics-row {
    display: grid; grid-template-columns: repeat(3, 1fr);
    gap: 1.5rem; margin-bottom: 2.5rem;
}
.metric-tile {
    border: 1px solid var(--border-md);
    padding: 1.75rem;
    position: relative;
}
.metric-tile::after {
    content: '';
    position: absolute;
    bottom: 0; left: 0; right: 0; height: 3px;
    background: var(--paper-dkr);
}
.mt-n {
    font-family: var(--ff-display);
    font-size: 4.5rem; font-weight: 700;
    color: var(--rust);
    line-height: .9; letter-spacing: -.03em;
    margin-bottom: .5rem;
}
.mt-l {
    font-family: var(--ff-body);
    font-size: .65rem; font-weight: 700;
    letter-spacing: .2em; text-transform: uppercase;
    color: var(--ink-light);
}

/* ── history in results ── */
.history-strip { }
.h-title {
    font-family: var(--ff-display);
    font-size: 1.6rem; font-weight: 700;
    color: var(--paper-lt);
    margin-bottom: 1rem;
    letter-spacing: -.01em;
}
.h-item {
    display: flex; align-items: center; justify-content: space-between;
    padding: .6rem 0;
    border-bottom: 1px solid rgba(255,255,255,.08);
}
.h-item:last-child { border-bottom: none; }
.h-time {
    font-family: 'Courier New', monospace;
    font-size: .72rem; font-weight: 700;
    color: var(--paper-lt);
}
.h-detail {
    font-family: var(--ff-body);
    font-size: .7rem;
    color: rgba(244,236,214,.45);
}
.h-badge {
    font-family: var(--ff-body);
    font-size: .6rem; font-weight: 700;
    letter-spacing: .12em; text-transform: uppercase;
    padding: .18rem .6rem;
}
.h-ok  { color: #7EC49A; background: rgba(58,107,74,.25); }
.h-err { color: #E09888; background: rgba(166,52,24,.25); }

/* ── prog status ── */
.prog-status {
    font-family: 'Courier New', monospace;
    font-size: .76rem;
    color: var(--rust-lt);
    letter-spacing: .06em;
    padding: .4rem 0;
    text-align: center;
}
.prog-detail {
    display: flex; justify-content: space-between;
    font-family: 'Courier New', monospace;
    font-size: .65rem;
    color: rgba(244,236,214,.35);
    margin-top: .2rem;
}

/* ── expander ── */
.streamlit-expanderHeader {
    font-family: var(--ff-body) !important;
    font-size: .82rem !important;
    color: rgba(244,236,214,.5) !important;
    background: rgba(255,255,255,.04) !important;
    border-radius: 1px !important;
}

/* ══════════════════════════════
   SCROLL-REVEAL utility
══════════════════════════════ */
.reveal {
    opacity: 0;
    transform: translateY(28px);
    transition: opacity .7s ease, transform .7s ease;
}
.reveal.visible {
    opacity: 1;
    transform: translateY(0);
}
.reveal-delay-1 { transition-delay: .1s; }
.reveal-delay-2 { transition-delay: .2s; }
.reveal-delay-3 { transition-delay: .3s; }
.reveal-delay-4 { transition-delay: .4s; }

/* ── KEYFRAMES ── */
@keyframes revealUp {
    from { opacity:0; transform:translateY(20px); }
    to   { opacity:1; transform:translateY(0); }
}
@keyframes revealFade {
    from { opacity:0; }
    to   { opacity:1; }
}
@keyframes scrollPulse {
    0%, 100% { opacity:.3; transform:scaleY(1); }
    50%       { opacity:1;  transform:scaleY(1.15); }
}
@keyframes marquee {
    from { transform:translateX(0); }
    to   { transform:translateX(-50%); }
}

/* ── RESPONSIVE ── */
@media (max-width: 900px) {
    .v-section, .v-section-narrow, .v-dark, .v-results { padding: 3.5rem 1.25rem; }
    .v-nav { padding: 0 1.25rem; }
    .v-hero { padding: 4rem 1.25rem 3rem; }
    .upload-grid { grid-template-columns: 1fr; }
    .cfg-grid { grid-template-columns: 1fr; gap: 1.5rem; }
    .metrics-row { grid-template-columns: 1fr; }
    .hero-h { font-size: 3.5rem; }
}
@media (max-width: 640px) {
    .hero-h { font-size: 2.8rem; }
    .hero-script { font-size: 2rem; }
    .sec-title { font-size: 2.4rem; }
    .dark-title { font-size: 2rem; }
}
"""

st.markdown(f"<style>{CSS}</style>", unsafe_allow_html=True)

# ── Scroll-reveal JS ──────────────────────────────────────────────────────────
st.markdown("""
<script>
(function() {
    function init() {
        var els = document.querySelectorAll('.reveal');
        if (!els.length) { setTimeout(init, 300); return; }
        var io = new IntersectionObserver(function(entries) {
            entries.forEach(function(e) {
                if (e.isIntersecting) { e.target.classList.add('visible'); io.unobserve(e.target); }
            });
        }, { threshold: 0.12 });
        els.forEach(function(el) { io.observe(el); });
    }
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        setTimeout(init, 200);
    }
})();
</script>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# NAV
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="v-nav">
  <div class="v-nav-inner">
    <div class="nav-logo">BillFlow <sup>Studio</sup></div>
    <div class="nav-links">
      <a class="nav-a" href="#">Upload</a>
      <a class="nav-a" href="#">Configure</a>
      <a class="nav-a" href="#">History</a>
      <span class="nav-badge">Process Now</span>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# HERO
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="v-hero">
  <div class="hero-vol">Vol. I &nbsp;·&nbsp; <span>Report Automation</span> &nbsp;·&nbsp; Est. 2024</div>
  <h1 class="hero-h">AUTOMATE<br><em>Your Billing</em></h1>
  <span class="hero-script">Reports, reconciled.</span>
  <div class="hero-rule"></div>
  <p class="hero-sub">
    Upload your three source documents, configure the month and mode,
    and let BillFlow consolidate every job and invoice into a clean, 
    ready-to-use Excel report in seconds.
  </p>
  <div class="hero-actions">
    <span class="hero-btn-primary">Begin Upload ↓</span>
    <span class="hero-btn-ghost">Learn More</span>
  </div>
  <div class="hero-scroll">
    <span class="hero-scroll-label">Scroll</span>
    <div class="hero-scroll-line"></div>
  </div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# MARQUEE BAND
# ══════════════════════════════════════════════════════════════════════════════
band = "Bill Report &nbsp; <span>✦</span> &nbsp; Job Register &nbsp; <span>✦</span> &nbsp; E-Invoice &nbsp; <span>✦</span> &nbsp; Reconcile &nbsp; <span>✦</span> &nbsp; Export &nbsp; <span>✦</span> &nbsp; "
st.markdown(f"""
<div class="v-band">
  <div class="band-track">
    <span class="band-item">{band}{band}{band}{band}</span>
  </div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 01 — UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="v-section">
  <div class="reveal sec-kicker">Step One</div>
  <div class="reveal reveal-delay-1 sec-title">Upload<br><em>Your Documents</em></div>
  <div class="reveal reveal-delay-2 sec-sub">
    Three source files power the reconciliation. Each plays a distinct role — 
    the master ledger, the job register, and the invoice index.
  </div>
</div>
""", unsafe_allow_html=True)

# ── Upload columns ────────────────────────────────────────────────────────────
uc1, uc2, uc3 = st.columns(3, gap="medium")

with uc1:
    st.markdown("""
    <div class="reveal upload-tile">
      <div class="tile-accent" style="background:#3A6B4A;"></div>
      <div class="tile-body">
        <div class="tile-num">I</div>
        <span class="tile-kicker">Primary Register</span>
        <div class="tile-title">Bill Report</div>
        <div class="tile-desc">
          Master billing workbook with monthly tabs — 
          all consolidated data is written into this file.
        </div>
        <div class="tile-rule"></div>
      </div>
    </div>
    """, unsafe_allow_html=True)
    bill_file = st.file_uploader("Bill Report .xlsx", type="xlsx",
                                  key="bill", label_visibility="collapsed")
    if bill_file:
        st.markdown('<div class="tile-status loaded">✓ &nbsp; Loaded</div>', unsafe_allow_html=True)
        if st.session_state.show_preview:
            v = validate_file(bill_file, "bill")
            if v["valid"]:
                sheets_n = len(v["sheets"]) if v["sheets"] else "—"
                st.markdown(f"""<div class="tile-preview">
                    <div class="tp-row"><span>Rows</span><span class="tp-val">{v['rows']}</span></div>
                    <div class="tp-row"><span>Columns</span><span class="tp-val">{v['cols']}</span></div>
                    <div class="tp-row"><span>Sheets</span><span class="tp-val">{sheets_n}</span></div>
                </div>""", unsafe_allow_html=True)
            bill_file.seek(0)
    else:
        st.markdown('<div class="tile-status">○ &nbsp; Awaiting file</div>', unsafe_allow_html=True)

with uc2:
    st.markdown("""
    <div class="reveal reveal-delay-1 upload-tile">
      <div class="tile-accent" style="background:#A63418;"></div>
      <div class="tile-body">
        <div class="tile-num">II</div>
        <span class="tile-kicker">Job Register</span>
        <div class="tile-title">Job Report</div>
        <div class="tile-desc">
          Shipper details, container numbers, invoice 
          references and SB/BE data for each job.
        </div>
        <div class="tile-rule"></div>
      </div>
    </div>
    """, unsafe_allow_html=True)
    job_file = st.file_uploader("Job Report .xlsx", type="xlsx",
                                 key="job", label_visibility="collapsed")
    if job_file:
        st.markdown('<div class="tile-status loaded">✓ &nbsp; Loaded</div>', unsafe_allow_html=True)
        if st.session_state.show_preview:
            v = validate_file(job_file, "job")
            if v["valid"]:
                st.markdown(f"""<div class="tile-preview">
                    <div class="tp-row"><span>Rows</span><span class="tp-val">{v['rows']}</span></div>
                    <div class="tp-row"><span>Columns</span><span class="tp-val">{v['cols']}</span></div>
                </div>""", unsafe_allow_html=True)
            job_file.seek(0)
    else:
        st.markdown('<div class="tile-status">○ &nbsp; Awaiting file</div>', unsafe_allow_html=True)

with uc3:
    st.markdown("""
    <div class="reveal reveal-delay-2 upload-tile">
      <div class="tile-accent" style="background:#8A6A18;"></div>
      <div class="tile-body">
        <div class="tile-num">III</div>
        <span class="tile-kicker">Invoice Data</span>
        <div class="tile-title">E-Invoice Report</div>
        <div class="tile-desc">
          Electronic invoice records — bill numbers 
          and issuance dates matched to each job.
        </div>
        <div class="tile-rule"></div>
      </div>
    </div>
    """, unsafe_allow_html=True)
    einv_file = st.file_uploader("E-Invoice .xlsx", type="xlsx",
                                  key="einv", label_visibility="collapsed")
    if einv_file:
        st.markdown('<div class="tile-status loaded">✓ &nbsp; Loaded</div>', unsafe_allow_html=True)
        if st.session_state.show_preview:
            v = validate_file(einv_file, "einv")
            if v["valid"]:
                st.markdown(f"""<div class="tile-preview">
                    <div class="tp-row"><span>Rows</span><span class="tp-val">{v['rows']}</span></div>
                    <div class="tp-row"><span>Columns</span><span class="tp-val">{v['cols']}</span></div>
                </div>""", unsafe_allow_html=True)
            einv_file.seek(0)
    else:
        st.markdown('<div class="tile-status">○ &nbsp; Awaiting file</div>', unsafe_allow_html=True)

# preview toggle
st.markdown('<div style="max-width:1160px;margin:1rem auto;padding:0 0 1rem;">', unsafe_allow_html=True)
_, pt_col = st.columns([5, 1])
with pt_col:
    sp = st.checkbox("Show preview", value=st.session_state.show_preview, key="prev_chk")
    st.session_state.show_preview = sp
st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 02 — DARK CONFIG
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="v-dark">
  <div class="v-dark-inner">
    <div class="reveal dark-kicker">Step Two</div>
    <div class="reveal reveal-delay-1 dark-title">Configure<br><em>&amp; Run</em></div>
  </div>
</div>
""", unsafe_allow_html=True)

# Config fields — inside dark section using a 2-col layout
cfg_left, cfg_right = st.columns([5, 2], gap="large")

with cfg_left:
    st.markdown('<div class="v-dark" style="padding:0 2.5rem 3rem;">', unsafe_allow_html=True)
    st.markdown('<div style="max-width:1160px; margin:0 auto;">', unsafe_allow_html=True)

    # file status
    files_ready = bill_file and job_file and einv_file
    b_c = "#7EC49A" if bill_file  else "rgba(244,236,214,.25)"
    j_c = "#7EC49A" if job_file   else "rgba(244,236,214,.25)"
    e_c = "#7EC49A" if einv_file  else "rgba(244,236,214,.25)"
    b_i = "✓" if bill_file  else "○"
    j_i = "✓" if job_file   else "○"
    e_i = "✓" if einv_file  else "○"
    st.markdown(f"""
    <div class="reveal file-status-row">
      <span class="fs-item" style="color:{b_c};">{b_i} &nbsp;Bill Report</span>
      <span class="fs-item" style="color:{j_c};">{j_i} &nbsp;Job Report</span>
      <span class="fs-item" style="color:{e_c};">{e_i} &nbsp;E-Invoice</span>
    </div>
    """, unsafe_allow_html=True)

    cfg_c1, cfg_c2 = st.columns(2, gap="large")
    with cfg_c1:
        st.markdown('<span class="cfg-label">Billing Month</span>', unsafe_allow_html=True)
        month = st.selectbox("Month", MONTHS,
                             index=st.session_state.selected_month,
                             label_visibility="collapsed", key="month_sel")
        st.session_state.selected_month = MONTHS.index(month)

    with cfg_c2:
        st.markdown('<span class="cfg-label">Update Mode</span>', unsafe_allow_html=True)
        update_mode = st.radio("Mode", ["Append", "Overwrite"],
                               label_visibility="collapsed", key="umode")

    if update_mode == "Append":
        st.markdown('<div class="mode-note mode-ok reveal">New entries appended — existing records remain untouched.</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="mode-note mode-warn reveal">All existing entries for this month will be cleared and replaced.</div>', unsafe_allow_html=True)

    st.markdown('<div style="height:1.5rem;"></div>', unsafe_allow_html=True)
    st.markdown('<div class="exec-wrap reveal">', unsafe_allow_html=True)
    process_clicked = st.button("✦  Execute Report Generation", key="proc_main",
                                 use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('</div></div>', unsafe_allow_html=True)

with cfg_right:
    # History panel — also dark themed
    st.markdown('<div class="v-dark" style="padding:0 1.5rem 3rem 0;">', unsafe_allow_html=True)
    if st.session_state.processing_history:
        st.markdown('<div class="h-title">Activity</div>', unsafe_allow_html=True)
        for entry in reversed(st.session_state.processing_history[-6:]):
            bc = "h-ok" if entry["status"] == "success" else "h-err"
            bm = "DONE" if entry["status"] == "success" else "ERR"
            st.markdown(f"""
            <div class="h-item">
              <div>
                <div class="h-time">{entry["time"]} · {entry["month"]}</div>
                <div class="h-detail">+{entry["added"]} added &nbsp;·&nbsp; {entry.get("duration","—")}</div>
              </div>
              <span class="h-badge {bc}">{bm}</span>
            </div>
            """, unsafe_allow_html=True)
        if st.button("Clear History", key="clr_hist"):
            st.session_state.processing_history = []
            st.rerun()
    else:
        st.markdown("""
        <div style="font-family:var(--ff-serif);font-style:italic;font-size:.9rem;
                    color:rgba(244,236,214,.3);padding-top:1rem;line-height:1.6;">
            No history yet — processed runs will appear here.
        </div>
        """, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PROCESSING LOGIC
# ══════════════════════════════════════════════════════════════════════════════
if process_clicked:
    if not (bill_file and job_file and einv_file):
        st.error("Please upload all three documents before running.")
    else:
        # progress lives in the dark section context
        pb = st.progress(0)
        se = st.empty()
        de = st.empty()
        t0 = datetime.now()

        def upd(p, msg, det=""):
            pb.progress(p)
            se.markdown(f'<div class="prog-status">{msg}</div>', unsafe_allow_html=True)
            if det:
                de.markdown(f'<div class="prog-detail"><span>{det}</span><span>{p}%</span></div>',
                            unsafe_allow_html=True)

        try:
            upd(10, "Loading workbooks…", "Reading Excel files")
            bill_wb = load_workbook(bill_file)
            job_wb  = load_workbook(job_file)
            einv_wb = load_workbook(einv_file)

            upd(22, "Locating sheet…", f"Looking for tab '{month}'")
            if month not in bill_wb.sheetnames:
                st.error(f"Sheet '{month}' not found. Available: {', '.join(bill_wb.sheetnames)}")
                st.stop()

            bill_ws = bill_wb[month]
            job_ws  = job_wb.active
            einv_ws = einv_wb.active

            upd(35, "Mapping columns…", "Analysing header rows")
            bill_cols = col_map(bill_ws, BILL_HEADER_ROW)
            job_cols  = col_map(job_ws,  SRC_HEADER_ROW)
            einv_cols = col_map(einv_ws, SRC_HEADER_ROW)

            upd(45, "Validating structure…", "Checking required fields")
            for c in ["JOB NO","JOB DATE","SHIPPER NAME","INVOICE NO","SB / BE NO","CONTAINER NO"]:
                if c not in job_cols: st.error(f"Missing in Job Report: '{c}'"); st.stop()
            for c in ["JOB NO","BILL NO","BILL DATE"]:
                if c not in einv_cols: st.error(f"Missing in E-Invoice: '{c}'"); st.stop()
            for c in ["JOB NO","JOB DATE","PARTY NAME","INVOICE NO","SB / BE NO","CONTAINER NO","BILL NO","BILL DATE"]:
                if c not in bill_cols: st.error(f"Missing in Bill Report: '{c}'"); st.stop()

            append_mode = (update_mode == "Append")
            existing_jobs = set()

            if append_mode:
                upd(55, "Scanning existing entries…", "Detecting duplicates")
                existing_jobs = get_existing_jobs(bill_ws, bill_cols, BILL_DATA_START)
                row = next_empty(bill_ws, BILL_DATA_START)
                last_sr = 0
                for r in range(BILL_DATA_START, row):
                    sv = bill_ws.cell(r, 1).value
                    if sv and isinstance(sv, (int, float)): last_sr = max(last_sr, int(sv))
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
                if raw and bno: bill_map[clean_job(raw)].append((str(bno), bd))

            upd(88, "Writing consolidated report…", "Generating output")
            added = 0; skipped = 0

            for jn, data in jobs.items():
                if append_mode and jn in existing_jobs: skipped += 1; continue
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
                        main, suf = (fb.rsplit("/",1)[0], "/"+fb.rsplit("/",1)[1]) if "/" in fb else (fb, "")
                        extra = [str(bills[i][0]).split("/")[0][-4:] for i in range(1, len(bills))]
                        merged = main + "/" + "/".join(extra) + suf
                    vd = [b[1] for b in bills if b[1]]
                    bill_ws.cell(row, bill_cols["BILL NO"]).value   = merged
                    bill_ws.cell(row, bill_cols["BILL DATE"]).value = fmt_date(max(vd) if vd else None)
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

            # ── RESULTS SECTION ──────────────────────────────────────────────
            st.markdown(f"""
            <div class="v-results">
              <div class="v-results-inner">
                <div class="reveal sec-kicker">Complete</div>
                <div class="reveal reveal-delay-1 results-title">
                  Report<br><em>Generated</em>
                </div>
                <div class="metrics-row reveal reveal-delay-2">
                  <div class="metric-tile">
                    <div class="mt-n">{added}</div>
                    <div class="mt-l">Entries Added</div>
                  </div>
                  <div class="metric-tile">
                    <div class="mt-n">{skipped}</div>
                    <div class="mt-l">Skipped</div>
                  </div>
                  <div class="metric-tile">
                    <div class="mt-n">{dur:.1f}s</div>
                    <div class="mt-l">Duration</div>
                  </div>
                </div>
              </div>
            </div>
            """, unsafe_allow_html=True)

            dl1, dl2, _ = st.columns([1, 1, 2], gap="medium")
            # wrap downloads in the results light area
            st.markdown('<div style="background:var(--paper-lt);padding:0 2.5rem 4rem;border-top:none;">', unsafe_allow_html=True)
            st.markdown('<div style="max-width:1160px;margin:0 auto;display:flex;gap:1rem;">', unsafe_allow_html=True)
            with dl1:
                with open(out, "rb") as f:
                    st.download_button("↓ Download Excel", data=f, file_name=out,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True, key="dl_xl")
            with dl2:
                try:
                    df  = pd.read_excel(out, sheet_name=month, header=BILL_HEADER_ROW - 1)
                    csv = df.to_csv(index=False)
                    st.download_button("↓ Download CSV", data=csv,
                        file_name=f"UPDATED_BILL_REPORT_{month}.csv",
                        mime="text/csv", use_container_width=True, key="dl_csv")
                except: pass
            st.markdown('</div></div>', unsafe_allow_html=True)

        except Exception as e:
            pb.empty(); se.empty(); de.empty()
            st.session_state.processing_history.append({
                "time": datetime.now().strftime("%H:%M:%S"), "month": month,
                "status": "error", "added": 0, "skipped": 0, "duration": "—",
            })
            st.error(f"Processing error: {str(e)}")
            with st.expander("View full error details"):
                st.exception(e)

# ══════════════════════════════════════════════════════════════════════════════
# FOOTER
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div style="background:var(--ink);padding:2rem 2.5rem;margin-top:0;">
  <div style="max-width:1160px;margin:0 auto;
              display:flex;align-items:center;justify-content:space-between;
              flex-wrap:wrap;gap:1rem;">
    <div style="font-family:var(--ff-display);font-size:1.3rem;font-weight:700;
                color:var(--paper-dk);letter-spacing:-.01em;">
      BillFlow <span style="font-family:var(--ff-hand);font-size:1rem;
                            color:var(--rust-lt);font-weight:600;">Studio</span>
    </div>
    <div style="font-family:var(--ff-body);font-size:.68rem;font-weight:500;
                letter-spacing:.18em;text-transform:uppercase;
                color:rgba(244,236,214,.25);">
      Report Automation Suite &nbsp;·&nbsp; All rights reserved
    </div>
  </div>
</div>
""", unsafe_allow_html=True)