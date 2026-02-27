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

for k, v in {
    'selected_month': max(0, datetime.now().month - 2),
    'processing_history': [],
    'show_preview': False,
    'dark_mode': False,
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

dm = st.session_state.dark_mode

if dm:
    BG="#141008"; BG2="#1E1508"; BG3="#2A1E0C"
    BORDER="rgba(255,220,140,.10)"; BORDER2="rgba(255,220,140,.18)"
    INK="#F0E4C4"; INK2="#C8A870"; INK3="#8A7050"
    RUST="#D4602A"; RUST_LT="#E8804A"; RUST_PL="rgba(212,96,42,.18)"
    SAGE="#4A8A5A"; SAGE_PL="rgba(74,138,90,.18)"
    GOLD="#C89828"; GOLD_PL="rgba(200,152,40,.18)"
    UP_BG="rgba(255,220,140,.05)"; UP_HBRD="#D4602A"
    UP_TXT="rgba(240,228,196,.35)"; UP_BTN="rgba(255,220,140,.08)"; UP_BTXT="rgba(240,228,196,.55)"
    SEL_BG="rgba(255,220,140,.06)"; SEL_BRD="rgba(255,220,140,.15)"; SEL_TXT="#F0E4C4"
    RAD_BG="rgba(255,220,140,.06)"; RAD_BRD="rgba(255,220,140,.14)"
    RAD_TXT="rgba(240,228,196,.45)"; RAD_SEL="#D4602A"; RAD_STX="#fff"
    CHK_TXT="rgba(240,228,196,.55)"
    BTN_BG="rgba(255,220,140,.07)"; BTN_BRD="rgba(255,220,140,.18)"; BTN_TXT="rgba(240,228,196,.55)"
    EXEC_BG="#D4602A"; EXEC_BD="#D4602A"; EXEC_TX="#fff"
    DL_BG="#4A8A5A"; DL_BD="#4A8A5A"
    GREEN_STRIPE="#4A8A5A"; RED_STRIPE="#D4602A"; GOLD_STRIPE="#C89828"
    PROG_TRK="rgba(255,220,140,.1)"; PROG_TXT="#D4602A"
    TOGGLE_IC="☀"; TOGGLE_LBL="Light Mode"
    BAND_BG="#F0E4C4"; BAND_TXT="#2A1E0C"; BAND_ACC="#D4602A"
    FOOT_BG="#0C0804"; FOOT_TXT="rgba(240,228,196,.22)"; FOOT_LOGO="#C8A870"
    MODE_OK_C="#7ED49A"; MODE_OK_BG="rgba(74,138,90,.15)"; MODE_OK_BD="#4A8A5A"
    MODE_WN_C="#E8A888"; MODE_WN_BG="rgba(212,96,42,.15)"; MODE_WN_BD="#D4602A"
    FS_WAIT="rgba(240,228,196,.25)"
    TILE_NUM_C="rgba(255,220,140,.07)"; TILE_TIT_C="#F0E4C4"
    TILE_DSC_C="#8A7050"; TILE_KCK_C="#8A7050"
    HIST_TIT_C="#F0E4C4"; HIST_TM_C="#F0E4C4"
    HIST_DT_C="#8A7050"; HIST_NO_C="#6A5038"
    EXP_BG="rgba(255,220,140,.05)"; EXP_TXT="rgba(240,228,196,.45)"
    ERR_COLOR="#E8A888"
else:
    BG="#F4ECD6"; BG2="#FAF5E4"; BG3="#EDE0C2"
    BORDER="rgba(30,15,2,.12)"; BORDER2="rgba(30,15,2,.22)"
    INK="#1E0F02"; INK2="#5C3418"; INK3="#A07040"
    RUST="#A63418"; RUST_LT="#C4501E"; RUST_PL="rgba(166,52,24,.08)"
    SAGE="#3A6B4A"; SAGE_PL="rgba(58,107,74,.08)"
    GOLD="#8A6A18"; GOLD_PL="rgba(138,106,24,.08)"
    UP_BG="#FAF5E4"; UP_HBRD="#A63418"
    UP_TXT="rgba(30,15,2,.38)"; UP_BTN="#EDE0C2"; UP_BTXT="#5C3418"
    SEL_BG="#FAF5E4"; SEL_BRD="rgba(30,15,2,.20)"; SEL_TXT="#1E0F02"
    RAD_BG="#EDE0C2"; RAD_BRD="rgba(30,15,2,.15)"
    RAD_TXT="#8A7050"; RAD_SEL="#1E0F02"; RAD_STX="#F4ECD6"
    CHK_TXT="#5C3418"
    BTN_BG="#EDE0C2"; BTN_BRD="rgba(30,15,2,.20)"; BTN_TXT="#5C3418"
    EXEC_BG="#1E0F02"; EXEC_BD="#1E0F02"; EXEC_TX="#F4ECD6"
    DL_BG="#3A6B4A"; DL_BD="#3A6B4A"
    GREEN_STRIPE="#3A6B4A"; RED_STRIPE="#A63418"; GOLD_STRIPE="#8A6A18"
    PROG_TRK="rgba(30,15,2,.10)"; PROG_TXT="#A63418"
    TOGGLE_IC="◑"; TOGGLE_LBL="Dark Mode"
    BAND_BG="#1E0F02"; BAND_TXT="#EDE0C2"; BAND_ACC="#A63418"
    FOOT_BG="#1E0F02"; FOOT_TXT="rgba(244,236,214,.22)"; FOOT_LOGO="#C8A870"
    MODE_OK_C="#2E5A3C"; MODE_OK_BG="rgba(58,107,74,.10)"; MODE_OK_BD="#3A6B4A"
    MODE_WN_C="#8A2A12"; MODE_WN_BG="rgba(166,52,24,.10)"; MODE_WN_BD="#A63418"
    FS_WAIT="rgba(30,15,2,.30)"
    TILE_NUM_C="rgba(30,15,2,.06)"; TILE_TIT_C="#1E0F02"
    TILE_DSC_C="#A07040"; TILE_KCK_C="#A07040"
    HIST_TIT_C="#1E0F02"; HIST_TM_C="#1E0F02"
    HIST_DT_C="#A07040"; HIST_NO_C="#A07040"
    EXP_BG="#EDE0C2"; EXP_TXT="#5C3418"
    ERR_COLOR="#8A2A12"

CSS = f"""
@import url('https://fonts.googleapis.com/css2?family=Cormorant:ital,wght@0,300;0,400;0,600;0,700;1,300;1,400;1,700&family=Libre+Baskerville:ital,wght@0,400;0,700;1,400&family=Jost:wght@300;400;500;600;700&family=Caveat:wght@500;600;700&display=swap');

header[data-testid="stHeader"],
[data-testid="stHeader"],
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"],
#MainMenu, footer, .stDeployButton {{
    display: none !important; height: 0 !important; min-height: 0 !important;
}}
*, *::before, *::after {{ box-sizing: border-box; }}
html {{ scroll-behavior:smooth; scrollbar-gutter:stable; }}
::-webkit-scrollbar {{ width:5px; }}
::-webkit-scrollbar-track {{ background:{BG3}; }}
::-webkit-scrollbar-thumb {{ background:{BORDER2}; border-radius:3px; }}
::-webkit-scrollbar-thumb:hover {{ background:{INK3}; }}
body,
[data-testid="stAppViewContainer"],
.main,
[data-testid="stAppViewBlockContainer"],
[data-testid="stMainBlockContainer"] {{
    background:{BG} !important; color:{INK} !important;
    font-family:'Jost',sans-serif !important;
}}
.main .block-container {{ padding:0 !important; max-width:100% !important; }}
[data-testid="stHorizontalBlock"] {{ gap:1.25rem !important; padding:0 !important; align-items:stretch !important; }}
[data-testid="stColumn"] {{ padding:0 !important; }}
[data-testid="stVerticalBlock"] {{ gap:0 !important; }}

/* NAV */
.bf-nav {{
    background:{BG}; border-bottom:1px solid {BORDER};
    padding:0 2.5rem; position:sticky; top:0; z-index:999;
}}
.bf-nav-inner {{
    max-width:1200px; margin:0 auto;
    height:60px; display:flex; align-items:center; justify-content:space-between;
}}
.nav-logo {{ font-family:'Cormorant',serif; font-size:1.6rem; font-weight:700; color:{INK}; display:flex; align-items:baseline; gap:.3rem; }}
.nav-logo-script {{ font-family:'Caveat',cursive; font-size:1rem; font-weight:600; color:{RUST}; }}
.nav-link {{ font-family:'Jost',sans-serif; font-size:.75rem; font-weight:600; letter-spacing:.12em; text-transform:uppercase; color:{INK3}; cursor:pointer; transition:color .15s; }}
.nav-link:hover {{ color:{INK}; }}
.nav-cta {{ font-family:'Jost',sans-serif; font-size:.72rem; font-weight:700; letter-spacing:.14em; text-transform:uppercase; color:{BG2}; background:{INK}; padding:.45rem 1.3rem; cursor:pointer; transition:background .2s; }}
.nav-cta:hover {{ background:{RUST}; }}

/* HERO */
.bf-hero {{ background:{BG}; padding:6rem 2.5rem 5rem; text-align:center; border-bottom:1px solid {BORDER}; position:relative; overflow:hidden; }}
.bf-hero::before {{ content:''; position:absolute; top:1.5rem; left:6%; right:6%; height:1px; background:linear-gradient(90deg,transparent,{BORDER2} 25%,{BORDER2} 75%,transparent); }}
.bf-hero::after  {{ content:''; position:absolute; top:1.85rem; left:6%; right:6%; height:1px; background:linear-gradient(90deg,transparent,{BORDER} 25%,{BORDER} 75%,transparent); }}
.hero-vol {{ font-family:'Jost',sans-serif; font-size:.65rem; font-weight:700; letter-spacing:.3em; text-transform:uppercase; color:{INK3}; margin-bottom:1.25rem; animation:fadeUp .5s .05s ease both; }}
.hero-vol em {{ color:{RUST}; font-style:normal; }}
.hero-h1 {{ font-family:'Cormorant',serif; font-size:clamp(4rem,9vw,8.5rem); font-weight:700; color:{INK}; line-height:.92; letter-spacing:-.025em; margin-bottom:.3rem; animation:fadeUp .55s .1s ease both; }}
.hero-h1 em {{ font-style:italic; font-weight:300; color:{INK2}; }}
.hero-script {{ font-family:'Caveat',cursive; font-size:clamp(2rem,4.5vw,3.8rem); font-weight:600; color:{RUST}; display:block; line-height:1.1; margin-bottom:1.5rem; animation:fadeUp .55s .15s ease both; }}
.hero-divider {{ display:flex; align-items:center; gap:1rem; justify-content:center; margin-bottom:1.5rem; animation:fadeUp .55s .2s ease both; }}
.hero-divider-line {{ width:60px; height:1px; background:{BORDER2}; }}
.hero-divider-dot {{ color:{RUST}; font-size:.7rem; }}
.hero-body {{ font-family:'Libre Baskerville',serif; font-style:italic; font-size:clamp(.9rem,1.5vw,1.05rem); color:{INK3}; max-width:480px; margin:0 auto 2.25rem; line-height:1.7; animation:fadeUp .6s .25s ease both; }}
.hero-btns {{ display:flex; align-items:center; gap:1rem; justify-content:center; animation:fadeUp .6s .3s ease both; }}
.hero-btn-p {{ font-family:'Jost',sans-serif; font-size:.78rem; font-weight:700; letter-spacing:.12em; text-transform:uppercase; color:{BG2}; background:{INK}; padding:.88rem 2.25rem; border:2px solid {INK}; cursor:pointer; transition:all .2s; display:inline-block; }}
.hero-btn-p:hover {{ background:{RUST}; border-color:{RUST}; }}
.hero-btn-g {{ font-family:'Jost',sans-serif; font-size:.75rem; font-weight:600; letter-spacing:.1em; text-transform:uppercase; color:{INK2}; background:transparent; padding:.88rem 1.8rem; border:1px solid {BORDER2}; cursor:pointer; transition:all .2s; display:inline-block; }}
.hero-btn-g:hover {{ border-color:{INK}; color:{INK}; }}
.hero-scroll-hint {{ margin-top:3.5rem; display:flex; flex-direction:column; align-items:center; gap:.45rem; animation:fadeUp .7s .5s ease both; }}
.hsh-label {{ font-family:'Jost',sans-serif; font-size:.6rem; font-weight:700; letter-spacing:.25em; text-transform:uppercase; color:{INK3}; }}
.hsh-line {{ width:1px; height:36px; background:linear-gradient(180deg,{BORDER2},transparent); animation:scrollPulse 2.2s ease-in-out infinite; }}

/* BAND */
.bf-band {{ background:{BAND_BG}; padding:.8rem 0; overflow:hidden; white-space:nowrap; }}
.band-track {{ display:inline-flex; animation:marquee 28s linear infinite; }}
.band-item {{ font-family:'Cormorant',serif; font-size:1rem; font-weight:700; letter-spacing:.08em; text-transform:uppercase; color:{BAND_TXT}; padding:0 1.75rem; }}
.band-item em {{ color:{BAND_ACC}; font-style:normal; }}

/* SECTIONS */
.bf-sec {{ background:{BG}; padding:4.5rem 2.5rem 2rem; border-bottom:1px solid {BORDER}; }}
.bf-sec-inner {{ max-width:1200px; margin:0 auto; }}
.bf-sec-dark {{ background:{BG2}; padding:4.5rem 2.5rem 2rem; border-bottom:1px solid {BORDER}; }}
.bf-sec-dark-inner {{ max-width:1200px; margin:0 auto; }}
.sec-eyebrow {{ font-family:'Jost',sans-serif; font-size:.62rem; font-weight:700; letter-spacing:.28em; text-transform:uppercase; color:{RUST}; display:flex; align-items:center; gap:.65rem; margin-bottom:.85rem; }}
.sec-eyebrow::before {{ content:''; width:28px; height:1px; background:{RUST}; flex-shrink:0; }}
.sec-headline {{ font-family:'Cormorant',serif; font-size:clamp(2.2rem,4.5vw,4rem); font-weight:700; color:{INK}; line-height:.94; letter-spacing:-.025em; margin-bottom:.6rem; }}
.sec-headline em {{ font-style:italic; font-weight:300; color:{INK2}; }}
.sec-sub {{ font-family:'Libre Baskerville',serif; font-style:italic; font-size:.95rem; color:{INK3}; line-height:1.65; max-width:520px; margin-bottom:2.5rem; }}

/* TILES */
.tile-wrap {{ background:{BG2}; border:1px solid {BORDER}; border-top:4px solid; padding:1.5rem 1.5rem 1.25rem; margin-bottom:.75rem; transition:box-shadow .2s; position:relative; overflow:hidden; }}
.tile-wrap:hover {{ box-shadow:0 8px 32px rgba(0,0,0,.12); }}
.tile-num-bg {{ position:absolute; right:.8rem; top:.3rem; font-family:'Cormorant',serif; font-size:5.5rem; font-weight:700; color:{TILE_NUM_C}; line-height:1; letter-spacing:-.04em; pointer-events:none; user-select:none; }}
.tile-kicker {{ font-family:'Jost',sans-serif; font-size:.6rem; font-weight:700; letter-spacing:.2em; text-transform:uppercase; color:{TILE_KCK_C}; margin-bottom:.55rem; display:block; }}
.tile-title {{ font-family:'Cormorant',serif; font-size:1.55rem; font-weight:700; color:{TILE_TIT_C}; line-height:1.1; margin-bottom:.35rem; letter-spacing:-.01em; }}
.tile-desc {{ font-family:'Libre Baskerville',serif; font-style:italic; font-size:.8rem; color:{TILE_DSC_C}; line-height:1.55; }}
.tile-rule {{ height:1px; background:{BORDER}; margin:1rem 0 .75rem; }}
.tile-status-ok {{ display:inline-flex; align-items:center; gap:.4rem; font-family:'Jost',sans-serif; font-size:.65rem; font-weight:700; letter-spacing:.12em; text-transform:uppercase; padding:.28rem .75rem; background:{SAGE_PL}; border:1px solid {SAGE}; color:{SAGE}; }}
.tile-status-wait {{ display:inline-flex; align-items:center; gap:.4rem; font-family:'Jost',sans-serif; font-size:.65rem; font-weight:700; letter-spacing:.12em; text-transform:uppercase; padding:.28rem .75rem; background:transparent; border:1px solid {BORDER2}; color:{INK3}; }}
.tile-preview {{ background:{BG3}; padding:.6rem .8rem; margin-top:.55rem; font-family:'Courier New',monospace; font-size:.68rem; }}
.tp-row {{ display:flex; justify-content:space-between; color:{INK2}; padding:.14rem 0; border-bottom:1px solid {BORDER}; }}
.tp-row:last-child {{ border-bottom:none; }}
.tp-val {{ font-weight:700; color:{RUST}; }}

/* FILE UPLOADER */
[data-testid="stFileUploader"] section {{ background:{UP_BG} !important; border:1.5px dashed {BORDER2} !important; border-radius:0 !important; padding:.7rem 1rem !important; transition:border-color .2s !important; }}
[data-testid="stFileUploader"] section:hover {{ border-color:{UP_HBRD} !important; background:{BG3} !important; }}
[data-testid="stFileUploader"] section p,
[data-testid="stFileUploader"] section span,
[data-testid="stFileUploader"] section small {{ font-family:'Jost',sans-serif !important; font-size:.78rem !important; color:{UP_TXT} !important; }}
[data-testid="stFileUploader"] section button {{ background:{UP_BTN} !important; border:1px solid {BORDER2} !important; color:{UP_BTXT} !important; font-family:'Jost',sans-serif !important; font-size:.75rem !important; border-radius:0 !important; }}

/* CONFIG */
.cfg-label {{ font-family:'Jost',sans-serif; font-size:.62rem; font-weight:700; letter-spacing:.2em; text-transform:uppercase; color:{INK3}; display:block; margin-bottom:.5rem; }}
[data-testid="stSelectbox"] > div > div {{ background:{SEL_BG} !important; border:1px solid {SEL_BRD} !important; border-radius:0 !important; color:{SEL_TXT} !important; font-family:'Courier New',monospace !important; font-size:.92rem !important; font-weight:600 !important; }}
[data-testid="stSelectbox"] > div > div:hover {{ border-color:{RUST} !important; }}
[data-testid="stSelectbox"] > div > div > div {{ color:{SEL_TXT} !important; }}
[data-testid="stSelectbox"] svg {{ fill:{INK3} !important; }}
[data-testid="stRadio"] > div {{ display:flex !important; gap:0 !important; flex-direction:row !important; background:{RAD_BG} !important; border:1px solid {RAD_BRD} !important; border-radius:0 !important; padding:3px !important; }}
[data-testid="stRadio"] label > div:first-child,
[data-testid="stRadio"] input[type="radio"] {{ display:none !important; }}
[data-testid="stRadio"] label {{ background:transparent !important; border:none !important; padding:.6rem 1.5rem !important; font-family:'Jost',sans-serif !important; font-size:.85rem !important; font-weight:500 !important; color:{RAD_TXT} !important; cursor:pointer !important; transition:all .18s !important; flex:1 !important; text-align:center !important; letter-spacing:.04em !important; line-height:1.4 !important; }}
[data-testid="stRadio"] label:hover {{ color:{INK} !important; background:{BG3} !important; }}
[data-testid="stRadio"] label:has(input:checked) {{ background:{RAD_SEL} !important; color:{RAD_STX} !important; font-weight:700 !important; letter-spacing:.06em !important; box-shadow:0 2px 10px rgba(0,0,0,.2) !important; }}
.mode-note {{ font-family:'Libre Baskerville',serif; font-style:italic; font-size:.85rem; line-height:1.6; padding:.7rem 1rem; margin-top:.65rem; border-left:3px solid; }}
.mode-ok {{ color:{MODE_OK_C}; background:{MODE_OK_BG}; border-color:{MODE_OK_BD}; }}
.mode-warn {{ color:{MODE_WN_C}; background:{MODE_WN_BG}; border-color:{MODE_WN_BD}; }}
[data-testid="stCheckbox"] label span,
[data-testid="stCheckbox"] span {{ font-family:'Jost',sans-serif !important; font-size:.83rem !important; color:{CHK_TXT} !important; }}
.fs-row {{ display:flex; align-items:center; gap:2rem; padding:.7rem 1rem; background:{BG3}; border:1px solid {BORDER}; margin-bottom:1.25rem; }}
.fs-item {{ font-family:'Jost',sans-serif; font-size:.72rem; font-weight:700; letter-spacing:.08em; text-transform:uppercase; display:flex; align-items:center; gap:.4rem; }}

/* BUTTONS */
.stButton > button {{ background:{BTN_BG} !important; border:1px solid {BTN_BRD} !important; color:{BTN_TXT} !important; font-family:'Jost',sans-serif !important; font-size:.74rem !important; font-weight:600 !important; letter-spacing:.1em !important; text-transform:uppercase !important; border-radius:0 !important; padding:.6rem 1.5rem !important; transition:all .18s !important; width:100% !important; }}
.stButton > button:hover {{ background:{BG3} !important; border-color:{INK3} !important; color:{INK} !important; }}
.exec-zone .stButton > button {{ background:{EXEC_BG} !important; border:2px solid {EXEC_BD} !important; color:{EXEC_TX} !important; font-size:.84rem !important; font-weight:700 !important; padding:1.05rem 2rem !important; letter-spacing:.14em !important; box-shadow:0 4px 20px rgba(0,0,0,.18) !important; }}
.exec-zone .stButton > button:hover {{ background:{RUST_LT} !important; border-color:{RUST_LT} !important; transform:translateY(-2px) !important; box-shadow:0 8px 32px rgba(0,0,0,.25) !important; }}
[data-testid="stProgress"] > div {{ background:{PROG_TRK} !important; border-radius:0 !important; height:3px !important; }}
[data-testid="stProgress"] > div > div {{ background:linear-gradient(90deg,{RUST},{RUST_LT}) !important; border-radius:0 !important; }}
.prog-txt {{ font-family:'Courier New',monospace; font-size:.75rem; color:{PROG_TXT}; text-align:center; padding:.35rem 0; letter-spacing:.06em; }}
.prog-det {{ display:flex; justify-content:space-between; font-family:'Courier New',monospace; font-size:.64rem; color:{INK3}; margin-top:.15rem; }}
.stDownloadButton > button {{ background:{DL_BG} !important; border:2px solid {DL_BD} !important; color:#fff !important; font-family:'Jost',sans-serif !important; font-size:.74rem !important; font-weight:700 !important; letter-spacing:.1em !important; text-transform:uppercase !important; border-radius:0 !important; padding:.75rem 1.5rem !important; width:100% !important; transition:all .18s !important; }}
.stDownloadButton > button:hover {{ filter:brightness(1.12) !important; transform:translateY(-1px) !important; }}
.streamlit-expanderHeader {{ font-family:'Jost',sans-serif !important; font-size:.8rem !important; color:{EXP_TXT} !important; background:{EXP_BG} !important; border-radius:0 !important; }}
[data-testid="stAlert"] {{ font-family:'Jost',sans-serif !important; font-size:.85rem !important; border-radius:0 !important; }}

/* HISTORY */
.hist-panel {{ background:{BG3}; border:1px solid {BORDER}; padding:1.5rem 1.4rem; }}
.hist-panel-title {{ font-family:'Cormorant',serif; font-size:1.45rem; font-weight:700; color:{HIST_TIT_C}; letter-spacing:-.01em; margin-bottom:1.1rem; display:flex; align-items:center; gap:.6rem; }}
.hist-panel-title::after {{ content:''; flex:1; height:1px; background:{BORDER}; }}
.hist-entry {{ display:flex; align-items:center; justify-content:space-between; padding:.55rem 0; border-bottom:1px solid {BORDER}; }}
.hist-entry:last-child {{ border-bottom:none; }}
.he-info {{ display:flex; flex-direction:column; gap:2px; }}
.he-time {{ font-family:'Courier New',monospace; font-size:.72rem; font-weight:700; color:{HIST_TM_C}; }}
.he-detail {{ font-family:'Jost',sans-serif; font-size:.68rem; color:{HIST_DT_C}; }}
.he-badge {{ font-family:'Jost',sans-serif; font-size:.6rem; font-weight:700; letter-spacing:.12em; text-transform:uppercase; padding:.2rem .6rem; }}
.he-ok  {{ color:{SAGE};  background:{SAGE_PL};  border:1px solid {SAGE}; }}
.he-err {{ color:{RUST};  background:{RUST_PL};  border:1px solid {RUST}; }}
.hist-empty {{ font-family:'Libre Baskerville',serif; font-style:italic; font-size:.85rem; color:{HIST_NO_C}; line-height:1.6; }}

/* RESULTS */
.results-section {{ background:{BG2}; border-top:1px solid {BORDER}; padding:4rem 2.5rem; }}
.results-inner {{ max-width:1200px; margin:0 auto; }}
.results-headline {{ font-family:'Cormorant',serif; font-size:clamp(3.5rem,7vw,6.5rem); font-weight:700; color:{INK}; line-height:.88; letter-spacing:-.03em; margin-bottom:2.5rem; }}
.results-headline em {{ font-style:italic; font-weight:300; color:{RUST}; }}
.metrics-row {{ display:grid; grid-template-columns:repeat(3,1fr); gap:1.25rem; margin-bottom:2.5rem; }}
.metric-tile {{ background:{BG}; border:1px solid {BORDER}; padding:1.6rem 1.4rem; position:relative; overflow:hidden; }}
.metric-tile::after {{ content:''; position:absolute; bottom:0; left:0; right:0; height:3px; background:{RUST}; }}
.mt-num {{ font-family:'Cormorant',serif; font-size:4.5rem; font-weight:700; color:{RUST}; line-height:.88; letter-spacing:-.03em; margin-bottom:.45rem; }}
.mt-lbl {{ font-family:'Jost',sans-serif; font-size:.62rem; font-weight:700; letter-spacing:.2em; text-transform:uppercase; color:{INK3}; }}

/* FOOTER */
.bf-footer {{ background:{FOOT_BG}; padding:1.75rem 2.5rem; }}
.bf-footer-inner {{ max-width:1200px; margin:0 auto; display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:1rem; }}
.footer-logo {{ font-family:'Cormorant',serif; font-size:1.3rem; font-weight:700; color:{FOOT_LOGO}; display:flex; align-items:baseline; gap:.3rem; }}
.footer-script {{ font-family:'Caveat',cursive; font-size:.9rem; font-weight:600; color:{RUST_LT}; }}
.footer-copy {{ font-family:'Jost',sans-serif; font-size:.65rem; font-weight:600; letter-spacing:.18em; text-transform:uppercase; color:{FOOT_TXT}; }}

@keyframes fadeUp {{ from{{ opacity:0; transform:translateY(16px); }} to{{ opacity:1; transform:translateY(0); }} }}
@keyframes scrollPulse {{ 0%,100%{{ opacity:.25; transform:scaleY(1); }} 50%{{ opacity:.9; transform:scaleY(1.2); }} }}
@keyframes marquee {{ from{{ transform:translateX(0); }} to{{ transform:translateX(-50%); }} }}

@media(max-width:900px){{
    .bf-sec,.bf-sec-dark,.bf-hero,.results-section{{ padding-left:1.25rem; padding-right:1.25rem; }}
    .hero-h1{{ font-size:3.2rem; }}
    .metrics-row{{ grid-template-columns:1fr 1fr; }}
}}
@media(max-width:580px){{
    .hero-h1{{ font-size:2.4rem; }}
    .hero-btns{{ flex-direction:column; }}
    .metrics-row{{ grid-template-columns:1fr; }}
}}
"""

st.markdown(f"<style>{CSS}</style>", unsafe_allow_html=True)

# NAV
nav_l, nav_m, nav_r = st.columns([3,2,3])
with nav_l:
    st.markdown(f'<div style="background:{BG};border-bottom:1px solid {BORDER};padding:0 0 0 2.5rem;height:60px;display:flex;align-items:center;gap:2.5rem;"><span class="nav-link">Upload</span><span class="nav-link">Configure</span><span class="nav-link">History</span></div>', unsafe_allow_html=True)
with nav_m:
    st.markdown(f'<div style="background:{BG};border-bottom:1px solid {BORDER};border-left:1px solid {BORDER};border-right:1px solid {BORDER};height:60px;display:flex;align-items:center;justify-content:center;"><span class="nav-logo">BillFlow <span class="nav-logo-script">Studio</span></span></div>', unsafe_allow_html=True)
with nav_r:
    nr1, nr2 = st.columns([2,2])
    with nr1:
        st.markdown(f'<div style="background:{BG};border-bottom:1px solid {BORDER};height:60px;display:flex;align-items:center;justify-content:flex-end;padding-right:.75rem;"><span class="nav-cta">Process Now</span></div>', unsafe_allow_html=True)
    with nr2:
        st.markdown(f'<div style="background:{BG};border-bottom:1px solid {BORDER};height:60px;display:flex;align-items:center;padding-right:2.5rem;">', unsafe_allow_html=True)
        if st.button(f"{TOGGLE_IC}  {TOGGLE_LBL}", key="theme_toggle"):
            st.session_state.dark_mode = not st.session_state.dark_mode
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

# HERO
st.markdown(f"""
<div class="bf-hero">
    <div class="hero-vol">Vol. I &nbsp;·&nbsp; <em>Report Automation</em> &nbsp;·&nbsp; Est. 2024</div>
    <h1 class="hero-h1">AUTOMATE<br><em>Your Billing</em></h1>
    <span class="hero-script">Reports, reconciled.</span>
    <div class="hero-divider"><div class="hero-divider-line"></div><span class="hero-divider-dot">✦</span><div class="hero-divider-line"></div></div>
    <p class="hero-body">Upload three source files, configure the billing month and update mode, then let BillFlow consolidate every job and invoice into a clean Excel report — in seconds.</p>
    <div class="hero-btns">
        <span class="hero-btn-p">Begin Upload ↓</span>
        <span class="hero-btn-g">How it works</span>
    </div>
    <div class="hero-scroll-hint">
        <span class="hsh-label">Scroll to begin</span>
        <div class="hsh-line"></div>
    </div>
</div>
""", unsafe_allow_html=True)

# BAND
items = ["Bill Report","Job Register","E-Invoice","Reconcile","Export","Automate"] * 4
band_html = "".join([f'<span class="band-item">{t} &nbsp;<em>✦</em></span>' for t in items])
st.markdown(f'<div class="bf-band"><div class="band-track">{band_html}{band_html}</div></div>', unsafe_allow_html=True)

# SECTION 01
st.markdown(f'<div class="bf-sec"><div class="bf-sec-inner"><div class="sec-eyebrow">Step One</div><div class="sec-headline">Upload<br><em>Your Documents</em></div><p class="sec-sub">Three source files power the reconciliation — the master ledger, the job register, and the invoice index. All files must be .xlsx format.</p></div></div>', unsafe_allow_html=True)

st.markdown(f'<div style="background:{BG};padding:0 2.5rem 3rem;"><div style="max-width:1200px;margin:0 auto;">', unsafe_allow_html=True)
uc1, uc2, uc3 = st.columns(3, gap="medium")

with uc1:
    st.markdown(f'<div class="tile-wrap" style="border-top-color:{GREEN_STRIPE};"><div class="tile-num-bg">I</div><span class="tile-kicker">Primary Register</span><div class="tile-title">Bill Report</div><div class="tile-desc">Master billing workbook with monthly tabs — consolidated data is written into this file.</div><div class="tile-rule"></div></div>', unsafe_allow_html=True)
    bill_file = st.file_uploader("Bill Report", type="xlsx", key="bill", label_visibility="collapsed")
    if bill_file:
        st.markdown('<div class="tile-status-ok">✓ &nbsp;Loaded</div>', unsafe_allow_html=True)
        if st.session_state.show_preview:
            v = validate_file(bill_file, "bill")
            if v["valid"]:
                sn = len(v["sheets"]) if v["sheets"] else "—"
                st.markdown(f'<div class="tile-preview"><div class="tp-row"><span>Rows</span><span class="tp-val">{v["rows"]}</span></div><div class="tp-row"><span>Cols</span><span class="tp-val">{v["cols"]}</span></div><div class="tp-row"><span>Sheets</span><span class="tp-val">{sn}</span></div></div>', unsafe_allow_html=True)
            bill_file.seek(0)
    else:
        st.markdown('<div class="tile-status-wait">○ &nbsp;Awaiting file</div>', unsafe_allow_html=True)

with uc2:
    st.markdown(f'<div class="tile-wrap" style="border-top-color:{RED_STRIPE};"><div class="tile-num-bg">II</div><span class="tile-kicker">Job Register</span><div class="tile-title">Job Report</div><div class="tile-desc">Shipper details, container numbers, invoice references and SB/BE data for each job entry.</div><div class="tile-rule"></div></div>', unsafe_allow_html=True)
    job_file = st.file_uploader("Job Report", type="xlsx", key="job", label_visibility="collapsed")
    if job_file:
        st.markdown('<div class="tile-status-ok">✓ &nbsp;Loaded</div>', unsafe_allow_html=True)
        if st.session_state.show_preview:
            v = validate_file(job_file, "job")
            if v["valid"]:
                st.markdown(f'<div class="tile-preview"><div class="tp-row"><span>Rows</span><span class="tp-val">{v["rows"]}</span></div><div class="tp-row"><span>Cols</span><span class="tp-val">{v["cols"]}</span></div></div>', unsafe_allow_html=True)
            job_file.seek(0)
    else:
        st.markdown('<div class="tile-status-wait">○ &nbsp;Awaiting file</div>', unsafe_allow_html=True)

with uc3:
    st.markdown(f'<div class="tile-wrap" style="border-top-color:{GOLD_STRIPE};"><div class="tile-num-bg">III</div><span class="tile-kicker">Invoice Data</span><div class="tile-title">E-Invoice Report</div><div class="tile-desc">Electronic invoice records — bill numbers and issuance dates matched per job reference.</div><div class="tile-rule"></div></div>', unsafe_allow_html=True)
    einv_file = st.file_uploader("E-Invoice", type="xlsx", key="einv", label_visibility="collapsed")
    if einv_file:
        st.markdown('<div class="tile-status-ok">✓ &nbsp;Loaded</div>', unsafe_allow_html=True)
        if st.session_state.show_preview:
            v = validate_file(einv_file, "einv")
            if v["valid"]:
                st.markdown(f'<div class="tile-preview"><div class="tp-row"><span>Rows</span><span class="tp-val">{v["rows"]}</span></div><div class="tp-row"><span>Cols</span><span class="tp-val">{v["cols"]}</span></div></div>', unsafe_allow_html=True)
            einv_file.seek(0)
    else:
        st.markdown('<div class="tile-status-wait">○ &nbsp;Awaiting file</div>', unsafe_allow_html=True)

st.markdown('</div></div>', unsafe_allow_html=True)

# SECTION 02
st.markdown(f'<div class="bf-sec-dark"><div class="bf-sec-dark-inner"><div class="sec-eyebrow">Step Two</div><div class="sec-headline">Configure<br><em>&amp; Execute</em></div><p class="sec-sub">Select billing month and update mode, then run the pipeline in a single click.</p></div></div>', unsafe_allow_html=True)

st.markdown(f'<div style="background:{BG2};padding:0 2.5rem 4rem;border-bottom:1px solid {BORDER};"><div style="max-width:1200px;margin:0 auto;">', unsafe_allow_html=True)

b_c = SAGE if bill_file  else FS_WAIT
j_c = SAGE if job_file   else FS_WAIT
e_c = SAGE if einv_file  else FS_WAIT
b_i = "✓" if bill_file  else "○"
j_i = "✓" if job_file   else "○"
e_i = "✓" if einv_file  else "○"
st.markdown(f'<div class="fs-row"><span class="fs-item" style="color:{b_c};">{b_i}&thinsp; Bill Report</span><span class="fs-item" style="color:{j_c};">{j_i}&thinsp; Job Report</span><span class="fs-item" style="color:{e_c};">{e_i}&thinsp; E-Invoice</span></div>', unsafe_allow_html=True)

cfg_main, cfg_side = st.columns([3,2], gap="large")

with cfg_main:
    p1, p2 = st.columns(2, gap="large")
    with p1:
        st.markdown('<span class="cfg-label">Billing Month</span>', unsafe_allow_html=True)
        month = st.selectbox("Month", MONTHS, index=st.session_state.selected_month, label_visibility="collapsed", key="month_sel")
        st.session_state.selected_month = MONTHS.index(month)
    with p2:
        st.markdown('<span class="cfg-label">Update Mode</span>', unsafe_allow_html=True)
        update_mode = st.radio("Mode", ["Append","Overwrite"], label_visibility="collapsed", key="umode")

    if update_mode == "Append":
        st.markdown('<div class="mode-note mode-ok">New job entries will be appended below existing records — nothing is overwritten or deleted.</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="mode-note mode-warn">All existing entries for the selected month will be cleared before writing new data.</div>', unsafe_allow_html=True)

    st.markdown(f'<div style="height:1rem;margin-top:1.25rem;border-top:1px solid {BORDER};padding-top:1.25rem;">', unsafe_allow_html=True)
    sp = st.checkbox("Show file preview after upload", value=st.session_state.show_preview, key="prev_chk")
    st.session_state.show_preview = sp
    st.markdown('<div style="height:.75rem;"></div>', unsafe_allow_html=True)
    st.markdown('<div class="exec-zone">', unsafe_allow_html=True)
    process_clicked = st.button("✦  Execute Report Generation", key="proc_main", use_container_width=True)
    st.markdown('</div></div>', unsafe_allow_html=True)

with cfg_side:
    st.markdown(f'<div class="hist-panel"><div class="hist-panel-title">Activity Log</div>', unsafe_allow_html=True)
    if st.session_state.processing_history:
        for entry in reversed(st.session_state.processing_history[-6:]):
            bc = "he-ok" if entry["status"]=="success" else "he-err"
            bm = "DONE" if entry["status"]=="success" else "ERROR"
            st.markdown(f'<div class="hist-entry"><div class="he-info"><span class="he-time">{entry["time"]} · {entry["month"]}</span><span class="he-detail">+{entry["added"]} added · {entry.get("duration","—")}</span></div><span class="he-badge {bc}">{bm}</span></div>', unsafe_allow_html=True)
        if st.button("Clear log", key="clr_hist"):
            st.session_state.processing_history = []; st.rerun()
    else:
        st.markdown(f'<p class="hist-empty">No runs yet. Processed reports will be logged here with timing and entry counts.</p>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('</div></div>', unsafe_allow_html=True)

# PROCESSING
if process_clicked:
    if not (bill_file and job_file and einv_file):
        st.error("All three documents must be uploaded before running.")
    else:
        st.markdown(f'<div style="background:{BG2};padding:1rem 2.5rem 0;"><div style="max-width:1200px;margin:0 auto;">', unsafe_allow_html=True)
        pb = st.progress(0); se = st.empty(); de = st.empty()
        st.markdown('</div></div>', unsafe_allow_html=True)
        t0 = datetime.now()

        def upd(p, msg, det=""):
            pb.progress(p)
            se.markdown(f'<div style="background:{BG2};padding:0 2.5rem;"><div style="max-width:1200px;margin:0 auto;"><div class="prog-txt">{msg}</div></div></div>', unsafe_allow_html=True)
            if det:
                de.markdown(f'<div style="background:{BG2};padding:0 2.5rem .5rem;"><div style="max-width:1200px;margin:0 auto;"><div class="prog-det"><span>{det}</span><span>{p}%</span></div></div></div>', unsafe_allow_html=True)

        try:
            upd(10,"Loading workbooks…","Reading Excel files")
            bill_wb=load_workbook(bill_file); job_wb=load_workbook(job_file); einv_wb=load_workbook(einv_file)
            upd(22,f"Locating sheet '{month}'…","Checking tabs")
            if month not in bill_wb.sheetnames:
                st.error(f"Sheet '{month}' not found. Available: {', '.join(bill_wb.sheetnames)}"); st.stop()
            bill_ws=bill_wb[month]; job_ws=job_wb.active; einv_ws=einv_wb.active
            upd(35,"Mapping columns…","Analysing headers")
            bill_cols=col_map(bill_ws,BILL_HEADER_ROW); job_cols=col_map(job_ws,SRC_HEADER_ROW); einv_cols=col_map(einv_ws,SRC_HEADER_ROW)
            upd(45,"Validating structure…","Checking required fields")
            for c in ["JOB NO","JOB DATE","SHIPPER NAME","INVOICE NO","SB / BE NO","CONTAINER NO"]:
                if c not in job_cols: st.error(f"Missing in Job Report: '{c}'"); st.stop()
            for c in ["JOB NO","BILL NO","BILL DATE"]:
                if c not in einv_cols: st.error(f"Missing in E-Invoice: '{c}'"); st.stop()
            for c in ["JOB NO","JOB DATE","PARTY NAME","INVOICE NO","SB / BE NO","CONTAINER NO","BILL NO","BILL DATE"]:
                if c not in bill_cols: st.error(f"Missing in Bill Report: '{c}'"); st.stop()
            append_mode=(update_mode=="Append"); existing_jobs=set()
            if append_mode:
                upd(55,"Scanning existing entries…","Detecting duplicates")
                existing_jobs=get_existing_jobs(bill_ws,bill_cols,BILL_DATA_START)
                row=next_empty(bill_ws,BILL_DATA_START)
                last_sr=0
                for r in range(BILL_DATA_START,row):
                    sv=bill_ws.cell(r,1).value
                    if sv and isinstance(sv,(int,float)): last_sr=max(last_sr,int(sv))
                sr=last_sr+1
            else:
                upd(55,"Clearing sheet…","Removing existing entries")
                for r in range(BILL_DATA_START,bill_ws.max_row+1):
                    for c in range(1,bill_ws.max_column+1): bill_ws.cell(r,c).value=None
                row=BILL_DATA_START; sr=1
            upd(65,"Reading job register…","Parsing entries")
            jobs={}
            for r in range(SRC_DATA_START,job_ws.max_row+1):
                raw=job_ws.cell(r,job_cols["JOB NO"]).value
                if not raw: continue
                jn=clean_job(raw)
                if jn in jobs: continue
                jobs[jn]={"JOB DATE":parse_date(job_ws.cell(r,job_cols["JOB DATE"]).value),"PARTY NAME":job_ws.cell(r,job_cols["SHIPPER NAME"]).value,"INVOICE NO":job_ws.cell(r,job_cols["INVOICE NO"]).value,"SB / BE NO":job_ws.cell(r,job_cols["SB / BE NO"]).value,"CONTAINER NO":job_ws.cell(r,job_cols["CONTAINER NO"]).value}
            upd(78,"Reading invoice register…","Parsing bills")
            bill_map=defaultdict(list)
            for r in range(SRC_DATA_START,einv_ws.max_row+1):
                raw=einv_ws.cell(r,einv_cols["JOB NO"]).value; bno=einv_ws.cell(r,einv_cols["BILL NO"]).value; bd=parse_date(einv_ws.cell(r,einv_cols["BILL DATE"]).value)
                if raw and bno: bill_map[clean_job(raw)].append((str(bno),bd))
            upd(88,"Writing report…","Generating output")
            added=0; skipped=0
            for jn,data in jobs.items():
                if append_mode and jn in existing_jobs: skipped+=1; continue
                bill_ws.cell(row,1).value=sr; bill_ws.cell(row,bill_cols["JOB NO"]).value=jn
                for field in ["JOB DATE","PARTY NAME","INVOICE NO","SB / BE NO","CONTAINER NO"]:
                    val=data[field]
                    if field=="JOB DATE": val=fmt_date(val)
                    bill_ws.cell(row,bill_cols[field]).value=val
                if jn in bill_map:
                    bills=bill_map[jn]
                    if len(bills)==1: merged=bills[0][0]
                    else:
                        fb=str(bills[0][0]); main,suf=(fb.rsplit("/",1)[0],"/"+fb.rsplit("/",1)[1]) if "/" in fb else (fb,"")
                        extra=[str(bills[i][0]).split("/")[0][-4:] for i in range(1,len(bills))]
                        merged=main+"/"+"/".join(extra)+suf
                    vd=[b[1] for b in bills if b[1]]
                    bill_ws.cell(row,bill_cols["BILL NO"]).value=merged; bill_ws.cell(row,bill_cols["BILL DATE"]).value=fmt_date(max(vd) if vd else None)
                row+=1; sr+=1; added+=1
            upd(96,"Saving workbook…","Writing to disk")
            out=f"UPDATED_BILL_REPORT_{month}.xlsx"; bill_wb.save(out)
            pb.progress(100); t1=datetime.now(); dur=(t1-t0).total_seconds()
            se.empty(); de.empty(); pb.empty()
            st.session_state.processing_history.append({"time":t1.strftime("%H:%M:%S"),"month":month,"status":"success","added":added,"skipped":skipped,"duration":f"{dur:.1f}s"})

            st.markdown(f"""
<div class="results-section">
<div class="results-inner">
    <div class="sec-eyebrow">Complete</div>
    <div class="results-headline">Report<br><em>Generated</em></div>
    <div class="metrics-row">
        <div class="metric-tile"><div class="mt-num">{added}</div><div class="mt-lbl">Entries Added</div></div>
        <div class="metric-tile"><div class="mt-num">{skipped}</div><div class="mt-lbl">Skipped</div></div>
        <div class="metric-tile"><div class="mt-num">{dur:.1f}s</div><div class="mt-lbl">Duration</div></div>
    </div>
</div>
</div>""", unsafe_allow_html=True)

            st.markdown(f'<div style="background:{BG2};padding:0 2.5rem 4rem;"><div style="max-width:1200px;margin:0 auto;">', unsafe_allow_html=True)
            dl1,dl2,_ = st.columns([1,1,2],gap="medium")
            with dl1:
                with open(out,"rb") as f:
                    st.download_button("↓ Excel",data=f,file_name=out,mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True,key="dl_xl")
            with dl2:
                try:
                    df=pd.read_excel(out,sheet_name=month,header=BILL_HEADER_ROW-1); csv=df.to_csv(index=False)
                    st.download_button("↓ CSV",data=csv,file_name=f"UPDATED_BILL_REPORT_{month}.csv",mime="text/csv",use_container_width=True,key="dl_csv")
                except: pass
            st.markdown('</div></div>', unsafe_allow_html=True)

        except Exception as e:
            pb.empty(); se.empty(); de.empty()
            st.session_state.processing_history.append({"time":datetime.now().strftime("%H:%M:%S"),"month":month,"status":"error","added":0,"skipped":0,"duration":"—"})
            st.error(f"Processing error: {str(e)}")
            with st.expander("Full error details"): st.exception(e)

# FOOTER
st.markdown(f'<div class="bf-footer"><div class="bf-footer-inner"><div class="footer-logo">BillFlow <span class="footer-script">Studio</span></div><div class="footer-copy">Report Automation Suite &nbsp;·&nbsp; All rights reserved</div></div></div>', unsafe_allow_html=True)