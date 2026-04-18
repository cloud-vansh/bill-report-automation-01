import streamlit as st
from openpyxl import load_workbook
from collections import defaultdict
from datetime import datetime
import pandas as pd
import re

st.set_page_config(
    page_title="BillFlow · Studio",
    page_icon="✦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

for k, v in {
    'selected_month': max(0, datetime.now().month - 2),
    'processing_history': [],
    'show_preview': False,
    'dark_mode': False,
    'last_result': None,
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
def norm_hdr(x):
    if not x: return ""
    s = str(x).strip().upper()
    s = re.sub(r'\s+', ' ', s)
    s = s.replace(".", "").replace("#", "").replace("_", " ")
    return s

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

def col_map(ws):
    for r in range(1, 25):
        m = {}
        hits = 0
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v:
                nv = norm_hdr(v)
                if nv == "SHIPPER NAME": nv = "PARTY NAME"
                m[nv] = c
                if nv in ["JOB NO", "BILL NO", "INVOICE NO", "PARTY NAME", "CONTAINER NO", "BILL DATE", "JOB DATE", "SB / BE NO"]:
                    hits += 1
        if hits >= 2:
            return m, r
    return {}, 0

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
    RAD_TXT="rgba(240,228,196,.55)"; RAD_SEL="#D4602A"; RAD_STX="#fff"
    CHK_TXT="rgba(240,228,196,.55)"
    BTN_BG="rgba(255,220,140,.07)"; BTN_BRD="rgba(255,220,140,.18)"; BTN_TXT="rgba(240,228,196,.55)"
    EXEC_BG="#D4602A"; EXEC_BD="#D4602A"; EXEC_TX="#fff"
    GREEN_STRIPE="#4A8A5A"; RED_STRIPE="#D4602A"; GOLD_STRIPE="#C89828"
    PROG_TRK="rgba(255,220,140,.1)"; PROG_TXT="#D4602A"
    TOGGLE_IC="☀"; TOGGLE_LBL="Light Mode"
    MODE_OK_C="#7ED49A"; MODE_OK_BG="rgba(74,138,90,.15)"; MODE_OK_BD="#4A8A5A"
    MODE_WN_C="#E8A888"; MODE_WN_BG="rgba(212,96,42,.15)"; MODE_WN_BD="#D4602A"
    TILE_NUM_C="rgba(255,220,140,.05)"; TILE_TIT_C="#F0E4C4"
    TILE_DSC_C="#8A7050"; TILE_KCK_C="#8A7050"
    EXP_BG="transparent"; EXP_TXT="rgba(240,228,196,.45)"
else:
    BG="#F4ECD6"; BG2="#FAF5E4"; BG3="#EDE0C2"
    BORDER="rgba(30,15,2,.12)"; BORDER2="rgba(30,15,2,.22)"
    INK="#1E0F02"; INK2="#5C3418"; INK3="#A07040"
    RUST="#A63418"; RUST_LT="#C4501E"; RUST_PL="rgba(166,52,24,.08)"
    SAGE="#3A6B4A"; SAGE_PL="rgba(58,107,74,.08)"
    GOLD="#8A6A18"; GOLD_PL="rgba(138,106,24,.08)"
    UP_BG="#FAF5E4"; UP_HBRD="#A63418"
    UP_TXT="#5C3418"; UP_BTN="#EDE0C2"; UP_BTXT="#3A1E08"
    SEL_BG="#FAF5E4"; SEL_BRD="rgba(30,15,2,.20)"; SEL_TXT="#1E0F02"
    RAD_BG="#EDE0C2"; RAD_BRD="rgba(30,15,2,.15)"
    RAD_TXT="#5C3418"; RAD_SEL="#1E0F02"; RAD_STX="#F4ECD6"
    CHK_TXT="#3A1E08"
    BTN_BG="#1E0F02"; BTN_BRD="#1E0F02"; BTN_TXT="#F4ECD6"
    EXEC_BG="#1E0F02"; EXEC_BD="#1E0F02"; EXEC_TX="#F4ECD6"
    GREEN_STRIPE="#3A6B4A"; RED_STRIPE="#A63418"; GOLD_STRIPE="#8A6A18"
    PROG_TRK="rgba(30,15,2,.10)"; PROG_TXT="#A63418"
    TOGGLE_IC="◑"; TOGGLE_LBL="Dark Mode"
    MODE_OK_C="#1A3D28"; MODE_OK_BG="rgba(58,107,74,.13)"; MODE_OK_BD="#3A6B4A"
    MODE_WN_C="#5A1A08"; MODE_WN_BG="rgba(166,52,24,.12)"; MODE_WN_BD="#A63418"
    TILE_NUM_C="rgba(30,15,2,.06)"; TILE_TIT_C="#1E0F02"
    TILE_DSC_C="#5C3418"; TILE_KCK_C="#7A5030"
    EXP_BG="transparent"; EXP_TXT="#5C3418"

CSS = f"""
@import url('https://fonts.googleapis.com/css2?family=Cormorant:ital,wght@0,300;0,400;0,600;0,700;1,300;1,400;1,700&family=Libre+Baskerville:ital,wght@0,400;0,700&family=Jost:wght@300;400;500;600;700&family=Caveat:wght@500;600;700&display=swap');

header[data-testid="stHeader"],
[data-testid="stHeader"],
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"],
#MainMenu, footer, .stDeployButton, [data-testid="collapsedControl"] {{
    display: none !important; height: 0 !important; min-height: 0 !important;
}}
*, *::before, *::after {{ box-sizing: border-box; }}
html {{ scroll-behavior:smooth; scrollbar-gutter:stable; }}
body,
[data-testid="stAppViewContainer"],
.main,
[data-testid="stAppViewBlockContainer"],
[data-testid="stMainBlockContainer"] {{
    background:{BG} !important; color:{INK} !important;
    font-family:'Jost',sans-serif !important;
}}

/* HERO AND BRANDING */
.hero-wrap {{ padding: 2.5rem 1rem 3rem; text-align: center; border-bottom:1px solid {BORDER}; }}
.hero-logo {{ font-family:'Cormorant',serif; font-size:1.6rem; font-weight:700; color:{INK}; position:absolute; top:2rem; left:2.5rem; }}
.hero-script {{ font-family:'Caveat',cursive; font-size:1.2rem; font-weight:600; color:{RUST}; }}
.sub-script {{ font-family:'Caveat',cursive; font-size:1.6rem; font-weight:600; color:{RUST}; display:block; margin: 3rem auto 0.5rem; }}
.hero-title {{ font-family:'Cormorant',serif; font-size:5.5rem; font-weight:700; color:{INK}; line-height:.95; letter-spacing:-.03em; margin-bottom:1.5rem; }}
.hero-title em {{ font-style:italic; font-weight:300; color:{INK2}; }}
.hero-desc {{ font-family:'Libre Baskerville',serif; font-style:italic; font-size:1.15rem; color:{INK3}; max-width:800px; margin:0 auto; line-height:1.6; padding:0 1rem; }}

/* LAYOUT SPACING */
.sec-wrap {{ padding: 3rem 2.5rem 1rem; max-width:1600px; margin: 0 auto; }}
.sec-title {{ font-family:'Jost',sans-serif; font-size:.7rem; font-weight:700; letter-spacing:.25em; text-transform:uppercase; color:{INK3}; text-align:center; margin-bottom:2.5rem; }}

/* TILES 3-COLUMN */
.tile-wrap {{ background:{BG2}; border:1px solid {BORDER}; border-top:4px solid; padding:2rem 1.5rem; position:relative; overflow:hidden; display:flex; flex-direction:column; height: 100%; }}
.tile-num-bg {{ position:absolute; right:1.2rem; top:1.2rem; font-family:'Cormorant',serif; font-size:7rem; font-weight:700; color:{TILE_NUM_C}; line-height:.8; letter-spacing:-.04em; pointer-events:none; user-select:none; }}
.tile-kicker {{ font-family:'Jost',sans-serif; font-size:.6rem; font-weight:700; letter-spacing:.2em; text-transform:uppercase; color:{TILE_KCK_C}; display:block; margin-bottom:.5rem; }}
.tile-title {{ font-family:'Cormorant',serif; font-size:1.6rem; font-weight:700; color:{TILE_TIT_C}; line-height:1; margin-bottom:.5rem; z-index:2; position:relative; }}
.tile-desc {{ font-family:'Libre Baskerville',serif; font-style:italic; font-size:.85rem; color:{TILE_DSC_C}; line-height:1.5; z-index:2; position:relative; margin-bottom:2.5rem; }}

.tile-status-ok {{ display:inline-flex; align-items:center; gap:.4rem; font-family:'Jost',sans-serif; font-size:.65rem; font-weight:700; letter-spacing:.1em; text-transform:uppercase; padding:.3rem .8rem; background:{SAGE_PL}; border:1px solid {SAGE}; color:{SAGE}; margin-top:auto; }}
.tile-status-wait {{ display:inline-flex; align-items:center; gap:.4rem; font-family:'Jost',sans-serif; font-size:.65rem; font-weight:700; letter-spacing:.1em; text-transform:uppercase; padding:.3rem .8rem; background:transparent; border:1px dashed {BORDER2}; color:{INK3}; margin-top:auto; }}
.tp-row {{ display:flex; justify-content:space-between; color:{INK2}; padding:.2rem 0; border-bottom:1px solid {BORDER}; font-family:'Courier New',monospace; font-size:.7rem; margin-top:.5rem; }}
.tp-val {{ font-weight:700; color:{RUST}; }}

/* FILE UPLOADER LIGHT MODE PATCH */
[data-testid="stFileUploader"] {{ margin-top:-1rem; z-index:2; position:relative; }}
[data-testid="stFileUploader"] section {{ background:transparent !important; border:1px dashed {BORDER2} !important; padding:1rem !important; min-height:0 !important; cursor:pointer !important; transition:0.2s all !important; }}
[data-testid="stFileUploader"] section:hover {{ background:{BG3} !important; border-color:{INK2} !important; }}
[data-testid="stFileUploader"] section p, 
[data-testid="stFileUploader"] section span, 
[data-testid="stFileUploader"] section small, 
[data-testid="stFileUploader"] section div {{ font-family:'Jost',sans-serif !important; font-size:.75rem !important; color:{INK} !important; margin:0 !important; font-weight:600 !important; }}
[data-testid="stFileUploader"] section button {{ background:{BG3} !important; border:1px solid {BORDER} !important; color:{INK} !important; font-size:.65rem !important; font-weight:700 !important; padding:.3rem .8rem !important; margin-top:.5rem !important; text-transform:uppercase !important; border-radius:0 !important; }}

/* CONFIG AREA */
.cfg-box {{ background:{BG2}; border:1px solid {BORDER}; padding:3rem; max-width:800px; margin:0 auto; box-shadow:0 15px 40px rgba(0,0,0,0.03); }}
.cfg-label {{ font-family:'Jost',sans-serif; font-size:.65rem; font-weight:700; letter-spacing:.2em; text-transform:uppercase; color:{INK3}; display:block; margin-bottom:.7rem; text-align:center; }}

[data-testid="stSelectbox"] > div > div {{ background:{SEL_BG} !important; border:1px solid {SEL_BRD} !important; border-radius:0 !important; color:{SEL_TXT} !important; font-family:'Courier New',monospace !important; font-size:1.1rem !important; font-weight:700 !important; text-align:center !important; }}
[data-testid="stSelectbox"] > div > div > div {{ justify-content:center !important; }}

/* RADIO FIX FOR LIGHT/DARK MODE */
[data-testid="stRadio"] > div {{ display:flex !important; flex-direction:row !important; background:{RAD_BG} !important; border:1px solid {RAD_BRD} !important; border-radius:0 !important; padding:4px !important; gap:0 !important; height:100% !important; }}
[data-testid="stRadio"] label > div:first-child, [data-testid="stRadio"] input[type="radio"] {{ display:none !important; }}
[data-testid="stRadio"] label {{ padding:.6rem !important; font-family:'Jost',sans-serif !important; cursor:pointer !important; flex:1 !important; text-align:center !important; margin:0 !important; border-radius:0 !important; transition:0.2s all !important; }}
[data-testid="stRadio"] label p {{ color:{RAD_TXT} !important; font-weight:600 !important; font-size:.85rem !important; margin:0 !important; transition:0.2s color !important; }}
[data-testid="stRadio"] label:has(input:checked) {{ background:{RAD_SEL} !important; }}
[data-testid="stRadio"] label:has(input:checked) p {{ color:{RAD_STX} !important; font-weight:700 !important; }}

.mode-note {{ font-family:'Libre Baskerville',serif; font-style:italic; font-size:.85rem; line-height:1.6; padding:1.2rem; margin:1.5rem 0; border-left:3px solid; text-align:center; }}
.mode-ok {{ color:{MODE_OK_C}; background:{MODE_OK_BG}; border-color:{MODE_OK_BD}; }}
.mode-warn {{ color:{MODE_WN_C}; background:{MODE_WN_BG}; border-color:{MODE_WN_BD}; }}

/* BUTTON */
.stButton button, .stDownloadButton button {{ background:{EXEC_BG} !important; border:1px solid {EXEC_BD} !important; color:{EXEC_TX} !important; font-family:'Jost',sans-serif !important; font-size:.85rem !important; font-weight:700 !important; letter-spacing:.15em !important; text-transform:uppercase !important; border-radius:0 !important; padding:1rem !important; box-shadow:0 8px 25px rgba(0,0,0,.08) !important; width:100% !important; transition:0.3s all !important; }}
.stButton button:hover, .stDownloadButton button:hover {{ transform:translateY(-2px) !important; filter:brightness(1.15) !important; box-shadow:0 12px 35px rgba(0,0,0,.15) !important; color:{EXEC_TX} !important; border-color:{EXEC_BD} !important; }}

.theme-toggle-wrap {{ position:absolute; top:2rem; right:2.5rem; z-index:100; max-width:140px; }}

/* PROGRESS */
[data-testid="stProgress"] > div {{ background:{PROG_TRK} !important; border-radius:0 !important; height:3px !important; margin-top:2rem !important; }}
[data-testid="stProgress"] > div > div {{ background:{RUST} !important; border-radius:0 !important; }}
.prog-txt {{ font-family:'Courier New',monospace; font-size:.8rem; color:{PROG_TXT}; margin-bottom:.3rem; font-weight:700; text-align:center; letter-spacing:0.05em; }}
.prog-det {{ font-family:'Courier New',monospace; font-size:.7rem; color:{INK3}; display:flex; justify-content:space-between; }}

/* RESULTS */
.res-box {{ background:{BG2}; border:1px solid {BORDER}; border-top:4px solid {RUST}; padding:3rem; margin-top:3rem; text-align:center; max-width:800px; margin-left:auto; margin-right:auto; box-shadow:0 20px 50px rgba(0,0,0,.05); }}
.res-title {{ font-family:'Cormorant',serif; font-size:3rem; font-weight:700; color:{INK}; margin-bottom:2.5rem; line-height:1; letter-spacing:-.02em; }}
.res-metrics {{ display:flex; justify-content:space-around; align-items:center; gap:2rem; margin-bottom:3rem; border-bottom:1px solid {BORDER}; padding-bottom:3rem; }}
.res-m {{ display:flex; flex-direction:column; gap:.5rem; }}
.res-m-val {{ font-family:'Cormorant',serif; font-size:4rem; font-weight:700; color:{RUST}; line-height:.9; }}
.res-m-lbl {{ font-family:'Jost',sans-serif; font-size:.65rem; font-weight:700; letter-spacing:.2em; text-transform:uppercase; color:{INK3}; }}

.history-wrap {{ text-align:center; padding:4rem 0; border-top:1px dashed {BORDER}; margin-top:4rem; }}
.he-title {{ font-family:'Jost',sans-serif; font-size:.65rem; font-weight:700; letter-spacing:.2em; text-transform:uppercase; color:{INK3}; margin-bottom:2rem; }}
.he-row {{ font-family:'Courier New',monospace; font-size:.75rem; color:{INK2}; margin-bottom:.5rem; display:flex; justify-content:center; gap:1rem; align-items:center; }}
.he-ok {{ color:{SAGE}; font-weight:700; }}
.he-err {{ color:{RUST}; font-weight:700; }}
"""

st.markdown(f"<style>{CSS}</style>", unsafe_allow_html=True)

# THEME BUTTON (Absolute Top Right via flex wrapper)
st.markdown('<div class="theme-toggle-wrap">', unsafe_allow_html=True)
if st.button(f"{TOGGLE_IC}  {TOGGLE_LBL}", key="theme_toggle", use_container_width=True):
    st.session_state.dark_mode = not st.session_state.dark_mode
    st.rerun()
st.markdown('</div>', unsafe_allow_html=True)

# HERO
st.markdown(f'''
<span class="hero-logo">BillFlow <span class="hero-script" style="display:inline;">Studio</span></span>
<div class="hero-wrap">
    <span class="sub-script">Data Concierge</span>
    <div class="hero-title">Automated Bill <em>Reconciliation</em></div>
    <div class="hero-desc">Supply your Master Ledger alongside standard Job & E-Invoice extracts. BillFlow dynamically synthesizes the parameters and regenerates completely mapped outputs instantaneously.</div>
</div>
''', unsafe_allow_html=True)

# UPLOADERS
st.markdown('<div class="sec-wrap"><div class="sec-title">1. Source Documents</div>', unsafe_allow_html=True)

c1, c2, c3 = st.columns(3, gap="large")

with c1:
    st.markdown(f'''
    <div class="tile-wrap" style="border-top-color:{GREEN_STRIPE};">
        <div class="tile-num-bg">I</div>
        <div>
            <span class="tile-kicker">Master Ledger</span>
            <div class="tile-title">Bill Report</div>
            <div class="tile-desc">Master billing workbook containing localized monthly sheets.</div>
        </div>
    ''', unsafe_allow_html=True)
    bill_file = st.file_uploader("Bill Report", type="xlsx", key="bill", label_visibility="collapsed")
    if bill_file:
        v = validate_file(bill_file, "bill")
        if v["valid"]:
            sn = len(v["sheets"]) if v["sheets"] else "—"
            st.markdown(f'<div class="tile-status-ok">✓ Loaded securely</div><div class="tp-row"><span>Rows</span><span class="tp-val">{v["rows"]}</span></div><div class="tp-row"><span>Sheets</span><span class="tp-val">{sn}</span></div>', unsafe_allow_html=True)
        bill_file.seek(0)
    else:
        st.markdown('<div class="tile-status-wait">○ Awaiting Document</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

with c2:
    st.markdown(f'''
    <div class="tile-wrap" style="border-top-color:{RED_STRIPE};">
        <div class="tile-num-bg">II</div>
        <div>
            <span class="tile-kicker">Job Details</span>
            <div class="tile-title">Job Report</div>
            <div class="tile-desc">Shipper names, container allocations, and reference ID.</div>
        </div>
    ''', unsafe_allow_html=True)
    job_file = st.file_uploader("Job Report", type="xlsx", key="job", label_visibility="collapsed")
    if job_file:
        v = validate_file(job_file, "job")
        if v["valid"]:
            st.markdown(f'<div class="tile-status-ok">✓ Loaded securely</div><div class="tp-row"><span>Rows</span><span class="tp-val">{v["rows"]}</span></div><div class="tp-row"><span>Cols</span><span class="tp-val">{v["cols"]}</span></div>', unsafe_allow_html=True)
        job_file.seek(0)
    else:
        st.markdown('<div class="tile-status-wait">○ Awaiting Document</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

with c3:
    st.markdown(f'''
    <div class="tile-wrap" style="border-top-color:{GOLD_STRIPE};">
        <div class="tile-num-bg">III</div>
        <div>
            <span class="tile-kicker">Issuance</span>
            <div class="tile-title">E-Invoice Data</div>
            <div class="tile-desc">Bill metrics directly tied strictly to organic job variables.</div>
        </div>
    ''', unsafe_allow_html=True)
    einv_file = st.file_uploader("E-Invoice", type="xlsx", key="einv", label_visibility="collapsed")
    if einv_file:
        v = validate_file(einv_file, "einv")
        if v["valid"]:
            st.markdown(f'<div class="tile-status-ok">✓ Loaded securely</div><div class="tp-row"><span>Rows</span><span class="tp-val">{v["rows"]}</span></div><div class="tp-row"><span>Cols</span><span class="tp-val">{v["cols"]}</span></div>', unsafe_allow_html=True)
        einv_file.seek(0)
    else:
        st.markdown('<div class="tile-status-wait">○ Awaiting Document</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# CONFIG AND EXECUTION
st.markdown('<div class="sec-wrap" style="padding-top:0;"><div class="sec-title">2. Configuration & Execution</div></div>', unsafe_allow_html=True)

cx_left, cx_center, cx_right = st.columns([1, 4, 1])

with cx_center:
    mcols = st.columns(2, gap="large")
    with mcols[0]:
        st.markdown('<span class="cfg-label">Target Cycle</span>', unsafe_allow_html=True)
        month = st.selectbox("Month", MONTHS, index=st.session_state.selected_month, label_visibility="collapsed")
        st.session_state.selected_month = MONTHS.index(month)
    with mcols[1]:
        st.markdown('<span class="cfg-label">Injection Strategy</span>', unsafe_allow_html=True)
        update_mode = st.radio("Mode", ["Append","Overwrite"], index=1, label_visibility="collapsed")

    if update_mode == "Append":
        st.markdown('<div class="mode-note mode-ok">Safely appends unique missing nodes directly without stripping foundational table data.</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="mode-note mode-warn">Destructively resets the primary workbook tab prior to fresh compilation. Standard operating mode.</div>', unsafe_allow_html=True)

    exe = st.button("✦ Generate Report ✦", use_container_width=True)

prog_ui = st.empty()

if exe:
    st.session_state.last_result = None
    if not (bill_file and job_file and einv_file):
        st.error("Protocol Exception: All three origin fragments must be satisfied before compiling.")
    else:
        t0 = datetime.now()

        def upd(p, msg, det=""):
            html = f"""
            <div style="max-width:600px; margin:2rem auto;">
                <div class="prog-txt">{msg}</div>
                <div class="prog-det"><span>{det}</span><span>{p}%</span></div>
            </div>
            """
            with prog_ui.container():
                st.markdown(html, unsafe_allow_html=True)
                st.progress(p)

        try:
            upd(10,"Authorizing Data Handlers…","Buffer instantiation")
            bill_wb=load_workbook(bill_file); job_wb=load_workbook(job_file); einv_wb=load_workbook(einv_file)
            upd(22,f"Scanning Ledger Domain '{month}'…","Parsing structure")
            if month not in bill_wb.sheetnames:
                st.error(f"Sheet '{month}' not located in Master Ledger. Handlers detected: {', '.join(bill_wb.sheetnames)}")
                st.stop()
                
            bill_ws=bill_wb[month]; job_ws=job_wb.active; einv_ws=einv_wb.active
            
            upd(35,"Analyzing Header Formats…","Robust string normalization")
            bill_cols, bill_hr = col_map(bill_ws)
            job_cols, job_hr = col_map(job_ws)
            einv_cols, einv_hr = col_map(einv_ws)
            
            bill_ds = (bill_hr + 1) if bill_hr else BILL_DATA_START
            job_ds = (job_hr + 1) if job_hr else SRC_DATA_START
            einv_ds = (einv_hr + 1) if einv_hr else SRC_DATA_START

            upd(45,"Validating Logical Parity…","Validating required nodes")
            for c in ["JOB NO","JOB DATE","PARTY NAME","INVOICE NO","SB / BE NO","CONTAINER NO"]:
                if c not in job_cols: st.error(f"Missing in Job Component: '{c}'"); st.stop()
            for c in ["JOB NO","BILL NO","BILL DATE"]:
                if c not in einv_cols: st.error(f"Missing in E-Invoice Component: '{c}'"); st.stop()
            for c in ["JOB NO","JOB DATE","PARTY NAME","INVOICE NO","SB / BE NO","CONTAINER NO","BILL NO","BILL DATE"]:
                if c not in bill_cols: st.error(f"Missing in Master Ledger: '{c}'"); st.stop()
                
            append_mode=(update_mode=="Append"); existing_jobs=set()
            if append_mode:
                upd(55,"Mapping Existing Indexes…","Generating append offset")
                existing_jobs=get_existing_jobs(bill_ws,bill_cols,bill_ds)
                row=next_empty(bill_ws,bill_ds)
                last_sr=0
                for r in range(bill_ds,row):
                    sv=bill_ws.cell(r,1).value
                    if sv and isinstance(sv,(int,float)): last_sr=max(last_sr,int(sv))
                sr=last_sr+1
            else:
                upd(55,"Truncating Primary Ledger…","Sanitizing output destination")
                for r in range(bill_ds,bill_ws.max_row+1):
                    for c in range(1,bill_ws.max_column+1): bill_ws.cell(r,c).value=None
                row=bill_ds; sr=1
                
            upd(65,"Ingesting Job Data…","Translating values to buffer")
            jobs={}
            for r in range(job_ds,job_ws.max_row+1):
                raw=job_ws.cell(r,job_cols["JOB NO"]).value
                if not raw: continue
                jn=clean_job(raw)
                if jn in jobs: continue
                jobs[jn]={"JOB DATE":parse_date(job_ws.cell(r,job_cols["JOB DATE"]).value),"PARTY NAME":job_ws.cell(r,job_cols["PARTY NAME"]).value,"INVOICE NO":job_ws.cell(r,job_cols["INVOICE NO"]).value,"SB / BE NO":job_ws.cell(r,job_cols["SB / BE NO"]).value,"CONTAINER NO":job_ws.cell(r,job_cols["CONTAINER NO"]).value}
                
            upd(78,"Linking Correlated Elements…","Integrating cross-referenced metrics")
            bill_map=defaultdict(list)
            for r in range(einv_ds,einv_ws.max_row+1):
                raw=einv_ws.cell(r,einv_cols["JOB NO"]).value; bno=einv_ws.cell(r,einv_cols["BILL NO"]).value; bd=parse_date(einv_ws.cell(r,einv_cols["BILL DATE"]).value)
                if raw and bno: bill_map[clean_job(raw)].append((str(bno),bd))
                
            upd(88,"Synthesizing Final Output…","Injecting matrix into active workbook")
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
                
            upd(96,"Solidifying Write Stream…","Flushing virtual memory to .XLSX encoding")
            out=f"UPDATED_BILL_REPORT_{month}.xlsx"; bill_wb.save(out)
            
            t1=datetime.now(); dur=(t1-t0).total_seconds()
            st.session_state.processing_history.append({"time":t1.strftime("%H:%M:%S"),"month":month,"status":"success","added":added,"skipped":skipped,"duration":f"{dur:.1f}s"})
            
            st.session_state.last_result = {
                "added": added,
                "skipped": skipped,
                "dur": f"{dur:.1f}s",
                "out": out,
                "month": month,
                "hdr": (bill_hr-1 if bill_hr else 2)
            }

            prog_ui.empty()
        except Exception as e:
            prog_ui.empty()
            st.session_state.processing_history.append({"time":datetime.now().strftime("%H:%M:%S"),"month":month,"status":"error","added":0,"skipped":0,"duration":"—"})
            st.error(f"Compiler Fault: {str(e)}")
            with st.expander("System Diagnostic Trace"):
                st.exception(e)

# PERSISTENT RESULTS RENDER CENTERED
if st.session_state.get('last_result'):
    lr = st.session_state.last_result
    st.markdown(f"""
    <div class="res-box">
        <div class="res-title">Compilation Successful</div>
        <div class="res-metrics">
            <div class="res-m"><span class="res-m-val">{lr['added']}</span><span class="res-m-lbl">Rows Integrated</span></div>
            <div class="res-m"><span class="res-m-val">{lr['dur']}</span><span class="res-m-lbl">Run Duration</span></div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Download buttons perfectly aligned beneath result box
    d_wrap = st.container()
    d1, d2, d3, d4 = d_wrap.columns([1, 2, 2, 1], gap="large")
    
    try:
        with open(lr['out'], "rb") as f:
            xl_data = f.read()
            d2.download_button("↓ Secure Download (Excel)", data=xl_data, file_name=lr['out'], use_container_width=True)
            
        df = pd.read_excel(lr['out'], sheet_name=lr['month'], header=lr['hdr'])
        csv_data = df.to_csv(index=False)
        d3.download_button("↓ Raw Extract (CSV)", data=csv_data, file_name=f"UPDATED_BILL_REPORT_{lr['month']}.csv", mime="text/csv", use_container_width=True)
    except Exception as e:
        d_wrap.error("Export fragments deleted organically. Re-run compilation to regenerate.", icon="⚠️")

# EMBEDDED HISTORY (Hidden strictly until content available)
if st.session_state.processing_history:
    st.markdown('<div class="history-wrap"><div class="he-title">Audit Protocol · Recent Activities</div>', unsafe_allow_html=True)
    for entry in reversed(st.session_state.processing_history[-5:]):
        ok = '<span class="he-ok">SUCCESS</span>' if entry['status']=='success' else '<span class="he-err">CRITICAL ABORT</span>'
        dur = entry.get('duration','—')
        st.markdown(f'<div class="he-row"><span>{entry["time"]}</span><span>·</span><span>{entry["month"]}</span><span>·</span>{ok}<span>·</span><span>+{entry["added"]} rows mapped in {dur}</span></div>', unsafe_allow_html=True)
    
    hc1, hc2, hc3 = st.columns([2, 1, 2])
    with hc2:
        st.write("") # spacing
        if st.button("Flush Data Audit Log", use_container_width=True):
            st.session_state.processing_history = []
            st.session_state.last_result = None
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True) # close sec wrap