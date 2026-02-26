import streamlit as st
from openpyxl import load_workbook
from collections import defaultdict
from datetime import datetime
import pandas as pd

# ─── PAGE CONFIG ───
st.set_page_config(
    page_title="BillFlow Studio",
    page_icon="✦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ─── SESSION STATE ───
if 'selected_month'      not in st.session_state: st.session_state.selected_month      = datetime.now().month - 1
if 'processing_history'  not in st.session_state: st.session_state.processing_history  = []
if 'show_preview'        not in st.session_state: st.session_state.show_preview        = False
if 'show_settings'       not in st.session_state: st.session_state.show_settings       = False

# ─── DESIGN SYSTEM ───
THEME = """
@import url('https://fonts.googleapis.com/css2?family=Abril+Fatface&family=Dancing+Script:wght@500;600;700&family=Jost:wght@300;400;500;600&family=Courier+Prime:ital,wght@0,400;0,700;1,400&family=Playfair+Display:ital,wght@0,700;0,800;1,700&display=swap');

:root {
    /* Cream & Brown palette — straight from Havanza */
    --cream:         #FAF2E2;
    --cream-dark:    #F2E6CC;
    --cream-deeper:  #E8D8B4;
    --espresso:      #3B1A09;
    --brown-mid:     #6B3520;
    --brown-light:   #9B6040;
    --brown-pale:    #C8A882;
    --sienna:        #C4622A;
    --sienna-light:  #E07840;
    --sienna-pale:   rgba(196,98,42,0.1);

    /* Card frame colours — like Havanza's gallery borders */
    --frame-green:   #5B9E6A;
    --frame-orange:  #D4733A;
    --frame-yellow:  #C8A82A;
    --frame-teal:    #3A8A8A;

    /* Typography */
    --font-display:  'Abril Fatface', Georgia, serif;
    --font-script:   'Dancing Script', cursive;
    --font-serif:    'Playfair Display', Georgia, serif;
    --font-body:     'Jost', sans-serif;
    --font-mono:     'Courier Prime', monospace;

    /* Shadows */
    --shadow-sm:  0 2px 8px rgba(59,26,9,0.08);
    --shadow-md:  0 4px 20px rgba(59,26,9,0.12);
    --shadow-lg:  0 8px 40px rgba(59,26,9,0.16);
    --shadow-xl:  0 16px 64px rgba(59,26,9,0.20);

    --radius-sm:  6px;
    --radius-md:  12px;
    --radius-lg:  20px;
    --radius-pill:50px;
}

/* ── GLOBAL ── */
*, *::before, *::after { box-sizing: border-box; }

html, body,
[data-testid="stAppViewContainer"],
[data-testid="stAppViewBlockContainer"],
.main .block-container {
    font-family: var(--font-body) !important;
    background: var(--cream) !important;
    color: var(--espresso) !important;
}

.main { background: var(--cream) !important; }
.block-container {
    max-width: 1380px !important;
    padding: 0 0 3rem 0 !important;
}

/* ── HIDE STREAMLIT CHROME ── */
#MainMenu, footer, .stDeployButton,
[data-testid="stToolbar"], [data-testid="stDecoration"],
[data-testid="stStatusWidget"] { visibility: hidden !important; display: none !important; }

/* ── SCROLLBAR ── */
::-webkit-scrollbar { width: 4px; }
::-webkit-scrollbar-track { background: var(--cream-dark); }
::-webkit-scrollbar-thumb { background: var(--brown-pale); border-radius: 4px; }

/* ══════════════════════════════════
   NAV BAR  (top strip)
══════════════════════════════════ */
.hv-nav {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0 2.5rem;
    height: 68px;
    background: var(--cream);
    border-bottom: 1px solid rgba(59,26,9,0.12);
    position: sticky; top: 0; z-index: 100;
}
.nav-left, .nav-right {
    display: flex;
    align-items: center;
    gap: 2rem;
}
.nav-link {
    font-family: var(--font-body);
    font-size: 0.88rem;
    font-weight: 500;
    color: var(--brown-mid);
    letter-spacing: 0.02em;
    cursor: pointer;
    transition: color 0.2s;
}
.nav-link:hover { color: var(--espresso); }
.nav-wordmark {
    font-family: var(--font-body);
    font-size: 1rem;
    font-weight: 600;
    letter-spacing: 0.22em;
    color: var(--espresso);
    text-transform: uppercase;
    display: flex;
    align-items: center;
    gap: 0.55rem;
}
.nav-wordmark-icon { color: var(--sienna); font-size: 1.1rem; }
.nav-pill {
    font-family: var(--font-body);
    font-size: 0.82rem;
    font-weight: 600;
    color: var(--cream) !important;
    background: var(--espresso);
    border: none;
    padding: 0.55rem 1.4rem;
    border-radius: var(--radius-pill);
    letter-spacing: 0.04em;
    cursor: pointer;
    transition: background 0.2s, transform 0.15s;
}
.nav-pill:hover { background: var(--brown-mid); transform: translateY(-1px); }

/* ══════════════════════════════════
   HERO BANNER
══════════════════════════════════ */
.hv-hero {
    background: var(--cream-dark);
    padding: 3.5rem 2.5rem 2.5rem;
    text-align: center;
    position: relative;
    overflow: hidden;
    border-bottom: 1px solid rgba(59,26,9,0.08);
}
/* Subtle radial warm glow */
.hv-hero::before {
    content: '';
    position: absolute;
    inset: 0;
    background: radial-gradient(ellipse 70% 80% at 50% 100%, rgba(196,98,42,0.08) 0%, transparent 70%);
    pointer-events: none;
}
.hero-script {
    font-family: var(--font-script);
    font-size: 1.6rem;
    color: var(--sienna);
    display: block;
    margin-bottom: 0.25rem;
    animation: fadeInDown 0.5s ease both;
}
.hero-title {
    font-family: var(--font-display);
    font-size: clamp(2.8rem, 6vw, 5rem);
    color: var(--espresso);
    line-height: 0.95;
    letter-spacing: -0.01em;
    margin: 0 auto 0.4rem;
    animation: fadeInUp 0.55s ease 0.1s both;
}
.hero-title em {
    font-family: var(--font-serif);
    font-style: italic;
    font-size: 0.9em;
    color: var(--brown-mid);
}
.hero-sub {
    font-family: var(--font-body);
    font-size: 0.92rem;
    font-weight: 400;
    color: var(--brown-light);
    letter-spacing: 0.06em;
    margin-bottom: 1.75rem;
    animation: fadeInUp 0.6s ease 0.2s both;
}
.hero-cta-row {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 1rem;
    animation: fadeInUp 0.65s ease 0.3s both;
}
.hero-btn {
    display: inline-flex;
    align-items: center;
    gap: 0.55rem;
    background: var(--espresso);
    color: var(--cream);
    font-family: var(--font-body);
    font-size: 0.82rem;
    font-weight: 600;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    padding: 0.85rem 2rem;
    border-radius: var(--radius-pill);
    box-shadow: var(--shadow-md);
    transition: background 0.2s, transform 0.15s, box-shadow 0.2s;
}
.hero-btn:hover {
    background: var(--brown-mid);
    transform: translateY(-2px);
    box-shadow: var(--shadow-lg);
}
.hero-btn-arrow {
    width: 24px; height: 24px;
    background: var(--sienna);
    border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    font-size: 0.75rem;
}
@keyframes fadeInDown {
    from { opacity:0; transform:translateY(-12px); }
    to   { opacity:1; transform:translateY(0); }
}
@keyframes fadeInUp {
    from { opacity:0; transform:translateY(14px); }
    to   { opacity:1; transform:translateY(0); }
}

/* ══════════════════════════════════
   SECTION TITLE  ("Moments & emotions" style)
══════════════════════════════════ */
.hv-section-row {
    display: flex;
    align-items: center;
    gap: 1.25rem;
    padding: 0 2.5rem;
    margin: 2.5rem 0 1.5rem;
}
.hv-section-title {
    font-family: var(--font-serif);
    font-style: italic;
    font-size: 1.9rem;
    font-weight: 700;
    color: var(--espresso);
    white-space: nowrap;
}
.hv-section-rule {
    flex: 1;
    height: 1px;
    background: rgba(59,26,9,0.15);
}
.hv-section-tag {
    font-family: var(--font-body);
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    color: var(--brown-light);
    background: var(--cream-dark);
    border: 1px solid rgba(59,26,9,0.15);
    padding: 0.3rem 0.85rem;
    border-radius: var(--radius-pill);
}

/* ══════════════════════════════════
   UPLOAD CARDS  (colorful bordered frames)
══════════════════════════════════ */
.cards-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 1.5rem;
    padding: 0 2.5rem;
}
.upload-card {
    background: var(--cream);
    border-radius: var(--radius-lg);
    overflow: hidden;
    box-shadow: var(--shadow-md);
    transition: transform 0.25s, box-shadow 0.25s;
    position: relative;
}
.upload-card:hover {
    transform: translateY(-4px);
    box-shadow: var(--shadow-xl);
}
.card-frame-top {
    height: 6px;
    width: 100%;
}
.card-body { padding: 1.6rem 1.5rem 1.25rem; }
.card-badge {
    display: inline-flex;
    align-items: center;
    gap: 0.4rem;
    font-family: var(--font-body);
    font-size: 0.68rem;
    font-weight: 600;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    padding: 0.28rem 0.75rem;
    border-radius: var(--radius-pill);
    margin-bottom: 1rem;
}
.card-title {
    font-family: var(--font-serif);
    font-size: 1.35rem;
    font-weight: 800;
    color: var(--espresso);
    margin-bottom: 0.3rem;
    line-height: 1.2;
}
.card-desc {
    font-family: var(--font-body);
    font-size: 0.84rem;
    color: var(--brown-light);
    line-height: 1.5;
    margin-bottom: 1rem;
}
.card-loaded {
    display: inline-flex;
    align-items: center;
    gap: 0.4rem;
    font-family: var(--font-body);
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.08em;
    padding: 0.3rem 0.8rem;
    border-radius: var(--radius-pill);
    margin-bottom: 0.65rem;
}
/* Per-card colour theming */
.card-green .card-frame-top   { background: var(--frame-green); }
.card-green .card-badge        { color: var(--frame-green); background: rgba(91,158,106,0.1); border: 1px solid rgba(91,158,106,0.25); }
.card-green .card-loaded       { color: var(--frame-green); background: rgba(91,158,106,0.1); border: 1px solid rgba(91,158,106,0.25); }

.card-orange .card-frame-top  { background: var(--frame-orange); }
.card-orange .card-badge       { color: var(--frame-orange); background: rgba(212,115,58,0.1); border: 1px solid rgba(212,115,58,0.25); }
.card-orange .card-loaded      { color: var(--frame-orange); background: rgba(212,115,58,0.1); border: 1px solid rgba(212,115,58,0.25); }

.card-yellow .card-frame-top  { background: var(--frame-yellow); }
.card-yellow .card-badge       { color: var(--frame-yellow); background: rgba(200,168,42,0.1); border: 1px solid rgba(200,168,42,0.25); }
.card-yellow .card-loaded      { color: var(--frame-yellow); background: rgba(200,168,42,0.1); border: 1px solid rgba(200,168,42,0.25); }

/* preview mini-table */
.preview-mini {
    background: var(--cream-dark);
    border-radius: var(--radius-sm);
    padding: 0.65rem 0.85rem;
    margin-top: 0.55rem;
}
.pm-row {
    display: flex; justify-content: space-between;
    font-family: var(--font-mono);
    font-size: 0.7rem;
    color: var(--brown-mid);
    padding: 0.2rem 0;
    border-bottom: 1px solid rgba(59,26,9,0.07);
}
.pm-row:last-child { border-bottom: none; }
.pm-val { font-weight: 700; color: var(--sienna); }

/* ══════════════════════════════════
   CONFIG PANEL
══════════════════════════════════ */
.config-wrap {
    padding: 0 2.5rem;
    margin-top: 2.5rem;
}
.config-card {
    background: var(--espresso);
    border-radius: var(--radius-lg);
    padding: 2.25rem 2rem;
    box-shadow: var(--shadow-xl);
    position: relative;
    overflow: hidden;
}
.config-card::before {
    content: '';
    position: absolute;
    top: -60px; right: -60px;
    width: 200px; height: 200px;
    border-radius: 50%;
    background: rgba(196,98,42,0.15);
}
.config-card::after {
    content: '';
    position: absolute;
    bottom: -80px; left: -40px;
    width: 240px; height: 240px;
    border-radius: 50%;
    background: rgba(255,255,255,0.03);
}
.config-script {
    font-family: var(--font-script);
    font-size: 1.3rem;
    color: var(--sienna-light);
    display: block;
    margin-bottom: 0.1rem;
}
.config-title {
    font-family: var(--font-display);
    font-size: 2rem;
    color: var(--cream);
    line-height: 1.1;
    margin-bottom: 1.75rem;
    position: relative; z-index: 1;
}
.config-label {
    font-family: var(--font-body);
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    color: var(--brown-pale);
    display: block;
    margin-bottom: 0.5rem;
    position: relative; z-index: 1;
}
.config-divider {
    height: 1px;
    background: rgba(255,255,255,0.08);
    margin: 1.5rem 0;
    position: relative; z-index: 1;
}
.mode-note {
    font-family: var(--font-body);
    font-style: italic;
    font-size: 0.8rem;
    margin-top: 0.6rem;
    padding: 0.55rem 0.85rem;
    border-radius: var(--radius-sm);
    position: relative; z-index: 1;
    line-height: 1.45;
}
.mode-note-green { color: #8EC49A; background: rgba(91,158,106,0.12); border: 1px solid rgba(91,158,106,0.2); }
.mode-note-warn  { color: #E0A880; background: rgba(196,98,42,0.15);  border: 1px solid rgba(196,98,42,0.25); }

/* ══════════════════════════════════
   STREAMLIT OVERRIDES
══════════════════════════════════ */

/* Selectbox */
[data-testid="stSelectbox"] > div > div {
    background: rgba(255,255,255,0.07) !important;
    border: 1px solid rgba(255,255,255,0.18) !important;
    border-radius: var(--radius-md) !important;
    color: var(--cream) !important;
    font-family: var(--font-mono) !important;
    font-size: 0.9rem !important;
    letter-spacing: 0.08em !important;
}
[data-testid="stSelectbox"] > div > div:hover {
    border-color: var(--sienna-light) !important;
}
[data-testid="stSelectbox"] svg { fill: var(--brown-pale) !important; }
[data-testid="stSelectbox"] > div > div > div {
    color: var(--cream) !important;
}

/* Radio */
[data-testid="stRadio"] > div {
    display: flex !important;
    gap: 0.6rem !important;
    flex-direction: row !important;
    position: relative; z-index: 1;
}
[data-testid="stRadio"] label {
    background: rgba(255,255,255,0.06) !important;
    border: 1px solid rgba(255,255,255,0.15) !important;
    border-radius: var(--radius-sm) !important;
    padding: 0.65rem 1rem !important;
    font-family: var(--font-body) !important;
    font-size: 0.88rem !important;
    color: var(--brown-pale) !important;
    cursor: pointer !important;
    transition: all 0.2s !important;
    flex: 1 !important;
    text-align: center !important;
}
[data-testid="stRadio"] label:hover {
    border-color: var(--sienna-light) !important;
    color: var(--cream) !important;
    background: rgba(255,255,255,0.1) !important;
}
[data-testid="stRadio"] label:has(input:checked) {
    background: var(--sienna) !important;
    border-color: var(--sienna) !important;
    color: var(--cream) !important;
    font-weight: 600 !important;
}

/* Checkbox */
[data-testid="stCheckbox"] {
    font-family: var(--font-body) !important;
    font-size: 0.88rem !important;
    color: var(--brown-mid) !important;
}

/* File uploader */
[data-testid="stFileUploader"] section {
    background: rgba(255,255,255,0.5) !important;
    border: 1.5px dashed rgba(59,26,9,0.2) !important;
    border-radius: var(--radius-md) !important;
    padding: 0.6rem 0.85rem !important;
    transition: all 0.2s !important;
}
[data-testid="stFileUploader"] section:hover {
    border-color: var(--sienna) !important;
    background: rgba(255,255,255,0.75) !important;
}
[data-testid="stFileUploader"] section p,
[data-testid="stFileUploader"] section span,
[data-testid="stFileUploader"] section small {
    font-family: var(--font-body) !important;
    font-size: 0.8rem !important;
    color: var(--brown-light) !important;
}
[data-testid="stFileUploader"] section button {
    background: var(--cream-dark) !important;
    border: 1px solid rgba(59,26,9,0.2) !important;
    color: var(--brown-mid) !important;
    font-family: var(--font-body) !important;
    font-size: 0.78rem !important;
    border-radius: var(--radius-sm) !important;
}

/* Buttons */
.stButton > button {
    background: var(--cream) !important;
    border: 1px solid rgba(59,26,9,0.2) !important;
    color: var(--espresso) !important;
    font-family: var(--font-body) !important;
    font-size: 0.82rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.06em !important;
    text-transform: uppercase !important;
    border-radius: var(--radius-pill) !important;
    padding: 0.65rem 1.5rem !important;
    transition: all 0.2s !important;
    box-shadow: var(--shadow-sm) !important;
}
.stButton > button:hover {
    background: var(--espresso) !important;
    color: var(--cream) !important;
    border-color: var(--espresso) !important;
    transform: translateY(-1px) !important;
    box-shadow: var(--shadow-md) !important;
}
.exec-btn .stButton > button {
    background: var(--sienna) !important;
    border-color: var(--sienna) !important;
    color: var(--cream) !important;
    font-size: 0.88rem !important;
    padding: 0.8rem 2rem !important;
    box-shadow: 0 6px 24px rgba(196,98,42,0.4) !important;
    width: 100% !important;
}
.exec-btn .stButton > button:hover {
    background: var(--sienna-light) !important;
    border-color: var(--sienna-light) !important;
    box-shadow: 0 8px 32px rgba(196,98,42,0.5) !important;
}

/* Download button */
.stDownloadButton > button {
    background: var(--frame-green) !important;
    border-color: var(--frame-green) !important;
    color: white !important;
    font-family: var(--font-body) !important;
    font-size: 0.82rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.06em !important;
    text-transform: uppercase !important;
    border-radius: var(--radius-pill) !important;
    padding: 0.7rem 1.5rem !important;
    transition: all 0.2s !important;
    width: 100% !important;
}
.stDownloadButton > button:hover {
    background: #4A8A58 !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 16px rgba(91,158,106,0.35) !important;
}

/* Progress */
[data-testid="stProgress"] > div {
    background: rgba(255,255,255,0.12) !important;
    border-radius: 3px !important;
    height: 4px !important;
}
[data-testid="stProgress"] > div > div {
    background: linear-gradient(90deg, var(--sienna), var(--sienna-light), var(--frame-yellow)) !important;
    border-radius: 3px !important;
    box-shadow: 0 0 10px rgba(196,98,42,0.4) !important;
}

/* Native alerts */
[data-testid="stAlert"] {
    font-family: var(--font-body) !important;
    font-size: 0.85rem !important;
    border-radius: var(--radius-md) !important;
}

/* Expander */
.streamlit-expanderHeader {
    font-family: var(--font-body) !important;
    font-size: 0.85rem !important;
    color: var(--brown-mid) !important;
    background: var(--cream-dark) !important;
}

/* ══════════════════════════════════
   HISTORY STRIP
══════════════════════════════════ */
.history-strip {
    display: flex;
    flex-direction: column;
    gap: 0.55rem;
    margin-top: 1.5rem;
    position: relative; z-index: 1;
}
.hist-pill {
    display: flex;
    align-items: center;
    justify-content: space-between;
    background: rgba(255,255,255,0.07);
    border: 1px solid rgba(255,255,255,0.1);
    border-radius: var(--radius-pill);
    padding: 0.55rem 1rem;
}
.hp-left {
    font-family: var(--font-mono);
    font-size: 0.7rem;
    color: var(--brown-pale);
}
.hp-time { font-weight: 700; color: var(--cream); margin-right: 0.5rem; }
.hp-badge {
    font-family: var(--font-body);
    font-size: 0.62rem;
    font-weight: 700;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    padding: 0.2rem 0.6rem;
    border-radius: var(--radius-pill);
}
.hp-ok  { color: #8EC49A; background: rgba(91,158,106,0.2); }
.hp-err { color: #E0A880; background: rgba(196,98,42,0.2); }

/* ══════════════════════════════════
   RESULTS SECTION
══════════════════════════════════ */
.results-banner {
    background: var(--cream-dark);
    border: 1px solid rgba(59,26,9,0.12);
    border-radius: var(--radius-lg);
    padding: 2rem 2.5rem;
    margin: 2rem 2.5rem 0;
    box-shadow: var(--shadow-md);
    animation: fadeInUp 0.4s ease both;
    position: relative;
    overflow: hidden;
}
.results-banner::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 4px;
    background: linear-gradient(90deg, var(--frame-green), var(--frame-yellow), var(--frame-orange));
}
.rb-script {
    font-family: var(--font-script);
    font-size: 1.4rem;
    color: var(--sienna);
    display: block;
    margin-bottom: 0.1rem;
}
.rb-title {
    font-family: var(--font-display);
    font-size: 2.2rem;
    color: var(--espresso);
    line-height: 1.1;
    margin-bottom: 1.5rem;
}
.metrics-row {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 1rem;
}
.metric-tile {
    background: var(--cream);
    border: 1px solid rgba(59,26,9,0.1);
    border-radius: var(--radius-md);
    padding: 1.25rem;
    text-align: center;
    box-shadow: var(--shadow-sm);
}
.mt-num {
    font-family: var(--font-display);
    font-size: 3rem;
    color: var(--sienna);
    line-height: 1;
    margin-bottom: 0.35rem;
}
.mt-label {
    font-family: var(--font-body);
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: var(--brown-light);
}

/* ══════════════════════════════════
   PROGRESS STATUS
══════════════════════════════════ */
.prog-status {
    font-family: var(--font-mono);
    font-size: 0.78rem;
    color: var(--sienna-light);
    text-align: center;
    padding: 0.4rem;
    letter-spacing: 0.08em;
}
.prog-detail {
    display: flex;
    justify-content: space-between;
    font-family: var(--font-mono);
    font-size: 0.68rem;
    color: var(--brown-light);
    margin-top: 0.3rem;
}

/* SETTINGS PANEL */
.settings-wrap {
    background: var(--cream-dark);
    border: 1px solid rgba(59,26,9,0.12);
    border-radius: var(--radius-lg);
    padding: 1.5rem 2rem;
    margin: 0 2.5rem 1.5rem;
    box-shadow: var(--shadow-sm);
}

/* DOWNLOAD CARD */
.dl-card {
    background: var(--cream-dark);
    border: 1px solid rgba(59,26,9,0.12);
    border-radius: var(--radius-lg);
    padding: 1.5rem;
    box-shadow: var(--shadow-sm);
}
.dl-title {
    font-family: var(--font-serif);
    font-style: italic;
    font-size: 1.2rem;
    color: var(--espresso);
    margin-bottom: 1rem;
}

/* RESPONSIVE */
@media (max-width: 900px) {
    .cards-grid { grid-template-columns: 1fr; padding: 0 1.25rem; }
    .hv-nav { padding: 0 1.25rem; }
    .hv-hero { padding: 2.5rem 1.25rem 2rem; }
    .hv-section-row, .config-wrap { padding: 0 1.25rem; }
    .metrics-row { grid-template-columns: repeat(3, 1fr); }
    .results-banner { margin: 1.5rem 1.25rem 0; }
    .settings-wrap { margin: 0 1.25rem 1.5rem; }
}
@media (max-width: 640px) {
    .hero-title { font-size: 2.4rem; }
    .metrics-row { grid-template-columns: 1fr; }
    .mt-num { font-size: 2.4rem; }
    [data-testid="stRadio"] > div { flex-direction: column !important; }
}
"""

st.markdown(f"<style>{THEME}</style>", unsafe_allow_html=True)

# ═══════════ HELPER ═══════════
def validate_file(file, file_type):
    try:
        wb = load_workbook(file)
        ws = wb.active
        rows = ws.max_row; cols = ws.max_column
        header_row = 7 if file_type in ["job", "einv"] else 3
        headers = []
        for c in range(1, min(cols + 1, 20)):
            v = ws.cell(header_row, c).value
            if v: headers.append(str(v).strip())
        return {"valid": True, "rows": rows, "cols": cols, "headers": headers,
                "sheets": wb.sheetnames if file_type == "bill" else None}
    except Exception as e:
        return {"valid": False, "error": str(e)}

MONTHS = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"]

# ═══════════ NAV ═══════════
st.markdown("""
<div class="hv-nav">
    <div class="nav-left">
        <span class="nav-link">Documents</span>
        <span class="nav-link">History</span>
        <span class="nav-link">Settings</span>
    </div>
    <div class="nav-wordmark">
        <span class="nav-wordmark-icon">✦</span>
        BILLFLOW
    </div>
    <div class="nav-right">
        <span class="nav-link">Reports</span>
        <span class="nav-link">Export</span>
        <span class="nav-pill">Process Now</span>
    </div>
</div>
""", unsafe_allow_html=True)

# ═══════════ HERO ═══════════
st.markdown("""
<div class="hv-hero">
    <span class="hero-script">Seamless & Smart</span>
    <h1 class="hero-title">AUTOMATE YOUR<br><em>Bill Reports</em></h1>
    <p class="hero-sub">Upload · Reconcile · Export &nbsp;—&nbsp; in seconds</p>
    <div class="hero-cta-row">
        <div class="hero-btn">
            Get Started
            <span class="hero-btn-arrow">↗</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# Settings toggle
sc1, sc2 = st.columns([8, 1])
with sc2:
    st.markdown("<div style='padding:1rem 2.5rem 0 0; text-align:right;'>", unsafe_allow_html=True)
    if st.button("⚙ Settings" if not st.session_state.show_settings else "✕ Close", key="stog"):
        st.session_state.show_settings = not st.session_state.show_settings
    st.markdown("</div>", unsafe_allow_html=True)

# Settings panel
if st.session_state.show_settings:
    st.markdown('<div class="settings-wrap">', unsafe_allow_html=True)
    sa, sb = st.columns(2, gap="large")
    with sa:
        st.markdown('<span style="font-family:var(--font-serif);font-style:italic;font-size:1rem;color:var(--espresso);">File Preview</span>', unsafe_allow_html=True)
        sp = st.checkbox("Show file statistics after upload", value=st.session_state.show_preview, key="prev_chk")
        st.session_state.show_preview = sp
    with sb:
        st.markdown('<span style="font-family:var(--font-serif);font-style:italic;font-size:1rem;color:var(--espresso);">Processing History</span>', unsafe_allow_html=True)
        if st.button("Clear History", key="clr"):
            st.session_state.processing_history = []
            st.success("History cleared.")
    st.markdown('</div>', unsafe_allow_html=True)

# ═══════════ SECTION TITLE ═══════════
st.markdown("""
<div class="hv-section-row">
    <span class="hv-section-title">Documents &amp; intake</span>
    <div class="hv-section-rule"></div>
    <span class="hv-section-tag">Step 01</span>
</div>
""", unsafe_allow_html=True)

# ═══════════ UPLOAD CARDS ═══════════
st.markdown('<div class="cards-grid">', unsafe_allow_html=True)

# Card 1 — Bill Report (green frame)
st.markdown("""
<div class="upload-card card-green">
    <div class="card-frame-top"></div>
    <div class="card-body">
        <div class="card-badge">01 &nbsp;·&nbsp; Primary Register</div>
        <div class="card-title">Bill Report</div>
        <div class="card-desc">Main billing workbook with monthly tabs — the master ledger into which data is written.</div>
    </div>
</div>
""", unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)  # close grid temporarily for file uploaders

# Re-open with actual columns for uploaders
col1, col2, col3 = st.columns(3, gap="medium")

with col1:
    st.markdown("""
    <div style="padding: 0 2.5rem 0 2.5rem; margin-top:-2rem;">
        <div class="upload-card card-green" style="padding:0;">
            <div class="card-frame-top"></div>
            <div class="card-body">
                <div class="card-badge">01 &nbsp;·&nbsp; Primary Register</div>
                <div class="card-title">Bill Report</div>
                <div class="card-desc">Main billing workbook with monthly tabs — the master ledger.</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    bill_file = st.file_uploader("", type="xlsx", key="bill", label_visibility="collapsed")
    if bill_file:
        st.markdown('<div class="card-loaded" style="color:#5B9E6A;background:rgba(91,158,106,0.1);border:1px solid rgba(91,158,106,0.25);display:inline-flex;align-items:center;gap:0.4rem;font-family:var(--font-body);font-size:0.72rem;font-weight:600;letter-spacing:0.08em;padding:0.3rem 0.8rem;border-radius:50px;">✓ Document Loaded</div>', unsafe_allow_html=True)
        if st.session_state.show_preview:
            v = validate_file(bill_file, "bill")
            if v["valid"]:
                st.markdown(f'<div class="preview-mini"><div class="pm-row"><span>Rows</span><span class="pm-val">{v["rows"]}</span></div><div class="pm-row"><span>Columns</span><span class="pm-val">{v["cols"]}</span></div><div class="pm-row"><span>Sheets</span><span class="pm-val">{len(v["sheets"])}</span></div></div>', unsafe_allow_html=True)
            bill_file.seek(0)

with col2:
    st.markdown("""
    <div style="padding: 0;">
        <div class="upload-card card-orange" style="padding:0;">
            <div class="card-frame-top"></div>
            <div class="card-body">
                <div class="card-badge">02 &nbsp;·&nbsp; Job Register</div>
                <div class="card-title">Job Report</div>
                <div class="card-desc">Shipper details, container numbers, invoice refs and SB/BE tracking data.</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    job_file = st.file_uploader("", type="xlsx", key="job", label_visibility="collapsed")
    if job_file:
        st.markdown('<div style="color:#D4733A;background:rgba(212,115,58,0.1);border:1px solid rgba(212,115,58,0.25);display:inline-flex;align-items:center;gap:0.4rem;font-family:var(--font-body);font-size:0.72rem;font-weight:600;letter-spacing:0.08em;padding:0.3rem 0.8rem;border-radius:50px;">✓ Document Loaded</div>', unsafe_allow_html=True)
        if st.session_state.show_preview:
            v = validate_file(job_file, "job")
            if v["valid"]:
                st.markdown(f'<div class="preview-mini"><div class="pm-row"><span>Rows</span><span class="pm-val">{v["rows"]}</span></div><div class="pm-row"><span>Columns</span><span class="pm-val">{v["cols"]}</span></div></div>', unsafe_allow_html=True)
            job_file.seek(0)

with col3:
    st.markdown("""
    <div style="padding: 0;">
        <div class="upload-card card-yellow" style="padding:0;">
            <div class="card-frame-top"></div>
            <div class="card-body">
                <div class="card-badge">03 &nbsp;·&nbsp; Invoice Data</div>
                <div class="card-title">E-Invoice Report</div>
                <div class="card-desc">Electronic invoice records — bill numbers and issuance dates by job.</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    einv_file = st.file_uploader("", type="xlsx", key="einv", label_visibility="collapsed")
    if einv_file:
        st.markdown('<div style="color:#C8A82A;background:rgba(200,168,42,0.1);border:1px solid rgba(200,168,42,0.25);display:inline-flex;align-items:center;gap:0.4rem;font-family:var(--font-body);font-size:0.72rem;font-weight:600;letter-spacing:0.08em;padding:0.3rem 0.8rem;border-radius:50px;">✓ Document Loaded</div>', unsafe_allow_html=True)
        if st.session_state.show_preview:
            v = validate_file(einv_file, "einv")
            if v["valid"]:
                st.markdown(f'<div class="preview-mini"><div class="pm-row"><span>Rows</span><span class="pm-val">{v["rows"]}</span></div><div class="pm-row"><span>Columns</span><span class="pm-val">{v["cols"]}</span></div></div>', unsafe_allow_html=True)
            einv_file.seek(0)

# ═══════════ CONFIG SECTION ═══════════
st.markdown("""
<div class="hv-section-row" style="margin-top:2.5rem;">
    <span class="hv-section-title">Configure &amp; run</span>
    <div class="hv-section-rule"></div>
    <span class="hv-section-tag">Step 02</span>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="config-wrap">', unsafe_allow_html=True)
cfg_left, cfg_right = st.columns([3, 1], gap="large")

with cfg_left:
    st.markdown('<div class="config-card">', unsafe_allow_html=True)
    st.markdown('<span class="config-script">Ready to process</span>', unsafe_allow_html=True)
    st.markdown('<div class="config-title">Set your<br>parameters</div>', unsafe_allow_html=True)

    cl1, cl2 = st.columns(2, gap="large")
    with cl1:
        st.markdown('<span class="config-label">Billing Month</span>', unsafe_allow_html=True)
        month = st.selectbox("", MONTHS, index=st.session_state.selected_month,
                             label_visibility="collapsed", key="month_sel")
        st.session_state.selected_month = MONTHS.index(month)
    with cl2:
        st.markdown('<span class="config-label">Update Mode</span>', unsafe_allow_html=True)
        update_mode = st.radio("", ["Append", "Overwrite"],
                               label_visibility="collapsed", key="umode")

    if update_mode == "Append":
        st.markdown('<div class="mode-note mode-note-green">New job entries will be appended — existing records remain untouched.</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="mode-note mode-note-warn">All existing entries for this month will be cleared and replaced.</div>', unsafe_allow_html=True)

    st.markdown('<div class="config-divider"></div>', unsafe_allow_html=True)
    st.markdown('<div class="exec-btn">', unsafe_allow_html=True)
    process_clicked = st.button("✦  Execute Report Generation", key="proc_main", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

with cfg_right:
    if st.session_state.processing_history:
        st.markdown("""
        <div style="background:var(--espresso);border-radius:var(--radius-lg);padding:1.5rem 1.25rem;box-shadow:var(--shadow-lg);">
            <span style="font-family:var(--font-script);font-size:1.1rem;color:var(--sienna-light);display:block;margin-bottom:0.2rem;">Recent</span>
            <div style="font-family:var(--font-display);font-size:1.3rem;color:var(--cream);margin-bottom:1rem;">Activity</div>
        """, unsafe_allow_html=True)
        for entry in reversed(st.session_state.processing_history[-4:]):
            bc = "hp-ok" if entry["status"] == "success" else "hp-err"
            bm = "OK" if entry["status"] == "success" else "ERR"
            st.markdown(f"""
            <div class="hist-pill">
                <div class="hp-left">
                    <span class="hp-time">{entry["time"]}</span>
                    {entry["month"]} · +{entry["added"]}
                </div>
                <div class="hp-badge {bc}">{bm}</div>
            </div>
            """, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# ═══════════ PROCESSING CONSTANTS ═══════════
BILL_HEADER_ROW = 3; BILL_DATA_START = 4
SRC_HEADER_ROW  = 7; SRC_DATA_START  = 8

def norm(x): return str(x).strip().upper() if x else ""
def clean_job(job): return norm(job).split("/")[0]

def parse_date(date_val):
    if not date_val: return None
    if isinstance(date_val, datetime): return date_val.date()
    ds = str(date_val).strip()
    if not ds or ds.upper() == 'NONE': return None
    for fmt in ['%Y-%m-%d','%d-%m-%Y','%d/%m/%Y','%m/%d/%Y','%d.%m.%Y',
                '%Y/%m/%d','%d-%b-%Y','%d %b %Y','%d-%m-%y','%d/%m/%y']:
        try: return datetime.strptime(ds, fmt).date()
        except: continue
    return None

def format_date_for_excel(d):
    if not d: return None
    if isinstance(d, datetime): return d.strftime('%d/%m/%y')
    try: return d.strftime('%d/%m/%y')
    except: return None

def col_map(ws, header_row):
    m = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v: m[norm(v)] = c
    return m

def get_existing_jobs(ws, bc, start_row):
    ex = set()
    for r in range(start_row, ws.max_row + 1):
        j = ws.cell(r, bc["JOB NO"]).value
        if j: ex.add(clean_job(j))
    return ex

def find_next_empty_row(ws, start_row):
    for r in range(start_row, ws.max_row + 100):
        if not ws.cell(r, 1).value and not ws.cell(r, 2).value: return r
    return ws.max_row + 1

# ═══════════ PROCESSING ═══════════
if process_clicked:
    if not (bill_file and job_file and einv_file):
        st.error("Please upload all three documents before running.")
    else:
        pb = st.progress(0); se = st.empty(); de = st.empty()
        t0 = datetime.now()

        def upd(p, msg, det=""):
            pb.progress(p)
            se.markdown(f'<div class="prog-status">{msg}</div>', unsafe_allow_html=True)
            if det:
                de.markdown(f'<div class="prog-detail"><span>{det}</span><span>{p}%</span></div>', unsafe_allow_html=True)

        try:
            upd(10, "Loading workbooks…", "Reading Excel files")
            bill_wb = load_workbook(bill_file)
            job_wb  = load_workbook(job_file)
            einv_wb = load_workbook(einv_file)

            upd(25, "Locating sheets…", "Checking month tab")
            if month not in bill_wb.sheetnames:
                st.error(f"Sheet '{month}' not found in Bill Report."); st.stop()

            bill_ws = bill_wb[month]; job_ws = job_wb.active; einv_ws = einv_wb.active

            upd(35, "Mapping columns…", "Analysing headers")
            bill_cols = col_map(bill_ws, BILL_HEADER_ROW)
            job_cols  = col_map(job_ws,  SRC_HEADER_ROW)
            einv_cols = col_map(einv_ws, SRC_HEADER_ROW)

            upd(45, "Validating…", "Checking required fields")
            for col in ["JOB NO","JOB DATE","SHIPPER NAME","INVOICE NO","SB / BE NO","CONTAINER NO"]:
                if col not in job_cols: st.error(f"Missing in Job Report: {col}"); st.stop()
            for col in ["JOB NO","BILL NO","BILL DATE"]:
                if col not in einv_cols: st.error(f"Missing in E-Invoice: {col}"); st.stop()
            for col in ["JOB NO","JOB DATE","PARTY NAME","INVOICE NO","SB / BE NO","CONTAINER NO","BILL NO","BILL DATE"]:
                if col not in bill_cols: st.error(f"Missing in Bill Report: {col}"); st.stop()

            append_mode = update_mode == "Append"
            existing_jobs = set()

            if append_mode:
                upd(55, "Scanning existing entries…", "Checking duplicates")
                existing_jobs = get_existing_jobs(bill_ws, bill_cols, BILL_DATA_START)
                row = find_next_empty_row(bill_ws, BILL_DATA_START)
                last_sr = 0
                for r in range(BILL_DATA_START, row):
                    sv = bill_ws.cell(r, 1).value
                    if sv and isinstance(sv, (int, float)): last_sr = max(last_sr, int(sv))
                sr = last_sr + 1
            else:
                upd(55, "Clearing sheet…", "Preparing fresh data")
                for r in range(BILL_DATA_START, bill_ws.max_row + 1):
                    for c in range(1, bill_ws.max_column + 1): bill_ws.cell(r, c).value = None
                row = BILL_DATA_START; sr = 1

            upd(65, "Reading job register…", "Parsing entries")
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

            upd(78, "Reading invoice register…", "Parsing bills")
            bill_map = defaultdict(list)
            for r in range(SRC_DATA_START, einv_ws.max_row + 1):
                raw = einv_ws.cell(r, einv_cols["JOB NO"]).value
                bno = einv_ws.cell(r, einv_cols["BILL NO"]).value
                bd  = parse_date(einv_ws.cell(r, einv_cols["BILL DATE"]).value)
                if raw and bno: bill_map[clean_job(raw)].append((str(bno), bd))

            upd(88, "Writing report…", "Generating output")
            added = 0; skipped = 0

            for jn, data in jobs.items():
                if append_mode and jn in existing_jobs:
                    skipped += 1; continue
                bill_ws.cell(row, 1).value = sr
                bill_ws.cell(row, bill_cols["JOB NO"]).value = jn
                for field in ["JOB DATE","PARTY NAME","INVOICE NO","SB / BE NO","CONTAINER NO"]:
                    val = data[field]
                    if field == "JOB DATE": val = format_date_for_excel(val)
                    bill_ws.cell(row, bill_cols[field]).value = val
                if jn in bill_map:
                    bills = bill_map[jn]
                    if len(bills) == 1:
                        merged = bills[0][0]
                    else:
                        fb = str(bills[0][0])
                        main, suf = (fb.rsplit("/",1)[0], "/"+fb.rsplit("/",1)[1]) if "/" in fb else (fb, "")
                        extra = []
                        for i in range(1, len(bills)):
                            bn = str(bills[i][0])
                            if "/" in bn: bn = bn.split("/")[0]
                            extra.append(bn[-4:] if len(bn) >= 4 else bn)
                        merged = main + "/" + "/".join(extra) + suf
                    vd = [b[1] for b in bills if b[1]]
                    ld = max(vd) if vd else None
                    bill_ws.cell(row, bill_cols["BILL NO"]).value   = merged
                    bill_ws.cell(row, bill_cols["BILL DATE"]).value = format_date_for_excel(ld)
                row += 1; sr += 1; added += 1

            upd(96, "Saving file…", "Writing to disk")
            out = f"UPDATED_BILL_REPORT_{month}.xlsx"
            bill_wb.save(out)
            pb.progress(100)
            t1 = datetime.now(); dur = (t1 - t0).total_seconds()
            se.empty(); de.empty(); pb.empty()

            st.session_state.processing_history.append({
                "time": t1.strftime("%H:%M:%S"), "month": month,
                "status": "success", "added": added, "skipped": skipped,
                "duration": f"{dur:.1f}s"
            })

            # Results
            st.markdown(f"""
            <div class="results-banner">
                <span class="rb-script">All done!</span>
                <div class="rb-title">REPORT<br>GENERATED</div>
                <div class="metrics-row">
                    <div class="metric-tile">
                        <div class="mt-num">{added}</div>
                        <div class="mt-label">Entries Added</div>
                    </div>
                    <div class="metric-tile">
                        <div class="mt-num">{skipped}</div>
                        <div class="mt-label">Skipped</div>
                    </div>
                    <div class="metric-tile">
                        <div class="mt-num">{dur:.1f}s</div>
                        <div class="mt-label">Duration</div>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            _, dl_col, _ = st.columns([2, 1, 2])
            with dl_col:
                st.markdown('<div style="margin-top:1rem;padding:0 2.5rem;">', unsafe_allow_html=True)
                with open(out, "rb") as f:
                    st.download_button("↓ Download Excel", data=f, file_name=out,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True, key="dl_xl")
                try:
                    df = pd.read_excel(out, sheet_name=month, header=BILL_HEADER_ROW - 1)
                    csv = df.to_csv(index=False)
                    st.download_button("↓ Download CSV", data=csv,
                        file_name=f"UPDATED_BILL_REPORT_{month}.csv",
                        mime="text/csv", use_container_width=True, key="dl_csv")
                except: pass
                st.markdown('</div>', unsafe_allow_html=True)

        except Exception as e:
            pb.empty(); se.empty(); de.empty()
            st.session_state.processing_history.append({
                "time": datetime.now().strftime("%H:%M:%S"), "month": month,
                "status": "error", "added": 0, "skipped": 0, "duration": "—"
            })
            st.error(f"Processing error: {str(e)}")
            with st.expander("View error details"):
                st.exception(e)