import streamlit as st
from openpyxl import load_workbook
from collections import defaultdict
from datetime import datetime
import pandas as pd

# ================= PAGE CONFIG =================
st.set_page_config(
    page_title="Ledger — Bill Report Automation",
    page_icon="📜",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ================= SESSION STATE =================
if 'selected_month' not in st.session_state:
    st.session_state.selected_month = datetime.now().month - 1
if 'processing_history' not in st.session_state:
    st.session_state.processing_history = []
if 'show_preview' not in st.session_state:
    st.session_state.show_preview = False
if 'show_settings' not in st.session_state:
    st.session_state.show_settings = False

# ================= THEME =================
THEME = """
@import url('https://fonts.googleapis.com/css2?family=Cormorant+Garant:ital,wght@0,300;0,400;0,500;0,600;0,700;1,300;1,400;1,600&family=IM+Fell+English:ital@0;1&family=Courier+Prime:ital,wght@0,400;0,700;1,400&family=Crimson+Pro:ital,wght@0,300;0,400;0,500;0,600;1,400&display=swap');

:root {
    /* Parchment palette */
    --parch-0:       #FDFAF3;
    --parch-1:       #F7EDD8;
    --parch-2:       #EFE0BE;
    --parch-3:       #E4D0A2;
    --parch-4:       #D4BB84;

    /* Ink shades */
    --ink-0:         #2A1C08;
    --ink-1:         #3D2B12;
    --ink-2:         #5C4024;
    --ink-3:         #7A5B38;
    --ink-4:         #9B7B58;
    --ink-5:         #BDA07A;

    /* Accent: antique crimson */
    --accent:        #8B3A2A;
    --accent-soft:   #A04A38;
    --accent-pale:   rgba(139,58,42,0.1);
    --accent-border: rgba(139,58,42,0.3);

    /* Verdigris (aged copper green) */
    --verdant:       #4A7B6A;
    --verdant-pale:  rgba(74,123,106,0.12);
    --verdant-border:rgba(74,123,106,0.3);

    /* Aged gold */
    --gilt:          #9A6F2A;
    --gilt-soft:     #C49040;

    /* Borders */
    --rule-light:    rgba(139,110,70,0.18);
    --rule-mid:      rgba(139,110,70,0.32);
    --rule-strong:   rgba(139,110,70,0.55);

    /* Shadows */
    --shadow-sm:     0 1px 4px rgba(42,28,8,0.08);
    --shadow-md:     0 3px 12px rgba(42,28,8,0.12);
    --shadow-lg:     0 6px 28px rgba(42,28,8,0.16);

    /* Type */
    --font-display:  'Cormorant Garant', Georgia, serif;
    --font-body:     'Crimson Pro', Georgia, serif;
    --font-fell:     'IM Fell English', Georgia, serif;
    --font-mono:     'Courier Prime', 'Courier New', monospace;

    --radius-sm:  4px;
    --radius-md:  8px;
    --radius-lg:  14px;
    --radius-xl:  20px;
}

/* ─── GLOBAL RESET ─── */
*, *::before, *::after { box-sizing: border-box; }

html, body,
[data-testid="stAppViewContainer"],
[data-testid="stAppViewBlockContainer"],
.main .block-container {
    font-family: var(--font-body) !important;
    background: var(--parch-1) !important;
    color: var(--ink-0) !important;
}

.main { background: var(--parch-1) !important; }
.block-container {
    max-width: 1380px !important;
    padding: 2rem 2.5rem !important;
}

/* ─── PARCHMENT GRAIN OVERLAY ─── */
.stApp::before {
    content: '';
    position: fixed; inset: 0;
    background-image:
        url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='400' height='400'%3E%3Cfilter id='noise'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.9' numOctaves='4' stitchTiles='stitch'/%3E%3CfeColorMatrix type='saturate' values='0'/%3E%3C/filter%3E%3Crect width='400' height='400' filter='url(%23noise)' opacity='0.04'/%3E%3C/svg%3E");
    pointer-events: none;
    z-index: 0;
}

/* Soft vignette */
.stApp::after {
    content: '';
    position: fixed; inset: 0;
    background: radial-gradient(ellipse at center, transparent 60%, rgba(42,28,8,0.07) 100%);
    pointer-events: none;
    z-index: 0;
}

/* ─── HIDE STREAMLIT CHROME ─── */
#MainMenu, footer, .stDeployButton,
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"] {
    visibility: hidden !important;
    display: none !important;
}

/* ─── SCROLLBAR ─── */
::-webkit-scrollbar { width: 5px; }
::-webkit-scrollbar-track { background: var(--parch-2); }
::-webkit-scrollbar-thumb { background: var(--parch-4); border-radius: 3px; }

/* ─── HEADER ─── */
.ledger-header {
    text-align: center;
    padding: 2.5rem 2rem 1.75rem;
    position: relative;
    margin-bottom: 2rem;
}
.lh-ornament {
    font-family: var(--font-display);
    color: var(--gilt);
    font-size: 1.1rem;
    letter-spacing: 0.3em;
    margin-bottom: 0.5rem;
    display: block;
}
.lh-double-rule {
    height: 3px;
    border-top: 1px solid var(--rule-strong);
    border-bottom: 1px solid var(--rule-mid);
    margin: 0.75rem auto;
    width: 100%;
}
.lh-title {
    font-family: var(--font-display);
    font-size: 3.4rem;
    font-weight: 600;
    color: var(--ink-0);
    letter-spacing: -0.01em;
    line-height: 1.05;
    margin-bottom: 0.2rem;
}
.lh-title em {
    font-style: italic;
    color: var(--accent);
}
.lh-subtitle {
    font-family: var(--font-fell);
    font-style: italic;
    font-size: 1.05rem;
    color: var(--ink-3);
    letter-spacing: 0.04em;
    margin-bottom: 0.75rem;
}
.lh-meta {
    font-family: var(--font-mono);
    font-size: 0.72rem;
    color: var(--ink-4);
    letter-spacing: 0.16em;
    text-transform: uppercase;
}

/* ─── SECTION HEADING ─── */
.sect-heading {
    font-family: var(--font-display);
    font-size: 1.35rem;
    font-weight: 500;
    color: var(--ink-1);
    display: flex;
    align-items: center;
    gap: 0.75rem;
    margin-bottom: 1rem;
    letter-spacing: 0.01em;
}
.sect-heading::after {
    content: '';
    flex: 1;
    height: 1px;
    background: var(--rule-mid);
}
.sect-heading span {
    font-family: var(--font-mono);
    font-size: 0.65rem;
    color: var(--ink-4);
    letter-spacing: 0.18em;
    text-transform: uppercase;
    border: 1px solid var(--rule-mid);
    padding: 0.2rem 0.55rem;
    border-radius: 3px;
    background: var(--parch-0);
}

/* ─── CARD ─── */
.v-card {
    background: var(--parch-0);
    border: 1px solid var(--rule-mid);
    border-radius: var(--radius-lg);
    padding: 1.75rem;
    box-shadow: var(--shadow-md);
    position: relative;
    overflow: hidden;
}
.v-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0; height: 3px;
    background: linear-gradient(90deg,
        transparent 0%,
        var(--rule-mid) 15%,
        var(--gilt) 50%,
        var(--rule-mid) 85%,
        transparent 100%
    );
    opacity: 0.6;
}

/* ─── ORNAMENTAL DIVIDER ─── */
.orn-divider {
    text-align: center;
    color: var(--ink-5);
    font-size: 1rem;
    margin: 1.1rem 0;
    display: flex;
    align-items: center;
    gap: 0.6rem;
}
.orn-divider::before, .orn-divider::after {
    content: '';
    flex: 1;
    height: 1px;
    background: var(--rule-light);
}

/* ─── UPLOAD ITEMS ─── */
.upload-entry {
    background: var(--parch-1);
    border: 1px solid var(--rule-light);
    border-radius: var(--radius-md);
    padding: 1.15rem 1.25rem 0.9rem;
    margin-bottom: 0.85rem;
    transition: border-color 0.25s, box-shadow 0.25s, background 0.25s;
    position: relative;
}
.upload-entry:last-of-type { margin-bottom: 0; }
.upload-entry:hover {
    border-color: var(--rule-strong);
    background: var(--parch-0);
    box-shadow: var(--shadow-sm);
}

.ue-row {
    display: flex;
    align-items: baseline;
    gap: 0.85rem;
    margin-bottom: 0.2rem;
}
.ue-numeral {
    font-family: var(--font-mono);
    font-size: 0.68rem;
    color: var(--ink-5);
    background: var(--parch-2);
    border: 1px solid var(--rule-light);
    padding: 0.1rem 0.45rem;
    border-radius: 3px;
    letter-spacing: 0.05em;
    flex-shrink: 0;
}
.ue-title {
    font-family: var(--font-display);
    font-size: 1.15rem;
    font-weight: 600;
    color: var(--ink-1);
}
.ue-desc {
    font-family: var(--font-body);
    font-size: 0.88rem;
    color: var(--ink-4);
    font-style: italic;
    padding-left: 2.8rem;
    margin-bottom: 0.6rem;
    line-height: 1.4;
}
.file-loaded {
    display: inline-flex;
    align-items: center;
    gap: 0.4rem;
    font-family: var(--font-mono);
    font-size: 0.68rem;
    color: var(--verdant);
    background: var(--verdant-pale);
    border: 1px solid var(--verdant-border);
    padding: 0.25rem 0.7rem;
    border-radius: 3px;
    margin-left: 2.8rem;
    margin-top: 0.35rem;
    letter-spacing: 0.08em;
    text-transform: uppercase;
}

/* ─── PREVIEW TABLE ─── */
.preview-ledger {
    background: var(--parch-1);
    border: 1px solid var(--rule-light);
    border-radius: var(--radius-sm);
    margin-top: 0.55rem;
    margin-left: 2.8rem;
    overflow: hidden;
}
.pl-row {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 0.35rem 0.85rem;
    border-bottom: 1px solid var(--rule-light);
    font-family: var(--font-mono);
    font-size: 0.71rem;
    color: var(--ink-3);
}
.pl-row:last-child { border-bottom: none; }
.pl-val { font-weight: 700; color: var(--accent); }

/* ─── CONFIG LABELS ─── */
.conf-label {
    font-family: var(--font-fell);
    font-size: 1rem;
    font-style: italic;
    color: var(--ink-2);
    display: block;
    margin-bottom: 0.45rem;
    letter-spacing: 0.02em;
}

/* ─── ALERTS ─── */
.v-alert-info {
    background: rgba(74,123,106,0.06);
    border: 1px solid var(--verdant-border);
    border-left: 3px solid var(--verdant);
    border-radius: var(--radius-sm);
    padding: 0.6rem 0.9rem;
    font-family: var(--font-body);
    font-style: italic;
    font-size: 0.88rem;
    color: var(--verdant);
    margin-top: 0.5rem;
    line-height: 1.4;
}
.v-alert-warn {
    background: var(--accent-pale);
    border: 1px solid var(--accent-border);
    border-left: 3px solid var(--accent);
    border-radius: var(--radius-sm);
    padding: 0.6rem 0.9rem;
    font-family: var(--font-body);
    font-style: italic;
    font-size: 0.88rem;
    color: var(--accent);
    margin-top: 0.5rem;
    line-height: 1.4;
}

/* ─── SELECTBOX ─── */
[data-testid="stSelectbox"] > div > div {
    background: var(--parch-1) !important;
    border: 1px solid var(--rule-mid) !important;
    border-radius: var(--radius-md) !important;
    color: var(--ink-1) !important;
    font-family: var(--font-mono) !important;
    font-size: 0.9rem !important;
    letter-spacing: 0.06em !important;
}
[data-testid="stSelectbox"] > div > div:hover {
    border-color: var(--accent) !important;
}

/* ─── RADIO ─── */
[data-testid="stRadio"] > div {
    display: flex !important;
    gap: 0.6rem !important;
    flex-direction: row !important;
}
[data-testid="stRadio"] label {
    background: var(--parch-1) !important;
    border: 1px solid var(--rule-mid) !important;
    border-radius: var(--radius-sm) !important;
    padding: 0.65rem 1rem !important;
    font-family: var(--font-body) !important;
    font-size: 0.92rem !important;
    color: var(--ink-3) !important;
    cursor: pointer !important;
    transition: all 0.2s !important;
    flex: 1 !important;
    text-align: center !important;
}
[data-testid="stRadio"] label:hover {
    border-color: var(--accent-soft) !important;
    color: var(--ink-1) !important;
    background: var(--parch-0) !important;
}
[data-testid="stRadio"] label:has(input:checked) {
    background: var(--accent-pale) !important;
    border-color: var(--accent) !important;
    color: var(--accent) !important;
    font-weight: 600 !important;
}

/* ─── CHECKBOX ─── */
[data-testid="stCheckbox"] {
    font-family: var(--font-body) !important;
    font-size: 0.92rem !important;
    color: var(--ink-2) !important;
}

/* ─── FILE UPLOADER ─── */
[data-testid="stFileUploader"] section {
    background: var(--parch-1) !important;
    border: 1px dashed var(--rule-mid) !important;
    border-radius: var(--radius-md) !important;
    padding: 0.55rem 0.85rem !important;
    transition: border-color 0.2s !important;
}
[data-testid="stFileUploader"] section:hover {
    border-color: var(--accent-soft) !important;
    background: var(--parch-0) !important;
}
[data-testid="stFileUploader"] section p,
[data-testid="stFileUploader"] section span,
[data-testid="stFileUploader"] section small {
    font-family: var(--font-body) !important;
    font-style: italic !important;
    font-size: 0.82rem !important;
    color: var(--ink-4) !important;
}
[data-testid="stFileUploader"] section button {
    background: var(--parch-2) !important;
    border: 1px solid var(--rule-mid) !important;
    color: var(--ink-2) !important;
    font-family: var(--font-body) !important;
    font-size: 0.82rem !important;
    border-radius: var(--radius-sm) !important;
}

/* ─── BUTTONS ─── */
.stButton > button {
    background: var(--parch-0) !important;
    border: 1px solid var(--rule-mid) !important;
    color: var(--ink-2) !important;
    font-family: var(--font-display) !important;
    font-size: 1rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.04em !important;
    border-radius: var(--radius-sm) !important;
    padding: 0.7rem 1.4rem !important;
    width: 100% !important;
    transition: all 0.22s !important;
    box-shadow: var(--shadow-sm) !important;
}
.stButton > button:hover {
    background: var(--parch-1) !important;
    border-color: var(--rule-strong) !important;
    box-shadow: var(--shadow-md) !important;
    transform: translateY(-1px) !important;
}

.execute-btn .stButton > button {
    background: var(--accent) !important;
    border-color: var(--accent) !important;
    color: var(--parch-0) !important;
    font-size: 1.05rem !important;
    letter-spacing: 0.07em !important;
    box-shadow: 0 4px 16px rgba(139,58,42,0.28) !important;
}
.execute-btn .stButton > button:hover {
    background: var(--accent-soft) !important;
    box-shadow: 0 6px 22px rgba(139,58,42,0.35) !important;
}

/* ─── DOWNLOAD BUTTON ─── */
.stDownloadButton > button {
    background: var(--verdant-pale) !important;
    border: 1px solid var(--verdant-border) !important;
    color: var(--verdant) !important;
    font-family: var(--font-display) !important;
    font-size: 0.95rem !important;
    font-weight: 600 !important;
    border-radius: var(--radius-sm) !important;
    padding: 0.65rem 1rem !important;
    width: 100% !important;
    transition: all 0.2s !important;
}
.stDownloadButton > button:hover {
    background: rgba(74,123,106,0.2) !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 14px rgba(74,123,106,0.2) !important;
}

/* ─── PROGRESS ─── */
[data-testid="stProgress"] > div {
    background: var(--parch-3) !important;
    border-radius: 2px !important;
    height: 4px !important;
}
[data-testid="stProgress"] > div > div {
    background: linear-gradient(90deg, var(--accent), var(--gilt-soft)) !important;
    border-radius: 2px !important;
    box-shadow: 0 0 8px rgba(139,58,42,0.3) !important;
}

/* ─── NATIVE ALERTS ─── */
[data-testid="stAlert"] {
    background: var(--verdant-pale) !important;
    border: 1px solid var(--verdant-border) !important;
    border-radius: var(--radius-md) !important;
    font-family: var(--font-body) !important;
    font-style: italic !important;
    color: var(--verdant) !important;
}

/* ─── METRICS ─── */
.m-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 1rem;
    margin: 1.25rem 0;
}
.m-cell {
    background: var(--parch-0);
    border: 1px solid var(--rule-mid);
    border-radius: var(--radius-md);
    padding: 1.25rem 1rem;
    text-align: center;
    box-shadow: var(--shadow-sm);
    position: relative;
    overflow: hidden;
}
.m-cell::after {
    content: '';
    position: absolute;
    bottom: 0; left: 15%; right: 15%;
    height: 1px;
    background: var(--rule-mid);
}
.m-number {
    font-family: var(--font-display);
    font-size: 2.8rem;
    font-weight: 600;
    color: var(--accent);
    line-height: 1;
    margin-bottom: 0.3rem;
    letter-spacing: -0.02em;
}
.m-label {
    font-family: var(--font-fell);
    font-style: italic;
    font-size: 0.88rem;
    color: var(--ink-4);
}

/* ─── RESULT BANNER ─── */
.result-wrap {
    background: var(--parch-0);
    border: 1px solid var(--rule-mid);
    border-radius: var(--radius-xl);
    padding: 2rem 2.25rem;
    margin-bottom: 1.75rem;
    box-shadow: var(--shadow-lg);
    position: relative;
    overflow: hidden;
    animation: fadeUp 0.45s cubic-bezier(0.22, 1, 0.36, 1);
}
.result-wrap::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0; height: 2px;
    background: linear-gradient(90deg,
        transparent, var(--rule-mid) 20%,
        var(--gilt) 50%,
        var(--rule-mid) 80%, transparent
    );
}
.rw-title {
    font-family: var(--font-display);
    font-size: 1.75rem;
    font-weight: 600;
    color: var(--ink-0);
    margin-bottom: 0.2rem;
    letter-spacing: -0.01em;
}
.rw-title em { color: var(--accent); font-style: italic; }
.rw-sub {
    font-family: var(--font-fell);
    font-style: italic;
    font-size: 0.95rem;
    color: var(--ink-4);
}
@keyframes fadeUp {
    from { opacity: 0; transform: translateY(14px); }
    to   { opacity: 1; transform: translateY(0); }
}

/* ─── HISTORY ─── */
.hist-entry {
    background: var(--parch-0);
    border: 1px solid var(--rule-light);
    border-radius: var(--radius-sm);
    padding: 0.75rem 1.1rem;
    margin-bottom: 0.55rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
    transition: border-color 0.2s;
}
.hist-entry:hover { border-color: var(--rule-mid); }
.he-time {
    font-family: var(--font-mono);
    font-size: 0.74rem;
    color: var(--ink-2);
    font-weight: 700;
}
.he-info {
    font-family: var(--font-body);
    font-style: italic;
    font-size: 0.83rem;
    color: var(--ink-4);
    margin-top: 1px;
}
.he-badge {
    font-family: var(--font-mono);
    font-size: 0.63rem;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    padding: 0.22rem 0.55rem;
    border-radius: 3px;
}
.he-ok  { color: var(--verdant); background: var(--verdant-pale); border: 1px solid var(--verdant-border); }
.he-err { color: var(--accent);  background: var(--accent-pale);  border: 1px solid var(--accent-border); }

/* ─── PROGRESS STATUS ─── */
.prog-status {
    font-family: var(--font-mono);
    font-size: 0.75rem;
    color: var(--accent);
    text-align: center;
    padding: 0.4rem;
    letter-spacing: 0.08em;
}
.prog-detail {
    display: flex;
    justify-content: space-between;
    font-family: var(--font-mono);
    font-size: 0.68rem;
    color: var(--ink-5);
    margin-top: 0.25rem;
}

/* ─── SETTINGS ─── */
.settings-panel {
    background: var(--parch-0);
    border: 1px solid var(--rule-mid);
    border-radius: var(--radius-lg);
    padding: 1.5rem 1.75rem;
    margin-bottom: 1.75rem;
    box-shadow: var(--shadow-sm);
}

/* ─── EXPANDER ─── */
.streamlit-expanderHeader {
    font-family: var(--font-body) !important;
    font-style: italic !important;
    font-size: 0.88rem !important;
    color: var(--ink-3) !important;
    background: var(--parch-1) !important;
    border-radius: var(--radius-sm) !important;
}
.streamlit-expanderContent {
    background: var(--parch-1) !important;
    border: 1px solid var(--rule-light) !important;
    font-family: var(--font-body) !important;
}

/* ─── RESPONSIVE ─── */
@media (max-width: 960px) {
    .block-container { padding: 1rem 1.25rem !important; }
    .lh-title { font-size: 2.4rem; }
    .m-grid { grid-template-columns: repeat(3, 1fr); }
}
@media (max-width: 640px) {
    .lh-title { font-size: 1.9rem; }
    .m-grid { grid-template-columns: 1fr; }
    .m-number { font-size: 2.2rem; }
}
"""

st.markdown(f"<style>{THEME}</style>", unsafe_allow_html=True)

# ═══════════ HELPER: VALIDATE ═══════════
def validate_file(file, file_type):
    try:
        wb = load_workbook(file)
        ws = wb.active
        rows = ws.max_row
        cols = ws.max_column
        header_row = 7 if file_type in ["job", "einv"] else 3
        headers = []
        for col in range(1, min(cols + 1, 20)):
            val = ws.cell(header_row, col).value
            if val:
                headers.append(str(val).strip())
        return {
            "valid": True, "rows": rows, "cols": cols,
            "headers": headers,
            "sheets": wb.sheetnames if file_type == "bill" else None
        }
    except Exception as e:
        return {"valid": False, "error": str(e)}

# ═══════════ HEADER ═══════════
st.markdown("""
<div class="ledger-header">
    <span class="lh-ornament">❧ &nbsp; &nbsp; ❧</span>
    <div class="lh-double-rule"></div>
    <h1 class="lh-title">The Bill Report <em>Ledger</em></h1>
    <p class="lh-subtitle">Automated reconciliation of invoices, jobs &amp; billing records</p>
    <div class="lh-double-rule"></div>
    <div class="lh-meta">Est. Automation System &nbsp;·&nbsp; Report Processing Suite</div>
</div>
""", unsafe_allow_html=True)

# Settings toggle row
tog1, tog2 = st.columns([6, 1])
with tog2:
    stlbl = "⚙ Settings" if not st.session_state.show_settings else "✕ Close"
    if st.button(stlbl, key="settings_tog"):
        st.session_state.show_settings = not st.session_state.show_settings

# ═══════════ SETTINGS ═══════════
if st.session_state.show_settings:
    st.markdown('<div class="settings-panel">', unsafe_allow_html=True)
    s1, s2 = st.columns(2, gap="large")
    with s1:
        st.markdown('<span class="conf-label">File Preview</span>', unsafe_allow_html=True)
        show_prev = st.checkbox("Show file statistics after upload", value=st.session_state.show_preview, key="prev_chk")
        st.session_state.show_preview = show_prev
    with s2:
        st.markdown('<span class="conf-label">Processing History</span>', unsafe_allow_html=True)
        if st.button("Clear History", key="clr_hist"):
            st.session_state.processing_history = []
            st.success("History cleared.")
    st.markdown('</div>', unsafe_allow_html=True)

# ═══════════ MAIN COLUMNS ═══════════
MONTHS = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"]

left_col, right_col = st.columns([3, 2], gap="large")

# ─────────── LEFT: DOCUMENT INTAKE ───────────
with left_col:
    st.markdown('<div class="sect-heading"><span>I</span> Document Intake</div>', unsafe_allow_html=True)
    st.markdown('<div class="v-card">', unsafe_allow_html=True)

    # Bill Report
    st.markdown("""
    <div class="upload-entry">
        <div class="ue-row">
            <span class="ue-numeral">I</span>
            <span class="ue-title">Bill Report</span>
        </div>
        <div class="ue-desc">Primary billing register — Excel workbook containing monthly ledger sheets</div>
    </div>
    """, unsafe_allow_html=True)
    bill_file = st.file_uploader("", type="xlsx", key="bill", label_visibility="collapsed")
    if bill_file:
        st.markdown('<div class="file-loaded">✦ Document Received</div>', unsafe_allow_html=True)
        if st.session_state.show_preview:
            v = validate_file(bill_file, "bill")
            if v["valid"]:
                st.markdown(f"""
                <div class="preview-ledger">
                    <div class="pl-row"><span>Total Rows</span><span class="pl-val">{v['rows']}</span></div>
                    <div class="pl-row"><span>Columns</span><span class="pl-val">{v['cols']}</span></div>
                    <div class="pl-row"><span>Sheets</span><span class="pl-val">{len(v['sheets'])}</span></div>
                </div>
                """, unsafe_allow_html=True)
            bill_file.seek(0)

    st.markdown('<div class="orn-divider">✦</div>', unsafe_allow_html=True)

    # Job Report
    st.markdown("""
    <div class="upload-entry">
        <div class="ue-row">
            <span class="ue-numeral">II</span>
            <span class="ue-title">Job Report</span>
        </div>
        <div class="ue-desc">Job particulars — shipper names, container numbers &amp; invoice references</div>
    </div>
    """, unsafe_allow_html=True)
    job_file = st.file_uploader("", type="xlsx", key="job", label_visibility="collapsed")
    if job_file:
        st.markdown('<div class="file-loaded">✦ Document Received</div>', unsafe_allow_html=True)
        if st.session_state.show_preview:
            v = validate_file(job_file, "job")
            if v["valid"]:
                st.markdown(f"""
                <div class="preview-ledger">
                    <div class="pl-row"><span>Total Rows</span><span class="pl-val">{v['rows']}</span></div>
                    <div class="pl-row"><span>Columns</span><span class="pl-val">{v['cols']}</span></div>
                </div>
                """, unsafe_allow_html=True)
            job_file.seek(0)

    st.markdown('<div class="orn-divider">✦</div>', unsafe_allow_html=True)

    # E-Invoice
    st.markdown("""
    <div class="upload-entry">
        <div class="ue-row">
            <span class="ue-numeral">III</span>
            <span class="ue-title">E-Invoice Report</span>
        </div>
        <div class="ue-desc">Electronic invoice register — bill numbers, issuance dates &amp; job correspondence</div>
    </div>
    """, unsafe_allow_html=True)
    einv_file = st.file_uploader("", type="xlsx", key="einv", label_visibility="collapsed")
    if einv_file:
        st.markdown('<div class="file-loaded">✦ Document Received</div>', unsafe_allow_html=True)
        if st.session_state.show_preview:
            v = validate_file(einv_file, "einv")
            if v["valid"]:
                st.markdown(f"""
                <div class="preview-ledger">
                    <div class="pl-row"><span>Total Rows</span><span class="pl-val">{v['rows']}</span></div>
                    <div class="pl-row"><span>Columns</span><span class="pl-val">{v['cols']}</span></div>
                </div>
                """, unsafe_allow_html=True)
            einv_file.seek(0)

    st.markdown('</div>', unsafe_allow_html=True)

# ─────────── RIGHT: CONFIGURATION ───────────
with right_col:
    st.markdown('<div class="sect-heading"><span>II</span> Configuration</div>', unsafe_allow_html=True)
    st.markdown('<div class="v-card">', unsafe_allow_html=True)

    st.markdown('<span class="conf-label">Billing Period</span>', unsafe_allow_html=True)
    month = st.selectbox("", MONTHS, index=st.session_state.selected_month,
                         label_visibility="collapsed", key="month_sel")
    st.session_state.selected_month = MONTHS.index(month)

    st.markdown('<div class="orn-divider">· · ·</div>', unsafe_allow_html=True)

    st.markdown('<span class="conf-label">Update Mode</span>', unsafe_allow_html=True)
    update_mode = st.radio("", ["Append", "Overwrite"],
                           label_visibility="collapsed", key="umode")
    if update_mode == "Append":
        st.markdown('<div class="v-alert-info">New entries shall be appended to the existing record without disturbing prior data.</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="v-alert-warn">All existing entries in the selected month shall be expunged and replaced anew.</div>', unsafe_allow_html=True)

    st.markdown('<div class="orn-divider">· · ·</div>', unsafe_allow_html=True)

    st.markdown('<div class="execute-btn">', unsafe_allow_html=True)
    process_clicked = st.button("✦  Execute Report Generation", key="proc_main", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

    # Activity log
    if st.session_state.processing_history:
        st.markdown("<br/>", unsafe_allow_html=True)
        st.markdown('<div class="sect-heading"><span>III</span> Activity Log</div>', unsafe_allow_html=True)
        st.markdown('<div class="v-card">', unsafe_allow_html=True)
        for entry in reversed(st.session_state.processing_history[-4:]):
            badge = "he-ok" if entry["status"] == "success" else "he-err"
            mark  = "DONE"  if entry["status"] == "success" else "ERR"
            st.markdown(f"""
            <div class="hist-entry">
                <div>
                    <div class="he-time">{entry["time"]} &nbsp;·&nbsp; {entry["month"]}</div>
                    <div class="he-info">+{entry['added']} added &nbsp;·&nbsp; {entry['skipped']} skipped &nbsp;·&nbsp; {entry.get('duration','—')}</div>
                </div>
                <div class="he-badge {badge}">{mark}</div>
            </div>
            """, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

# ═══════════ PROCESSING CONSTANTS ═══════════
BILL_HEADER_ROW = 3
BILL_DATA_START = 4
SRC_HEADER_ROW  = 7
SRC_DATA_START  = 8

# ═══════════ HELPERS ═══════════
def norm(x):
    return str(x).strip().upper() if x else ""

def clean_job(job):
    return norm(job).split("/")[0]

def parse_date(date_val):
    if not date_val: return None
    if isinstance(date_val, datetime): return date_val.date()
    date_str = str(date_val).strip()
    if not date_str or date_str.upper() == 'NONE': return None
    for fmt in ['%Y-%m-%d','%d-%m-%Y','%d/%m/%Y','%m/%d/%Y','%d.%m.%Y',
                '%Y/%m/%d','%d-%b-%Y','%d %b %Y','%d-%m-%y','%d/%m/%y']:
        try:
            return datetime.strptime(date_str, fmt).date()
        except (ValueError, TypeError):
            continue
    return None

def format_date_for_excel(date_val):
    if not date_val: return None
    if isinstance(date_val, datetime): return date_val.strftime('%d/%m/%y')
    try: return date_val.strftime('%d/%m/%y')
    except: return None

def col_map(ws, header_row):
    m = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v: m[norm(v)] = c
    return m

def get_existing_jobs(ws, bill_cols, start_row):
    existing = set()
    for r in range(start_row, ws.max_row + 1):
        j = ws.cell(r, bill_cols["JOB NO"]).value
        if j: existing.add(clean_job(j))
    return existing

def find_next_empty_row(ws, start_row):
    for r in range(start_row, ws.max_row + 100):
        if not ws.cell(r, 1).value and not ws.cell(r, 2).value: return r
    return ws.max_row + 1

# ═══════════ PROCESSING ═══════════
if process_clicked:
    if not (bill_file and job_file and einv_file):
        st.error("Please furnish all three documents before proceeding.")
    else:
        progress_bar = st.progress(0)
        status_el    = st.empty()
        detail_el    = st.empty()
        start_time   = datetime.now()

        def upd(pct, msg, det=""):
            progress_bar.progress(pct)
            status_el.markdown(f'<div class="prog-status">{msg}</div>', unsafe_allow_html=True)
            if det:
                detail_el.markdown(
                    f'<div class="prog-detail"><span>{det}</span><span>{pct}%</span></div>',
                    unsafe_allow_html=True
                )

        try:
            upd(10, "Loading workbooks…", "Reading Excel documents")
            bill_wb = load_workbook(bill_file)
            job_wb  = load_workbook(job_file)
            einv_wb = load_workbook(einv_file)

            upd(25, "Locating sheets…", "Verifying month tab")
            if month not in bill_wb.sheetnames:
                st.error(f"The sheet '{month}' was not found in the Bill Report.")
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
                    st.error(f"Column absent in Job Report: {col}"); st.stop()
            for col in ["JOB NO","BILL NO","BILL DATE"]:
                if col not in einv_cols:
                    st.error(f"Column absent in E-Invoice: {col}"); st.stop()
            for col in ["JOB NO","JOB DATE","PARTY NAME","INVOICE NO","SB / BE NO","CONTAINER NO","BILL NO","BILL DATE"]:
                if col not in bill_cols:
                    st.error(f"Column absent in Bill Report: {col}"); st.stop()

            append_mode = update_mode == "Append"
            existing_jobs = set()

            if append_mode:
                upd(55, "Scanning existing entries…", "Checking for duplicates")
                existing_jobs = get_existing_jobs(bill_ws, bill_cols, BILL_DATA_START)
                row = find_next_empty_row(bill_ws, BILL_DATA_START)
                last_sr = 0
                for r in range(BILL_DATA_START, row):
                    sv = bill_ws.cell(r, 1).value
                    if sv and isinstance(sv, (int, float)): last_sr = max(last_sr, int(sv))
                sr = last_sr + 1
            else:
                upd(55, "Clearing ledger sheet…", "Preparing for fresh entries")
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
                raw   = einv_ws.cell(r, einv_cols["JOB NO"]).value
                bno   = einv_ws.cell(r, einv_cols["BILL NO"]).value
                bdate = parse_date(einv_ws.cell(r, einv_cols["BILL DATE"]).value)
                if raw and bno:
                    bill_map[clean_job(raw)].append((str(bno), bdate))

            upd(88, "Composing final report…", "Writing consolidated entries")
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
                        main, suffix = (fb.rsplit("/",1)[0], "/"+fb.rsplit("/",1)[1]) if "/" in fb else (fb, "")
                        extra = []
                        for i in range(1, len(bills)):
                            bn = str(bills[i][0])
                            if "/" in bn: bn = bn.split("/")[0]
                            extra.append(bn[-4:] if len(bn) >= 4 else bn)
                        merged = main + "/" + "/".join(extra) + suffix
                    vd = [b[1] for b in bills if b[1]]
                    ld = max(vd) if vd else None
                    bill_ws.cell(row, bill_cols["BILL NO"]).value   = merged
                    bill_ws.cell(row, bill_cols["BILL DATE"]).value = format_date_for_excel(ld)
                row += 1; sr += 1; added += 1

            upd(96, "Committing to file…", "Saving workbook")
            out = f"UPDATED_BILL_REPORT_{month}.xlsx"
            bill_wb.save(out)
            progress_bar.progress(100)

            end_time = datetime.now()
            duration = (end_time - start_time).total_seconds()

            status_el.empty(); detail_el.empty(); progress_bar.empty()

            st.session_state.processing_history.append({
                "time": end_time.strftime("%H:%M:%S"), "month": month,
                "status": "success", "added": added, "skipped": skipped,
                "duration": f"{duration:.1f}s"
            })

            r1, r2 = st.columns([3, 1], gap="large")

            with r1:
                st.markdown(f"""
                <div class="result-wrap">
                    <div class="rw-title">Report <em>Complete</em></div>
                    <div class="rw-sub">Processed in {duration:.1f} seconds &nbsp;·&nbsp; {month} ledger updated</div>
                    <div class="m-grid">
                        <div class="m-cell">
                            <div class="m-number">{added}</div>
                            <div class="m-label">Entries Added</div>
                        </div>
                        <div class="m-cell">
                            <div class="m-number">{skipped}</div>
                            <div class="m-label">Skipped</div>
                        </div>
                        <div class="m-cell">
                            <div class="m-number">{added + skipped}</div>
                            <div class="m-label">Total Examined</div>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

            with r2:
                st.markdown('<div class="v-card">', unsafe_allow_html=True)
                st.markdown(
                    '<div class="sect-heading" style="font-size:1rem; margin-bottom:1rem;">'
                    '<span>↓</span> Export</div>',
                    unsafe_allow_html=True
                )
                with open(out, "rb") as f:
                    st.download_button("Download Excel", data=f, file_name=out,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True, key="dl_xl")
                try:
                    df  = pd.read_excel(out, sheet_name=month, header=BILL_HEADER_ROW - 1)
                    csv = df.to_csv(index=False)
                    st.download_button("Download CSV", data=csv,
                        file_name=f"UPDATED_BILL_REPORT_{month}.csv",
                        mime="text/csv", use_container_width=True, key="dl_csv")
                except:
                    pass
                st.markdown('</div>', unsafe_allow_html=True)

        except Exception as e:
            progress_bar.empty(); status_el.empty(); detail_el.empty()
            st.session_state.processing_history.append({
                "time": datetime.now().strftime("%H:%M:%S"), "month": month,
                "status": "error", "added": 0, "skipped": 0, "duration": "—"
            })
            st.error(f"An error was encountered during processing: {str(e)}")
            with st.expander("View error details"):
                st.exception(e)