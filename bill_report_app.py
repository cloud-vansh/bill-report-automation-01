import streamlit as st
from openpyxl import load_workbook
from collections import defaultdict
from datetime import datetime
import pandas as pd

# ================= PAGE CONFIG =================
st.set_page_config(
    page_title="BillFlow — Report Automation",
    page_icon="⬡",
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
if 'show_results' not in st.session_state:
    st.session_state.show_results = False
if 'processing_results' not in st.session_state:
    st.session_state.processing_results = None

# ================= DESIGN SYSTEM =================
THEME = """
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;500;600;700;800&family=DM+Mono:wght@300;400;500&family=DM+Sans:wght@300;400;500;600&display=swap');

:root {
    --bg-void:        #06060A;
    --bg-base:        #0C0C12;
    --bg-surface:     #111119;
    --bg-elevated:    #17171F;
    --bg-hover:       #1C1C26;
    --border-dim:     #1E1E2C;
    --border-mid:     #2A2A3C;
    --border-bright:  #3A3A52;
    --gold:           #E8B84B;
    --gold-dim:       #C49A35;
    --gold-glow:      rgba(232,184,75,0.15);
    --gold-glow-sm:   rgba(232,184,75,0.08);
    --teal:           #3DD6C8;
    --teal-dim:       #2AADA0;
    --teal-glow:      rgba(61,214,200,0.12);
    --red:            #FF5757;
    --red-glow:       rgba(255,87,87,0.12);
    --amber-soft:     #FBB040;
    --text-primary:   #F0EEF8;
    --text-secondary: #8B899E;
    --text-muted:     #4A4860;
    --text-micro:     #333248;
    --mono:           'DM Mono', monospace;
    --sans:           'DM Sans', sans-serif;
    --display:        'Syne', sans-serif;
    --radius-sm:      6px;
    --radius-md:      10px;
    --radius-lg:      16px;
    --radius-xl:      22px;
}

* { box-sizing: border-box; margin: 0; padding: 0; }

html, body, [data-testid="stAppViewContainer"],
[data-testid="stAppViewBlockContainer"],
.main .block-container {
    font-family: var(--sans);
    background: var(--bg-void) !important;
    color: var(--text-primary);
}

/* ─── HIDE STREAMLIT CHROME ─── */
#MainMenu, footer, header, .stDeployButton,
[data-testid="stToolbar"], [data-testid="stDecoration"],
[data-testid="stStatusWidget"] { visibility: hidden !important; display: none !important; }

/* ─── SCROLLBAR ─── */
::-webkit-scrollbar { width: 4px; }
::-webkit-scrollbar-track { background: var(--bg-void); }
::-webkit-scrollbar-thumb { background: var(--border-mid); border-radius: 2px; }

/* ─── MAIN CONTAINER ─── */
.main { background: var(--bg-void) !important; }
.block-container {
    max-width: 1400px !important;
    padding: 2rem 2.5rem !important;
}

/* ─── DOT GRID OVERLAY ─── */
.stApp::before {
    content: '';
    position: fixed;
    inset: 0;
    background-image: radial-gradient(circle, var(--text-micro) 1px, transparent 1px);
    background-size: 28px 28px;
    pointer-events: none;
    z-index: 0;
    opacity: 0.6;
}

/* ─── HEADER ─── */
.header-wrap {
    position: relative;
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 1.5rem 2rem;
    background: var(--bg-surface);
    border: 1px solid var(--border-dim);
    border-radius: var(--radius-xl);
    margin-bottom: 1.75rem;
    overflow: hidden;
}
.header-wrap::before {
    content: '';
    position: absolute;
    left: 0; top: 0; bottom: 0;
    width: 3px;
    background: linear-gradient(180deg, var(--gold) 0%, transparent 100%);
    border-radius: 0 0 0 var(--radius-xl);
}
.header-accent-bar {
    position: absolute;
    right: 2rem; top: 50%; transform: translateY(-50%);
    display: flex; gap: 6px; align-items: center;
}
.accent-dot { width: 6px; height: 6px; border-radius: 50%; }

.hd-wordmark {
    font-family: var(--display);
    font-size: 1.5rem;
    font-weight: 800;
    letter-spacing: -0.02em;
    color: var(--text-primary);
}
.hd-wordmark span { color: var(--gold); }
.hd-sub {
    font-family: var(--mono);
    font-size: 0.68rem;
    color: var(--text-muted);
    letter-spacing: 0.12em;
    text-transform: uppercase;
    margin-top: 2px;
}

.hd-status {
    display: flex;
    align-items: center;
    gap: 0.5rem;
    font-family: var(--mono);
    font-size: 0.72rem;
    color: var(--teal);
    background: var(--teal-glow);
    border: 1px solid rgba(61,214,200,0.2);
    padding: 0.4rem 0.9rem;
    border-radius: 30px;
}
.pulse {
    width: 6px; height: 6px;
    background: var(--teal);
    border-radius: 50%;
    animation: pulse 2s infinite;
}
@keyframes pulse {
    0%, 100% { opacity: 1; box-shadow: 0 0 0 0 rgba(61,214,200,0.4); }
    50% { opacity: 0.7; box-shadow: 0 0 0 5px rgba(61,214,200,0); }
}

/* ─── SECTION LABEL ─── */
.section-label {
    font-family: var(--mono);
    font-size: 0.65rem;
    letter-spacing: 0.18em;
    text-transform: uppercase;
    color: var(--text-muted);
    margin-bottom: 0.75rem;
    display: flex;
    align-items: center;
    gap: 0.6rem;
}
.section-label::after {
    content: '';
    flex: 1;
    height: 1px;
    background: var(--border-dim);
}

/* ─── CARD BASE ─── */
.card {
    background: var(--bg-surface);
    border: 1px solid var(--border-dim);
    border-radius: var(--radius-lg);
    padding: 1.5rem;
    position: relative;
    transition: border-color 0.2s;
}
.card:hover { border-color: var(--border-mid); }

/* ─── UPLOAD ITEMS ─── */
.upload-item {
    background: var(--bg-elevated);
    border: 1px solid var(--border-dim);
    border-radius: var(--radius-md);
    padding: 1.1rem 1.25rem 0.85rem;
    margin-bottom: 0.75rem;
    position: relative;
    transition: border-color 0.25s, background 0.25s;
}
.upload-item:last-child { margin-bottom: 0; }
.upload-item:hover {
    border-color: var(--gold-dim);
    background: var(--bg-hover);
}
.upload-item::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 1px;
    background: linear-gradient(90deg, transparent, var(--gold-glow-sm), transparent);
    border-radius: var(--radius-md) var(--radius-md) 0 0;
}

.ui-header {
    display: flex;
    align-items: center;
    gap: 0.75rem;
    margin-bottom: 0.35rem;
}
.ui-index {
    font-family: var(--mono);
    font-size: 0.65rem;
    color: var(--text-muted);
    background: var(--bg-surface);
    border: 1px solid var(--border-mid);
    width: 22px; height: 22px;
    border-radius: 4px;
    display: flex; align-items: center; justify-content: center;
    flex-shrink: 0;
}
.ui-title {
    font-family: var(--display);
    font-size: 0.9rem;
    font-weight: 600;
    color: var(--text-primary);
}
.ui-desc {
    font-family: var(--sans);
    font-size: 0.75rem;
    color: var(--text-secondary);
    margin-bottom: 0.7rem;
    padding-left: 2.25rem;
}

.file-loaded-chip {
    display: inline-flex; align-items: center; gap: 0.4rem;
    font-family: var(--mono); font-size: 0.68rem;
    color: var(--teal);
    background: var(--teal-glow);
    border: 1px solid rgba(61,214,200,0.25);
    padding: 0.25rem 0.65rem;
    border-radius: 20px;
    margin-top: 0.35rem;
    margin-left: 2.25rem;
}

/* ─── PREVIEW BOX ─── */
.preview-box {
    background: var(--bg-void);
    border: 1px solid var(--border-dim);
    border-radius: var(--radius-sm);
    padding: 0.75rem 1rem;
    margin-top: 0.5rem;
    margin-left: 2.25rem;
}
.pb-row {
    display: flex; justify-content: space-between;
    font-family: var(--mono); font-size: 0.7rem;
    color: var(--text-secondary);
    padding: 0.2rem 0;
    border-bottom: 1px solid var(--border-dim);
}
.pb-row:last-child { border-bottom: none; }
.pb-val { color: var(--gold); font-weight: 500; }

/* ─── CONFIG CARD ─── */
.config-label {
    font-family: var(--mono);
    font-size: 0.68rem;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: var(--text-muted);
    margin-bottom: 0.4rem;
    display: block;
}

/* ─── SELECTBOX ─── */
[data-testid="stSelectbox"] > div > div {
    background: var(--bg-elevated) !important;
    border: 1px solid var(--border-mid) !important;
    border-radius: var(--radius-md) !important;
    color: var(--text-primary) !important;
    font-family: var(--mono) !important;
    font-size: 0.85rem !important;
}
[data-testid="stSelectbox"] > div > div:hover {
    border-color: var(--gold-dim) !important;
}

/* ─── RADIO ─── */
[data-testid="stRadio"] > div {
    display: flex !important;
    gap: 0.6rem !important;
    flex-direction: row !important;
}
[data-testid="stRadio"] label {
    background: var(--bg-elevated) !important;
    border: 1px solid var(--border-mid) !important;
    border-radius: var(--radius-md) !important;
    padding: 0.6rem 1rem !important;
    font-family: var(--mono) !important;
    font-size: 0.78rem !important;
    color: var(--text-secondary) !important;
    cursor: pointer !important;
    transition: all 0.2s !important;
    flex: 1 !important;
    text-align: center !important;
}
[data-testid="stRadio"] label:hover {
    border-color: var(--gold-dim) !important;
    color: var(--text-primary) !important;
}
[data-testid="stRadio"] label:has(input:checked) {
    background: var(--gold-glow) !important;
    border-color: var(--gold) !important;
    color: var(--gold) !important;
}

/* ─── CHECKBOX ─── */
[data-testid="stCheckbox"] {
    font-family: var(--sans) !important;
    font-size: 0.82rem !important;
    color: var(--text-secondary) !important;
}
[data-testid="stCheckbox"] span { color: var(--text-secondary) !important; }

/* ─── FILE UPLOADER ─── */
[data-testid="stFileUploader"] section {
    background: var(--bg-void) !important;
    border: 1px dashed var(--border-mid) !important;
    border-radius: var(--radius-md) !important;
    padding: 0.5rem 0.75rem !important;
    transition: border-color 0.2s !important;
}
[data-testid="stFileUploader"] section:hover {
    border-color: var(--gold-dim) !important;
}
[data-testid="stFileUploader"] section p,
[data-testid="stFileUploader"] section span,
[data-testid="stFileUploader"] section small {
    font-family: var(--mono) !important;
    font-size: 0.72rem !important;
    color: var(--text-muted) !important;
}

/* ─── BUTTONS ─── */
.stButton > button {
    background: transparent !important;
    border: 1px solid var(--gold-dim) !important;
    color: var(--gold) !important;
    font-family: var(--display) !important;
    font-size: 0.85rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.05em !important;
    border-radius: var(--radius-md) !important;
    padding: 0.7rem 1.5rem !important;
    width: 100% !important;
    transition: all 0.2s !important;
    position: relative !important;
    overflow: hidden !important;
}
.stButton > button::before {
    content: '' !important;
    position: absolute !important;
    inset: 0 !important;
    background: var(--gold-glow) !important;
    opacity: 0 !important;
    transition: opacity 0.2s !important;
}
.stButton > button:hover {
    border-color: var(--gold) !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 8px 24px rgba(232,184,75,0.18) !important;
}
.stButton > button:hover::before { opacity: 1 !important; }

.process-btn > button,
div[data-testid="column"]:last-child .stButton > button.process-trigger {
    background: linear-gradient(135deg, var(--gold) 0%, var(--amber-soft) 100%) !important;
    border-color: transparent !important;
    color: var(--bg-void) !important;
    font-weight: 700 !important;
    box-shadow: 0 4px 20px rgba(232,184,75,0.3) !important;
}

/* ─── DOWNLOAD BUTTON ─── */
.stDownloadButton > button {
    background: var(--teal-glow) !important;
    border: 1px solid rgba(61,214,200,0.35) !important;
    color: var(--teal) !important;
    font-family: var(--display) !important;
    font-size: 0.82rem !important;
    font-weight: 600 !important;
    border-radius: var(--radius-md) !important;
    padding: 0.65rem 1rem !important;
    width: 100% !important;
    transition: all 0.2s !important;
}
.stDownloadButton > button:hover {
    background: rgba(61,214,200,0.2) !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 18px rgba(61,214,200,0.2) !important;
}

/* ─── ALERT BOXES ─── */
.alert-info {
    background: rgba(91,111,255,0.08);
    border: 1px solid rgba(91,111,255,0.2);
    border-left: 2px solid #5B6FFF;
    border-radius: var(--radius-sm);
    padding: 0.55rem 0.8rem;
    font-family: var(--mono);
    font-size: 0.71rem;
    color: #8899FF;
    margin-top: 0.5rem;
}
.alert-warn {
    background: rgba(232,184,75,0.07);
    border: 1px solid rgba(232,184,75,0.2);
    border-left: 2px solid var(--gold);
    border-radius: var(--radius-sm);
    padding: 0.55rem 0.8rem;
    font-family: var(--mono);
    font-size: 0.71rem;
    color: var(--gold-dim);
    margin-top: 0.5rem;
}

/* ─── PROGRESS BAR ─── */
[data-testid="stProgress"] > div {
    background: var(--bg-elevated) !important;
    border-radius: 3px !important;
    height: 3px !important;
}
[data-testid="stProgress"] > div > div {
    background: linear-gradient(90deg, var(--gold-dim), var(--gold), var(--amber-soft)) !important;
    border-radius: 3px !important;
    box-shadow: 0 0 12px rgba(232,184,75,0.4) !important;
}

/* ─── SUCCESS / ERROR MESSAGES ─── */
[data-testid="stAlert"] {
    background: var(--teal-glow) !important;
    border: 1px solid rgba(61,214,200,0.25) !important;
    border-radius: var(--radius-md) !important;
    font-family: var(--mono) !important;
    font-size: 0.8rem !important;
    color: var(--teal) !important;
}
[data-testid="stAlert"][data-baseweb="notification"][kind="error"] {
    background: var(--red-glow) !important;
    border-color: rgba(255,87,87,0.25) !important;
    color: var(--red) !important;
}
.stSuccess, .stError, .stInfo, .stWarning {
    font-family: var(--mono) !important;
    font-size: 0.8rem !important;
    border-radius: var(--radius-md) !important;
}

/* ─── METRIC CARDS ─── */
.metric-row {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 0.75rem;
    margin: 1rem 0;
}
.metric-card {
    background: var(--bg-elevated);
    border: 1px solid var(--border-dim);
    border-radius: var(--radius-md);
    padding: 1.1rem 1rem;
    text-align: center;
    position: relative;
    overflow: hidden;
}
.metric-card::after {
    content: '';
    position: absolute;
    bottom: 0; left: 0; right: 0; height: 2px;
    background: linear-gradient(90deg, transparent, var(--gold), transparent);
    opacity: 0.5;
}
.metric-val {
    font-family: var(--display);
    font-size: 2.2rem;
    font-weight: 800;
    color: var(--gold);
    line-height: 1;
    margin-bottom: 0.35rem;
}
.metric-lbl {
    font-family: var(--mono);
    font-size: 0.65rem;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: var(--text-muted);
}

/* ─── RESULT BANNER ─── */
.result-banner {
    background: var(--bg-surface);
    border: 1px solid rgba(61,214,200,0.3);
    border-radius: var(--radius-xl);
    padding: 1.5rem 2rem;
    margin-bottom: 1.5rem;
    position: relative;
    overflow: hidden;
    animation: slideIn 0.4s cubic-bezier(0.22, 1, 0.36, 1);
}
.result-banner::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0; height: 2px;
    background: linear-gradient(90deg, transparent, var(--teal), transparent);
}
.rb-title {
    font-family: var(--display);
    font-size: 1.2rem;
    font-weight: 700;
    color: var(--teal);
    margin-bottom: 0.25rem;
}
.rb-sub {
    font-family: var(--mono);
    font-size: 0.72rem;
    color: var(--text-muted);
}
@keyframes slideIn {
    from { opacity: 0; transform: translateY(-16px); }
    to   { opacity: 1; transform: translateY(0); }
}

/* ─── HISTORY ITEMS ─── */
.history-entry {
    background: var(--bg-elevated);
    border: 1px solid var(--border-dim);
    border-radius: var(--radius-md);
    padding: 0.75rem 1rem;
    margin-bottom: 0.5rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
    transition: border-color 0.2s;
}
.history-entry:hover { border-color: var(--border-mid); }
.he-left { display: flex; flex-direction: column; gap: 2px; }
.he-time {
    font-family: var(--mono);
    font-size: 0.72rem;
    color: var(--text-primary);
    font-weight: 500;
}
.he-detail {
    font-family: var(--mono);
    font-size: 0.65rem;
    color: var(--text-muted);
}
.he-badge {
    font-family: var(--mono);
    font-size: 0.63rem;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    padding: 0.25rem 0.6rem;
    border-radius: 4px;
}
.he-badge.ok {
    color: var(--teal);
    background: var(--teal-glow);
    border: 1px solid rgba(61,214,200,0.2);
}
.he-badge.err {
    color: var(--red);
    background: var(--red-glow);
    border: 1px solid rgba(255,87,87,0.2);
}

/* ─── SETTINGS PANEL ─── */
.settings-section {
    background: var(--bg-surface);
    border: 1px solid var(--border-dim);
    border-radius: var(--radius-lg);
    padding: 1.25rem 1.5rem;
    margin-bottom: 1.5rem;
}
.settings-row {
    display: flex; align-items: flex-start; justify-content: space-between;
    padding: 0.85rem 0;
    border-bottom: 1px solid var(--border-dim);
}
.settings-row:last-child { border-bottom: none; padding-bottom: 0; }
.sr-info { flex: 1; }
.sr-title {
    font-family: var(--display);
    font-size: 0.85rem;
    font-weight: 600;
    color: var(--text-primary);
    margin-bottom: 2px;
}
.sr-desc {
    font-family: var(--mono);
    font-size: 0.68rem;
    color: var(--text-muted);
}

/* ─── PROGRESS DETAIL ─── */
.prog-detail {
    display: flex;
    justify-content: space-between;
    font-family: var(--mono);
    font-size: 0.68rem;
    color: var(--text-muted);
    margin-top: 0.3rem;
}

/* ─── DIVIDER ─── */
.gold-divider {
    height: 1px;
    background: linear-gradient(90deg, transparent, var(--border-mid), transparent);
    margin: 1.25rem 0;
}

/* ─── STATUS TEXT ─── */
.status-msg {
    font-family: var(--mono);
    font-size: 0.75rem;
    color: var(--gold);
    text-align: center;
    padding: 0.5rem;
    animation: blink 1.5s step-end infinite;
}
@keyframes blink {
    0%, 100% { opacity: 1; }
    50% { opacity: 0.5; }
}

/* ─── EXPANDER ─── */
.streamlit-expanderHeader {
    font-family: var(--mono) !important;
    font-size: 0.78rem !important;
    color: var(--text-secondary) !important;
    background: var(--bg-elevated) !important;
    border-radius: var(--radius-md) !important;
}
.streamlit-expanderContent {
    background: var(--bg-elevated) !important;
    border: 1px solid var(--border-dim) !important;
    border-radius: 0 0 var(--radius-md) var(--radius-md) !important;
}

/* ─── RESPONSIVE ─── */
@media (max-width: 900px) {
    .block-container { padding: 1rem 1rem !important; }
    .header-wrap { flex-direction: column; gap: 1rem; align-items: flex-start; }
    .metric-row { grid-template-columns: repeat(3, 1fr); }
    .hd-accent { display: none; }
}
@media (max-width: 640px) {
    .metric-row { grid-template-columns: 1fr; }
    .metric-val { font-size: 1.75rem; }
}
"""

st.markdown(f"<style>{THEME}</style>", unsafe_allow_html=True)

# ================= HELPER: VALIDATE FILE =================
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

# ================= RESULTS BANNER =================
if st.session_state.get('show_results', False):
    results = st.session_state.processing_results
    rb_col1, rb_col2 = st.columns([3, 1], gap="medium")
    with rb_col1:
        st.markdown(f"""
        <div class="result-banner">
            <div class="rb-title">⬡ Processing Complete</div>
            <div class="rb-sub">Report generated successfully — ready for download</div>
            <div class="metric-row" style="margin-top:1rem; margin-bottom:0;">
                <div class="metric-card">
                    <div class="metric-val">{results['added']}</div>
                    <div class="metric-lbl">Jobs Added</div>
                </div>
                <div class="metric-card">
                    <div class="metric-val">{results['skipped']}</div>
                    <div class="metric-lbl">Skipped</div>
                </div>
                <div class="metric-card">
                    <div class="metric-val">{results['total']}</div>
                    <div class="metric-lbl">Total Processed</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    with rb_col2:
        st.markdown("""
        <div class="card" style="height:100%;">
            <div class="section-label">Downloads</div>
        """, unsafe_allow_html=True)
        with open(results['filename'], "rb") as f:
            st.download_button("↓ Excel Report", data=f, file_name=results['filename'],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="rb_dl_xl")
        if results.get('csv_data'):
            st.download_button("↓ CSV Export", data=results['csv_data'],
                file_name=results['csv_filename'], mime="text/csv", key="rb_dl_csv")
        if st.button("✕ Dismiss", key="dismiss_rb"):
            st.session_state.show_results = False
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)

# ================= HEADER =================
st.markdown("""
<div class="header-wrap">
    <div>
        <div class="hd-wordmark">Bill<span>Flow</span></div>
        <div class="hd-sub">Report Automation System · v2.0</div>
    </div>
    <div class="hd-status">
        <div class="pulse"></div>
        SYSTEM READY
    </div>
    <div class="header-accent-bar hd-accent">
        <div class="accent-dot" style="background:#E8B84B;"></div>
        <div class="accent-dot" style="background:#3DD6C8;"></div>
        <div class="accent-dot" style="background:#333248;"></div>
    </div>
</div>
""", unsafe_allow_html=True)

# Settings toggle
st_col1, st_col2 = st.columns([6, 1])
with st_col2:
    settings_label = "⚙ Settings" if not st.session_state.show_settings else "✕ Close"
    if st.button(settings_label, key="settings_toggle"):
        st.session_state.show_settings = not st.session_state.show_settings

# ================= SETTINGS PANEL =================
if st.session_state.show_settings:
    st.markdown('<div class="settings-section">', unsafe_allow_html=True)
    st.markdown('<div class="section-label">Preferences</div>', unsafe_allow_html=True)

    s1, s2 = st.columns(2, gap="medium")
    with s1:
        st.markdown("""
        <div class="settings-row">
            <div class="sr-info">
                <div class="sr-title">File Preview</div>
                <div class="sr-desc">Show file statistics after upload</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        show_prev = st.checkbox("Enable file preview", value=st.session_state.show_preview, key="prev_toggle")
        st.session_state.show_preview = show_prev

    with s2:
        st.markdown("""
        <div class="settings-row">
            <div class="sr-info">
                <div class="sr-title">Processing History</div>
                <div class="sr-desc">Manage operation log</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Clear History", key="clear_hist"):
            st.session_state.processing_history = []
            st.success("History cleared")

    st.markdown('</div>', unsafe_allow_html=True)

# ================= MAIN LAYOUT =================
left_col, right_col = st.columns([3, 2], gap="large")

# ──────────────── LEFT: UPLOADS ────────────────
with left_col:
    st.markdown('<div class="section-label">Document Intake</div>', unsafe_allow_html=True)
    st.markdown('<div class="card">', unsafe_allow_html=True)

    # ── BILL REPORT ──
    st.markdown("""
    <div class="upload-item">
        <div class="ui-header">
            <div class="ui-index">01</div>
            <div class="ui-title">Bill Report</div>
        </div>
        <div class="ui-desc">Primary billing report — Excel workbook with monthly sheets</div>
    </div>
    """, unsafe_allow_html=True)
    bill_file = st.file_uploader("", type="xlsx", key="bill", label_visibility="collapsed")
    if bill_file:
        st.markdown('<div class="file-loaded-chip">✓ LOADED</div>', unsafe_allow_html=True)
        if st.session_state.show_preview:
            v = validate_file(bill_file, "bill")
            if v["valid"]:
                st.markdown(f"""
                <div class="preview-box">
                    <div class="pb-row"><span>Rows</span><span class="pb-val">{v['rows']}</span></div>
                    <div class="pb-row"><span>Columns</span><span class="pb-val">{v['cols']}</span></div>
                    <div class="pb-row"><span>Sheets</span><span class="pb-val">{len(v['sheets'])}</span></div>
                </div>
                """, unsafe_allow_html=True)
            bill_file.seek(0)

    st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)

    # ── JOB REPORT ──
    st.markdown("""
    <div class="upload-item">
        <div class="ui-header">
            <div class="ui-index">02</div>
            <div class="ui-title">Job Report</div>
        </div>
        <div class="ui-desc">Job details, shipper data, container numbers</div>
    </div>
    """, unsafe_allow_html=True)
    job_file = st.file_uploader("", type="xlsx", key="job", label_visibility="collapsed")
    if job_file:
        st.markdown('<div class="file-loaded-chip">✓ LOADED</div>', unsafe_allow_html=True)
        if st.session_state.show_preview:
            v = validate_file(job_file, "job")
            if v["valid"]:
                st.markdown(f"""
                <div class="preview-box">
                    <div class="pb-row"><span>Rows</span><span class="pb-val">{v['rows']}</span></div>
                    <div class="pb-row"><span>Columns</span><span class="pb-val">{v['cols']}</span></div>
                </div>
                """, unsafe_allow_html=True)
            job_file.seek(0)

    st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)

    # ── E-INVOICE ──
    st.markdown("""
    <div class="upload-item">
        <div class="ui-header">
            <div class="ui-index">03</div>
            <div class="ui-title">E-Invoice Report</div>
        </div>
        <div class="ui-desc">Electronic invoice data — bill numbers and dates</div>
    </div>
    """, unsafe_allow_html=True)
    einv_file = st.file_uploader("", type="xlsx", key="einv", label_visibility="collapsed")
    if einv_file:
        st.markdown('<div class="file-loaded-chip">✓ LOADED</div>', unsafe_allow_html=True)
        if st.session_state.show_preview:
            v = validate_file(einv_file, "einv")
            if v["valid"]:
                st.markdown(f"""
                <div class="preview-box">
                    <div class="pb-row"><span>Rows</span><span class="pb-val">{v['rows']}</span></div>
                    <div class="pb-row"><span>Columns</span><span class="pb-val">{v['cols']}</span></div>
                </div>
                """, unsafe_allow_html=True)
            einv_file.seek(0)

    st.markdown('</div>', unsafe_allow_html=True)  # close card

# ──────────────── RIGHT: CONFIG ────────────────
with right_col:
    st.markdown('<div class="section-label">Configuration</div>', unsafe_allow_html=True)
    st.markdown('<div class="card">', unsafe_allow_html=True)

    # Month selector
    st.markdown('<span class="config-label">Billing Month</span>', unsafe_allow_html=True)
    MONTHS = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"]
    month = st.selectbox("", MONTHS, index=st.session_state.selected_month,
                         label_visibility="collapsed", key="month_select")
    st.session_state.selected_month = MONTHS.index(month)

    st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)

    # Update mode
    st.markdown('<span class="config-label">Update Mode</span>', unsafe_allow_html=True)
    update_mode = st.radio("", ["Append", "Overwrite"],
                           label_visibility="collapsed", key="update_mode_radio")
    if update_mode == "Append":
        st.markdown('<div class="alert-info">ℹ New jobs will be added to existing data</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="alert-warn">⚠ All existing data will be replaced</div>', unsafe_allow_html=True)

    st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)

    # Process button
    st.markdown('<div class="process-btn">', unsafe_allow_html=True)
    process_clicked = st.button("⬡  Execute Processing", key="process_main", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)  # close card

    # ── HISTORY ──
    if st.session_state.processing_history:
        st.markdown('<br/>', unsafe_allow_html=True)
        st.markdown('<div class="section-label">Activity Log</div>', unsafe_allow_html=True)
        st.markdown('<div class="card">', unsafe_allow_html=True)

        for entry in reversed(st.session_state.processing_history[-4:]):
            badge_cls = "ok" if entry["status"] == "success" else "err"
            badge_text = "OK" if entry["status"] == "success" else "ERR"
            st.markdown(f"""
            <div class="history-entry">
                <div class="he-left">
                    <div class="he-time">{entry["time"]} · {entry["month"]}</div>
                    <div class="he-detail">+{entry["added"]} added &nbsp;·&nbsp; {entry["skipped"]} skipped &nbsp;·&nbsp; {entry.get("duration","—")}</div>
                </div>
                <div class="he-badge {badge_cls}">{badge_text}</div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown('</div>', unsafe_allow_html=True)

# ================= CONSTANTS =================
BILL_HEADER_ROW = 3
BILL_DATA_START = 4
SRC_HEADER_ROW  = 7
SRC_DATA_START  = 8

# ================= HELPERS =================
def norm(x):
    return str(x).strip().upper() if x else ""

def clean_job(job):
    return norm(job).split("/")[0]

def parse_date(date_val):
    if not date_val:
        return None
    if isinstance(date_val, datetime):
        return date_val.date()
    date_str = str(date_val).strip()
    if not date_str or date_str.upper() == 'NONE':
        return None
    for fmt in ['%Y-%m-%d','%d-%m-%Y','%d/%m/%Y','%m/%d/%Y','%d.%m.%Y',
                '%Y/%m/%d','%d-%b-%Y','%d %b %Y','%d-%m-%y','%d/%m/%y']:
        try:
            return datetime.strptime(date_str, fmt).date()
        except (ValueError, TypeError):
            continue
    return None

def format_date_for_excel(date_val):
    if not date_val: return None
    if isinstance(date_val, datetime):
        return date_val.strftime('%d/%m/%y')
    try:
        return date_val.strftime('%d/%m/%y')
    except:
        return None

def col_map(ws, header_row):
    m = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v:
            m[norm(v)] = c
    return m

def get_existing_jobs(ws, bill_cols, start_row):
    existing = set()
    for r in range(start_row, ws.max_row + 1):
        job_val = ws.cell(r, bill_cols["JOB NO"]).value
        if job_val:
            existing.add(clean_job(job_val))
    return existing

def find_next_empty_row(ws, start_row):
    for r in range(start_row, ws.max_row + 100):
        if not ws.cell(r, 1).value and not ws.cell(r, 2).value:
            return r
    return ws.max_row + 1

# ================= PROCESSING =================
if process_clicked:
    if not (bill_file and job_file and einv_file):
        st.error("⚠ Please upload all three files to proceed.")
    else:
        progress_bar  = st.progress(0)
        status_text   = st.empty()
        detail_text   = st.empty()
        start_time    = datetime.now()

        try:
            def update_prog(pct, msg, detail=""):
                progress_bar.progress(pct)
                status_text.markdown(f'<div class="status-msg">{msg}</div>', unsafe_allow_html=True)
                if detail:
                    detail_text.markdown(f'<div class="prog-detail"><span>{detail}</span><span>{pct}%</span></div>', unsafe_allow_html=True)

            update_prog(10, "Loading workbooks…", "Reading Excel files")
            bill_wb = load_workbook(bill_file)
            job_wb  = load_workbook(job_file)
            einv_wb = load_workbook(einv_file)

            update_prog(25, "Validating structure…", "Checking sheet names")
            if month not in bill_wb.sheetnames:
                st.error(f"Sheet '{month}' not found in Bill Report.")
                st.stop()

            bill_ws = bill_wb[month]
            job_ws  = job_wb.active
            einv_ws = einv_wb.active

            update_prog(35, "Mapping columns…", "Analysing headers")
            bill_cols = col_map(bill_ws, BILL_HEADER_ROW)
            job_cols  = col_map(job_ws,  SRC_HEADER_ROW)
            einv_cols = col_map(einv_ws, SRC_HEADER_ROW)

            update_prog(45, "Validating columns…", "Checking required fields")
            for col in ["JOB NO","JOB DATE","SHIPPER NAME","INVOICE NO","SB / BE NO","CONTAINER NO"]:
                if col not in job_cols:
                    st.error(f"Missing column in Job Report: {col}")
                    st.stop()
            for col in ["JOB NO","BILL NO","BILL DATE"]:
                if col not in einv_cols:
                    st.error(f"Missing column in E-Invoice: {col}")
                    st.stop()
            for col in ["JOB NO","JOB DATE","PARTY NAME","INVOICE NO","SB / BE NO","CONTAINER NO","BILL NO","BILL DATE"]:
                if col not in bill_cols:
                    st.error(f"Missing column in Bill Report: {col}")
                    st.stop()

            append_mode   = update_mode == "Append"
            existing_jobs = set()

            if append_mode:
                update_prog(55, "Scanning existing data…", "Checking for duplicates")
                existing_jobs = get_existing_jobs(bill_ws, bill_cols, BILL_DATA_START)
                row     = find_next_empty_row(bill_ws, BILL_DATA_START)
                last_sr = 0
                for r in range(BILL_DATA_START, row):
                    sr_val = bill_ws.cell(r, 1).value
                    if sr_val and isinstance(sr_val, (int, float)):
                        last_sr = max(last_sr, int(sr_val))
                sr = last_sr + 1
            else:
                update_prog(55, "Clearing sheet…", "Preparing fresh data")
                for r in range(BILL_DATA_START, bill_ws.max_row + 1):
                    for c in range(1, bill_ws.max_column + 1):
                        bill_ws.cell(r, c).value = None
                row = BILL_DATA_START
                sr  = 1

            update_prog(65, "Reading Job Report…", "Parsing job entries")
            jobs = {}
            for r in range(SRC_DATA_START, job_ws.max_row + 1):
                raw_job = job_ws.cell(r, job_cols["JOB NO"]).value
                if not raw_job: continue
                job_no = clean_job(raw_job)
                if job_no in jobs: continue
                jobs[job_no] = {
                    "JOB DATE":    parse_date(job_ws.cell(r, job_cols["JOB DATE"]).value),
                    "PARTY NAME":  job_ws.cell(r, job_cols["SHIPPER NAME"]).value,
                    "INVOICE NO":  job_ws.cell(r, job_cols["INVOICE NO"]).value,
                    "SB / BE NO":  job_ws.cell(r, job_cols["SB / BE NO"]).value,
                    "CONTAINER NO":job_ws.cell(r, job_cols["CONTAINER NO"]).value,
                }

            update_prog(78, "Reading E-Invoice…", "Parsing invoice entries")
            bill_map = defaultdict(list)
            for r in range(SRC_DATA_START, einv_ws.max_row + 1):
                raw_job  = einv_ws.cell(r, einv_cols["JOB NO"]).value
                bill_no  = einv_ws.cell(r, einv_cols["BILL NO"]).value
                bill_date= parse_date(einv_ws.cell(r, einv_cols["BILL DATE"]).value)
                if raw_job and bill_no:
                    bill_map[clean_job(raw_job)].append((str(bill_no), bill_date))

            update_prog(88, "Writing report…", "Generating output")
            added_count   = 0
            skipped_count = 0

            for job_no, data in jobs.items():
                if append_mode and job_no in existing_jobs:
                    skipped_count += 1
                    continue
                bill_ws.cell(row, 1).value = sr
                bill_ws.cell(row, bill_cols["JOB NO"]).value = job_no
                for field in ["JOB DATE","PARTY NAME","INVOICE NO","SB / BE NO","CONTAINER NO"]:
                    val = data[field]
                    if field == "JOB DATE":
                        val = format_date_for_excel(val)
                    bill_ws.cell(row, bill_cols[field]).value = val
                if job_no in bill_map:
                    bills = bill_map[job_no]
                    if len(bills) == 1:
                        merged = bills[0][0]
                    else:
                        first_bill = str(bills[0][0])
                        if "/" in first_bill:
                            main_part, suffix = first_bill.rsplit("/", 1)
                            suffix = "/" + suffix
                        else:
                            main_part, suffix = first_bill, ""
                        extra = []
                        for i in range(1, len(bills)):
                            bn = str(bills[i][0])
                            if "/" in bn: bn = bn.split("/")[0]
                            extra.append(bn[-4:] if len(bn) >= 4 else bn)
                        merged = main_part + "/" + "/".join(extra) + suffix
                    valid_dates = [b[1] for b in bills if b[1]]
                    latest_date = max(valid_dates) if valid_dates else None
                    bill_ws.cell(row, bill_cols["BILL NO"]).value   = merged
                    bill_ws.cell(row, bill_cols["BILL DATE"]).value = format_date_for_excel(latest_date)
                row += 1
                sr  += 1
                added_count += 1

            update_prog(96, "Saving file…", "Writing to disk")
            out = f"UPDATED_BILL_REPORT_{month}.xlsx"
            bill_wb.save(out)
            progress_bar.progress(100)

            end_time = datetime.now()
            duration = (end_time - start_time).total_seconds()

            status_text.empty()
            detail_text.empty()
            progress_bar.empty()

            st.session_state.processing_history.append({
                "time": end_time.strftime("%H:%M:%S"),
                "month": month, "status": "success",
                "added": added_count, "skipped": skipped_count,
                "duration": f"{duration:.1f}s"
            })

            # ── Inline result section ──
            r1, r2 = st.columns([3, 1], gap="medium")
            with r1:
                st.success(f"Processing complete in {duration:.1f}s")
                st.markdown(f"""
                <div class="metric-row">
                    <div class="metric-card">
                        <div class="metric-val">{added_count}</div>
                        <div class="metric-lbl">Jobs Added</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-val">{skipped_count}</div>
                        <div class="metric-lbl">Skipped</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-val">{added_count + skipped_count}</div>
                        <div class="metric-lbl">Total</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

            with r2:
                st.markdown('<div class="card"><div class="section-label">Export</div>', unsafe_allow_html=True)
                with open(out, "rb") as f:
                    st.download_button("↓ Excel", data=f, file_name=out,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True, key="dl_xl")
                try:
                    df  = pd.read_excel(out, sheet_name=month, header=BILL_HEADER_ROW - 1)
                    csv = df.to_csv(index=False)
                    st.download_button("↓ CSV", data=csv,
                        file_name=f"UPDATED_BILL_REPORT_{month}.csv",
                        mime="text/csv", use_container_width=True, key="dl_csv")
                except:
                    pass
                st.markdown('</div>', unsafe_allow_html=True)

        except Exception as e:
            progress_bar.empty()
            status_text.empty()
            detail_text.empty()
            st.session_state.processing_history.append({
                "time": datetime.now().strftime("%H:%M:%S"),
                "month": month, "status": "error",
                "added": 0, "skipped": 0, "duration": "—"
            })
            st.error(f"Processing failed: {str(e)}")
            with st.expander("Error details"):
                st.exception(e)