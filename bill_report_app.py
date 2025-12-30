import streamlit as st
from openpyxl import load_workbook
from collections import defaultdict
from datetime import datetime
import pandas as pd

# ================= PAGE CONFIG =================
st.set_page_config(
    page_title="Bill Report Automation",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ================= SESSION STATE =================
if 'selected_month' not in st.session_state:
    st.session_state.selected_month = datetime.now().month - 1
if 'processing_history' not in st.session_state:
    st.session_state.processing_history = []
if 'dark_mode' not in st.session_state:
    st.session_state.dark_mode = False
if 'show_preview' not in st.session_state:
    st.session_state.show_preview = False
if 'show_settings' not in st.session_state:
    st.session_state.show_settings = False
if 'show_results' not in st.session_state:
    st.session_state.show_results = False
if 'processing_results' not in st.session_state:
    st.session_state.processing_results = None

# ================= CUSTOM CSS =================
theme = """
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    * {
        font-family: 'Inter', sans-serif;
    }
    
    .main {
        padding: 1.25rem;
        background: #f8f9fa;
    }
    
    /* Header */
    .app-header {
        background: white;
        border-radius: 12px;
        padding: 1.25rem 1.5rem;
        margin-bottom: 1rem;
        border: 1px solid #e5e7eb;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    
    .header-left {
        flex: 1;
    }
    
    .header-title {
        font-size: 1.35rem;
        font-weight: 700;
        color: #111827;
        margin: 0 0 0.15rem 0;
    }
    
    .header-subtitle {
        color: #6b7280;
        font-size: 0.8rem;
        margin: 0;
    }
    
    .header-actions {
        display: flex;
        gap: 0.5rem;
    }
    
    .icon-btn {
        width: 36px;
        height: 36px;
        background: #f3f4f6;
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        cursor: pointer;
        transition: all 0.2s ease;
        font-size: 16px;
    }
    
    .icon-btn:hover {
        background: #e5e7eb;
    }
    
    /* Bento Cards */
    .bento-card {
        background: white;
        border-radius: 12px;
        padding: 1.25rem;
        border: 1px solid #e5e7eb;
        height: 100%;
    }
    
    .bento-card-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 1rem;
        padding-bottom: 0.75rem;
        border-bottom: 1px solid #f3f4f6;
    }
    
    .card-header-left {
        display: flex;
        align-items: center;
        gap: 0.6rem;
    }
    
    .card-icon {
        width: 32px;
        height: 32px;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 16px;
        flex-shrink: 0;
    }
    
    .card-title {
        font-size: 0.95rem;
        font-weight: 600;
        color: #111827;
        margin: 0;
    }
    
    .card-action {
        font-size: 0.75rem;
        color: #667eea;
        cursor: pointer;
        font-weight: 500;
    }
    
    .card-action:hover {
        text-decoration: underline;
    }
    
    /* Upload Items */
    .upload-section {
        display: grid;
        gap: 0.75rem;
    }
    
    .upload-item {
        background: #f9fafb;
        border: 2px solid #e5e7eb;
        border-radius: 10px;
        padding: 1rem;
        transition: all 0.2s ease;
    }
    
    .upload-item:hover {
        border-color: #667eea;
        background: #f0f4ff;
    }
    
    .upload-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 0.5rem;
    }
    
    .upload-left {
        display: flex;
        align-items: center;
        gap: 0.6rem;
    }
    
    .upload-icon {
        width: 28px;
        height: 28px;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 6px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 14px;
        flex-shrink: 0;
    }
    
    .upload-title {
        font-size: 0.85rem;
        font-weight: 600;
        color: #111827;
        margin: 0;
    }
    
    .upload-desc {
        font-size: 0.75rem;
        color: #6b7280;
        margin-bottom: 0.6rem;
    }
    
    .status-badge {
        display: inline-flex;
        align-items: center;
        gap: 0.35rem;
        padding: 0.3rem 0.65rem;
        background: #d1fae5;
        color: #065f46;
        border-radius: 6px;
        font-size: 0.72rem;
        font-weight: 500;
    }
    
    .validation-badge {
        font-size: 0.7rem;
        padding: 0.25rem 0.5rem;
        border-radius: 4px;
        font-weight: 500;
    }
    
    .badge-success {
        background: #d1fae5;
        color: #065f46;
    }
    
    .badge-warning {
        background: #fef3c7;
        color: #92400e;
    }
    
    .badge-error {
        background: #fee2e2;
        color: #991b1b;
    }
    
    /* Preview Box */
    .preview-box {
        background: #f9fafb;
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        padding: 0.75rem;
        margin-top: 0.5rem;
        font-size: 0.75rem;
    }
    
    .preview-header {
        font-weight: 600;
        color: #374151;
        margin-bottom: 0.5rem;
    }
    
    .preview-stats {
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 0.5rem;
    }
    
    .preview-stat {
        display: flex;
        justify-content: space-between;
        color: #6b7280;
    }
    
    .preview-stat-value {
        font-weight: 600;
        color: #111827;
    }
    
    /* Config Section */
    .config-item {
        margin-bottom: 1rem;
    }
    
    .config-item:last-child {
        margin-bottom: 0;
    }
    
    .config-label {
        font-size: 0.8rem;
        font-weight: 600;
        color: #374151;
        margin-bottom: 0.4rem;
        display: block;
    }
    
    /* Alert Boxes */
    .info-alert {
        background: #eff6ff;
        border: 1px solid #bfdbfe;
        border-left: 3px solid #3b82f6;
        padding: 0.6rem 0.8rem;
        border-radius: 6px;
        font-size: 0.75rem;
        color: #1e40af;
        margin-top: 0.4rem;
    }
    
    .warning-alert {
        background: #fef3c7;
        border: 1px solid #fde68a;
        border-left: 3px solid #f59e0b;
        padding: 0.6rem 0.8rem;
        border-radius: 6px;
        font-size: 0.75rem;
        color: #92400e;
        margin-top: 0.4rem;
    }
    
    /* Buttons */
    .stButton>button {
        width: 100%;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        padding: 0.75rem 1.5rem;
        font-size: 0.9rem;
        font-weight: 600;
        border-radius: 10px;
        transition: all 0.2s ease;
        margin-top: 0.75rem;
    }
    
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(102, 126, 234, 0.4);
    }
    
    .stDownloadButton>button {
        width: 100%;
        background: linear-gradient(135deg, #10b981 0%, #059669 100%);
        color: white;
        border: none;
        padding: 0.75rem 1.5rem;
        font-size: 0.9rem;
        font-weight: 600;
        border-radius: 10px;
    }
    
    .stDownloadButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(16, 185, 129, 0.4);
    }
    
    /* Radio buttons */
    .stRadio > div {
        display: flex;
        gap: 0.6rem;
    }
    
    .stRadio > div > label {
        background: #f9fafb;
        padding: 0.6rem 0.85rem;
        border-radius: 6px;
        border: 2px solid #e5e7eb;
        transition: all 0.2s ease;
        cursor: pointer;
        flex: 1;
        text-align: center;
        font-size: 0.8rem;
    }
    
    .stRadio > div > label:hover {
        border-color: #c7d2fe;
        background: #f0f4ff;
    }
    
    .stRadio > div > label:has(input:checked) {
        background: #ede9fe;
        border-color: #667eea;
        font-weight: 600;
    }
    
    /* Select box */
    .stSelectbox > div > div {
        border-radius: 6px;
        border: 2px solid #e5e7eb;
        font-size: 0.85rem;
    }
    
    /* Progress */
    .stProgress > div > div {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        border-radius: 8px;
    }
    
    .progress-detail {
        font-size: 0.75rem;
        color: #6b7280;
        margin-top: 0.25rem;
        display: flex;
        justify-content: space-between;
    }
    
    /* Metrics */
    .metrics-grid {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 0.75rem;
        margin: 1rem 0;
    }
    
    .metric-box {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 10px;
        padding: 1.25rem;
        text-align: center;
    }
    
    .metric-value {
        font-size: 2rem;
        font-weight: 700;
        color: #111827;
        margin-bottom: 0.2rem;
    }
    
    .metric-label {
        font-size: 0.75rem;
        color: #6b7280;
        font-weight: 500;
    }
    
    /* History */
    .history-item {
        background: #f9fafb;
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        padding: 0.75rem;
        margin-bottom: 0.5rem;
        font-size: 0.75rem;
    }
    
    .history-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 0.25rem;
    }
    
    .history-time {
        font-weight: 600;
        color: #111827;
    }
    
    .history-status {
        padding: 0.2rem 0.5rem;
        border-radius: 4px;
        font-size: 0.7rem;
    }
    
    .history-details {
        color: #6b7280;
    }
    
    /* Settings Panel */
    .settings-panel {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 12px;
        padding: 1.25rem;
        margin-bottom: 1rem;
    }
    
    .settings-item {
        margin-bottom: 1rem;
        padding-bottom: 1rem;
        border-bottom: 1px solid #f3f4f6;
    }
    
    .settings-item:last-child {
        margin-bottom: 0;
        padding-bottom: 0;
        border-bottom: none;
    }
    
    .settings-label {
        font-size: 0.85rem;
        font-weight: 600;
        color: #111827;
        margin-bottom: 0.25rem;
    }
    
    .settings-desc {
        font-size: 0.75rem;
        color: #6b7280;
    }
    
    /* File uploader */
    [data-testid="stFileUploader"] {
        background: transparent;
    }
    
    [data-testid="stFileUploader"] section {
        border: 2px dashed #d1d5db;
        border-radius: 6px;
        padding: 0.6rem;
        background: white;
    }
    
    [data-testid="stFileUploader"] section:hover {
        border-color: #667eea;
    }
    
    [data-testid="stFileUploader"] section button {
        font-size: 0.8rem;
        padding: 0.4rem 0.8rem;
    }
    
    [data-testid="stFileUploader"] section small {
        font-size: 0.7rem;
    }
    
    /* Expander */
    .streamlit-expanderHeader {
        font-size: 0.85rem;
        font-weight: 600;
    }
    
    /* Modal/Results styling */
    .success-banner {
        animation: slideDown 0.3s ease-out;
    }
    
    @keyframes slideDown {
        from {
            opacity: 0;
            transform: translateY(-20px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }
    
    .modal-metrics {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 1rem;
        margin-bottom: 1.5rem;
    }
    
    .modal-metric {
        background: #f9fafb;
        border: 1px solid #e5e7eb;
        border-radius: 12px;
        padding: 1.25rem;
        text-align: center;
    }
    
    .modal-metric-value {
        font-size: 2.5rem;
        font-weight: 700;
        color: #111827;
        margin-bottom: 0.25rem;
    }
    
    .modal-metric-label {
        font-size: 0.85rem;
        color: #6b7280;
        font-weight: 500;
    }
    
    .modal-section-title {
        font-size: 1rem;
        font-weight: 600;
        color: #111827;
        margin-bottom: 1rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    
    /* Hide streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stDeployButton {display: none;}
    
    /* Fix container width */
    .element-container {
        width: 100%;
    }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.5rem;
    }
    
    .stTabs [data-baseweb="tab"] {
        font-size: 0.85rem;
        padding: 0.5rem 1rem;
    }
"""

st.markdown(f"<style>{theme}</style>", unsafe_allow_html=True)

# ================= RESULTS BANNER =================
if st.session_state.get('show_results', False):
    results = st.session_state.processing_results
    
    st.markdown("""
    <div class="bento-card success-banner" style="background: linear-gradient(135deg, #d1fae5 0%, #a7f3d0 100%); border-color: #10b981; margin-bottom: 1rem;">
        <div style="text-align: center; padding: 0.5rem 0;">
            <h2 style="color: #065f46; font-size: 1.5rem; margin: 0 0 0.25rem 0;">✅ Processing Complete!</h2>
            <p style="color: #047857; font-size: 0.9rem; margin: 0;">Your report is ready for download</p>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Results metrics and downloads
    result_col1, result_col2 = st.columns([2, 1], gap="medium")
    
    with result_col1:
        st.markdown(f"""
        <div class="bento-card">
            <div class="modal-section-title">📊 Processing Summary</div>
            <div class="modal-metrics" style="margin-bottom: 0;">
                <div class="modal-metric">
                    <div class="modal-metric-value">{results['added']}</div>
                    <div class="modal-metric-label">Jobs Added</div>
                </div>
                <div class="modal-metric">
                    <div class="modal-metric-value">{results['skipped']}</div>
                    <div class="modal-metric-label">Jobs Skipped</div>
                </div>
                <div class="modal-metric">
                    <div class="modal-metric-value">{results['total']}</div>
                    <div class="modal-metric-label">Total Processed</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with result_col2:
        st.markdown("""
        <div class="bento-card">
            <div class="modal-section-title">📥 Download Files</div>
        """, unsafe_allow_html=True)
        
        with open(results['filename'], "rb") as f:
            st.download_button(
                "📥 Download Excel",
                data=f,
                file_name=results['filename'],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="result_download_excel"
            )
        
        if results.get('csv_data'):
            st.download_button(
                "📄 Download CSV",
                data=results['csv_data'],
                file_name=results['csv_filename'],
                mime="text/csv",
                use_container_width=True,
                key="result_download_csv"
            )
        
        if st.button("✕ Dismiss", use_container_width=True, key="dismiss_results"):
            st.session_state.show_results = False
            st.rerun()
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("---")

# ================= HELPER FUNCTIONS =================
def validate_file(file, file_type):
    """Validate uploaded file and return stats"""
    try:
        wb = load_workbook(file)
        ws = wb.active
        rows = ws.max_row
        cols = ws.max_column
        
        # Get headers
        headers = []
        header_row = 7 if file_type in ["job", "einv"] else 3
        for col in range(1, min(cols + 1, 20)):
            val = ws.cell(header_row, col).value
            if val:
                headers.append(str(val).strip())
        
        return {
            "valid": True,
            "rows": rows,
            "cols": cols,
            "headers": headers,
            "sheets": wb.sheetnames if file_type == "bill" else None
        }
    except Exception as e:
        return {
            "valid": False,
            "error": str(e)
        }

# ================= HEADER =================
st.markdown("""
<div class="app-header">
    <div class="header-left">
        <h1 class="header-title">Bill Report Automation</h1>
        <p class="header-subtitle">Streamline your billing workflow with intelligent processing</p>
    </div>
</div>
""", unsafe_allow_html=True)

# Settings toggle
col1, col2 = st.columns([5, 1])
with col2:
    if st.button("⚙️ Settings" if not st.session_state.show_settings else "✕ Close", key="settings_toggle"):
        st.session_state.show_settings = not st.session_state.show_settings

# ================= SETTINGS PANEL =================
if st.session_state.show_settings:
    st.markdown('<div class="settings-panel">', unsafe_allow_html=True)
    
    st.markdown('<div class="settings-item">', unsafe_allow_html=True)
    st.markdown('<div class="settings-label">File Preview</div>', unsafe_allow_html=True)
    st.markdown('<div class="settings-desc">Show file statistics after upload</div>', unsafe_allow_html=True)
    show_preview = st.checkbox("Enable file preview", value=st.session_state.show_preview, key="preview_toggle")
    st.session_state.show_preview = show_preview
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="settings-item">', unsafe_allow_html=True)
    st.markdown('<div class="settings-label">Processing History</div>', unsafe_allow_html=True)
    st.markdown('<div class="settings-desc">Keep track of all processing operations</div>', unsafe_allow_html=True)
    if st.button("Clear History", key="clear_history"):
        st.session_state.processing_history = []
        st.success("History cleared", icon="✓")
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)

# ================= MAIN LAYOUT =================
left_col, right_col = st.columns([2, 1], gap="medium")

# ================= LEFT COLUMN: FILE UPLOADS =================
with left_col:
    st.markdown("""
    <div class="bento-card">
        <div class="bento-card-header">
            <div class="card-header-left">
                <div class="card-icon">📁</div>
                <h2 class="card-title">Upload Documents</h2>
            </div>
        </div>
        <div class="upload-section">
    """, unsafe_allow_html=True)
    
    # Bill Report
    st.markdown("""
        <div class="upload-item">
            <div class="upload-header">
                <div class="upload-left">
                    <div class="upload-icon">📄</div>
                    <h3 class="upload-title">Bill Report</h3>
                </div>
            </div>
            <div class="upload-desc">Main billing report Excel file</div>
    """, unsafe_allow_html=True)
    bill_file = st.file_uploader("", type="xlsx", key="bill", label_visibility="collapsed")
    if bill_file:
        st.markdown('<div class="status-badge">✓ File Loaded</div>', unsafe_allow_html=True)
        if st.session_state.show_preview:
            validation = validate_file(bill_file, "bill")
            if validation["valid"]:
                st.markdown(f"""
                <div class="preview-box">
                    <div class="preview-header">📊 File Statistics</div>
                    <div class="preview-stats">
                        <div class="preview-stat">
                            <span>Total Rows:</span>
                            <span class="preview-stat-value">{validation["rows"]}</span>
                        </div>
                        <div class="preview-stat">
                            <span>Columns:</span>
                            <span class="preview-stat-value">{validation["cols"]}</span>
                        </div>
                        <div class="preview-stat">
                            <span>Sheets:</span>
                            <span class="preview-stat-value">{len(validation["sheets"])}</span>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
            bill_file.seek(0)
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Job Report
    st.markdown("""
        <div class="upload-item">
            <div class="upload-header">
                <div class="upload-left">
                    <div class="upload-icon">💼</div>
                    <h3 class="upload-title">Job Report</h3>
                </div>
            </div>
            <div class="upload-desc">Job details and information file</div>
    """, unsafe_allow_html=True)
    job_file = st.file_uploader("", type="xlsx", key="job", label_visibility="collapsed")
    if job_file:
        st.markdown('<div class="status-badge">✓ File Loaded</div>', unsafe_allow_html=True)
        if st.session_state.show_preview:
            validation = validate_file(job_file, "job")
            if validation["valid"]:
                st.markdown(f"""
                <div class="preview-box">
                    <div class="preview-header">📊 File Statistics</div>
                    <div class="preview-stats">
                        <div class="preview-stat">
                            <span>Total Rows:</span>
                            <span class="preview-stat-value">{validation["rows"]}</span>
                        </div>
                        <div class="preview-stat">
                            <span>Columns:</span>
                            <span class="preview-stat-value">{validation["cols"]}</span>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
            job_file.seek(0)
    st.markdown('</div>', unsafe_allow_html=True)
    
    # E-Invoice Report
    st.markdown("""
        <div class="upload-item">
            <div class="upload-header">
                <div class="upload-left">
                    <div class="upload-icon">📋</div>
                    <h3 class="upload-title">E-Invoice Report</h3>
                </div>
            </div>
            <div class="upload-desc">Electronic invoice data file</div>
    """, unsafe_allow_html=True)
    einv_file = st.file_uploader("", type="xlsx", key="einv", label_visibility="collapsed")
    if einv_file:
        st.markdown('<div class="status-badge">✓ File Loaded</div>', unsafe_allow_html=True)
        if st.session_state.show_preview:
            validation = validate_file(einv_file, "einv")
            if validation["valid"]:
                st.markdown(f"""
                <div class="preview-box">
                    <div class="preview-header">📊 File Statistics</div>
                    <div class="preview-stats">
                        <div class="preview-stat">
                            <span>Total Rows:</span>
                            <span class="preview-stat-value">{validation["rows"]}</span>
                        </div>
                        <div class="preview-stat">
                            <span>Columns:</span>
                            <span class="preview-stat-value">{validation["cols"]}</span>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
            einv_file.seek(0)
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('</div></div>', unsafe_allow_html=True)

# ================= RIGHT COLUMN: CONFIGURATION & PROCESS =================
with right_col:
    # Configuration Card
    st.markdown("""
    <div class="bento-card">
        <div class="bento-card-header">
            <div class="card-header-left">
                <div class="card-icon">⚙️</div>
                <h2 class="card-title">Configuration</h2>
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    # Month Selection
    st.markdown('<div class="config-item">', unsafe_allow_html=True)
    st.markdown('<span class="config-label">Select Month</span>', unsafe_allow_html=True)
    month = st.selectbox(
        "",
        ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"],
        index=st.session_state.selected_month,
        label_visibility="collapsed"
    )
    st.session_state.selected_month = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"].index(month)
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Update Mode
    st.markdown('<div class="config-item">', unsafe_allow_html=True)
    st.markdown('<span class="config-label">Update Mode</span>', unsafe_allow_html=True)
    update_mode = st.radio(
        "",
        ["Append", "Overwrite"],
        label_visibility="collapsed"
    )
    
    if update_mode == "Append":
        st.markdown('<div class="info-alert">ℹ️ New jobs will be added to existing data</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="warning-alert">⚠️ All existing data will be cleared</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Process Button
    process_clicked = st.button("🚀 Process Report", use_container_width=True)
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Processing History
    if st.session_state.processing_history:
        st.markdown("""
        <div class="bento-card" style="margin-top: 1rem;">
            <div class="bento-card-header">
                <div class="card-header-left">
                    <div class="card-icon">📜</div>
                    <h2 class="card-title">Recent Activity</h2>
                </div>
            </div>
        """, unsafe_allow_html=True)
        
        for i, entry in enumerate(reversed(st.session_state.processing_history[-3:])):
            status_color = "badge-success" if entry["status"] == "success" else "badge-error"
            st.markdown(f"""
            <div class="history-item">
                <div class="history-header">
                    <div class="history-time">{entry["time"]}</div>
                    <div class="history-status {status_color}">{entry["status"].upper()}</div>
                </div>
                <div class="history-details">{entry["month"]} • {entry["added"]} added • {entry["skipped"]} skipped</div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)

# ================= CONSTANTS =================
BILL_HEADER_ROW = 3
BILL_DATA_START = 4
SRC_HEADER_ROW = 7
SRC_DATA_START = 8

# ================= HELPER FUNCTIONS =================
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
    date_formats = [
        '%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y', '%m/%d/%Y',
        '%d.%m.%Y', '%Y/%m/%d', '%d-%b-%Y', '%d %b %Y',
        '%d-%m-%y', '%d/%m/%y',
    ]
    for fmt in date_formats:
        try:
            return datetime.strptime(date_str, fmt).date()
        except (ValueError, TypeError):
            continue
    return None

def format_date_for_excel(date_val):
    if not date_val:
        return None
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

# ================= PROCESSING LOGIC =================
if process_clicked:
    if not (bill_file and job_file and einv_file):
        st.error("⚠️ Please upload all three files to continue", icon="⚠️")
    else:
        progress_bar = st.progress(0)
        status_text = st.empty()
        progress_detail = st.empty()
        
        start_time = datetime.now()

        try:
            # Load workbooks
            status_text.info("📂 Loading workbooks...")
            progress_detail.markdown('<div class="progress-detail"><span>Reading Excel files...</span><span>10%</span></div>', unsafe_allow_html=True)
            progress_bar.progress(10)
            
            bill_wb = load_workbook(bill_file)
            job_wb = load_workbook(job_file)
            einv_wb = load_workbook(einv_file)
            
            progress_bar.progress(25)

            if month not in bill_wb.sheetnames:
                st.error(f"❌ Sheet '{month}' not found in Bill Report", icon="❌")
                st.stop()

            bill_ws = bill_wb[month]
            job_ws = job_wb.active
            einv_ws = einv_wb.active

            # Map columns
            status_text.info("🗺️ Mapping columns...")
            progress_detail.markdown('<div class="progress-detail"><span>Analyzing file structure...</span><span>35%</span></div>', unsafe_allow_html=True)
            bill_cols = col_map(bill_ws, BILL_HEADER_ROW)
            job_cols = col_map(job_ws, SRC_HEADER_ROW)
            einv_cols = col_map(einv_ws, SRC_HEADER_ROW)
            progress_bar.progress(35)

            # Validate columns
            status_text.info("✓ Validating structure...")
            progress_detail.markdown('<div class="progress-detail"><span>Checking required columns...</span><span>45%</span></div>', unsafe_allow_html=True)
            for col in ["JOB NO","JOB DATE","SHIPPER NAME","INVOICE NO","SB / BE NO","CONTAINER NO"]:
                if col not in job_cols:
                    st.error(f"❌ Missing column in Job Report: {col}", icon="❌")
                    st.stop()

            for col in ["JOB NO","BILL NO","BILL DATE"]:
                if col not in einv_cols:
                    st.error(f"❌ Missing column in E-Invoice: {col}", icon="❌")
                    st.stop()

            for col in ["JOB NO","JOB DATE","PARTY NAME","INVOICE NO","SB / BE NO","CONTAINER NO","BILL NO","BILL DATE"]:
                if col not in bill_cols:
                    st.error(f"❌ Missing column in Bill Report: {col}", icon="❌")
                    st.stop()

            progress_bar.progress(45)

            # Determine mode
            append_mode = update_mode == "Append"
            existing_jobs = set()
            
            if append_mode:
                status_text.info("🔗 Checking existing data...")
                progress_detail.markdown('<div class="progress-detail"><span>Scanning for duplicates...</span><span>55%</span></div>', unsafe_allow_html=True)
                existing_jobs = get_existing_jobs(bill_ws, bill_cols, BILL_DATA_START)
                row = find_next_empty_row(bill_ws, BILL_DATA_START)
                last_sr = 0
                for r in range(BILL_DATA_START, row):
                    sr_val = bill_ws.cell(r, 1).value
                    if sr_val and isinstance(sr_val, (int, float)):
                        last_sr = max(last_sr, int(sr_val))
                sr = last_sr + 1
            else:
                status_text.info("🔄 Clearing existing data...")
                progress_detail.markdown('<div class="progress-detail"><span>Preparing fresh sheet...</span><span>55%</span></div>', unsafe_allow_html=True)
                for r in range(BILL_DATA_START, bill_ws.max_row + 1):
                    for c in range(1, bill_ws.max_column + 1):
                        bill_ws.cell(r, c).value = None
                row = BILL_DATA_START
                sr = 1

            progress_bar.progress(55)

            # Read Job Report
            status_text.info("📗 Reading Job Report...")
            progress_detail.markdown('<div class="progress-detail"><span>Processing job entries...</span><span>70%</span></div>', unsafe_allow_html=True)
            jobs = {}
            for r in range(SRC_DATA_START, job_ws.max_row + 1):
                raw_job = job_ws.cell(r, job_cols["JOB NO"]).value
                if not raw_job:
                    continue
                job_no = clean_job(raw_job)
                if job_no in jobs:
                    continue
                job_date_raw = job_ws.cell(r, job_cols["JOB DATE"]).value
                job_date = parse_date(job_date_raw)
                jobs[job_no] = {
                    "JOB DATE": job_date,
                    "PARTY NAME": job_ws.cell(r, job_cols["SHIPPER NAME"]).value,
                    "INVOICE NO": job_ws.cell(r, job_cols["INVOICE NO"]).value,
                    "SB / BE NO": job_ws.cell(r, job_cols["SB / BE NO"]).value,
                    "CONTAINER NO": job_ws.cell(r, job_cols["CONTAINER NO"]).value,
                }

            progress_bar.progress(70)

            # Read E-Invoice
            status_text.info("📕 Reading E-Invoice Report...")
            progress_detail.markdown('<div class="progress-detail"><span>Processing invoices...</span><span>80%</span></div>', unsafe_allow_html=True)
            bill_map = defaultdict(list)
            for r in range(SRC_DATA_START, einv_ws.max_row + 1):
                raw_job = einv_ws.cell(r, einv_cols["JOB NO"]).value
                bill_no = einv_ws.cell(r, einv_cols["BILL NO"]).value
                bill_date_raw = einv_ws.cell(r, einv_cols["BILL DATE"]).value
                bill_date = parse_date(bill_date_raw)
                if raw_job and bill_no:
                    bill_map[clean_job(raw_job)].append((str(bill_no), bill_date))

            progress_bar.progress(80)

            # Write Bill Report
            status_text.info("✏️ Writing to Bill Report...")
            progress_detail.markdown('<div class="progress-detail"><span>Generating final report...</span><span>90%</span></div>', unsafe_allow_html=True)
            added_count = 0
            skipped_count = 0

            for job_no, data in jobs.items():
                if append_mode and job_no in existing_jobs:
                    skipped_count += 1
                    continue

                bill_ws.cell(row, 1).value = sr
                bill_ws.cell(row, bill_cols["JOB NO"]).value = job_no

                for field in ["JOB DATE","PARTY NAME","INVOICE NO","SB / BE NO","CONTAINER NO"]:
                    value = data[field]
                    if field == "JOB DATE":
                        value = format_date_for_excel(value)
                    bill_ws.cell(row, bill_cols[field]).value = value

                if job_no in bill_map:
                    bills = bill_map[job_no]
                    
                    # Bill merging logic for format: EUR2501759/25-26
                    if len(bills) == 1:
                        merged = bills[0][0]
                    else:
                        first_bill = str(bills[0][0])
                        
                        # Split first bill to get main part and suffix (e.g., EUR2501759 and /25-26)
                        if "/" in first_bill:
                            main_part, suffix = first_bill.rsplit("/", 1)
                            suffix = "/" + suffix
                        else:
                            main_part = first_bill
                            suffix = ""
                        
                        # Collect last 4 digits from additional bills
                        additional_digits = []
                        for i in range(1, len(bills)):
                            bill_num = str(bills[i][0])
                            # Remove suffix if present
                            if "/" in bill_num:
                                bill_num = bill_num.split("/")[0]
                            # Get last 4 digits
                            last_four = bill_num[-4:] if len(bill_num) >= 4 else bill_num
                            additional_digits.append(last_four)
                        
                        # Merge: main_part/digit1/digit2.../suffix
                        merged = main_part + "/" + "/".join(additional_digits) + suffix
                    
                    valid_dates = [b[1] for b in bills if b[1] is not None]
                    latest_date = max(valid_dates) if valid_dates else None
                    bill_ws.cell(row, bill_cols["BILL NO"]).value = merged
                    bill_ws.cell(row, bill_cols["BILL DATE"]).value = format_date_for_excel(latest_date)

                row += 1
                sr += 1
                added_count += 1

            # Save file
            status_text.info("💾 Saving file...")
            progress_detail.markdown('<div class="progress-detail"><span>Writing to disk...</span><span>100%</span></div>', unsafe_allow_html=True)
            out = f"UPDATED_BILL_REPORT_{month}.xlsx"
            bill_wb.save(out)
            progress_bar.progress(100)
            
            # Calculate processing time
            end_time = datetime.now()
            duration = (end_time - start_time).total_seconds()

            status_text.empty()
            progress_bar.empty()
            progress_detail.empty()
            
            # Add to history
            st.session_state.processing_history.append({
                "time": end_time.strftime("%H:%M:%S"),
                "month": month,
                "status": "success",
                "added": added_count,
                "skipped": skipped_count,
                "duration": f"{duration:.1f}s"
            })

            # Success message with results layout
            result_left, result_right = st.columns([2, 1], gap="medium")
            
            with result_left:
                st.success(f"✅ Processing complete in {duration:.1f} seconds!", icon="✅")
                
                # Results metrics
                st.markdown(f"""
                <div class="metrics-grid">
                    <div class="metric-box">
                        <div class="metric-value">{added_count}</div>
                        <div class="metric-label">Jobs Added</div>
                    </div>
                    <div class="metric-box">
                        <div class="metric-value">{skipped_count}</div>
                        <div class="metric-label">Jobs Skipped</div>
                    </div>
                    <div class="metric-box">
                        <div class="metric-value">{added_count + skipped_count}</div>
                        <div class="metric-label">Total Processed</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
            
            with result_right:
                # Download section
                st.markdown("""
                <div class="bento-card">
                    <div class="bento-card-header">
                        <div class="card-header-left">
                            <div class="card-icon">📥</div>
                            <h2 class="card-title">Download</h2>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
                
                with open(out, "rb") as f:
                    st.download_button(
                        "📥 Download Excel",
                        data=f,
                        file_name=out,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="download_excel"
                    )
                
                # Optional: Generate CSV
                try:
                    df = pd.read_excel(out, sheet_name=month, header=BILL_HEADER_ROW-1)
                    csv = df.to_csv(index=False)
                    st.download_button(
                        "📄 Download CSV",
                        data=csv,
                        file_name=f"UPDATED_BILL_REPORT_{month}.csv",
                        mime="text/csv",
                        use_container_width=True,
                        key="download_csv"
                    )
                except:
                    pass
                
                st.markdown('</div>', unsafe_allow_html=True)

        except Exception as e:
            # Add to history as error
            st.session_state.processing_history.append({
                "time": datetime.now().strftime("%H:%M:%S"),
                "month": month,
                "status": "error",
                "added": 0,
                "skipped": 0,
                "duration": "0s"
            })
            
            st.error(f"❌ Error: {str(e)}", icon="❌")
            with st.expander("View Error Details"):
                st.exception(e)