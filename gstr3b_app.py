"""GSTR-3B Consolidation Streamlit App."""

import streamlit as st
import tempfile
import os
from pathlib import Path
from gstr3b_parser import extract_gstr3b_data, month_order
from gstr3b_excel import build_excel, MONTHS

st.set_page_config(
    page_title="GSTR-3B Consolidator",
    page_icon="📊",
    layout="wide",
)

# ── Styling ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1F3864 0%, #2E74B5 100%);
        padding: 20px 30px;
        border-radius: 12px;
        margin-bottom: 24px;
        color: white;
    }
    .main-header h1 { color: white; margin: 0; font-size: 28px; }
    .main-header p  { color: #B8CCE4; margin: 4px 0 0 0; font-size: 14px; }

    .metric-card {
        background: #F0F4F8;
        border-left: 4px solid #2E74B5;
        padding: 14px 18px;
        border-radius: 8px;
        margin-bottom: 10px;
    }
    .metric-card .label { font-size: 12px; color: #555; text-transform: uppercase; font-weight: 600; }
    .metric-card .value { font-size: 22px; color: #1F3864; font-weight: 700; }

    .file-chip {
        display: inline-block;
        background: #E8F1FB;
        border: 1px solid #2E74B5;
        border-radius: 20px;
        padding: 4px 14px;
        font-size: 12px;
        color: #1F3864;
        margin: 3px;
        font-weight: 600;
    }
    .success-box {
        background: #E8F5E9;
        border-left: 4px solid #2E7D32;
        padding: 14px 18px;
        border-radius: 8px;
        color: #1B5E20;
        font-weight: 600;
    }
    .warning-box {
        background: #FFF8E1;
        border-left: 4px solid #F9A825;
        padding: 10px 16px;
        border-radius: 8px;
        color: #795548;
        font-size: 13px;
    }
    div[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }
    .stTabs [data-baseweb="tab"] { font-size: 14px; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>📊 GSTR-3B Consolidator</h1>
    <p>Upload multiple GSTR-3B PDFs → get a fully consolidated Excel workbook (FY-wise, all sections)</p>
</div>
""", unsafe_allow_html=True)

# ── Upload Section ─────────────────────────────────────────────────────────────
with st.container():
    col_up, col_info = st.columns([2, 1])
    with col_up:
        uploaded_files = st.file_uploader(
            "Upload GSTR-3B PDF files (one or more months)",
            type=["pdf"],
            accept_multiple_files=True,
            help="Upload individual GSTR-3B PDFs. You can upload all 12 months at once."
        )
    with col_info:
        st.markdown("""
        **What gets extracted:**
        - Section 3.1 – Outward & RCM Supplies
        - Section 3.1.1 – ECO / Section 9(5)
        - Section 3.2 – Inter-state Supplies
        - Section 4 – Eligible ITC (A, B, C, D)
        - Section 5 & 5.1 – Exempt Inward + Interest
        - Section 6.1 – Payment of Tax
        """)

if not uploaded_files:
    st.markdown("""
    <div class="warning-box">
    ⬆️ Please upload one or more GSTR-3B PDF files to begin. 
    Files should follow the naming pattern: <code>GSTR3B_&lt;GSTIN&gt;_&lt;MMYYYY&gt;.pdf</code>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ── Parse Files ────────────────────────────────────────────────────────────────
all_data = []
errors = []
progress = st.progress(0, text="Reading PDFs…")

with tempfile.TemporaryDirectory() as tmpdir:
    for i, f in enumerate(uploaded_files):
        progress.progress((i + 1) / len(uploaded_files), text=f"Parsing: {f.name}")
        tmp_path = Path(tmpdir) / f.name
        tmp_path.write_bytes(f.read())
        try:
            d = extract_gstr3b_data(str(tmp_path))
            d['filename'] = f.name
            all_data.append(d)
        except Exception as e:
            errors.append(f"{f.name}: {e}")

progress.empty()

if errors:
    with st.expander("⚠️ Parsing errors", expanded=True):
        for e in errors:
            st.error(e)

if not all_data:
    st.error("No data could be extracted. Please check your PDF files.")
    st.stop()

# Sort by financial year order
all_data.sort(key=lambda d: month_order(d['period']))

# ── Summary Metrics ────────────────────────────────────────────────────────────
st.markdown("### 📋 Parsed Summary")

cols = st.columns(4)
with cols[0]:
    st.markdown(f"""<div class="metric-card">
        <div class="label">Files Parsed</div>
        <div class="value">{len(all_data)}</div>
    </div>""", unsafe_allow_html=True)

with cols[1]:
    gstin = all_data[0].get('gstin', 'N/A')
    st.markdown(f"""<div class="metric-card">
        <div class="label">GSTIN</div>
        <div class="value" style="font-size:16px;">{gstin}</div>
    </div>""", unsafe_allow_html=True)

with cols[2]:
    name = all_data[0].get('legal_name', 'N/A')[:24]
    st.markdown(f"""<div class="metric-card">
        <div class="label">Legal Name</div>
        <div class="value" style="font-size:14px;">{name}</div>
    </div>""", unsafe_allow_html=True)

with cols[3]:
    months_parsed = [d['period'].capitalize() for d in all_data]
    st.markdown(f"""<div class="metric-card">
        <div class="label">Months</div>
        <div class="value" style="font-size:15px;">{', '.join(months_parsed)}</div>
    </div>""", unsafe_allow_html=True)

# Month chips
chips = ''.join(f'<span class="file-chip">{d["period"].capitalize()}</span>' for d in all_data)
st.markdown(chips, unsafe_allow_html=True)

# ── Data Preview Tabs ──────────────────────────────────────────────────────────
import pandas as pd

st.markdown("### 🔍 Data Preview")
tabs = st.tabs(["📤 Outward Supplies (3.1)", "💰 Eligible ITC (4)", "🏦 Payment of Tax (6.1)"])

with tabs[0]:
    rows = []
    for d in all_data:
        rows.append({
            'Month': d['period'].capitalize(),
            '(a) Taxable Value': d['3_1a'][0],
            '(a) IGST': d['3_1a'][1],
            '(a) CGST': d['3_1a'][2],
            '(a) SGST': d['3_1a'][3],
            '(b) Zero-Rated TV': d['3_1b'][0],
            '(b) IGST': d['3_1b'][1],
            '(c) Nil/Exempt TV': d['3_1c'][0],
            '(d) RCM TV': d['3_1d'][0],
            '(d) CGST': d['3_1d'][2],
            '(d) SGST': d['3_1d'][3],
        })
    df = pd.DataFrame(rows)
    st.dataframe(df.style.format({c: '{:,.2f}' for c in df.columns if c != 'Month'}),
                 use_container_width=True)

with tabs[1]:
    rows = []
    for d in all_data:
        rows.append({
            'Month': d['period'].capitalize(),
            'A(3) RCM CGST': d['itc_a3'][1],
            'A(3) RCM SGST': d['itc_a3'][2],
            'A(5) Other IGST': d['itc_a5'][0],
            'A(5) Other CGST': d['itc_a5'][1],
            'A(5) Other SGST': d['itc_a5'][2],
            'B(1) Rev IGST': d['itc_b1'][0],
            'B(1) Rev CGST': d['itc_b1'][1],
            'C Net IGST': d['itc_c'][0],
            'C Net CGST': d['itc_c'][1],
            'C Net SGST': d['itc_c'][2],
        })
    df = pd.DataFrame(rows)
    st.dataframe(df.style.format({c: '{:,.2f}' for c in df.columns if c != 'Month'}),
                 use_container_width=True)

with tabs[2]:
    rows = []
    for d in all_data:
        rows.append({
            'Month': d['period'].capitalize(),
            '(A)IGST Payable': d.get('pay_a_igst', [0]*10)[0],
            '(A)IGST Cash': d.get('pay_a_igst', [0]*10)[7],
            '(A)CGST Payable': d.get('pay_a_cgst', [0]*10)[0],
            '(A)CGST Cash': d.get('pay_a_cgst', [0]*10)[7],
            '(A)SGST Payable': d.get('pay_a_sgst', [0]*10)[0],
            '(A)SGST Cash': d.get('pay_a_sgst', [0]*10)[7],
            '(B)CGST RCM': d.get('pay_b_cgst', [0]*10)[0],
            '(B)SGST RCM': d.get('pay_b_sgst', [0]*10)[0],
        })
    df = pd.DataFrame(rows)
    st.dataframe(df.style.format({c: '{:,.2f}' for c in df.columns if c != 'Month'}),
                 use_container_width=True)

# ── Generate Excel ─────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("### 📥 Generate Consolidated Excel")

col_btn, col_opt = st.columns([1, 2])
with col_opt:
    gstin_label = all_data[0].get('gstin', 'GSTIN')
    year_label  = all_data[0].get('year', '2025-26')
    filename = st.text_input("Output filename", value=f"GSTR3B_Consolidated_{gstin_label}_{year_label}.xlsx")

with col_btn:
    generate = st.button("⚡ Generate Excel", type="primary", use_container_width=True)

if generate or st.session_state.get('excel_ready'):
    with st.spinner("Building Excel workbook…"):
        import tempfile, os
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            out_path = tmp.name
        build_excel(all_data, out_path)
        with open(out_path, 'rb') as f:
            excel_bytes = f.read()
        try:
            os.unlink(out_path)
        except Exception:
            pass
        st.session_state['excel_ready'] = True
        st.session_state['excel_bytes'] = excel_bytes
        st.session_state['excel_name'] = filename

    st.markdown("""<div class="success-box">
        ✅ Excel workbook generated successfully! Click below to download.
    </div>""", unsafe_allow_html=True)

    st.download_button(
        label="⬇️  Download Excel Workbook",
        data=st.session_state['excel_bytes'],
        file_name=st.session_state['excel_name'],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.markdown("""
    **Workbook contains 4 sheets:**
    | Sheet | Contents |
    |---|---|
    | Outward Supplies (3.1) | Section 3.1, 3.1.1 & 3.2 — all 12 months + totals |
    | Eligible ITC (4) | ITC Available, Reversed, Net, Other Details |
    | Exempt & Interest (5) | Sec 5 inward + Sec 5.1 interest & late fee |
    | Payment of Tax (6.1) | Section 6.1 (A) & (B) — all 8 tax rows per month |
    """)

# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("GSTR-3B Consolidator • Supports full FY 2025-26 (Apr–Mar) • All amounts in ₹")
