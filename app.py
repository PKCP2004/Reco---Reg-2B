from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

from reconciliation import export_results, read_upload, reconcile, standardize_frame


st.set_page_config(
    page_title="GST Reconciliation Studio",
    page_icon=str(Path(__file__).parent / "assets" / "PKP_logo.png"),
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    :root {
        --ink: #f4f7fa;
        --muted: #a9b7c4;
        --line: #304250;
        --paper: #101820;
        --crimson: #a90032;
        --cyan: #5bd3cf;
    }
    .stApp {background: #101820; color: var(--ink);}
    [data-testid="stHeader"] {background: rgba(16,24,32,.94);}
    .block-container {max-width: 1440px; padding-top: 1.4rem; padding-bottom: 3rem;}
    h1, h2, h3, h4, p, label, [data-testid="stMetricLabel"] {letter-spacing: 0;}
    h1, h2, h3, h4, p, label, span, small {color: var(--ink);}
    .stMarkdown, .stMarkdown p, [data-testid="stCaptionContainer"] {font-size: .88rem; line-height: 1.5;}
    h2, h3 {color: #ffffff; font-weight: 760;}
    [data-testid="stSidebar"] {background: #17212b; border-right: 1px solid #2b3a48;}
    [data-testid="stSidebar"] * {color: #eaf0f5;}
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {color: #aebdca;}
    [data-testid="stSidebar"] hr {border-color: #344552;}
    [data-testid="stNumberInput"] input, [data-baseweb="select"] > div {
        background: #0d151c !important; color: #ffffff !important;
        border-color: #425565 !important;
    }
    [data-testid="stNumberInput"] button {background: #22313d !important; color: #ffffff !important;}

    .hero {
        position: relative;
        overflow: hidden;
        border-radius: 10px;
        padding: 28px 34px 30px 34px;
        margin-bottom: 18px;
        background: linear-gradient(120deg,#17212b 0%,#243746 70%,#31505a 100%);
        color: #fff;
        border: 1px solid #3c5663;
        box-shadow: 0 12px 28px rgba(20,38,52,.14);
    }
    .hero:after {
        content: "";
        position: absolute;
        width: 340px; height: 340px;
        right: -110px; top: -180px;
        border: 1px solid rgba(91,211,207,.28);
        transform: rotate(45deg);
        box-shadow: 0 0 0 18px rgba(91,211,207,.05), 0 0 0 36px rgba(91,211,207,.035);
        pointer-events: none;
    }
    .brand-row {display:flex; align-items:center; gap:14px; margin-bottom:38px; position:relative; z-index:1;}
    .brand-mark {
        width:50px; height:50px; border-radius:8px;
        display:flex; align-items:center; justify-content:center;
        background:#000; overflow:hidden;
        border:1px solid rgba(255,255,255,.45);
    }
    .brand-mark img {width:100%; height:100%; object-fit:contain;}
    .brand-name {font-size:14px; font-weight:800; letter-spacing:.02em;}
    .brand-sub {font-size:10px; opacity:.9; margin-top:4px;}
    .hero-title {font-size:2rem; font-weight:780; letter-spacing:0; margin-bottom:8px; position:relative; z-index:1;}
    .hero-subtitle {font-size:14px; opacity:.82; position:relative; z-index:1;}

    .section-card {
        border:1px solid var(--line); border-radius:8px; background:#17232d;
        padding:18px 20px; box-shadow:0 3px 12px rgba(0,0,0,.18);
    }
    .section {font-weight:780; font-size:1rem; margin: 8px 0 10px 0; color: #ffffff;}
    [data-testid="stMetric"] {
        border:1px solid var(--line); border-radius:8px;
        padding:13px 14px; background:#17232d; box-shadow: 0 2px 7px rgba(0,0,0,.18);
    }
    [data-testid="stMetricLabel"] {color: var(--muted) !important; font-size: .74rem; text-transform: uppercase;}
    [data-testid="stMetricValue"] {color: #ffffff !important; font-weight: 760;}
    [data-testid="stFileUploader"] {border: 0; background: transparent;}
    [data-testid="stFileUploader"] label {display: block; margin: 0 0 8px 0; color: #ffffff !important; font-weight: 700; line-height: 1.35;}
    [data-testid="stFileUploaderDropzone"] {min-height: 72px; padding: 14px 16px !important; border: 1px dashed #536a7a; border-radius: 8px; background: #17232d; box-sizing: border-box;}
    [data-testid="stFileUploaderDropzoneInstructions"], [data-testid="stFileUploaderDropzoneInstructions"] * {color: #d6e0e7 !important;}
    [data-testid="stFileUploaderDropzone"] button {
        background: #17212b !important; border: 1px solid #17212b !important;
        color: #ffffff !important; font-weight: 700 !important; min-height: 2.1rem;
    }
    [data-testid="stAlert"] {border-radius: 7px; border-left-width: 3px;}
    [data-testid="stAlert"] p, [data-testid="stAlert"] div {color: #edf3f7 !important;}
    [data-testid="stDataFrame"] {border: 1px solid var(--line); border-radius: 8px; overflow: hidden; background: #17232d;}
    [data-testid="stDataFrame"] * {color: #edf3f7;}
    [data-testid="stTabs"] button {font-weight: 650; color: var(--muted);}
    [data-testid="stTabs"] button[aria-selected="true"] {color: var(--crimson);}
    [data-testid="stProgressBar"] > div > div {background: var(--cyan);}
    [data-testid="stProgressBar"] > div {background: #273743;}
    .stButton > button, .stDownloadButton > button {
        border-radius: 6px; min-height: 2.65rem; padding: .55rem 1rem; font-weight: 700;
        border: 1px solid #b7c4cf; box-shadow: 0 2px 5px rgba(20,38,52,.08);
        background: #22313d !important; color: #ffffff !important;
    }
    .stButton > button p, .stDownloadButton > button p {color: inherit !important;}
    .stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
        background: var(--crimson) !important; border-color: var(--crimson) !important; color: #ffffff !important;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {border-color: var(--cyan) !important; color: #087f8c !important;}
    .stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover {background: #7f0027 !important; color: #ffffff !important;}
    [data-testid="stFileDownloadButton"] button {background: #22313d !important; color: #ffffff !important;}
    [data-testid="stFileDownloadButton"] button p {color: #ffffff !important;}
    .hint {
        border-left:3px solid var(--cyan); background:#172a35;
        padding:11px 15px; border-radius:6px; color:#d5e2e9; margin:12px 0 18px 0;
    }
    .stepbar {
        display:flex; align-items:center; justify-content:space-between;
        border:1px solid var(--line); border-radius:8px; background:#17232d;
        padding:13px 20px; margin-bottom:18px; box-shadow: 0 2px 8px rgba(20,38,52,.035);
    }
    .step {display:flex; align-items:center; gap:9px; font-size:12px; color:#d5e0e7; font-weight:650; white-space:nowrap;}
    .step-num {
        width:28px; height:28px; border-radius:50%;
        display:flex; align-items:center; justify-content:center;
        background:#e8f3f4; color:#087f8c; font-weight:800; font-size:12px;
    }
    .step-line {height:1px; background:#dfe3e8; flex:1; margin:0 16px;}
    .footer {text-align:center; color:#7d8b99; font-size:11px; padding-top:24px;}
    .resource-links {display:flex; gap:10px; flex-wrap:wrap; margin-top:10px;}
    .resource-links a {color:#73e0db !important; text-decoration:none; font-weight:700; font-size:12px;}
    .resource-links a:hover {color:#ffffff !important; text-decoration:underline;}
    @media (max-width: 760px) {
        .hero {padding:22px 20px 24px;}
        .hero-title {font-size:1.55rem;}
        .stepbar {overflow-x:auto; justify-content:flex-start; gap:12px;}
        .step-line {min-width:24px; margin:0;}
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
    <div class="brand-row">
        <div class="brand-mark"><img src="data:image/png;base64,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" /></div>
        <div>
            <div class="brand-name">PUSHPAK KUMAR</div>
            <div class="brand-sub">GST Reconciliation &amp; Reporting</div>
        </div>
    </div>
    <div class="hero-title">GST Reconciliation Studio</div>
    <div class="hero-subtitle">Professional GSTR-2B &rarr; Purchase Register reconciliation, review and reporting.</div>
    <div class="resource-links">
        <a href="https://www.pushpakkumar.com" target="_blank">www.pushpakkumar.com</a>
        <a href="https://www.linkedin.com/in/pkcp22/" target="_blank">LinkedIn &middot; Explore more tools</a>
    </div>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### Reconciliation")
    tolerance = st.number_input(
        "Tax value tolerance (₹)",
        min_value=0.0,
        max_value=1000.0,
        value=1.0,
        step=0.50,
        help="Differences up to this amount are treated as within tolerance.",
    )
    st.divider()
    st.markdown("### Matching logic")
    st.caption("1. Exact GSTIN + invoice number")
    st.caption("2. Invoice/tax-value candidate matching")
    st.caption("3. GSTIN mismatch detection")
    st.caption("4. Register-only / 2B-only identification")
    st.divider()
    st.caption("Use the included sample files first to test the complete workflow.")
    st.markdown(
        '<div class="resource-links"><a href="https://www.pushpakkumar.com" target="_blank">Website</a><a href="https://www.linkedin.com/in/pkcp22/" target="_blank">LinkedIn</a></div>',
        unsafe_allow_html=True,
    )

st.markdown("""
<div class="stepbar">
    <div class="step"><span class="step-num">1</span><span>Select Files</span></div>
    <div class="step-line"></div>
    <div class="step"><span class="step-num">2</span><span>Reconcile</span></div>
    <div class="step-line"></div>
    <div class="step"><span class="step-num">3</span><span>Review</span></div>
    <div class="step-line"></div>
    <div class="step"><span class="step-num">4</span><span>Download Report</span></div>
</div>
""", unsafe_allow_html=True)

left, right = st.columns(2, gap="large")

with left:
    st.markdown('<div class="section">GSTR-2B</div>', unsafe_allow_html=True)
    gstr2b_file = st.file_uploader(
        "Upload GSTR-2B Excel/CSV",
        type=["xlsx", "xls", "csv"],
        key="gstr2b",
    )
    if gstr2b_file:
        st.success(f"Loaded: {gstr2b_file.name} ({gstr2b_file.size / 1_048_576:.1f} MB)")

with right:
    st.markdown('<div class="section">Inward Purchase Register</div>', unsafe_allow_html=True)
    register_file = st.file_uploader(
        "Upload purchase register Excel/CSV",
        type=["xlsx", "xls", "csv"],
        key="register",
    )
    if register_file:
        st.success(f"Loaded: {register_file.name} ({register_file.size / 1_048_576:.1f} MB)")

st.markdown("""
<div class="hint">
<b>Recommended columns:</b> GSTIN, Invoice No, Invoice Date, 2B Month/Return Period,
Taxable Value, IGST, CGST, SGST, Cess, Invoice Value.
<br>
Column names can vary — the application automatically recognizes common aliases.
</div>
""", unsafe_allow_html=True)

run = st.button("Run Reconciliation", type="primary", use_container_width=True)

if run:
    if not gstr2b_file or not register_file:
        st.error("Please upload both GSTR-2B and the Inward Purchase Register.")
    else:
        progress = st.progress(0, text="Starting reconciliation…")
        status_box = st.empty()
        try:
            status_box.info("Step 1 of 5 • Reading GSTR-2B file")
            progress.progress(0.08, text="Reading GSTR-2B file…")
            gstr2b_raw = read_upload(gstr2b_file)
            status_box.success(f"GSTR-2B loaded: {len(gstr2b_raw):,} source rows")

            status_box.info("Step 2 of 5 • Reading inward purchase register")
            progress.progress(0.16, text="Reading inward purchase register…")
            register_raw = read_upload(register_file)
            status_box.success(f"Inward register loaded: {len(register_raw):,} source rows")

            status_box.info("Step 3 of 5 • Standardizing columns and values")
            progress.progress(0.24, text="Standardizing columns and values…")
            gstr2b = standardize_frame(gstr2b_raw, "GSTR-2B")
            progress.progress(0.25, text=f"Standardized {len(gstr2b):,} GSTR-2B rows")
            register = standardize_frame(register_raw, "Inward Register")

            def update_progress(value, message):
                progress.progress(min(max(value, 0.0), 1.0), text=message)
                status_box.info(message)

            status_box.info("Step 4 of 5 • Matching invoices and checking exceptions")
            update_progress(0.25, "Starting invoice matching…")
            results = reconcile(gstr2b, register, tolerance=tolerance, progress_callback=update_progress)
            st.session_state["results"] = results

            status_box.info("Step 5 of 5 • Preparing report")
            progress.progress(0.96, text="Preparing professional Excel report…")
            def update_export_status(message):
                status_box.info(f"Step 5 of 5 • {message}")

            st.session_state["excel_bytes"] = export_results(results, progress_callback=update_export_status)
            progress.progress(1.0, text="Completed • Reconciliation report ready")
            status_box.success("Reconciliation completed successfully.")
        except Exception as exc:
            progress.empty()
            status_box.empty()
            st.exception(exc)

results = st.session_state.get("results")

if results is not None:
    reco = results["reconciliation"]

    st.divider()
    st.subheader("Reconciliation Summary")

    matched_count = len(results["matched"])
    mismatch_count = len(results["mismatches"])
    register_only = len(results["missing_in_2b"])
    two_b_only = len(results["missing_in_register"])
    total = len(results["inward_register_source"])

    clean_rate = (matched_count / total * 100) if total else 0

    cols = st.columns(5)
    cols[0].metric("Inward Register Rows", total)
    cols[1].metric("Matched", matched_count)
    cols[2].metric("Review Required", mismatch_count)
    cols[3].metric("Register Only", register_only)
    cols[4].metric("Clean Match Rate", f"{clean_rate:.1f}%")

    if two_b_only:
        st.info(f"{two_b_only:,} GSTR-2B invoice(s) were not found in the inward register.")

    tab1, tab2, tab3, tab4 = st.tabs(
        ["Reconciliation", "Review Queue", "GSTR-2B Data", "Register Data"]
    )

    with tab1:
        st.dataframe(reco, use_container_width=True, hide_index=True)

    with tab2:
        if mismatch_count:
            st.warning(f"{mismatch_count} row(s) need review.")
            statuses = sorted(results["mismatches"]["status"].dropna().unique().tolist())
            selected = st.multiselect(
                "Filter status",
                statuses,
                default=statuses,
            )
            review = results["mismatches"][
                results["mismatches"]["status"].isin(selected)
            ]
            st.dataframe(review, use_container_width=True, hide_index=True)
        else:
            st.success("No mismatch items were identified.")

    with tab3:
        st.dataframe(
            results["gstr2b_source"],
            use_container_width=True,
            hide_index=True,
        )

    with tab4:
        st.dataframe(
            results["inward_register_source"],
            use_container_width=True,
            hide_index=True,
        )

    st.divider()
    excel_bytes = st.session_state.get("excel_bytes") or export_results(results)
    st.download_button(
        "Download Reconciliation Excel",
        data=excel_bytes,
        file_name="GST_Reconciliation_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
else:
    st.info("Upload both files and click **Run Reconciliation** to begin.")

st.divider()
st.subheader("Sample Files")

sample_dir = Path(__file__).parent / "sample_data"
sample_files = [
    ("Sample GSTR-2B", sample_dir / "Sample_GSTR_2B.xlsx"),
    ("Sample Inward Register", sample_dir / "Sample_Inward_Register.xlsx"),
]

c1, c2 = st.columns(2)
for col, (label, path) in zip((c1, c2), sample_files):
    with col:
        st.download_button(
            label,
            data=path.read_bytes(),
            file_name=path.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

st.caption("Tip: run the supplied sample files first. They contain matched, value-mismatch, GSTIN-mismatch and missing-in-source cases.")

st.markdown(
    '<div class="footer">Designed &amp; developed by <b>PUSHPAK KUMAR</b> &nbsp;•&nbsp; GST Reconciliation &amp; Reporting<br><span class="resource-links" style="justify-content:center"><a href="https://www.pushpakkumar.com" target="_blank">www.pushpakkumar.com</a><a href="https://www.linkedin.com/in/pkcp22/" target="_blank">LinkedIn &middot; Explore more tools</a></span></div>',
    unsafe_allow_html=True,
)
