"""
Streamlit Application for OTIS Excel Report Formatter.
Provides a modern dashboard UI for uploading OTIS Test Sequence files,
previewing parsed report blocks, converting to formatted inspection reports,
and downloading the final output.
"""

import os
import sys
import time
import io
import json
import streamlit as st
import pandas as pd

# Ensure base directory in sys.path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import config
from logger import logger, log_upload
from utils import ensure_directories, format_output_filename
from parser import parse_test_sequence_workbook
from formatter import convert_otis_report

# Page Configuration
st.set_page_config(
    page_title="OTIS Excel Report Formatter",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Ensure required directories exist
ensure_directories()

# Custom CSS Styling
CUSTOM_CSS = """
<style>
    /* Global Styling */
    .main {
        background-color: #0f172a;
        color: #f8fafc;
        font-family: 'Inter', system-ui, -apple-system, sans-serif;
    }
    
    /* Header Card */
    .header-card {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 24px;
        margin-bottom: 24px;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.3);
    }
    
    .header-title {
        color: #38bdf8;
        font-size: 28px;
        font-weight: 700;
        margin-bottom: 6px;
    }
    
    .header-subtitle {
        color: #94a3b8;
        font-size: 15px;
    }
    
    /* Metric Cards */
    .metric-container {
        display: flex;
        gap: 16px;
        margin-bottom: 24px;
    }
    
    .metric-card {
        background-color: #1e293b;
        border: 1px solid #334155;
        border-radius: 10px;
        padding: 16px;
        flex: 1;
        text-align: center;
    }
    
    .metric-val {
        color: #38bdf8;
        font-size: 26px;
        font-weight: 700;
    }
    
    .metric-lbl {
        color: #94a3b8;
        font-size: 13px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    /* Status Badges */
    .status-badge-success {
        background-color: #064e3b;
        color: #34d399;
        padding: 4px 10px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 600;
    }
    
    .status-badge-warning {
        background-color: #78350f;
        color: #fbbf24;
        padding: 4px 10px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 600;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def init_session_state():
    """Initializes Streamlit session state variables."""
    if "uploaded_file_bytes" not in st.session_state:
        st.session_state.uploaded_file_bytes = None
    if "uploaded_filename" not in st.session_state:
        st.session_state.uploaded_filename = None
    if "parsed_blocks" not in st.session_state:
        st.session_state.parsed_blocks = None
    if "parse_warnings" not in st.session_state:
        st.session_state.parse_warnings = []
    if "parse_summary" not in st.session_state:
        st.session_state.parse_summary = None
    if "generated_output_path" not in st.session_state:
        st.session_state.generated_output_path = None
    if "generated_bytes" not in st.session_state:
        st.session_state.generated_bytes = None
    if "conversion_summary" not in st.session_state:
        st.session_state.conversion_summary = None


def reset_application():
    """Resets all session state variables."""
    st.session_state.uploaded_file_bytes = None
    st.session_state.uploaded_filename = None
    st.session_state.parsed_blocks = None
    st.session_state.parse_warnings = []
    st.session_state.parse_summary = None
    st.session_state.generated_output_path = None
    st.session_state.generated_bytes = None
    st.session_state.conversion_summary = None
    st.rerun()


def main():
    init_session_state()

    # --- SIDEBAR ---
    with st.sidebar:
        # Display Application Logo
        logo_path = os.path.join(config.ASSETS_DIR, "logo.png")
        if os.path.exists(logo_path):
            st.image(logo_path, use_container_width=True)
        else:
            st.title("OTIS Formatter")

        st.markdown("---")
        st.subheader("⚙️ Controls & Upload")

        uploaded_file = st.file_uploader(
            "Upload OTIS Excel",
            type=["xlsx", "xls"],
            help="Upload OTIS Test Sequence Excel file"
        )

        use_sample = st.checkbox("Use Sample Test Sequence File", value=False)

        if use_sample and st.session_state.uploaded_file_bytes is None:
            sample_path = os.path.join(config.SAMPLE_DIR, "sample_test_sequence.xlsx")
            if os.path.exists(sample_path):
                with open(sample_path, "rb") as f:
                    st.session_state.uploaded_file_bytes = f.read()
                st.session_state.uploaded_filename = "sample_test_sequence.xlsx"
                st.success("Sample file loaded!")

        if uploaded_file is not None:
            file_bytes = uploaded_file.getvalue()
            if st.session_state.uploaded_file_bytes != file_bytes:
                st.session_state.uploaded_file_bytes = file_bytes
                st.session_state.uploaded_filename = uploaded_file.name
                # Reset previous conversion state
                st.session_state.parsed_blocks = None
                st.session_state.generated_output_path = None
                st.session_state.generated_bytes = None
                log_upload(uploaded_file.name, len(file_bytes))

        # Output Folder selector/display
        st.markdown("### 📁 Output Directory")
        output_dir_input = st.text_input("Output Path", value=config.OUTPUT_DIR)

        st.markdown("---")

        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            gen_clicked = st.button("🚀 Generate Report", use_container_width=True, type="primary")
        with col_btn2:
            reset_clicked = st.button("🔄 Reset", use_container_width=True)

        if reset_clicked:
            reset_application()

    # --- MAIN PAGE HEADER ---
    st.markdown("""
        <div class="header-card">
            <div class="header-title">OTIS Excel Report Formatter</div>
            <div class="header-subtitle">
                Automated Enterprise Conversion of OTIS Test Sequence Files into Inspection Reports
            </div>
        </div>
    """, unsafe_allow_html=True)

    if st.session_state.uploaded_file_bytes is None:
        st.info("👋 Please upload an OTIS Test Sequence Excel file from the sidebar or select 'Use Sample Test Sequence File' to test immediately.")

        # Show feature overview card
        with st.expander("📌 Features & Overview", expanded=True):
            st.markdown("""
            - **Automatic Worksheet Detection**: Locates the `Test Sequence` sheet automatically.
            - **Smart Header Mapping**: Identifies column headers regardless of exact position.
            - **Report Block Parsing**: Aggregates test rows by `Sl.No` (1.1, 1.2, 2.1, etc.).
            - **Automated Min/Max Lookups**: Calculates voltage tolerance ranges (18V, 24V, 28V, 30V, 48V, 54V, 60V).
            - **5-Row Report Formatting**: Recreates OTIS standard 5-row structured blocks with merged cell formatting.
            - **Print & Layout Optimization**: Applies Landscape orientation, A4 paper size, freeze top rows, and explicit column widths.
            """)
        return

    # Process file parsing if not already parsed
    if st.session_state.parsed_blocks is None:
        try:
            with st.spinner("Parsing OTIS Test Sequence file..."):
                blocks, warnings, summary = parse_test_sequence_workbook(st.session_state.uploaded_file_bytes)
                st.session_state.parsed_blocks = blocks
                st.session_state.parse_warnings = warnings
                st.session_state.parse_summary = summary
        except Exception as e:
            st.error(f"❌ Parsing Error: {str(e)}")
            logger.error("Parsing failed in app.py", exc_info=True)
            return

    # --- UPLOADED FILE INFO & METRICS ---
    st.markdown("### 📄 Uploaded File Summary")
    f_col1, f_col2, f_col3, f_col4 = st.columns(4)
    with f_col1:
        st.metric("Filename", st.session_state.uploaded_filename)
    with f_col2:
        st.metric("Total Reports", len(st.session_state.parsed_blocks))
    with f_col3:
        st.metric("Parsed Rows", st.session_state.parse_summary.get("total_rows_parsed", 0))
    with f_col4:
        status_str = "Warnings Found" if st.session_state.parse_warnings else "Valid"
        st.metric("Status", status_str)

    if st.session_state.parse_warnings:
        with st.expander("⚠️ Parsing Warnings", expanded=False):
            for w in st.session_state.parse_warnings:
                st.warning(w)

    # --- GENERATE REPORT ACTION ---
    if gen_clicked:
        progress_bar = st.progress(0)
        status_text = st.empty()

        try:
            status_text.text("Step 1/3: Reading and validating report blocks...")
            progress_bar.progress(30)
            time.sleep(0.2)

            status_text.text("Step 2/3: Constructing 5-row report layout & openpyxl styles...")
            progress_bar.progress(70)

            out_path = os.path.join(output_dir_input, format_output_filename(st.session_state.uploaded_filename))

            saved_path, blocks, conv_summary = convert_otis_report(
                input_file=st.session_state.uploaded_file_bytes,
                output_file_path=out_path,
                original_filename=st.session_state.uploaded_filename
            )

            with open(saved_path, "rb") as f:
                st.session_state.generated_bytes = f.read()

            st.session_state.generated_output_path = saved_path
            st.session_state.conversion_summary = conv_summary

            progress_bar.progress(100)
            status_text.text("Conversion complete!")
            st.success(f"🎉 Report generated successfully! Saved to: `{saved_path}`")

        except Exception as e:
            st.error(f"❌ Conversion failed: {str(e)}")
            logger.error("Conversion failed in app.py", exc_info=True)

    # --- DISPLAY GENERATED REPORT METRICS & DOWNLOAD ---
    if st.session_state.conversion_summary:
        st.markdown("---")
        st.markdown("### 📊 Report Generation Summary")

        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            st.metric("Total Reports", st.session_state.conversion_summary["total_reports"])
        with c2:
            st.metric("Rows Written", st.session_state.conversion_summary["total_rows_written"])
        with c3:
            st.metric("Voltage Rows", st.session_state.conversion_summary["voltage_rows"])
        with c4:
            st.metric("Current Rows", st.session_state.conversion_summary["current_rows"])
        with c5:
            st.metric("Time (s)", st.session_state.conversion_summary["processing_time_seconds"])

        out_name = format_output_filename(st.session_state.uploaded_filename)
        st.download_button(
            label=f"📥 Download {out_name}",
            data=st.session_state.generated_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    # --- PREVIEW TABLE ---
    st.markdown("---")
    st.markdown("### 🔍 Parsed Reports Preview")

    table_data = []
    for b in st.session_state.parsed_blocks:
        table_data.append({
            "Report": str(b.report_no or ""),
            "Test Name": str(b.test_name or ""),
            "Voltage": str(b.voltage_raw or (f"{b.voltage}V" if b.voltage is not None else "")),
            "Current": str(b.current_raw or (f"{b.current}A" if b.current is not None else "")),
            "Point": str(b.measurement_point or ""),
            "PSU Device": str(b.psu_device or ""),
            "Min": str(b.min_voltage if b.min_voltage is not None else "-"),
            "Max": str(b.max_voltage if b.max_voltage is not None else "-"),
            "Status": "Warnings" if b.warnings else "OK"
        })

    df_preview = pd.DataFrame(table_data).astype(str)
    st.dataframe(df_preview, use_container_width=True, height=350)

    # --- JSON SUMMARY VIEWER ---
    with st.expander("📋 View Parsed Reports JSON Summary"):
        json_blocks = [b.to_dict() for b in st.session_state.parsed_blocks]
        st.json(json_blocks)


if __name__ == "__main__":
    main()
