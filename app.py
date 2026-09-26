"""
Streamlit Application for Universal AI Excel Format Converter.
Allows users to upload any Excel workbook, enter custom output column names manually,
map input columns, specify formatting rules, and generate formatted Excel output sheets.
"""

import os
import sys
import time
import io
import streamlit as st
import pandas as pd

# Ensure base directory in sys.path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import config
from logger import logger, log_transformation, log_error
from utils import ensure_directories, format_output_filename
from validators import validate_workbook_file
from parser import get_workbook_sheet_names, parse_excel_sheet
from mapping_engine import suggest_column_mappings
from prompt_engine import parse_user_prompt, parse_custom_columns, PromptInstructions
from transformer import transform_dataset, TransformedData
from template_engine import TemplateEngine
from formatter import export_transformed_data

# Page Configuration
st.set_page_config(
    page_title="Universal AI Excel Format Converter",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

ensure_directories()

# Custom White and Blue Professional Theme CSS
CUSTOM_CSS = """
<style>
    .main {
        background-color: #f8fafc;
        color: #1e293b;
        font-family: 'Inter', system-ui, -apple-system, sans-serif;
    }
    .header-card {
        background: linear-gradient(135deg, #1e40af 0%, #1e3a8a 100%);
        border-radius: 14px;
        padding: 24px;
        color: #ffffff;
        margin-bottom: 24px;
        box-shadow: 0 10px 15px -3px rgba(30, 64, 175, 0.2);
    }
    .header-title {
        font-size: 28px;
        font-weight: 700;
        margin-bottom: 6px;
        color: #ffffff;
    }
    .header-subtitle {
        font-size: 15px;
        color: #93c5fd;
    }
    .section-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 20px;
        margin-bottom: 20px;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05);
    }
    .section-title {
        color: #1e40af;
        font-size: 18px;
        font-weight: 700;
        margin-bottom: 12px;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def init_session_state():
    if "input_file_bytes" not in st.session_state:
        st.session_state.input_file_bytes = None
    if "input_filename" not in st.session_state:
        st.session_state.input_filename = None
    if "template_file_bytes" not in st.session_state:
        st.session_state.template_file_bytes = None
    if "selected_sheet" not in st.session_state:
        st.session_state.selected_sheet = None
    if "df_input" not in st.session_state:
        st.session_state.df_input = None
    if "input_summary" not in st.session_state:
        st.session_state.input_summary = None
    if "custom_columns_text" not in st.session_state:
        st.session_state.custom_columns_text = "Test Name, Max Value, Min Value, Device, Value, Actual"
    if "manual_mappings" not in st.session_state:
        st.session_state.manual_mappings = {}
    if "user_prompt" not in st.session_state:
        st.session_state.user_prompt = ""
    if "transformed_data" not in st.session_state:
        st.session_state.transformed_data = None
    if "output_bytes" not in st.session_state:
        st.session_state.output_bytes = None
    if "output_filepath" not in st.session_state:
        st.session_state.output_filepath = None


def reset_app():
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()


def main():
    init_session_state()

    # --- SIDEBAR ---
    with st.sidebar:
        logo_path = os.path.join(config.ASSETS_DIR, "logo.png")
        if os.path.exists(logo_path):
            st.image(logo_path, use_container_width=True)
        else:
            st.title("Universal Converter")

        st.markdown("---")
        st.subheader("📥 1. Input Excel File")
        uploaded_file = st.file_uploader("Upload Input Workbook", type=["xlsx", "xls"], help="Upload any Excel file to convert")

        if uploaded_file is not None:
            f_bytes = uploaded_file.getvalue()
            if st.session_state.input_file_bytes != f_bytes:
                st.session_state.input_file_bytes = f_bytes
                st.session_state.input_filename = uploaded_file.name
                st.session_state.df_input = None
                st.session_state.transformed_data = None

        st.markdown("### 📄 2. Format Template (Optional)")
        template_file = st.file_uploader("Upload Excel Template", type=["xlsx"], help="Optional template defines fonts, borders, fills")
        if template_file is not None:
            st.session_state.template_file_bytes = template_file.getvalue()
            st.success("Template loaded!")

        st.markdown("---")
        st.subheader("⚙️ 3. Output Format")
        output_format = st.selectbox("Choose Export Format", options=config.SUPPORTED_OUTPUT_FORMATS, index=0)

        st.markdown("---")
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            gen_clicked = st.button("🚀 Convert", use_container_width=True, type="primary")
        with col_btn2:
            reset_clicked = st.button("🔄 Reset", use_container_width=True)

        if reset_clicked:
            reset_app()

    # --- MAIN PAGE HEADER ---
    st.markdown("""
        <div class="header-card">
            <div class="header-title">Universal AI Excel Format Converter</div>
            <div class="header-subtitle">
                Enter your desired custom output column names and convert any Excel file automatically
            </div>
        </div>
    """, unsafe_allow_html=True)

    if st.session_state.input_file_bytes is None:
        st.info("👋 Please upload an Excel workbook from the sidebar to begin.")

        with st.expander("📌 How to Use", expanded=True):
            st.markdown("""
            1. **Upload your Excel workbook** in the sidebar.
            2. **Enter your desired output column names** (e.g. `Test Name, Max Value, Min Value, Device, Value, Actual`).
            3. **Review & customize AI column mappings** to map your uploaded columns to your new output columns.
            4. **Click 'Convert'** to generate and download your formatted Excel output sheet!
            """)
        return

    # Parse Excel Sheet if not already loaded
    if st.session_state.df_input is None:
        try:
            is_valid, errs = validate_workbook_file(st.session_state.input_file_bytes, st.session_state.input_filename)
            if not is_valid:
                for e in errs:
                    st.error(e)
                return

            sheet_list = get_workbook_sheet_names(st.session_state.input_file_bytes)
            selected_sheet = sheet_list[0]

            df, _, summary = parse_excel_sheet(st.session_state.input_file_bytes, sheet_name=selected_sheet)
            st.session_state.df_input = df
            st.session_state.selected_sheet = selected_sheet
            st.session_state.input_summary = summary
        except Exception as e:
            st.error(f"❌ Failed to parse Excel file: {str(e)}")
            logger.error("Excel parsing failed in app.py", exc_info=True)
            return

    # --- STEP 1: SHEET INSPECTOR ---
    st.markdown("<div class='section-card'><div class='section-title'>📁 Step 1: Input Worksheet Inspector</div>", unsafe_allow_html=True)
    sheet_list = st.session_state.input_summary.get("available_sheets", [st.session_state.selected_sheet])

    if len(sheet_list) > 1:
        chosen_sheet = st.selectbox("Select Worksheet", options=sheet_list, index=sheet_list.index(st.session_state.selected_sheet))
        if chosen_sheet != st.session_state.selected_sheet:
            df, _, summary = parse_excel_sheet(st.session_state.input_file_bytes, sheet_name=chosen_sheet)
            st.session_state.df_input = df
            st.session_state.selected_sheet = chosen_sheet
            st.session_state.input_summary = summary
            st.rerun()

    m1, m2, m3 = st.columns(3)
    with m1:
        st.metric("Worksheet", st.session_state.selected_sheet)
    with m2:
        st.metric("Total Input Rows", len(st.session_state.df_input))
    with m3:
        st.metric("Total Input Columns", len(st.session_state.df_input.columns))

    with st.expander("🔍 View Input Sheet Data (Top 50 Rows)", expanded=False):
        st.dataframe(st.session_state.df_input.head(50), use_container_width=True, height=250)
    st.markdown("</div>", unsafe_allow_html=True)

    # --- STEP 2: CUSTOM COLUMN CONFIGURATION ---
    st.markdown("<div class='section-card'><div class='section-title'>✏️ Step 2: Enter Your Desired Output Columns</div>", unsafe_allow_html=True)
    st.caption("Type the exact column names you want to display in your output Excel format sheet (separated by commas).")

    cols_input = st.text_area(
        "Output Column Names (Comma-separated)",
        value=st.session_state.custom_columns_text,
        height=70,
        help="Type your exact custom column names, e.g.: Test Name, Max Value, Min Value, Device, Value, Actual"
    )
    st.session_state.custom_columns_text = cols_input

    # Quick Preset Column Fillers
    st.markdown("**Quick Preset Column Templates (Click to fill):**")
    qc1, qc2, qc3, qc4 = st.columns(4)
    with qc1:
        if st.button("📋 Test / Inspection Columns"):
            st.session_state.custom_columns_text = "Sl No, Test Name, Device, Value, Min Value, Max Value, Actual"
            st.rerun()
    with qc2:
        if st.button("🎓 Student Marks Columns"):
            st.session_state.custom_columns_text = "Student Name, USN, Department, Semester, Total Marks, Percentage, Grade"
            st.rerun()
    with qc3:
        if st.button("🧾 Invoice / Billing Columns"):
            st.session_state.custom_columns_text = "Invoice No, Item Description, Quantity, Unit Price, Total, GST (18%), Grand Total"
            st.rerun()
    with qc4:
        if st.button("📦 Inventory BOM Columns"):
            st.session_state.custom_columns_text = "Part No, Part Description, Quantity, Supplier, Unit Cost, Total Value"
            st.rerun()

    target_columns = parse_custom_columns(st.session_state.custom_columns_text)
    st.markdown(f"**Detected Output Columns ({len(target_columns)}):** " + ", ".join([f"`{c}`" for c in target_columns]))
    st.markdown("</div>", unsafe_allow_html=True)

    # --- STEP 3: AI COLUMN MAPPING MATRIX ---
    st.markdown("<div class='section-card'><div class='section-title'>🤖 Step 3: Map Input Columns to Output Columns</div>", unsafe_allow_html=True)
    st.caption("AI has auto-suggested mappings below. You can override any column using the dropdown selection.")

    input_cols = list(st.session_state.df_input.columns)
    ai_suggestions = suggest_column_mappings(input_cols, target_columns)

    final_column_map = {}
    mapping_rows = []

    grid_cols = st.columns(min(len(target_columns), 3) or 1)
    for idx, target in enumerate(target_columns):
        c_idx = idx % len(grid_cols)
        sug_col, conf = ai_suggestions.get(target, ("", 0.0))

        options = ["(Blank / Formula)"] + input_cols
        default_idx = options.index(sug_col) if sug_col in options else 0

        with grid_cols[c_idx]:
            chosen_col = st.selectbox(
                f"Output Column: **{target}**",
                options=options,
                index=default_idx,
                key=f"map_select_{target}"
            )
            actual_input = "" if chosen_col == "(Blank / Formula)" else chosen_col
            final_column_map[target] = actual_input

    st.session_state.manual_mappings = final_column_map
    st.markdown("</div>", unsafe_allow_html=True)

    # --- STEP 4: OPTIONAL RULES & FORMULAS ---
    st.markdown("<div class='section-card'><div class='section-title'>✍️ Step 4: Optional Transformation Rules (Optional)</div>", unsafe_allow_html=True)
    prompt_input = st.text_input(
        "Enter optional calculation or layout instructions (optional)",
        value=st.session_state.user_prompt,
        placeholder="Example: Merge Sl No and Test Name. Calculate Percentage = Total Marks / 500 * 100."
    )
    st.session_state.user_prompt = prompt_input
    st.markdown("</div>", unsafe_allow_html=True)

    # --- STEP 5: TRANSFORMED PREVIEW & EXPORT ---
    st.markdown("<div class='section-card'><div class='section-title'>📊 Step 5: Transformed Output Preview & Download</div>", unsafe_allow_html=True)

    if gen_clicked or st.session_state.transformed_data is not None:
        if gen_clicked:
            try:
                with st.spinner("Building custom Excel output format..."):
                    start_t = time.time()
                    instructions = parse_user_prompt(st.session_state.user_prompt, custom_columns_text=st.session_state.custom_columns_text)
                    col_map = st.session_state.manual_mappings

                    t_data = transform_dataset(st.session_state.df_input, col_map, instructions)

                    tmpl_engine = TemplateEngine(st.session_state.template_file_bytes) if st.session_state.template_file_bytes else None

                    out_name = format_output_filename(st.session_state.input_filename, output_format.lower())
                    out_path = os.path.join(config.OUTPUT_DIR, out_name)

                    saved_path = export_transformed_data(t_data, out_path, output_format, tmpl_engine)

                    with open(saved_path, "rb") as f:
                        st.session_state.output_bytes = f.read()

                    st.session_state.output_filepath = saved_path
                    st.session_state.transformed_data = t_data

                    elapsed = time.time() - start_t
                    log_transformation(
                        filename=st.session_state.input_filename,
                        preset="Custom Columns",
                        prompt=st.session_state.custom_columns_text,
                        input_rows=len(st.session_state.df_input),
                        output_rows=len(t_data.df_grid),
                        elapsed_time=elapsed
                    )
                    st.success(f"🎉 Report generated successfully! Saved to `{saved_path}`")

            except Exception as e:
                st.error(f"❌ Transformation failed: {str(e)}")
                logger.error("Transformation error in app.py", exc_info=True)
                return

        if st.session_state.transformed_data is not None:
            t_res = st.session_state.transformed_data

            tc1, tc2, tc3 = st.columns(3)
            with tc1:
                st.metric("Total Output Rows", len(t_res.df_grid))
            with tc2:
                st.metric("Total Output Columns", len(t_res.df_grid.columns))
            with tc3:
                st.metric("Output File Format", output_format)

            st.markdown("#### Formatted Output Data Table Preview")

            # Convert preview table data to string object types to avoid PyArrow display errors
            df_preview_disp = t_res.df_grid.head(50).astype(str)
            st.dataframe(df_preview_disp, use_container_width=True, height=350)

            out_filename = os.path.basename(st.session_state.output_filepath)
            st.download_button(
                label=f"📥 Download {out_filename}",
                data=st.session_state.output_bytes,
                file_name=out_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
    st.markdown("</div>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
