"""
Streamlit Application for Universal AI Excel Format Converter.
Provides a modern dashboard UI for uploading any Excel file, configuring AI column mappings,
interpreting natural language prompts, applying templates, and downloading transformed reports.
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
from logger import logger, log_transformation, log_error
from utils import ensure_directories, format_output_filename
from validators import validate_workbook_file, validate_dataframe_headers, validate_column_mapping
from parser import get_workbook_sheet_names, parse_excel_sheet
from mapping_engine import suggest_column_mappings
from prompt_engine import parse_user_prompt, PromptInstructions
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
    .metric-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 16px;
        text-align: center;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05);
    }
    .metric-val {
        color: #2563eb;
        font-size: 26px;
        font-weight: 700;
    }
    .metric-lbl {
        color: #64748b;
        font-size: 12px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
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
    if "manual_mappings" not in st.session_state:
        st.session_state.manual_mappings = {}
    if "user_prompt" not in st.session_state:
        st.session_state.user_prompt = ""
    if "selected_preset" not in st.session_state:
        st.session_state.selected_preset = "Auto Grid"
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
        st.subheader("📥 1. Input Source")

        uploaded_file = st.file_uploader("Upload Input Excel", type=["xlsx", "xls"], help="Upload any input Excel file")

        sample_choice = st.selectbox(
            "Or Choose Sample Dataset",
            options=["None", "OTIS Test Sequence", "Student Marks", "Sales Invoice Data", "Employee Attendance", "Inventory Stock"]
        )

        if sample_choice != "None" and st.session_state.input_file_bytes is None:
            sample_map = {
                "OTIS Test Sequence": "OTIS_Test_Sequence.xlsx",
                "Student Marks": "Student_Marks_Sheet.xlsx",
                "Sales Invoice Data": "Sales_Data.xlsx",
                "Employee Attendance": "Employee_Attendance.xlsx",
                "Inventory Stock": "Inventory_Stock.xlsx"
            }
            s_filename = sample_map[sample_choice]
            s_path = os.path.join(config.SAMPLE_DIR, s_filename)
            if os.path.exists(s_path):
                with open(s_path, "rb") as f:
                    st.session_state.input_file_bytes = f.read()
                st.session_state.input_filename = s_filename
                st.success(f"Loaded '{s_filename}'!")

        if uploaded_file is not None:
            f_bytes = uploaded_file.getvalue()
            if st.session_state.input_file_bytes != f_bytes:
                st.session_state.input_file_bytes = f_bytes
                st.session_state.input_filename = uploaded_file.name
                st.session_state.df_input = None
                st.session_state.transformed_data = None

        st.markdown("### 📄 2. Format Template (Optional)")
        template_file = st.file_uploader("Upload Excel Template", type=["xlsx"], help="Optional template defines formatting, fonts, borders")
        if template_file is not None:
            st.session_state.template_file_bytes = template_file.getvalue()
            st.success("Template loaded!")

        st.markdown("---")
        st.subheader("⚙️ 3. Transformation Configuration")

        preset = st.selectbox("Preset Layout", options=list(config.PRESET_TEMPLATES.keys()), index=0)
        st.session_state.selected_preset = preset

        output_format = st.selectbox("Output Format", options=config.SUPPORTED_OUTPUT_FORMATS, index=0)

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
                Enterprise Data Transformation & Automated Format Conversion System
            </div>
        </div>
    """, unsafe_allow_html=True)

    if st.session_state.input_file_bytes is None:
        st.info("👋 Welcome! Please upload an Excel workbook from the sidebar or choose a sample dataset to get started.")

        # Quick Demo Cards
        st.markdown("### 💡 What would you like to convert today?")
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown("#### 🏢 OTIS Test Sequence")
            st.caption("Converts raw test sequence rows into formatted 5-row inspection blocks with voltage/current limits.")
        with c2:
            st.markdown("#### 🎓 Student Marks")
            st.caption("Maps student marksheets to result sheets with calculated Totals, Percentages, and Grades.")
        with c3:
            st.markdown("#### 🧾 Sales & Billing")
            st.caption("Transforms sales orders into formatted Invoice statements with Subtotal, GST, and Grand Totals.")
        return

    # Parse Excel Sheet if not already parsed
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

    # --- DASHBOARD TABS ---
    tab1, tab2, tab3, tab4 = st.columns([1, 1, 1, 1])

    st.markdown("---")

    # Main Tabs Interface
    main_tab1, main_tab2, main_tab3, main_tab4 = st.tabs([
        "📁 1. Sheet Preview",
        "🤖 2. AI Column Mapping",
        "✍️ 3. Natural Language Prompt",
        "📊 4. Transformed Preview & Export"
    ])

    # --- TAB 1: SHEET PREVIEW ---
    with main_tab1:
        st.markdown("### 📁 Uploaded Sheet Inspector")
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
            st.metric("Total Rows", len(st.session_state.df_input))
        with m3:
            st.metric("Total Columns", len(st.session_state.df_input.columns))

        st.markdown("#### Data Preview (Top 50 Rows)")
        st.dataframe(st.session_state.df_input.head(50), use_container_width=True, height=350)

    # --- TAB 2: AI COLUMN MAPPING ---
    with main_tab2:
        st.markdown("### 🤖 AI Column Mapping Engine")
        st.caption("Automatic AI suggestions mapping input columns to output columns. Override any mapping using the dropdowns.")

        # Interpret Prompt to get target columns
        instructions = parse_user_prompt(st.session_state.user_prompt, default_preset=st.session_state.selected_preset)
        input_cols = list(st.session_state.df_input.columns)
        target_cols = instructions.target_columns

        ai_suggestions = suggest_column_mappings(input_cols, target_cols)

        mapping_table = []
        final_column_map = {}

        for target in target_cols:
            sug_col, conf = ai_suggestions.get(target, ("", 0.0))

            # Selectbox for manual override
            options = ["(Blank / None)"] + input_cols
            default_idx = options.index(sug_col) if sug_col in options else 0

            chosen_col = st.selectbox(
                f"Target: '{target}'",
                options=options,
                index=default_idx,
                key=f"map_select_{target}"
            )

            actual_input = "" if chosen_col == "(Blank / None)" else chosen_col
            final_column_map[target] = actual_input

            confidence_label = f"{int(conf*100)}% Match" if conf > 0 else "Manual / Default"
            mapping_table.append({
                "Target Column": target,
                "Mapped Input Column": actual_input or "(Blank)",
                "AI Confidence": confidence_label
            })

        st.session_state.manual_mappings = final_column_map
        st.dataframe(pd.DataFrame(mapping_table), use_container_width=True)

    # --- TAB 3: PROMPT ENGINE ---
    with main_tab3:
        st.markdown("### ✍️ Natural Language Transformation Prompt")
        st.caption("Describe your desired output layout in plain English.")

        prompt_input = st.text_area(
            "Enter Transformation Prompt",
            value=st.session_state.user_prompt,
            height=120,
            placeholder="Example: Convert to Student Result sheet with columns: Student Name, USN, Total Marks, Percentage, Grade. Calculate Percentage = Total Marks / 500 * 100."
        )
        st.session_state.user_prompt = prompt_input

        # Quick Example Buttons
        st.markdown("#### Quick Example Prompts")
        qp1, qp2, qp3 = st.columns(3)
        with qp1:
            if st.button("Sample: OTIS Inspection Report"):
                st.session_state.user_prompt = "Arrange this Excel into OTIS report format with columns: Sl No, Test Name, Device, Value, Min, Max, Actual. Merge Sl No and Test Name cells."
                st.rerun()
        with qp2:
            if st.button("Sample: Student Result Sheet"):
                st.session_state.user_prompt = "Convert to Student Result sheet with columns: Student Name, USN, Department, Semester, Total Marks, Percentage, Grade. Calculate Percentage = Total Marks / 500 * 100."
                st.rerun()
        with qp3:
            if st.button("Sample: Sales Invoice"):
                st.session_state.user_prompt = "Convert to Invoice format with columns: Invoice No, Item Description, Quantity, Unit Price, Total, GST (18%), Grand Total."
                st.rerun()

        # Parsed Rules Inspector
        if prompt_input:
            p_inst = parse_user_prompt(prompt_input, default_preset=st.session_state.selected_preset)
            with st.expander("📋 View Parsed Prompt Rules", expanded=True):
                st.json(p_inst.to_dict())

    # --- TAB 4: TRANSFORMED PREVIEW & EXPORT ---
    with main_tab4:
        st.markdown("### 📊 Transformed Output & Export")

        if gen_clicked or st.session_state.transformed_data is not None:
            if gen_clicked:
                try:
                    with st.spinner("Executing transformation engine..."):
                        start_t = time.time()
                        instructions = parse_user_prompt(st.session_state.user_prompt, default_preset=st.session_state.selected_preset)
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
                            preset=instructions.preset_name,
                            prompt=st.session_state.user_prompt,
                            input_rows=len(st.session_state.df_input),
                            output_rows=len(t_data.df_grid),
                            elapsed_time=elapsed
                        )
                        st.success(f"🎉 Transformation complete! File saved to `{saved_path}`")

                except Exception as e:
                    st.error(f"❌ Transformation failed: {str(e)}")
                    logger.error("Transformation error in app.py", exc_info=True)
                    return

            if st.session_state.transformed_data is not None:
                t_res = st.session_state.transformed_data

                # Summary Cards
                tc1, tc2, tc3 = st.columns(3)
                with tc1:
                    st.metric("Input Rows", t_res.summary_metrics.get("input_rows", 0))
                with tc2:
                    st.metric("Transformed Rows", t_res.summary_metrics.get("output_rows", 0))
                with tc3:
                    st.metric("Output Preset", t_res.layout_type)

                st.markdown("#### Transformed Data Preview")
                st.dataframe(t_res.df_grid.head(50), use_container_width=True, height=350)

                out_filename = os.path.basename(st.session_state.output_filepath)
                st.download_button(
                    label=f"📥 Download {out_filename}",
                    data=st.session_state.output_bytes,
                    file_name=out_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )


if __name__ == "__main__":
    main()
