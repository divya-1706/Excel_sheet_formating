"""
Universal Formatter Module for AI Excel Format Converter.
Generates production-quality openpyxl Excel workbooks, CSV, and ODS exports.
Applies styles, borders, column width auto-fitting, row heights, and print-ready setup.
"""

from typing import Union, Optional, List, Dict, Any
import os
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
import pandas as pd

import config
from logger import logger
from styles import (
    HEADER_BLUE_FONT, HEADER_FONT, BODY_FONT, BODY_BOLD_FONT,
    HEADER_BLUE_FILL, HEADER_FILL, THIN_BORDER, HEADER_BORDER,
    HEADER_ALIGNMENT, BODY_LEFT_ALIGNMENT, BODY_CENTER_ALIGNMENT, BODY_RIGHT_ALIGNMENT,
    apply_range_borders
)
from template_engine import TemplateEngine
from transformer import TransformedData


def build_universal_workbook(
    transformed_data: TransformedData,
    template_engine: Optional[TemplateEngine] = None
) -> openpyxl.Workbook:
    """
    Builds styled openpyxl Workbook from TransformedData container.
    Clones template formatting if TemplateEngine is provided.
    """
    logger.info("Building universal openpyxl workbook...")

    wb = openpyxl.Workbook()
    ws: Worksheet = wb.active
    ws.title = "Formatted Report"

    df = transformed_data.df_grid

    # Write Header Row
    headers = list(df.columns)
    ws.append(headers)
    ws.row_dimensions[1].height = config.HEADER_ROW_HEIGHT

    for col_idx, col_name in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = HEADER_BLUE_FONT
        cell.fill = HEADER_BLUE_FILL
        cell.border = HEADER_BORDER
        cell.alignment = HEADER_ALIGNMENT

    # Write Data Rows
    current_row = 2
    for _, row in df.iterrows():
        ws.row_dimensions[current_row].height = config.DATA_ROW_HEIGHT
        for col_idx, val in enumerate(row, 1):
            cell_val = "" if pd.isna(val) else val
            cell = ws.cell(row=current_row, column=col_idx, value=cell_val)
            cell.font = BODY_FONT
            cell.border = THIN_BORDER

            # Alignment heuristics
            if isinstance(cell_val, (int, float)):
                cell.alignment = BODY_RIGHT_ALIGNMENT
            elif str(cell_val).strip().isdigit():
                cell.alignment = BODY_CENTER_ALIGNMENT
            else:
                cell.alignment = BODY_LEFT_ALIGNMENT

        current_row += 1

    # Apply Auto Column Widths
    for col_idx, col_name in enumerate(headers, 1):
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        max_len = max(len(str(col_name)), 12)
        for r in range(2, min(current_row, 100)):
            val = str(ws.cell(row=r, column=col_idx).value or "")
            if len(val) > max_len:
                max_len = len(val)
        ws.column_dimensions[col_letter].width = min(max_len + 4, 45)

    # Apply Template styling if loaded
    if template_engine and template_engine.is_loaded():
        template_engine.apply_template_to_worksheet(ws)

    # Page Setup & Print Formatting
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Freeze panes on Header row
    ws.freeze_panes = "A2"

    # Repeat header row on print
    ws.print_title_rows = "1:1"

    # Show Gridlines
    if ws.views.sheetView:
        ws.views.sheetView[0].showGridLines = True

    logger.info(f"Universal workbook constructed successfully. Total rows: {current_row - 1}")
    return wb


def export_transformed_data(
    transformed_data: TransformedData,
    output_filepath: str,
    output_format: str = "XLSX",
    template_engine: Optional[TemplateEngine] = None
) -> str:
    """Exports transformed dataset to specified output file format (XLSX, CSV, ODS)."""
    fmt = output_format.upper()
    logger.info(f"Exporting data to format '{fmt}' at '{output_filepath}'...")

    if fmt == "CSV":
        transformed_data.df_grid.to_csv(output_filepath, index=False)
    elif fmt == "ODS":
        transformed_data.df_grid.to_excel(output_filepath, index=False, engine="openpyxl")
    else:
        # Default XLSX with openpyxl styling
        wb = build_universal_workbook(transformed_data, template_engine)
        wb.save(output_filepath)

    logger.info(f"Export complete: '{output_filepath}'")
    return output_filepath
