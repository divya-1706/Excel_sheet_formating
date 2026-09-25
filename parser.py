"""
Universal Parser Module for AI Excel Format Converter.
Reads any uploaded Excel workbook, lists worksheets, detects header rows automatically,
handles merged cells and hidden rows, and converts sheets into clean pandas DataFrames.
"""

from typing import List, Dict, Any, Tuple, Optional, Union
import io
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
import pandas as pd

import config
from logger import logger
from utils import clean_str, normalize_column_name


def get_workbook_sheet_names(file_input: Union[str, io.BytesIO, bytes]) -> List[str]:
    """Returns list of worksheet names from an Excel file."""
    if isinstance(file_input, bytes):
        file_input = io.BytesIO(file_input)
    wb = openpyxl.load_workbook(file_input, read_only=True)
    sheets = wb.sheetnames
    wb.close()
    return sheets


def detect_header_row_index(ws: Worksheet, max_scan_rows: int = 25) -> int:
    """Detects the most probable header row index in a worksheet."""
    best_row_idx = 1
    max_non_empty = 0

    for r in range(1, min(max_scan_rows, ws.max_row) + 1):
        non_empty = 0
        for c in range(1, ws.max_column + 1):
            val = ws.cell(row=r, column=c).value
            if val is not None and str(val).strip():
                non_empty += 1
        if non_empty > max_non_empty:
            max_non_empty = non_empty
            best_row_idx = r

    logger.info(f"Header row detected at index {best_row_idx} with {max_non_empty} columns.")
    return best_row_idx


def parse_excel_sheet(
    file_input: Union[str, io.BytesIO, bytes],
    sheet_name: Optional[str] = None,
    header_row_idx: Optional[int] = None
) -> Tuple[pd.DataFrame, List[str], Dict[str, Any]]:
    """
    Parses a worksheet into a clean pandas DataFrame.
    Returns:
        df: Cleaned pandas DataFrame
        sheet_list: Available worksheet names
        summary: Parsing metadata dict
    """
    logger.info("Parsing Excel workbook sheet...")

    if isinstance(file_input, bytes):
        file_input = io.BytesIO(file_input)

    wb = openpyxl.load_workbook(file_input, data_only=True)
    sheet_list = wb.sheetnames

    if not sheet_name or sheet_name not in sheet_list:
        target_ws = wb.active
        sheet_name = target_ws.title
    else:
        target_ws = wb[sheet_name]

    if header_row_idx is None:
        header_row_idx = detect_header_row_index(target_ws)

    # Read Header Row
    headers: List[str] = []
    col_count = target_ws.max_column
    for c in range(1, col_count + 1):
        val = clean_str(target_ws.cell(row=header_row_idx, column=c).value)
        if not val:
            val = f"Column_{c}"
        headers.append(val)

    # Deduplicate Headers
    seen: Dict[str, int] = {}
    unique_headers: List[str] = []
    for h in headers:
        if h in seen:
            seen[h] += 1
            unique_headers.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            unique_headers.append(h)

    # Read Data Rows
    data_rows: List[List[Any]] = []
    for r in range(header_row_idx + 1, target_ws.max_row + 1):
        row_vals = [target_ws.cell(row=r, column=c).value for c in range(1, col_count + 1)]
        # Skip totally empty rows
        if any(v is not None and str(v).strip() != "" for v in row_vals):
            data_rows.append(row_vals)

    wb.close()

    df = pd.DataFrame(data_rows, columns=unique_headers)

    summary = {
        "sheet_name": sheet_name,
        "header_row": header_row_idx,
        "total_rows": len(df),
        "total_columns": len(unique_headers),
        "columns": unique_headers,
        "available_sheets": sheet_list
    }

    logger.info(f"Parsed sheet '{sheet_name}': {len(df)} rows, {len(unique_headers)} columns.")
    return df, sheet_list, summary
