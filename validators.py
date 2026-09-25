"""
Validators module for Universal AI Excel Format Converter.
Performs pre-flight checks on workbooks, sheets, headers, and column mappings.
Ensures application stability and prevents unexpected runtime failures.
"""

import os
import io
from typing import List, Dict, Any, Tuple, Optional, Union
import openpyxl
import pandas as pd
from logger import logger


def validate_workbook_file(file_input: Union[str, io.BytesIO, bytes], filename: str = "uploaded_file.xlsx") -> Tuple[bool, List[str]]:
    """
    Validates file extension and openpyxl readability.
    Returns (is_valid, error_list).
    """
    errors: List[str] = []

    if isinstance(file_input, str):
        if not os.path.exists(file_input):
            errors.append(f"Input file not found: '{file_input}'")
            return False, errors
        ext = os.path.splitext(file_input)[1].lower()
        if ext not in (".xlsx", ".xls", ".xlsm"):
            errors.append(f"Unsupported file extension '{ext}'. Only .xlsx, .xls, and .xlsm files are supported.")
            return False, errors

    try:
        if isinstance(file_input, bytes):
            file_input = io.BytesIO(file_input)
        wb = openpyxl.load_workbook(file_input, read_only=True)
        sheet_names = wb.sheetnames
        wb.close()
        if not sheet_names:
            errors.append("Workbook contains no worksheets.")
            return False, errors
    except Exception as e:
        logger.error(f"Workbook validation failed for '{filename}'", exc_info=True)
        errors.append(f"Corrupt or invalid Excel workbook: {str(e)}")
        return False, errors

    return True, []


def validate_dataframe_headers(df: pd.DataFrame) -> Tuple[List[str], List[str]]:
    """
    Validates DataFrame structure and detects duplicate or missing header issues.
    Returns (warnings, errors).
    """
    warnings: List[str] = []
    errors: List[str] = []

    if df is None or df.empty:
        errors.append("Uploaded worksheet contains no data rows.")
        return warnings, errors

    columns = [str(c).strip() for c in df.columns]

    # Check duplicate columns
    seen = set()
    duplicates = set()
    for col in columns:
        if col in seen:
            duplicates.add(col)
        else:
            seen.add(col)

    if duplicates:
        warnings.append(f"Duplicate column headers detected: {list(duplicates)}. Automatic suffixes will be applied.")

    # Check unnamed columns
    unnamed = [c for c in columns if "unnamed" in c.lower() or not c]
    if len(unnamed) > len(columns) / 2:
        warnings.append("High number of unnamed columns detected. Please verify header row detection.")

    return warnings, errors


def validate_column_mapping(input_cols: List[str], target_cols: List[str], column_map: Dict[str, str]) -> Tuple[List[str], List[str]]:
    """
    Validates column mapping matrix between input and target format.
    Returns (warnings, errors).
    """
    warnings: List[str] = []
    errors: List[str] = []

    unmapped = [target for target in target_cols if target not in column_map.values() and target not in column_map]
    if unmapped:
        warnings.append(f"The following target columns could not be mapped from input: {unmapped}. They will be populated with default blank values.")

    return warnings, errors
