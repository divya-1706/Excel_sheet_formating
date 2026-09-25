"""
Utility functions for OTIS Excel Report Formatter.
Includes string parsing, number extraction, voltage lookup calculations,
and filesystem helper utilities.
"""

import os
import re
import math
from typing import Optional, Tuple, Any
import config
from logger import logger


def ensure_directories() -> None:
    """Ensures all required project directories exist."""
    directories = [
        config.INPUT_DIR,
        config.OUTPUT_DIR,
        config.LOG_DIR,
        config.SAMPLE_DIR,
        config.ASSETS_DIR,
    ]
    for d in directories:
        os.makedirs(d, exist_ok=True)


def clean_str(val: Any) -> str:
    """
    Cleans an input value into a stripped string.
    Returns empty string if value is None or NaN.
    """
    if val is None:
        return ""
    if isinstance(val, float) and math.isnan(val):
        return ""
    s = str(val).strip()
    if s.lower() in ("none", "nan", "null"):
        return ""
    return s


def extract_number(text: str) -> Optional[float]:
    """
    Extracts the first floating point or integer number from a string.
    Examples:
        '24V' -> 24.0
        'SET CURRENT 1.5A' -> 1.5
        'P30V' -> 30.0
    """
    if not text:
        return None
    # Match numbers like 24, 24.5, 1.5
    match = re.search(r"[-+]?\d*\.?\d+", text)
    if match:
        try:
            val = float(match.group(0))
            return val
        except ValueError:
            return None
    return None


def extract_voltage(text_or_val: Any) -> Optional[float]:
    """
    Extracts voltage float value from raw text or numeric cell.
    Supports patterns like:
    - 'SET Voltage 24V' -> 24.0
    - 'SET Voltage 18' -> 18.0
    - '24V' -> 24.0
    - 24 -> 24.0
    """
    if text_or_val is None:
        return None
    if isinstance(text_or_val, (int, float)):
        if isinstance(text_or_val, float) and math.isnan(text_or_val):
            return None
        return float(text_or_val)

    s = clean_str(text_or_val)
    if not s:
        return None

    # Specific regex pattern for voltage in description or string cell
    match = re.search(r"voltage\s*[:=]?\s*([0-9]+(?:\.[0-9]+)?)", s, re.IGNORECASE)
    if match:
        try:
            return float(match.group(1))
        except ValueError:
            pass

    # Match 24V or 24 V
    match_v = re.search(r"([0-9]+(?:\.[0-9]+)?)\s*v(?:olts?)?", s, re.IGNORECASE)
    if match_v:
        try:
            return float(match_v.group(1))
        except ValueError:
            pass

    # Generic number fallback
    return extract_number(s)


def extract_current(text_or_val: Any) -> Optional[float]:
    """
    Extracts current float value from raw text or numeric cell.
    Supports patterns like:
    - 'SET CURRENT 1.5A' -> 1.5
    - '1.5A' -> 1.5
    - 1.5 -> 1.5
    """
    if text_or_val is None:
        return None
    if isinstance(text_or_val, (int, float)):
        if isinstance(text_or_val, float) and math.isnan(text_or_val):
            return None
        return float(text_or_val)

    s = clean_str(text_or_val)
    if not s:
        return None

    # Specific regex pattern for current
    match = re.search(r"current\s*[:=]?\s*([0-9]+(?:\.[0-9]+)?)", s, re.IGNORECASE)
    if match:
        try:
            return float(match.group(1))
        except ValueError:
            pass

    # Match 1.5A or 1.5 A
    match_a = re.search(r"([0-9]+(?:\.[0-9]+)?)\s*a(?:mps?)?", s, re.IGNORECASE)
    if match_a:
        try:
            return float(match_a.group(1))
        except ValueError:
            pass

    return extract_number(s)


def get_voltage_min_max(voltage_val: Optional[float]) -> Tuple[Optional[float], Optional[float]]:
    """
    Calculates Min and Max allowed voltage values given a nominal voltage.
    First checks config.VOLTAGE_LOOKUP.
    If exact nominal voltage is found in lookup table (e.g., 24 -> (25.0, 27.0)), returns that.
    If not found but voltage is provided, computes a standard tolerance range (nominal - 1V to nominal + 1V or ±5%).
    If voltage_val is None, returns (None, None).
    """
    if voltage_val is None:
        return None, None

    v_rounded = round(float(voltage_val), 1)

    # Check exact matching in VOLTAGE_LOOKUP
    if v_rounded in config.VOLTAGE_LOOKUP:
        return config.VOLTAGE_LOOKUP[v_rounded]

    # Check integer key matching (e.g. 24.0 vs 24)
    v_int = float(round(v_rounded))
    if v_int in config.VOLTAGE_LOOKUP:
        return config.VOLTAGE_LOOKUP[v_int]

    # Fallback tolerance calculation if nominal voltage is custom
    min_v = round(v_rounded * 0.95 + 1.0, 1)
    max_v = round(v_rounded * 1.05 + 1.0, 1)
    logger.warning(
        f"Nominal voltage {voltage_val}V not found in VOLTAGE_LOOKUP table. "
        f"Using fallback min/max calculation: Min={min_v}, Max={max_v}"
    )
    return min_v, max_v


def format_number_display(val: Optional[float], unit: str = "") -> str:
    """
    Formats float or int for table display.
    Examples:
        (24.0, 'V') -> '24V'
        (1.5, 'A') -> '1.5A'
        (None, 'V') -> ''
    """
    if val is None:
        return ""
    if val == int(val):
        formatted = str(int(val))
    else:
        formatted = f"{val:.2f}".rstrip('0').rstrip('.')
    return f"{formatted}{unit}"


def format_output_filename(original_filename: str) -> str:
    """
    Generates formatted output filename based on original uploaded filename.
    Example: 'OTIS_Test_Seq.xlsx' -> 'Formatted_OTIS_Test_Seq.xlsx'
    """
    basename = os.path.basename(original_filename)
    name, ext = os.path.splitext(basename)
    if not ext:
        ext = ".xlsx"
    clean_name = re.sub(r"^Formatted_", "", name, flags=re.IGNORECASE)
    return f"{config.OUTPUT_FILENAME_PREFIX}{clean_name}{ext}"
