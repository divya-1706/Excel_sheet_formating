"""
Utility functions for Universal AI Excel Format Converter.
Handles string parsing, number extraction, voltage lookup calculations,
file name formatting, and folder initialization.
"""

import os
import re
import math
from typing import Optional, Tuple, Any, List
import config
from logger import logger


def ensure_directories() -> None:
    """Ensures all required project directories exist."""
    directories = [
        config.INPUT_DIR,
        config.OUTPUT_DIR,
        config.LOG_DIR,
        config.SAMPLE_DIR,
        config.TEMPLATES_DIR,
        config.ASSETS_DIR,
    ]
    for d in directories:
        os.makedirs(d, exist_ok=True)


def clean_str(val: Any) -> str:
    """Cleans input into a stripped string. Returns empty string if None/NaN."""
    if val is None:
        return ""
    if isinstance(val, float) and math.isnan(val):
        return ""
    s = str(val).strip()
    if s.lower() in ("none", "nan", "null"):
        return ""
    return s


def normalize_column_name(col_name: Any) -> str:
    """Normalizes column name: lowercase, strip punctuation and extra spaces."""
    s = clean_str(col_name).lower()
    s = re.sub(r"[_\-.\/]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def extract_number(text: Any) -> Optional[float]:
    """Extracts the first float or integer number from text."""
    if text is None:
        return None
    if isinstance(text, (int, float)):
        if isinstance(text, float) and math.isnan(text):
            return None
        return float(text)
    s = clean_str(text)
    match = re.search(r"[-+]?\d*\.?\d+", s)
    if match:
        try:
            return float(match.group(0))
        except ValueError:
            return None
    return None


def extract_voltage(text_or_val: Any) -> Optional[float]:
    """Extracts numeric voltage value (e.g. 'SET Voltage 24V' -> 24.0)."""
    if text_or_val is None:
        return None
    if isinstance(text_or_val, (int, float)):
        if isinstance(text_or_val, float) and math.isnan(text_or_val):
            return None
        return float(text_or_val)

    s = clean_str(text_or_val)
    if not s:
        return None

    match = re.search(r"voltage\s*[:=]?\s*([0-9]+(?:\.[0-9]+)?)", s, re.IGNORECASE)
    if match:
        try:
            return float(match.group(1))
        except ValueError:
            pass

    match_v = re.search(r"([0-9]+(?:\.[0-9]+)?)\s*v(?:olts?)?", s, re.IGNORECASE)
    if match_v:
        try:
            return float(match_v.group(1))
        except ValueError:
            pass

    return extract_number(s)


def extract_current(text_or_val: Any) -> Optional[float]:
    """Extracts numeric current value (e.g. 'SET CURRENT 1.5A' -> 1.5)."""
    if text_or_val is None:
        return None
    if isinstance(text_or_val, (int, float)):
        if isinstance(text_or_val, float) and math.isnan(text_or_val):
            return None
        return float(text_or_val)

    s = clean_str(text_or_val)
    if not s:
        return None

    match = re.search(r"current\s*[:=]?\s*([0-9]+(?:\.[0-9]+)?)", s, re.IGNORECASE)
    if match:
        try:
            return float(match.group(1))
        except ValueError:
            pass

    match_a = re.search(r"([0-9]+(?:\.[0-9]+)?)\s*a(?:mps?)?", s, re.IGNORECASE)
    if match_a:
        try:
            return float(match_a.group(1))
        except ValueError:
            pass

    return extract_number(s)


def get_voltage_min_max(voltage_val: Optional[float]) -> Tuple[Optional[float], Optional[float]]:
    """Calculates Min and Max allowed voltage from VOLTAGE_LOOKUP or tolerance fallback."""
    if voltage_val is None:
        return None, None

    v_rounded = round(float(voltage_val), 1)
    if v_rounded in config.VOLTAGE_LOOKUP:
        return config.VOLTAGE_LOOKUP[v_rounded]

    v_int = float(round(v_rounded))
    if v_int in config.VOLTAGE_LOOKUP:
        return config.VOLTAGE_LOOKUP[v_int]

    min_v = round(v_rounded * 0.95 + 1.0, 1)
    max_v = round(v_rounded * 1.05 + 1.0, 1)
    return min_v, max_v


def format_number_display(val: Optional[float], unit: str = "") -> str:
    """Formats float/int for UI display."""
    if val is None:
        return ""
    if val == int(val):
        formatted = str(int(val))
    else:
        formatted = f"{val:.2f}".rstrip('0').rstrip('.')
    return f"{formatted}{unit}"


def format_output_filename(original_filename: str, ext: str = ".xlsx") -> str:
    """Generates formatted output filename based on original uploaded filename."""
    basename = os.path.basename(original_filename)
    name, _ = os.path.splitext(basename)
    clean_name = re.sub(r"^Converted_", "", name, flags=re.IGNORECASE)
    if not ext.startswith("."):
        ext = f".{ext}"
    return f"{config.OUTPUT_FILENAME_PREFIX}{clean_name}{ext}"
