"""
Config module for OTIS Excel Report Formatter application.
Stores application constants, formatting defaults, and voltage lookup tables.
"""

import os
from typing import Dict, Tuple

# Base paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
LOG_DIR = os.path.join(BASE_DIR, "logs")
SAMPLE_DIR = os.path.join(BASE_DIR, "sample")
ASSETS_DIR = os.path.join(BASE_DIR, "assets")

# File & Sheet Constants
DEFAULT_SHEET_NAME = "Test Sequence"
OUTPUT_FILENAME_PREFIX = "Formatted_"

# Expected Input Column Keywords (case-insensitive search)
COLUMN_KEYWORDS = {
    "sl_no": [
        "sl.no", "sl no", "sl_no", "s.no", "sno", "sr.no", "sl. no.",
        "sl.no.", "sl. no", "si.no", "si no", "si. no.", "sl", "no", "no.",
        "seq", "seq no", "seq.no", "sequence", "step", "item", "test no", "test.no", "s. no."
    ],
    "test_name": ["test name", "testname", "test_name", "sequence name", "test name/description"],
    "point": ["point", "measurement point", "test point", "node", "channel", "tp"],
    "description": ["description", "test description", "instruction", "action", "procedure", "specification"],
    "voltage": ["voltage", "volts", "volt", "voltage (v)", "voltage(v)"],
    "current": ["current", "amps", "amp", "current (a)", "current(a)", "curr"],
    "device": ["device", "equipment", "instrument", "device name", "tool", "source"],
    "remarks": ["remarks", "notes", "remark", "result", "comments"],
    "actual_value": ["actual value", "actual", "measured value", "reading", "value", "observed"]
}

# Voltage Min/Max Lookup Table (Nominal Voltage -> (Min, Max))
# Default mapping per OTIS specification:
# 18 -> (19.0, 21.0), 24 -> (25.0, 27.0), 28 -> (29.0, 31.0),
# 30 -> (31.0, 33.0), 48 -> (49.0, 51.0), 54 -> (55.0, 57.0), 60 -> (61.0, 63.0)
VOLTAGE_LOOKUP: Dict[float, Tuple[float, float]] = {
    18.0: (19.0, 21.0),
    24.0: (25.0, 27.0),
    28.0: (29.0, 31.0),
    30.0: (31.0, 33.0),
    48.0: (49.0, 51.0),
    54.0: (55.0, 57.0),
    60.0: (61.0, 63.0),
}

# Output Report Table Headers
OUTPUT_COLUMNS = [
    "Sl No",
    "Test Name",
    "Device",
    "Value",
    "Min",
    "Max",
    "Actual"
]

# Formatting & Styling Configuration
FONT_FAMILY = "Calibri"
HEADER_FONT_SIZE = 11
BODY_FONT_SIZE = 10

HEADER_FILL_COLOR = "D9D9D9"  # Light grey
BORDER_COLOR = "000000"       # Black

# Standard Column Widths (matching OTIS template specification)
COLUMN_WIDTHS = {
    "A": 10.0,  # Sl No
    "B": 35.0,  # Test Name
    "C": 18.0,  # Device
    "D": 28.0,  # Value
    "E": 10.0,  # Min
    "F": 10.0,  # Max
    "G": 12.0   # Actual
}

# Row Heights
HEADER_ROW_HEIGHT = 26.0
DATA_ROW_HEIGHT = 20.0

# Device Names Default Fallbacks
DEFAULT_PSU_DEVICE = "PSU.01"
DEFAULT_WAVESHARE_DEVICE = "Waveshare"
DEFAULT_DMM_DEVICE = "DMM"
