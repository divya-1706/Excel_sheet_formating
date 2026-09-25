"""
Config module for Universal AI Excel Format Converter.
Stores constants, directory paths, column synonym dictionaries,
voltage tables, default styles, and transformation presets.
"""

import os
from typing import Dict, List, Tuple, Any

# Base paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
LOG_DIR = os.path.join(BASE_DIR, "logs")
SAMPLE_DIR = os.path.join(BASE_DIR, "sample")
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
ASSETS_DIR = os.path.join(BASE_DIR, "assets")

# Default Constants
DEFAULT_SHEET_NAME = "Sheet1"
OUTPUT_FILENAME_PREFIX = "Converted_"

# Comprehensive Synonym Dictionaries for AI Column Mapping
COLUMN_SYNONYMS: Dict[str, List[str]] = {
    "sl_no": [
        "sl.no", "sl no", "sl_no", "s.no", "sno", "sr.no", "sl. no.",
        "sl.no.", "sl. no", "si.no", "si no", "si. no.", "sl", "no", "no.",
        "seq", "seq no", "seq.no", "sequence", "step", "item no", "test no", "serial", "serial no", "serial number"
    ],
    "test_name": [
        "test name", "testname", "test_name", "sequence name", "test name/description",
        "item name", "product name", "description", "item description", "parameter", "feature", "task name"
    ],
    "name": [
        "name", "full name", "student name", "employee name", "customer name", "client name",
        "contact name", "person name", "emp_name", "student_name"
    ],
    "id": [
        "id", "usn", "emp_id", "employee id", "student id", "customer id", "roll no", "registration no",
        "code", "item code", "sku", "part no", "part number"
    ],
    "point": [
        "point", "measurement point", "test point", "node", "channel", "tp", "location", "pin"
    ],
    "voltage": [
        "voltage", "volts", "volt", "voltage (v)", "voltage(v)", "v", "battery voltage", "voltage_val"
    ],
    "current": [
        "current", "amps", "amp", "current (a)", "current(a)", "curr", "load current"
    ],
    "device": [
        "device", "equipment", "instrument", "device name", "tool", "source", "psu", "meter"
    ],
    "quantity": [
        "qty", "quantity", "count", "units", "amount", "num_units", "items_count"
    ],
    "price": [
        "price", "unit price", "cost", "unit cost", "rate", "fee", "amount", "charge"
    ],
    "total": [
        "total", "total amount", "total cost", "subtotal", "total price", "grand total", "sum"
    ],
    "status": [
        "status", "result", "state", "pass/fail", "condition", "remarks", "grade", "verdict"
    ],
    "date": [
        "date", "created date", "entry date", "timestamp", "datetime", "dob", "invoice date", "due date"
    ],
    "marks": [
        "marks", "score", "total marks", "obtained marks", "points", "grade points"
    ],
    "percentage": [
        "percentage", "percent", "%", "pct", "score percentage", "attendance percentage"
    ]
}

# Voltage Min/Max Lookup Table for OTIS Layout
VOLTAGE_LOOKUP: Dict[float, Tuple[float, float]] = {
    18.0: (19.0, 21.0),
    24.0: (25.0, 27.0),
    28.0: (29.0, 31.0),
    30.0: (31.0, 33.0),
    48.0: (49.0, 51.0),
    54.0: (55.0, 57.0),
    60.0: (61.0, 63.0),
}

# Supported Output Formats
SUPPORTED_OUTPUT_FORMATS = ["XLSX", "CSV", "ODS"]

# Default Styling Configurations
FONT_FAMILY = "Calibri"
HEADER_FONT_SIZE = 11
BODY_FONT_SIZE = 10

# Blue and White Professional Color Palette
PRIMARY_BLUE = "1E40AF"      # Deep Blue
SECONDARY_BLUE = "3B82F6"    # Sky Blue
LIGHT_BLUE_FILL = "EFF6FF"   # Light Tint
HEADER_FILL_COLOR = "D9D9D9"  # Standard Light Grey
BORDER_COLOR_HEX = "000000"   # Black

# Row Heights
HEADER_ROW_HEIGHT = 26.0
DATA_ROW_HEIGHT = 20.0

# Standard Column Widths
DEFAULT_COLUMN_WIDTH = 18.0
OTIS_COLUMN_WIDTHS = {
    "A": 10.0,  # Sl No
    "B": 35.0,  # Test Name
    "C": 18.0,  # Device
    "D": 28.0,  # Value
    "E": 10.0,  # Min
    "F": 10.0,  # Max
    "G": 12.0   # Actual
}

# Transformation Layout Presets
PRESET_TEMPLATES = {
    "Auto Grid": "Standard column-mapped grid format with calculated fields",
    "OTIS Inspection Report": "5-row per test sequence block with merged cells, PSU/Waveshare/DMM rows",
    "Invoice / Billing": "Header metadata + itemized billing table + subtotal/tax/grand total rows",
    "Student Result Sheet": "Student list with calculated Total Marks, Percentage, and Grade rules",
    "Inventory BOM": "Bill of Materials with categorized grouping, unit costs, and total valuation"
}
