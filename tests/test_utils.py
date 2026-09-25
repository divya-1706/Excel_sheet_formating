"""
Unit tests for utils.py module.
"""

import unittest
import os
import sys

# Ensure parent directory is in sys.path
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from utils import (
    clean_str, extract_number, extract_voltage, extract_current,
    get_voltage_min_max, format_number_display, format_output_filename
)


class TestUtils(unittest.TestCase):

    def test_clean_str(self):
        self.assertEqual(clean_str("  hello  "), "hello")
        self.assertEqual(clean_str(None), "")
        self.assertEqual(clean_str("None"), "")
        self.assertEqual(clean_str("NaN"), "")
        self.assertEqual(clean_str(24), "24")

    def test_extract_number(self):
        self.assertEqual(extract_number("24V"), 24.0)
        self.assertEqual(extract_number("SET CURRENT 1.5A"), 1.5)
        self.assertIsNone(extract_number("no numbers here"))

    def test_extract_voltage(self):
        self.assertEqual(extract_voltage("SET Voltage 24V"), 24.0)
        self.assertEqual(extract_voltage("SET Voltage 18"), 18.0)
        self.assertEqual(extract_voltage("28V"), 28.0)
        self.assertEqual(extract_voltage(48.0), 48.0)
        self.assertIsNone(extract_voltage(""))

    def test_extract_current(self):
        self.assertEqual(extract_current("SET CURRENT 1.5A"), 1.5)
        self.assertEqual(extract_current("2.0A"), 2.0)
        self.assertEqual(extract_current(3.0), 3.0)

    def test_get_voltage_min_max(self):
        # Exact lookup test cases
        self.assertEqual(get_voltage_min_max(18.0), (19.0, 21.0))
        self.assertEqual(get_voltage_min_max(24.0), (25.0, 27.0))
        self.assertEqual(get_voltage_min_max(28.0), (29.0, 31.0))
        self.assertEqual(get_voltage_min_max(30.0), (31.0, 33.0))
        self.assertEqual(get_voltage_min_max(48.0), (49.0, 51.0))
        self.assertEqual(get_voltage_min_max(54.0), (55.0, 57.0))
        self.assertEqual(get_voltage_min_max(60.0), (61.0, 63.0))

    def test_format_output_filename(self):
        self.assertTrue(format_output_filename("test_seq.xlsx").endswith("test_seq.xlsx"))
        self.assertTrue(format_output_filename("Converted_test.xlsx").endswith("test.xlsx"))


if __name__ == "__main__":
    unittest.main()
