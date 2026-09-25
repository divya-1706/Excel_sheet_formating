"""
Unit tests for parser.py module.
"""

import unittest
import os
import sys

# Ensure parent directory is in sys.path
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import config
from parser import parse_test_sequence_workbook, _is_new_report_row, _extract_report_number


class TestParser(unittest.TestCase):

    def test_is_new_report_row(self):
        self.assertTrue(_is_new_report_row("1.1"))
        self.assertTrue(_is_new_report_row("1.2"))
        self.assertTrue(_is_new_report_row("2.1"))
        self.assertTrue(_is_new_report_row("10.3"))
        self.assertTrue(_is_new_report_row("12.15"))
        self.assertFalse(_is_new_report_row(""))
        self.assertFalse(_is_new_report_row("Sl.No"))
        self.assertFalse(_is_new_report_row("Test Description"))

    def test_extract_report_number(self):
        self.assertEqual(_extract_report_number("Sl.No 1.1"), "1.1")
        self.assertEqual(_extract_report_number("1.2"), "1.2")
        self.assertEqual(_extract_report_number("10.3"), "10.3")

    def test_parse_sample_workbook(self):
        sample_path = os.path.join(config.SAMPLE_DIR, "sample_test_sequence.xlsx")
        self.assertTrue(os.path.exists(sample_path), "Sample workbook missing")

        blocks, warnings, summary = parse_test_sequence_workbook(sample_path)
        self.assertGreaterEqual(len(blocks), 5)
        self.assertEqual(blocks[0].report_no, "1.1")
        self.assertEqual(blocks[0].voltage, 24.0)
        self.assertEqual(blocks[0].current, 1.5)
        self.assertEqual(blocks[0].min_voltage, 25.0)
        self.assertEqual(blocks[0].max_voltage, 27.0)


if __name__ == "__main__":
    unittest.main()
