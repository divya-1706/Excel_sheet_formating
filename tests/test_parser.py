"""
Unit tests for parser.py module.
"""

import unittest
import os
import sys

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import config
from parser import parse_excel_sheet, get_workbook_sheet_names


class TestParser(unittest.TestCase):

    def test_parse_sample_student_marks(self):
        sample_path = os.path.join(config.SAMPLE_DIR, "Student_Marks_Sheet.xlsx")
        self.assertTrue(os.path.exists(sample_path), "Sample workbook missing")

        sheets = get_workbook_sheet_names(sample_path)
        self.assertGreaterEqual(len(sheets), 1)

        df, _, summary = parse_excel_sheet(sample_path, sheet_name=sheets[0])
        self.assertEqual(len(df), 4)
        self.assertIn("Student Name", df.columns)
        self.assertIn("USN", df.columns)
        self.assertIn("Total Marks", df.columns)


if __name__ == "__main__":
    unittest.main()
