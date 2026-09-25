"""
Unit tests for validators.py module.
"""

import unittest
import os
import sys

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import config
from validators import validate_workbook_file, validate_column_mapping


class TestValidators(unittest.TestCase):

    def test_validate_workbook_file(self):
        sample_file = os.path.join(config.SAMPLE_DIR, "Student_Marks_Sheet.xlsx")
        is_valid, errs = validate_workbook_file(sample_file)
        self.assertTrue(is_valid)
        self.assertEqual(len(errs), 0)

    def test_validate_column_mapping(self):
        input_cols = ["Name", "Score"]
        target_cols = ["Name", "Score", "Grade"]
        col_map = {"Name": "Name", "Score": "Score"}
        warnings, errors = validate_column_mapping(input_cols, target_cols, col_map)
        self.assertGreaterEqual(len(warnings), 1)


if __name__ == "__main__":
    unittest.main()
