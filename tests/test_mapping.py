"""
Unit tests for mapping_engine.py module.
"""

import unittest
import os
import sys

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from mapping_engine import suggest_column_mappings, compute_string_similarity


class TestMappingEngine(unittest.TestCase):

    def test_compute_similarity(self):
        self.assertEqual(compute_string_similarity("Student Name", "Student Name"), 1.0)
        self.assertGreaterEqual(compute_string_similarity("Battery Voltage", "Voltage"), 0.8)

    def test_suggest_column_mappings(self):
        input_cols = ["USN", "Student Name", "Total Marks", "QTY", "Battery Voltage"]
        target_cols = ["Voltage", "Student Name", "Quantity"]

        mappings = suggest_column_mappings(input_cols, target_cols)
        self.assertEqual(mappings["Voltage"][0], "Battery Voltage")
        self.assertEqual(mappings["Student Name"][0], "Student Name")
        self.assertEqual(mappings["Quantity"][0], "QTY")


if __name__ == "__main__":
    unittest.main()
