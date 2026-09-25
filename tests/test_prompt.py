"""
Unit tests for prompt_engine.py module.
"""

import unittest
import os
import sys

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from prompt_engine import parse_user_prompt


class TestPromptEngine(unittest.TestCase):

    def test_parse_prompt_preset_detection(self):
        inst_otis = parse_user_prompt("Convert to OTIS report layout")
        self.assertEqual(inst_otis.preset_name, "OTIS Inspection Report")

        inst_student = parse_user_prompt("Convert to Student Result sheet with columns: Student Name, USN, Total Marks, Percentage")
        self.assertEqual(inst_student.preset_name, "Student Result Sheet")
        self.assertIn("Percentage", inst_student.target_columns)

    def test_parse_prompt_calculations(self):
        inst = parse_user_prompt("Calculate Percentage = Total Marks / 500 * 100")
        self.assertEqual(len(inst.calculations), 1)
        self.assertEqual(inst.calculations[0].target_column, "Percentage")


if __name__ == "__main__":
    unittest.main()
