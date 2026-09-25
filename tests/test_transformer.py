"""
Unit tests for transformer.py module.
"""

import unittest
import os
import sys
import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from prompt_engine import parse_user_prompt
from transformer import transform_dataset


class TestTransformer(unittest.TestCase):

    def test_student_result_transformation(self):
        df_in = pd.DataFrame([
            {"USN": "101", "Student Name": "Alice", "Total Marks": 450},
            {"USN": "102", "Student Name": "Bob", "Total Marks": 350},
        ])

        col_map = {"Student Name": "Student Name", "USN": "USN", "Total Marks": "Total Marks"}
        instructions = parse_user_prompt("Convert to Student Result sheet")

        res = transform_dataset(df_in, col_map, instructions)
        self.assertEqual(len(res.df_grid), 2)
        self.assertIn("Percentage", res.df_grid.columns)
        self.assertIn("Grade", res.df_grid.columns)
        self.assertEqual(res.df_grid.at[0, "Percentage"], 90.0)
        self.assertEqual(res.df_grid.at[0, "Grade"], "A+")


if __name__ == "__main__":
    unittest.main()
