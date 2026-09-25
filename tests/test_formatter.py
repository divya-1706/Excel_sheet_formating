"""
Unit tests for formatter.py module.
"""

import unittest
import os
import sys
import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import config
from prompt_engine import parse_user_prompt
from transformer import transform_dataset
from formatter import export_transformed_data


class TestFormatter(unittest.TestCase):

    def test_export_transformed_data(self):
        df_in = pd.DataFrame([
            {"Invoice No": "INV-01", "Item Description": "Widget A", "Quantity": 5, "Unit Price": 100.0}
        ])
        col_map = {"Invoice No": "Invoice No", "Item Description": "Item Description", "Quantity": "Quantity", "Unit Price": "Unit Price"}
        instructions = parse_user_prompt("Convert to Invoice format")
        res = transform_dataset(df_in, col_map, instructions)

        out_file = os.path.join(config.OUTPUT_DIR, "test_formatter_out.xlsx")
        if os.path.exists(out_file):
            os.remove(out_file)

        saved = export_transformed_data(res, out_file, "XLSX")
        self.assertTrue(os.path.exists(saved))
        if os.path.exists(out_file):
            os.remove(out_file)


if __name__ == "__main__":
    unittest.main()
