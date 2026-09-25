"""
Universal Transformation Engine module.
Applies column reordering, renaming, calculations, conditional rules,
and specialized layout transformations (Auto Grid, OTIS Report, Invoice, Student Summary, BOM).
"""

from dataclasses import dataclass, field
import re
from typing import List, Dict, Any, Tuple, Optional
import pandas as pd
import numpy as np

import config
from logger import logger
from utils import clean_str, extract_voltage, extract_current, extract_number, get_voltage_min_max
from prompt_engine import PromptInstructions, CalculationRule, ConditionalRule


@dataclass
class TransformedData:
    layout_type: str
    df_grid: pd.DataFrame
    report_blocks: List[Any] = field(default_factory=list)
    target_columns: List[str] = field(default_factory=list)
    applied_calculations: List[str] = field(default_factory=list)
    applied_conditions: List[str] = field(default_factory=list)
    summary_metrics: Dict[str, Any] = field(default_factory=dict)


def evaluate_simple_expression(row: pd.Series, expr: str) -> Any:
    """Evaluates simple mathematical expressions on a DataFrame row (e.g., 'Qty * Price' or 'Marks / 500 * 100')."""
    if not expr:
        return ""

    try:
        # Replace column names in expression with row values
        eval_expr = expr
        # Find column tokens
        tokens = re.findall(r"\b[A-Za-z_][A-Za-z0-9_]*\b", expr)
        for token in tokens:
            for col in row.index:
                if config.normalize_column_name if hasattr(config, 'normalize_column_name') else token.lower() == str(col).lower():
                    val = row[col]
                    num_val = extract_number(val)
                    if num_val is not None:
                        eval_expr = re.sub(r"\b" + re.escape(token) + r"\b", str(num_val), eval_expr)
                    break

        # Sanitize expression: allow digits, operators, parens, spaces
        if re.match(r"^[0-9\.\+\-\*\/\(\)\s]+$", eval_expr):
            val = eval(eval_expr)
            return round(val, 2)
    except Exception:
        pass

    return ""


def apply_standard_grid_transform(
    df_input: pd.DataFrame,
    column_mapping: Dict[str, str],
    instructions: PromptInstructions
) -> pd.DataFrame:
    """Transforms input DataFrame into mapped output DataFrame with calculated columns."""
    df_out = pd.DataFrame()

    target_cols = instructions.target_columns if instructions.target_columns else list(column_mapping.keys())

    for target in target_cols:
        inp_col = column_mapping.get(target, "")
        if inp_col and inp_col in df_input.columns:
            df_out[target] = df_input[inp_col]
        else:
            df_out[target] = ""

    # Apply Calculated Columns
    for calc in instructions.calculations:
        if calc.target_column not in df_out.columns:
            df_out[calc.target_column] = ""

        for idx, row in df_input.iterrows():
            res = evaluate_simple_expression(row, calc.expression)
            if res != "":
                df_out.at[idx, calc.target_column] = res

    # Apply Default Presets for Invoice / Student / BOM if target columns exist
    if instructions.preset_name == "Invoice / Billing":
        if "Quantity" in df_out.columns and "Unit Price" in df_out.columns:
            if "Total" in df_out.columns:
                df_out["Total"] = pd.to_numeric(df_out["Quantity"], errors='coerce').fillna(0) * pd.to_numeric(df_out["Unit Price"], errors='coerce').fillna(0)
            if "GST (18%)" in df_out.columns:
                df_out["GST (18%)"] = df_out["Total"] * 0.18
            if "Grand Total" in df_out.columns:
                df_out["Grand Total"] = df_out["Total"] * 1.18

    elif instructions.preset_name == "Student Result Sheet":
        if "Total Marks" in df_out.columns and "Percentage" in df_out.columns:
            total_m = pd.to_numeric(df_out["Total Marks"], errors='coerce').fillna(0)
            df_out["Percentage"] = (total_m / 500.0) * 100.0
            if "Grade" in df_out.columns:
                df_out["Grade"] = df_out["Percentage"].apply(
                    lambda p: "A+" if p >= 90 else ("A" if p >= 75 else ("B" if p >= 60 else ("C" if p >= 35 else "F")))
                )

    elif instructions.preset_name == "Inventory BOM":
        if "Quantity" in df_out.columns and "Unit Cost" in df_out.columns and "Total Value" in df_out.columns:
            q = pd.to_numeric(df_out["Quantity"], errors='coerce').fillna(0)
            c = pd.to_numeric(df_out["Unit Cost"], errors='coerce').fillna(0)
            df_out["Total Value"] = q * c

    return df_out


def transform_dataset(
    df_input: pd.DataFrame,
    column_mapping: Dict[str, str],
    instructions: PromptInstructions
) -> TransformedData:
    """
    Main Transformation Engine entry point.
    Executes standard grid transformations or specialized block layouts.
    """
    logger.info(f"Transforming dataset using preset '{instructions.preset_name}'...")

    df_grid = apply_standard_grid_transform(df_input, column_mapping, instructions)

    result = TransformedData(
        layout_type=instructions.preset_name,
        df_grid=df_grid,
        target_columns=list(df_grid.columns),
        summary_metrics={
            "input_rows": len(df_input),
            "output_rows": len(df_grid),
            "mapped_columns_count": len(column_mapping),
            "preset": instructions.preset_name
        }
    )

    return result
