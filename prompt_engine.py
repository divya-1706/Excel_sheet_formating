"""
Prompt Engine module for Universal AI Excel Format Converter.
Parses natural language user instructions and user-defined custom column lists
into structured transformation rules, calculated columns, merge instructions, and conditional rules.
"""

from dataclasses import dataclass, field
import re
from typing import List, Dict, Any, Optional
import config
from logger import logger


@dataclass
class CalculationRule:
    target_column: str
    expression: str
    formula_type: str = "custom"


@dataclass
class ConditionalRule:
    column_name: str
    condition: str
    value: str
    action_type: str
    action_target: str


@dataclass
class PromptInstructions:
    preset_name: str = "Custom Columns"
    target_columns: List[str] = field(default_factory=list)
    merge_instructions: List[str] = field(default_factory=list)
    calculations: List[CalculationRule] = field(default_factory=list)
    conditional_rules: List[ConditionalRule] = field(default_factory=list)
    raw_prompt: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "preset_name": self.preset_name,
            "target_columns": self.target_columns,
            "merge_instructions": self.merge_instructions,
            "calculations": [{"target": c.target_column, "expr": c.expression} for c in self.calculations],
            "conditional_rules": [{"col": r.column_name, "cond": r.condition, "val": r.value, "target": r.action_target} for r in self.conditional_rules],
            "raw_prompt": self.raw_prompt
        }


def parse_custom_columns(col_text: str) -> List[str]:
    """
    Parses user-entered custom column text into a list of clean column names.
    Supports comma-separated, semicolon-separated, or newline-separated strings.
    Example: 'Sl No, Test Name, Device, Value, Min Value, Max Value, Actual'
    -> ['Sl No', 'Test Name', 'Device', 'Value', 'Min Value', 'Max Value', 'Actual']
    """
    if not col_text:
        return []
    parts = re.split(r"[,;\n]", col_text.strip())
    cols = [p.strip() for p in parts if p.strip()]
    return cols


def parse_user_prompt(prompt_text: str, custom_columns_text: str = "") -> PromptInstructions:
    """
    Parses user natural language instructions and custom column inputs.
    Extracts target column names, calculation formulas, and merge rules.
    """
    instructions = PromptInstructions(raw_prompt=prompt_text or "")

    clean_p = (prompt_text or "").strip()
    p_lower = clean_p.lower()

    if "otis" in p_lower:
        instructions.preset_name = "OTIS Inspection Report"
    elif "student" in p_lower or "result" in p_lower or "mark" in p_lower:
        instructions.preset_name = "Student Result Sheet"
    elif "invoice" in p_lower or "billing" in p_lower or "sale" in p_lower:
        instructions.preset_name = "Invoice / Billing"

    # 1. Custom User-Defined Columns
    custom_cols = parse_custom_columns(custom_columns_text)
    if custom_cols:
        instructions.target_columns = custom_cols

    if not clean_p:
        return instructions

    p_lower = clean_p.lower()

    # 2. Extract Requested Columns from prompt if not provided in custom_columns_text
    if not instructions.target_columns:
        col_match = re.search(r"columns?\s*[:=]\s*([^\.\n]+)", clean_p, re.IGNORECASE)
        if col_match:
            raw_cols = col_match.group(1)
            parsed_cols = parse_custom_columns(raw_cols)
            if parsed_cols:
                instructions.target_columns = parsed_cols

    # Default fallback columns if neither provided
    if not instructions.target_columns:
        if "otis" in p_lower:
            instructions.target_columns = ["Sl No", "Test Name", "Device", "Value", "Min", "Max", "Actual"]
        elif "student" in p_lower or "result" in p_lower or "mark" in p_lower:
            instructions.target_columns = ["Student Name", "USN", "Department", "Semester", "Total Marks", "Percentage", "Grade"]
        elif "invoice" in p_lower or "billing" in p_lower or "sale" in p_lower:
            instructions.target_columns = ["Invoice No", "Item Description", "Quantity", "Unit Price", "Total", "GST (18%)", "Grand Total"]

    # 3. Extract Calculations (e.g. Percentage = Total Marks / 500 * 100)
    calc_matches = re.findall(r"(\b[A-Za-z0-9_\s]+\b)\s*=\s*([^\n,;\.]+)", clean_p)
    for target, expr in calc_matches:
        t_clean = re.sub(r"^(?:calculate|compute|set)\s+", "", target.strip(), flags=re.IGNORECASE).strip()
        e_clean = expr.strip()
        if t_clean.lower() not in ("columns", "columns:", "format", "preset"):
            instructions.calculations.append(CalculationRule(target_column=t_clean, expression=e_clean))

    # 4. Extract Merge Rules
    if "merge" in p_lower:
        merge_matches = re.findall(r"merge\s+([^\n,\.]+)", clean_p, re.IGNORECASE)
        for m in merge_matches:
            instructions.merge_instructions.append(f"Merge {m.strip()}")

    # 5. Extract Conditional Rules
    cond_matches = re.findall(r"if\s+([A-Za-z0-9_\s]+)\s+(is|>|<|==|above|below)\s+([A-Za-z0-9_%]+)\s+(color|add|set)\s+([A-Za-z0-9_\s]+)", clean_p, re.IGNORECASE)
    for col, cond, val, act, tgt in cond_matches:
        instructions.conditional_rules.append(ConditionalRule(
            column_name=col.strip(),
            condition=cond.strip(),
            value=val.strip(),
            action_type=act.strip().lower(),
            action_target=tgt.strip()
        ))

    logger.info(f"Parsed prompt: Target Columns={instructions.target_columns}")
    return instructions
