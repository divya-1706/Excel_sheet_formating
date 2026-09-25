"""
Prompt Engine module for Universal AI Excel Format Converter.
Parses natural language user instructions into structured transformation rules,
layout presets, calculated columns, merge instructions, and conditional rules.
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
    formula_type: str = "custom"  # total, percentage, gst, custom


@dataclass
class ConditionalRule:
    column_name: str
    condition: str
    value: str
    action_type: str  # color, set_val
    action_target: str


@dataclass
class PromptInstructions:
    preset_name: str = "Auto Grid"
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


def parse_user_prompt(prompt_text: str, default_preset: str = "Auto Grid") -> PromptInstructions:
    """
    Parses natural language user prompt into structured PromptInstructions object.
    Extracts layout preset, target column names, calculation formulas, and merge rules.
    """
    if not prompt_text:
        return PromptInstructions(preset_name=default_preset, raw_prompt="")

    clean_p = prompt_text.strip()
    instructions = PromptInstructions(preset_name=default_preset, raw_prompt=clean_p)

    # 1. Preset Detection
    p_lower = clean_p.lower()
    if "otis" in p_lower or "test sequence" in p_lower:
        instructions.preset_name = "OTIS Inspection Report"
        instructions.target_columns = ["Sl No", "Test Name", "Device", "Value", "Min", "Max", "Actual"]
        instructions.merge_instructions = ["Merge Sl No vertically", "Merge Test Name vertically"]
    elif "invoice" in p_lower or "billing" in p_lower or "sale" in p_lower:
        instructions.preset_name = "Invoice / Billing"
        instructions.target_columns = ["Invoice No", "Item Description", "Quantity", "Unit Price", "Total", "GST (18%)", "Grand Total"]
    elif "student" in p_lower or "result" in p_lower or "mark" in p_lower or "grade" in p_lower:
        instructions.preset_name = "Student Result Sheet"
        instructions.target_columns = ["Student Name", "USN", "Department", "Semester", "Total Marks", "Percentage", "Grade"]
    elif "inventory" in p_lower or "bom" in p_lower or "bill of material" in p_lower:
        instructions.preset_name = "Inventory BOM"
        instructions.target_columns = ["Part No", "Part Description", "Quantity", "Supplier", "Unit Cost", "Total Value"]

    # 2. Extract Requested Columns if specified with "columns:" or comma list
    col_match = re.search(r"columns?\s*[:=]\s*([^\.\n]+)", clean_p, re.IGNORECASE)
    if col_match:
        raw_cols = col_match.group(1)
        parsed_cols = [c.strip() for c in re.split(r"[,;\n]", raw_cols) if c.strip()]
        if parsed_cols:
            instructions.target_columns = parsed_cols

    # 3. Extract Calculations
    # Match patterns like: Calculate Percentage = Total / 500 * 100 or Total = Qty * Price or GST = Total * 0.18
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
    # Match: If Status is PASS color green / If Percentage > 90 add Grade A
    cond_matches = re.findall(r"if\s+([A-Za-z0-9_\s]+)\s+(is|>|<|==|above|below)\s+([A-Za-z0-9_%]+)\s+(color|add|set)\s+([A-Za-z0-9_\s]+)", clean_p, re.IGNORECASE)
    for col, cond, val, act, tgt in cond_matches:
        instructions.conditional_rules.append(ConditionalRule(
            column_name=col.strip(),
            condition=cond.strip(),
            value=val.strip(),
            action_type=act.strip().lower(),
            action_target=tgt.strip()
        ))

    logger.info(f"Parsed prompt: Preset='{instructions.preset_name}', Columns={instructions.target_columns}")
    return instructions
