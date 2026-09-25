"""
AI Mapping Engine module for Universal AI Excel Format Converter.
Performs smart column detection, normalization, synonym dictionary lookup,
and fuzzy matching similarity scoring between input and output column schemas.
"""

from difflib import SequenceMatcher
from typing import List, Dict, Tuple, Any, Optional
import config
from logger import logger
from utils import normalize_column_name


def compute_string_similarity(s1: str, s2: str) -> float:
    """Computes similarity score between two strings (0.0 to 1.0)."""
    n1 = normalize_column_name(s1)
    n2 = normalize_column_name(s2)

    if not n1 or not n2:
        return 0.0

    if n1 == n2:
        return 1.0

    # Token overlap check
    t1 = set(n1.split())
    t2 = set(n2.split())
    if t1 == t2:
        return 0.95

    # Substring check
    if n1 in n2 or n2 in n1:
        return 0.85

    # Sequence matcher fallback
    return SequenceMatcher(None, n1, n2).ratio()


def find_synonym_category(col_name: str) -> Optional[str]:
    """Finds matching synonym category for a column name from config.COLUMN_SYNONYMS."""
    norm = normalize_column_name(col_name)
    for category, synonyms in config.COLUMN_SYNONYMS.items():
        for syn in synonyms:
            if syn == norm or syn in norm or norm in syn:
                return category
    return None


def suggest_column_mappings(input_columns: List[str], target_columns: List[str]) -> Dict[str, Tuple[str, float]]:
    """
    Generates AI suggested column mappings from target columns to input columns.
    Returns dictionary:
        { target_col: (matched_input_col, confidence_score) }
    """
    mappings: Dict[str, Tuple[str, float]] = {}

    for target in target_columns:
        best_match: Optional[str] = None
        best_score: float = 0.0

        target_cat = find_synonym_category(target)

        for inp in input_columns:
            # Check category match
            inp_cat = find_synonym_category(inp)
            cat_bonus = 0.3 if (target_cat and inp_cat and target_cat == inp_cat) else 0.0

            similarity = compute_string_similarity(target, inp)
            final_score = min(1.0, similarity + cat_bonus)

            if final_score > best_score:
                best_score = final_score
                best_match = inp

        if best_match and best_score >= 0.35:
            mappings[target] = (best_match, round(best_score, 2))
        else:
            mappings[target] = ("", 0.0)

    logger.info(f"AI Column Mapping Engine generated {len(mappings)} mappings.")
    return mappings
