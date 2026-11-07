"""
formulas.py
===========
Live Excel Formula Generation Module.
Writes real dynamic Excel formulas (XLOOKUP, VLOOKUP, SUMIFS, COUNTIFS, IFERROR, INDEX-MATCH)
into openpyxl cell grids so the exported workbook remains alive and responsive.
Provides bilingual (Hindi & English) explanations for Excel-learning.
"""

from __future__ import annotations
from typing import Dict, List, Optional, Tuple, Any
import openpyxl
from openpyxl.utils import get_column_letter


def build_xlookup_formula(
    lookup_cell: str,
    lookup_range: str,
    return_range: str,
    if_not_found: str = '""'
) -> str:
    """Builds standard modern Excel =XLOOKUP(...) formula."""
    return f'=XLOOKUP({lookup_cell}, {lookup_range}, {return_range}, {if_not_found}, 0)'


def build_vlookup_formula(
    lookup_cell: str,
    table_range: str,
    col_index: int,
    exact_match: bool = True
) -> str:
    """Builds standard Excel =IFERROR(VLOOKUP(...), "") formula."""
    range_lookup = "FALSE" if exact_match else "TRUE"
    return f'=IFERROR(VLOOKUP({lookup_cell}, {table_range}, {col_index}, {range_lookup}), "")'


def build_sumifs_formula(
    sum_range: str,
    criteria_range: str,
    criteria_cell: str
) -> str:
    """Builds Excel =SUMIFS(...) formula for conditional totals."""
    return f'=SUMIFS({sum_range}, {criteria_range}, {criteria_cell})'


def build_countifs_formula(
    criteria_range: str,
    criteria_cell: str
) -> str:
    """Builds Excel =COUNTIFS(...) formula."""
    return f'=COUNTIFS({criteria_range}, {criteria_cell})'


def build_calculated_cell_formula(
    formula_template: str,
    row_idx: int,
    col_letter_map: Dict[str, str]
) -> str:
    """
    Replaces column names in formulas like 'Quantity * Unit_Price'
    with Excel cell references like 'D2 * E2'.
    """
    expr = formula_template
    # Replace column names from longest to shortest to avoid partial replacements
    sorted_cols = sorted(col_letter_map.keys(), key=len, reverse=True)
    for col_name in sorted_cols:
        cell_ref = f"{col_letter_map[col_name]}{row_idx}"
        # Word boundary replacement
        import re
        pattern = r'\b' + re.escape(col_name) + r'\b'
        expr = re.sub(pattern, cell_ref, expr, flags=re.IGNORECASE)
    return f"={expr}"


def explain_formula(formula_type: str, details: Dict[str, Any]) -> Dict[str, str]:
    """
    Provides plain Hindi and English educational explanations for Excel formulas.
    Supports Excel-Learning Mode ('Ye kaise kiya?').
    """
    f_type = formula_type.upper()
    
    if f_type in ["XLOOKUP", "VLOOKUP"]:
        lookup_col = details.get("lookup_col", "Key")
        target_col = details.get("target_col", "Value")
        source_tbl = details.get("source_table", "Prices")
        
        return {
            "en": (
                f"Formula: =XLOOKUP({lookup_col}2, {source_tbl}!A:A, {source_tbl}!B:B, '')\n"
                f"Explanation: Excel searches for '{lookup_col}' in the source table and retrieves the matching '{target_col}'. "
                f"If the value is missing, it returns blank instead of #N/A."
            ),
            "hi": (
                f"फॉर्मूला: =XLOOKUP({lookup_col}2, {source_tbl}!A:A, {source_tbl}!B:B, '')\n"
                f"सरल व्याख्या: यह फॉर्मूला {source_tbl} टेबल में '{lookup_col}' को ढूंढता है और उसके सामने वाला '{target_col}' निकाल लाता है। "
                f"अगर नाम न मिले तो एरर (#N/A) के बदले सेल खाली रहेगा।"
            )
        }
    
    elif f_type in ["SUMIFS", "PIVOT"]:
        sum_col = details.get("sum_col", "Amount")
        group_col = details.get("group_col", "City")
        return {
            "en": (
                f"Formula: =SUMIFS({sum_col}_Range, {group_col}_Range, Criteria_Cell)\n"
                f"Explanation: Calculates the total of '{sum_col}' only where '{group_col}' matches the given row category."
            ),
            "hi": (
                f"फॉर्मूला: =SUMIFS({sum_col}_रेंज, {group_col}_रेंज, क्राइटेरिया_सेल)\n"
                f"सरल व्याख्या: यह केवल उन पंक्तियों के '{sum_col}' को जोड़ता है जहां '{group_col}' उस शहर या कैटेगरी से मेल खाता है।"
            )
        }
    
    elif f_type in ["CALCULATE"]:
        formula = details.get("formula", "A * B")
        target = details.get("target_col", "Total")
        return {
            "en": (
                f"Formula: ={formula}\n"
                f"Explanation: Evaluates '{formula}' dynamically for each row to populate column '{target}'."
            ),
            "hi": (
                f"फॉर्मूला: ={formula}\n"
                f"सरल व्याख्या: प्रत्येक पंक्ति में दी गई संख्याओं का गुणा या जोड़ करके '{target}' कॉलम बनाता है।"
            )
        }
    
    return {
        "en": f"Dynamic Excel formula applied.",
        "hi": f"डायनामिक एक्सेल फॉर्मूला लागू किया गया।"
    }
