"""
actions.py
==========
Core Data Manipulation Engine for Voice-Controlled Excel Assistant.
Executes predefined, safe pandas and openpyxl operations based on validated JSON plans.
No arbitrary eval/exec is ever permitted.

Features:
- VLOOKUP / XLOOKUP (Exact & Fuzzy join, trimming, casing, type coercion)
- Sorting (Multi-column, asc/desc)
- Set Operations (Common values, Unique to A/B, Duplicates across tables, Fuzzy matching)
- Filtering (equals, contains, gt, lt, between, date range)
- Cleaning (Deduplication, blanks filling/removal, trimming, case conversions)
- Pivot & GroupBy (sum, mean, count, min, max)
- Column Transformations (safe math formulas, splitting, merging, renaming, dropping)
- Sheet / Table Merging (concatenate / append)
- Excel Export with Conditional Formatting (openpyxl highlighting)
"""

from __future__ import annotations
import re
import io
from typing import Any, Dict, List, Optional, Tuple, Union
import pandas as pd
import numpy as np

try:
    from rapidfuzz import process, fuzz
except ImportError:
    process = None
    fuzz = None

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils.dataframe import dataframe_to_rows
except ImportError:
    openpyxl = None


class ActionExecutionError(Exception):
    """Custom exception containing user-friendly Hindi & English error messages."""
    def __init__(self, message_en: str, message_hi: Optional[str] = None):
        super().__init__(message_en)
        self.message_en = message_en
        self.message_hi = message_hi or message_en


def normalize_lookup_key(val: Any) -> str:
    """Normalize keys for robust matching across strings, numbers, and floats."""
    if pd.isna(val) or val is None:
        return ""
    # Convert float like 123.0 to '123' if it represents an integer
    if isinstance(val, (float, np.floating)) and val.is_integer():
        return str(int(val)).strip().lower()
    return str(val).strip().lower()


def find_column_case_insensitive(df: pd.DataFrame, target_col: str) -> Optional[str]:
    """Find a column in a dataframe ignoring case and extra whitespace."""
    target_clean = str(target_col).strip().lower()
    for col in df.columns:
        if str(col).strip().lower() == target_clean:
            return str(col)
    return None


def execute_vlookup(
    left_df: pd.DataFrame,
    right_df: pd.DataFrame,
    left_key: str,
    right_key: str,
    return_columns: List[str],
    fuzzy: bool = False,
    similarity_threshold: float = 80.0
) -> Tuple[pd.DataFrame, str]:
    """
    Performs a VLOOKUP / XLOOKUP equivalent (Left Join).
    Handles:
    - Whitespace trimming
    - Case differences
    - Number vs text data type mismatches
    - Optional fuzzy matching via rapidfuzz
    """
    left_actual_key = find_column_case_insensitive(left_df, left_key)
    if not left_actual_key:
        raise ActionExecutionError(
            f"Left table column '{left_key}' not found. Available: {list(left_df.columns)}",
            f"Left table mein column '{left_key}' nahi mila. Uplabdh columns: {list(left_df.columns)}"
        )

    right_actual_key = find_column_case_insensitive(right_df, right_key)
    if not right_actual_key:
        raise ActionExecutionError(
            f"Right table column '{right_key}' not found. Available: {list(right_df.columns)}",
            f"Right table mein column '{right_key}' nahi mila. Uplabdh columns: {list(right_df.columns)}"
        )

    # Validate return columns
    actual_return_cols: List[str] = []
    for col in return_columns:
        matched = find_column_case_insensitive(right_df, col)
        if matched and matched != right_actual_key:
            actual_return_cols.append(matched)

    if not actual_return_cols:
        raise ActionExecutionError(
            f"None of the return columns {return_columns} were found in the right table. Available: {list(right_df.columns)}",
            f"Return columns {return_columns} right table mein nahi mile. Uplabdh columns: {list(right_df.columns)}"
        )

    result_df = left_df.copy()

    if fuzzy and process and fuzz:
        # Fuzzy matching using rapidfuzz
        right_lookup_map: Dict[str, Any] = {}
        right_keys_list = [normalize_lookup_key(k) for k in right_df[right_actual_key].dropna().unique()]

        # Pre-build row map for right table (deduplicated by key)
        right_sub = right_df.drop_duplicates(subset=[right_actual_key])
        right_sub_dict = {}
        for _, r in right_sub.iterrows():
            k_norm = normalize_lookup_key(r[right_actual_key])
            right_sub_dict[k_norm] = {c: r[c] for c in actual_return_cols}

        matched_rows: List[Dict[str, Any]] = []
        for val in left_df[left_actual_key]:
            val_norm = normalize_lookup_key(val)
            if not val_norm:
                matched_rows.append({c: np.nan for c in actual_return_cols})
                continue

            # Exact match first
            if val_norm in right_sub_dict:
                matched_rows.append(right_sub_dict[val_norm])
                continue

            # Fuzzy match fallback
            match = process.extractOne(val_norm, right_keys_list, scorer=fuzz.token_sort_ratio)
            if match and match[1] >= similarity_threshold:
                matched_key = match[0]
                matched_rows.append(right_sub_dict.get(matched_key, {c: np.nan for c in actual_return_cols}))
            else:
                matched_rows.append({c: np.nan for c in actual_return_cols})

        match_df = pd.DataFrame(matched_rows, index=result_df.index)
        for c in actual_return_cols:
            result_col_name = c if c not in result_df.columns else f"{c}_lookup"
            result_df[result_col_name] = match_df[c]

        msg = f"VLOOKUP completed with fuzzy matching (threshold {similarity_threshold}%). Brought columns: {actual_return_cols}."
    else:
        # Standard robust match (whitespace trimmed, case-insensitive, type-coerced)
        left_temp_key = "_lookup_norm_left_key"
        right_temp_key = "_lookup_norm_right_key"

        result_df[left_temp_key] = result_df[left_actual_key].apply(normalize_lookup_key)
        right_copy = right_df.copy()
        right_copy[right_temp_key] = right_copy[right_actual_key].apply(normalize_lookup_key)

        # Drop duplicates on the right lookup key to prevent row explosion
        right_subset = right_copy.drop_duplicates(subset=[right_temp_key])[[right_temp_key] + actual_return_cols]

        merged = pd.merge(
            result_df,
            right_subset,
            how="left",
            left_on=left_temp_key,
            right_on=right_temp_key,
            suffixes=("", "_from_lookup")
        )

        merged.drop(columns=[left_temp_key, right_temp_key], inplace=True, errors="ignore")
        result_df = merged
        msg = f"VLOOKUP successfully matched '{left_actual_key}' with '{right_actual_key}'. Added columns: {actual_return_cols}."

    return result_df, msg


def execute_sort(
    df: pd.DataFrame,
    columns: List[str],
    ascending: Union[bool, List[bool]] = True
) -> Tuple[pd.DataFrame, str]:
    """Sorts dataframe by one or multiple columns."""
    actual_cols: List[str] = []
    for col in columns:
        matched = find_column_case_insensitive(df, col)
        if not matched:
            raise ActionExecutionError(
                f"Sort column '{col}' not found. Available columns: {list(df.columns)}",
                f"Sort karne ke liye column '{col}' nahi mila. Uplabdh columns: {list(df.columns)}"
            )
        actual_cols.append(matched)

    if isinstance(ascending, list) and len(ascending) != len(actual_cols):
        ascending = ascending[0] if len(ascending) > 0 else True

    sorted_df = df.sort_values(by=actual_cols, ascending=ascending).reset_index(drop=True)
    asc_text = "ascending (A-Z, 0-9)" if (ascending is True or ascending == [True]) else "descending (Z-A, 9-0)"
    return sorted_df, f"Data sorted by {actual_cols} in {asc_text} order."


def execute_find_similar_or_common(
    table_a: pd.DataFrame,
    table_b: pd.DataFrame,
    col_a: str,
    col_b: str,
    mode: str = "common",  # 'common', 'unique_to_a', 'unique_to_b', 'duplicates_all'
    fuzzy: bool = False,
    similarity_threshold: float = 80.0
) -> Tuple[pd.DataFrame, str]:
    """
    Finds common or unique items across two tables.
    Modes:
    - common: Rows in table A whose key exists in table B (Inner join / Intersection)
    - unique_to_a: Rows in table A whose key does NOT exist in table B
    - unique_to_b: Rows in table B whose key does NOT exist in table A
    - duplicates_all: Merged records appearing more than once
    """
    actual_col_a = find_column_case_insensitive(table_a, col_a)
    actual_col_b = find_column_case_insensitive(table_b, col_b)

    if not actual_col_a:
        raise ActionExecutionError(f"Column '{col_a}' not found in Table A. Available: {list(table_a.columns)}")
    if not actual_col_b:
        raise ActionExecutionError(f"Column '{col_b}' not found in Table B. Available: {list(table_b.columns)}")

    # Extract normalized sets
    norm_a = table_a[actual_col_a].dropna().apply(normalize_lookup_key).unique()
    norm_b = table_b[actual_col_b].dropna().apply(normalize_lookup_key).unique()

    if fuzzy and process and fuzz:
        matched_in_b = set()
        for item_a in norm_a:
            if not item_a:
                continue
            if item_a in norm_b:
                matched_in_b.add(item_a)
            else:
                best = process.extractOne(item_a, norm_b, scorer=fuzz.token_sort_ratio)
                if best and best[1] >= similarity_threshold:
                    matched_in_b.add(item_a)
        common_keys = matched_in_b
    else:
        common_keys = set(norm_a).intersection(set(norm_b))

    temp_key = "_norm_filter_key"
    temp_df_a = table_a.copy()
    temp_df_a[temp_key] = temp_df_a[actual_col_a].apply(normalize_lookup_key)

    temp_df_b = table_b.copy()
    temp_df_b[temp_key] = temp_df_b[actual_col_b].apply(normalize_lookup_key)

    if mode == "common":
        result = temp_df_a[temp_df_a[temp_key].isin(common_keys)].copy()
        result.drop(columns=[temp_key], inplace=True)
        msg = f"Found {len(result)} common rows where '{actual_col_a}' in Table A matches '{actual_col_b}' in Table B."
    elif mode == "unique_to_a":
        result = temp_df_a[~temp_df_a[temp_key].isin(common_keys)].copy()
        result.drop(columns=[temp_key], inplace=True)
        msg = f"Found {len(result)} rows unique to Table A (absent in Table B)."
    elif mode == "unique_to_b":
        result = temp_df_b[~temp_df_b[temp_key].isin(common_keys)].copy()
        result.drop(columns=[temp_key], inplace=True)
        msg = f"Found {len(result)} rows unique to Table B (absent in Table A)."
    else:
        # Default intersection
        result = temp_df_a[temp_df_a[temp_key].isin(common_keys)].copy()
        result.drop(columns=[temp_key], inplace=True)
        msg = f"Extracted {len(result)} common items."

    return result.reset_index(drop=True), msg


def execute_filter(
    df: pd.DataFrame,
    column: str,
    operator: str,
    value: Any,
    value2: Optional[Any] = None
) -> Tuple[pd.DataFrame, str]:
    """
    Filters rows by condition:
    operators: 'equals', 'not_equals', 'contains', 'not_contains', 'greater_than',
               'less_than', 'greater_equal', 'less_equal', 'between', 'date_range'
    """
    actual_col = find_column_case_insensitive(df, column)
    if not actual_col:
        raise ActionExecutionError(
            f"Filter column '{column}' not found. Available: {list(df.columns)}",
            f"Filter karne ke liye column '{column}' nahi mila. Uplabdh columns: {list(df.columns)}"
        )

    series = df[actual_col]
    op = operator.lower().strip()
    initial_rows = len(df)

    if op in ["equals", "==", "eq"]:
        if isinstance(value, str):
            mask = series.astype(str).str.strip().str.lower() == str(value).strip().lower()
        else:
            mask = series == value
        result = df[mask]
        msg = f"Filtered rows where '{actual_col}' equals '{value}' ({len(result)} of {initial_rows} rows kept)."

    elif op in ["not_equals", "!=", "neq"]:
        if isinstance(value, str):
            mask = series.astype(str).str.strip().str.lower() != str(value).strip().lower()
        else:
            mask = series != value
        result = df[mask]
        msg = f"Filtered rows where '{actual_col}' does not equal '{value}'."

    elif op in ["contains", "like"]:
        val_str = str(value).strip().lower()
        mask = series.astype(str).str.lower().str.contains(val_str, na=False, regex=False)
        result = df[mask]
        msg = f"Filtered rows where '{actual_col}' contains '{value}'."

    elif op in ["not_contains"]:
        val_str = str(value).strip().lower()
        mask = ~series.astype(str).str.lower().str.contains(val_str, na=False, regex=False)
        result = df[mask]
        msg = f"Filtered rows where '{actual_col}' does not contain '{value}'."

    elif op in ["greater_than", ">", "gt"]:
        num_val = float(value)
        num_series = pd.to_numeric(series, errors="coerce")
        result = df[num_series > num_val]
        msg = f"Filtered rows where '{actual_col}' > {num_val}."

    elif op in ["less_than", "<", "lt"]:
        num_val = float(value)
        num_series = pd.to_numeric(series, errors="coerce")
        result = df[num_series < num_val]
        msg = f"Filtered rows where '{actual_col}' < {num_val}."

    elif op in ["greater_equal", ">=", "gte"]:
        num_val = float(value)
        num_series = pd.to_numeric(series, errors="coerce")
        result = df[num_series >= num_val]
        msg = f"Filtered rows where '{actual_col}' >= {num_val}."

    elif op in ["less_equal", "<=", "lte"]:
        num_val = float(value)
        num_series = pd.to_numeric(series, errors="coerce")
        result = df[num_series <= num_val]
        msg = f"Filtered rows where '{actual_col}' <= {num_val}."

    elif op in ["between"]:
        low, high = float(value), float(value2 if value2 is not None else value)
        num_series = pd.to_numeric(series, errors="coerce")
        result = df[(num_series >= low) & (num_series <= high)]
        msg = f"Filtered rows where '{actual_col}' between {low} and {high}."

    elif op in ["date_range"]:
        date_series = pd.to_datetime(series, errors="coerce")
        d1 = pd.to_datetime(value)
        d2 = pd.to_datetime(value2) if value2 else d1
        result = df[(date_series >= d1) & (date_series <= d2)]
        msg = f"Filtered rows where date '{actual_col}' between {value} and {value2}."

    else:
        raise ActionExecutionError(f"Unsupported filter operator: '{operator}'")

    return result.reset_index(drop=True), msg


def execute_cleaning(
    df: pd.DataFrame,
    operation: str,
    columns: Optional[List[str]] = None,
    fill_value: Any = None
) -> Tuple[pd.DataFrame, str]:
    """
    Cleans data:
    - 'remove_duplicates': drop duplicate rows
    - 'drop_blanks': drop rows where specified columns (or any) are blank
    - 'fill_blanks': fill NaN/blank values with fill_value or mean/median
    - 'trim_spaces': strip leading and trailing whitespace from string columns
    - 'change_case': transform text case ('upper', 'lower', 'title')
    """
    result = df.copy()
    op = operation.lower().strip()
    init_count = len(df)

    actual_cols = []
    if columns:
        for c in columns:
            found = find_column_case_insensitive(df, c)
            if found:
                actual_cols.append(found)

    if op in ["remove_duplicates", "dedup", "drop_duplicates"]:
        subset = actual_cols if actual_cols else None
        result = result.drop_duplicates(subset=subset).reset_index(drop=True)
        removed = init_count - len(result)
        msg = f"Removed {removed} duplicate rows. Remaining: {len(result)} rows."

    elif op in ["drop_blanks", "remove_blanks"]:
        subset = actual_cols if actual_cols else None
        result = result.dropna(subset=subset).reset_index(drop=True)
        removed = init_count - len(result)
        msg = f"Dropped {removed} blank/null rows. Remaining: {len(result)} rows."

    elif op in ["fill_blanks"]:
        target_cols = actual_cols if actual_cols else list(result.columns)
        for col in target_cols:
            if fill_value == "mean":
                mean_val = pd.to_numeric(result[col], errors="coerce").mean()
                result[col] = result[col].fillna(mean_val)
            elif fill_value == "median":
                med_val = pd.to_numeric(result[col], errors="coerce").median()
                result[col] = result[col].fillna(med_val)
            elif fill_value == "mode":
                mode_val = result[col].mode()[0] if not result[col].mode().empty else ""
                result[col] = result[col].fillna(mode_val)
            else:
                val = 0 if fill_value is None else fill_value
                result[col] = result[col].fillna(val)
        msg = f"Filled blank cells in {target_cols} with '{fill_value}'."

    elif op in ["trim_spaces", "strip"]:
        target_cols = actual_cols if actual_cols else list(result.select_dtypes(include=["object"]).columns)
        for col in target_cols:
            result[col] = result[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
        msg = f"Trimmed whitespace from columns: {target_cols}."

    elif op in ["change_case"]:
        case_type = str(fill_value).lower() if fill_value else "title"
        target_cols = actual_cols if actual_cols else list(result.select_dtypes(include=["object"]).columns)
        for col in target_cols:
            if case_type == "upper":
                result[col] = result[col].apply(lambda x: x.upper() if isinstance(x, str) else x)
            elif case_type == "lower":
                result[col] = result[col].apply(lambda x: x.lower() if isinstance(x, str) else x)
            else:
                result[col] = result[col].apply(lambda x: x.title() if isinstance(x, str) else x)
        msg = f"Converted case to {case_type} for columns: {target_cols}."

    else:
        raise ActionExecutionError(f"Unsupported cleaning action: '{operation}'")

    return result, msg


def execute_pivot_table(
    df: pd.DataFrame,
    index_columns: List[str],
    values_columns: List[str],
    agg_func: str = "sum",
    columns_pivot: Optional[List[str]] = None
) -> Tuple[pd.DataFrame, str]:
    """
    Creates a pivot table or group-by summary.
    agg_func: 'sum', 'count', 'mean' / 'average', 'min', 'max'
    """
    actual_indexes = [find_column_case_insensitive(df, c) for c in index_columns if find_column_case_insensitive(df, c)]
    if not actual_indexes:
        raise ActionExecutionError(f"Group-by index columns {index_columns} not found.")

    actual_values = [find_column_case_insensitive(df, c) for c in values_columns if find_column_case_insensitive(df, c)]
    if not actual_values:
        raise ActionExecutionError(f"Values columns {values_columns} not found in dataframe.")

    func_map = {
        "sum": "sum",
        "avg": "mean",
        "average": "mean",
        "mean": "mean",
        "count": "count",
        "min": "min",
        "max": "max"
    }
    agg = func_map.get(agg_func.lower().strip(), "sum")

    # Coerce values columns to numeric if performing sum or average
    df_calc = df.copy()
    if agg in ["sum", "mean", "min", "max"]:
        for v in actual_values:
            df_calc[v] = pd.to_numeric(df_calc[v], errors="coerce")

    actual_pivot_cols = None
    if columns_pivot:
        actual_pivot_cols = [find_column_case_insensitive(df, c) for c in columns_pivot if find_column_case_insensitive(df, c)]

    if actual_pivot_cols:
        pivot_df = pd.pivot_table(
            df_calc,
            index=actual_indexes,
            columns=actual_pivot_cols,
            values=actual_values,
            aggfunc=agg,
            fill_value=0
        )
        pivot_df = pivot_df.reset_index()
    else:
        # Straightforward groupby
        grouped = df_calc.groupby(actual_indexes, as_index=False)[actual_values].agg(agg)
        pivot_df = grouped

    msg = f"Created summary table grouped by {actual_indexes} calculating {agg} of {actual_values}."
    return pivot_df, msg


def execute_column_transform(
    df: pd.DataFrame,
    action_type: str,
    target_column: str,
    source_columns: Optional[List[str]] = None,
    formula: Optional[str] = None,
    delimiter: Optional[str] = None
) -> Tuple[pd.DataFrame, str]:
    """
    Column transformations:
    - 'calculate': Safely evaluate mathematical expressions (e.g. Total = Quantity * Price)
    - 'merge': Combine multiple columns into one with a separator
    - 'split': Split one column into multiple based on a delimiter
    - 'rename': Rename column
    - 'delete': Drop column
    """
    result = df.copy()
    action = action_type.lower().strip()

    if action in ["calculate", "add_column"]:
        if not formula:
            raise ActionExecutionError("Calculated column requires a formula (e.g. 'Quantity * Rate').")

        # Parse formula safely without arbitrary code execution
        # Example formula: "Quantity * Price" or "A + B" or "A - B"
        # Match tokens against dataframe column names
        tokens = re.findall(r"([A-Za-z0-9_\s]+|[\+\-\*/\(\)])", formula)
        
        # Build safe pandas arithmetic evaluation
        # Replace column names with df['col']
        eval_expression = formula
        for col in df.columns:
            # Replace exact word occurrences
            pattern = r'\b' + re.escape(str(col)) + r'\b'
            if re.search(pattern, eval_expression, re.IGNORECASE):
                # Ensure numeric column
                result[col] = pd.to_numeric(result[col], errors="coerce").fillna(0)
                eval_expression = re.sub(pattern, f"result['{col}']", eval_expression, flags=re.IGNORECASE)

        # Allow only safe arithmetic operators and syntax
        allowed_chars = set("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_[]'\"+-*/. ()")
        if not all(ch in allowed_chars for ch in eval_expression):
            raise ActionExecutionError(f"Formula contains forbidden characters: '{formula}'")

        try:
            # Execute with restricted scope
            computed_series = eval(eval_expression, {"__builtins__": None, "result": result, "np": np})
            result[target_column] = computed_series
            msg = f"Added calculated column '{target_column}' using formula: {formula}."
        except Exception as e:
            raise ActionExecutionError(f"Failed to calculate formula '{formula}': {str(e)}")

    elif action in ["merge", "combine"]:
        if not source_columns:
            raise ActionExecutionError("Merge columns requires 'source_columns'.")
        sep = delimiter if delimiter is not None else " "
        actual_sources = [find_column_case_insensitive(df, c) for c in source_columns if find_column_case_insensitive(df, c)]
        result[target_column] = result[actual_sources].astype(str).agg(sep.join, axis=1)
        msg = f"Merged columns {actual_sources} into '{target_column}' using separator '{sep}'."

    elif action in ["split"]:
        actual_target = find_column_case_insensitive(df, target_column)
        if not actual_target:
            raise ActionExecutionError(f"Column '{target_column}' not found.")
        sep = delimiter if delimiter else ","
        split_df = result[actual_target].astype(str).str.split(sep, expand=True)
        col_names = [f"{actual_target}_{i+1}" for i in range(split_df.shape[1])]
        for i, name in enumerate(col_names):
            result[name] = split_df[i].str.strip()
        msg = f"Split column '{actual_target}' into {col_names} by '{sep}'."

    elif action in ["rename"]:
        actual_target = find_column_case_insensitive(df, target_column)
        new_name = formula or (source_columns[0] if source_columns else None)
        if not actual_target or not new_name:
            raise ActionExecutionError("Rename requires target column and new name.")
        result.rename(columns={actual_target: new_name}, inplace=True)
        msg = f"Renamed column '{actual_target}' to '{new_name}'."

    elif action in ["delete", "drop"]:
        actual_target = find_column_case_insensitive(df, target_column)
        if not actual_target:
            raise ActionExecutionError(f"Column '{target_column}' not found to delete.")
        result.drop(columns=[actual_target], inplace=True)
        msg = f"Deleted column '{actual_target}'."

    else:
        raise ActionExecutionError(f"Unsupported column transformation: '{action_type}'")

    return result, msg


def execute_merge_sheets(
    tables: List[pd.DataFrame],
    sheet_names: List[str],
    merge_type: str = "append"  # 'append' (union/concat) or 'outer'
) -> Tuple[pd.DataFrame, str]:
    """Merges or concatenates multiple sheets or tables."""
    if not tables:
        raise ActionExecutionError("No tables provided to merge.")

    if merge_type in ["append", "concat"]:
        annotated = []
        for df, name in zip(tables, sheet_names):
            df_copy = df.copy()
            df_copy["_SourceSheet"] = name
            annotated.append(df_copy)
        merged = pd.concat(annotated, ignore_index=True)
        msg = f"Appended {len(tables)} sheets ({sheet_names}) into one table. Total rows: {len(merged)}."
    else:
        # Outer join on shared columns
        merged = pd.concat(tables, axis=1)
        msg = f"Merged {len(tables)} sheets side-by-side."

    return merged, msg


def export_dataframe_with_formatting(
    df: pd.DataFrame,
    highlight_rules: Optional[List[Dict[str, Any]]] = None
) -> io.BytesIO:
    """
    Exports DataFrame to an Excel file (.xlsx) with:
    - Styled headers (dark navy background, bold white text)
    - Auto-fitted column widths
    - Conditional highlighting (yellow for duplicates, green for positive conditions)
    - Clean borders and gridlines
    """
    output = io.BytesIO()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Assistant Result"
    ws.views.sheetView[0].showGridLines = True

    # Style definitions
    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    border_thin = Side(style="thin", color="CBD5E1")
    cell_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)

    highlight_yellow = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid") # soft yellow
    highlight_green = PatternFill(start_color="BBF7D0", end_color="BBF7D0", fill_type="solid")  # soft green
    highlight_coral = PatternFill(start_color="FECDD3", end_color="FECDD3", fill_type="solid")  # soft red/pink

    # Write dataframe rows
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=1):
        ws.append(row)
        for c_idx in range(1, len(row) + 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.border = cell_border
            if r_idx == 1:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            else:
                cell.font = data_font
                cell.alignment = Alignment(vertical="center")

    # Apply conditional highlighting rules if provided
    if highlight_rules:
        for rule in highlight_rules:
            rule_type = rule.get("type", "")
            target_col = rule.get("column", "")
            col_idx = None
            if target_col:
                for idx, col_name in enumerate(df.columns, start=1):
                    if str(col_name).strip().lower() == str(target_col).strip().lower():
                        col_idx = idx
                        break

            color_name = rule.get("color", "yellow").lower()
            fill_style = highlight_green if color_name == "green" else (highlight_coral if color_name in ["red", "coral"] else highlight_yellow)

            if rule_type == "duplicates" and col_idx:
                # Find duplicates in that column
                dup_mask = df.duplicated(subset=[df.columns[col_idx - 1]], keep=False)
                for r_idx, is_dup in enumerate(dup_mask, start=2):
                    if is_dup:
                        ws.cell(row=r_idx, column=col_idx).fill = fill_style

            elif rule_type == "condition" and col_idx:
                threshold = rule.get("threshold", 0)
                cond_op = rule.get("operator", "gt")
                for r_idx in range(2, len(df) + 2):
                    val = ws.cell(row=r_idx, column=col_idx).value
                    try:
                        num_val = float(val)
                        matches = (num_val > threshold) if cond_op == "gt" else (num_val < threshold)
                        if matches:
                            ws.cell(row=r_idx, column=col_idx).fill = fill_style
                    except (ValueError, TypeError):
                        pass

    # Auto-adjust column width and freeze top row
    ws.freeze_panes = "A2"
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            val_str = str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(output)
    output.seek(0)
    return output


def execute_compare_versions(
    df_old: pd.DataFrame,
    df_new: pd.DataFrame,
    key_column: Optional[str] = None
) -> Tuple[pd.DataFrame, str, Dict[str, Any]]:
    """
    Compares two versions of a dataset (Old vs New).
    Detects added rows, removed rows, and modified values.
    """
    actual_key = find_column_case_insensitive(df_old, key_column) if key_column else None
    if not actual_key:
        # Default to first column as key
        actual_key = df_old.columns[0]

    old_norm = df_old.copy()
    new_norm = df_new.copy()

    old_norm["_key_norm"] = old_norm[actual_key].apply(normalize_lookup_key)
    new_norm["_key_norm"] = new_norm[actual_key].apply(normalize_lookup_key)

    old_keys = set(old_norm["_key_norm"].dropna())
    new_keys = set(new_norm["_key_norm"].dropna())

    added_keys = new_keys - old_keys
    removed_keys = old_keys - new_keys
    common_keys = old_keys.intersection(new_keys)

    diff_records = []

    # Added rows
    for _, r in new_norm[new_norm["_key_norm"].isin(added_keys)].iterrows():
        rec = {c: r[c] for c in df_new.columns}
        rec["_Diff_Status"] = "ADDED (New in Version 2)"
        diff_records.append(rec)

    # Removed rows
    for _, r in old_norm[old_norm["_key_norm"].isin(removed_keys)].iterrows():
        rec = {c: r[c] for c in df_old.columns}
        rec["_Diff_Status"] = "REMOVED (Missing in Version 2)"
        diff_records.append(rec)

    # Modified rows
    old_indexed = old_norm.set_index("_key_norm")
    new_indexed = new_norm.set_index("_key_norm")
    shared_cols = [c for c in df_old.columns if c in df_new.columns and c != actual_key]

    modified_count = 0
    for k in common_keys:
        old_r = old_indexed.loc[k]
        new_r = new_indexed.loc[k]
        if isinstance(old_r, pd.DataFrame):
            old_r = old_r.iloc[0]
        if isinstance(new_r, pd.DataFrame):
            new_r = new_r.iloc[0]

        changes = []
        for col in shared_cols:
            val_o = str(old_r[col]) if pd.notna(old_r[col]) else ""
            val_n = str(new_r[col]) if pd.notna(new_r[col]) else ""
            if val_o.strip() != val_n.strip():
                changes.append(f"{col}: '{val_o}' ➔ '{val_n}'")

        if changes:
            modified_count += 1
            rec = {c: new_r[c] for c in df_new.columns}
            rec["_Diff_Status"] = f"MODIFIED ({', '.join(changes)})"
            diff_records.append(rec)

    result_df = pd.DataFrame(diff_records) if diff_records else pd.DataFrame(columns=list(df_new.columns) + ["_Diff_Status"])
    summary_msg = f"Comparison complete: {len(added_keys)} added, {len(removed_keys)} removed, {modified_count} modified rows."
    summary_hi = f"फ़ाइल तुलना पूरी हुई: {len(added_keys)} नई जोड़ी गई, {len(removed_keys)} हटाई गई, {modified_count} बदली गई पंक्तियाँ।"

    meta = {
        "added_count": len(added_keys),
        "removed_count": len(removed_keys),
        "modified_count": modified_count,
        "output_name": "Version_Changes"
    }

    return result_df, summary_msg, meta


def execute_detect_anomalies(
    df: pd.DataFrame,
    columns: Optional[List[str]] = None
) -> Tuple[pd.DataFrame, str]:
    """
    Detects unusual anomalies or possible data-entry mistakes:
    - Extreme statistical outliers (3x IQR)
    - Negative values in positive-only columns (price, quantity)
    - Formatting irregularities
    """
    target_cols = [find_column_case_insensitive(df, c) for c in (columns or list(df.columns)) if find_column_case_insensitive(df, c)]
    num_cols = [c for c in target_cols if pd.api.types.is_numeric_dtype(df[c])]

    anomalous_indices = set()
    reasons: Dict[int, List[str]] = {}

    # 1. Negative numbers check
    for c in num_cols:
        if any(term in c.lower() for term in ["price", "qty", "quantity", "stock", "rate", "cost"]):
            neg_mask = df[c] < 0
            for idx in df[neg_mask].index:
                anomalous_indices.add(idx)
                reasons.setdefault(idx, []).append(f"Negative {c} ({df.loc[idx, c]})")

    # 2. Extreme Statistical Outliers (IQR)
    for c in num_cols:
        s = df[c].dropna()
        if len(s) >= 8:
            q25, q75 = s.quantile(0.25), s.quantile(0.75)
            iqr = q75 - q25
            if iqr > 0:
                low = q25 - (2.5 * iqr)
                high = q75 + (2.5 * iqr)
                out_mask = (df[c] < low) | (df[c] > high)
                for idx in df[out_mask].index:
                    anomalous_indices.add(idx)
                    reasons.setdefault(idx, []).append(f"Extreme {c} outlier ({df.loc[idx, c]})")

    result = df.loc[list(anomalous_indices)].copy()
    if not result.empty:
        result["_Anomaly_Reason"] = ["; ".join(reasons[i]) for i in result.index]
        msg = f"Detected {len(result)} suspicious / anomalous rows."
    else:
        result = pd.DataFrame(columns=list(df.columns) + ["_Anomaly_Reason"])
        msg = "No obvious anomalies or extreme outliers found."

    return result.reset_index(drop=True), msg


def execute_dry_run_preview(
    action_plan: Dict[str, Any],
    tables_store: Dict[str, pd.DataFrame]
) -> Tuple[pd.DataFrame, str, Dict[str, Any]]:
    """
    Simulates destructive actions (remove duplicates, delete rows, filter)
    and returns a preview DataFrame with visual flags for what will change.
    """
    action = action_plan.get("action", "").lower().strip()
    table_name = action_plan.get("table") or list(tables_store.keys())[0]
    df = tables_store[table_name]

    preview_df = df.copy()

    if action in ["clean", "remove_duplicates"]:
        op = action_plan.get("operation", "remove_duplicates")
        if op in ["remove_duplicates", "dedup"]:
            dup_mask = preview_df.duplicated(keep="first")
            preview_df["_DryRun_Action"] = np.where(dup_mask, "WILL_BE_DELETED", "KEPT")
            msg = f"Dry Run Preview: {dup_mask.sum()} duplicate rows will be deleted if confirmed."
            return preview_df, msg, {"rows_affected": int(dup_mask.sum())}

    elif action == "filter":
        col = action_plan.get("column", df.columns[0])
        op = action_plan.get("operator", "equals")
        val = action_plan.get("value")
        actual_col = find_column_case_insensitive(df, col)
        
        # Test mask
        mask = df[actual_col].astype(str).str.lower() == str(val).lower() if op == "equals" else (df[actual_col] > float(val))
        preview_df["_DryRun_Action"] = np.where(mask, "KEPT", "WILL_BE_REMOVED")
        msg = f"Dry Run Preview: {(~mask).sum()} rows will be removed if filter confirmed."
        return preview_df, msg, {"rows_affected": int((~mask).sum())}

    preview_df["_DryRun_Action"] = "NO_CHANGE"
    return preview_df, "Dry run preview generated.", {"rows_affected": 0}


def get_excel_learning_explanation(action_plan: Dict[str, Any]) -> Dict[str, str]:
    """
    Excel-Learning Mode ('Ye kaise kiya?'):
    Explains the manual Excel steps and equivalent formula in simple Hindi & English.
    """
    action = action_plan.get("action", "").lower()

    if action in ["vlookup", "xlookup"]:
        return {
            "en": (
                "How to do this in Excel:\n"
                "1. Click in cell C2 of your main sheet.\n"
                "2. Type formula: =XLOOKUP(B2, Prices!A:A, Prices!C:C, \"Not Found\")\n"
                "3. Press Enter and double click the bottom-right corner of the cell to fill down.\n"
                "Alternative (Legacy VLOOKUP): =VLOOKUP(B2, Prices!A:C, 3, FALSE)"
            ),
            "hi": (
                "एक्सेल में यह कैसे करें:\n"
                "1. अपनी मुख्य शीट के C2 सेल में जाएं।\n"
                "2. यह फॉर्मूला लिखें: =XLOOKUP(B2, Prices!A:A, Prices!C:C, \"Not Found\")\n"
                "3. Enter दबाएं और सेल के कोने पर डबल क्लिक करके पूरे कॉलम में लागू करें।\n"
                "पुराना तरीका: =VLOOKUP(B2, Prices!A:C, 3, FALSE)"
            )
        }

    elif action == "sort":
        return {
            "en": (
                "How to do this in Excel:\n"
                "1. Select your data table.\n"
                "2. Go to the 'Data' tab in the top ribbon.\n"
                "3. Click 'Sort' (or Sort A to Z / Z to A).\n"
                "4. Select the column name and sort order (Ascending or Descending)."
            ),
            "hi": (
                "एक्सेल में यह कैसे करें:\n"
                "1. अपने पूरे डेटा को सेलेक्ट करें।\n"
                "2. ऊपर रिबन में 'Data' टैब पर क्लिक करें।\n"
                "3. 'Sort' बटन पर क्लिक करें।\n"
                "4. जिस कॉलम से सॉर्ट करना है उसे चुनें और Ascending (A-Z) या Descending (Z-A) चुनें।"
            )
        }

    elif action == "pivot":
        return {
            "en": (
                "How to do this in Excel:\n"
                "1. Select your data table.\n"
                "2. Go to 'Insert' ➔ 'PivotTable'.\n"
                "3. Drag your category column (e.g. City) into the 'Rows' box.\n"
                "4. Drag your numeric column (e.g. Quantity) into the 'Values' box (set to Sum)."
            ),
            "hi": (
                "एक्सेल में यह कैसे करें:\n"
                "1. अपने डेटा को सेलेक्ट करें।\n"
                "2. ऊपर 'Insert' ➔ 'PivotTable' पर क्लिक करें।\n"
                "3. दाईं तरफ Rows बॉक्स में शहर (City) को खींचें।\n"
                "4. Values बॉक्स में मात्रा (Quantity) को खींचें और Sum चुनें।"
            )
        }

    elif action in ["clean", "remove_duplicates"]:
        return {
            "en": (
                "How to do this in Excel:\n"
                "1. Select the table.\n"
                "2. Go to 'Data' tab ➔ click 'Remove Duplicates'.\n"
                "3. Check the columns to check for duplicates and click OK."
            ),
            "hi": (
                "एक्सेल में यह कैसे करें:\n"
                "1. अपने डेटा को सेलेक्ट करें।\n"
                "2. 'Data' टैब में जाएं ➔ 'Remove Duplicates' पर क्लिक करें।\n"
                "3. जिन कॉलम्स में डुप्लिकेट चेक करना है उन्हें टिक करके OK दबाएं।"
            )
        }

    return {
        "en": "Excel manual steps available for all standard data operations.",
        "hi": "एक्सेल के सभी स्टेप्स आसानी से सीखे जा सकते हैं।"
    }


def dispatch_action_plan(
    action_plan: Dict[str, Any],
    tables_store: Dict[str, pd.DataFrame]
) -> Tuple[pd.DataFrame, str, Dict[str, Any]]:
    """
    Validates and dispatches a structured JSON action plan from LLM.
    Returns:
    - result_df: The updated/new DataFrame
    - status_message: Clear message in Hindi/English
    - metadata: {before_rows: int, after_rows: int, output_name: str}
    """
    action = action_plan.get("action", "").lower().strip()
    output_name = action_plan.get("output_name", "Result")

    if not action:
        raise ActionExecutionError("Action plan missing 'action' field.")

    # 1. VLOOKUP
    if action in ["vlookup", "xlookup"]:
        left_tbl = action_plan.get("left_table")
        right_tbl = action_plan.get("right_table")
        if left_tbl not in tables_store:
            raise ActionExecutionError(f"Table '{left_tbl}' not found. Available: {list(tables_store.keys())}")
        if right_tbl not in tables_store:
            raise ActionExecutionError(f"Table '{right_tbl}' not found. Available: {list(tables_store.keys())}")

        left_df = tables_store[left_tbl]
        right_df = tables_store[right_tbl]
        before_rows = len(left_df)

        res_df, msg = execute_vlookup(
            left_df=left_df,
            right_df=right_df,
            left_key=action_plan.get("left_key", ""),
            right_key=action_plan.get("right_key", ""),
            return_columns=action_plan.get("return_columns", []),
            fuzzy=action_plan.get("fuzzy", False),
            similarity_threshold=action_plan.get("similarity_threshold", 80.0)
        )
        return res_df, msg, {"before_rows": before_rows, "after_rows": len(res_df), "output_name": output_name}

    # 2. SORT
    elif action in ["sort", "sorting"]:
        target_tbl = action_plan.get("table") or list(tables_store.keys())[0]
        if target_tbl not in tables_store:
            raise ActionExecutionError(f"Table '{target_tbl}' not found.")
        df = tables_store[target_tbl]
        before_rows = len(df)
        cols = action_plan.get("columns", [])
        asc = action_plan.get("ascending", True)
        res_df, msg = execute_sort(df, columns=cols, ascending=asc)
        return res_df, msg, {"before_rows": before_rows, "after_rows": len(res_df), "output_name": output_name}

    # 3. COMMON / UNIQUE / SIMILAR
    elif action in ["find_common", "find_unique", "common_items"]:
        tbl_a = action_plan.get("table_a") or list(tables_store.keys())[0]
        tbl_b = action_plan.get("table_b") or (list(tables_store.keys())[1] if len(tables_store) > 1 else tbl_a)
        df_a = tables_store[tbl_a]
        df_b = tables_store[tbl_b]
        before_rows = len(df_a)
        res_df, msg = execute_find_similar_or_common(
            table_a=df_a,
            table_b=df_b,
            col_a=action_plan.get("col_a", ""),
            col_b=action_plan.get("col_b", ""),
            mode=action_plan.get("mode", "common"),
            fuzzy=action_plan.get("fuzzy", False),
            similarity_threshold=action_plan.get("similarity_threshold", 80.0)
        )
        return res_df, msg, {"before_rows": before_rows, "after_rows": len(res_df), "output_name": output_name}

    # 4. FILTER
    elif action in ["filter", "filtering"]:
        target_tbl = action_plan.get("table") or list(tables_store.keys())[0]
        df = tables_store[target_tbl]
        before_rows = len(df)
        res_df, msg = execute_filter(
            df=df,
            column=action_plan.get("column", ""),
            operator=action_plan.get("operator", "equals"),
            value=action_plan.get("value"),
            value2=action_plan.get("value2")
        )
        return res_df, msg, {"before_rows": before_rows, "after_rows": len(res_df), "output_name": output_name}

    # 5. CLEANING
    elif action in ["clean", "remove_duplicates", "drop_blanks", "fill_blanks", "trim_spaces", "change_case"]:
        target_tbl = action_plan.get("table") or list(tables_store.keys())[0]
        df = tables_store[target_tbl]
        before_rows = len(df)
        sub_op = action_plan.get("operation") or action
        res_df, msg = execute_cleaning(
            df=df,
            operation=sub_op,
            columns=action_plan.get("columns"),
            fill_value=action_plan.get("fill_value")
        )
        return res_df, msg, {"before_rows": before_rows, "after_rows": len(res_df), "output_name": output_name}

    # 6. PIVOT / GROUPBY
    elif action in ["pivot", "groupby", "pivot_table", "summary"]:
        target_tbl = action_plan.get("table") or list(tables_store.keys())[0]
        df = tables_store[target_tbl]
        before_rows = len(df)
        res_df, msg = execute_pivot_table(
            df=df,
            index_columns=action_plan.get("index_columns", []),
            values_columns=action_plan.get("values_columns", []),
            agg_func=action_plan.get("agg_func", "sum"),
            columns_pivot=action_plan.get("columns_pivot")
        )
        return res_df, msg, {"before_rows": before_rows, "after_rows": len(res_df), "output_name": output_name}

    # 7. COLUMN TRANSFORM (Calculated / Merge / Split / Rename / Delete)
    elif action in ["calculate", "merge_columns", "split_column", "rename_column", "delete_column"]:
        target_tbl = action_plan.get("table") or list(tables_store.keys())[0]
        df = tables_store[target_tbl]
        before_rows = len(df)
        res_df, msg = execute_column_transform(
            df=df,
            action_type=action,
            target_column=action_plan.get("target_column", ""),
            source_columns=action_plan.get("source_columns"),
            formula=action_plan.get("formula"),
            delimiter=action_plan.get("delimiter")
        )
        return res_df, msg, {"before_rows": before_rows, "after_rows": len(res_df), "output_name": output_name}

    # 8. MERGE SHEETS
    elif action in ["merge_sheets", "append_sheets", "combine_sheets"]:
        sheets = action_plan.get("sheets", list(tables_store.keys()))
        dfs = [tables_store[s] for s in sheets if s in tables_store]
        before_rows = sum(len(d) for d in dfs)
        res_df, msg = execute_merge_sheets(tables=dfs, sheet_names=sheets)
        return res_df, msg, {"before_rows": before_rows, "after_rows": len(res_df), "output_name": output_name}

    # 9. COMPARE VERSIONS
    elif action in ["compare_versions", "diff_tables"]:
        t_old = action_plan.get("table_old") or list(tables_store.keys())[0]
        t_new = action_plan.get("table_new") or (list(tables_store.keys())[1] if len(tables_store) > 1 else t_old)
        df_old = tables_store[t_old]
        df_new = tables_store[t_new]
        res_df, msg, meta = execute_compare_versions(df_old, df_new, action_plan.get("key_column"))
        return res_df, msg, {"before_rows": len(df_old), "after_rows": len(res_df), "output_name": output_name}

    # 10. ANOMALY DETECTION
    elif action in ["detect_anomalies", "find_anomalies", "outliers"]:
        target_tbl = action_plan.get("table") or list(tables_store.keys())[0]
        df = tables_store[target_tbl]
        res_df, msg = execute_detect_anomalies(df, action_plan.get("columns"))
        return res_df, msg, {"before_rows": len(df), "after_rows": len(res_df), "output_name": output_name}

    else:
        raise ActionExecutionError(
            f"Unknown action: '{action}'. Supported actions: vlookup, sort, find_common, filter, clean, pivot, calculate, merge_sheets, compare_versions, detect_anomalies.",
            f"Kripya sahi action chune: '{action}' samajh nahi aaya."
        )

