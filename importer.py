"""
importer.py
===========
Multi-format Data Importer for Data-to-Portal Auto Filler.
Reads data from:
1. Excel (.xlsx, .xls) - supports multiple sheets
2. CSV (.csv) - handles UTF-8 and Hindi Unicode text
3. Word (.docx) - extracts structured tables AND "Label: Value" paragraphs into a clean DataFrame
"""

from __future__ import annotations
import io
import os
import re
from typing import Dict, List, Optional, Tuple, Any
import pandas as pd

try:
    import docx
except ImportError:
    docx = None


def load_excel_or_csv(file_obj: Any, filename: str) -> Dict[str, pd.DataFrame]:
    """
    Reads an uploaded Excel or CSV file and returns a dictionary of {sheet_name: DataFrame}.
    Preserves string formatting for leading-zero PIN codes, mobile numbers, and Aadhaar numbers.
    """
    sheets_dict: Dict[str, pd.DataFrame] = {}
    ext = os.path.splitext(filename)[1].lower()

    if ext == ".csv":
        df = pd.read_csv(file_obj, dtype=str, keep_default_na=False)
        clean_name = os.path.splitext(os.path.basename(filename))[0]
        sheets_dict[clean_name] = df
    else:
        xls = pd.ExcelFile(file_obj)
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str, keep_default_na=False)
            sheets_dict[sheet_name] = df

    return sheets_dict


def load_word_document(file_obj: Any, filename: str = "document.docx") -> pd.DataFrame:
    """
    Extracts structured records from a Word (.docx) file.
    Supports two common government/citizen document formats:
    1. Word Tables: First row is treated as headers, subsequent rows as records.
       (Also handles 2-column Key-Value tables!)
    2. Paragraph "Label: Value" or "Label - Value" blocks:
       e.g.:
       Name: Ramesh Kumar
       Father Name: Suresh Kumar
       Mobile: 9876543210
    """
    if docx is None:
        raise RuntimeError("python-docx is not installed. Run 'pip install python-docx'.")

    doc = docx.Document(file_obj)
    extracted_records: List[Dict[str, str]] = []

    # 1. Extract from Word Tables
    for table in doc.tables:
        if not table.rows:
            continue

        # Check if it's a 2-column Key-Value table vs standard header table
        first_row_cells = [c.text.strip() for c in table.rows[0].cells]
        if len(first_row_cells) == 2 and any(
            k in first_row_cells[0].lower() for k in ["name", "नाम", "field", "label", "आवेदक"]
        ) and not any(k in first_row_cells[1].lower() for k in ["father", "पिता", "mobile", "मोबाइल"]):
            # Treat as key-value pairs for a single record
            kv_record: Dict[str, str] = {}
            for row in table.rows:
                cells = [c.text.strip() for c in row.cells]
                if len(cells) >= 2 and cells[0]:
                    key = cells[0].rstrip(":-").strip()
                    kv_record[key] = cells[1]
            if kv_record:
                extracted_records.append(kv_record)
        else:
            # Standard table: row 0 = headers, row 1..N = data rows
            headers = [c.text.strip() or f"Column_{i+1}" for i, c in enumerate(table.rows[0].cells)]
            for row in table.rows[1:]:
                vals = [c.text.strip() for c in row.cells]
                if any(vals):
                    rec = {headers[i]: (vals[i] if i < len(vals) else "") for i in range(len(headers))}
                    extracted_records.append(rec)

    # 2. Extract from "Label: Value" Paragraphs
    current_para_record: Dict[str, str] = {}
    kv_pattern = re.compile(r"^([^:\-—]{2,40})\s*[:\-—]\s*(.+)$")

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            # Blank line separates applicant blocks if we collected enough fields
            if len(current_para_record) >= 3:
                extracted_records.append(current_para_record)
                current_para_record = {}
            continue

        # Check for "---" separator
        if text.startswith("---") or text.startswith("==="):
            if current_para_record:
                extracted_records.append(current_para_record)
                current_para_record = {}
            continue

        match = kv_pattern.match(text)
        if match:
            key = match.group(1).strip()
            val = match.group(2).strip()
            # If the same key repeats (e.g. 'Name' or 'नाम'), start a new record
            if key in current_para_record and len(current_para_record) >= 2:
                extracted_records.append(current_para_record)
                current_para_record = {}
            current_para_record[key] = val

    if current_para_record:
        extracted_records.append(current_para_record)

    if not extracted_records:
        return pd.DataFrame()

    return pd.DataFrame(extracted_records).fillna("")


def filter_rows_by_selection(
    df: pd.DataFrame,
    mode: str = "all",
    start_row: int = 1,
    end_row: Optional[int] = None,
    selected_indices: Optional[List[int]] = None
) -> pd.DataFrame:
    """
    Filters DataFrame rows based on user selection:
    - 'all': all rows
    - 'range': 1-indexed inclusive row range [start_row, end_row]
    - 'selected': specific 0-indexed row indices
    """
    if df.empty:
        return df

    if mode == "range":
        s_idx = max(0, start_row - 1)
        e_idx = end_row if end_row is not None else len(df)
        return df.iloc[s_idx:e_idx].copy()

    elif mode == "selected" and selected_indices is not None:
        valid_indices = [i for i in selected_indices if 0 <= i < len(df)]
        return df.iloc[valid_indices].copy()

    return df.copy()
