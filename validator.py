"""
validator.py
============
Pre-Flight Citizen & Portal Data Validator.
Validates rows before portal submission:
- Mandatory fields (non-blank check)
- Mobile number (10 digits, starts with 6-9)
- Aadhaar / JanAadhaar number (12 digits or 10 digits)
- PIN code (6 digits)
- Date format validation (DD/MM/YYYY or YYYY-MM-DD)
- Duplicate rows detection
Provides clear Hindi and English error messages per row.
"""

from __future__ import annotations
import re
from typing import Dict, List, Any, Optional, Tuple
import pandas as pd


class ValidationIssue:
    def __init__(
        self,
        row_index: int,
        column: str,
        issue_type: str,
        message_en: str,
        message_hi: str,
        current_value: Any
    ):
        self.row_index = row_index  # 1-based display row
        self.column = column
        self.issue_type = issue_type
        self.message_en = message_en
        self.message_hi = message_hi
        self.current_value = current_value

    def to_dict(self) -> Dict[str, Any]:
        return {
            "row": self.row_index,
            "column": self.column,
            "type": self.issue_type,
            "value": str(self.current_value),
            "message_en": self.message_en,
            "message_hi": self.message_hi
        }


def detect_column_role(col_name: str) -> str:
    """Infers semantic validation role from column header (English, Hindi, Hinglish)."""
    c = col_name.strip().lower()
    if any(k in c for k in ["mobile", "phone", "contact", "मोबाइल", "फ़ोन", "फोन"]):
        return "mobile"
    if any(k in c for k in ["aadhaar", "aadhar", "uid", "आधार"]):
        return "aadhaar"
    if any(k in c for k in ["pincode", "pin code", "pin", "पिन", "postal"]):
        return "pincode"
    if any(k in c for k in ["dob", "birth", "date", "तिथि", "जन्म"]):
        return "date"
    if any(k in c for k in ["name", "applicant", "नाम", "आवेदक"]):
        return "mandatory_name"
    return "general"


def validate_dataframe(
    df: pd.DataFrame,
    mandatory_columns: Optional[List[str]] = None
) -> Dict[str, Any]:
    """
    Runs full validation across all rows and returns:
    - is_valid: bool
    - total_issues: int
    - issues: List of dicts with Hindi & English messages
    - duplicate_rows: List of 1-based row numbers
    - invalid_cells_map: Dict["row_col", issue_dict] for red UI highlighting
    """
    issues: List[ValidationIssue] = []
    invalid_cells_map: Dict[str, Dict[str, Any]] = {}

    if df.empty:
        return {
            "is_valid": True,
            "total_issues": 0,
            "issues": [],
            "duplicate_rows": [],
            "invalid_cells_map": {}
        }

    # Infer mandatory columns if not explicitly provided
    inferred_mandatory = mandatory_columns or [
        col for col in df.columns if detect_column_role(col) in ["mandatory_name", "mobile", "aadhaar"]
    ]

    # 1. Check Duplicate Rows
    dup_mask = df.duplicated(keep="first")
    duplicate_rows: List[int] = []
    for idx, is_dup in enumerate(dup_mask):
        if is_dup:
            row_num = idx + 1
            duplicate_rows.append(row_num)
            first_col = df.columns[0]
            issue = ValidationIssue(
                row_index=row_num,
                column=first_col,
                issue_type="duplicate",
                message_en=f"Row {row_num}: Duplicate row detected.",
                message_hi=f"Row {row_num}: यह पंक्ति डुप्लिकेट (दोहराई गई) है।",
                current_value=df.iloc[idx][first_col]
            )
            issues.append(issue)
            invalid_cells_map[f"{idx}_{first_col}"] = issue.to_dict()

    # 2. Row-by-Row Field Validation
    for idx in range(len(df)):
        row_num = idx + 1
        row_data = df.iloc[idx]

        for col in df.columns:
            raw_val = row_data[col]
            val_str = str(raw_val).strip() if pd.notna(raw_val) else ""
            if val_str.lower() in ["nan", "none", "null"]:
                val_str = ""

            # Strip trailing .0 from excel float strings
            if re.match(r"^\d+\.0$", val_str):
                val_str = val_str[:-2]

            role = detect_column_role(col)

            # Check Mandatory Blank
            if col in inferred_mandatory and not val_str:
                issue = ValidationIssue(
                    row_index=row_num,
                    column=col,
                    issue_type="blank_mandatory",
                    message_en=f"Row {row_num}: Mandatory column '{col}' is blank.",
                    message_hi=f"Row {row_num}: ज़रूरी कॉलम '{col}' खाली है।",
                    current_value=""
                )
                issues.append(issue)
                invalid_cells_map[f"{idx}_{col}"] = issue.to_dict()
                continue

            if not val_str:
                continue

            # Mobile Validation (10 digits)
            if role == "mobile":
                digits = re.sub(r"\D", "", val_str)
                if digits.startswith("91") and len(digits) == 12:
                    digits = digits[2:]
                if len(digits) != 10 or digits[0] not in "6789":
                    issue = ValidationIssue(
                        row_index=row_num,
                        column=col,
                        issue_type="invalid_mobile",
                        message_en=f"Row {row_num}: Mobile number has {len(digits)} digits ('{val_str}'), must be a valid 10-digit number.",
                        message_hi=f"Row {row_num}: मोबाइल नंबर {len(digits)} अंक का है ('{val_str}'), 10 अंक का होना चाहिए।",
                        current_value=val_str
                    )
                    issues.append(issue)
                    invalid_cells_map[f"{idx}_{col}"] = issue.to_dict()

            # Aadhaar Validation (12 digits)
            elif role == "aadhaar":
                digits = re.sub(r"\D", "", val_str)
                if len(digits) != 12:
                    issue = ValidationIssue(
                        row_index=row_num,
                        column=col,
                        issue_type="invalid_aadhaar",
                        message_en=f"Row {row_num}: Aadhaar number has {len(digits)} digits, must be 12 digits.",
                        message_hi=f"Row {row_num}: आधार नंबर {len(digits)} अंक का है, पूरे 12 अंक होना चाहिए।",
                        current_value=val_str
                    )
                    issues.append(issue)
                    invalid_cells_map[f"{idx}_{col}"] = issue.to_dict()

            # PIN Code Validation (6 digits)
            elif role == "pincode":
                digits = re.sub(r"\D", "", val_str)
                if len(digits) != 6 or digits[0] == "0":
                    issue = ValidationIssue(
                        row_index=row_num,
                        column=col,
                        issue_type="invalid_pincode",
                        message_en=f"Row {row_num}: PIN code '{val_str}' is invalid (must be 6 digits).",
                        message_hi=f"Row {row_num}: पिन कोड '{val_str}' गलत है (6 अंक का होना चाहिए)।",
                        current_value=val_str
                    )
                    issues.append(issue)
                    invalid_cells_map[f"{idx}_{col}"] = issue.to_dict()

            # Date Validation
            elif role == "date":
                dt = pd.to_datetime(val_str, dayfirst=True, errors="coerce")
                if pd.isna(dt):
                    issue = ValidationIssue(
                        row_index=row_num,
                        column=col,
                        issue_type="invalid_date",
                        message_en=f"Row {row_num}: Date '{val_str}' is not in a valid DD/MM/YYYY format.",
                        message_hi=f"Row {row_num}: तारीख '{val_str}' सही फॉर्मेट (DD/MM/YYYY) में नहीं है।",
                        current_value=val_str
                    )
                    issues.append(issue)
                    invalid_cells_map[f"{idx}_{col}"] = issue.to_dict()

    return {
        "is_valid": len(issues) == 0,
        "total_issues": len(issues),
        "issues": [i.to_dict() for i in issues],
        "duplicate_rows": duplicate_rows,
        "invalid_cells_map": invalid_cells_map
    }
