"""
quality.py
==========
One-Click Data Quality Audit and Smart Standardization Engine.
Detects data health issues and provides safe automatic fixes with preview.
"""

from __future__ import annotations
import re
from typing import Dict, List, Any, Tuple, Optional
import pandas as pd
import numpy as np

try:
    from rapidfuzz import process, fuzz
except ImportError:
    process = None
    fuzz = None


# Known city normalization dictionary for common Indian variants
KNOWN_CITY_CANONICAL = [
    "Mumbai", "Delhi", "Bengaluru", "Kolkata", "Chennai", 
    "Hyderabad", "Pune", "Ahmedabad", "Jaipur", "Lucknow",
    "Chandigarh", "Indore", "Gurugram", "Noida", "Bhopal"
]


def audit_data_quality(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Performs a comprehensive, one-click health audit on a DataFrame.
    Returns structured stats on blanks, duplicates, mixed types, spaces, outliers, and bad formats.
    """
    total_rows = len(df)
    total_cells = df.size if total_rows > 0 else 0

    # 1. Duplicate rows
    duplicate_rows_count = int(df.duplicated().sum())

    # 2. Blank cells per column
    blanks_by_column: Dict[str, int] = {}
    for col in df.columns:
        cnt = int(df[col].isna().sum() + (df[col].astype(str).str.strip() == '').sum() - df[col].isna().sum())
        if cnt > 0:
            blanks_by_column[col] = cnt

    total_blanks = sum(blanks_by_column.values()) + int(df.isna().sum().sum())

    # 3. Leading / trailing extra spaces
    columns_with_extra_spaces: List[str] = []
    space_count = 0
    for col in df.select_dtypes(include=["object"]).columns:
        has_spaces = df[col].dropna().apply(lambda x: isinstance(x, str) and (x != x.strip()))
        cnt = int(has_spaces.sum())
        if cnt > 0:
            columns_with_extra_spaces.append(f"{col} ({cnt} rows)")
            space_count += cnt

    # 4. Mixed types in column (e.g. strings mixed with numbers)
    mixed_type_columns: List[str] = []
    for col in df.columns:
        non_nulls = df[col].dropna()
        if len(non_nulls) > 0:
            types_found = set(type(v).__name__ for v in non_nulls.head(100))
            # If float and int are mixed, it's normal numeric, but if str and int/float mix:
            if 'str' in types_found and ('int' in types_found or 'float' in types_found):
                mixed_type_columns.append(col)

    # 5. Invalid email addresses
    invalid_emails: Dict[str, int] = {}
    email_regex = re.compile(r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$')
    for col in df.columns:
        if 'email' in col.lower():
            invalids = df[col].dropna().apply(lambda x: not bool(email_regex.match(str(x).strip())))
            cnt = int(invalids.sum())
            if cnt > 0:
                invalid_emails[col] = cnt

    # 6. Invalid phone numbers (non-10 digits or malformed)
    invalid_phones: Dict[str, int] = {}
    for col in df.columns:
        if any(term in col.lower() for term in ['phone', 'mobile', 'contact', 'tel']):
            def is_bad_phone(val: Any) -> bool:
                digits = re.sub(r'\D', '', str(val))
                return len(digits) not in [10, 11, 12]  # Indian mobile standard
            invalids = df[col].dropna().apply(is_bad_phone)
            cnt = int(invalids.sum())
            if cnt > 0:
                invalid_phones[col] = cnt

    # 7. Statistical Outliers via IQR
    outliers_detected: Dict[str, int] = {}
    for col in df.select_dtypes(include=[np.number]).columns:
        series = df[col].dropna()
        if len(series) >= 8:
            q25, q75 = series.quantile(0.25), series.quantile(0.75)
            iqr = q75 - q25
            if iqr > 0:
                lower_bound = q25 - (1.5 * iqr)
                upper_bound = q75 + (1.5 * iqr)
                outlier_cnt = int(((series < lower_bound) | (series > upper_bound)).sum())
                if outlier_cnt > 0:
                    outliers_detected[col] = outlier_cnt

    # Overall Health Score (0 - 100)
    penalty = (
        (duplicate_rows_count * 5) +
        (total_blanks * 1.5) +
        (space_count * 1) +
        (len(mixed_type_columns) * 10) +
        (sum(invalid_emails.values()) * 5) +
        (sum(invalid_phones.values()) * 5)
    )
    health_score = max(10, min(100, int(100 - (penalty / max(1, total_rows) * 100))))

    return {
        "health_score": health_score,
        "total_rows": total_rows,
        "duplicate_rows": duplicate_rows_count,
        "total_blanks": total_blanks,
        "blanks_by_column": blanks_by_column,
        "extra_spaces_columns": columns_with_extra_spaces,
        "mixed_type_columns": mixed_type_columns,
        "invalid_emails": invalid_emails,
        "invalid_phones": invalid_phones,
        "outliers_detected": outliers_detected
    }


def auto_fix_data_quality(
    df: pd.DataFrame,
    options: Optional[Dict[str, bool]] = None
) -> Tuple[pd.DataFrame, List[str]]:
    """
    Applies automatic, standard data cleaning:
    - Trims spaces from all strings
    - Drops duplicate rows
    - Standardizes names to Title Case
    - Standardizes phone numbers to +91 format
    - Standardizes dates to DD-MM-YYYY format
    - Normalizes known city spelling variants via fuzzy matching
    """
    opts = options or {
        "trim_spaces": True,
        "remove_duplicates": True,
        "standardize_case": True,
        "standardize_phones": True,
        "standardize_dates": True,
        "fuzzy_cities": True
    }

    result = df.copy()
    logs: List[str] = []

    # 1. Trim whitespace
    if opts.get("trim_spaces"):
        for col in result.select_dtypes(include=["object"]).columns:
            result[col] = result[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
        logs.append("Trimmed whitespace across all text columns.")

    # 2. Remove duplicates
    if opts.get("remove_duplicates"):
        init_len = len(result)
        result = result.drop_duplicates().reset_index(drop=True)
        removed = init_len - len(result)
        if removed > 0:
            logs.append(f"Removed {removed} duplicate rows.")

    # 3. Standardize text case (Title Case for names/cities)
    if opts.get("standardize_case"):
        for col in result.columns:
            if any(term in col.lower() for term in ["name", "customer", "city", "product", "item"]):
                result[col] = result[col].apply(
                    lambda x: x.title() if isinstance(x, str) and not x.isupper() else x
                )
        logs.append("Standardized names and categories to Proper Case.")

    # 4. Standardize phone numbers (+91 XXXXX XXXXX)
    if opts.get("standardize_phones"):
        for col in result.columns:
            if any(term in col.lower() for term in ["phone", "mobile", "contact"]):
                def format_in_phone(val: Any) -> Any:
                    if pd.isna(val) or val is None:
                        return val
                    digits = re.sub(r'\D', '', str(val))
                    if len(digits) == 10:
                        return f"+91 {digits[:5]} {digits[5:]}"
                    elif len(digits) == 12 and digits.startswith("91"):
                        return f"+91 {digits[2:7]} {digits[7:]}"
                    return val
                result[col] = result[col].apply(format_in_phone)
        logs.append("Standardized mobile numbers to +91 Indian format.")

    # 5. Standardize dates to DD-MM-YYYY
    if opts.get("standardize_dates"):
        for col in result.columns:
            if any(term in col.lower() for term in ["date", "time", "dob"]):
                try:
                    parsed_dates = pd.to_datetime(result[col], errors="coerce", dayfirst=True)
                    # If majority are valid dates, format them
                    if parsed_dates.notna().sum() > len(result) * 0.5:
                        result[col] = parsed_dates.dt.strftime("%d-%m-%Y")
                        logs.append(f"Standardized column '{col}' to DD-MM-YYYY format.")
                except Exception:
                    pass

    # 6. Fuzzy city standardisation (e.g. 'Dilli', 'delhi ' -> 'Delhi')
    if opts.get("fuzzy_cities") and process and fuzz:
        for col in result.columns:
            if "city" in col.lower():
                def normalize_city(val: Any) -> Any:
                    if not isinstance(val, str) or not val.strip():
                        return val
                    clean = val.strip().title()
                    # Check against known canonical cities
                    match = process.extractOne(clean, KNOWN_CITY_CANONICAL, scorer=fuzz.token_sort_ratio)
                    if match and match[1] >= 80:
                        return match[0]
                    return clean
                result[col] = result[col].apply(normalize_city)
        logs.append("Normalized city spelling variants with fuzzy matching.")

    return result, logs
