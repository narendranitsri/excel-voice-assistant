"""
reconcile.py
============
PART C: Combined Mode (Fill + Verify & Reconcile).
Compares the user's local Excel records against the live data fetched from the Website
and generates:
1. Updated Status DataFrame (Submitted -> Approved / Pending / Rejected / Objection)
2. A reconciliation summary with exact counts of Matched, Mismatched, and Missing records
3. A color-coded Excel workbook containing a "Differences" sheet where mismatched cells
   are highlighted in yellow/coral and matched cells in green.
"""

from __future__ import annotations
import io
from typing import Dict, List, Any, Optional, Tuple
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows


def reconcile_excel_with_portal_data(
    local_df: pd.DataFrame,
    portal_df: pd.DataFrame,
    key_column_local: str,
    key_column_portal: str,
    compare_columns: Optional[List[Tuple[str, str]]] = None
) -> Dict[str, Any]:
    """
    C2: Compares `local_df` (our Excel file) with `portal_df` (data fetched from the website)
    using `key_column_local` == `key_column_portal`.
    Returns:
    - summary: {total_local, matched_exact, mismatched_rows, missing_on_portal}
    - differences_df: DataFrame showing row-by-row comparison and exact field mismatches
    - cell_highlights: Map of (row_idx, col_name) -> 'mismatch' | 'missing' | 'match'
    """
    if local_df.empty:
        return {
            "summary": {"total_local": 0, "matched_exact": 0, "mismatched_rows": 0, "missing_on_portal": 0},
            "differences_df": pd.DataFrame(),
            "cell_highlights": {}
        }

    def norm_key(v: Any) -> str:
        s = str(v).strip()
        if s.endswith(".0"):
            s = s[:-2]
        return s.upper()

    portal_map: Dict[str, Dict[str, Any]] = {}
    if not portal_df.empty and key_column_portal in portal_df.columns:
        for _, r in portal_df.iterrows():
            k = norm_key(r[key_column_portal])
            if k:
                portal_map[k] = r.to_dict()

    # Auto-detect shared columns if compare_columns not specified
    if not compare_columns:
        compare_columns = []
        for l_col in local_df.columns:
            if l_col == key_column_local:
                continue
            for p_col in portal_df.columns:
                if p_col == key_column_portal:
                    continue
                if l_col.strip().lower().replace("_", "") == p_col.strip().lower().replace("_", ""):
                    compare_columns.append((l_col, p_col))
                    break

    diff_rows: List[Dict[str, Any]] = []
    cell_highlights: Dict[Tuple[int, str], str] = {}

    matched_exact = 0
    mismatched_rows = 0
    missing_on_portal = 0

    for idx, l_row in local_df.iterrows():
        k_val = norm_key(l_row.get(key_column_local, ""))
        rec: Dict[str, Any] = {
            key_column_local: k_val,
            "Reconcile_Status": "",
            "Portal_Application_Status": "",
            "Mismatch_Details": ""
        }

        if k_val not in portal_map:
            missing_on_portal += 1
            rec["Reconcile_Status"] = "MISSING ON PORTAL (पोर्टल पर नहीं मिला)"
            cell_highlights[(idx, "Reconcile_Status")] = "missing"
            for l_col, _ in compare_columns:
                rec[f"{l_col} (Excel)"] = str(l_row.get(l_col, ""))
                rec[f"{l_col} (Website)"] = "—"
            diff_rows.append(rec)
            continue

        p_row = portal_map[k_val]
        rec["Portal_Application_Status"] = str(
            p_row.get("Portal_Status") or p_row.get("Application_Status") or p_row.get("Status") or "Verified"
        )

        mismatches: List[str] = []
        for l_col, p_col in compare_columns:
            val_l = str(l_row.get(l_col, "")).strip()
            val_p = str(p_row.get(p_col, "")).strip()
            if val_l.endswith(".0"):
                val_l = val_l[:-2]
            if val_p.endswith(".0"):
                val_p = val_p[:-2]

            rec[f"{l_col} (Excel)"] = val_l
            rec[f"{l_col} (Website)"] = val_p

            if val_l.lower() != val_p.lower():
                mismatches.append(f"{l_col}: Excel='{val_l}' vs Web='{val_p}'")
                cell_highlights[(idx, f"{l_col} (Website)")] = "mismatch"
            else:
                cell_highlights[(idx, f"{l_col} (Website)")] = "match"

        if mismatches:
            mismatched_rows += 1
            rec["Reconcile_Status"] = "MISMATCH FOUND (अंतर मिला)"
            rec["Mismatch_Details"] = "; ".join(mismatches)
            cell_highlights[(idx, "Reconcile_Status")] = "mismatch"
        else:
            matched_exact += 1
            rec["Reconcile_Status"] = "EXACT MATCH (सही मिलान)"
            rec["Mismatch_Details"] = "All fields verified"
            cell_highlights[(idx, "Reconcile_Status")] = "match"

        diff_rows.append(rec)

    differences_df = pd.DataFrame(diff_rows)
    return {
        "summary": {
            "total_local": len(local_df),
            "matched_exact": matched_exact,
            "mismatched_rows": mismatched_rows,
            "missing_on_portal": missing_on_portal
        },
        "differences_df": differences_df,
        "cell_highlights": cell_highlights
    }


def export_reconciliation_excel(
    local_df: pd.DataFrame,
    portal_df: pd.DataFrame,
    reconcile_result: Dict[str, Any]
) -> bytes:
    """
    Exports a multi-sheet Excel workbook with:
    1. 'Differences' sheet (colour-coded: Red/Coral for missing, Yellow for mismatched cells, Green for verified)
    2. 'Original_Excel_Copy' sheet
    3. 'Fetched_Website_Data' sheet
    """
    wb = openpyxl.Workbook()
    ws_diff = wb.active
    ws_diff.title = "Differences"
    ws_diff.views.sheetView[0].showGridLines = True

    diff_df: pd.DataFrame = reconcile_result.get("differences_df", pd.DataFrame())
    highlights: Dict[Tuple[int, str], str] = reconcile_result.get("cell_highlights", {})

    header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_match = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    fill_mismatch = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
    fill_missing = PatternFill(start_color="FECDD3", end_color="FECDD3", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )

    if not diff_df.empty:
        cols = list(diff_df.columns)
        for r_idx, row in enumerate(dataframe_to_rows(diff_df, index=False, header=True), start=1):
            ws_diff.append(row)
            for c_idx in range(1, len(row) + 1):
                cell = ws_diff.cell(row=r_idx, column=c_idx)
                cell.border = thin_border
                if r_idx == 1:
                    cell.fill = header_fill
                    cell.font = header_font
                else:
                    data_row_idx = r_idx - 2
                    col_name = cols[c_idx - 1]
                    tag = highlights.get((data_row_idx, col_name))
                    if tag == "mismatch":
                        cell.fill = fill_mismatch
                    elif tag == "missing":
                        cell.fill = fill_missing
                    elif tag == "match":
                        cell.fill = fill_match

        ws_diff.freeze_panes = "A2"
        for col in ws_diff.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=12)
            ws_diff.column_dimensions[col[0].column_letter].width = min(max(max_len + 4, 14), 45)

    # Sheet 2: Original Copy
    ws_orig = wb.create_sheet("Original_Excel_Copy")
    for r in dataframe_to_rows(local_df, index=False, header=True):
        ws_orig.append(r)

    # Sheet 3: Website Data
    ws_web = wb.create_sheet("Fetched_Website_Data")
    for r in dataframe_to_rows(portal_df, index=False, header=True):
        ws_web.append(r)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()
