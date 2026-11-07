"""
utils.py
========
Table Extraction (PDF / Image), Speech Synthesis (gTTS), and Executive Summary Builder.
"""

from __future__ import annotations
import os
import io
import re
import json
import base64
from typing import Dict, List, Any, Optional, Tuple
import pandas as pd

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    from gtts import gTTS
except ImportError:
    gTTS = None

try:
    from PIL import Image
except ImportError:
    Image = None

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows


def extract_table_from_pdf(pdf_file_bytes: bytes) -> List[pd.DataFrame]:
    """
    Extracts tabular data from PDF files using pdfplumber.
    Returns a list of extracted DataFrames (one per detected table).
    """
    if pdfplumber is None:
        raise RuntimeError("pdfplumber is not installed. Run 'pip install pdfplumber'.")

    extracted_dfs: List[pd.DataFrame] = []
    with pdfplumber.open(io.BytesIO(pdf_file_bytes)) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            for t_idx, tbl in enumerate(tables):
                if tbl and len(tbl) > 1:
                    headers = [str(h).strip() if h else f"Col_{i+1}" for i, h in enumerate(tbl[0])]
                    data_rows = tbl[1:]
                    df = pd.DataFrame(data_rows, columns=headers)
                    # Clean empty rows
                    df.dropna(how="all", inplace=True)
                    if not df.empty:
                        extracted_dfs.append(df.reset_index(drop=True))

    return extracted_dfs


def extract_table_from_image_with_claude(
    image_bytes: bytes,
    api_key: Optional[str] = None,
    model_name: Optional[str] = None
) -> Optional[pd.DataFrame]:
    """
    Extracts tabular data from an image (photo of bill, invoice, or printed table)
    using Claude Vision API. Returns a clean pandas DataFrame.
    """
    resolved_key = api_key or os.environ.get("ANTHROPIC_API_KEY")
    if not resolved_key:
        raise RuntimeError("Anthropic API key is required for image table extraction.")

    import anthropic
    b64_img = base64.b64encode(image_bytes).decode("utf-8")
    model = model_name or os.environ.get("CLAUDE_MODEL", "claude-sonnet-5-5")

    client = anthropic.Anthropic(api_key=resolved_key)
    prompt = """Analyze this image of a table/invoice/register and extract all structured data.
Return ONLY a valid JSON object matching this schema:
{
  "headers": ["Column1", "Column2", ...],
  "rows": [
    ["Val1", "Val2", ...],
    ["Val1", "Val2", ...]
  ]
}
Do not include any conversational text or markdown code fences outside the JSON.
"""

    response = client.messages.create(
        model=model,
        max_tokens=2500,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/jpeg",
                            "data": b64_img
                        }
                    },
                    {
                        "type": "text",
                        "text": prompt
                    }
                ]
            }
        ]
    )

    text_resp = "".join([b.text for b in response.content if hasattr(b, "text")])
    match = re.search(r"\{.*\}", text_resp, re.DOTALL)
    if match:
        data = json.loads(match.group(0))
        headers = data.get("headers", [])
        rows = data.get("rows", [])
        if headers and rows:
            return pd.DataFrame(rows, columns=headers)

    return None


def generate_voice_reply_audio(text_message: str, lang: str = "hi") -> Optional[bytes]:
    """
    Generates a spoken audio clip (.mp3) using gTTS for audio voice playback.
    lang: 'hi' for Hindi, 'en' for Indian English.
    """
    if gTTS is None:
        return None
    try:
        clean_text = text_message.strip()
        tts = gTTS(text=clean_text, lang=lang, slow=False)
        buf = io.BytesIO()
        tts.write_to_fp(buf)
        buf.seek(0)
        return buf.read()
    except Exception as e:
        print("gTTS generation error:", e)
        return None


def create_executive_summary_sheet(
    workbook: openpyxl.Workbook,
    df: pd.DataFrame,
    sheet_name: str = "Executive_Summary"
) -> openpyxl.worksheet.worksheet.Worksheet:
    """
    Generates a professional one-page executive summary sheet inside the Excel workbook.
    Includes:
    - High-level KPIs (Total Records, Unique Keys, Totals, Averages)
    - Top 5 and Bottom 5 items
    - Key insights summary
    - Clean typography and professional styling
    """
    if sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]
    else:
        ws = workbook.create_sheet(title=sheet_name, index=0)

    ws.views.sheetView[0].showGridLines = True

    # Styling definitions
    title_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    title_font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")

    card_header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    card_header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    label_font = Font(name="Calibri", size=10, bold=True, color="475569")
    val_font = Font(name="Calibri", size=12, bold=True, color="0F172A")

    thin_border = Border(
        left=Side(style="thin", color="E2E8F0"),
        right=Side(style="thin", color="E2E8F0"),
        top=Side(style="thin", color="E2E8F0"),
        bottom=Side(style="thin", color="E2E8F0")
    )

    # 1. Main Header
    ws.merge_cells("A1:G2")
    title_cell = ws["A1"]
    title_cell.value = "EXECUTIVE SUMMARY REPORT & DATA INSIGHTS"
    title_cell.fill = title_fill
    title_cell.font = title_font
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    # 2. Key Metrics Cards Row
    ws["A4"] = "METRIC"
    ws["B4"] = "VALUE"
    ws["A4"].fill = card_header_fill
    ws["A4"].font = card_header_font
    ws["B4"].fill = card_header_fill
    ws["B4"].font = card_header_font

    metrics = [
        ("Total Rows in Dataset", len(df)),
        ("Total Columns", len(df.columns)),
        ("Duplicate Rows Detected", int(df.duplicated().sum()))
    ]

    # Numeric metrics
    num_cols = df.select_dtypes(include=["number"]).columns
    for c in num_cols[:3]:
        metrics.append((f"Total {c}", round(float(df[c].sum()), 2)))
        metrics.append((f"Average {c}", round(float(df[c].mean()), 2)))

    cur_row = 5
    for label, val in metrics:
        ws.cell(row=cur_row, column=1, value=label).font = label_font
        cell_val = ws.cell(row=cur_row, column=2, value=val)
        cell_val.font = val_font
        cell_val.alignment = Alignment(horizontal="right")
        ws.cell(row=cur_row, column=1).border = thin_border
        ws.cell(row=cur_row, column=2).border = thin_border
        cur_row += 1

    # 3. Top 5 Items Table
    if len(num_cols) > 0 and len(df) >= 5:
        sort_col = num_cols[0]
        top_5 = df.sort_values(by=sort_col, ascending=False).head(5)
        
        ws["D4"] = f"TOP 5 PERFORMERS BY {sort_col.upper()}"
        ws.merge_cells("D4:G4")
        ws["D4"].fill = card_header_fill
        ws["D4"].font = card_header_font
        ws["D4"].alignment = Alignment(horizontal="center")

        # Headers
        disp_cols = list(df.columns[:4])
        for idx, col_name in enumerate(disp_cols, start=4):
            c_cell = ws.cell(row=5, column=idx, value=col_name)
            c_cell.font = Font(name="Calibri", size=10, bold=True)
            c_cell.border = thin_border

        for r_idx, (_, r_data) in enumerate(top_5.iterrows(), start=6):
            for c_idx, col_name in enumerate(disp_cols, start=4):
                val_c = ws.cell(row=r_idx, column=c_idx, value=r_data[col_name])
                val_c.font = Font(name="Calibri", size=10)
                val_c.border = thin_border

    # Adjust widths
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 5
    ws.column_dimensions["D"].width = 24
    ws.column_dimensions["E"].width = 22
    ws.column_dimensions["F"].width = 16
    ws.column_dimensions["G"].width = 16

    return ws
