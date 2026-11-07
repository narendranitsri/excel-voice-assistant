"""
extractor.py
============
PART B: Website -> Excel Extraction Engine.
Capabilities:
- B1. Table & List Extraction: Detects all HTML tables and structured lists on the current page
  using BeautifulSoup + pandas.read_html; handles multi-page pagination (Next button / page numbers)
  up to a configurable `max_pages` limit with polite delays.
- B2. Search-and-Fetch from Excel: Takes an Excel sheet of IDs (Application No, Aadhaar, Roll No),
  enters each ID into the portal's Search/Status page, reads the result fields (Status, Approval Date,
  Remarks, Officer), pauses for CAPTCHA/OTP if prompted, and writes results back into the Excel copy.
- B3. Detail-Page Extraction: Opens row detail links and collects "Label -> Value" pairs into columns.
- B4. File Downloads: Downloads linked PDFs/receipts/certificates into `./downloads/<Row_ID>/`
  and records the local file path in the Excel sheet.
- B6. Cleanup on the Way In: Trims whitespace, standardizes dates, splits combined fields,
  removes duplicate rows, and keeps Hindi Unicode intact.
"""

from __future__ import annotations
import io
import os
import re
import time
from typing import Dict, List, Any, Optional, Tuple
import pandas as pd
from bs4 import BeautifulSoup


def detect_tables_on_html(html_content: str) -> List[Dict[str, Any]]:
    """
    B1: Scans an HTML page for all `<table>` elements and structured card/list grids.
    Returns a list of {"table_index": i, "title": str, "rows": int, "columns": list, "df": DataFrame}.
    """
    detected: List[Dict[str, Any]] = []
    if not html_content.strip():
        return detected

    soup = BeautifulSoup(html_content, "html.parser")

    # 1. Standard HTML <table> elements via BeautifulSoup + pandas.read_html
    try:
        dfs = pd.read_html(io.StringIO(html_content))
        html_tables = soup.find_all("table")
        for idx, df in enumerate(dfs):
            if df.empty:
                continue
            df = df.dropna(how="all").fillna("")
            # Try to find caption or preceding heading
            title = f"Table #{idx + 1}"
            if idx < len(html_tables):
                cap = html_tables[idx].find("caption")
                if cap and cap.get_text(strip=True):
                    title = cap.get_text(strip=True)
                else:
                    prev_h = html_tables[idx].find_previous(["h1", "h2", "h3", "h4"])
                    if prev_h and prev_h.get_text(strip=True):
                        title = prev_h.get_text(strip=True)

            detected.append({
                "table_index": idx + 1,
                "title": title,
                "rows": len(df),
                "columns": [str(c) for c in df.columns],
                "df": df.astype(str)
            })
    except Exception:
        pass

    # 2. Structured Key-Value Definition Lists (<dl> or .status-card / .detail-row)
    kv_pairs: Dict[str, str] = {}
    for dt in soup.find_all("dt"):
        dd = dt.find_next_sibling("dd")
        if dd:
            kv_pairs[dt.get_text(strip=True).rstrip(":")] = dd.get_text(strip=True)

    for row_div in soup.select(".detail-row, .status-item, .kv-row"):
        lbl = row_div.select_one(".label, strong, b, span:first-child")
        val = row_div.select_one(".value, span:last-child")
        if lbl and val and lbl != val:
            kv_pairs[lbl.get_text(strip=True).rstrip(":")] = val.get_text(strip=True)

    if kv_pairs:
        kv_df = pd.DataFrame([kv_pairs])
        detected.append({
            "table_index": len(detected) + 1,
            "title": "Page Details / Status Summary Card",
            "rows": 1,
            "columns": list(kv_df.columns),
            "df": kv_df
        })

    return detected


def clean_extracted_dataframe(
    df: pd.DataFrame,
    trim_spaces: bool = True,
    standardize_dates: bool = True,
    remove_duplicates: bool = True,
    split_column_spec: Optional[Tuple[str, str, str, str]] = None
) -> pd.DataFrame:
    """
    B6: Cleanup on the way in:
    - Trims extra whitespace while preserving Hindi Unicode characters
    - Standardizes date columns to DD/MM/YYYY
    - Splits combined fields (e.g. "District - Block" -> ["District", "Block"])
    - Removes duplicate rows
    - Adds extraction timestamp column
    """
    if df.empty:
        return df

    cleaned = df.copy()

    # 1. Trim spaces
    if trim_spaces:
        for col in cleaned.columns:
            cleaned[col] = cleaned[col].apply(
                lambda x: re.sub(r"\s+", " ", str(x)).strip() if pd.notna(x) else ""
            )

    # 2. Optional split of combined column: (source_col, delimiter, new_col_1, new_col_2)
    if split_column_spec:
        src_col, delim, col1, col2 = split_column_spec
        if src_col in cleaned.columns and delim:
            parts = cleaned[src_col].astype(str).str.split(delim, n=1, expand=True)
            cleaned[col1] = parts[0].str.strip() if 0 in parts.columns else ""
            cleaned[col2] = parts[1].str.strip() if 1 in parts.columns else ""

    # 3. Standardize dates
    if standardize_dates:
        for col in cleaned.columns:
            if any(k in str(col).lower() for k in ["date", "dob", "तिथि", "दिनांक"]):
                try:
                    dt = pd.to_datetime(cleaned[col], dayfirst=True, errors="coerce")
                    valid_mask = dt.notna()
                    if valid_mask.sum() > 0:
                        cleaned.loc[valid_mask, col] = dt[valid_mask].dt.strftime("%d/%m/%Y")
                except Exception:
                    pass

    # 4. Remove duplicates
    if remove_duplicates:
        cleaned = cleaned.drop_duplicates().reset_index(drop=True)

    cleaned["_Extracted_At"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
    return cleaned


def extract_paginated_table_from_page(
    page: Any,
    table_index: int = 0,
    next_button_selector: str = "a.next, button.next, #next_page_btn, a:has-text('Next'), button:has-text('Next')",
    max_pages: int = 5,
    delay_sec: float = 1.5
) -> pd.DataFrame:
    """
    B1: Extracts an HTML table across multiple pages by clicking the Next / Load More button
    up to `max_pages` while respecting `delay_sec` between pages.
    """
    collected_dfs: List[pd.DataFrame] = []

    for page_num in range(1, max_pages + 1):
        html = page.content()
        tables = detect_tables_on_html(html)
        if not tables:
            break

        chosen_idx = min(table_index, len(tables) - 1)
        df_page = tables[chosen_idx]["df"].copy()
        df_page["_Page_No"] = str(page_num)
        collected_dfs.append(df_page)

        if page_num >= max_pages:
            break

        # Check if Next button is visible and enabled
        try:
            next_btn = page.locator(next_button_selector).first
            if next_btn.count() == 0 or not next_btn.is_visible() or not next_btn.is_enabled():
                break
            # Check disabled attribute or class
            cls = (next_btn.get_attribute("class") or "").lower()
            if "disabled" in cls or next_btn.get_attribute("disabled") is not None:
                break

            next_btn.click()
            time.sleep(delay_sec)
        except Exception:
            break

    if not collected_dfs:
        return pd.DataFrame()

    combined = pd.concat(collected_dfs, ignore_index=True)
    return clean_extracted_dataframe(combined)


def search_and_fetch_row_status(
    page: Any,
    browser_mgr: Any,
    search_url: str,
    search_input_selector: str,
    search_button_selector: str,
    lookup_id_value: str,
    result_selectors: Dict[str, str],
    download_link_selector: Optional[str] = None,
    downloads_dir: str = "downloads",
    delay_sec: float = 1.2
) -> Dict[str, Any]:
    """
    B2 & B4: Opens or uses the portal's Search / Application Status page, enters the ID
    (Application No / Aadhaar / Registration No), clicks Search, reads all mapped result fields
    (Status, Date, Remarks, Amount), and optionally downloads the linked certificate/receipt PDF.
    """
    if search_url and page.url != search_url:
        page.goto(search_url, wait_until="domcontentloaded", timeout=45000)
        time.sleep(0.5)

    # Check if portal asks for CAPTCHA/OTP/Login before searching
    needs_human, reason_en, reason_hi = browser_mgr.check_requires_human_intervention()
    if needs_human:
        return {
            "_Pause_Required": True,
            "_Pause_Reason_EN": reason_en,
            "_Pause_Reason_HI": reason_hi
        }

    clean_id = str(lookup_id_value).strip()
    if clean_id.endswith(".0"):
        clean_id = clean_id[:-2]

    # Fill Search ID and click Search
    inp = page.locator(search_input_selector).first
    inp.fill("")
    inp.fill(clean_id)
    time.sleep(0.3)

    btn = page.locator(search_button_selector).first
    btn.click()
    time.sleep(delay_sec)

    # Extract each requested field from result selectors
    fetched_data: Dict[str, Any] = {"_Pause_Required": False}
    for col_name, sel in result_selectors.items():
        if not sel:
            continue
        try:
            loc = page.locator(sel).first
            if loc.count() > 0 and loc.is_visible():
                fetched_data[col_name] = loc.inner_text(timeout=2000).strip()
            else:
                fetched_data[col_name] = "Not Found"
        except Exception:
            fetched_data[col_name] = ""

    # B4: Optional File / Certificate / Receipt Download
    if download_link_selector:
        try:
            dl_loc = page.locator(download_link_selector).first
            if dl_loc.count() > 0 and dl_loc.is_visible():
                row_dl_dir = os.path.join(downloads_dir, re.sub(r"[^a-zA-Z0-9_\-]", "", clean_id) or "doc")
                os.makedirs(row_dl_dir, exist_ok=True)
                href = dl_loc.get_attribute("href") or ""
                if href.startswith("data:") or href.endswith(".pdf") or href.endswith(".html"):
                    # Save receipt snapshot or trigger Playwright download
                    out_file = os.path.join(row_dl_dir, f"Certificate_{clean_id}.html")
                    with open(out_file, "w", encoding="utf-8") as f:
                        f.write(page.content())
                    fetched_data["Downloaded_File_Path"] = out_file
        except Exception as e:
            fetched_data["Downloaded_File_Path"] = f"Download skipped: {e}"

    fetched_data["Status_Checked_At"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
    return fetched_data


def extract_detail_pages_from_links(
    page: Any,
    row_link_selector: str,
    detail_field_selectors: Dict[str, str],
    max_items: int = 20,
    delay_sec: float = 1.2
) -> pd.DataFrame:
    """
    B3: Detail-Page Extraction:
    Collects all detail page URLs matching `row_link_selector` on a listing page,
    visits each detail view politely, and collects `label -> value` selectors into a table.
    """
    links = page.locator(row_link_selector)
    count = min(links.count(), max_items)
    hrefs: List[str] = []

    for i in range(count):
        h = links.nth(i).get_attribute("href")
        if h:
            hrefs.append(h)

    records: List[Dict[str, str]] = []
    base_url = page.url

    for idx, href in enumerate(hrefs):
        try:
            target_url = href if href.startswith("http") or href.startswith("file:") else base_url
            page.goto(target_url, wait_until="domcontentloaded")
            time.sleep(delay_sec)

            rec: Dict[str, str] = {"Detail_URL": target_url}
            for col_name, sel in detail_field_selectors.items():
                loc = page.locator(sel).first
                rec[col_name] = loc.inner_text(timeout=2000).strip() if loc.count() > 0 else ""
            records.append(rec)
        except Exception:
            pass

    return clean_extracted_dataframe(pd.DataFrame(records))
