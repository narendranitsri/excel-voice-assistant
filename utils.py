"""
utils.py
========
Privacy Masking, Rate Limiting, Voice Command Parser, and Bilingual (Hindi/English) UI Helpers.
"""

from __future__ import annotations
import re
import time
import os
from typing import Any, Dict, List, Optional


def mask_sensitive_value(value: Any, col_name: str = "") -> str:
    """
    Masks sensitive PII (Aadhaar, Mobile, Bank Account, PAN) so logs and AI prompts
    never expose raw citizen data.
    Examples:
    - Aadhaar (12 digits): XXXXXXXX1234
    - Mobile (10 digits): XXXXXX7890
    - Bank Account: XXXXXXXX5678
    """
    if value is None:
        return ""
    val_str = str(value).strip()
    if not val_str or val_str.lower() == "nan":
        return ""

    col_lower = col_name.lower()
    digits_only = re.sub(r"\D", "", val_str)

    # Check if Aadhaar (12 digits or column name indicates Aadhaar/JanAadhaar)
    if "aadhaar" in col_lower or "aadhar" in col_lower or "जनाधार" in col_lower or "आधार" in col_lower or len(digits_only) == 12:
        if len(digits_only) >= 4:
            return "X" * (len(digits_only) - 4) + digits_only[-4:]

    # Check if Mobile / Phone (10 digits)
    if any(k in col_lower for k in ["mobile", "phone", "contact", "मोबाइल", "फ़ोन"]) or len(digits_only) == 10:
        if len(digits_only) >= 4:
            return "X" * (len(digits_only) - 4) + digits_only[-4:]

    # Check if Bank Account / IFSC / PAN
    if any(k in col_lower for k in ["account", "bank", "खाता", "pan"]):
        if len(val_str) > 4:
            return "X" * (len(val_str) - 4) + val_str[-4:]

    return val_str


def sanitize_sample_rows_for_ai(rows: List[Dict[str, Any]]) -> List[Dict[str, str]]:
    """
    Takes 1-2 sample rows and masks all sensitive numbers before sending to Gemini for field mapping.
    """
    masked_rows: List[Dict[str, str]] = []
    for row in rows[:2]:
        clean_row: Dict[str, str] = {}
        for col, val in row.items():
            clean_row[str(col)] = mask_sensitive_value(val, str(col))
        masked_rows.append(clean_row)
    return masked_rows


def format_value_for_portal(
    value: Any,
    transform_rule: Optional[str] = None
) -> str:
    """
    Applies automatic value formatting required by portals:
    - 'date_ddmmyyyy': converts dates to DD/MM/YYYY
    - 'date_yyyymmdd': converts dates to YYYY-MM-DD (for HTML <input type="date">)
    - 'upper': UPPERCASE
    - 'lower': lowercase
    - 'title': Proper Case
    - 'digits_only': strips spaces and dashes
    """
    if value is None:
        return ""
    val_str = str(value).strip()
    if val_str.lower() in ["nan", "none", "null"]:
        return ""

    # Clean trailing .0 from numeric strings read by pandas (e.g. 9876543210.0)
    if re.match(r"^\d+\.0$", val_str):
        val_str = val_str[:-2]

    if not transform_rule or transform_rule == "none":
        return val_str

    rule = transform_rule.lower()
    if rule == "upper":
        return val_str.upper()
    elif rule == "lower":
        return val_str.lower()
    elif rule == "title":
        return val_str.title()
    elif rule == "digits_only":
        return re.sub(r"\D", "", val_str)
    elif rule in ["date_ddmmyyyy", "date_yyyymmdd"]:
        import pandas as pd
        try:
            dt = pd.to_datetime(val_str, dayfirst=True, errors="coerce")
            if pd.notna(dt):
                return dt.strftime("%d/%m/%Y") if rule == "date_ddmmyyyy" else dt.strftime("%Y-%m-%d")
        except Exception:
            pass

    return val_str


def parse_voice_assistant_command(command_text: str) -> Dict[str, Any]:
    """
    Parses Hindi / English / Hinglish voice or text commands for both Part A (Fill) and Part B (Extract):
    - "agli row bharo" / "fill next row" -> {"intent": "next_row"}
    - "pehli 10 rows bharo" / "fill first 10 rows" -> {"intent": "range_rows", "count": 10}
    - "sirf failed rows dobara chalao" / "retry failed" -> {"intent": "retry_failed"}
    - "photo upload skip karo" -> {"intent": "skip_photos", "value": True}
    - "is page ki table Excel mein nikalo" / "extract table" -> {"intent": "extract_table"}
    - "sabhi application ka status check karo" / "check all status" -> {"intent": "check_all_status"}
    - "reconcile karo" / "compare data" -> {"intent": "reconcile"}
    - "pause karo" / "ruk jao" -> {"intent": "pause"}
    - "scan page" / "form scan karo" -> {"intent": "scan_page"}
    """
    cmd = command_text.strip().lower()
    if not cmd:
        return {"intent": "unknown", "message": "Empty command"}

    if any(p in cmd for p in ["table excel", "टेबल एक्सेल", "extract table", "nikalo", "निकालो"]):
        return {
            "intent": "extract_table",
            "message_hi": "वेब पेज से टेबल निकालकर एक्सेल में सेव की जा रही है...",
            "message_en": "Extracting HTML table from page into Excel..."
        }

    if any(p in cmd for p in ["status check", "स्टेटस चेक", "application ka status", "check all status", "स्थिति जांचो"]):
        return {
            "intent": "check_all_status",
            "message_hi": "सभी आवेदन संख्याओं (Application IDs) का स्टेटस वेबसाइट से चेक किया जा रहा है...",
            "message_en": "Checking application status for all IDs on the portal..."
        }

    if any(p in cmd for p in ["reconcile", "compare", "मिलान", "अंतर"]):
        return {
            "intent": "reconcile",
            "message_hi": "एक्सेल और वेबसाइट के डेटा का मिलान (Reconcile) किया जा रहा है...",
            "message_en": "Reconciling Excel data with website records..."
        }

    if any(p in cmd for p in ["agli row", "next row", "अगली रो", "ek row"]):
        return {"intent": "next_row", "message_hi": "अगली रो (Row) भरी जा रही है...", "message_en": "Filling next row..."}

    match_num = re.search(r"(\d+)", cmd)
    if any(p in cmd for p in ["pehli", "first", "प्रथम", "shuru ki"]) and match_num:
        n = int(match_num.group(1))
        return {"intent": "range_rows", "count": n, "message_hi": f"पहली {n} पंक्तियाँ (Rows) सेलेक्ट की गईं।", "message_en": f"Selected first {n} rows."}

    if any(p in cmd for p in ["failed", "फेल", "dobara", "retry"]):
        return {"intent": "retry_failed", "message_hi": "केवल फेल हुई पंक्तियों को दोबारा चलाया जा रहा है।", "message_en": "Retrying failed rows only."}

    if any(p in cmd for p in ["photo", "फोटो", "document", "skip"]):
        return {"intent": "skip_photos", "value": True, "message_hi": "फोटो अपलोड स्किप कर दिया गया है।", "message_en": "Photo upload skipped."}

    if any(p in cmd for p in ["pause", "ruk", "रुक", "stop", "रोको"]):
        return {"intent": "pause", "message_hi": "ऑटो-फिलर रोक दिया गया है।", "message_en": "Auto-filler paused."}

    if any(p in cmd for p in ["scan", "स्कैन", "teach"]):
        return {"intent": "scan_page", "message_hi": "वेब पेज के फॉर्म फील्ड्स स्कैन किए जा रहे हैं...", "message_en": "Scanning form fields on page..."}

    return {"intent": "unknown", "message_hi": f"कमांड समझ नहीं आया: '{command_text}'", "message_en": f"Command not recognized: '{command_text}'"}
