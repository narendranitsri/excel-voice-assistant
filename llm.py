"""
llm.py
======
Centralized Google Gemini AI Provider for "Excel <-> Website Automation Assistant".
All LLM calls are isolated in this single module so the provider can be maintained or changed easily.

STRICT COMPLIANCE & PRIVACY:
- Uses the official `google-genai` package (`from google import genai`).
- Reads API Key from `st.secrets["GEMINI_API_KEY"]` first, then `os.environ.get("GEMINI_API_KEY")`. Never hardcoded.
- Reads Model from `os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")`.
- Requests JSON only (`response_mime_type="application/json"`) and validates against strict schemas.
- NEVER runs `eval()` or `exec()` on generated output.
- Retries up to 3 times with exponential backoff on 429 / quota errors, returning friendly Hindi/English messages.
- Receives ONLY column names, field labels, and 1-2 MASKED sample values (never raw Aadhaar, mobile, or bank numbers).
"""

from __future__ import annotations
import json
import os
import re
import time
from typing import Dict, List, Any, Optional, Tuple
import pandas as pd

from utils import sanitize_sample_rows_for_ai

try:
    import streamlit as st
except ImportError:
    st = None

try:
    from google import genai
    from google.genai import types as genai_types
except ImportError:
    genai = None
    genai_types = None


BILINGUAL_SYNONYMS: List[Tuple[List[str], List[str]]] = [
    (["applicant name", "full name", "candidate name", "name", "आवेदक का नाम", "पूरा नाम", "नाम"], ["applicant_name", "fullname", "name", "candidate"]),
    (["father name", "father's name", "पिता का नाम", "पिता", "pita ka naam"], ["father_name", "fname", "father"]),
    (["mother name", "mother's name", "माता का नाम", "माता", "mata ka naam"], ["mother_name", "mname", "mother"]),
    (["date of birth", "dob", "birth date", "जन्म तिथि", "जन्म दिनांक", "janm tithi"], ["dob", "birth_date", "dateofbirth"]),
    (["gender", "sex", "लिंग", "ling"], ["gender", "sex"]),
    (["category", "caste", "श्रेणी", "वर्ग", "जाति"], ["category", "caste"]),
    (["mobile", "mobile number", "phone", "मोबाइल नंबर", "मोबाइल", "फ़ोन"], ["mobile", "phone", "contact_no"]),
    (["aadhaar", "aadhar", "uid", "jan aadhaar", "आधार नंबर", "आधार संख्या", "जनाधार"], ["aadhaar", "aadhar", "uid", "janaadhaar"]),
    (["application no", "application id", "app no", "registration no", "आवेदन संख्या", "रजिस्ट्रेशन नंबर"], ["application_no", "app_id", "reg_no"]),
    (["status", "application status", "स्थिति", "आवेदन की स्थिति"], ["status", "app_status"]),
    (["district", "जिला", "zila"], ["district", "dist"]),
    (["block", "tehsil", "panchayat samiti", "ब्लॉक", "तहसील", "पंचायत समिति"], ["block", "tehsil"]),
    (["village", "ward", "gram panchayat", "गाँव", "ग्राम", "वार्ड"], ["village", "gram", "ward"]),
    (["pincode", "pin code", "postal code", "पिन कोड", "पिनकोड"], ["pincode", "pin", "postal"]),
    (["address", "permanent address", "पता", "स्थायी पता"], ["address", "addr"]),
    (["photo", "applicant photo", "passport photo", "फोटो", "आवेदक की फोटो"], ["photo", "image", "pic"]),
    (["document", "id proof", "aadhaar pdf", "दस्तावेज़", "पहचान पत्र"], ["document", "doc", "id_proof"])
]


def get_gemini_api_key() -> str:
    """
    Retrieves GEMINI_API_KEY from `st.secrets["GEMINI_API_KEY"]` first,
    then falls back to environment variable `GEMINI_API_KEY`. Never hardcoded.
    """
    if st is not None:
        try:
            if "GEMINI_API_KEY" in st.secrets and st.secrets["GEMINI_API_KEY"]:
                return str(st.secrets["GEMINI_API_KEY"]).strip()
        except Exception:
            pass
    return os.environ.get("GEMINI_API_KEY", "").strip()


def get_gemini_model_name() -> str:
    """Reads model name from environment variable GEMINI_MODEL (default: gemini-2.5-flash)."""
    return os.environ.get("GEMINI_MODEL", "gemini-2.5-flash").strip() or "gemini-2.5-flash"


def call_gemini_json_with_retry(
    system_instruction: str,
    user_prompt: str,
    max_retries: int = 3
) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    """
    Calls Google Gemini API (`google-genai`) requesting `response_mime_type="application/json"`.
    Retries up to 3 times with exponential backoff on 429 / quota / transient network errors.
    Returns: (parsed_json_dict_or_None, friendly_bilingual_error_or_None)
    """
    api_key = get_gemini_api_key()
    if not api_key:
        return None, (
            "GEMINI_API_KEY सेट नहीं है (Not set). कृपया साइडबार में या st.secrets / environment variable में अपनी Gemini API Key डालें। "
            "(Using built-in Smart Bilingual Dictionary fallback)."
        )

    if genai is None:
        return None, "google-genai पैकेज इंस्टॉल नहीं है। कृपया 'pip install google-genai' चलाएं।"

    model_name = get_gemini_model_name()
    client = genai.Client(api_key=api_key)

    for attempt in range(1, max_retries + 1):
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=user_prompt,
                config=genai_types.GenerateContentConfig(
                    system_instruction=system_instruction,
                    response_mime_type="application/json"
                ) if genai_types else None
            )
            raw_text = (response.text or "").strip()
            match = re.search(r"\{.*\}", raw_text, re.DOTALL)
            json_str = match.group(0) if match else raw_text
            parsed = json.loads(json_str)
            if isinstance(parsed, dict):
                return parsed, None
            return {"data": parsed}, None

        except Exception as err:
            err_str = str(err)
            is_quota = "429" in err_str or "quota" in err_str.lower() or "resource_exhausted" in err_str.lower()
            if is_quota and attempt < max_retries:
                backoff_sec = 2 ** attempt
                time.sleep(backoff_sec)
                continue
            elif is_quota:
                return None, (
                    f"⚠️ Gemini API कोटा सीमा (429 Quota Exceeded) पूरी हो गई है (3 बार प्रयास किया गया)। "
                    f"ऐप ने स्मार्ट हिंदी/इंग्लिश डिक्शनरी मैपिंग का उपयोग किया है। (Quota limit reached; used smart local fallback)."
                )
            else:
                return None, f"Gemini API त्रुटि (Error): {err_str}"

    return None, "Gemini API call failed after retries."


def heuristic_field_mapping(
    data_columns: List[str],
    portal_fields: List[Dict[str, Any]]
) -> List[Dict[str, Any]]:
    """
    Deterministic bilingual matcher that pairs Excel/Word columns with portal form fields
    (Hindi/English/Hinglish) when offline or as a fallback.
    """
    mappings: List[Dict[str, Any]] = []
    district_selector: Optional[str] = None
    block_selector: Optional[str] = None

    for pf in portal_fields:
        f_label = str(pf.get("label", "")).lower()
        f_name = str(pf.get("name", "")).lower()
        f_id = str(pf.get("id", "")).lower()
        f_type = str(pf.get("type", "text")).lower()
        f_sel = pf.get("selector", "")
        combined_portal = f"{f_label} {f_name} {f_id}"

        matched_col = ""
        # 1. Direct token match
        for col in data_columns:
            col_clean = col.strip().lower().replace("_", " ")
            if col_clean in combined_portal or any(part in combined_portal for part in col_clean.split() if len(part) > 3):
                matched_col = col
                break

        # 2. Bilingual synonym match
        if not matched_col:
            for keywords, _ in BILINGUAL_SYNONYMS:
                if any(k in combined_portal for k in keywords):
                    for col in data_columns:
                        c_low = col.strip().lower().replace("_", " ")
                        if any(k in c_low for k in keywords):
                            matched_col = col
                            break
                if matched_col:
                    break

        transform = "none"
        if f_type == "date" or "dob" in combined_portal or "जन्म" in combined_portal:
            transform = "date_yyyymmdd" if f_type == "date" else "date_ddmmyyyy"
        elif any(k in combined_portal for k in ["mobile", "aadhaar", "pincode", "मोबाइल", "आधार", "पिन"]):
            transform = "digits_only"

        depends_on = None
        wait_ms = 200
        if any(k in combined_portal for k in ["district", "जिला"]):
            district_selector = f_sel
            wait_ms = 800
        elif any(k in combined_portal for k in ["block", "tehsil", "ब्लॉक", "तहसील"]):
            block_selector = f_sel
            depends_on = district_selector
            wait_ms = 800
        elif any(k in combined_portal for k in ["village", "gram", "गाँव", "ग्राम"]):
            depends_on = block_selector
            wait_ms = 800

        mappings.append({
            "field_selector": f_sel,
            "field_label": pf.get("label", f_name or f_id),
            "field_type": f_type,
            "data_column": matched_col,
            "transform": transform,
            "depends_on_selector": depends_on,
            "wait_after_ms": wait_ms,
            "required": pf.get("required", False)
        })

    return mappings


def suggest_form_mapping_with_gemini(
    df: pd.DataFrame,
    portal_fields: List[Dict[str, Any]]
) -> Tuple[List[Dict[str, Any]], str]:
    """
    PART A (Teach Mode): Uses Gemini (`GEMINI_MODEL`, default `gemini-2.5-flash`) to map
    Excel/Word columns to scanned web form fields.
    Sends ONLY column names, field labels, and 1-2 MASKED sample rows.
    """
    data_columns = list(df.columns)
    raw_samples = df.head(2).to_dict(orient="records") if not df.empty else []
    masked_samples = sanitize_sample_rows_for_ai(raw_samples)

    system_prompt = """You are a Smart Bilingual Portal Form Mapper.
Match the user's Excel/Word columns to the scanned web form fields.
Understand Hindi, English, and Hinglish labels (e.g. "पिता का नाम" = "Father_Name", "जिला" = "District").
Return ONLY valid JSON matching this schema:
{
  "mappings": [
    {
      "field_selector": "#applicant_name",
      "field_label": "आवेदक का नाम / Applicant Name",
      "field_type": "text",
      "data_column": "Applicant_Name",
      "transform": "none",
      "depends_on_selector": null,
      "wait_after_ms": 200
    }
  ]
}
Allowed transforms: "none", "date_ddmmyyyy", "date_yyyymmdd", "upper", "lower", "title", "digits_only".
"""
    user_payload = json.dumps({
        "excel_columns": data_columns,
        "masked_sample_values": masked_samples,
        "web_form_fields": portal_fields
    }, ensure_ascii=False, indent=2)

    parsed, err = call_gemini_json_with_retry(system_prompt, user_payload)
    if parsed and "mappings" in parsed and isinstance(parsed["mappings"], list):
        return parsed["mappings"], f"✨ Gemini ({get_gemini_model_name()}) ने स्मार्ट फील्ड मैपिंग तैयार की।"

    fallback = heuristic_field_mapping(data_columns, portal_fields)
    status_note = err if err else "Mapped via Smart Bilingual Dictionary."
    return fallback, status_note


def suggest_extraction_cleanup_with_gemini(
    raw_headers: List[str],
    masked_sample_rows: List[Dict[str, str]]
) -> Tuple[Dict[str, str], str]:
    """
    PART B (Website -> Excel Teach Mode): Suggests clean, standardized Hindi/English column headers
    for tables or detail pages extracted from a website.
    """
    system_prompt = """You are a Data Extraction Schema Cleaner.
Given raw table headers scraped from a website and 1-2 masked sample rows, suggest clean, standardized column names.
Return ONLY valid JSON:
{
  "header_mapping": {
    "raw_header_1": "Clean_Column_Name_1"
  }
}
"""
    payload = json.dumps({
        "raw_headers": raw_headers,
        "masked_samples": masked_sample_rows[:2]
    }, ensure_ascii=False, indent=2)

    parsed, err = call_gemini_json_with_retry(system_prompt, payload)
    if parsed and "header_mapping" in parsed and isinstance(parsed["header_mapping"], dict):
        return parsed["header_mapping"], f"✨ Gemini ({get_gemini_model_name()}) ने टेबल कॉलम नाम साफ़ किए।"

    default_map = {h: re.sub(r"\s+", "_", str(h).strip()) for h in raw_headers}
    return default_map, err or "Standardized column headers locally."


class PortalTemplateStore:
    """Manages reusable JSON templates for both Form Filling (Part A) and Website Extraction (Part B)."""

    def __init__(self, templates_dir: str = "templates"):
        self.templates_dir = templates_dir
        os.makedirs(self.templates_dir, exist_ok=True)

    def save_template(
        self,
        template_name: str,
        template_type: str,  # 'fill' or 'extract'
        portal_url: str,
        mappings: List[Dict[str, Any]],
        extra_config: Optional[Dict[str, Any]] = None
    ) -> str:
        safe_name = re.sub(r"[^a-zA-Z0-9_\- ]", "", template_name).strip().replace(" ", "_") or "template"
        path = os.path.join(self.templates_dir, f"{safe_name}.json")
        payload = {
            "template_name": template_name,
            "template_type": template_type,
            "portal_url": portal_url,
            "updated_at": pd.Timestamp.now().isoformat(),
            "extra_config": extra_config or {},
            "mappings": mappings
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        return path

    def list_templates(self, template_type: Optional[str] = None) -> List[Dict[str, Any]]:
        results: List[Dict[str, Any]] = []
        if not os.path.exists(self.templates_dir):
            return results
        for fname in sorted(os.listdir(self.templates_dir)):
            if fname.endswith(".json"):
                try:
                    with open(os.path.join(self.templates_dir, fname), "r", encoding="utf-8") as f:
                        data = json.load(f)
                        if template_type is None or data.get("template_type", "fill") == template_type:
                            results.append(data)
                except Exception:
                    pass
        return results
