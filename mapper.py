"""
mapper.py
=========
Smart AI Field Mapper & Template Manager.
Maps Excel/Word columns to scanned Portal Form Fields using:
- Anthropic Claude API (configurable via CLAUDE_MODEL env var, default: claude-sonnet-5-5, key: ANTHROPIC_API_KEY)
- Google Gemini API (GEMINI_API_KEY, model: gemini-3.8-flash)
- Deterministic Bilingual Heuristic Fallback (understands Hindi/English/Hinglish labels like "पिता का नाम" = "Father Name")

STRICT SAFETY & PRIVACY:
- Only sends column headers and 1-2 MASKED sample rows (Aadhaar/Mobile/Bank masked as XXXXXXXX1234).
- Returns strictly validated JSON mapping. NEVER runs eval() or exec() on generated output.
"""

from __future__ import annotations
import json
import os
import re
from typing import Dict, List, Any, Optional, Tuple
import pandas as pd
from utils import sanitize_sample_rows_for_ai

try:
    import anthropic
except ImportError:
    anthropic = None

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
    (["email", "email id", "ईमेल"], ["email", "mail"]),
    (["district", "जिला", "zila"], ["district", "dist"]),
    (["block", "tehsil", "panchayat samiti", "ब्लॉक", "तहसील", "पंचायत समिति"], ["block", "tehsil"]),
    (["village", "ward", "gram panchayat", "गाँव", "ग्राम", "वार्ड"], ["village", "gram", "ward"]),
    (["pincode", "pin code", "postal code", "पिन कोड", "पिनकोड"], ["pincode", "pin", "postal"]),
    (["address", "permanent address", "पता", "स्थायी पता"], ["address", "addr"]),
    (["photo", "applicant photo", "passport photo", "फोटो", "आवेदक की फोटो"], ["photo", "image", "pic"]),
    (["document", "id proof", "aadhaar pdf", "दस्तावेज़", "पहचान पत्र"], ["document", "doc", "id_proof"])
]


MAPPING_SYSTEM_PROMPT = """You are a Smart Portal Form Field Mapper for an automated form-filling assistant.
Your task is to match the user's Data Columns (from Excel/Word) to the target Web Portal Form Fields.
Understand English, Hindi (Devanagari), and Hinglish equivalents (for example: "पिता का नाम" matches "Father Name", "जिला" matches "District", "जन्म तिथि" matches "DOB").

Return ONLY a valid JSON object with a "mappings" list following this exact schema:
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
Allowed values for "transform": "none", "date_ddmmyyyy", "date_yyyymmdd", "upper", "lower", "title", "digits_only".
If a dropdown depends on another dropdown (e.g. Block depends on District, Village depends on Block), set "depends_on_selector" to the parent dropdown's selector and set "wait_after_ms": 800.
Never return executable code. Return ONLY valid JSON.
"""


def heuristic_field_mapping(
    data_columns: List[str],
    portal_fields: List[Dict[str, Any]]
) -> List[Dict[str, Any]]:
    """
    Deterministic bilingual matcher that pairs Excel/Word columns with portal form fields
    even without an internet connection or API key.
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
        # 1. Exact or substring match
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

        # Determine automatic transform rule
        transform = "none"
        if f_type == "date" or "dob" in combined_portal or "जन्म" in combined_portal:
            transform = "date_yyyymmdd" if f_type == "date" else "date_ddmmyyyy"
        elif any(k in combined_portal for k in ["mobile", "aadhaar", "pincode", "मोबाइल", "आधार", "पिन"]):
            transform = "digits_only"

        # Track cascading dependent dropdowns (District -> Block -> Village)
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


def suggest_mapping_with_ai(
    df: pd.DataFrame,
    portal_fields: List[Dict[str, Any]],
    provider: str = "auto"
) -> Tuple[List[Dict[str, Any]], str]:
    """
    Uses Anthropic Claude (CLAUDE_MODEL, default claude-sonnet-5-5) or Google Gemini
    to generate intelligent bilingual column-to-field mappings.
    Only sends column names and 2 masked sample rows.
    """
    data_columns = list(df.columns)
    raw_samples = df.head(2).to_dict(orient="records") if not df.empty else []
    masked_samples = sanitize_sample_rows_for_ai(raw_samples)

    user_payload = json.dumps({
        "data_columns": data_columns,
        "masked_sample_rows": masked_samples,
        "portal_form_fields": portal_fields
    }, ensure_ascii=False, indent=2)

    anthropic_key = os.environ.get("ANTHROPIC_API_KEY", "")
    gemini_key = os.environ.get("GEMINI_API_KEY", "")
    claude_model = os.environ.get("CLAUDE_MODEL", "claude-sonnet-5-5")

    # 1. Try Anthropic Claude if configured
    if (provider in ["auto", "claude"]) and anthropic_key and anthropic:
        try:
            client = anthropic.Anthropic(api_key=anthropic_key)
            resp = client.messages.create(
                model=claude_model,
                max_tokens=2000,
                system=MAPPING_SYSTEM_PROMPT,
                messages=[{"role": "user", "content": user_payload}]
            )
            text = "".join([b.text for b in resp.content if hasattr(b, "text")])
            match = re.search(r"\{.*\}", text, re.DOTALL)
            if match:
                parsed = json.loads(match.group(0))
                if "mappings" in parsed and isinstance(parsed["mappings"], list):
                    return parsed["mappings"], f"Mapped using Claude ({claude_model})"
        except Exception as e:
            print(f"Claude mapping fallback notice: {e}")

    # 2. Try Google Gemini if configured
    if (provider in ["auto", "gemini"]) and gemini_key and genai:
        try:
            client = genai.Client(api_key=gemini_key)
            resp = client.models.generate_content(
                model="gemini-3.8-flash",
                contents=user_payload,
                config=genai_types.GenerateContentConfig(
                    system_instruction=MAPPING_SYSTEM_PROMPT,
                    response_mime_type="application/json"
                ) if genai_types else None
            )
            match = re.search(r"\{.*\}", resp.text or "", re.DOTALL)
            if match:
                parsed = json.loads(match.group(0))
                if "mappings" in parsed and isinstance(parsed["mappings"], list):
                    return parsed["mappings"], "Mapped using Google Gemini (gemini-3.8-flash)"
        except Exception as e:
            print(f"Gemini mapping fallback notice: {e}")

    # 3. Fallback to deterministic bilingual heuristic mapper
    fallback_mappings = heuristic_field_mapping(data_columns, portal_fields)
    return fallback_mappings, "Mapped using Smart Bilingual Dictionary (Hindi/English)"


class PortalTemplateStore:
    """Saves and loads reusable JSON Portal Mapping Templates in ./templates/."""

    def __init__(self, templates_dir: str = "templates"):
        self.templates_dir = templates_dir
        os.makedirs(self.templates_dir, exist_ok=True)

    def save_template(
        self,
        template_name: str,
        portal_url: str,
        mappings: List[Dict[str, Any]],
        photo_settings: Optional[Dict[str, Any]] = None
    ) -> str:
        safe_name = re.sub(r"[^a-zA-Z0-9_\- ]", "", template_name).strip().replace(" ", "_")
        if not safe_name:
            safe_name = "portal_template"
        path = os.path.join(self.templates_dir, f"{safe_name}.json")
        payload = {
            "template_name": template_name,
            "portal_url": portal_url,
            "updated_at": pd.Timestamp.now().isoformat(),
            "photo_settings": photo_settings or {
                "max_kb": 50,
                "width": 200,
                "height": 230,
                "format": "JPEG",
                "match_by_column": "Aadhaar_No"
            },
            "mappings": mappings
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        return path

    def list_templates(self) -> List[Dict[str, Any]]:
        results: List[Dict[str, Any]] = []
        if not os.path.exists(self.templates_dir):
            return results
        for fname in sorted(os.listdir(self.templates_dir)):
            if fname.endswith(".json"):
                fpath = os.path.join(self.templates_dir, fname)
                try:
                    with open(fpath, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        data["_filename"] = fname
                        results.append(data)
                except Exception:
                    pass
        return results
