"""
llm_parser.py
=============
Converts natural language instructions in Hindi, English, or Hinglish into
safe, structured JSON action plans using Google Gemini OR Anthropic Claude.

Supports:
- Google Gemini API (GEMINI_API_KEY, model: gemini-3.8-flash)
- Anthropic Claude API (ANTHROPIC_API_KEY, model configurable via CLAUDE_MODEL)

Guarantees:
- Never outputs arbitrary executable code.
- Only schema names, columns, and 5 sample rows are sent to LLM (privacy & token efficiency).
- Supports Multi-Step Commands (returns an ordered list of JSON actions).
- Supports 'Ask Your Data' Q&A queries.
- Detects ambiguous requests and generates short, clarifying questions.
"""

from __future__ import annotations
import json
import os
import re
from typing import Any, Dict, List, Optional, Tuple, Union
import pandas as pd

# Gemini SDKs (Google GenAI)
try:
    from google import genai
    from google.genai import types as genai_types
except ImportError:
    genai = None
    genai_types = None

try:
    import google.generativeai as legacy_genai
except ImportError:
    legacy_genai = None

# Claude SDK
try:
    import anthropic
except ImportError:
    anthropic = None


SYSTEM_PROMPT = """You are the AI planner for a Voice-Controlled Excel Assistant.
Translate user voice or text instructions (in English, Hindi, or Hinglish) into a structured JSON action plan.
NEVER return python code, bash scripts, or free-form text. Return ONLY a valid JSON object.

MODEL INSTRUCTION SCHEMA:

1. SINGLE ACTION:
{
  "action": "vlookup" | "sort" | "find_common" | "filter" | "clean" | "pivot" | "calculate" | "merge_sheets" | "compare_versions" | "detect_anomalies",
  ... (parameters for action)
}

2. MULTI-STEP PIPELINE:
If the user gives multiple commands in one sentence (e.g. "Duplicate hatao, Price se sort karo, phir city wise pivot banao"), return:
{
  "is_multi_step": true,
  "steps": [
    {"action": "clean", "operation": "remove_duplicates", "output_name": "Step1_Cleaned"},
    {"action": "sort", "columns": ["Price"], "ascending": false, "output_name": "Step2_Sorted"},
    {"action": "pivot", "index_columns": ["City"], "values_columns": ["Amount"], "agg_func": "sum", "output_name": "Step3_Pivot"}
  ]
}

3. ASK YOUR DATA (Q&A Mode):
If the user asks an analytical question about the data (e.g. "Sabse zyada sale kis city mein hui?" or "Top 5 customers kaun hain?"):
{
  "action": "qa_query",
  "question": "User question here",
  "query_type": "top_n" | "aggregation" | "filter_summary",
  "target_column": "City",
  "metric_column": "Sales",
  "agg_func": "sum" | "max" | "min" | "mean",
  "limit": 5,
  "answer_template_hi": "सबसे ज़्यादा बिक्री {target_column} ({metric_column}) में हुई है।",
  "answer_template_en": "Highest sales occurred in {target_column} with total {metric_column}."
}

4. AMBIGUITY HANDLING:
If the user's instruction is ambiguous (e.g. multiple columns look like "Price" or "City", or table is not identifiable), DO NOT GUESS:
{
  "needs_clarification": true,
  "question_en": "Did you mean 'Unit_Price' or 'Total_Price' for this action?",
  "question_hi": "क्या आपका मतलब 'Unit_Price' से था या 'Total_Price' से?",
  "options": ["Unit_Price", "Total_Price"]
}

Always return strict, valid JSON. No text outside the JSON.
"""


def summarize_tables_for_prompt(tables_store: Dict[str, pd.DataFrame]) -> str:
    """Summarizes sheets, columns, and 5 sample rows for prompt payload."""
    summary_parts = []
    for name, df in tables_store.items():
        sample_rows = df.head(5).to_dict(orient="records")
        clean_sample = []
        for row in sample_rows:
            clean_row = {}
            for k, v in row.items():
                if pd.isna(v):
                    clean_row[k] = None
                else:
                    clean_row[k] = str(v)
            clean_sample.append(clean_row)

        summary_parts.append(
            f"--- TABLE / SHEET: '{name}' ---\n"
            f"Columns: {list(df.columns)}\n"
            f"Total rows in sheet: {len(df)}\n"
            f"Sample 5 rows:\n{json.dumps(clean_sample, indent=2)}\n"
        )
    return "\n".join(summary_parts)


def parse_instruction_with_gemini(
    instruction: str,
    tables_store: Dict[str, pd.DataFrame],
    api_key: Optional[str] = None,
    latest_result_name: Optional[str] = None,
    model_name: str = "gemini-3.8-flash"
) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    """
    Parses natural language instruction using Google Gemini API.
    Model: gemini-3.8-flash (fast, highly capable for structured tool & reasoning tasks).
    """
    resolved_key = api_key or os.environ.get("GEMINI_API_KEY")
    if not resolved_key:
        return None, (
            "Google Gemini API Key is missing. Please set the GEMINI_API_KEY environment variable "
            "or enter it in the settings sidebar."
        )

    tables_summary = summarize_tables_for_prompt(tables_store)
    user_prompt = f"""
AVAILABLE TABLES IN CURRENT WORKSPACE:
{tables_summary}

LATEST ACTIVE TABLE RESULT:
{latest_result_name if latest_result_name else list(tables_store.keys())[0]}

USER INSTRUCTION (Can be in English, Hindi, or Hinglish):
"{instruction}"

Produce the structured JSON action plan to execute this instruction.
Return ONLY valid JSON matching the schema.
"""

    # 1. Try google.genai modern SDK
    if genai is not None:
        try:
            client = genai.Client(api_key=resolved_key)
            response = client.models.generate_content(
                model=model_name,
                contents=user_prompt,
                config=genai_types.GenerateContentConfig(
                    system_instruction=SYSTEM_PROMPT,
                    response_mime_type="application/json"
                ) if genai_types else None
            )

            text_resp = response.text or ""
            match = re.search(r"\{.*\}", text_resp, re.DOTALL)
            if match:
                return json.loads(match.group(0)), None
            return json.loads(text_resp), None
        except Exception as e:
            # Fall through to legacy SDK if needed
            print(f"google.genai SDK attempt: {e}")

    # 2. Try google.generativeai legacy SDK
    if legacy_genai is not None:
        try:
            legacy_genai.configure(api_key=resolved_key)
            # Use gemini-1.5-flash alias or model
            model = legacy_genai.GenerativeModel(
                model_name="gemini-1.5-flash",
                system_instruction=SYSTEM_PROMPT,
                generation_config={"response_mime_type": "application/json"}
            )
            response = model.generate_content(user_prompt)
            text_resp = response.text or ""
            match = re.search(r"\{.*\}", text_resp, re.DOTALL)
            if match:
                return json.loads(match.group(0)), None
            return json.loads(text_resp), None
        except Exception as e:
            return None, f"Gemini API error: {str(e)}"

    return None, "Neither 'google-genai' nor 'google-generativeai' package is installed. Run 'pip install google-genai'."


def parse_instruction_with_claude(
    instruction: str,
    tables_store: Dict[str, pd.DataFrame],
    api_key: Optional[str] = None,
    latest_result_name: Optional[str] = None,
    model_override: Optional[str] = None
) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    """Parses instruction using Anthropic Claude."""
    resolved_key = api_key or os.environ.get("ANTHROPIC_API_KEY")
    if not resolved_key:
        return None, (
            "Anthropic API key is missing. Please set ANTHROPIC_API_KEY or use your Gemini key."
        )

    if not anthropic:
        return None, "anthropic package is not installed. Run 'pip install anthropic'."

    model_name = model_override or os.environ.get("CLAUDE_MODEL", "claude-sonnet-5-5")
    tables_summary = summarize_tables_for_prompt(tables_store)

    user_prompt = f"""
AVAILABLE TABLES IN CURRENT WORKSPACE:
{tables_summary}

LATEST ACTIVE TABLE RESULT:
{latest_result_name if latest_result_name else list(tables_store.keys())[0]}

USER INSTRUCTION (Can be in English, Hindi, or Hinglish):
"{instruction}"

Produce the structured JSON action plan.
"""

    try:
        client = anthropic.Anthropic(api_key=resolved_key)
        response = client.messages.create(
            model=model_name,
            max_tokens=1800,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_prompt}]
        )
        content_text = "".join([b.text for b in response.content if hasattr(b, "text")])
        match = re.search(r"\{.*\}", content_text, re.DOTALL)
        if match:
            return json.loads(match.group(0)), None
        return None, f"Could not parse structured JSON from Claude response."
    except Exception as exc:
        return None, f"Error calling Claude: {str(exc)}"


def parse_instruction_unified(
    instruction: str,
    tables_store: Dict[str, pd.DataFrame],
    provider: str = "auto",
    gemini_key: Optional[str] = None,
    claude_key: Optional[str] = None,
    latest_result_name: Optional[str] = None
) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    """
    Unified entry point:
    Supports provider = 'gemini' | 'claude' | 'auto'.
    If 'auto', checks if GEMINI_API_KEY or ANTHROPIC_API_KEY is available.
    """
    has_gemini = bool(gemini_key or os.environ.get("GEMINI_API_KEY"))
    has_claude = bool(claude_key or os.environ.get("ANTHROPIC_API_KEY"))

    target_provider = provider.lower()
    if target_provider == "auto":
        target_provider = "gemini" if has_gemini else ("claude" if has_claude else "gemini")

    if target_provider == "gemini":
        return parse_instruction_with_gemini(
            instruction=instruction,
            tables_store=tables_store,
            api_key=gemini_key,
            latest_result_name=latest_result_name
        )
    else:
        return parse_instruction_with_claude(
            instruction=instruction,
            tables_store=tables_store,
            api_key=claude_key,
            latest_result_name=latest_result_name
        )
