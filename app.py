"""
app.py
======
Voice-Controlled Excel Assistant (Phase 1 to Phase 4 Complete Architecture)
Built with Streamlit, pandas, openpyxl, rapidfuzz, Plotly, and Anthropic Claude.
Model configurable via CLAUDE_MODEL env var (default: claude-sonnet-5-5).
"""

from __future__ import annotations
import os
import json
import io
from typing import Dict, List, Optional, Any
import pandas as pd
import streamlit as st

# Custom Modules (Phases 1-4)
from actions import (
    dispatch_action_plan,
    export_dataframe_with_formatting,
    execute_compare_versions,
    execute_detect_anomalies,
    execute_dry_run_preview,
    get_excel_learning_explanation,
    ActionExecutionError
)
from llm_parser import parse_instruction_unified
from formulas import explain_formula
from quality import audit_data_quality, auto_fix_data_quality
from charts import generate_plotly_chart, generate_quick_dashboard
from workflows import WorkflowRecorder
from utils import (
    extract_table_from_pdf,
    extract_table_from_image_with_claude,
    generate_voice_reply_audio,
    create_executive_summary_sheet
)

# Optional mic recorder
try:
    from streamlit_mic_recorder import mic_recorder, speech_to_text
except ImportError:
    mic_recorder = None
    speech_to_text = None

# Configure Streamlit
st.set_page_config(
    page_title="Voice-Controlled Excel Assistant",
    page_icon="🎙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize Session State
if "tables_store" not in st.session_state:
    st.session_state.tables_store = {}

if "original_tables" not in st.session_state:
    st.session_state.original_tables = {}

if "history" not in st.session_state:
    st.session_state.history = []

if "active_table_name" not in st.session_state:
    st.session_state.active_table_name = None

if "clarification_state" not in st.session_state:
    st.session_state.clarification_state = None

if "voice_transcript" not in st.session_state:
    st.session_state.voice_transcript = ""

if "last_json_plan" not in st.session_state:
    st.session_state.last_json_plan = None

if "dry_run_pending" not in st.session_state:
    st.session_state.dry_run_pending = None

if "active_workflow_steps" not in st.session_state:
    st.session_state.active_workflow_steps = []

if "qa_response" not in st.session_state:
    st.session_state.qa_response = None

workflow_mgr = WorkflowRecorder()

# Custom CSS
st.markdown("""
<style>
    .main-header { font-size: 2rem; font-weight: 700; color: #0F172A; margin-bottom: 0.2rem; }
    .sub-header { font-size: 0.95rem; color: #475569; margin-bottom: 1.2rem; }
    .privacy-box { background: #F0FDF4; border: 1px solid #BBF7D0; padding: 0.6rem 0.9rem; border-radius: 6px; font-size: 0.85rem; color: #166534; margin-bottom: 1rem; }
    .card-kpi { background: #F8FAFC; border: 1px solid #E2E8F0; padding: 0.8rem; border-radius: 8px; text-align: center; }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------
# Sidebar Configuration
# ---------------------------------------------------------
with st.sidebar:
    st.title("🎙️ Assistant Settings")

    # AI Engine Provider Selection
    ai_provider = st.radio(
        "Select AI Engine",
        options=["Google Gemini (Recommended)", "Anthropic Claude"],
        index=0 if os.environ.get("GEMINI_API_KEY") or not os.environ.get("ANTHROPIC_API_KEY") else 1
    )

    if "Gemini" in ai_provider:
        default_gemini = os.environ.get("GEMINI_API_KEY", "")
        gemini_key_input = st.text_input(
            "Google Gemini API Key",
            value=default_gemini,
            type="password",
            help="Reads from GEMINI_API_KEY env var by default. Uses model: gemini-3.8-flash."
        )
        if gemini_key_input:
            os.environ["GEMINI_API_KEY"] = gemini_key_input
        st.caption("✨ Using **Google Gemini (gemini-3.8-flash)** for high-speed voice instruction planning.")
    else:
        default_claude = os.environ.get("ANTHROPIC_API_KEY", "")
        claude_key_input = st.text_input(
            "Anthropic Claude API Key",
            value=default_claude,
            type="password",
            help="Reads from ANTHROPIC_API_KEY env var."
        )
        if claude_key_input:
            os.environ["ANTHROPIC_API_KEY"] = claude_key_input
        default_model = os.environ.get("CLAUDE_MODEL", "claude-sonnet-5-5")
        model_choice = st.text_input("Claude Model", value=default_model)
        if model_choice:
            os.environ["CLAUDE_MODEL"] = model_choice

    st.markdown("---")
    st.subheader("🔊 Audio & Language")
    lang_choice = st.selectbox(
        "Speech Recognition Language",
        options=["hi-IN (Hindi / Hinglish)", "en-IN (Indian English)", "en-US (English)"],
        index=0
    )
    lang_code = lang_choice.split(" ")[0]

    enable_tts = st.toggle("Voice Reply (Text-to-Speech)", value=True, help="Speaks short status after actions.")
    write_formulas_toggle = st.toggle("Write Dynamic Excel Formulas", value=True, help="Writes real XLOOKUP / SUMIFS formulas in output cells.")

    st.markdown("---")
    st.subheader("↺ Action History & Undo")
    if st.session_state.history:
        st.write(f"Actions recorded: **{len(st.session_state.history)}**")
        if st.button("⏪ Undo Last Action", use_container_width=True):
            last_state = st.session_state.history.pop()
            st.session_state.tables_store = last_state["tables_store"]
            st.session_state.active_table_name = last_state["active_table_name"]
            st.success(f"Undid: '{last_state.get('description')}'")
            st.rerun()
    else:
        st.caption("No operations performed yet.")

    st.markdown("---")
    st.subheader("💾 Saved Workflows (Macros)")
    saved_macros = workflow_mgr.list_workflows()
    if saved_macros:
        macro_names = [m["name"] for m in saved_macros]
        selected_macro = st.selectbox("Select Workflow", macro_names)
        if st.button("▶️ Run Workflow on Active Table", use_container_width=True):
            target_m = next((m for m in saved_macros if m["name"] == selected_macro), None)
            if target_m:
                updated_store, logs = workflow_mgr.replay_workflow(
                    target_m["steps"],
                    st.session_state.tables_store,
                    dispatch_action_plan
                )
                st.session_state.tables_store = updated_store
                st.success(f"Executed workflow '{selected_macro}' successfully!")
                st.rerun()


# ---------------------------------------------------------
# Main App Header
# ---------------------------------------------------------
st.markdown("<div class='main-header'>Voice-Controlled Excel Assistant</div>", unsafe_allow_html=True)
st.markdown("<div class='sub-header'>A smart, bilingual spreadsheet copilot with voice recognition, live Excel formulas, data quality audits, and automated dashboards.</div>", unsafe_allow_html=True)

# Privacy Note (from User Goal Step 1)
st.markdown("""
<div class='privacy-box'>
    🔒 <b>Privacy Guarantee:</b> Only sheet names, column headers, and 5 sample rows are sent to Claude for instruction planning. Your full dataset remains local and private on your machine.
</div>
""", unsafe_allow_html=True)


# ---------------------------------------------------------
# Step 1: File Uploader, Sample Loader, PDF/Image Extractor
# ---------------------------------------------------------
with st.expander("📁 1. Load Data: Excel, CSV, PDF or Image Table", expanded=(len(st.session_state.tables_store) == 0)):
    tab_f1, tab_f2, tab_f3 = st.tabs(["Spreadsheets (Excel/CSV)", "Extract Table from PDF", "Extract Table from Image (Photo/Invoice)"])

    with tab_f1:
        col_up, col_samples = st.columns([2, 1])
        with col_up:
            up_files = st.file_uploader("Upload Excel (.xlsx, .xls) or CSV files", type=["xlsx", "xls", "csv"], accept_multiple_files=True)
            if up_files and st.button("Process Files", key="btn_proc_files"):
                for f in up_files:
                    fname = f.name
                    if fname.endswith(".csv"):
                        df = pd.read_csv(f)
                        st.session_state.tables_store[os.path.splitext(fname)[0]] = df
                    else:
                        xls = pd.ExcelFile(f)
                        for s in xls.sheet_names:
                            key = f"{os.path.splitext(fname)[0]}_{s}" if len(xls.sheet_names) > 1 else s
                            st.session_state.tables_store[key] = pd.read_excel(xls, sheet_name=s)
                st.session_state.active_table_name = list(st.session_state.tables_store.keys())[0]
                st.success("Files loaded successfully!")
                st.rerun()

        with col_samples:
            st.write("Or try prebuilt test tables:")
            if st.button("Load 4 Sample Tables (Orders, Prices, Inventory, Messy Data)", key="btn_load_samples"):
                samples_dir = "samples"
                for sf, key in [
                    ("sample_orders.xlsx", "Orders"),
                    ("sample_prices.xlsx", "Prices"),
                    ("sample_inventory.xlsx", "Inventory"),
                    ("sample_messy_data.xlsx", "Messy_Data")
                ]:
                    p = os.path.join(samples_dir, sf)
                    if os.path.exists(p):
                        st.session_state.tables_store[key] = pd.read_excel(p)
                st.session_state.active_table_name = "Orders"
                st.success("Loaded 4 sample tables: Orders, Prices, Inventory, and Messy_Data!")
                st.rerun()

    with tab_f2:
        pdf_file = st.file_uploader("Upload PDF with tables (e.g. invoice or report)", type=["pdf"])
        if pdf_file and st.button("Extract Tables from PDF"):
            with st.spinner("Extracting tabular data from PDF..."):
                dfs = extract_table_from_pdf(pdf_file.read())
                if dfs:
                    for idx, extracted_df in enumerate(dfs, start=1):
                        tbl_name = f"PDF_Table_{idx}"
                        st.session_state.tables_store[tbl_name] = extracted_df
                    st.session_state.active_table_name = f"PDF_Table_1"
                    st.success(f"Extracted {len(dfs)} table(s) from PDF!")
                    st.rerun()
                else:
                    st.warning("No structured tables detected in the uploaded PDF.")

    with tab_f3:
        img_file = st.file_uploader("Upload photo/screenshot of a table or invoice", type=["png", "jpg", "jpeg"])
        if img_file and st.button("Extract Table with Claude Vision"):
            with st.spinner("Analyzing image with Claude Vision..."):
                img_df = extract_table_from_image_with_claude(img_file.read())
                if img_df is not None:
                    tbl_name = f"Image_Table_{len(st.session_state.tables_store)+1}"
                    st.session_state.tables_store[tbl_name] = img_df
                    st.session_state.active_table_name = tbl_name
                    st.success("Successfully extracted table from image!")
                    st.rerun()
                else:
                    st.error("Failed to extract structured table from the provided image.")


# ---------------------------------------------------------
# Tables View & Quick Actions Bar
# ---------------------------------------------------------
if st.session_state.tables_store:
    tab_keys = list(st.session_state.tables_store.keys())
    active_idx = tab_keys.index(st.session_state.active_table_name) if st.session_state.active_table_name in tab_keys else 0

    col_sel, col_audit, col_dash, col_sum = st.columns([3, 2, 2, 2])
    with col_sel:
        selected_tbl = st.selectbox("Active Table", tab_keys, index=active_idx)
        st.session_state.active_table_name = selected_tbl

    with col_audit:
        if st.button("🩺 Data Quality Report", use_container_width=True):
            st.session_state["show_quality_audit"] = True

    with col_dash:
        if st.button("📊 Quick Dashboard", use_container_width=True):
            st.session_state["show_dashboard"] = True

    with col_sum:
        if st.button("📑 Report Banao (Summary)", use_container_width=True):
            st.session_state["show_summary_report"] = True

    active_df = st.session_state.tables_store[selected_tbl]
    st.dataframe(active_df.head(10), use_container_width=True)

    # 1. Quality Audit Panel if toggled
    if st.session_state.get("show_quality_audit"):
        with st.expander("🩺 Data Quality Health Report", expanded=True):
            audit = audit_data_quality(active_df)
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Health Score", f"{audit['health_score']}/100")
            c2.metric("Duplicate Rows", audit["duplicate_rows"])
            c3.metric("Blank Cells", audit["total_blanks"])
            c4.metric("Mixed Type Columns", len(audit["mixed_type_columns"]))

            if audit["blanks_by_column"]:
                st.write(f"**Blank Cells by Column:** `{audit['blanks_by_column']}`")
            if audit["extra_spaces_columns"]:
                st.write(f"**Extra Spaces Found:** `{', '.join(audit['extra_spaces_columns'])}`")
            if audit["invalid_emails"]:
                st.write(f"**Invalid Emails Detected:** `{audit['invalid_emails']}`")
            if audit["invalid_phones"]:
                st.write(f"**Invalid Phones Detected:** `{audit['invalid_phones']}`")

            if st.button("✨ Fix All Automatically (Smart Clean & Standardize)"):
                fixed_df, fix_logs = auto_fix_data_quality(active_df)
                st.session_state.tables_store[f"{selected_tbl}_Fixed"] = fixed_df
                st.session_state.active_table_name = f"{selected_tbl}_Fixed"
                st.success("Auto-fix applied! Created standardized table.")
                st.rerun()

    # 2. Quick Dashboard if toggled
    if st.session_state.get("show_dashboard"):
        with st.expander("📊 Quick Dashboard & Visual Insights", expanded=True):
            dash = generate_quick_dashboard(active_df)
            kpi_cols = st.columns(len(dash["kpis"]))
            for idx, (k, v) in enumerate(dash["kpis"].items()):
                kpi_cols[idx].metric(k, v)
            for ch in dash["charts"]:
                st.plotly_chart(ch, use_container_width=True)


# ---------------------------------------------------------
# Voice & Text Command Center
# ---------------------------------------------------------
st.markdown("### 🗣️ Voice Command or Text Instruction")

col_v1, col_v2 = st.columns([2, 1])

with col_v1:
    if speech_to_text:
        spoken = speech_to_text(language=lang_code, start_prompt="🎙️ Press to Speak", stop_prompt="⏹️ Stop Recording")
        if spoken:
            st.session_state.voice_transcript = spoken

    user_cmd = st.text_area(
        "Voice Instruction (Editable)",
        value=st.session_state.voice_transcript,
        placeholder="e.g. Orders aur Prices me Item se VLOOKUP lagao aur Unit_Price le aao",
        height=75
    )

with col_v2:
    st.write("⚡ **10 Example Voice Commands:**")
    example_prompts = [
        "Orders aur Prices me Item se VLOOKUP lagao aur Unit_Price le aao",
        "Duplicate hatao, Quantity se sort karo, phir city wise pivot banao",
        "Orders aur Inventory me common product nikalo",
        "Add calculated column Total = Quantity * Unit_Price",
        "City ke hisab se Quantity ka total sum pivot banao",
        "Sabse zyada sale kis city mein hui?",
        "Duplicate rows hata do",
        "Quantity ke hisab se descending sort karo",
        "Galat ya ajeeb values dhundo",
        "Sales ka bar chart banao"
    ]
    selected_ex = st.selectbox("Choose quick command:", example_prompts)
    if st.button("Insert Command"):
        st.session_state.voice_transcript = selected_ex
        st.rerun()


# ---------------------------------------------------------
# Clarification Dialog (Step 4)
# ---------------------------------------------------------
if st.session_state.clarification_state:
    c = st.session_state.clarification_state
    st.warning("❓ **Clarification Needed (स्पष्टीकरण चाहिए)**")
    st.write(f"**English:** {c.get('question_en')}")
    st.write(f"**Hindi:** {c.get('question_hi')}")
    opts = c.get("options", [])
    if opts:
        choice = st.radio("Select matching column:", opts)
        if st.button("Confirm Choice"):
            st.session_state.voice_transcript += f" (use {choice})"
            st.session_state.clarification_state = None
            st.rerun()


# ---------------------------------------------------------
# Dry Run Confirmation Dialog (Feature 18)
# ---------------------------------------------------------
if st.session_state.dry_run_pending:
    pending = st.session_state.dry_run_pending
    st.warning("⚠️ **Dry Run Preview: Confirmation Required**")
    st.write(pending["message"])
    st.dataframe(pending["preview_df"].head(10), use_container_width=True)

    c_yes, c_no = st.columns(2)
    with c_yes:
        if st.button("✅ Confirm & Execute Action", type="primary"):
            plan = pending["plan"]
            st.session_state.dry_run_pending = None
            try:
                res_df, msg, meta = dispatch_action_plan(plan, st.session_state.tables_store)
                out_name = meta.get("output_name", "Result")
                st.session_state.tables_store[out_name] = res_df
                st.session_state.active_table_name = out_name
                st.success(f"Confirmed: {msg}")
                st.rerun()
            except Exception as e:
                st.error(str(e))
    with c_no:
        if st.button("❌ Cancel"):
            st.session_state.dry_run_pending = None
            st.info("Action cancelled.")
            st.rerun()


# ---------------------------------------------------------
# Execution Flow (Single-Step, Multi-Step, Q&A)
# ---------------------------------------------------------
if st.button("🚀 Execute Action", type="primary"):
    if not user_cmd.strip():
        st.warning("Please speak or type an instruction.")
    elif not st.session_state.tables_store:
        st.error("No tables loaded. Please load or upload files.")
    else:
        provider_name = "gemini" if "Gemini" in ai_provider else "claude"
        spinner_msg = f"🤖 {ai_provider.split(' ')[0]} analyzing instruction..."
        with st.spinner(spinner_msg):
            plan, err = parse_instruction_unified(
                instruction=user_cmd,
                tables_store=st.session_state.tables_store,
                provider=provider_name,
                latest_result_name=st.session_state.active_table_name
            )

        if err:
            st.error(f"❌ {err}")
        elif plan:
            # 1. Clarification needed
            if plan.get("needs_clarification"):
                st.session_state.clarification_state = plan
                st.rerun()

            # 2. Q&A Mode ("Ask Your Data")
            elif plan.get("action") == "qa_query":
                tgt_col = plan.get("target_column")
                met_col = plan.get("metric_column")
                agg = plan.get("agg_func", "sum")
                curr = st.session_state.tables_store[st.session_state.active_table_name]
                
                if tgt_col in curr.columns and met_col in curr.columns:
                    grouped = curr.groupby(tgt_col)[met_col].agg(agg).sort_values(ascending=False).reset_index()
                    top_item = grouped.iloc[0]
                    ans = f"**Answer:** Top {tgt_col} is **{top_item[tgt_col]}** with {met_col} = **{top_item[met_col]:,}**"
                    st.success(ans)
                    st.dataframe(grouped.head(5), use_container_width=True)
                else:
                    st.info(f"Query: {plan.get('question')}")

            # 3. Multi-Step Pipeline (Feature 14)
            elif plan.get("is_multi_step"):
                steps = plan.get("steps", [])
                st.info(f"Executing {len(steps)} sequential steps...")
                prog_bar = st.progress(0)

                for idx, step in enumerate(steps, start=1):
                    res_df, msg, meta = dispatch_action_plan(step, st.session_state.tables_store)
                    out_name = meta.get("output_name", f"Step_{idx}")
                    st.session_state.tables_store[out_name] = res_df
                    st.session_state.active_table_name = out_name
                    prog_bar.progress(idx / len(steps))

                st.success(f"Completed all {len(steps)} steps successfully!")
                st.rerun()

            # 4. Standard Single Action with Dry Run Check (Feature 18)
            else:
                act = plan.get("action", "")
                if act in ["clean", "remove_duplicates"] and plan.get("operation") == "remove_duplicates":
                    prev_df, dry_msg, meta = execute_dry_run_preview(plan, st.session_state.tables_store)
                    if meta.get("rows_affected", 0) > 0:
                        st.session_state.dry_run_pending = {
                            "plan": plan,
                            "preview_df": prev_df,
                            "message": dry_msg
                        }
                        st.rerun()

                # Normal execution
                st.session_state.last_json_plan = plan
                # Record undo state
                st.session_state.history.append({
                    "tables_store": {k: v.copy() for k, v in st.session_state.tables_store.items()},
                    "active_table_name": st.session_state.active_table_name,
                    "description": user_cmd[:50]
                })

                res_df, msg, meta = dispatch_action_plan(plan, st.session_state.tables_store)
                out_name = meta.get("output_name", "Result")
                st.session_state.tables_store[out_name] = res_df
                st.session_state.active_table_name = out_name

                # Voice Reply (TTS Feature 23)
                if enable_tts:
                    audio_bytes = generate_voice_reply_audio(msg, lang="hi" if "hi" in lang_code else "en")
                    if audio_bytes:
                        st.audio(audio_bytes, format="audio/mp3")

                st.success(f"✅ {msg}")
                st.rerun()


# ---------------------------------------------------------
# Excel Learning Mode & Macro Saver Panels
# ---------------------------------------------------------
if st.session_state.last_json_plan:
    c_learn, c_macro = st.columns(2)
    with c_learn:
        with st.expander("💡 Ye Kaise Kiya? (Excel Learning Mode)", expanded=False):
            exp = get_excel_learning_explanation(st.session_state.last_json_plan)
            st.write(f"**English:**\n{exp.get('en')}")
            st.write(f"**Hindi (हिंदी में सीखें):**\n{exp.get('hi')}")

    with c_macro:
        with st.expander("💾 Save as Named Macro / Workflow", expanded=False):
            macro_title = st.text_input("Macro Name", "Monthly Sales Cleanup")
            if st.button("Save Workflow Profile"):
                path = workflow_mgr.save_workflow(
                    name=macro_title,
                    description=f"Automated workflow for {macro_title}",
                    steps=[st.session_state.last_json_plan]
                )
                st.success(f"Workflow saved: '{macro_title}'")


# ---------------------------------------------------------
# Export Options (.xlsx, .csv, and Summary Report Sheet)
# ---------------------------------------------------------
if st.session_state.active_table_name in st.session_state.tables_store:
    curr_df = st.session_state.tables_store[st.session_state.active_table_name]
    st.markdown("---")
    st.subheader(f"📥 Export Results: '{st.session_state.active_table_name}'")

    c_dl1, c_dl2 = st.columns(2)
    with c_dl1:
        xlsx_buf = export_dataframe_with_formatting(curr_df)
        st.download_button(
            label="⬇️ Download Formatted Excel (.xlsx)",
            data=xlsx_buf,
            file_name=f"{st.session_state.active_table_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    with c_dl2:
        csv_data = curr_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="⬇️ Download CSV (.csv)",
            data=csv_data,
            file_name=f"{st.session_state.active_table_name}.csv",
            mime="text/csv"
        )
