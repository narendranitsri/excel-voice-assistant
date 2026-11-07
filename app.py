"""
app.py
Excel <-> Website Automation Assistant (Streamlit Main Entry Point)

Supports three core modes via top Mode Switch:
1. "Excel to Website" (Part A): Import Excel/CSV/Word, validate data, Teach Mode form scan + Gemini mapping, Semi-Auto/Auto Playwright filling with cascading dropdowns, photo compression & upload, human Login/OTP/CAPTCHA pause, and result Excel/PDF export.
2. "Website to Excel" (Part B): Extract HTML tables (with pagination), Search-and-Fetch status by Application No/Aadhaar from an input Excel, Detail-Page scraping, certificate/PDF downloads, and clean Excel export.
3. "Both (fill and verify)" (Part C): Submit or verify previously submitted rows against live portal status, plus full Excel vs. Website reconciliation with a color-coded "Differences" sheet.
"""

import os
import io
import json
import time
import pandas as pd
import streamlit as st

from importer import load_spreadsheet_sheets, load_word_document
from validator import validate_dataset, ValidationIssue
from browser import PortalBrowserManager
from scanner import scan_live_page_fields, scan_html_string_fields
from llm import (
    get_gemini_api_key,
    get_gemini_model_name,
    generate_field_mapping_with_gemini,
    suggest_extraction_selectors_with_gemini,
    PortalTemplateStore,
)
from uploader import match_document_for_row, compress_and_resize_image
from filler import fill_row_into_portal
from runner import BatchExecutionState
from extractor import (
    detect_html_tables_from_string,
    extract_paginated_table_live,
    search_and_fetch_from_excel_live,
    extract_detail_pages_live,
    clean_extracted_dataframe,
)
from reconcile import reconcile_excel_vs_website, export_reconciliation_excel
from report import export_results_excel, generate_pdf_summary_report
from utils import (
    mask_sensitive_value,
    mask_dataframe_for_display,
    parse_voice_assistant_command,
)

st.set_page_config(
    page_title="Excel <-> Website Automation Assistant",
    page_icon="🔄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Session State Initialization
# ---------------------------------------------------------------------------
if "ui_lang" not in st.session_state:
    st.session_state.ui_lang = "Bilingual (Hindi + English)"
if "active_df" not in st.session_state:
    st.session_state.active_df = None
if "raw_excel_bytes" not in st.session_state:
    st.session_state.raw_excel_bytes = None
if "scanned_fields" not in st.session_state:
    st.session_state.scanned_fields = []
if "field_mappings" not in st.session_state:
    st.session_state.field_mappings = []
if "browser_mgr" not in st.session_state:
    st.session_state.browser_mgr = PortalBrowserManager(profile_dir="./browser_profile")
if "batch_state" not in st.session_state:
    st.session_state.batch_state = BatchExecutionState()
if "extracted_df" not in st.session_state:
    st.session_state.extracted_df = None
if "reconcile_summary_df" not in st.session_state:
    st.session_state.reconcile_summary_df = None
if "reconcile_diff_df" not in st.session_state:
    st.session_state.reconcile_diff_df = None

template_store = PortalTemplateStore("./templates")

# ---------------------------------------------------------------------------
# Sidebar Configuration & Safety Controls
# ---------------------------------------------------------------------------
with st.sidebar:
    st.title("⚙️ Assistant Settings")
    st.caption("Human-in-the-Loop Portal & Excel Automation")

    st.session_state.ui_lang = st.selectbox(
        "भाषा / Interface Language",
        ["Bilingual (Hindi + English)", "English", "हिंदी (Hindi)"],
        index=0,
    )

    st.divider()
    st.subheader("🔑 Google Gemini API (`llm.py`)")
    active_key = get_gemini_api_key()
    active_model = get_gemini_model_name()
    if active_key:
        st.success(f"GEMINI_API_KEY detected (Model: `{active_model}`)")
    else:
        st.warning("GEMINI_API_KEY not found in st.secrets or environment.")
        manual_key = st.text_input("Enter GEMINI_API_KEY (Session Only):", type="password")
        if manual_key:
            os.environ["GEMINI_API_KEY"] = manual_key
            st.rerun()

    st.divider()
    st.subheader("🌐 Playwright Headed Browser")
    default_demo_path = os.path.abspath("./samples/demo_portal.html")
    default_url = f"file://{default_demo_path}" if os.path.exists(default_demo_path) else "https://sso.rajasthan.gov.in"
    portal_url = st.text_input("Target Portal URL:", value=default_url)

    col_b1, col_b2 = st.columns(2)
    with col_b1:
        if st.button("🚀 Launch Browser", use_container_width=True):
            ok, msg = st.session_state.browser_mgr.launch_headed_browser(portal_url)
            if ok:
                st.toast(msg)
            else:
                st.error(msg)
    with col_b2:
        if st.button("🛑 Close Browser", use_container_width=True):
            st.session_state.browser_mgr.close()
            st.toast("Browser closed.")

    st.divider()
    st.subheader("🛡️ Safety & Rate Limiting")
    action_delay = st.slider("Delay between fields (sec):", 0.5, 5.0, 1.2, 0.1)
    semi_auto_mode = st.toggle("Semi-Automatic Mode (Stop before Submit)", value=True)
    auto_submit_confirmed = False
    if not semi_auto_mode:
        st.warning("⚠️ Auto-Submit Enabled! The app will click Submit automatically after filling each row.")
        auto_submit_confirmed = st.checkbox("I confirm I want Auto-Submit without manual review per row.")

    mask_pii_ui = st.toggle("Mask Aadhaar/Mobile/Bank in Preview & Logs", value=True)

    st.divider()
    with st.expander("🔒 Privacy & Terms Notice", expanded=False):
        st.markdown(
            """
            - **No Bypass Rule**: This app **never** bypasses or automates Login, OTP, or CAPTCHA. When detected, it pauses so you can complete them in the browser.
            - **Zero Password Storage**: Passwords and OTPs are never read or saved.
            - **PII Masking**: Sensitive numbers (Aadhaar, Mobile, Bank Account) are masked (`XXXXXXXX1234`) in logs and never sent in full to Gemini.
            - **Terms of Use**: Only use this tool on websites and portals where you are authorized to submit or access data.
            """
        )

# ---------------------------------------------------------------------------
# Top Header & Mode Switch
# ---------------------------------------------------------------------------
st.title("🔄 Excel ↔ Website Automation Assistant")
st.markdown(
    "**Two-Way Smart Assistant**: Fill government/enterprise web forms from Excel/Word **(Part A)**, extract web tables & application statuses into Excel **(Part B)**, or verify & reconcile both **(Part C)**."
)

mode = st.radio(
    "Select Workflow Mode / कार्य मोड चुनें:",
    [
        "Excel to Website",
        "Website to Excel",
        "Both (fill and verify)",
    ],
    horizontal=True,
)

# Optional Voice / Text Command Bar
with st.expander("🎙️ Voice / Quick Command Bar (Hindi, English, Hinglish)", expanded=False):
    col_v1, col_v2 = st.columns([3, 1])
    with col_v1:
        voice_text = st.text_input(
            "Speak or type a command (e.g., 'agli row bharo', 'isko chhodo', 'ruko', 'table nikalo', 'status check karo'):",
            placeholder="अगली पंक्ति भरें / agli row bharo / table nikalo / status check karo",
        )
    with col_v2:
        try:
            from streamlit_mic_recorder import speech_to_text
            spoken = speech_to_text(language="hi-IN", start_prompt="🎤 Speak Command", stop_prompt="⏹️ Stop", key="mic_cmd")
            if spoken:
                voice_text = spoken
        except Exception:
            st.caption("Install `streamlit-mic-recorder` for live mic button.")

    if voice_text:
        parsed_cmd = parse_voice_assistant_command(voice_text)
        st.info(f"**Recognized Command**: `{parsed_cmd['intent']}` — {parsed_cmd['message_hi']} ({parsed_cmd['message_en']})")

st.divider()


# ===========================================================================
# MODE 1: EXCEL TO WEBSITE (PART A)
# ===========================================================================
if mode == "Excel to Website":
    tab_import, tab_teach, tab_run, tab_report = st.tabs([
        "1️⃣ Data Import & Validation",
        "2️⃣ Teach Mode & Field Mapping",
        "3️⃣ Semi-Auto Form Filler",
        "4️⃣ Reports & Updated Excel",
    ])

    # --- TAB 1: DATA IMPORT & VALIDATION ---
    with tab_import:
        st.subheader("📂 Step 1: Upload Excel (.xlsx/.csv) or Word (.docx)")
        col_u1, col_u2 = st.columns([2, 1])
        with col_u1:
            uploaded_file = st.file_uploader(
                "Choose Excel, CSV, or Word file",
                type=["xlsx", "xls", "csv", "docx"],
            )
        with col_u2:
            st.markdown("**Or load bundled sample data:**")
            if st.button("📥 Load Sample Applicants (.xlsx)", use_container_width=True):
                sample_path = "./samples/sample_applicants.xlsx"
                if os.path.exists(sample_path):
                    with open(sample_path, "rb") as f:
                        raw_bytes = f.read()
                    st.session_state.raw_excel_bytes = raw_bytes
                    sheets = load_spreadsheet_sheets(raw_bytes, "sample_applicants.xlsx")
                    first_sheet = list(sheets.keys())[0]
                    st.session_state.active_df = sheets[first_sheet]
                    st.success(f"Loaded `{sample_path}` ({len(st.session_state.active_df)} rows)")

        if uploaded_file is not None:
            raw_bytes = uploaded_file.read()
            fname = uploaded_file.name
            if fname.lower().endswith(".docx"):
                doc_res = load_word_document(raw_bytes, fname)
                st.info(f"Extracted {len(doc_res['tables'])} table(s) and {len(doc_res['key_values'])} Label:Value fields from Word.")
                st.session_state.active_df = doc_res["combined_df"]
            else:
                st.session_state.raw_excel_bytes = raw_bytes
                sheets = load_spreadsheet_sheets(raw_bytes, fname)
                sheet_name = st.selectbox("Select Sheet:", list(sheets.keys()))
                st.session_state.active_df = sheets[sheet_name]

        if st.session_state.active_df is not None:
            df = st.session_state.active_df
            st.markdown("#### ✏️ Editable Data Preview (Fix any errors directly below)")
            display_df = mask_dataframe_for_display(df) if mask_pii_ui else df
            edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True, key="part_a_editor")
            st.session_state.active_df = edited_df

            # Row selection
            col_r1, col_r2, col_r3 = st.columns(3)
            with col_r1:
                row_filter_mode = st.selectbox("Rows to Process:", ["All Rows", "Row Range", "Only Pending Rows"])
            with col_r2:
                start_row = st.number_input("Start Row (1-indexed):", min_value=1, max_value=max(1, len(edited_df)), value=1)
            with col_r3:
                end_row = st.number_input("End Row (1-indexed):", min_value=1, max_value=max(1, len(edited_df)), value=len(edited_df))

            # Pre-flight validation
            st.markdown("#### ✅ Pre-Flight Data Validation (Aadhaar, Mobile, PIN, Dates, Duplicates)")
            mandatory_cols = st.multiselect(
                "Mandatory Columns (अनिवार्य कॉलम):",
                options=list(edited_df.columns),
                default=[c for c in edited_df.columns if any(k in c.lower() for k in ["name", "aadhaar", "mobile", "district"])][:4],
            )
            issues = validate_dataset(edited_df, mandatory_columns=mandatory_cols)
            if not issues:
                st.success("🎉 All rows passed validation checks! (सभी पंक्तियाँ सही हैं)")
            else:
                st.warning(f"Found {len(issues)} validation issue(s) across rows. Fix them in the table above before filling:")
                issue_rows = [
                    {
                        "Excel Row": i.row_index + 1,
                        "Column": i.column,
                        "Severity": i.severity.upper(),
                        "Message (EN)": i.message_en,
                        "संदेश (HI)": i.message_hi,
                    }
                    for i in issues
                ]
                st.dataframe(pd.DataFrame(issue_rows), use_container_width=True)

    # --- TAB 2: TEACH MODE & FIELD MAPPING ---
    with tab_teach:
        st.subheader("🎓 Step 2: Teach Mode (Scan Form Fields & Map to Excel Columns)")
        col_t1, col_t2 = st.columns(2)

        with col_t1:
            if st.button("🔍 Scan Open Browser Page Fields", use_container_width=True):
                if st.session_state.browser_mgr.page:
                    st.session_state.scanned_fields = scan_live_page_fields(st.session_state.browser_mgr.page)
                    st.success(f"Scanned {len(st.session_state.scanned_fields)} form fields from live browser!")
                else:
                    # Fallback to scanning local demo_portal.html if browser isn't open
                    demo_path = "./samples/demo_portal.html"
                    if os.path.exists(demo_path):
                        with open(demo_path, "r", encoding="utf-8") as f:
                            st.session_state.scanned_fields = scan_html_string_fields(f.read())
                        st.info(f"Browser not open — scanned {len(st.session_state.scanned_fields)} fields from `samples/demo_portal.html`.")
                    else:
                        st.error("Please launch the browser first or load a saved template.")

        with col_t2:
            saved_templates = template_store.list_templates()
            selected_tpl = st.selectbox("📂 Load Saved Portal Template:", ["-- Select Template --"] + saved_templates)
            if selected_tpl != "-- Select Template --" and st.button("📥 Apply Template"):
                tpl_data = template_store.load_template(selected_tpl)
                if tpl_data:
                    st.session_state.field_mappings = tpl_data.get("mappings", [])
                    st.success(f"Loaded template `{selected_tpl}` ({len(st.session_state.field_mappings)} mappings)!")

        if st.session_state.scanned_fields and st.session_state.active_df is not None:
            if st.button("✨ Auto-Map Columns with Google Gemini (`llm.py`)", type="primary"):
                with st.spinner("Sending masked column names & field labels to Gemini..."):
                    res = generate_field_mapping_with_gemini(
                        df=st.session_state.active_df,
                        scanned_fields=st.session_state.scanned_fields,
                    )
                    st.session_state.field_mappings = res.get("mappings", [])
                    if res.get("used_fallback"):
                        st.warning(res.get("message_hi", "Used heuristic fallback."))
                    else:
                        st.success(res.get("message_hi", "Mapping generated!"))

        if st.session_state.field_mappings:
            st.markdown("#### 🔗 Review & Edit Column → Portal Field Mapping")
            map_df = pd.DataFrame(st.session_state.field_mappings)
            edited_map_df = st.data_editor(map_df, num_rows="dynamic", use_container_width=True, key="mapping_editor")
            st.session_state.field_mappings = edited_map_df.to_dict(orient="records")

            col_s1, col_s2 = st.columns([2, 1])
            with col_s1:
                new_tpl_name = st.text_input("Template Name to Save:", value="SSO_Registration_Template")
            with col_s2:
                if st.button("💾 Save Template JSON", use_container_width=True):
                    path = template_store.save_template(
                        name=new_tpl_name,
                        portal_url=portal_url,
                        mode="excel_to_website",
                        mappings=st.session_state.field_mappings,
                    )
                    st.success(f"Saved template to `{path}`")

    # --- TAB 3: SEMI-AUTO FORM FILLER ---
    with tab_run:
        st.subheader("🤖 Step 3: Interactive Row-by-Row Form Filling")
        if st.session_state.active_df is None:
            st.info("Please upload or load an Excel/Word file in Tab 1 first.")
        else:
            df = st.session_state.active_df
            photos_folder = st.text_input("📁 Applicant Photos / Documents Folder:", value="./samples/photos")
            max_photo_kb = st.number_input("Max Photo Size (KB) for Auto-Resize/Compression:", value=50, min_value=10, max_value=500)

            curr_idx = st.number_input(
                "Current Row Index (0-based):",
                min_value=0,
                max_value=max(0, len(df) - 1),
                value=st.session_state.batch_state.current_row_index,
            )
            st.session_state.batch_state.current_row_index = curr_idx
            current_row = df.iloc[curr_idx].to_dict()

            st.markdown(f"#### 👤 Current Row #{curr_idx + 1} Preview")
            masked_preview = {k: (mask_sensitive_value(v, k) if mask_pii_ui else v) for k, v in current_row.items()}
            st.json(masked_preview)

            # Check matched photo
            matched_photo = match_document_for_row(current_row, photos_folder)
            if matched_photo:
                st.caption(f"📎 Matched Document/Photo: `{matched_photo}` (Will be auto-compressed ≤ {max_photo_kb} KB)")

            # Check for Login / OTP / CAPTCHA pause
            if st.session_state.browser_mgr.page:
                obstacle = st.session_state.browser_mgr.detect_security_obstacle()
                if obstacle["paused"]:
                    st.error(
                        f"⏸️ **PAUSED ({obstacle['reason']})**: Please login / enter OTP / captcha in the browser, then click Continue."
                    )
                    if st.button("▶️ Continue After Login / OTP / CAPTCHA", type="primary"):
                        st.rerun()

            col_act1, col_act2, col_act3, col_act4 = st.columns(4)
            with col_act1:
                if st.button("✍️ Fill Current Row in Browser", type="primary", use_container_width=True):
                    if not st.session_state.browser_mgr.page:
                        st.error("Please click 'Launch Browser' in the sidebar first!")
                    else:
                        fill_res = fill_row_into_portal(
                            page=st.session_state.browser_mgr.page,
                            row_data=current_row,
                            mappings=st.session_state.field_mappings,
                            photos_dir=photos_folder,
                            max_photo_kb=max_photo_kb,
                            action_delay=action_delay,
                        )
                        if fill_res["status"] == "PAUSED_FOR_HUMAN":
                            st.warning(fill_res["message"])
                        elif fill_res["status"] == "READY_FOR_REVIEW":
                            st.success(
                                "✅ Form filled! Stopped before Submit (Semi-Automatic Mode). Review in browser and click 'Approve & Submit' below."
                            )
                            if not semi_auto_mode and auto_submit_confirmed:
                                sub_res = st.session_state.browser_mgr.click_submit_and_capture(row_index=curr_idx)
                                st.session_state.batch_state.record_row_outcome(
                                    row_index=curr_idx,
                                    status=sub_res["status"],
                                    reference_no=sub_res["reference_no"],
                                    screenshot_path=sub_res["screenshot_path"],
                                    error_msg=sub_res.get("error", ""),
                                )
                                st.success(f"Auto-Submitted! Ref No: {sub_res['reference_no']}")

            with col_act2:
                if st.button("✅ Approve & Submit Row", use_container_width=True):
                    if st.session_state.browser_mgr.page:
                        sub_res = st.session_state.browser_mgr.click_submit_and_capture(row_index=curr_idx)
                        st.session_state.batch_state.record_row_outcome(
                            row_index=curr_idx,
                            status=sub_res["status"],
                            reference_no=sub_res["reference_no"],
                            screenshot_path=sub_res["screenshot_path"],
                            error_msg=sub_res.get("error", ""),
                        )
                        st.session_state.batch_state.current_row_index = min(len(df) - 1, curr_idx + 1)
                        st.success(f"Submitted Row #{curr_idx + 1}! Application No: `{sub_res['reference_no']}`")
                    else:
                        st.error("Browser is not open.")

            with col_act3:
                if st.button("⏭️ Skip Row (अगली पंक्ति)", use_container_width=True):
                    st.session_state.batch_state.record_row_outcome(
                        row_index=curr_idx,
                        status="Skipped",
                        reference_no="",
                        screenshot_path="",
                        error_msg="Skipped by user",
                    )
                    st.session_state.batch_state.current_row_index = min(len(df) - 1, curr_idx + 1)
                    st.rerun()

            with col_act4:
                if st.button("💾 Save Checkpoint", use_container_width=True):
                    cp_path = st.session_state.batch_state.save_checkpoint()
                    st.toast(f"Saved batch checkpoint to `{cp_path}`")

    # --- TAB 4: REPORTS & UPDATED EXCEL ---
    with tab_report:
        st.subheader("📊 Step 4: Execution Log, Proof Screenshots & Updated Excel")
        outcomes = st.session_state.batch_state.row_outcomes
        if not outcomes:
            st.info("No rows processed yet in this session.")
        else:
            out_df = pd.DataFrame.from_dict(outcomes, orient="index")
            st.dataframe(out_df, use_container_width=True)

            if st.session_state.active_df is not None:
                updated_excel_bytes = export_results_excel(
                    original_df=st.session_state.active_df,
                    outcomes=outcomes,
                    raw_original_bytes=st.session_state.raw_excel_bytes,
                )
                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    st.download_button(
                        "📥 Download Updated Excel (with Status & Application No)",
                        data=updated_excel_bytes,
                        file_name="applicants_results_updated.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                with col_d2:
                    pdf_bytes = generate_pdf_summary_report(outcomes)
                    st.download_button(
                        "📄 Download PDF Audit Summary Report",
                        data=pdf_bytes,
                        file_name="automation_audit_report.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )


# ===========================================================================
# MODE 2: WEBSITE TO EXCEL (PART B)
# ===========================================================================
elif mode == "Website to Excel":
    tab_table, tab_search, tab_detail = st.tabs([
        "1️⃣ Table & List Extractor (Multi-Page)",
        "2️⃣ Search-and-Fetch Status from Excel",
        "3️⃣ Detail-Page & File Downloader",
    ])

    # --- TAB 1: TABLE & LIST EXTRACTOR ---
    with tab_table:
        st.subheader("📋 Extract HTML Tables & Lists into Clean Excel")
        col_e1, col_e2 = st.columns(2)
        with col_e1:
            max_pages = st.number_input("Max Pages to Follow (Pagination):", min_value=1, max_value=100, value=3)
            next_selector = st.text_input("Next Page Button Selector:", value="button#next-page, a.next-page")
        with col_e2:
            clean_numbers = st.checkbox("Convert ₹ / commas to pure numbers", value=True)
            standardize_dates = st.checkbox("Standardize dates to DD/MM/YYYY", value=True)

        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("🔍 Extract Table from Live Browser", type="primary", use_container_width=True):
                if st.session_state.browser_mgr.page:
                    raw_df = extract_paginated_table_live(
                        page=st.session_state.browser_mgr.page,
                        next_button_selector=next_selector,
                        max_pages=max_pages,
                        delay_sec=action_delay,
                    )
                    st.session_state.extracted_df = clean_extracted_dataframe(
                        raw_df, clean_currency=clean_numbers, standardize_dates=standardize_dates
                    )
                    st.success(f"Extracted {len(st.session_state.extracted_df)} rows across up to {max_pages} page(s)!")
                else:
                    st.error("Please launch the browser first or click 'Extract from Local Demo Portal' on the right.")

        with col_btn2:
            if st.button("🧪 Extract Table from Local Demo Portal (`demo_portal.html`)", use_container_width=True):
                demo_path = "./samples/demo_portal.html"
                if os.path.exists(demo_path):
                    with open(demo_path, "r", encoding="utf-8") as f:
                        tables = detect_html_tables_from_string(f.read())
                    if tables:
                        raw_df = tables[0]
                        st.session_state.extracted_df = clean_extracted_dataframe(
                            raw_df, clean_currency=clean_numbers, standardize_dates=standardize_dates
                        )
                        st.success(f"Detected {len(tables)} table(s) — loaded {len(st.session_state.extracted_df)} rows!")

        if st.session_state.extracted_df is not None and not st.session_state.extracted_df.empty:
            st.markdown("#### 📊 Extracted Table Preview")
            st.dataframe(st.session_state.extracted_df, use_container_width=True)
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                st.session_state.extracted_df.to_excel(writer, index=False, sheet_name="Extracted_Data")
            st.download_button(
                "📥 Download Extracted Table (.xlsx)",
                data=buf.getvalue(),
                file_name="website_extracted_table.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # --- TAB 2: SEARCH-AND-FETCH FROM EXCEL ---
    with tab_search:
        st.subheader("🔎 Search-and-Fetch (Input Excel of IDs → Query Website → Save Status)")
        st.markdown(
            "Upload an Excel file with **Application No / Aadhaar / Roll No**, and the assistant will search each ID on the portal, extract the result fields, and write them back into your Excel file."
        )
        search_file = st.file_uploader("Upload Excel with IDs to Search:", type=["xlsx", "csv"], key="search_excel")
        if search_file:
            sheets = load_spreadsheet_sheets(search_file.read(), search_file.name)
            s_df = sheets[list(sheets.keys())[0]]
            st.session_state.active_df = s_df

        if st.session_state.active_df is not None:
            s_df = st.session_state.active_df
            id_col = st.selectbox("Select ID Column to Search:", list(s_df.columns))
            col_s1, col_s2 = st.columns(2)
            with col_s1:
                search_input_sel = st.text_input("Portal Search Box Selector:", value="#search-app-no")
                search_btn_sel = st.text_input("Portal Search Button Selector:", value="#btn-search-status")
            with col_s2:
                result_selectors_json = st.text_area(
                    "Result Field Selectors (JSON Column -> CSS Selector):",
                    value=json.dumps(
                        {
                            "Portal_Status": "#res-status",
                            "Approval_Date": "#res-date",
                            "Certificate_Amount": "#res-amount",
                            "Remarks": "#res-remarks",
                        },
                        indent=2,
                    ),
                    height=130,
                )

            if st.button("🚀 Run Search-and-Fetch Batch", type="primary"):
                if not st.session_state.browser_mgr.page:
                    st.error("Please launch the browser first so the assistant can search each ID.")
                else:
                    selectors_map = json.loads(result_selectors_json)
                    updated_df = search_and_fetch_from_excel_live(
                        page=st.session_state.browser_mgr.page,
                        df=s_df,
                        id_column=id_col,
                        search_input_selector=search_input_sel,
                        search_button_selector=search_btn_sel,
                        result_selectors=selectors_map,
                        delay_sec=action_delay,
                    )
                    st.session_state.active_df = updated_df
                    st.success("Search-and-Fetch complete!")
                    st.dataframe(updated_df, use_container_width=True)

                    buf = io.BytesIO()
                    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                        updated_df.to_excel(writer, index=False, sheet_name="Status_Fetched")
                    st.download_button(
                        "📥 Download Updated Excel with Fetched Statuses",
                        data=buf.getvalue(),
                        file_name="search_and_fetch_results.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

    # --- TAB 3: DETAIL-PAGE & FILE DOWNLOADER ---
    with tab_detail:
        st.subheader("📑 Detail-Page Scraper & Certificate/PDF Downloader")
        st.markdown(
            "When a portal table has a **'View Details'** or **'Download Certificate/Receipt'** link on each row, configure selectors below to open each detail view and save files named by ID."
        )
        row_link_sel = st.text_input("Row Detail Link Selector:", value="a.view-detail-link")
        download_link_sel = st.text_input("Download PDF/Receipt Selector (optional):", value="a.download-receipt-btn")
        downloads_folder = st.text_input("Save Downloaded Files To:", value="./downloads")

        if st.button("🚀 Extract Detail Pages & Download Files"):
            if not st.session_state.browser_mgr.page:
                st.error("Please launch the browser first.")
            else:
                detail_df = extract_detail_pages_live(
                    page=st.session_state.browser_mgr.page,
                    detail_link_selector=row_link_sel,
                    field_selectors={
                        "Application_No": ".detail-app-no",
                        "Applicant_Name": ".detail-name",
                        "Current_Status": ".detail-status",
                        "Verified_By": ".detail-officer",
                    },
                    download_selector=download_link_sel,
                    download_dir=downloads_folder,
                    delay_sec=action_delay,
                )
                st.dataframe(detail_df, use_container_width=True)


# ===========================================================================
# MODE 3: BOTH (FILL & VERIFY + RECONCILE) (PART C)
# ===========================================================================
else:
    st.subheader("🔄 Combined Mode: Fill & Verify + Excel vs. Website Reconciliation")
    st.markdown(
        "Compare your master Excel sheet with data extracted from the website to automatically flag **Approved / Pending / Rejected** changes and highlight field-level discrepancies."
    )

    col_c1, col_c2 = st.columns(2)
    with col_c1:
        excel_file_c = st.file_uploader("1️⃣ Upload Master Excel File:", type=["xlsx", "csv"], key="rec_excel")
    with col_c2:
        web_file_c = st.file_uploader("2️⃣ Upload Website Extracted Excel (or use extracted table):", type=["xlsx", "csv"], key="rec_web")

    if st.button("🧪 Load Sample Master vs. Portal Data for Reconciliation Demo"):
        sample_path = "./samples/sample_applicants.xlsx"
        if os.path.exists(sample_path):
            with open(sample_path, "rb") as f:
                sheets = load_spreadsheet_sheets(f.read(), "sample_applicants.xlsx")
            master_df = sheets[list(sheets.keys())[0]].copy()
            web_df = master_df.copy()
            if len(web_df) >= 2:
                web_df.loc[0, "Status"] = "Approved"
                web_df.loc[1, "Amount"] = "5500"
            summary_df, diff_df = reconcile_excel_vs_website(
                excel_df=master_df,
                website_df=web_df,
                key_column=list(master_df.columns)[0],
            )
            st.session_state.reconcile_summary_df = summary_df
            st.session_state.reconcile_diff_df = diff_df
            st.success("Reconciliation completed on sample dataset!")

    if excel_file_c and web_file_c:
        s1 = load_spreadsheet_sheets(excel_file_c.read(), excel_file_c.name)
        s2 = load_spreadsheet_sheets(web_file_c.read(), web_file_c.name)
        df_master = s1[list(s1.keys())[0]]
        df_web = s2[list(s2.keys())[0]]
        common_cols = [c for c in df_master.columns if c in df_web.columns]
        key_col = st.selectbox("Select Primary Key Column (e.g., Application_No or Aadhaar):", common_cols or list(df_master.columns))
        if st.button("🔍 Run Reconciliation (Compare Excel vs. Website)", type="primary"):
            summary_df, diff_df = reconcile_excel_vs_website(df_master, df_web, key_column=key_col)
            st.session_state.reconcile_summary_df = summary_df
            st.session_state.reconcile_diff_df = diff_df

    if st.session_state.reconcile_summary_df is not None:
        st.markdown("#### 📋 Reconciliation Summary")
        st.dataframe(st.session_state.reconcile_summary_df, use_container_width=True)
        st.markdown("#### ⚠️ Field-Level Differences (`Differences` Sheet)")
        st.dataframe(st.session_state.reconcile_diff_df, use_container_width=True)

        rec_bytes = export_reconciliation_excel(
            summary_df=st.session_state.reconcile_summary_df,
            differences_df=st.session_state.reconcile_diff_df,
        )
        st.download_button(
            "📥 Download Color-Coded Reconciliation Workbook (`Differences` Sheet)",
            data=rec_bytes,
            file_name="reconciliation_differences.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
