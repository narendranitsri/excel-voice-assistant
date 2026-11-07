# 🔄 Excel ↔ Website Automation Assistant (Python + Streamlit + Playwright + Gemini)

A complete, human-in-the-loop **Two-Way Automation Assistant** in Python that moves data between **Excel/Word** and **Government / Enterprise Web Portals** (such as SSO Rajasthan, Raj eMitra, scholarship/pension portals, or internal web applications):

1. **Part A (`Excel/Word → Website`)**: Read rows from Excel (`.xlsx`/`.csv`) or Word (`.docx`), validate mandatory fields (10-digit Mobile, 12-digit Aadhaar, 6-digit PIN, `DD/MM/YYYY` dates, duplicates), scan web forms in **Teach Mode**, map columns using Google Gemini (`llm.py`), fill forms row-by-row in a visible Chromium browser (including cascading `District → Block → Village` dropdowns and auto-compressed photo/PDF uploads), pause before Submit for your review (**Semi-Automatic Mode**), capture proof screenshots, and write the Application/Acknowledgement No. back into a copy of your Excel file.
2. **Part B (`Website → Excel`)**: Extract HTML tables & lists across multiple pages (`max_pages` pagination), run **Search-and-Fetch** (looping over an Excel list of Application IDs/Aadhaar numbers, searching each on the portal, and saving the status/dates/remarks back into Excel), scrape Detail Pages, download PDF certificates/receipts named by ID, and clean numbers/dates automatically.
3. **Part C (`Both: Fill & Verify + Reconcile`)**: Re-check live portal statuses for previously submitted rows (`Submitted` → `Approved`/`Pending`/`Rejected`) and reconcile your master Excel sheet against website data, generating a color-coded **`Differences`** Excel workbook.

---

## 📁 Project File Structure

```text
.
├── app.py                  # Main Streamlit UI (Mode Switch: Excel->Web, Web->Excel, Both Fill & Verify)
├── importer.py             # Excel (.xlsx/.csv) multi-sheet & Word (.docx table + "Label: Value") loader
├── validator.py            # Pre-flight validation (Mobile, Aadhaar, PIN, Dates, Blanks, Duplicates)
├── browser.py              # Playwright Headed Chromium manager, persistent profile, Login/OTP/CAPTCHA pause
├── scanner.py              # Teach Mode DOM & iframe form field scanner
├── llm.py                  # Centralized Google Gemini API (google-genai) JSON mapping & retry handler
├── uploader.py             # Photo/document matcher & Pillow iterative KB/dimension compressor
├── filler.py               # Playwright form-filling engine (cascading dropdowns, Unicode Hindi, uploads)
├── runner.py               # Batch execution state, checkpointing, Semi-Auto vs Auto-Submit workflow
├── extractor.py            # Part B: HTML table detection, pagination, Search-and-Fetch, Detail-page & PDF saver
├── reconcile.py            # Part C: Excel vs Website reconciliation & color-coded "Differences" workbook
├── report.py               # Non-destructive Excel result writer & fpdf2 PDF audit report generator
├── utils.py                # PII masking (XXXXXXXX1234), value normalizers, Hindi/English voice command parser
├── requirements.txt        # Python dependencies
├── README.md               # Setup guide, Privacy & Terms notice, and Hindi step-by-step walkthrough
├── templates/              # Saved portal JSON templates (reusable without calling Gemini again)
│   ├── SSO_Citizen_Registration.json
│   └── SSO_Status_Search_Extract.json
└── samples/                # Bundled test files & local demo portal
    ├── demo_portal.html    # Local SSO/eMitra test portal (Registration Form + Status Search + Paginated Table)
    ├── sample_applicants.xlsx
    ├── sample_applicants.docx
    └── photos/
        ├── 101_Ramesh_Kumar.jpg
        ├── 102_Sunita_Devi.jpg
        └── 103_Mahesh_Sharma.jpg
```

---

## 🚀 Beginner-Friendly Setup Instructions

### 1. Install Python 3.10+
Make sure Python 3.10 or newer is installed on your computer:
```bash
python --version
```

### 2. Install Required Packages
From the project root folder, run:
```bash
pip install -r requirements.txt
```

### 3. Install Playwright Chromium Browser
Download the Chromium browser binary used for headed automation:
```bash
playwright install chromium
```

### 4. Configure Your Google Gemini API Key (`GEMINI_API_KEY`)
All AI calls live inside `llm.py` using the official `google-genai` SDK. The app checks `st.secrets["GEMINI_API_KEY"]` first, and then the environment variable `GEMINI_API_KEY`.

**Option A — Environment Variable (Windows PowerShell):**
```powershell
$env:GEMINI_API_KEY="your_gemini_api_key_here"
$env:GEMINI_MODEL="gemini-2.5-flash"
```

**Option B — Environment Variable (Linux / macOS):**
```bash
export GEMINI_API_KEY="your_gemini_api_key_here"
export GEMINI_MODEL="gemini-2.5-flash"
```

**Option C — Streamlit Secrets (`.streamlit/secrets.toml`):**
```toml
GEMINI_API_KEY = "your_gemini_api_key_here"
```

### 5. Run the Application
```bash
streamlit run app.py
```

---

## 🔒 Privacy, Security & Terms of Use Notice

1. **Human-in-the-Loop for Login, OTP & CAPTCHA**:
   - This application **never** attempts to bypass, solve, or automate Login, OTP, or CAPTCHA challenges.
   - Whenever a login screen, OTP box, or CAPTCHA is detected, execution pauses immediately with the message:
     > *"Please login / enter OTP / captcha in the browser, then click Continue"*
2. **Zero Password & OTP Storage**:
   - Your portal credentials and OTPs are entered directly by you in the persistent Chromium browser window (`./browser_profile`) and are never read or saved by the app.
3. **PII Masking (`XXXXXXXX1234`)**:
   - Sensitive columns (Aadhaar, Jan-Aadhaar, Mobile Number, Bank Account) are automatically masked in UI logs and error reports.
   - When calling Google Gemini (`llm.py`) to map columns to form fields, **only** column headers, form labels, and 1–2 masked sample values are sent—never your full spreadsheet or raw personal numbers.
4. **Rate Limiting & Website Terms of Use**:
   - A configurable delay (default `1–2 seconds` between actions/pages) prevents overloading portal servers.
   - Only use this tool on websites and portals where you are authorized to enter or retrieve data in accordance with their Terms of Use.
