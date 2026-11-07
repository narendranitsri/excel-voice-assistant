# 🎙️ Voice-Controlled Excel Assistant (स्मार्ट वॉइस एक्सेल असिस्टेंट)

A complete, production-grade, voice-controlled Excel automation copilot written in Python 3.10+ using **Streamlit**, **pandas**, **openpyxl**, **rapidfuzz**, **Plotly**, **pdfplumber**, and **Anthropic Claude**.

Supports voice and text instructions in **Hindi, English, and Hinglish**. The app translates natural instructions into structured JSON action plans, executes safe predefined pandas/openpyxl operations (never arbitrary `exec`/`eval`), writes dynamic Excel formulas, provides automated data quality audits, embeds native charts, and speaks voice summaries back to you.

---

## 🌟 Complete Architecture & Features (Phases 1 - 4)

### Phase 1: Core Operations
1. **Multi-File & Multi-Sheet Support**: Upload multiple `.xlsx`, `.xls`, or `.csv` files.
2. **VLOOKUP / XLOOKUP**: Left join matching with space trimming, case-insensitivity, and numeric-text coercion, plus fuzzy matching (`rapidfuzz`).
3. **Multi-Column Sorting**: Ascending (`A-Z`, `0-9`) or descending (`Z-A`, `9-0`).
4. **Common & Unique Items**: Find common records, records unique to Table A or B, and duplicates across tables.
5. **Filtering**: `equals`, `contains`, `>`, `<`, `between`, and date ranges.
6. **Data Cleaning**: Deduplication, drop/fill blanks, trim spaces, transform text casing (Proper Case, UPPER, lower).
7. **Pivot Tables & GroupBy**: Summary tables with `sum`, `count`, `average`, `min`, `max`.
8. **Calculated Columns**: Safely evaluated formulas (e.g. `Total = Quantity * Unit_Price`), column merging and splitting.
9. **Sheet & Table Merging**: Concatenate or append multiple sheets with origin tracking.
10. **Conditional Formatting**: Professional dark navy headers and soft yellow/green highlights on duplicates or matched values.
11. **Undo / Step History**: Revert actions with full state restoration.

### Phase 2: Live Formulas, Multi-Step & Macros
12. **Live Dynamic Excel Formulas**: Writes real `=XLOOKUP()`, `=SUMIFS()`, `=COUNTIFS()`, `=IFERROR()` formulas into output cells so the file stays dynamic. Includes a toggle for "Write formulas" vs "Write values only" and educational explanations.
14. **Multi-Step Commands**: One compound instruction (e.g. *"Duplicate hatao, Price se sort karo, phir city wise pivot banao"*) becomes an ordered checklist of actions executed step-by-step.
17. **Save as Macro (Workflow Recorder)**: Save any sequence of actions as a named JSON workflow (e.g. `Monthly Sales Cleanup`) to run on new files in one click.
18. **Dry Run + Confirm**: Before destructive actions (delete rows, remove duplicates), previews what will change with red highlights and requests confirmation.

### Phase 3: Analytics, Q&A & Standardization
13. **Ask Your Data (Q&A Mode)**: Questions like *"Sabse zyada sale kis city mein hui?"* or *"Top 5 customers kaun hain?"* are answered directly with text and small summary tables.
15. **Data Quality Report (One-Click)**: Comprehensive health audit (blanks, duplicates, mixed types, extra spaces, outliers, invalid emails/phones) with a 1-click **"Fix all automatically"** button.
16. **Auto Charts & Quick Dashboard**: Create custom Plotly charts or generate a 4-metric visual dashboard with 1 click.
20. **Smart Standardization**: Standardize dates (`DD-MM-YYYY`), phone numbers (`+91 XXXXX XXXXX`), names (`Title Case`), and city spelling variants via fuzzy matching (e.g. `"Delhi"`, `"delhi "`, `"Dilli"` ➔ `"Delhi"`).
22. **Anomaly / Outlier Detection**: *"Galat ya ajeeb values dhundo"* detects extreme outliers, negative rates, and data-entry mistakes.
26. **Executive Summary Sheet**: *"Report banao"* creates a one-page executive summary sheet with key metrics, top/bottom performers, and insights in the exported Excel file.

### Phase 4: Vision, Voice & Versioning
19. **Compare Two Versions**: Upload old and new versions of a file to view added, removed, and modified cells, and export a color-coded "Changes" sheet.
21. **Table Extraction from PDF & Images**: Extract tables from PDF documents (`pdfplumber`) or photos/invoices using Claude Vision.
23. **Voice Reply (Text-to-Speech)**: Speaks a short spoken summary in Hindi or English after each operation (with on/off toggle).
24. **Excel-Learning Mode**: *"Ye kaise kiya?"* explains the manual Excel steps and equivalent formulas in simple Hindi & English.
25. **Extra Output Options**: Freeze top header row, auto column widths, and export to `.xlsx` or `.csv`.

---

## 📁 Project Structure

```
├── app.py                   # Streamlit frontend & interactive voice/text controller
├── actions.py               # Predefined pandas, openpyxl & rapidfuzz execution engine
├── llm_parser.py            # Anthropic Claude JSON action planner (configurable CLAUDE_MODEL)
├── formulas.py              # Dynamic Excel formulas generator (=XLOOKUP, =SUMIFS)
├── quality.py               # Data quality audit and smart standardization engine
├── charts.py                # Plotly dashboards & openpyxl native chart embedding
├── workflows.py             # Macro recorder and workflow automation engine
├── utils.py                 # PDF/image extraction, gTTS voice reply, summary sheet builder
├── generate_samples.py      # Script to generate the test datasets
├── requirements.txt         # Complete Python dependencies
├── README.md                # Documentation and beginner setup guide
└── samples/                 # 4 sample Excel files
    ├── sample_orders.xlsx
    ├── sample_prices.xlsx
    ├── sample_inventory.xlsx
    └── sample_messy_data.xlsx
```

---

## 🚀 Beginner Setup Guide (English)

### Step 1: Install Python 3.10+
- Download Python 3.10 or higher from [python.org](https://www.python.org/downloads/).
- **Important:** Check the box **"Add Python to PATH"** during setup.

### Step 2: Open Terminal / Command Prompt and Navigate to the Folder
```bash
cd "path/to/voice-controlled-excel-assistant"
```

### Step 3: (Recommended) Create a Virtual Environment
```bash
python -m venv venv

# Windows:
venv\Scripts\activate

# macOS / Linux:
source venv/bin/activate
```

### Step 4: Install Dependencies
```bash
pip install -r requirements.txt
```

### Step 5: Set Your API Key (Google Gemini or Claude)

**Option A: Using Google Gemini (Recommended - Fast & Powerful)**
- **Windows Command Prompt (CMD):**
  ```cmd
  set GEMINI_API_KEY=your_actual_gemini_api_key_here
  ```
- **Windows PowerShell:**
  ```powershell
  $env:GEMINI_API_KEY="your_actual_gemini_api_key_here"
  ```
- **macOS / Linux:**
  ```bash
  export GEMINI_API_KEY="your_actual_gemini_api_key_here"
  ```

**Option B: Using Anthropic Claude**
- **Windows CMD:**
  ```cmd
  set ANTHROPIC_API_KEY=your_actual_anthropic_api_key_here
  set CLAUDE_MODEL=claude-sonnet-5-5
  ```
- **Windows PowerShell:**
  ```powershell
  $env:ANTHROPIC_API_KEY="your_actual_anthropic_api_key_here"
  $env:CLAUDE_MODEL="claude-sonnet-5-5"
  ```
- **macOS / Linux:**
  ```bash
  export ANTHROPIC_API_KEY="your_actual_anthropic_api_key_here"
  export CLAUDE_MODEL="claude-sonnet-5-5"
  ```
*(Note: You can also enter your Gemini or Claude key directly in the sidebar inside the app!)*

### Step 6: Start the Application
```bash
streamlit run app.py
```
Streamlit will automatically open your default web browser at `http://localhost:8501`.

---

## 🗣️ 10 Example Voice Commands You Can Try

1. **VLOOKUP / XLOOKUP:**
   > *"Orders aur Prices me Item se VLOOKUP lagao aur Unit_Price le aao"*
2. **Multi-Step Pipeline:**
   > *"Duplicate hatao, Quantity se sort karo, phir city wise pivot banao"*
3. **Common Items Discovery:**
   > *"Orders aur Inventory me kaun se product common hain wo nikalo"*
4. **Calculated Columns:**
   > *"Add calculated column Total = Quantity * Unit_Price"*
5. **Pivot Table:**
   > *"City ke hisab se Quantity ka total sum pivot banao"*
6. **Ask Your Data (Q&A):**
   > *"Sabse zyada sale kis city mein hui?"*
7. **Anomaly / Outlier Detection:**
   > *"Galat ya ajeeb values dhundo"*
8. **Auto Charts:**
   > *"Sales ka bar chart banao"*
9. **Sorting:**
   > *"Quantity ko bade se chota (descending) sort karo"*
10. **Excel Learning Mode:**
    > *"Ye kaise kiya?"*

---

## 🇮🇳 हिंदी में सेटअप और चलाने की आसान गाइड (Hindi Guide)

अगर आप पहली बार Python इस्तेमाल कर रहे हैं, तो इन आसान स्टेप्स को फॉलो करें:

### १. Python इंस्टॉल करें
1. [python.org](https://www.python.org/downloads/) पर जाएं और Python 3.10 या 3.11 डाउनलोड करें।
2. इंस्टॉल करते समय **"Add Python to PATH"** पर टिक ज़रूर लगाएं।

### २. टर्मिनल या कमांड प्रॉम्प्ट में जाएं
प्रोजेक्ट फोल्डर में जाकर टर्मिनल खोलें:
```bash
cd <aapke_folder_ka_path>
```

### ३. ज़रूरी लाइब्रेरीज़ इंस्टॉल करें
टर्मिनल में लिखें:
```bash
pip install -r requirements.txt
```

### ४. अपनी API Key (Gemini या Claude) सेट करें
- **Google Gemini के लिए (बेस्ट और आसान):**
  - **Windows (CMD):**
    ```cmd
    set GEMINI_API_KEY=AIzaSy...
    ```
  - **Mac / Linux:**
    ```bash
    export GEMINI_API_KEY=AIzaSy...
    ```
- **Anthropic Claude के लिए:**
  - **Windows (CMD):**
    ```cmd
    set ANTHROPIC_API_KEY=sk-ant-api03-...
    ```
  - **Mac / Linux:**
    ```bash
    export ANTHROPIC_API_KEY=sk-ant-api03-...
    ```
*(नोट: यदि आप टर्मिनल में सेट नहीं करना चाहते, तो आप सीधे ऐप के साइडबार में "Google Gemini" चुनकर अपनी Key पेस्ट कर सकते हैं)*

### ५. ऐप शुरू करें
यह कमांड रन करें:
```bash
streamlit run app.py
```
ब्राउज़र में ऐप खुल जाएगा!

### ६. इस्तेमाल करने का तरीका:
1. **"Load 4 Sample Tables"** पर क्लिक करें।
2. माइक का बटन दबाएं और हिंदी या हिंग्लिश में बोलें (जैसे: *"Orders aur Prices me VLOOKUP lagao aur Unit_Price le aao"* या *"Duplicate hatao, Price se sort karo"*).
3. बोला हुआ टेक्स्ट बॉक्स में दिखेगा, ज़रूरत हो तो एडिट करें और **"Execute Action"** दबाएं।
4. ऐप एक्शन पूरा करके आवाज़ में उत्तर देगा, और आप नई एक्सेल फाइल तुरंत डाउनलोड कर सकते हैं!

---

## ⚖️ License
Apache-2.0 License.
