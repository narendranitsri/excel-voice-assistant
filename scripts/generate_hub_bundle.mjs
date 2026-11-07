import fs from 'fs';

const filesMeta = [
  { key: 'app.py', path: 'app.py', lang: 'python', desc: 'Main Streamlit app with 3-mode switch: Excel->Website, Website->Excel, Both (Fill & Verify).' },
  { key: 'importer.py', path: 'importer.py', lang: 'python', desc: 'Loads Excel (.xlsx/.csv) sheets and Word (.docx) tables + "Label: Value" paragraphs.' },
  { key: 'validator.py', path: 'validator.py', lang: 'python', desc: 'Pre-flight row validator: 10-digit mobile, 12-digit Aadhaar, 6-digit PIN, DD/MM/YYYY dates, duplicates.' },
  { key: 'browser.py', path: 'browser.py', lang: 'python', desc: 'Playwright Chromium manager (headed mode, persistent profile, Login/OTP/CAPTCHA pause, Ack No parser).' },
  { key: 'scanner.py', path: 'scanner.py', lang: 'python', desc: 'Teach Mode DOM & iframe form field scanner.' },
  { key: 'llm.py', path: 'llm.py', lang: 'python', desc: 'Centralized Google Gemini API (google-genai) JSON mapping & extraction helper with 3x retry & PII masking.' },
  { key: 'uploader.py', path: 'uploader.py', lang: 'python', desc: 'Photo/document matcher & Pillow iterative image compressor (<= 50 KB / target dimensions).' },
  { key: 'filler.py', path: 'filler.py', lang: 'python', desc: 'Playwright form-filling engine with cascading dropdown wait logic, Unicode Hindi support, and rate limiting.' },
  { key: 'runner.py', path: 'runner.py', lang: 'python', desc: 'Batch execution state, checkpointing, proof screenshots, and Semi-Auto vs Auto-Submit control.' },
  { key: 'extractor.py', path: 'extractor.py', lang: 'python', desc: 'Part B: HTML table detection, multi-page pagination, Search-and-Fetch by ID from Excel, Detail-page & PDF saver.' },
  { key: 'reconcile.py', path: 'reconcile.py', lang: 'python', desc: 'Part C: Excel vs Website reconciliation and color-coded "Differences" Excel workbook writer.' },
  { key: 'report.py', path: 'report.py', lang: 'python', desc: 'Non-destructive Excel result writer and fpdf2 PDF audit summary generator.' },
  { key: 'utils.py', path: 'utils.py', lang: 'python', desc: 'PII masking (XXXXXXXX1234), value normalizers, and bilingual Hindi/English voice command parser.' },
  { key: 'requirements.txt', path: 'requirements.txt', lang: 'text', desc: 'Python dependencies: streamlit, playwright, pandas, openpyxl, python-docx, beautifulsoup4, google-genai, Pillow, fpdf2.' },
  { key: 'README.md', path: 'README.md', lang: 'markdown', desc: 'Setup instructions, Playwright install guide, GEMINI_API_KEY configuration, and Privacy Notice.' },
  { key: 'samples/demo_portal.html', path: 'samples/demo_portal.html', lang: 'html', desc: 'Local SSO/eMitra test portal with Registration Form, Application Status Search, and Paginated Table.' },
  { key: 'templates/SSO_Citizen_Registration.json', path: 'templates/SSO_Citizen_Registration.json', lang: 'json', desc: 'Saved Part A template mapping Excel columns to SSO registration fields & cascading dropdowns.' },
  { key: 'templates/SSO_Status_Search_Extract.json', path: 'templates/SSO_Status_Search_Extract.json', lang: 'json', desc: 'Saved Part B template for Search-and-Fetch status extraction and multi-page table scraping.' },
];

const bundleObj = {};
for (const item of filesMeta) {
  const content = fs.readFileSync(item.path, 'utf-8');
  bundleObj[item.key] = {
    name: item.key,
    language: item.lang,
    description: item.desc,
    content
  };
}

const tsContent = `// Auto-generated full source bundle for PythonProjectHub
export interface BundledProjectFile {
  name: string;
  language: string;
  description: string;
  content: string;
}

export const GENERATED_PYTHON_FILES: Record<string, BundledProjectFile> = ${JSON.stringify(bundleObj, null, 2)};
`;

fs.writeFileSync('src/generatedPythonBundle.ts', tsContent, 'utf-8');
console.log('Generated src/generatedPythonBundle.ts with', Object.keys(bundleObj).length, 'complete files.');
