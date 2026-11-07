import React, { useState } from 'react';
import JSZip from 'jszip';
import {
  FileCode,
  FolderArchive,
  Copy,
  Check,
  BookOpen,
  FileSpreadsheet,
  Sparkles,
  ShieldCheck
} from 'lucide-react';
import { GENERATED_PYTHON_FILES } from '../generatedPythonBundle';

export const PythonProjectHub: React.FC = () => {
  const [selectedFile, setSelectedFile] = useState<string>('app.py');
  const [copied, setCopied] = useState(false);
  const [isZipping, setIsZipping] = useState(false);
  const [activeTab, setActiveTab] = useState<'code' | 'guide'>('code');

  const fileData = GENERATED_PYTHON_FILES[selectedFile] || GENERATED_PYTHON_FILES['app.py'];

  const handleCopyCode = () => {
    navigator.clipboard.writeText(fileData.content);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleDownloadZip = async () => {
    setIsZipping(true);
    try {
      const link = document.createElement('a');
      link.href = '/excel_website_automation_assistant.zip';
      link.download = 'excel_website_automation_assistant.zip';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (err) {
      console.error('Fallback JSZip generation:', err);
      const zip = new JSZip();
      Object.entries(GENERATED_PYTHON_FILES).forEach(([filename, file]) => {
        zip.file(`excel_website_automation_assistant/${filename}`, file.content);
      });
      const zipBlob = await zip.generateAsync({ type: 'blob' });
      const url = URL.createObjectURL(zipBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'excel_website_automation_assistant.zip';
      a.click();
      URL.revokeObjectURL(url);
    } finally {
      setTimeout(() => setIsZipping(false), 500);
    }
  };

  return (
    <div className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden shadow-xl">
      {/* Top Banner with 1-Click ZIP Download */}
      <div className="px-6 py-4 bg-gradient-to-r from-slate-950 via-slate-900 to-slate-950 border-b border-slate-800 flex flex-wrap items-center justify-between gap-4">
        <div>
          <div className="flex items-center gap-2">
            <span className="w-2.5 h-2.5 rounded-full bg-emerald-400"></span>
            <h2 className="text-base font-semibold text-slate-100">
              Complete Python Project Bundle (`excel_website_automation_assistant.zip`)
            </h2>
          </div>
          <p className="text-xs text-slate-400 mt-0.5">
            All 13 Python modules (`app.py`, `importer.py`, `validator.py`, `browser.py`, `scanner.py`, `llm.py`, `uploader.py`, `filler.py`, `runner.py`, `extractor.py`, `reconcile.py`, `report.py`, `utils.py`), `requirements.txt`, `README.md`, `templates/` & `samples/`
          </p>
        </div>

        <div className="flex items-center gap-2.5">
          <button
            onClick={() => setActiveTab(t => (t === 'code' ? 'guide' : 'code'))}
            className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white border border-slate-700 rounded-lg text-xs font-medium transition-colors cursor-pointer"
          >
            <BookOpen className="w-3.5 h-3.5 text-emerald-400" />
            <span>{activeTab === 'code' ? 'View Setup & Hindi Guide' : 'View Full Python Code'}</span>
          </button>

          <button
            onClick={handleDownloadZip}
            disabled={isZipping}
            className="flex items-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-500 disabled:bg-slate-800 text-white rounded-lg text-xs font-semibold transition-all cursor-pointer shadow-md shadow-emerald-950/40"
          >
            {isZipping ? (
              <div className="w-4 h-4 border-2 border-white/20 border-t-white rounded-full animate-spin" />
            ) : (
              <FolderArchive className="w-4 h-4" />
            )}
            <span>Download Full Python .ZIP</span>
          </button>
        </div>
      </div>

      {activeTab === 'guide' ? (
        <div className="p-6 space-y-6 text-sm text-slate-300">
          {/* Quickstart Command Cards */}
          <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
            <div className="p-4 bg-slate-950/80 border border-slate-800 rounded-lg">
              <span className="text-xs font-semibold text-emerald-400">1. Install Python Packages</span>
              <pre className="mt-2 p-2.5 bg-slate-900 border border-slate-800 rounded text-xs font-mono text-slate-200 overflow-x-auto">
                pip install -r requirements.txt
              </pre>
            </div>

            <div className="p-4 bg-slate-950/80 border border-slate-800 rounded-lg">
              <span className="text-xs font-semibold text-emerald-400">2. Install Chromium</span>
              <pre className="mt-2 p-2.5 bg-slate-900 border border-slate-800 rounded text-xs font-mono text-slate-200 overflow-x-auto">
                playwright install chromium
              </pre>
            </div>

            <div className="p-4 bg-slate-950/80 border border-slate-800 rounded-lg">
              <span className="text-xs font-semibold text-emerald-400">3. Set Gemini Key</span>
              <pre className="mt-2 p-2.5 bg-slate-900 border border-slate-800 rounded text-xs font-mono text-slate-200 overflow-x-auto">
                export GEMINI_API_KEY="AIza..."
                export GEMINI_MODEL="gemini-2.5-flash"
              </pre>
            </div>

            <div className="p-4 bg-slate-950/80 border border-slate-800 rounded-lg">
              <span className="text-xs font-semibold text-emerald-400">4. Run Streamlit App</span>
              <pre className="mt-2 p-2.5 bg-slate-900 border border-slate-800 rounded text-xs font-mono text-slate-200 overflow-x-auto">
                streamlit run app.py
              </pre>
            </div>
          </div>

          {/* Safety & Privacy Rules */}
          <div className="p-5 bg-slate-950/60 border border-slate-800 rounded-xl space-y-2.5">
            <div className="flex items-center gap-2">
              <ShieldCheck className="w-4 h-4 text-emerald-400" />
              <h3 className="font-semibold text-slate-100 text-sm">Built-In Safety, Privacy & Human-in-the-Loop Rules</h3>
            </div>
            <ul className="list-disc pl-5 space-y-1 text-xs text-slate-300">
              <li><strong>No Login/OTP/CAPTCHA Bypass:</strong> Whenever the portal asks for Login, OTP, or CAPTCHA, the browser pauses with <em>"Please login / enter OTP / captcha in the browser, then click Continue"</em>.</li>
              <li><strong>Semi-Automatic Mode (Default):</strong> Fills all fields and uploads compressed photos, then stops before the Submit button so you can verify and click <strong>Approve & Submit</strong>.</li>
              <li><strong>PII Masking (`XXXXXXXX1234`):</strong> Aadhaar, Mobile, and Bank Account numbers are masked in logs and never sent in full to Gemini (`llm.py` sends only column headers, field labels, and 1–2 masked sample values).</li>
            </ul>
          </div>

          {/* Hindi Step-by-Step Guide */}
          <div className="p-5 bg-slate-950/60 border border-slate-800 rounded-xl space-y-3">
            <div className="flex items-center gap-2">
              <Sparkles className="w-4 h-4 text-amber-400" />
              <h3 className="font-semibold text-slate-100 text-sm">🇮🇳 हिंदी में स्टेप-बाय-स्टेप गाइड (दोनों मोड के उदाहरण)</h3>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs text-slate-300 leading-relaxed">
              <div className="p-4 bg-slate-900 border border-slate-800 rounded-lg space-y-1.5">
                <div className="font-semibold text-emerald-400">मोड A: Excel/Word से वेबसाइट फॉर्म भरना (SSO / eMitra Registration)</div>
                <ol className="list-decimal pl-4 space-y-1">
                  <li>ऊपर <strong>"Excel to Website"</strong> मोड चुनें और अपनी <code>sample_applicants.xlsx</code> या Word फाइल अपलोड करें।</li>
                  <li>ऐप अपने-आप गलत मोबाइल नंबर (10 अंक), आधार (12 अंक), पिन कोड (6 अंक) और खाली कॉलम लाल रंग में दिखा देगा—उन्हें टेबल में ही ठीक करें।</li>
                  <li>साइडबार से <strong>"Launch Browser"</strong> दबाएं। यदि पोर्टल पर लॉगिन/OTP/Captcha आए, तो ब्राउज़र में खुद लॉगिन करें और ऐप में <strong>Continue</strong> दबाएं।</li>
                  <li><strong>Teach Mode</strong> टैब में जाकर <em>"Scan Open Browser Page Fields"</em> या सेव किया हुआ <code>SSO_Citizen_Registration</code> टेम्पलेट चुनें। Gemini कॉलम को फॉर्म के खानों से जोड़ देगा।</li>
                  <li><strong>Semi-Auto Form Filler</strong> टैब में <em>"Fill Current Row in Browser"</em> दबाएं। ऐप जिला → ब्लॉक → गाँव चुनकर और फोटो को 50KB से कम करके अपलोड कर देगा और Submit से पहले रुक जाएगा।</li>
                  <li>जांच करने के बाद <strong>"Approve & Submit"</strong> दबाएं। एप्लीकेशन नंबर और स्क्रीनशॉट अपने-आप सेव हो जाएंगे और अपडेटेड एक्सेल डाउनलोड करें।</li>
                </ol>
              </div>
              <div className="p-4 bg-slate-900 border border-slate-800 rounded-lg space-y-1.5">
                <div className="font-semibold text-blue-400">मोड B: वेबसाइट से Excel में डेटा और स्टेटस निकालना (Website → Excel)</div>
                <ol className="list-decimal pl-4 space-y-1">
                  <li>ऊपर <strong>"Website to Excel"</strong> मोड चुनें।</li>
                  <li><strong>उदाहरण 1 (टेबल निकालना):</strong> <em>Table & List Extractor</em> टैब में जाएं, अधिकतम पेज (जैसे <code>max_pages = 3</code>) सेट करें और <strong>"Extract Table"</strong> दबाएं। सभी पेजों का डेटा साफ होकर (₹/कॉमा हटकर) Excel फाइल में डाउनलोड हो जाएगा।</li>
                  <li><strong>उदाहरण 2 (Search-and-Fetch द्वारा स्टेटस चेक करना):</strong> <em>Search-and-Fetch Status from Excel</em> टैब में अपनी एप्लीकेशन नंबर वाली Excel अपलोड करें।</li>
                  <li><code>Application_No</code> कॉलम चुनें और <strong>"Run Search-and-Fetch Batch"</strong> दबाएं। ऐप एक-एक करके हर एप्लीकेशन नंबर को पोर्टल पर सर्च करेगा, उसका <em>Status (Approved/Pending/Rejected), Date और Amount</em> पढ़ेगा और आपकी Excel फाइल में नए कॉलम जोड़कर सेव कर देगा।</li>
                  <li><strong>मोड C (Both - Fill & Verify):</strong> अपनी पुरानी Excel और वेबसाइट के ताज़ा डेटा की तुलना (Reconcile) करके रंगीन <code>Differences</code> शीट डाउनलोड करें।</li>
                </ol>
              </div>
            </div>
          </div>
        </div>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-12 divide-y md:divide-y-0 md:divide-x divide-slate-800">
          {/* File Tree Selector */}
          <div className="md:col-span-4 p-4 bg-slate-950/50 space-y-2">
            <div className="px-2 py-1 text-xs font-semibold uppercase tracking-wider text-slate-500">
              Project Files ({Object.keys(GENERATED_PYTHON_FILES).length})
            </div>
            <div className="space-y-1 max-h-[540px] overflow-y-auto pr-1">
              {Object.entries(GENERATED_PYTHON_FILES).map(([key, file]) => {
                const isSelected = selectedFile === key;
                return (
                  <button
                    key={key}
                    onClick={() => setSelectedFile(key)}
                    className={`w-full text-left px-3 py-2 rounded-lg text-xs font-mono transition-all flex items-start gap-2.5 cursor-pointer ${
                      isSelected
                        ? 'bg-slate-800 text-emerald-300 font-semibold border border-slate-700 shadow-sm'
                        : 'text-slate-400 hover:text-slate-200 hover:bg-slate-900/60'
                    }`}
                  >
                    <FileCode className={`w-4 h-4 mt-0.5 shrink-0 ${isSelected ? 'text-emerald-400' : 'text-slate-500'}`} />
                    <div className="min-w-0">
                      <div className="truncate">{file.name}</div>
                      <div className="text-[11px] text-slate-500 font-sans truncate mt-0.5">
                        {file.description}
                      </div>
                    </div>
                  </button>
                );
              })}
            </div>

            <div className="pt-3 border-t border-slate-800/80">
              <div className="px-2 py-1 text-xs font-semibold uppercase tracking-wider text-slate-500">
                Bundled Binary Samples in ZIP
              </div>
              <div className="space-y-1 mt-1 text-xs text-slate-400 pl-2">
                <div className="flex items-center gap-2 py-0.5">
                  <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-400" />
                  <span>samples/sample_applicants.xlsx</span>
                </div>
                <div className="flex items-center gap-2 py-0.5">
                  <FileSpreadsheet className="w-3.5 h-3.5 text-blue-400" />
                  <span>samples/sample_applicants.docx</span>
                </div>
                <div className="flex items-center gap-2 py-0.5">
                  <FileSpreadsheet className="w-3.5 h-3.5 text-amber-400" />
                  <span>samples/photos/*.jpg (3 Applicant Photos)</span>
                </div>
              </div>
            </div>
          </div>

          {/* Code Viewer */}
          <div className="md:col-span-8 flex flex-col min-w-0">
            <div className="px-5 py-2.5 bg-slate-950 border-b border-slate-800 flex items-center justify-between">
              <div className="flex items-center gap-2 text-xs font-mono text-slate-300">
                <FileCode className="w-4 h-4 text-emerald-400" />
                <span>{fileData.name}</span>
                <span className="text-slate-500 font-sans">({fileData.description})</span>
              </div>

              <button
                onClick={handleCopyCode}
                className="flex items-center gap-1.5 px-2.5 py-1 bg-slate-900 hover:bg-slate-800 text-slate-300 hover:text-white border border-slate-800 rounded text-xs transition-colors cursor-pointer shrink-0"
              >
                {copied ? <Check className="w-3.5 h-3.5 text-emerald-400" /> : <Copy className="w-3.5 h-3.5" />}
                <span>{copied ? 'Copied!' : 'Copy Full File'}</span>
              </button>
            </div>

            <pre className="p-5 text-xs font-mono text-slate-300 overflow-x-auto max-h-[600px] bg-slate-950/90 leading-relaxed selection:bg-emerald-950 selection:text-emerald-200">
              <code>{fileData.content}</code>
            </pre>
          </div>
        </div>
      )}
    </div>
  );
};
