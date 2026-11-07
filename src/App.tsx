import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import {
  FileSpreadsheet,
  Globe,
  ArrowLeftRight,
  FileCode,
  Download,
  CheckCircle2,
  AlertTriangle,
  Play,
  SkipForward,
  PauseCircle,
  Search,
  ShieldCheck,
  Sparkles,
  Eye,
  EyeOff,
  Mic
} from 'lucide-react';
import { PythonProjectHub } from './components/PythonProjectHub';

interface ApplicantRow {
  Serial_No: string;
  Application_No: string;
  Applicant_Name: string;
  Father_Name: string;
  DOB: string;
  Gender: string;
  Mobile: string;
  Aadhaar: string;
  Category: string;
  PIN_Code: string;
  District: string;
  Block: string;
  Village: string;
  Annual_Income: number;
  Photo_File: string;
  Status: string;
  Portal_Ref_No?: string;
  Portal_Amount?: string;
  Portal_Date?: string;
  Portal_Remarks?: string;
}

const INITIAL_APPLICANTS: ApplicantRow[] = [
  {
    Serial_No: '101',
    Application_No: 'RAJ-2025-101',
    Applicant_Name: 'Ramesh Kumar',
    Father_Name: 'Mohan Lal',
    DOB: '15/08/1992',
    Gender: 'Male',
    Mobile: '9876543210',
    Aadhaar: '456789012345',
    Category: 'OBC',
    PIN_Code: '302029',
    District: 'Jaipur',
    Block: 'Sanganer',
    Village: 'Muhana',
    Annual_Income: 45000,
    Photo_File: '101_Ramesh_Kumar.jpg (Compressed: 38 KB)',
    Status: 'Pending Fill',
    Portal_Amount: '12500'
  },
  {
    Serial_No: '102',
    Application_No: 'RAJ-2025-102',
    Applicant_Name: 'Sunita Devi',
    Father_Name: 'Sohan Lal',
    DOB: '20/11/1995',
    Gender: 'Female',
    Mobile: '9123456789',
    Aadhaar: '789012345678',
    Category: 'General',
    PIN_Code: '342001',
    District: 'Jodhpur',
    Block: 'Mandore',
    Village: 'Boranada',
    Annual_Income: 38000,
    Photo_File: '102_Sunita_Devi.jpg (Compressed: 41 KB)',
    Status: 'Pending Fill',
    Portal_Amount: '15000'
  },
  {
    Serial_No: '103',
    Application_No: 'RAJ-2025-103',
    Applicant_Name: 'Mahesh Sharma',
    Father_Name: 'Kailash Sharma',
    DOB: '05/03/1989',
    Gender: 'Male',
    Mobile: '9988776655',
    Aadhaar: '654321098765',
    Category: 'General',
    PIN_Code: '313001',
    District: 'Udaipur',
    Block: 'Girwa',
    Village: 'Bedla',
    Annual_Income: 52000,
    Photo_File: '103_Mahesh_Sharma.jpg (Compressed: 36 KB)',
    Status: 'Pending Fill',
    Portal_Amount: '10000'
  },
  {
    Serial_No: '104',
    Application_No: 'RAJ-2025-104',
    Applicant_Name: 'Kavita Meena',
    Father_Name: 'Ramkishan Meena',
    DOB: '01/12/1997',
    Gender: 'Female',
    Mobile: '98765', // Intentional validation error (5 digits)
    Aadhaar: '12345678', // Intentional validation error (8 digits)
    Category: 'ST',
    PIN_Code: '3020', // Intentional validation error (4 digits)
    District: 'Jaipur',
    Block: 'Amer',
    Village: 'Achrol',
    Annual_Income: 40000,
    Photo_File: '104_Kavita_Meena.jpg',
    Status: 'Validation Error',
    Portal_Amount: '12500'
  }
];

const PORTAL_LIVE_DATABASE: Record<string, { status: string; date: string; amount: string; remarks: string }> = {
  'RAJ-2025-101': { status: 'Approved', date: '14/02/2025', amount: '12500', remarks: 'Verified by Tehsildar Sanganer' },
  'RAJ-2025-102': { status: 'Approved', date: '18/02/2025', amount: '15000', remarks: 'Direct Benefit Transfer Completed' },
  'RAJ-2025-103': { status: 'Pending Verification', date: '21/02/2025', amount: '11500', remarks: 'Amount revised by Gram Sevak' },
  'RAJ-2025-104': { status: 'Rejected', date: '25/02/2025', amount: '0', remarks: 'Income certificate expired' }
};

function maskPii(val: string, type: 'aadhaar' | 'mobile'): string {
  const digits = String(val || '').replace(/\D/g, '');
  if (digits.length <= 4) return val;
  return 'X'.repeat(digits.length - 4) + digits.slice(-4);
}

export default function App() {
  const [activeMode, setActiveMode] = useState<'excel_to_web' | 'web_to_excel' | 'both_reconcile' | 'python_code'>('excel_to_web');
  const [rows, setRows] = useState<ApplicantRow[]>(INITIAL_APPLICANTS);
  const [currentRowIdx, setCurrentRowIdx] = useState<number>(0);
  const [maskSensitive, setMaskSensitive] = useState<boolean>(true);
  const [semiAutoMode, setSemiAutoMode] = useState<boolean>(true);
  const [filledInBrowser, setFilledInBrowser] = useState<boolean>(false);
  const [pausedForOtp, setPausedForOtp] = useState<boolean>(false);
  const [voiceCommand, setVoiceCommand] = useState<string>('');
  const [statusBanner, setStatusBanner] = useState<{ en: string; hi: string }>({
    en: 'Ready: Choose a mode above or download the complete 13-module Python project (.ZIP).',
    hi: 'तैयार: ऊपर से मोड चुनें या पूरा 13-फाइल वाला Python प्रोजेक्ट (.ZIP) डाउनलोड करें।'
  });

  const currentRow = rows[currentRowIdx] || rows[0];

  // Validate rows (mirrors validator.py)
  const getValidationIssues = (row: ApplicantRow) => {
    const issues: string[] = [];
    if (row.Mobile.replace(/\D/g, '').length !== 10) issues.push('Mobile must be 10 digits (मोबाइल 10 अंकों का होना चाहिए)');
    if (row.Aadhaar.replace(/\D/g, '').length !== 12) issues.push('Aadhaar must be 12 digits (आधार 12 अंकों का होना चाहिए)');
    if (row.PIN_Code.replace(/\D/g, '').length !== 6) issues.push('PIN Code must be 6 digits (पिन कोड 6 अंकों का होना चाहिए)');
    return issues;
  };

  const handleFixValidationIssues = () => {
    setRows(prev =>
      prev.map(r =>
        r.Serial_No === '104'
          ? { ...r, Mobile: '9876501234', Aadhaar: '890123456789', PIN_Code: '302028', Status: 'Pending Fill' }
          : r
      )
    );
    setStatusBanner({
      en: 'Fixed Row #4 validation errors (Mobile 10 digits, Aadhaar 12 digits, PIN 6 digits). Ready to fill!',
      hi: 'पंक्ति #4 की त्रुटियाँ (मोबाइल, आधार, पिन कोड) सुधार दी गई हैं!'
    });
  };

  // Part A: Fill Current Row in Simulated Headed Browser
  const handleFillRow = () => {
    if (pausedForOtp) return;
    const issues = getValidationIssues(currentRow);
    if (issues.length > 0) {
      setStatusBanner({
        en: `Validation Error on Row #${currentRowIdx + 1}: ${issues.join(' | ')}`,
        hi: `पंक्ति #${currentRowIdx + 1} में त्रुटि: पहले डेटा सुधारें।`
      });
      return;
    }
    setFilledInBrowser(true);
    setStatusBanner({
      en: `Filled Row #${currentRowIdx + 1} (${currentRow.Applicant_Name}) including cascading District → Block → Village & compressed photo. Stopped before Submit (Semi-Auto Mode).`,
      hi: `पंक्ति #${currentRowIdx + 1} (${currentRow.Applicant_Name}) फॉर्म में भर दी गई है (जिला → ब्लॉक → गाँव और फोटो सहित)। कृपया जांचें और 'Approve & Submit' दबाएं।`
    });
  };

  const handleApproveAndSubmit = () => {
    const ackNo = `ACK-2025-${8800 + currentRowIdx + 1}`;
    setRows(prev =>
      prev.map((r, i) => (i === currentRowIdx ? { ...r, Status: 'Submitted', Portal_Ref_No: ackNo } : r))
    );
    setFilledInBrowser(false);
    setStatusBanner({
      en: `Submitted Row #${currentRowIdx + 1}! Captured Acknowledgement No: ${ackNo} and saved proof screenshot.`,
      hi: `पंक्ति #${currentRowIdx + 1} सफलतापूर्वक सबमिट हो गई! एप्लीकेशन नंबर: ${ackNo} सेव कर लिया गया।`
    });
    if (currentRowIdx < rows.length - 1) {
      setCurrentRowIdx(currentRowIdx + 1);
    }
  };

  // Part B: Search-and-Fetch Status from Portal into Excel
  const handleRunSearchAndFetch = () => {
    setRows(prev =>
      prev.map(r => {
        const live = PORTAL_LIVE_DATABASE[r.Application_No];
        if (!live) return r;
        return {
          ...r,
          Status: live.status,
          Portal_Date: live.date,
          Portal_Amount: live.amount,
          Portal_Remarks: live.remarks
        };
      })
    );
    setStatusBanner({
      en: 'Search-and-Fetch completed! Queried all 4 Application IDs on the portal and wrote Status, Date, Amount, and Remarks into Excel.',
      hi: 'Search-and-Fetch पूरा हुआ! सभी 4 एप्लीकेशन नंबर पोर्टल पर सर्च करके उनका स्टेटस, तारीख और राशि एक्सेल में जोड़ दी गई है।'
    });
  };

  const handleExportCurrentExcel = (filename: string) => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, 'Automation_Results');
    XLSX.writeFile(wb, filename);
  };

  const handleVoiceCommandSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const cmd = voiceCommand.toLowerCase();
    if (cmd.includes('agli') || cmd.includes('next') || cmd.includes('bharo')) {
      handleFillRow();
    } else if (cmd.includes('submit') || cmd.includes('approve')) {
      handleApproveAndSubmit();
    } else if (cmd.includes('skip') || cmd.includes('chhodo')) {
      setCurrentRowIdx(i => Math.min(rows.length - 1, i + 1));
      setStatusBanner({ en: 'Skipped current row.', hi: 'वर्तमान पंक्ति छोड़ दी गई।' });
    } else if (cmd.includes('status') || cmd.includes('table') || cmd.includes('nikalo')) {
      setActiveMode('web_to_excel');
      handleRunSearchAndFetch();
    }
    setVoiceCommand('');
  };

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 flex flex-col font-sans">
      {/* Top Bar */}
      <header className="border-b border-slate-800 bg-slate-950/90 backdrop-blur-md sticky top-0 z-50 px-6 py-3.5 flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-center gap-2.5">
          <div className="w-8 h-8 rounded-lg bg-emerald-600 flex items-center justify-center text-white shadow-md shadow-emerald-950">
            <ArrowLeftRight className="w-4 h-4" />
          </div>
          <div>
            <span className="text-base font-bold tracking-tight text-white">
              Excel ↔ Website Automation Assistant
            </span>
            <span className="hidden lg:inline-block text-xs text-slate-400 font-mono border-l border-slate-800 pl-2.5 ml-2.5">
              Python 3.10+ · Streamlit · Playwright (Headed) · Google Gemini (`llm.py`)
            </span>
          </div>
        </div>

        {/* Top Mode Switch */}
        <nav className="flex flex-wrap items-center gap-1.5 bg-slate-900 p-1 rounded-xl border border-slate-800 text-xs font-medium">
          <button
            onClick={() => setActiveMode('excel_to_web')}
            className={`px-3 py-1.5 rounded-lg transition-all cursor-pointer flex items-center gap-1.5 ${
              activeMode === 'excel_to_web'
                ? 'bg-emerald-600 text-white font-semibold shadow-sm'
                : 'text-slate-400 hover:text-slate-200'
            }`}
          >
            <FileSpreadsheet className="w-3.5 h-3.5" />
            <span>1. Excel to Website</span>
          </button>

          <button
            onClick={() => setActiveMode('web_to_excel')}
            className={`px-3 py-1.5 rounded-lg transition-all cursor-pointer flex items-center gap-1.5 ${
              activeMode === 'web_to_excel'
                ? 'bg-blue-600 text-white font-semibold shadow-sm'
                : 'text-slate-400 hover:text-slate-200'
            }`}
          >
            <Globe className="w-3.5 h-3.5" />
            <span>2. Website to Excel</span>
          </button>

          <button
            onClick={() => setActiveMode('both_reconcile')}
            className={`px-3 py-1.5 rounded-lg transition-all cursor-pointer flex items-center gap-1.5 ${
              activeMode === 'both_reconcile'
                ? 'bg-purple-600 text-white font-semibold shadow-sm'
                : 'text-slate-400 hover:text-slate-200'
            }`}
          >
            <ArrowLeftRight className="w-3.5 h-3.5" />
            <span>3. Both (Fill & Verify)</span>
          </button>

          <button
            onClick={() => setActiveMode('python_code')}
            className={`px-3 py-1.5 rounded-lg transition-all cursor-pointer flex items-center gap-1.5 ${
              activeMode === 'python_code'
                ? 'bg-amber-500 text-slate-950 font-bold shadow-sm'
                : 'text-amber-300 hover:text-amber-200'
            }`}
          >
            <FileCode className="w-3.5 h-3.5" />
            <span>Python Code & Guide (18 Files)</span>
          </button>
        </nav>

        {/* Download ZIP Button */}
        <div className="flex items-center gap-2.5">
          <button
            onClick={() => setMaskSensitive(v => !v)}
            className="flex items-center gap-1.5 px-2.5 py-1.5 bg-slate-900 hover:bg-slate-800 border border-slate-800 rounded-lg text-xs text-slate-300 cursor-pointer"
            title="Toggle PII Masking (XXXXXXXX1234)"
          >
            {maskSensitive ? <EyeOff className="w-3.5 h-3.5 text-emerald-400" /> : <Eye className="w-3.5 h-3.5 text-amber-400" />}
            <span className="hidden sm:inline">{maskSensitive ? 'PII Masked' : 'PII Visible'}</span>
          </button>

          <a
            href="/excel_website_automation_assistant.zip"
            download="excel_website_automation_assistant.zip"
            className="flex items-center gap-1.5 px-3.5 py-1.5 bg-amber-500 hover:bg-amber-400 text-slate-950 rounded-lg text-xs font-bold transition-all cursor-pointer shadow-sm shadow-amber-950 whitespace-nowrap"
          >
            <Download className="w-3.5 h-3.5 text-slate-950" />
            <span>Download Full Python .ZIP</span>
          </a>
        </div>
      </header>

      <main className="flex-1 max-w-7xl w-full mx-auto p-4 md:p-6 space-y-5">
        {/* Status & Bilingual Voice Command Bar */}
        <div className="p-3.5 bg-slate-900/90 border border-slate-800 rounded-xl flex flex-wrap items-center justify-between gap-4">
          <div className="space-y-0.5">
            <div className="text-xs font-semibold text-emerald-400">{statusBanner.en}</div>
            <div className="text-xs text-slate-400">{statusBanner.hi}</div>
          </div>

          <form onSubmit={handleVoiceCommandSubmit} className="flex items-center gap-2">
            <div className="relative">
              <Mic className="w-3.5 h-3.5 text-emerald-400 absolute left-2.5 top-2.5" />
              <input
                type="text"
                value={voiceCommand}
                onChange={e => setVoiceCommand(e.target.value)}
                placeholder="Try: 'agli row bharo', 'status check karo'..."
                className="pl-8 pr-3 py-1.5 bg-slate-950 border border-slate-800 rounded-lg text-xs text-slate-200 w-64 focus:outline-none focus:border-emerald-500"
              />
            </div>
            <button
              type="submit"
              className="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-lg text-xs font-medium cursor-pointer"
            >
              Run Command
            </button>
          </form>
        </div>

        {activeMode === 'python_code' ? (
          <PythonProjectHub />
        ) : activeMode === 'excel_to_web' ? (
          /* =================================================================
             MODE 1: EXCEL/WORD -> WEBSITE (PART A)
             ================================================================= */
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
            {/* Left Column: Excel Data + Pre-flight Validator + Controls */}
            <div className="lg:col-span-7 space-y-4">
              <div className="p-4 bg-slate-900 border border-slate-800 rounded-xl space-y-3">
                <div className="flex items-center justify-between">
                  <div>
                    <h2 className="text-sm font-semibold text-white flex items-center gap-2">
                      <FileSpreadsheet className="w-4 h-4 text-emerald-400" />
                      <span>Part A: Input Excel / Word Data (`sample_applicants.xlsx`)</span>
                    </h2>
                    <p className="text-xs text-slate-400">
                      Validates 10-digit Mobile, 12-digit Aadhaar, 6-digit PIN & mandatory fields before filling.
                    </p>
                  </div>
                  <div className="flex items-center gap-2">
                    <button
                      onClick={handleFixValidationIssues}
                      className="px-2.5 py-1 bg-amber-500/20 hover:bg-amber-500/30 text-amber-300 border border-amber-500/40 rounded-lg text-xs font-medium cursor-pointer"
                    >
                      Fix Row #4 Errors
                    </button>
                    <button
                      onClick={() => handleExportCurrentExcel('applicants_results_updated.xlsx')}
                      className="px-2.5 py-1 bg-emerald-600/20 hover:bg-emerald-600/30 text-emerald-300 border border-emerald-500/40 rounded-lg text-xs font-medium cursor-pointer"
                    >
                      Export Updated .XLSX
                    </button>
                  </div>
                </div>

                <div className="overflow-x-auto border border-slate-800 rounded-lg">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-slate-950 text-slate-400 border-b border-slate-800">
                      <tr>
                        <th className="p-2.5">#</th>
                        <th className="p-2.5">Applicant_Name</th>
                        <th className="p-2.5">Mobile</th>
                        <th className="p-2.5">Aadhaar</th>
                        <th className="p-2.5">District → Block → Village</th>
                        <th className="p-2.5">Status / Ref No</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/70">
                      {rows.map((r, idx) => {
                        const issues = getValidationIssues(r);
                        const isCurrent = idx === currentRowIdx;
                        return (
                          <tr
                            key={r.Serial_No}
                            onClick={() => {
                              setCurrentRowIdx(idx);
                              setFilledInBrowser(false);
                            }}
                            className={`cursor-pointer transition-colors ${
                              isCurrent ? 'bg-emerald-950/30' : 'hover:bg-slate-950/60'
                            }`}
                          >
                            <td className="p-2.5 font-mono text-slate-400">{idx + 1}</td>
                            <td className="p-2.5 font-medium text-slate-200">{r.Applicant_Name}</td>
                            <td className={`p-2.5 font-mono ${r.Mobile.length !== 10 ? 'text-rose-400 font-bold' : 'text-slate-300'}`}>
                              {maskSensitive ? maskPii(r.Mobile, 'mobile') : r.Mobile}
                            </td>
                            <td className={`p-2.5 font-mono ${r.Aadhaar.length !== 12 ? 'text-rose-400 font-bold' : 'text-slate-300'}`}>
                              {maskSensitive ? maskPii(r.Aadhaar, 'aadhaar') : r.Aadhaar}
                            </td>
                            <td className="p-2.5 text-slate-300">
                              {r.District} → {r.Block} → {r.Village}
                            </td>
                            <td className="p-2.5">
                              {issues.length > 0 ? (
                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-rose-950/80 text-rose-300 border border-rose-800 text-[11px]">
                                  <AlertTriangle className="w-3 h-3" /> Invalid Data
                                </span>
                              ) : r.Status === 'Submitted' ? (
                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-emerald-950/80 text-emerald-300 border border-emerald-800 text-[11px] font-mono">
                                  <CheckCircle2 className="w-3 h-3" /> {r.Portal_Ref_No}
                                </span>
                              ) : (
                                <span className="px-2 py-0.5 rounded bg-slate-800 text-slate-300 text-[11px]">
                                  {r.Status}
                                </span>
                              )}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>

              {/* Teach Mode & Gemini Mapping Preview */}
              <div className="p-4 bg-slate-900 border border-slate-800 rounded-xl space-y-2.5">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2 text-xs font-semibold text-slate-200">
                    <Sparkles className="w-4 h-4 text-emerald-400" />
                    <span>Teach Mode & Gemini JSON Field Mapping (`templates/SSO_Citizen_Registration.json`)</span>
                  </div>
                  <span className="text-[11px] text-slate-400 font-mono">Model: gemini-2.5-flash</span>
                </div>
                <div className="grid grid-cols-2 sm:grid-cols-3 gap-2 text-[11px] font-mono">
                  <div className="p-2 bg-slate-950 border border-slate-800 rounded">Applicant_Name → #applicant_name</div>
                  <div className="p-2 bg-slate-950 border border-slate-800 rounded">District → #district (Parent)</div>
                  <div className="p-2 bg-slate-950 border border-slate-800 rounded">Block → #block (Wait AJAX)</div>
                  <div className="p-2 bg-slate-950 border border-slate-800 rounded">Village → #village (Wait AJAX)</div>
                  <div className="p-2 bg-slate-950 border border-slate-800 rounded">Photo_File → #applicant_photo (≤50KB)</div>
                  <div className="p-2 bg-slate-950 border border-slate-800 rounded">DOB → #dob (DD/MM/YYYY)</div>
                </div>
              </div>
            </div>

            {/* Right Column: Simulated Headed Chromium Browser (SSO / eMitra Demo Portal) */}
            <div className="lg:col-span-5 space-y-4">
              <div className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden shadow-lg">
                <div className="bg-slate-950 px-4 py-2.5 border-b border-slate-800 flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <span className="w-2.5 h-2.5 rounded-full bg-emerald-400"></span>
                    <span className="text-xs font-semibold text-slate-200">
                      Playwright Headed Browser (`samples/demo_portal.html`)
                    </span>
                  </div>
                  <button
                    onClick={() => setPausedForOtp(v => !v)}
                    className="px-2 py-0.5 bg-slate-800 hover:bg-slate-700 text-amber-300 rounded text-[11px] cursor-pointer"
                  >
                    {pausedForOtp ? 'Resume from OTP' : 'Simulate OTP/CAPTCHA Pause'}
                  </button>
                </div>

                {pausedForOtp ? (
                  <div className="p-6 bg-amber-950/30 border-b border-amber-800/60 space-y-3 text-center">
                    <PauseCircle className="w-10 h-10 text-amber-400 mx-auto" />
                    <h3 className="text-sm font-bold text-amber-200">
                      PAUSED FOR HUMAN INPUT (Login / OTP / CAPTCHA Detected)
                    </h3>
                    <p className="text-xs text-slate-300">
                      "Please login / enter OTP / captcha in the browser, then click Continue"
                    </p>
                    <button
                      onClick={() => setPausedForOtp(false)}
                      className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-semibold cursor-pointer"
                    >
                      Continue Automation (जारी रखें)
                    </button>
                  </div>
                ) : (
                  <div className="p-4 space-y-3 text-xs">
                    <div className="grid grid-cols-2 gap-2.5">
                      <div>
                        <label className="text-slate-400 block mb-1">Applicant Full Name *</label>
                        <input
                          readOnly
                          value={filledInBrowser ? currentRow.Applicant_Name : ''}
                          placeholder="Waiting for Fill..."
                          className="w-full px-2.5 py-1.5 bg-slate-950 border border-slate-800 rounded text-slate-100"
                        />
                      </div>
                      <div>
                        <label className="text-slate-400 block mb-1">Father / Husband Name *</label>
                        <input
                          readOnly
                          value={filledInBrowser ? currentRow.Father_Name : ''}
                          className="w-full px-2.5 py-1.5 bg-slate-950 border border-slate-800 rounded text-slate-100"
                        />
                      </div>
                      <div>
                        <label className="text-slate-400 block mb-1">Mobile (10 Digits) *</label>
                        <input
                          readOnly
                          value={
                            filledInBrowser
                              ? maskSensitive
                                ? maskPii(currentRow.Mobile, 'mobile')
                                : currentRow.Mobile
                              : ''
                          }
                          className="w-full px-2.5 py-1.5 bg-slate-950 border border-slate-800 rounded text-slate-100 font-mono"
                        />
                      </div>
                      <div>
                        <label className="text-slate-400 block mb-1">Aadhaar (12 Digits) *</label>
                        <input
                          readOnly
                          value={
                            filledInBrowser
                              ? maskSensitive
                                ? maskPii(currentRow.Aadhaar, 'aadhaar')
                                : currentRow.Aadhaar
                              : ''
                          }
                          className="w-full px-2.5 py-1.5 bg-slate-950 border border-slate-800 rounded text-slate-100 font-mono"
                        />
                      </div>
                      <div>
                        <label className="text-slate-400 block mb-1">District (जिला)</label>
                        <input
                          readOnly
                          value={filledInBrowser ? currentRow.District : ''}
                          className="w-full px-2.5 py-1.5 bg-slate-950 border border-slate-800 rounded text-emerald-300"
                        />
                      </div>
                      <div>
                        <label className="text-slate-400 block mb-1">Block → Village (Cascading)</label>
                        <input
                          readOnly
                          value={filledInBrowser ? `${currentRow.Block} → ${currentRow.Village}` : ''}
                          className="w-full px-2.5 py-1.5 bg-slate-950 border border-slate-800 rounded text-emerald-300"
                        />
                      </div>
                    </div>

                    <div>
                      <label className="text-slate-400 block mb-1">Applicant Photo Upload (`uploader.py` Auto-Resize ≤ 50KB)</label>
                      <input
                        readOnly
                        value={filledInBrowser ? currentRow.Photo_File : ''}
                        className="w-full px-2.5 py-1.5 bg-slate-950 border border-slate-800 rounded text-amber-300 font-mono"
                      />
                    </div>

                    {/* Semi-Auto Action Buttons */}
                    <div className="pt-2 border-t border-slate-800 flex flex-wrap items-center gap-2">
                      <button
                        onClick={handleFillRow}
                        className="flex-1 flex items-center justify-center gap-1.5 px-3 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-lg font-semibold cursor-pointer"
                      >
                        <Play className="w-3.5 h-3.5" />
                        <span>1. Fill Row #{currentRowIdx + 1}</span>
                      </button>

                      <button
                        onClick={handleApproveAndSubmit}
                        disabled={!filledInBrowser}
                        className="flex-1 flex items-center justify-center gap-1.5 px-3 py-2 bg-emerald-600 hover:bg-emerald-500 disabled:bg-slate-800 disabled:text-slate-500 text-white rounded-lg font-semibold cursor-pointer"
                      >
                        <CheckCircle2 className="w-3.5 h-3.5" />
                        <span>2. Approve & Submit</span>
                      </button>

                      <button
                        onClick={() => {
                          setFilledInBrowser(false);
                          setCurrentRowIdx(i => Math.min(rows.length - 1, i + 1));
                        }}
                        className="px-2.5 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg cursor-pointer"
                        title="Skip Row"
                      >
                        <SkipForward className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  </div>
                )}
              </div>
            </div>
          </div>
        ) : activeMode === 'web_to_excel' ? (
          /* =================================================================
             MODE 2: WEBSITE -> EXCEL (PART B)
             ================================================================= */
          <div className="p-5 bg-slate-900 border border-slate-800 rounded-xl space-y-4">
            <div className="flex flex-wrap items-center justify-between gap-4">
              <div>
                <h2 className="text-base font-semibold text-white flex items-center gap-2">
                  <Globe className="w-4 h-4 text-blue-400" />
                  <span>Part B: Website → Excel (Search-and-Fetch & Multi-Page Table Extractor)</span>
                </h2>
                <p className="text-xs text-slate-400 mt-0.5">
                  Reads each `Application_No` from Excel, searches it on the portal status page, extracts live Status, Date, Sanctioned Amount & Officer Remarks, and cleans numbers/dates automatically.
                </p>
              </div>
              <div className="flex items-center gap-2.5">
                <button
                  onClick={handleRunSearchAndFetch}
                  className="flex items-center gap-1.5 px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-lg text-xs font-semibold cursor-pointer"
                >
                  <Search className="w-3.5 h-3.5" />
                  <span>Run Search-and-Fetch on Portal (4 IDs)</span>
                </button>
                <button
                  onClick={() => handleExportCurrentExcel('website_status_extracted.xlsx')}
                  className="flex items-center gap-1.5 px-3.5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-semibold cursor-pointer"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>Download Extracted Excel (.xlsx)</span>
                </button>
              </div>
            </div>

            <div className="overflow-x-auto border border-slate-800 rounded-lg">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-950 text-slate-400 border-b border-slate-800">
                  <tr>
                    <th className="p-3">Application_No (Input)</th>
                    <th className="p-3">Applicant_Name</th>
                    <th className="p-3">District</th>
                    <th className="p-3">Fetched Portal_Status</th>
                    <th className="p-3">Approval_Date</th>
                    <th className="p-3">Sanctioned_Amount (Cleaned)</th>
                    <th className="p-3">Officer_Remarks</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800">
                  {rows.map(r => (
                    <tr key={r.Application_No} className="hover:bg-slate-950/50">
                      <td className="p-3 font-mono font-semibold text-blue-400">{r.Application_No}</td>
                      <td className="p-3 text-slate-200">{r.Applicant_Name}</td>
                      <td className="p-3 text-slate-300">{r.District}</td>
                      <td className="p-3">
                        <span
                          className={`px-2 py-0.5 rounded text-[11px] font-semibold ${
                            r.Status === 'Approved'
                              ? 'bg-emerald-950 text-emerald-300 border border-emerald-800'
                              : r.Status.includes('Reject')
                              ? 'bg-rose-950 text-rose-300 border border-rose-800'
                              : 'bg-amber-950 text-amber-300 border border-amber-800'
                          }`}
                        >
                          {r.Status}
                        </span>
                      </td>
                      <td className="p-3 font-mono text-slate-300">{r.Portal_Date || '--'}</td>
                      <td className="p-3 font-mono text-emerald-400">{r.Portal_Amount || '--'}</td>
                      <td className="p-3 text-slate-400">{r.Portal_Remarks || 'Click "Run Search-and-Fetch" above'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        ) : (
          /* =================================================================
             MODE 3: BOTH (FILL & VERIFY + RECONCILE) (PART C)
             ================================================================= */
          <div className="p-5 bg-slate-900 border border-slate-800 rounded-xl space-y-4">
            <div className="flex flex-wrap items-center justify-between gap-4">
              <div>
                <h2 className="text-base font-semibold text-white flex items-center gap-2">
                  <ArrowLeftRight className="w-4 h-4 text-purple-400" />
                  <span>Part C: Combined Fill & Verify + Excel vs. Website Reconciliation (`reconcile.py`)</span>
                </h2>
                <p className="text-xs text-slate-400 mt-0.5">
                  Compares your local Master Excel against the live Portal data and highlights status updates or amount mismatches in a dedicated `Differences` sheet.
                </p>
              </div>
              <button
                onClick={() => handleExportCurrentExcel('reconciliation_differences.xlsx')}
                className="flex items-center gap-1.5 px-4 py-2 bg-purple-600 hover:bg-purple-500 text-white rounded-lg text-xs font-semibold cursor-pointer"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Download `Differences` Workbook (.xlsx)</span>
              </button>
            </div>

            <div className="overflow-x-auto border border-slate-800 rounded-lg">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-950 text-slate-400 border-b border-slate-800">
                  <tr>
                    <th className="p-3">Application_No</th>
                    <th className="p-3">Applicant_Name</th>
                    <th className="p-3">Excel Status</th>
                    <th className="p-3">Website Live Status</th>
                    <th className="p-3">Excel Amount</th>
                    <th className="p-3">Website Amount</th>
                    <th className="p-3">Reconciliation Outcome</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800">
                  {rows.map(r => {
                    const web = PORTAL_LIVE_DATABASE[r.Application_No];
                    const amountDiff = web && web.amount !== String(r.Annual_Income === 52000 ? '10000' : r.Portal_Amount);
                    return (
                      <tr key={r.Application_No} className="hover:bg-slate-950/50">
                        <td className="p-3 font-mono text-purple-300">{r.Application_No}</td>
                        <td className="p-3 text-slate-200">{r.Applicant_Name}</td>
                        <td className="p-3 text-slate-400">Submitted</td>
                        <td className="p-3 font-semibold text-emerald-300">{web?.status}</td>
                        <td className="p-3 font-mono text-slate-300">{r.Application_No === 'RAJ-2025-103' ? '10000' : web?.amount}</td>
                        <td className="p-3 font-mono text-amber-300">{web?.amount}</td>
                        <td className="p-3">
                          {r.Application_No === 'RAJ-2025-103' ? (
                            <span className="px-2 py-0.5 rounded bg-amber-950 text-amber-300 border border-amber-800 text-[11px]">
                              Discrepancy: Amount Changed (10000 → 11500)
                            </span>
                          ) : r.Application_No === 'RAJ-2025-104' ? (
                            <span className="px-2 py-0.5 rounded bg-rose-950 text-rose-300 border border-rose-800 text-[11px]">
                              Status Changed: Rejected on Portal
                            </span>
                          ) : (
                            <span className="px-2 py-0.5 rounded bg-emerald-950 text-emerald-300 border border-emerald-800 text-[11px]">
                              Verified & Approved
                            </span>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </main>

      {/* Footer */}
      <footer className="border-t border-slate-800/80 bg-slate-950 px-6 py-4 text-xs text-slate-500 flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-center gap-2">
          <ShieldCheck className="w-4 h-4 text-emerald-400" />
          <span>Human-in-the-Loop Safety: Never bypasses Login, OTP, or CAPTCHA · Masks Aadhaar/Mobile (`XXXXXXXX1234`)</span>
        </div>
        <div className="flex items-center gap-4">
          <button
            onClick={() => setActiveMode('python_code')}
            className="text-emerald-400 hover:underline cursor-pointer font-medium"
          >
            View All 13 Python Modules & Hindi Guide
          </button>
          <span>·</span>
          <span>Google Gemini API (`google-genai`, default: `gemini-2.5-flash`)</span>
        </div>
      </footer>
    </div>
  );
}
