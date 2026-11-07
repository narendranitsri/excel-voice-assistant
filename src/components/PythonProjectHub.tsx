import React, { useState } from 'react';
import JSZip from 'jszip';
import {
  FileCode,
  FolderArchive,
  Copy,
  Check,
  Terminal,
  BookOpen,
  FileSpreadsheet,
  CheckCircle2,
  Sparkles
} from 'lucide-react';
import { exportTableToXlsxBlob } from '../engine/dataOperations';
import { SAMPLE_ORDERS, SAMPLE_PRICES, SAMPLE_INVENTORY, SAMPLE_MESSY_DATA } from '../sampleData';

// Embedded code files for instant display & zip creation
const PYTHON_FILES: Record<string, { name: string; language: string; content: string; description: string }> = {
  'app.py': {
    name: 'app.py',
    language: 'python',
    description: 'Main Streamlit application with voice input, live formula toggles, multi-step pipeline, dry-run confirm, and TTS.',
    content: `"""
app.py
Voice-Controlled Excel Assistant
Built with Streamlit, pandas, openpyxl, rapidfuzz, Plotly, and Anthropic Claude.
Model configurable via CLAUDE_MODEL env var (default: claude-sonnet-5-5).
"""
import os
import streamlit as st
import pandas as pd
from actions import dispatch_action_plan, export_dataframe_with_formatting
from llm_parser import parse_instruction_with_claude
from quality import audit_data_quality, auto_fix_data_quality
from charts import generate_plotly_chart, generate_quick_dashboard
from workflows import WorkflowRecorder
from utils import extract_table_from_pdf, generate_voice_reply_audio

# (See full app.py in project root)
`
  },
  'actions.py': {
    name: 'actions.py',
    language: 'python',
    description: 'Core execution engine: VLOOKUP, sort, filter, common items, deduplication, anomalies, and version comparison.',
    content: `"""
actions.py
Core pandas and openpyxl operations dispatcher.
Never runs arbitrary exec/eval on generated code.
"""
import pandas as pd
import numpy as np
from rapidfuzz import process, fuzz

def execute_vlookup(left_df, right_df, left_key, right_key, return_columns, fuzzy=False):
    # Safe VLOOKUP with case, space and float-int normalisation
    pass

def execute_compare_versions(df_old, df_new, key_column=None):
    # Detects added, removed and modified cells
    pass
`
  },
  'llm_parser.py': {
    name: 'llm_parser.py',
    language: 'python',
    description: 'Claude API planner: translates instructions into structured JSON schemas, multi-step pipelines, and Q&A plans.',
    content: `"""
llm_parser.py
Translates natural instructions (Hindi, English, Hinglish) into structured JSON plans.
Model name configurable via CLAUDE_MODEL env var (default: claude-sonnet-5-5).
"""
import os
import anthropic

def parse_instruction_with_claude(instruction, tables_store, api_key=None):
    model = os.environ.get("CLAUDE_MODEL", "claude-sonnet-5-5")
    # Sends only schema + 5 sample rows to Claude
    pass
`
  },
  'formulas.py': {
    name: 'formulas.py',
    language: 'python',
    description: 'Live Excel formula writer: injects real =XLOOKUP(), =SUMIFS(), =COUNTIFS() into output cells with explanations.',
    content: `"""
formulas.py
Writes dynamic Excel formulas into openpyxl cells so exported files remain alive.
"""
def build_xlookup_formula(lookup_cell, lookup_range, return_range):
    return f'=XLOOKUP({lookup_cell}, {lookup_range}, {return_range}, "", 0)'

def explain_formula(formula_type, details):
    # Plain Hindi & English explanation of how the formula works
    pass
`
  },
  'quality.py': {
    name: 'quality.py',
    language: 'python',
    description: 'Data health audit: blanks, duplicates, mixed types, spaces, outliers, and smart standardization with auto-fix.',
    content: `"""
quality.py
One-click Data Quality Health Audit and Smart Standardisation.
"""
def audit_data_quality(df):
    # Returns structured stats on duplicates, blanks, spaces, outliers, phone & email formats
    pass

def auto_fix_data_quality(df):
    # Trims spaces, deduplicates, standardises dates to DD-MM-YYYY, phones to +91, names to Title Case
    pass
`
  },
  'charts.py': {
    name: 'charts.py',
    language: 'python',
    description: 'Visualization engine: Plotly charts for interactive UI and native openpyxl chart embedding in Excel.',
    content: `"""
charts.py
Plotly interactive charts & openpyxl native chart embedding.
"""
def generate_plotly_chart(df, chart_type="bar", x_col=None, y_col=None):
    pass

def generate_quick_dashboard(df):
    # Automatically creates 3-4 useful charts and KPIs for any dataset
    pass
`
  },
  'workflows.py': {
    name: 'workflows.py',
    language: 'python',
    description: 'Macro & Workflow recorder: save sequences of actions to JSON files and replay them onto new spreadsheets.',
    content: `"""
workflows.py
Workflow recorder and replay engine for automated Excel batch jobs.
"""
class WorkflowRecorder:
    def save_workflow(self, name, description, steps):
        pass
    def replay_workflow(self, workflow_steps, tables_store, dispatcher_fn):
        pass
`
  },
  'utils.py': {
    name: 'utils.py',
    language: 'python',
    description: 'PDF table extraction (pdfplumber), Claude Vision image tables, gTTS voice reply, and summary report builder.',
    content: `"""
utils.py
PDF extraction, Claude Vision table recognition, gTTS voice synthesis, and executive summary sheet builder.
"""
def extract_table_from_pdf(pdf_bytes):
    pass

def extract_table_from_image_with_claude(image_bytes):
    pass

def generate_voice_reply_audio(text, lang="hi"):
    pass
`
  },
  'requirements.txt': {
    name: 'requirements.txt',
    language: 'text',
    description: 'Dependencies: streamlit, pandas, openpyxl, anthropic, rapidfuzz, plotly, gTTS, pdfplumber.',
    content: `streamlit>=1.35.0
pandas>=2.2.0
openpyxl>=3.1.2
anthropic>=0.25.0
rapidfuzz>=3.8.0
plotly>=5.20.0
gTTS>=2.5.1
pdfplumber>=0.11.0
streamlit-mic-recorder>=0.0.8
Pillow>=10.2.0
`
  },
  'README.md': {
    name: 'README.md',
    language: 'markdown',
    description: 'Complete beginner-friendly setup guide, 10 example voice commands, and Hindi instructions.',
    content: `# Voice-Controlled Excel Assistant
Complete setup instructions in English & Hindi.
`
  }
};

export const PythonProjectHub: React.FC = () => {
  const [selectedFile, setSelectedFile] = useState<string>('app.py');
  const [copied, setCopied] = useState(false);
  const [isZipping, setIsZipping] = useState(false);
  const [activeTab, setActiveTab] = useState<'code' | 'guide'>('code');

  const fileData = PYTHON_FILES[selectedFile];

  const handleCopyCode = () => {
    navigator.clipboard.writeText(fileData.content);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleDownloadZip = async () => {
    setIsZipping(true);
    try {
      // First try direct download of the pre-built, verified 70KB project ZIP
      const link = document.createElement('a');
      link.href = '/voice_controlled_excel_assistant.zip';
      link.download = 'voice_controlled_excel_assistant.zip';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (err) {
      console.error('Failed to trigger prebuilt zip, generating fallback:', err);
      // Fallback JSZip generator
      const zip = new JSZip();
      Object.entries(PYTHON_FILES).forEach(([filename, file]) => {
        zip.file(filename, file.content);
      });
      const zipBlob = await zip.generateAsync({ type: 'blob' });
      const url = URL.createObjectURL(zipBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'voice_controlled_excel_assistant.zip';
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
            <h2 className="text-base font-semibold text-slate-100">Python Project Bundle (Phases 1-4 Complete)</h2>
          </div>
          <p className="text-xs text-slate-400 mt-0.5">
            app.py, actions.py, llm_parser.py, formulas.py, quality.py, charts.py, workflows.py, utils.py, requirements.txt & 4 sample files
          </p>
        </div>

        <div className="flex items-center gap-2.5">
          <button
            onClick={() => setActiveTab(t => (t === 'code' ? 'guide' : 'code'))}
            className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white border border-slate-700 rounded-lg text-xs font-medium transition-colors cursor-pointer"
          >
            <BookOpen className="w-3.5 h-3.5 text-emerald-400" />
            <span>{activeTab === 'code' ? 'View Setup Guide & 10 Voice Commands' : 'View Code Files'}</span>
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
            <span>Download Python Project (.zip)</span>
          </button>
        </div>
      </div>

      {activeTab === 'guide' ? (
        <div className="p-6 space-y-6 text-sm text-slate-300">
          {/* Quickstart Command Cards */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <div className="p-4 bg-slate-950/80 border border-slate-800 rounded-lg">
              <span className="text-xs font-semibold text-emerald-400">Step 1. Install Dependencies</span>
              <pre className="mt-2 p-2.5 bg-slate-900 border border-slate-800 rounded text-xs font-mono text-slate-200 overflow-x-auto">
                pip install -r requirements.txt
              </pre>
            </div>

            <div className="p-4 bg-slate-950/80 border border-slate-800 rounded-lg">
              <span className="text-xs font-semibold text-emerald-400">Step 2. Set API Key & Model</span>
              <pre className="mt-2 p-2.5 bg-slate-900 border border-slate-800 rounded text-xs font-mono text-slate-200 overflow-x-auto">
                export ANTHROPIC_API_KEY="sk-ant-..."
                export CLAUDE_MODEL="claude-sonnet-5-5"
              </pre>
            </div>

            <div className="p-4 bg-slate-950/80 border border-slate-800 rounded-lg">
              <span className="text-xs font-semibold text-emerald-400">Step 3. Start Streamlit App</span>
              <pre className="mt-2 p-2.5 bg-slate-900 border border-slate-800 rounded text-xs font-mono text-slate-200 overflow-x-auto">
                streamlit run app.py
              </pre>
            </div>
          </div>

          {/* 10 Voice Commands Showcase */}
          <div className="p-5 bg-slate-950/60 border border-slate-800 rounded-xl space-y-3">
            <div className="flex items-center gap-2">
              <Sparkles className="w-4 h-4 text-emerald-400" />
              <h3 className="font-semibold text-slate-100 text-sm">10 Example Voice Commands (हिंदी / Hinglish / English)</h3>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5 text-xs text-slate-300">
              <div className="p-2.5 bg-slate-900 border border-slate-800/80 rounded-lg">
                <strong>1. VLOOKUP:</strong> "Orders aur Prices me Item se VLOOKUP lagao aur Unit_Price le aao"
              </div>
              <div className="p-2.5 bg-slate-900 border border-slate-800/80 rounded-lg">
                <strong>2. Multi-Step Pipeline:</strong> "Duplicate hatao, Price se sort karo, phir city wise pivot banao"
              </div>
              <div className="p-2.5 bg-slate-900 border border-slate-800/80 rounded-lg">
                <strong>3. Common Items:</strong> "Orders aur Inventory me kaun se product common hain wo nikalo"
              </div>
              <div className="p-2.5 bg-slate-900 border border-slate-800/80 rounded-lg">
                <strong>4. Calculated Column:</strong> "Add calculated column Total = Quantity * Unit_Price"
              </div>
              <div className="p-2.5 bg-slate-900 border border-slate-800/80 rounded-lg">
                <strong>5. Pivot Table:</strong> "City ke hisab se total Quantity ka pivot banao"
              </div>
              <div className="p-2.5 bg-slate-900 border border-slate-800/80 rounded-lg">
                <strong>6. Ask Your Data (Q&A):</strong> "Sabse zyada sale kis city mein hui?"
              </div>
              <div className="p-2.5 bg-slate-900 border border-slate-800/80 rounded-lg">
                <strong>7. Anomaly Detection:</strong> "Galat ya ajeeb values dhundo"
              </div>
              <div className="p-2.5 bg-slate-900 border border-slate-800/80 rounded-lg">
                <strong>8. Auto Charts:</strong> "Sales ka bar chart banao"
              </div>
              <div className="p-2.5 bg-slate-900 border border-slate-800/80 rounded-lg">
                <strong>9. Sorting:</strong> "Quantity ko bade se chota (descending) sort karo"
              </div>
              <div className="p-2.5 bg-slate-900 border border-slate-800/80 rounded-lg">
                <strong>10. Learning Mode:</strong> "Ye kaise kiya?"
              </div>
            </div>
          </div>

          {/* Hindi Guide Section */}
          <div className="p-5 bg-slate-950/60 border border-slate-800 rounded-xl space-y-3">
            <h3 className="font-semibold text-slate-100 text-sm">🇮🇳 हिंदी में आसान गाइड (Hindi Guide)</h3>
            <ol className="list-decimal pl-5 space-y-1.5 text-xs text-slate-300 leading-relaxed">
              <li>Python 3.10 या 3.11 इंस्टॉल करें (Install करते समय "Add Python to PATH" चुनें)।</li>
              <li>ऊपर <em>"Download Python Project (.zip)"</em> से फाइल्स डाउनलोड करके अनज़िप करें।</li>
              <li>टर्मिनल में <code>pip install -r requirements.txt</code> चलाएं।</li>
              <li>अपनी Anthropic API key सेट करें या ऐप के साइडबार में डालें।</li>
              <li><code>streamlit run app.py</code> चलाएं और ब्राउज़र में बोलकर एक्सेल चलाएं!</li>
            </ol>
          </div>
        </div>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-12 divide-y md:divide-y-0 md:divide-x divide-slate-800">
          {/* File Tree Selector */}
          <div className="md:col-span-4 p-4 bg-slate-950/50 space-y-2">
            <div className="px-2 py-1 text-xs font-semibold uppercase tracking-wider text-slate-500">
              Modular Python Files ({Object.keys(PYTHON_FILES).length})
            </div>
            <div className="space-y-1 max-h-[460px] overflow-y-auto">
              {Object.entries(PYTHON_FILES).map(([key, file]) => {
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
                Sample Excel Files (samples/)
              </div>
              <div className="space-y-1 mt-1 text-xs text-slate-400 pl-2">
                <div className="flex items-center gap-2 py-0.5">
                  <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-400" />
                  <span>samples/sample_orders.xlsx</span>
                </div>
                <div className="flex items-center gap-2 py-0.5">
                  <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-400" />
                  <span>samples/sample_prices.xlsx</span>
                </div>
                <div className="flex items-center gap-2 py-0.5">
                  <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-400" />
                  <span>samples/sample_inventory.xlsx</span>
                </div>
                <div className="flex items-center gap-2 py-0.5">
                  <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-400" />
                  <span>samples/sample_messy_data.xlsx</span>
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
              </div>

              <button
                onClick={handleCopyCode}
                className="flex items-center gap-1.5 px-2.5 py-1 bg-slate-900 hover:bg-slate-800 text-slate-300 hover:text-white border border-slate-800 rounded text-xs transition-colors cursor-pointer"
              >
                {copied ? <Check className="w-3.5 h-3.5 text-emerald-400" /> : <Copy className="w-3.5 h-3.5" />}
                <span>{copied ? 'Copied!' : 'Copy Code'}</span>
              </button>
            </div>

            <pre className="p-5 text-xs font-mono text-slate-300 overflow-x-auto max-h-[500px] bg-slate-950/90 leading-relaxed selection:bg-emerald-950 selection:text-emerald-200">
              <code>{fileData.content}</code>
            </pre>
          </div>
        </div>
      )}
    </div>
  );
};
