import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  FileSpreadsheet,
  Upload,
  Plus,
  RotateCcw,
  Layers,
  Sparkles,
  HelpCircle,
  FileCode,
  Download,
  Info,
  Activity,
  BarChart3,
  BookOpen,
  Volume2,
  VolumeX,
  CheckCircle2
} from 'lucide-react';

import { WorkspaceTable, ActionPlan, HistoryItem, HighlightRule } from './types';
import { getDefaultSampleTables } from './sampleData';
import {
  executeVLookup,
  executeSort,
  executeFindSimilarOrCommon,
  executeFilter,
  executeCleaning,
  executePivotTable,
  executeColumnTransform,
  executeMergeSheets,
  OperationResult
} from './engine/dataOperations';
import { parseNaturalInstruction } from './engine/nlpParser';
import { VoiceInput } from './components/VoiceInput';
import { DataGrid } from './components/DataGrid';
import { PythonProjectHub } from './components/PythonProjectHub';

export default function App() {
  // Main state
  const [tables, setTables] = useState<Record<string, WorkspaceTable>>(getDefaultSampleTables());
  const [activeTableId, setActiveTableId] = useState<string>('Orders');
  const [history, setHistory] = useState<HistoryItem[]>([]);
  const [activeView, setActiveView] = useState<'workspace' | 'python_project'>('workspace');

  // New Feature Toggles
  const [voiceReplyEnabled, setVoiceReplyEnabled] = useState(true);
  const [showQualityReport, setShowQualityReport] = useState(false);
  const [showDashboard, setShowDashboard] = useState(false);
  const [showLearningMode, setShowLearningMode] = useState(false);

  // Execution feedback state
  const [lastResultMeta, setLastResultMeta] = useState<{
    beforeRows?: number;
    afterRows?: number;
    statusMessage?: string;
    statusMessageHi?: string;
    highlightRules?: HighlightRule[];
    lastPlan?: ActionPlan;
    learningGuide?: { en: string; hi: string };
  }>({
    statusMessage: "Loaded sample tables: 'Orders', 'Prices', 'Inventory', and 'Messy_Data'. Speak or type to begin.",
    statusMessageHi: "नमूना टेबल्स लोड हो गई हैं। शुरू करने के लिए बोलें या निर्देश लिखें।"
  });

  const [isLoading, setIsLoading] = useState(false);
  const [clarification, setClarification] = useState<ActionPlan | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const activeTable = tables[activeTableId] || Object.values(tables)[0];

  // Browser Text-To-Speech (Feature 23)
  const speakFeedback = (text: string) => {
    if (!voiceReplyEnabled || !('speechSynthesis' in window)) return;
    try {
      window.speechSynthesis.cancel();
      const utterance = new SpeechSynthesisUtterance(text);
      utterance.rate = 1.0;
      utterance.pitch = 1.0;
      window.speechSynthesis.speak(utterance);
    } catch (e) {
      console.warn('TTS error:', e);
    }
  };

  // File Upload Handler
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const newTables: Record<string, WorkspaceTable> = { ...tables };

    Array.from(files).forEach(file => {
      const reader = new FileReader();
      reader.onload = evt => {
        try {
          const data = new Uint8Array(evt.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });

          workbook.SheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
            const json: Record<string, any>[] = XLSX.utils.sheet_to_json(worksheet, { defval: null });

            if (json.length > 0) {
              const baseName = file.name.replace(/\.[^/.]+$/, '');
              const tableKey = workbook.SheetNames.length > 1 ? `${baseName}_${sheetName}` : baseName;
              const columns = Object.keys(json[0]);

              newTables[tableKey] = {
                id: `table_${Date.now()}_${Math.random()}`,
                name: tableKey,
                fileName: file.name,
                columns,
                data: json,
                rowCount: json.length
              };
            }
          });

          setTables(newTables);
          const firstNew = Object.keys(newTables)[Object.keys(newTables).length - 1];
          setActiveTableId(firstNew);
          setLastResultMeta({
            statusMessage: `Successfully loaded file '${file.name}'.`,
            statusMessageHi: `फाइल '${file.name}' सफलतापूर्वक लोड हो गई।`
          });
        } catch (err: any) {
          console.error('File parsing error:', err);
          alert(`Error reading file ${file.name}: ${err.message}`);
        }
      };
      reader.readAsArrayBuffer(file);
    });
  };

  // Dispatch Action Plan
  const executePlan = (plan: ActionPlan) => {
    if (plan.action === 'clarification' || plan.needs_clarification) {
      setClarification(plan);
      return;
    }

    setClarification(null);

    // Save snapshot for Undo
    const snapshotTables: Record<string, WorkspaceTable> = JSON.parse(JSON.stringify(tables));
    const historyItem: HistoryItem = {
      id: `hist_${Date.now()}`,
      timestamp: new Date().toLocaleTimeString(),
      description: `${plan.action.toUpperCase()}: ${plan.output_name || 'Action'}`,
      actionPlan: plan,
      beforeRows: activeTable?.data.length || 0,
      afterRows: 0,
      tablesSnapshot: snapshotTables,
      activeTableId
    };

    try {
      let result: OperationResult;
      let learningGuide = {
        en: `Manual Excel: Apply formula or ribbon command corresponding to ${plan.action.toUpperCase()}.`,
        hi: `एक्सेल में यह करने के लिए ${plan.action.toUpperCase()} फॉर्मूला या रिबन टूल का उपयोग करें।`
      };

      if (plan.action === 'vlookup') {
        const left = tables[plan.left_table || activeTableId] || activeTable;
        const right = tables[plan.right_table || 'Prices'] || Object.values(tables)[1];
        if (!left || !right) throw new Error('Both left and right tables are required for VLOOKUP.');

        result = executeVLookup(
          left,
          right,
          plan.left_key || 'Item',
          plan.right_key || 'Item_Name',
          plan.return_columns || ['Unit_Price'],
          plan.fuzzy,
          plan.similarity_threshold,
          plan.output_name || `${left.name}_Lookup`
        );
        learningGuide = {
          en: `=XLOOKUP(A2, ${right.name}!A:A, ${right.name}!C:C, "") or =VLOOKUP(A2, ${right.name}!A:C, 3, FALSE)`,
          hi: `फॉर्मूला: =XLOOKUP(A2, ${right.name}!A:A, ${right.name}!C:C, "") या पुराना =VLOOKUP(A2, ${right.name}!A:C, 3, FALSE)`
        };
      } else if (plan.action === 'sort') {
        const target = tables[plan.table || activeTableId] || activeTable;
        result = executeSort(
          target,
          plan.columns || [target.columns[0]],
          plan.ascending ?? true,
          plan.output_name
        );
        learningGuide = {
          en: `Go to Data Tab ➔ Sort ➔ Choose column & order (Ascending A-Z or Descending Z-A).`,
          hi: `Data टैब में जाएं ➔ Sort पर क्लिक करें ➔ कॉलम और Ascending/Descending चुनें।`
        };
      } else if (plan.action === 'find_common') {
        const tA = tables[plan.table_a || activeTableId] || activeTable;
        const tB = tables[plan.table_b || 'Inventory'] || Object.values(tables)[1];
        result = executeFindSimilarOrCommon(
          tA,
          tB,
          plan.col_a || tA.columns[1],
          plan.col_b || tB.columns[0],
          plan.mode || 'common',
          plan.fuzzy,
          plan.similarity_threshold,
          plan.output_name
        );
      } else if (plan.action === 'filter') {
        const target = tables[plan.table || activeTableId] || activeTable;
        result = executeFilter(
          target,
          plan.column || target.columns[0],
          plan.operator || 'greater_than',
          plan.value,
          plan.value2,
          plan.output_name
        );
      } else if (plan.action === 'clean') {
        const target = tables[plan.table || activeTableId] || activeTable;
        result = executeCleaning(
          target,
          plan.operation || 'remove_duplicates',
          plan.columns,
          plan.fill_value,
          plan.output_name
        );
        learningGuide = {
          en: `Select data ➔ Data Tab ➔ click 'Remove Duplicates' ➔ select columns.`,
          hi: `पूरा टेबल चुनें ➔ Data टैब ➔ 'Remove Duplicates' दबाएं।`
        };
      } else if (plan.action === 'pivot') {
        const target = tables[plan.table || activeTableId] || activeTable;
        result = executePivotTable(
          target,
          plan.index_columns || ['City'],
          plan.values_columns || ['Quantity'],
          plan.agg_func || 'sum',
          plan.output_name
        );
        learningGuide = {
          en: `Insert Tab ➔ PivotTable ➔ Drag category to Rows and metric to Values.`,
          hi: `Insert टैब ➔ PivotTable ➔ श्रेणी को Rows में और संख्या को Values में डालें।`
        };
      } else if (plan.action === 'calculate') {
        const target = tables[plan.table || activeTableId] || activeTable;
        result = executeColumnTransform(
          target,
          'calculate',
          plan.target_column || 'Total',
          plan.formula || 'Quantity * Unit_Price',
          plan.source_columns,
          plan.delimiter,
          plan.output_name
        );
      } else if (plan.action === 'merge_sheets') {
        const toMerge = (plan.sheets || Object.keys(tables)).map(k => tables[k]).filter(Boolean);
        result = executeMergeSheets(toMerge, 'append', plan.output_name);
      } else {
        throw new Error(`Unsupported action: ${plan.action}`);
      }

      // Update tables with new/modified table
      const updatedTables = {
        ...tables,
        [result.table.name]: result.table
      };

      historyItem.afterRows = result.afterRows;
      setHistory(prev => [...prev, historyItem]);
      setTables(updatedTables);
      setActiveTableId(result.table.name);

      setLastResultMeta({
        beforeRows: result.beforeRows,
        afterRows: result.afterRows,
        statusMessage: result.statusMessage,
        statusMessageHi: result.statusMessageHi,
        highlightRules: result.highlightRules,
        lastPlan: plan,
        learningGuide
      });

      speakFeedback(result.statusMessage);
    } catch (err: any) {
      console.error('Execution error:', err);
      setLastResultMeta({
        statusMessage: `Error: ${err.message}`,
        statusMessageHi: `त्रुटि: ${err.message}`
      });
      speakFeedback(`Error: ${err.message}`);
    } finally {
      setIsLoading(false);
    }
  };

  const handleInstruction = (instruction: string) => {
    setIsLoading(true);
    const plan = parseNaturalInstruction(instruction, tables, activeTableId);
    executePlan(plan);
  };

  const handleUndo = () => {
    if (history.length === 0) return;
    const lastItem = history[history.length - 1];
    setTables(lastItem.tablesSnapshot);
    setActiveTableId(lastItem.activeTableId);
    setHistory(prev => prev.slice(0, -1));
    setLastResultMeta({
      statusMessage: `Undid: ${lastItem.description}`,
      statusMessageHi: `पिछला एक्शन वापस लिया गया।`
    });
    speakFeedback("Action undone.");
  };

  const handleResetSamples = () => {
    const samples = getDefaultSampleTables();
    setTables(samples);
    setActiveTableId('Orders');
    setHistory([]);
    setLastResultMeta({
      statusMessage: "Reset to default sample tables: Orders, Prices, Inventory, and Messy_Data.",
      statusMessageHi: "सैंपल टेबल्स रीसेट कर दी गईं।"
    });
  };

  // Quick Data Quality Stats for active table
  const getQualityStats = () => {
    if (!activeTable) return { duplicates: 0, blanks: 0, score: 100 };
    const duplicates = activeTable.data.filter((r, idx) =>
      activeTable.data.some((r2, idx2) => idx2 < idx && JSON.stringify(r) === JSON.stringify(r2))
    ).length;
    let blanks = 0;
    activeTable.data.forEach(r => {
      activeTable.columns.forEach(c => {
        if (r[c] === null || r[c] === undefined || String(r[c]).trim() === '') blanks++;
      });
    });
    const score = Math.max(10, Math.min(100, 100 - (duplicates * 10 + blanks * 5)));
    return { duplicates, blanks, score };
  };

  const qualityStats = getQualityStats();

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 flex flex-col font-sans">
      {/* Top Bar Contract */}
      <header className="border-b border-slate-800 bg-slate-950/80 backdrop-blur-md sticky top-0 z-50 px-6 py-3.5 flex items-center justify-between gap-8">
        <div className="flex items-center gap-2.5 whitespace-nowrap shrink-0">
          <div className="w-8 h-8 rounded-lg bg-emerald-600 flex items-center justify-center text-white shadow-md shadow-emerald-950">
            <FileSpreadsheet className="w-4 h-4" />
          </div>
          <span className="text-base font-bold tracking-tight text-white">
            Voice Excel AI
          </span>
          <span className="hidden sm:inline-block text-xs text-slate-500 font-mono border-l border-slate-800 pl-2.5">
            Full Python & Web Suite
          </span>
        </div>

        <nav className="hidden md:flex items-center gap-6 text-xs font-medium text-slate-400">
          <button
            onClick={() => setActiveView('workspace')}
            className={`transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
              activeView === 'workspace' ? 'text-emerald-400 font-semibold' : 'hover:text-slate-200'
            }`}
          >
            Assistant Workspace
          </button>
          <button
            onClick={() => setActiveView('python_project')}
            className={`transition-colors whitespace-nowrap shrink-0 cursor-pointer flex items-center gap-1.5 ${
              activeView === 'python_project' ? 'text-emerald-400 font-semibold' : 'hover:text-slate-200'
            }`}
          >
            <FileCode className="w-3.5 h-3.5 text-emerald-400" />
            <span>Python Project & Code Hub</span>
          </button>
        </nav>

        <div className="flex items-center gap-3 shrink-0">
          <button
            onClick={() => setVoiceReplyEnabled(v => !v)}
            className={`p-2 rounded-lg border text-xs transition-colors cursor-pointer flex items-center gap-1.5 ${
              voiceReplyEnabled ? 'bg-emerald-950/40 text-emerald-300 border-emerald-800/60' : 'bg-slate-900 text-slate-400 border-slate-800'
            }`}
            title="Toggle Voice Reply (Text-to-Speech)"
          >
            {voiceReplyEnabled ? <Volume2 className="w-4 h-4" /> : <VolumeX className="w-4 h-4" />}
            <span className="hidden sm:inline">Voice Reply</span>
          </button>

          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileUpload}
            multiple
            accept=".xlsx,.xls,.csv"
            className="hidden"
          />
          <button
            onClick={() => fileInputRef.current?.click()}
            className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-900 hover:bg-slate-800 text-slate-200 border border-slate-800 rounded-lg text-xs font-medium transition-colors cursor-pointer"
          >
            <Upload className="w-3.5 h-3.5 text-slate-400" />
            <span className="hidden sm:inline">Upload Files</span>
          </button>

          <a
            href="/voice_controlled_excel_assistant.zip"
            download="voice_controlled_excel_assistant.zip"
            className="flex items-center gap-1.5 px-3 py-1.5 bg-amber-500 hover:bg-amber-400 text-slate-950 rounded-lg text-xs font-bold transition-all cursor-pointer shadow-sm shadow-amber-950 whitespace-nowrap"
            title="Download full project ZIP containing all Python code & sample Excel files"
          >
            <Download className="w-3.5 h-3.5 text-slate-950" />
            <span>Download .ZIP</span>
          </a>

          <button
            onClick={handleResetSamples}
            className="px-3.5 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-semibold transition-all cursor-pointer shadow-sm shadow-emerald-950 whitespace-nowrap"
          >
            Load 4 Samples
          </button>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-4 md:p-6 space-y-6">
        {activeView === 'python_project' ? (
          <PythonProjectHub />
        ) : (
          <>
            {/* Sheet Tabs & Quick Action Bar */}
            <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-800/80 pb-3">
              <div className="flex items-center gap-2 overflow-x-auto pb-1">
                <span className="text-xs text-slate-400 font-medium mr-1 flex items-center gap-1">
                  <Layers className="w-3.5 h-3.5" />
                  <span>Tables:</span>
                </span>
                {Object.keys(tables).map(name => {
                  const isActive = activeTableId === name;
                  const tbl = tables[name];
                  return (
                    <button
                      key={name}
                      onClick={() => setActiveTableId(name)}
                      className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-all flex items-center gap-2 cursor-pointer whitespace-nowrap ${
                        isActive
                          ? 'bg-slate-800 text-emerald-400 border border-slate-700 shadow-sm'
                          : 'text-slate-400 hover:text-slate-200 hover:bg-slate-900'
                      }`}
                    >
                      <span>{name}</span>
                      <span className="text-[11px] font-mono tabular-nums text-slate-500">
                        {tbl.data.length}
                      </span>
                    </button>
                  );
                })}

                <button
                  onClick={() => fileInputRef.current?.click()}
                  className="px-2.5 py-1.5 rounded-lg text-xs text-slate-500 hover:text-slate-300 hover:bg-slate-900 flex items-center gap-1 cursor-pointer"
                  title="Upload another Excel or CSV file"
                >
                  <Plus className="w-3.5 h-3.5" />
                  <span>Add File</span>
                </button>
              </div>

              {/* Utility Badges */}
              <div className="flex items-center gap-2">
                <button
                  onClick={() => setShowQualityReport(v => !v)}
                  className={`flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-xs border transition-colors cursor-pointer ${
                    showQualityReport ? 'bg-emerald-950/60 text-emerald-300 border-emerald-700' : 'bg-slate-900 text-slate-400 border-slate-800 hover:text-slate-200'
                  }`}
                >
                  <Activity className="w-3.5 h-3.5 text-emerald-400" />
                  <span>Quality Audit ({qualityStats.score}%)</span>
                </button>

                <button
                  onClick={() => setShowDashboard(v => !v)}
                  className={`flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-xs border transition-colors cursor-pointer ${
                    showDashboard ? 'bg-blue-950/60 text-blue-300 border-blue-700' : 'bg-slate-900 text-slate-400 border-slate-800 hover:text-slate-200'
                  }`}
                >
                  <BarChart3 className="w-3.5 h-3.5 text-blue-400" />
                  <span>Dashboard</span>
                </button>

                <button
                  onClick={() => setShowLearningMode(v => !v)}
                  className={`flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-xs border transition-colors cursor-pointer ${
                    showLearningMode ? 'bg-amber-950/60 text-amber-300 border-amber-700' : 'bg-slate-900 text-slate-400 border-slate-800 hover:text-slate-200'
                  }`}
                >
                  <BookOpen className="w-3.5 h-3.5 text-amber-400" />
                  <span>Ye Kaise Kiya?</span>
                </button>
              </div>
            </div>

            {/* Quality Report Dropdown Card */}
            {showQualityReport && (
              <div className="p-4 bg-slate-900 border border-slate-800 rounded-xl space-y-3">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2 text-xs font-semibold text-slate-100">
                    <Activity className="w-4 h-4 text-emerald-400" />
                    <span>Data Quality Report for '{activeTable?.name}'</span>
                  </div>
                  <button
                    onClick={() => {
                      executePlan({ action: 'clean', operation: 'remove_duplicates' });
                      setShowQualityReport(false);
                    }}
                    className="px-3 py-1 bg-emerald-600 hover:bg-emerald-500 text-white rounded text-xs font-medium cursor-pointer"
                  >
                    Fix All Automatically (Clean & Standardize)
                  </button>
                </div>
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
                  <div className="p-3 bg-slate-950 border border-slate-800 rounded-lg">
                    <span className="text-slate-500">Health Score</span>
                    <div className="text-lg font-bold text-emerald-400 font-mono mt-0.5">{qualityStats.score}/100</div>
                  </div>
                  <div className="p-3 bg-slate-950 border border-slate-800 rounded-lg">
                    <span className="text-slate-500">Duplicate Rows</span>
                    <div className="text-lg font-bold text-slate-200 font-mono mt-0.5">{qualityStats.duplicates}</div>
                  </div>
                  <div className="p-3 bg-slate-950 border border-slate-800 rounded-lg">
                    <span className="text-slate-500">Blank Cells</span>
                    <div className="text-lg font-bold text-slate-200 font-mono mt-0.5">{qualityStats.blanks}</div>
                  </div>
                  <div className="p-3 bg-slate-950 border border-slate-800 rounded-lg">
                    <span className="text-slate-500">Total Columns</span>
                    <div className="text-lg font-bold text-slate-200 font-mono mt-0.5">{activeTable?.columns.length}</div>
                  </div>
                </div>
              </div>
            )}

            {/* Quick Dashboard View */}
            {showDashboard && activeTable && (
              <div className="p-4 bg-slate-900 border border-slate-800 rounded-xl space-y-4">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2 text-xs font-semibold text-slate-100">
                    <BarChart3 className="w-4 h-4 text-blue-400" />
                    <span>Quick Visual Dashboard: '{activeTable.name}'</span>
                  </div>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  {activeTable.columns.slice(0, 3).map((col, idx) => (
                    <div key={idx} className="p-3 bg-slate-950 border border-slate-800 rounded-lg">
                      <span className="text-xs text-slate-400 font-medium">Column: {col}</span>
                      <div className="mt-2 space-y-1.5 text-xs">
                        {activeTable.data.slice(0, 4).map((r, rIdx) => (
                          <div key={rIdx} className="flex justify-between text-slate-300">
                            <span className="truncate max-w-[120px]">{String(r[activeTable.columns[0]] ?? '')}</span>
                            <span className="font-mono tabular-nums text-emerald-400">{String(r[col] ?? '')}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Excel Learning Mode Card ("Ye Kaise Kiya?") */}
            {showLearningMode && lastResultMeta.learningGuide && (
              <div className="p-4 bg-amber-950/20 border border-amber-800/60 rounded-xl space-y-2">
                <div className="flex items-center gap-2 text-xs font-semibold text-amber-300">
                  <BookOpen className="w-4 h-4" />
                  <span>💡 Ye Kaise Kiya? (Excel Learning Mode)</span>
                </div>
                <div className="text-xs text-slate-300 space-y-1">
                  <p><strong>English:</strong> {lastResultMeta.learningGuide.en}</p>
                  <p><strong>हिंदी में सीखें:</strong> {lastResultMeta.learningGuide.hi}</p>
                </div>
              </div>
            )}

            {/* Voice & Text Controller */}
            <VoiceInput
              onExecute={handleInstruction}
              isLoading={isLoading}
              activeTableName={activeTable?.name}
              onUndo={handleUndo}
              canUndo={history.length > 0}
            />

            {/* Clarification Dialog */}
            {clarification && (
              <div className="p-4 bg-amber-950/30 border border-amber-800/60 rounded-xl space-y-3">
                <div className="flex items-start gap-2.5">
                  <HelpCircle className="w-5 h-5 text-amber-400 shrink-0 mt-0.5" />
                  <div>
                    <h3 className="text-xs font-semibold text-amber-200">
                      Clarification Needed (स्पष्टीकरण चाहिए)
                    </h3>
                    <p className="text-xs text-slate-300 mt-1">{clarification.question_en}</p>
                    {clarification.question_hi && (
                      <p className="text-xs text-slate-400 mt-0.5">{clarification.question_hi}</p>
                    )}
                  </div>
                </div>

                {clarification.options && (
                  <div className="flex flex-wrap gap-2 pl-7">
                    {clarification.options.map((opt, idx) => (
                      <button
                        key={idx}
                        onClick={() => {
                          const updatedPlan = {
                            ...clarification,
                            return_columns: [opt],
                            needs_clarification: false
                          };
                          executePlan(updatedPlan);
                        }}
                        className="px-3 py-1 bg-amber-900/40 hover:bg-amber-800/60 text-amber-200 border border-amber-700/60 rounded-md text-xs font-medium transition-colors cursor-pointer"
                      >
                        Select: {opt}
                      </button>
                    ))}
                  </div>
                )}
              </div>
            )}

            {/* Data Grid with Pagination and Before/After Rows */}
            {activeTable && (
              <DataGrid
                table={activeTable}
                beforeRows={lastResultMeta.beforeRows}
                afterRows={lastResultMeta.afterRows}
                statusMessage={lastResultMeta.statusMessage}
                statusMessageHi={lastResultMeta.statusMessageHi}
                highlightRules={lastResultMeta.highlightRules}
              />
            )}
          </>
        )}
      </main>

      {/* Footer */}
      <footer className="border-t border-slate-800/80 bg-slate-950 px-6 py-4 text-xs text-slate-500 flex flex-wrap items-center justify-between gap-4">
        <div>
          Voice-Controlled Excel Assistant · Live Formulas · Q&A · Quality Audits · Macro Automation
        </div>
        <div className="flex items-center gap-4">
          <button
            onClick={() => setActiveView('python_project')}
            className="hover:text-slate-300 transition-colors cursor-pointer"
          >
            Python Source Code & README
          </button>
          <span>·</span>
          <span>Configurable via CLAUDE_MODEL (default: claude-sonnet-5-5)</span>
        </div>
      </footer>
    </div>
  );
}
