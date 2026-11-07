import React, { useState } from 'react';
import { Download, Search, Table, ChevronLeft, ChevronRight, FileSpreadsheet, Layers, Sparkles } from 'lucide-react';
import { WorkspaceTable, HighlightRule } from '../types';
import { exportTableToXlsxBlob } from '../engine/dataOperations';

interface DataGridProps {
  table: WorkspaceTable;
  beforeRows?: number;
  afterRows?: number;
  highlightRules?: HighlightRule[];
  statusMessage?: string;
  statusMessageHi?: string;
}

export const DataGrid: React.FC<DataGridProps> = ({
  table,
  beforeRows,
  afterRows,
  highlightRules = [],
  statusMessage,
  statusMessageHi
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [page, setPage] = useState(1);
  const pageSize = 15;

  // Filtered rows
  const filteredData = table.data.filter(row => {
    if (!searchTerm) return true;
    return Object.values(row).some(val =>
      String(val).toLowerCase().includes(searchTerm.toLowerCase())
    );
  });

  const totalPages = Math.ceil(filteredData.length / pageSize) || 1;
  const currentPage = Math.min(page, totalPages);
  const paginatedData = filteredData.slice((currentPage - 1) * pageSize, currentPage * pageSize);

  const handleDownloadExcel = () => {
    const blob = exportTableToXlsxBlob(table);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${table.name}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleDownloadCsv = () => {
    const headers = table.columns.join(',');
    const rows = table.data.map(r =>
      table.columns.map(c => `"${String(r[c] ?? '').replace(/"/g, '""')}"`).join(',')
    );
    const csvContent = [headers, ...rows].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${table.name}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const isHighlighted = (colName: string, val: any): string | null => {
    for (const rule of highlightRules) {
      if (rule.column.toLowerCase() === colName.toLowerCase()) {
        if (rule.color === 'green') return 'bg-emerald-950/40 text-emerald-200 border-l border-emerald-500/40';
        if (rule.color === 'yellow') return 'bg-amber-950/40 text-amber-200 border-l border-amber-500/40';
        if (rule.color === 'coral') return 'bg-rose-950/40 text-rose-200 border-l border-rose-500/40';
      }
    }
    return null;
  };

  return (
    <div className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden shadow-lg">
      {/* Top Status & Row-Count Change Strip */}
      {statusMessage && (
        <div className="px-5 py-3.5 bg-slate-900/90 border-b border-slate-800 flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-2">
            <span className="w-2 h-2 rounded-full bg-emerald-400"></span>
            <div>
              <p className="text-xs font-semibold text-slate-100">{statusMessage}</p>
              {statusMessageHi && statusMessageHi !== statusMessage && (
                <p className="text-xs text-slate-400 mt-0.5">{statusMessageHi}</p>
              )}
            </div>
          </div>

          {beforeRows !== undefined && afterRows !== undefined && (
            <div className="flex items-center gap-2 text-xs font-mono tabular-nums px-3 py-1 bg-slate-950 border border-slate-800 rounded-md">
              <span className="text-slate-400">Rows:</span>
              <span className="text-slate-300">{beforeRows.toLocaleString()}</span>
              <span className="text-emerald-400 font-bold">➔</span>
              <span className="text-emerald-300 font-semibold">{afterRows.toLocaleString()}</span>
              {beforeRows !== afterRows && (
                <span className={`text-[11px] ml-1 ${afterRows > beforeRows ? 'text-emerald-400' : 'text-amber-400'}`}>
                  ({afterRows > beforeRows ? `+${afterRows - beforeRows}` : `-${beforeRows - afterRows}`})
                </span>
              )}
            </div>
          )}
        </div>
      )}

      {/* Grid Toolbar */}
      <div className="px-5 py-3 border-b border-slate-800 flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-1.5 text-slate-200 font-medium text-sm">
            <Table className="w-4 h-4 text-emerald-400" />
            <span>{table.name}</span>
          </div>

          <div className="flex items-center gap-2 text-xs text-slate-400">
            <span>·</span>
            <span className="font-mono tabular-nums">{table.data.length.toLocaleString()} rows</span>
            <span>·</span>
            <span className="font-mono tabular-nums">{table.columns.length} columns</span>
          </div>
        </div>

        <div className="flex items-center gap-2.5">
          {/* In-table Search */}
          <div className="relative">
            <Search className="w-3.5 h-3.5 text-slate-500 absolute left-2.5 top-2.5" />
            <input
              type="text"
              value={searchTerm}
              onChange={e => {
                setSearchTerm(e.target.value);
                setPage(1);
              }}
              placeholder="Search in table..."
              className="pl-8 pr-3 py-1.5 bg-slate-950/80 border border-slate-800 rounded-lg text-xs text-slate-200 placeholder-slate-500 focus:outline-none focus:border-slate-700 w-44"
            />
          </div>

          {/* Export Buttons */}
          <button
            onClick={handleDownloadExcel}
            className="flex items-center gap-1.5 px-3 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-medium transition-colors cursor-pointer shadow-sm"
            title="Download full result as styled Excel (.xlsx)"
          >
            <FileSpreadsheet className="w-3.5 h-3.5" />
            <span>Download .xlsx</span>
          </button>

          <button
            onClick={handleDownloadCsv}
            className="flex items-center gap-1.5 px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white border border-slate-700 rounded-lg text-xs transition-colors cursor-pointer"
            title="Download CSV"
          >
            <Download className="w-3.5 h-3.5" />
            <span>CSV</span>
          </button>
        </div>
      </div>

      {/* Spreadsheet Table */}
      <div className="overflow-x-auto max-h-[480px]">
        <table className="w-full text-left border-collapse text-xs">
          <thead>
            <tr className="bg-slate-950/90 sticky top-0 z-10 border-b border-slate-800">
              <th className="py-2.5 px-3 font-mono text-slate-500 border-r border-slate-800/80 w-12 text-center select-none">
                #
              </th>
              {table.columns.map((col, idx) => (
                <th
                  key={idx}
                  className="py-2.5 px-4 font-semibold text-slate-300 whitespace-nowrap border-r border-slate-800/80 last:border-r-0"
                >
                  <div className="flex items-center justify-between gap-2">
                    <span>{col}</span>
                    {highlightRules.some(r => r.column.toLowerCase() === col.toLowerCase()) && (
                      <Sparkles className="w-3 h-3 text-emerald-400 shrink-0" />
                    )}
                  </div>
                </th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-800/60 font-sans">
            {paginatedData.length === 0 ? (
              <tr>
                <td colSpan={table.columns.length + 1} className="py-8 text-center text-slate-500">
                  No matching records found.
                </td>
              </tr>
            ) : (
              paginatedData.map((row, rIdx) => {
                const rowNum = (currentPage - 1) * pageSize + rIdx + 1;
                return (
                  <tr
                    key={rIdx}
                    className="hover:bg-slate-800/50 transition-colors group"
                  >
                    <td className="py-2 px-3 font-mono tabular-nums text-slate-500 text-center border-r border-slate-800/60 select-none bg-slate-950/40">
                      {rowNum}
                    </td>
                    {table.columns.map((col, cIdx) => {
                      const val = row[col];
                      const highlightClass = isHighlighted(col, val);
                      const isNumeric = typeof val === 'number' || (!isNaN(Number(val)) && val !== '' && val !== null && typeof val !== 'boolean');

                      return (
                        <td
                          key={cIdx}
                          className={`py-2 px-4 border-r border-slate-800/60 last:border-r-0 whitespace-nowrap ${
                            highlightClass || 'text-slate-200'
                          } ${isNumeric ? 'font-mono tabular-nums' : ''}`}
                        >
                          {val === null || val === undefined ? (
                            <span className="text-slate-600 italic">null</span>
                          ) : typeof val === 'number' ? (
                            val.toLocaleString()
                          ) : (
                            String(val)
                          )}
                        </td>
                      );
                    })}
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>

      {/* Pagination Footer */}
      <div className="px-5 py-2.5 bg-slate-950/90 border-t border-slate-800 flex items-center justify-between text-xs text-slate-400">
        <div>
          Showing{' '}
          <span className="text-slate-200 font-mono tabular-nums">
            {filteredData.length > 0 ? (currentPage - 1) * pageSize + 1 : 0}
          </span>{' '}
          to{' '}
          <span className="text-slate-200 font-mono tabular-nums">
            {Math.min(currentPage * pageSize, filteredData.length)}
          </span>{' '}
          of{' '}
          <span className="text-slate-200 font-mono tabular-nums">
            {filteredData.length.toLocaleString()}
          </span>{' '}
          rows
        </div>

        <div className="flex items-center gap-1.5">
          <button
            onClick={() => setPage(p => Math.max(p - 1, 1))}
            disabled={currentPage === 1}
            className="p-1.5 bg-slate-800 hover:bg-slate-700 disabled:opacity-30 disabled:cursor-not-allowed rounded text-slate-300 transition-colors"
          >
            <ChevronLeft className="w-3.5 h-3.5" />
          </button>
          <span className="px-2 font-mono tabular-nums text-slate-300">
            Page {currentPage} of {totalPages}
          </span>
          <button
            onClick={() => setPage(p => Math.min(p + 1, totalPages))}
            disabled={currentPage >= totalPages}
            className="p-1.5 bg-slate-800 hover:bg-slate-700 disabled:opacity-30 disabled:cursor-not-allowed rounded text-slate-300 transition-colors"
          >
            <ChevronRight className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>
    </div>
  );
};
