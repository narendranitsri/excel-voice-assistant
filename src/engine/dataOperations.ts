import * as XLSX from 'xlsx';
import { WorkspaceTable, TableData, ActionPlan, HighlightRule } from '../types';

export interface OperationResult {
  table: WorkspaceTable;
  statusMessage: string;
  statusMessageHi: string;
  beforeRows: number;
  afterRows: number;
  highlightRules?: HighlightRule[];
}

// Key normalizer (whitespace stripped, case-insensitive, float integer safe)
export function normalizeKey(val: any): string {
  if (val === null || val === undefined) return '';
  const num = Number(val);
  if (!isNaN(num) && typeof val === 'number') {
    return Number.isInteger(num) ? String(num) : String(val).trim().toLowerCase();
  }
  return String(val).trim().toLowerCase();
}

// Case-insensitive column finder
export function findColumnName(table: WorkspaceTable, target: string): string | null {
  const cleanTarget = target.trim().toLowerCase();
  for (const col of table.columns) {
    if (col.trim().toLowerCase() === cleanTarget) {
      return col;
    }
  }
  return null;
}

// Simple Levenshtein distance & similarity ratio for fuzzy matching
export function calculateFuzzySimilarity(a: string, b: string): number {
  const s1 = a.trim().toLowerCase();
  const s2 = b.trim().toLowerCase();
  if (s1 === s2) return 100;
  if (!s1.length || !s2.length) return 0;

  const matrix: number[][] = [];
  for (let i = 0; i <= s2.length; i++) matrix[i] = [i];
  for (let j = 0; j <= s1.length; j++) matrix[0][j] = j;

  for (let i = 1; i <= s2.length; i++) {
    for (let j = 1; j <= s1.length; j++) {
      if (s2.charAt(i - 1) === s1.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j] + 1
        );
      }
    }
  }
  const distance = matrix[s2.length][s1.length];
  const maxLen = Math.max(s1.length, s2.length);
  return Math.round(((maxLen - distance) / maxLen) * 100);
}

// 1. VLOOKUP / XLOOKUP
export function executeVLookup(
  leftTable: WorkspaceTable,
  rightTable: WorkspaceTable,
  leftKey: string,
  rightKey: string,
  returnColumns: string[],
  fuzzy = false,
  threshold = 75,
  outputName = 'VLookup_Result'
): OperationResult {
  const actualLeftKey = findColumnName(leftTable, leftKey);
  const actualRightKey = findColumnName(rightTable, rightKey);

  if (!actualLeftKey) {
    throw new Error(`Left column '${leftKey}' not found in table '${leftTable.name}'. Available: ${leftTable.columns.join(', ')}`);
  }
  if (!actualRightKey) {
    throw new Error(`Right column '${rightKey}' not found in table '${rightTable.name}'. Available: ${rightTable.columns.join(', ')}`);
  }

  // Validate return columns
  const actualReturnCols: string[] = [];
  for (const rc of returnColumns) {
    const matched = findColumnName(rightTable, rc);
    if (matched && matched !== actualRightKey) {
      actualReturnCols.push(matched);
    }
  }

  if (actualReturnCols.length === 0) {
    throw new Error(`Return columns [${returnColumns.join(', ')}] not found in table '${rightTable.name}'. Available: ${rightTable.columns.join(', ')}`);
  }

  // Build lookup index for right table
  const rightMap = new Map<string, Record<string, any>>();
  for (const row of rightTable.data) {
    const norm = normalizeKey(row[actualRightKey]);
    if (norm && !rightMap.has(norm)) {
      rightMap.set(norm, row);
    }
  }

  const rightKeysList = Array.from(rightMap.keys());
  const beforeRows = leftTable.data.length;

  const newData: TableData = leftTable.data.map((leftRow) => {
    const leftValNorm = normalizeKey(leftRow[actualLeftKey]);
    let matchedRightRow: Record<string, any> | undefined = undefined;

    if (leftValNorm && rightMap.has(leftValNorm)) {
      matchedRightRow = rightMap.get(leftValNorm);
    } else if (fuzzy && leftValNorm) {
      // Find best fuzzy match
      let bestMatchKey = '';
      let highestScore = 0;
      for (const rk of rightKeysList) {
        const score = calculateFuzzySimilarity(leftValNorm, rk);
        if (score >= threshold && score > highestScore) {
          highestScore = score;
          bestMatchKey = rk;
        }
      }
      if (bestMatchKey) {
        matchedRightRow = rightMap.get(bestMatchKey);
      }
    }

    const mergedRow = { ...leftRow };
    for (const retCol of actualReturnCols) {
      const outputColName = leftTable.columns.includes(retCol) ? `${retCol}_lookup` : retCol;
      mergedRow[outputColName] = matchedRightRow ? matchedRightRow[retCol] : null;
    }
    return mergedRow;
  });

  const newColumns = [...leftTable.columns];
  for (const retCol of actualReturnCols) {
    const outputColName = leftTable.columns.includes(retCol) ? `${retCol}_lookup` : retCol;
    if (!newColumns.includes(outputColName)) {
      newColumns.push(outputColName);
    }
  }

  const newTable: WorkspaceTable = {
    id: `result_${Date.now()}`,
    name: outputName,
    columns: newColumns,
    data: newData,
    rowCount: newData.length
  };

  return {
    table: newTable,
    statusMessage: `VLOOKUP matched '${actualLeftKey}' with '${actualRightKey}'. Added columns: ${actualReturnCols.join(', ')}.`,
    statusMessageHi: `VLOOKUP safaltapoorvak poora hua: '${actualLeftKey}' se match karke columns [${actualReturnCols.join(', ')}] jode gaye.`,
    beforeRows,
    afterRows: newData.length,
    highlightRules: actualReturnCols.map(c => ({
      type: 'lookup_match',
      column: leftTable.columns.includes(c) ? `${c}_lookup` : c,
      color: 'green'
    }))
  };
}

// 2. SORT
export function executeSort(
  table: WorkspaceTable,
  columns: string[],
  ascending: boolean | boolean[] = true,
  outputName?: string
): OperationResult {
  const actualCols: string[] = [];
  for (const c of columns) {
    const matched = findColumnName(table, c);
    if (!matched) {
      throw new Error(`Sort column '${c}' not found in '${table.name}'. Available: ${table.columns.join(', ')}`);
    }
    actualCols.push(matched);
  }

  const sortedData = [...table.data].sort((rowA, rowB) => {
    for (let i = 0; i < actualCols.length; i++) {
      const col = actualCols[i];
      const isAsc = Array.isArray(ascending) ? (ascending[i] ?? true) : ascending;
      const valA = rowA[col];
      const valB = rowB[col];

      if (valA === valB) continue;
      if (valA === null || valA === undefined) return isAsc ? 1 : -1;
      if (valB === null || valB === undefined) return isAsc ? -1 : 1;

      // Numeric comparison
      const numA = Number(valA);
      const numB = Number(valB);
      if (!isNaN(numA) && !isNaN(numB)) {
        return isAsc ? numA - numB : numB - numA;
      }

      // String comparison
      const strA = String(valA).toLowerCase();
      const strB = String(valB).toLowerCase();
      const comp = strA.localeCompare(strB);
      return isAsc ? comp : -comp;
    }
    return 0;
  });

  const ascText = (Array.isArray(ascending) ? ascending[0] : ascending) ? 'ascending (A-Z, 0-9)' : 'descending (Z-A, 9-0)';

  return {
    table: {
      ...table,
      name: outputName || `${table.name}_Sorted`,
      data: sortedData
    },
    statusMessage: `Sorted by ${actualCols.join(', ')} in ${ascText} order.`,
    statusMessageHi: `Data ko ${actualCols.join(', ')} ke hisab se ${ascText} order me sort kar diya gaya.`,
    beforeRows: table.data.length,
    afterRows: sortedData.length
  };
}

// 3. FIND COMMON / UNIQUE ITEMS
export function executeFindSimilarOrCommon(
  tableA: WorkspaceTable,
  tableB: WorkspaceTable,
  colA: string,
  colB: string,
  mode: 'common' | 'unique_to_a' | 'unique_to_b' = 'common',
  fuzzy = false,
  threshold = 75,
  outputName = 'Common_Items'
): OperationResult {
  const actualColA = findColumnName(tableA, colA);
  const actualColB = findColumnName(tableB, colB);

  if (!actualColA) throw new Error(`Column '${colA}' not found in '${tableA.name}'.`);
  if (!actualColB) throw new Error(`Column '${colB}' not found in '${tableB.name}'.`);

  const setB = new Set<string>();
  tableB.data.forEach(r => {
    const k = normalizeKey(r[actualColB]);
    if (k) setB.add(k);
  });
  const listB = Array.from(setB);

  const matchedKeys = new Set<string>();
  tableA.data.forEach(r => {
    const k = normalizeKey(r[actualColA]);
    if (!k) return;
    if (setB.has(k)) {
      matchedKeys.add(k);
    } else if (fuzzy) {
      for (const itemB of listB) {
        if (calculateFuzzySimilarity(k, itemB) >= threshold) {
          matchedKeys.add(k);
          break;
        }
      }
    }
  });

  let resultData: TableData = [];
  let sourceTable = tableA;
  let activeCol = actualColA;

  if (mode === 'common') {
    resultData = tableA.data.filter(r => matchedKeys.has(normalizeKey(r[actualColA])));
  } else if (mode === 'unique_to_a') {
    resultData = tableA.data.filter(r => !matchedKeys.has(normalizeKey(r[actualColA])));
  } else if (mode === 'unique_to_b') {
    sourceTable = tableB;
    activeCol = actualColB;
    const setA = new Set(tableA.data.map(r => normalizeKey(r[actualColA])));
    resultData = tableB.data.filter(r => !setA.has(normalizeKey(r[actualColB])));
  }

  const newTable: WorkspaceTable = {
    id: `result_${Date.now()}`,
    name: outputName,
    columns: [...sourceTable.columns],
    data: resultData,
    rowCount: resultData.length
  };

  const modeText = mode === 'common' ? 'common items' : mode === 'unique_to_a' ? 'unique to Table A' : 'unique to Table B';

  return {
    table: newTable,
    statusMessage: `Found ${resultData.length} ${modeText} matching '${activeCol}'.`,
    statusMessageHi: `${resultData.length} ${modeText} nikal liye gaye.`,
    beforeRows: sourceTable.data.length,
    afterRows: resultData.length,
    highlightRules: [{ type: 'lookup_match', column: activeCol, color: 'green' }]
  };
}

// 4. FILTER
export function executeFilter(
  table: WorkspaceTable,
  column: string,
  operator: string,
  value: any,
  value2?: any,
  outputName?: string
): OperationResult {
  const actualCol = findColumnName(table, column);
  if (!actualCol) throw new Error(`Filter column '${column}' not found in '${table.name}'.`);

  const op = operator.toLowerCase();
  const beforeRows = table.data.length;

  const filteredData = table.data.filter(row => {
    const rawVal = row[actualCol];
    if (rawVal === null || rawVal === undefined) return false;

    if (op === 'equals' || op === '==') {
      return String(rawVal).trim().toLowerCase() === String(value).trim().toLowerCase();
    }
    if (op === 'not_equals' || op === '!=') {
      return String(rawVal).trim().toLowerCase() !== String(value).trim().toLowerCase();
    }
    if (op === 'contains') {
      return String(rawVal).toLowerCase().includes(String(value).toLowerCase());
    }
    if (op === 'not_contains') {
      return !String(rawVal).toLowerCase().includes(String(value).toLowerCase());
    }
    const num = Number(rawVal);
    const targetNum = Number(value);
    if (!isNaN(num) && !isNaN(targetNum)) {
      if (op === 'greater_than' || op === '>') return num > targetNum;
      if (op === 'less_than' || op === '<') return num < targetNum;
      if (op === 'greater_equal' || op === '>=') return num >= targetNum;
      if (op === 'less_equal' || op === '<=') return num <= targetNum;
      if (op === 'between') {
        const targetNum2 = Number(value2 ?? targetNum);
        return num >= targetNum && num <= targetNum2;
      }
    }
    return false;
  });

  return {
    table: {
      ...table,
      name: outputName || `${table.name}_Filtered`,
      data: filteredData,
      rowCount: filteredData.length
    },
    statusMessage: `Filtered ${filteredData.length} of ${beforeRows} rows where '${actualCol}' ${operator} ${value}.`,
    statusMessageHi: `Rows filter kar di gayi (${filteredData.length} bachi hain).`,
    beforeRows,
    afterRows: filteredData.length
  };
}

// 5. CLEANING
export function executeCleaning(
  table: WorkspaceTable,
  operation: string,
  columns?: string[],
  fillValue?: any,
  outputName?: string
): OperationResult {
  const op = operation.toLowerCase();
  const beforeRows = table.data.length;
  let cleanData = [...table.data];
  const targetCols = columns && columns.length ? columns.map(c => findColumnName(table, c)).filter(Boolean) as string[] : table.columns;

  let msg = '';
  let msgHi = '';

  if (op === 'remove_duplicates' || op === 'dedup') {
    const seen = new Set<string>();
    cleanData = cleanData.filter(row => {
      const hash = targetCols.map(c => normalizeKey(row[c])).join('|||');
      if (seen.has(hash)) return false;
      seen.add(hash);
      return true;
    });
    const removed = beforeRows - cleanData.length;
    msg = `Removed ${removed} duplicate rows. Remaining: ${cleanData.length}.`;
    msgHi = `${removed} duplicate rows hata di gayi. Bachi hui: ${cleanData.length}.`;
  } else if (op === 'drop_blanks') {
    cleanData = cleanData.filter(row => {
      return targetCols.every(c => row[c] !== null && row[c] !== undefined && String(row[c]).trim() !== '');
    });
    const removed = beforeRows - cleanData.length;
    msg = `Dropped ${removed} blank rows.`;
    msgHi = `${removed} blank rows hata di gayi.`;
  } else if (op === 'fill_blanks') {
    const fillWith = fillValue ?? 0;
    cleanData = cleanData.map(row => {
      const newRow = { ...row };
      for (const col of targetCols) {
        if (newRow[col] === null || newRow[col] === undefined || String(newRow[col]).trim() === '') {
          newRow[col] = fillWith;
        }
      }
      return newRow;
    });
    msg = `Filled blanks in [${targetCols.join(', ')}] with '${fillWith}'.`;
    msgHi = `Khali cells me '${fillWith}' bhar diya gaya.`;
  } else if (op === 'trim_spaces') {
    cleanData = cleanData.map(row => {
      const newRow = { ...row };
      for (const col of targetCols) {
        if (typeof newRow[col] === 'string') {
          newRow[col] = newRow[col].trim();
        }
      }
      return newRow;
    });
    msg = `Trimmed whitespace in columns: ${targetCols.join(', ')}.`;
    msgHi = `Extra spaces hata diye gaye.`;
  } else if (op === 'change_case') {
    const caseType = String(fillValue || 'title').toLowerCase();
    cleanData = cleanData.map(row => {
      const newRow = { ...row };
      for (const col of targetCols) {
        if (typeof newRow[col] === 'string') {
          if (caseType === 'upper') newRow[col] = newRow[col].toUpperCase();
          else if (caseType === 'lower') newRow[col] = newRow[col].toLowerCase();
          else newRow[col] = newRow[col].replace(/\w\S*/g, (txt: string) => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase());
        }
      }
      return newRow;
    });
    msg = `Converted text case to '${caseType}'.`;
    msgHi = `Text case ko '${caseType}' me badal diya gaya.`;
  }

  return {
    table: {
      ...table,
      name: outputName || `${table.name}_Cleaned`,
      data: cleanData,
      rowCount: cleanData.length
    },
    statusMessage: msg,
    statusMessageHi: msgHi,
    beforeRows,
    afterRows: cleanData.length
  };
}

// 6. PIVOT & GROUPBY
export function executePivotTable(
  table: WorkspaceTable,
  indexColumns: string[],
  valuesColumns: string[],
  aggFunc: 'sum' | 'count' | 'average' | 'min' | 'max' = 'sum',
  outputName = 'Pivot_Summary'
): OperationResult {
  const actualIndexCols = indexColumns.map(c => findColumnName(table, c)).filter(Boolean) as string[];
  const actualValueCols = valuesColumns.map(c => findColumnName(table, c)).filter(Boolean) as string[];

  if (actualIndexCols.length === 0) throw new Error(`Index columns [${indexColumns.join(', ')}] not found.`);
  if (actualValueCols.length === 0) throw new Error(`Value columns [${valuesColumns.join(', ')}] not found.`);

  // Group by index keys
  const groups = new Map<string, Record<string, any>[]>();
  for (const row of table.data) {
    const key = actualIndexCols.map(c => String(row[c] ?? '')).join('|||');
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key)!.push(row);
  }

  const pivotData: TableData = [];
  groups.forEach((rows, key) => {
    const keyParts = key.split('|||');
    const pivotRow: Record<string, any> = {};
    actualIndexCols.forEach((col, idx) => {
      pivotRow[col] = keyParts[idx];
    });

    actualValueCols.forEach(valCol => {
      const numbers = rows.map(r => Number(r[valCol])).filter(n => !isNaN(n));
      const colLabel = `${aggFunc.toUpperCase()}_of_${valCol}`;

      if (aggFunc === 'count') {
        pivotRow[colLabel] = rows.length;
      } else if (numbers.length === 0) {
        pivotRow[colLabel] = 0;
      } else if (aggFunc === 'sum') {
        pivotRow[colLabel] = Math.round(numbers.reduce((a, b) => a + b, 0) * 100) / 100;
      } else if (aggFunc === 'average') {
        const sum = numbers.reduce((a, b) => a + b, 0);
        pivotRow[colLabel] = Math.round((sum / numbers.length) * 100) / 100;
      } else if (aggFunc === 'min') {
        pivotRow[colLabel] = Math.min(...numbers);
      } else if (aggFunc === 'max') {
        pivotRow[colLabel] = Math.max(...numbers);
      }
    });
    pivotData.push(pivotRow);
  });

  const newColumns = [...actualIndexCols, ...actualValueCols.map(c => `${aggFunc.toUpperCase()}_of_${c}`)];

  const newTable: WorkspaceTable = {
    id: `pivot_${Date.now()}`,
    name: outputName,
    columns: newColumns,
    data: pivotData,
    rowCount: pivotData.length
  };

  return {
    table: newTable,
    statusMessage: `Created pivot summary grouped by [${actualIndexCols.join(', ')}] with ${aggFunc} of [${actualValueCols.join(', ')}].`,
    statusMessageHi: `Pivot table ban gayi: [${actualIndexCols.join(', ')}] par ${aggFunc} lagaya gaya.`,
    beforeRows: table.data.length,
    afterRows: pivotData.length
  };
}

// 7. CALCULATE / TRANSFORM COLUMN
export function executeColumnTransform(
  table: WorkspaceTable,
  action: string,
  targetColumn: string,
  formula?: string,
  sourceColumns?: string[],
  delimiter?: string,
  outputName?: string
): OperationResult {
  const beforeRows = table.data.length;
  let newData = [...table.data];
  let newColumns = [...table.columns];

  if (action === 'calculate') {
    if (!formula) throw new Error('Formula is required for calculated column.');

    // Safe mathematical expression evaluation
    // Supports patterns like: "Quantity * Unit_Price", "A + B", "Price * (1 + GST / 100)"
    newData = newData.map(row => {
      let evalStr = formula;
      for (const col of table.columns) {
        const regex = new RegExp(`\\b${col}\\b`, 'gi');
        if (regex.test(evalStr)) {
          const val = Number(row[col]) || 0;
          evalStr = evalStr.replace(regex, String(val));
        }
      }

      // Safe character guard
      if (!/^[0-9+\-*/().\s]+$/.test(evalStr)) {
        throw new Error(`Formula contains invalid characters: '${formula}'`);
      }

      let resVal = 0;
      try {
        // Safe evaluation of pure arithmetic numbers
        resVal = Function(`'use strict'; return (${evalStr});`)();
      } catch {
        resVal = 0;
      }

      return {
        ...row,
        [targetColumn]: typeof resVal === 'number' ? Math.round(resVal * 100) / 100 : resVal
      };
    });

    if (!newColumns.includes(targetColumn)) newColumns.push(targetColumn);

    return {
      table: {
        ...table,
        name: outputName || table.name,
        columns: newColumns,
        data: newData
      },
      statusMessage: `Added calculated column '${targetColumn}' = ${formula}.`,
      statusMessageHi: `Naya column '${targetColumn}' formula '${formula}' se calculate kar diya gaya.`,
      beforeRows,
      afterRows: newData.length,
      highlightRules: [{ type: 'lookup_match', column: targetColumn, color: 'green' }]
    };
  }

  throw new Error(`Unsupported column transform action: ${action}`);
}

// 8. MERGE SHEETS
export function executeMergeSheets(
  tables: WorkspaceTable[],
  mergeType = 'append',
  outputName = 'Merged_Sheets'
): OperationResult {
  if (tables.length === 0) throw new Error('No sheets provided to merge.');

  const allColumns = new Set<string>();
  tables.forEach(t => t.columns.forEach(c => allColumns.add(c)));
  allColumns.add('_SourceSheet');

  const mergedData: TableData = [];
  tables.forEach(t => {
    t.data.forEach(r => {
      mergedData.push({
        ...r,
        _SourceSheet: t.name
      });
    });
  });

  const beforeRows = tables.reduce((sum, t) => sum + t.data.length, 0);

  const newTable: WorkspaceTable = {
    id: `merged_${Date.now()}`,
    name: outputName,
    columns: Array.from(allColumns),
    data: mergedData,
    rowCount: mergedData.length
  };

  return {
    table: newTable,
    statusMessage: `Appended ${tables.length} tables into '${outputName}'. Total rows: ${mergedData.length}.`,
    statusMessageHi: `${tables.length} sheets ko ek sath jod diya gaya. Kul rows: ${mergedData.length}.`,
    beforeRows,
    afterRows: mergedData.length
  };
}

// Export workbook to XLSX Blob
export function exportTableToXlsxBlob(table: WorkspaceTable): Blob {
  const ws = XLSX.utils.json_to_sheet(table.data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, table.name.slice(0, 31));
  const wbOut = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  return new Blob([wbOut], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
