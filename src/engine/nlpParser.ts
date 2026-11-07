import { ActionPlan, WorkspaceTable } from '../types';

export function parseNaturalInstruction(
  input: string,
  tables: Record<string, WorkspaceTable>,
  activeTableId?: string
): ActionPlan {
  const text = input.trim();
  const lower = text.toLowerCase();
  const tableNames = Object.keys(tables);
  const activeTable = activeTableId && tables[activeTableId] ? tables[activeTableId] : tables[tableNames[0]];

  // 1. VLOOKUP / XLOOKUP
  if (lower.includes('vlookup') || lower.includes('xlookup') || lower.includes('lookup') || lower.includes('match karke') || lower.includes('jod do')) {
    // Find mentioned tables
    const mentionedTables = tableNames.filter(name => lower.includes(name.toLowerCase()));
    let leftTbl = mentionedTables[0] || (activeTable ? activeTable.name : tableNames[0]);
    let rightTbl = mentionedTables[1] || (tableNames.find(t => t !== leftTbl) || tableNames[0]);

    if (mentionedTables.length === 0 && tableNames.length >= 2) {
      leftTbl = tableNames[0];
      rightTbl = tableNames[1];
    }

    const tLeft = tables[leftTbl];
    const tRight = tables[rightTbl];

    // Find return columns mentioned in prompt
    let returnCols: string[] = [];
    if (tRight) {
      returnCols = tRight.columns.filter(col => lower.includes(col.toLowerCase()));
      // If none explicitly matched, check for common aliases (e.g. price, unit_price, cost, rate)
      if (returnCols.length === 0) {
        if (lower.includes('price') || lower.includes('kimat') || lower.includes('rate')) {
          const priceCol = tRight.columns.find(c => /price|rate|cost|amount/i.test(c));
          if (priceCol) returnCols.push(priceCol);
        }
        if (lower.includes('category') || lower.includes('shreni')) {
          const catCol = tRight.columns.find(c => /cat/i.test(c));
          if (catCol) returnCols.push(catCol);
        }
      }
    }

    // Find join keys
    let leftKey = '';
    let rightKey = '';

    if (tLeft && tRight) {
      // Find matching keys
      for (const colL of tLeft.columns) {
        if (lower.includes(colL.toLowerCase())) {
          leftKey = colL;
          break;
        }
      }
      for (const colR of tRight.columns) {
        if (lower.includes(colR.toLowerCase())) {
          rightKey = colR;
          break;
        }
      }

      // If keys not explicitly found in prompt, look for overlapping semantic names
      if (!leftKey || !rightKey) {
        const itemColL = tLeft.columns.find(c => /item|product|name|sku|id/i.test(c));
        const itemColR = tRight.columns.find(c => /item|product|name|sku|id/i.test(c));
        if (itemColL && itemColR) {
          leftKey = leftKey || itemColL;
          rightKey = rightKey || itemColR;
        }
      }
    }

    // Clarification check if key or return column is completely missing
    if (!leftKey || !rightKey || returnCols.length === 0) {
      if (tRight && (!returnCols || returnCols.length === 0)) {
        return {
          action: 'clarification',
          needs_clarification: true,
          question_en: `Which column do you want to bring from '${rightTbl}'?`,
          question_hi: `'${rightTbl}' se aap kaun sa column lana chahte hain?`,
          options: tRight.columns
        };
      }
    }

    return {
      action: 'vlookup',
      left_table: leftTbl,
      right_table: rightTbl,
      left_key: leftKey || (tLeft ? tLeft.columns[1] || tLeft.columns[0] : 'Item'),
      right_key: rightKey || (tRight ? tRight.columns[0] : 'Item_Name'),
      return_columns: returnCols.length > 0 ? returnCols : [tRight ? tRight.columns[2] || tRight.columns[1] : 'Unit_Price'],
      fuzzy: lower.includes('fuzzy') || lower.includes('approx') || lower.includes('lagbhag'),
      similarity_threshold: 75,
      output_name: `${leftTbl}_with_${(returnCols[0] || 'Lookup')}`
    };
  }

  // 2. SORT
  if (lower.includes('sort') || lower.includes('kram') || lower.includes('ascending') || lower.includes('descending') || lower.includes('chote se bada') || lower.includes('bade se chota') || lower.includes('tarkeeb')) {
    const isDesc = lower.includes('desc') || lower.includes('bade se') || lower.includes('badha') || lower.includes('bade se chota') || lower.includes('ghat-te');
    let sortCols: string[] = [];

    if (activeTable) {
      sortCols = activeTable.columns.filter(c => lower.includes(c.toLowerCase()));
      if (sortCols.length === 0) {
        // Fallback checks
        if (lower.includes('quantity') || lower.includes('matra')) {
          const qty = activeTable.columns.find(c => /qty|quantity/i.test(c));
          if (qty) sortCols.push(qty);
        } else if (lower.includes('price') || lower.includes('amount') || lower.includes('rakam')) {
          const pr = activeTable.columns.find(c => /price|amount|total/i.test(c));
          if (pr) sortCols.push(pr);
        } else {
          sortCols.push(activeTable.columns[activeTable.columns.length > 3 ? 3 : 0]);
        }
      }
    }

    return {
      action: 'sort',
      table: activeTable ? activeTable.name : tableNames[0],
      columns: sortCols,
      ascending: !isDesc,
      output_name: `${activeTable?.name || 'Data'}_Sorted`
    };
  }

  // 3. COMMON / UNIQUE ITEMS
  if (lower.includes('common') || lower.includes('similar') || lower.includes('dono table') || lower.includes('milte julte') || lower.includes('match hone wale') || lower.includes('unique to')) {
    const mentionedTables = tableNames.filter(name => lower.includes(name.toLowerCase()));
    const tblA = mentionedTables[0] || tableNames[0];
    const tblB = mentionedTables[1] || tableNames[1] || tableNames[0];

    const tA = tables[tblA];
    const tB = tables[tblB];

    let colA = '';
    let colB = '';

    if (tA && tB) {
      colA = tA.columns.find(c => lower.includes(c.toLowerCase())) || tA.columns.find(c => /item|product|name/i.test(c)) || tA.columns[1] || tA.columns[0];
      colB = tB.columns.find(c => lower.includes(c.toLowerCase())) || tB.columns.find(c => /product|item|name/i.test(c)) || tB.columns[0];
    }

    const isUniqueA = lower.includes('unique to a') || lower.includes('sirf pehle') || lower.includes('only in first');
    const isUniqueB = lower.includes('unique to b') || lower.includes('sirf doosre') || lower.includes('only in second');

    return {
      action: 'find_common',
      table_a: tblA,
      table_b: tblB,
      col_a: colA,
      col_b: colB,
      mode: isUniqueA ? 'unique_to_a' : isUniqueB ? 'unique_to_b' : 'common',
      fuzzy: lower.includes('fuzzy') || lower.includes('approx') || lower.includes('lagbhag'),
      similarity_threshold: 75,
      output_name: `Common_${tblA}_${tblB}`
    };
  }

  // 4. CLEAN DATA (Duplicates, Blanks, Trimming)
  if (lower.includes('duplicate') || lower.includes('dedup') || lower.includes('hata do') || lower.includes('khali') || lower.includes('blank') || lower.includes('trim') || lower.includes('uppercase') || lower.includes('lowercase')) {
    let operation: ActionPlan['operation'] = 'remove_duplicates';
    let fillVal: any = null;

    if (lower.includes('khali') || lower.includes('blank')) {
      if (lower.includes('bhar') || lower.includes('fill')) {
        operation = 'fill_blanks';
        fillVal = 0;
      } else {
        operation = 'drop_blanks';
      }
    } else if (lower.includes('trim') || lower.includes('space')) {
      operation = 'trim_spaces';
    } else if (lower.includes('uppercase') || lower.includes('bada akshar')) {
      operation = 'change_case';
      fillVal = 'upper';
    } else if (lower.includes('lowercase') || lower.includes('chota akshar')) {
      operation = 'change_case';
      fillVal = 'lower';
    }

    return {
      action: 'clean',
      table: activeTable ? activeTable.name : tableNames[0],
      operation,
      fill_value: fillVal,
      output_name: `${activeTable?.name || 'Data'}_Cleaned`
    };
  }

  // 5. PIVOT TABLE & GROUPBY
  if (lower.includes('pivot') || lower.includes('groupby') || lower.includes('group by') || lower.includes('summary') || lower.includes('ke hisab se') || lower.includes('kul jodh')) {
    let indexCols: string[] = [];
    let valueCols: string[] = [];

    if (activeTable) {
      // Find candidate index columns (e.g. City, Category, Customer, Warehouse)
      const catCols = activeTable.columns.filter(c => /city|cat|customer|warehouse|state/i.test(c));
      indexCols = activeTable.columns.filter(c => lower.includes(c.toLowerCase()));
      if (indexCols.length === 0 && catCols.length > 0) {
        indexCols.push(catCols[0]);
      } else if (indexCols.length === 0) {
        indexCols.push(activeTable.columns[activeTable.columns.length > 4 ? 4 : 1]);
      }

      // Find value columns (e.g. Quantity, Amount, Unit_Price, Stock_Count)
      const numCols = activeTable.columns.filter(c => /qty|quantity|amount|total|price|stock/i.test(c));
      valueCols = numCols.filter(c => lower.includes(c.toLowerCase()));
      if (valueCols.length === 0 && numCols.length > 0) {
        valueCols.push(numCols[0]);
      }
    }

    let agg: 'sum' | 'count' | 'average' | 'min' | 'max' = 'sum';
    if (lower.includes('count') || lower.includes('ginati')) agg = 'count';
    else if (lower.includes('avg') || lower.includes('average') || lower.includes('ausat')) agg = 'average';
    else if (lower.includes('min') || lower.includes('kam se kam')) agg = 'min';
    else if (lower.includes('max') || lower.includes('zyada se zyada')) agg = 'max';

    return {
      action: 'pivot',
      table: activeTable ? activeTable.name : tableNames[0],
      index_columns: indexCols.length > 0 ? indexCols : ['City'],
      values_columns: valueCols.length > 0 ? valueCols : ['Quantity'],
      agg_func: agg,
      output_name: `Pivot_by_${indexCols[0] || 'Group'}`
    };
  }

  // 6. CALCULATED COLUMNS
  if (lower.includes('calculate') || lower.includes('formula') || lower.includes('multiply') || lower.includes('guna') || lower.includes('add column') || lower.includes('jod do') || lower.includes('total =')) {
    // Check if user specified a formula like "Total = Quantity * Unit_Price"
    const formulaMatch = text.match(/([a-zA-Z0-9_]+)\s*=\s*(.+)/);
    let targetCol = 'Total';
    let formulaStr = 'Quantity * Unit_Price';

    if (formulaMatch) {
      targetCol = formulaMatch[1].trim();
      formulaStr = formulaMatch[2].trim();
    } else if (activeTable) {
      // Find two numeric columns to multiply
      const nums = activeTable.columns.filter(c => /qty|quantity|price|rate|cost|amount/i.test(c));
      if (nums.length >= 2) {
        formulaStr = `${nums[0]} * ${nums[1]}`;
      }
    }

    return {
      action: 'calculate',
      table: activeTable ? activeTable.name : tableNames[0],
      target_column: targetCol,
      formula: formulaStr,
      output_name: `${activeTable?.name || 'Data'}_Calculated`
    };
  }

  // 7. FILTER
  if (lower.includes('filter') || lower.includes('chhaanto') || lower.includes('greater than') || lower.includes('less than') || lower.includes('se zyada') || lower.includes('se kam') || lower.includes('jisme') || lower.includes('> ') || lower.includes('< ')) {
    let col = activeTable ? activeTable.columns[3] || activeTable.columns[0] : 'Quantity';
    let op: ActionPlan['operator'] = 'greater_than';
    let val: any = 1;

    // Check operators
    if (lower.includes('<') || lower.includes('less than') || lower.includes('kam')) op = 'less_than';
    else if (lower.includes('==') || lower.includes('equal') || lower.includes('barabar')) op = 'equals';
    else if (lower.includes('contains') || lower.includes('jisme')) op = 'contains';

    // Extract numbers
    const numMatch = text.match(/\d+(\.\d+)?/);
    if (numMatch) val = parseFloat(numMatch[0]);

    if (activeTable) {
      const matched = activeTable.columns.find(c => lower.includes(c.toLowerCase()));
      if (matched) col = matched;
    }

    return {
      action: 'filter',
      table: activeTable ? activeTable.name : tableNames[0],
      column: col,
      operator: op,
      value: val,
      output_name: `${activeTable?.name || 'Data'}_Filtered`
    };
  }

  // Default fallback: VLOOKUP or Sorting based on available tables
  if (tableNames.length >= 2) {
    return {
      action: 'vlookup',
      left_table: tableNames[0],
      right_table: tableNames[1],
      left_key: tables[tableNames[0]].columns[1] || 'Item',
      right_key: tables[tableNames[1]].columns[0] || 'Item_Name',
      return_columns: [tables[tableNames[1]].columns[2] || 'Unit_Price'],
      output_name: 'VLookup_Result'
    };
  }

  return {
    action: 'sort',
    table: tableNames[0],
    columns: [activeTable?.columns[3] || 'Quantity'],
    ascending: false,
    output_name: `${activeTable?.name || 'Data'}_Sorted`
  };
}
