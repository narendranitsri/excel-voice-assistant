export type TableData = Record<string, any>[];

export interface WorkspaceTable {
  id: string;
  name: string;
  fileName?: string;
  columns: string[];
  data: TableData;
  rowCount: number;
}

export type ActionType =
  | 'vlookup'
  | 'sort'
  | 'find_common'
  | 'filter'
  | 'clean'
  | 'pivot'
  | 'calculate'
  | 'merge_sheets'
  | 'clarification';

export interface ActionPlan {
  action: ActionType;
  output_name?: string;
  // VLOOKUP
  left_table?: string;
  right_table?: string;
  left_key?: string;
  right_key?: string;
  return_columns?: string[];
  fuzzy?: boolean;
  similarity_threshold?: number;
  // SORT
  table?: string;
  columns?: string[];
  ascending?: boolean | boolean[];
  // FIND COMMON / UNIQUE
  table_a?: string;
  table_b?: string;
  col_a?: string;
  col_b?: string;
  mode?: 'common' | 'unique_to_a' | 'unique_to_b';
  // FILTER
  column?: string;
  operator?: 'equals' | 'not_equals' | 'contains' | 'not_contains' | 'greater_than' | 'less_than' | 'between' | 'date_range';
  value?: any;
  value2?: any;
  // CLEAN
  operation?: 'remove_duplicates' | 'drop_blanks' | 'fill_blanks' | 'trim_spaces' | 'change_case';
  fill_value?: any;
  // PIVOT
  index_columns?: string[];
  values_columns?: string[];
  agg_func?: 'sum' | 'count' | 'average' | 'min' | 'max';
  columns_pivot?: string[];
  // CALCULATE
  target_column?: string;
  source_columns?: string[];
  formula?: string;
  delimiter?: string;
  // MERGE SHEETS
  sheets?: string[];
  // CLARIFICATION
  needs_clarification?: boolean;
  question_en?: string;
  question_hi?: string;
  options?: string[];
}

export interface HistoryItem {
  id: string;
  timestamp: string;
  description: string;
  descriptionHi?: string;
  actionPlan: ActionPlan;
  beforeRows: number;
  afterRows: number;
  tablesSnapshot: Record<string, WorkspaceTable>;
  activeTableId: string;
}

export interface HighlightRule {
  type: 'duplicates' | 'condition' | 'lookup_match';
  column: string;
  color: 'yellow' | 'green' | 'coral';
  conditionValue?: any;
  operator?: string;
}
