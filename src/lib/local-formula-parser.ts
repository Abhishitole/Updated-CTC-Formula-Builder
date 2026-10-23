import * as XLSX from 'xlsx';
import { PayHeadFormula, normalizeFormulaSyntax } from '../services/gemini';

export interface ParsedCell {
  col: string;
  colIndex: number;
  row: number;
  address: string;
  value: any;
  formula?: string;
}

export interface ExtractedPayHeadRow {
  rowNumber: number;
  payHead: string;
  value: any;
  formula?: string;
  colIndex: number;
  valueColIndex: number;
  cellAddress: string;
}

/**
 * Splits formula arguments at top level by comma, taking into account
 * nested parentheses and quotation marks.
 */
export function splitTopLevelArgs(inner: string): string[] {
  const args: string[] = [];
  let current = '';
  let parenDepth = 0;
  let inQuotes = false;
  let quoteChar = '';

  for (let i = 0; i < inner.length; i++) {
    const ch = inner[i];
    if ((ch === '"' || ch === "'") && (i === 0 || inner[i - 1] !== '\\')) {
      if (!inQuotes) {
        inQuotes = true;
        quoteChar = ch;
      } else if (ch === quoteChar) {
        inQuotes = false;
      }
      current += ch;
      continue;
    }

    if (inQuotes) {
      current += ch;
      continue;
    }

    if (ch === '(') {
      parenDepth++;
      current += ch;
    } else if (ch === ')') {
      parenDepth--;
      current += ch;
    } else if (ch === ',' && parenDepth === 0) {
      args.push(current.trim());
      current = '';
    } else {
      current += ch;
    }
  }

  if (current.trim()) {
    args.push(current.trim());
  }

  return args;
}

/**
 * Checks if a string is a standard Excel column header (e.g. "Sl No", "Particulars")
 * to avoid treating table header rows as salary components.
 */
function isCommonHeaderLabel(label: string): boolean {
  const clean = label.trim().toLowerCase();
  const headers = [
    'sl no', 'sl.no', 'sr no', 'sr.no', 's.no', 's no', 'serial no', 'no.', 'no',
    'salary component', 'salary components', 'particulars', 'component', 'components',
    'pay head', 'pay heads', 'payhead', 'payheads', 'earning', 'earnings',
    'deduction', 'deductions', 'monthly', 'monthly amount', 'per month', 'pm', 'p.m.',
    'annual', 'yearly', 'annual amount', 'per annum', 'pa', 'p.a.',
    'formula', 'calculation', 'percentage', 'percent', 'remarks', 'remark',
    'description', 'notes', 'type', 'amount', 'amt', 'value', 'values', 'head', 'heads',
    'ctc breakup', 'annexure'
  ];
  return headers.includes(clean);
}

/**
 * Normalizes text to a valid JavaScript variable name
 */
export function toJsVariableName(label: string): string {
  let clean = label
    .trim()
    .replace(/[^a-zA-Z0-9_ ]/g, '')
    .replace(/\s+/g, '_');

  // Strip leading numbers or underscores
  clean = clean.replace(/^[0-9_]+/, '');
  if (!clean) clean = 'Component';
  return clean;
}

/**
 * Parses raw Excel formula into human-readable CASE WHEN syntax
 * and substitutes cell references with Pay Head names.
 */
export function translateExcelFormula(
  rawFormula: string,
  cellToHead: Map<string, string>,
  rowToHead: Map<number, string>,
  allHeads: { payHead: string; jsVariableName: string; isBase: boolean }[]
): { readableFormula: string; jsExpression: string; description: string } {
  let f = rawFormula.trim();
  if (f.startsWith('=')) {
    f = f.substring(1).trim();
  }

  // Helper to replace cell references with Pay Head names
  const replaceCellRefs = (expr: string, useJsNames: boolean = false): string => {
    // Replace ranges like B2:B5 in SUM
    expr = expr.replace(/\b([A-Z]+)(\d+)\s*:\s*([A-Z]+)(\d+)\b/gi, (match, col1, row1, col2, row2) => {
      const r1 = parseInt(row1, 10);
      const r2 = parseInt(row2, 10);
      const start = Math.min(r1, r2);
      const end = Math.max(r1, r2);
      const headNames: string[] = [];

      for (let r = start; r <= end; r++) {
        const head = rowToHead.get(r) || cellToHead.get(`${col1.toUpperCase()}${r}`);
        if (head) {
          if (useJsNames) {
            const found = allHeads.find(h => h.payHead === head);
            headNames.push(found ? found.jsVariableName : toJsVariableName(head));
          } else {
            headNames.push(head);
          }
        }
      }

      if (headNames.length > 0) {
        return headNames.join(' + ');
      }
      return match;
    });

    // Replace individual cell references like B2, $B$2, C10
    expr = expr.replace(/\$?[A-Z]+\$?(\d+)/gi, (match, rowStr) => {
      const cleanCell = match.replace(/\$/g, '').toUpperCase();
      const rowNum = parseInt(rowStr, 10);
      const head = cellToHead.get(cleanCell) || rowToHead.get(rowNum);
      if (head) {
        if (useJsNames) {
          const found = allHeads.find(h => h.payHead === head);
          return found ? found.jsVariableName : toJsVariableName(head);
        }
        return head;
      }
      return match;
    });

    return expr;
  };

  // Convert percentages: 40% -> 0.4
  f = f.replace(/(\d+(?:\.\d+)?)%/g, (match, num) => {
    const val = parseFloat(num) / 100;
    return val.toString();
  });

  // Handle SUM functions
  f = f.replace(/\bSUM\s*\(([^)]+)\)/gi, (match, inner) => {
    const args = splitTopLevelArgs(inner);
    const convertedArgs = args.map(arg => {
      if (arg.includes(':')) {
        return replaceCellRefs(arg, false);
      }
      return arg.trim();
    });
    return convertedArgs.join(' + ');
  });

  // Handle ROUND functions: ROUND(expr, digits) -> expr
  f = f.replace(/\bROUND(?:UP|DOWN)?\s*\(([^,]+),\s*\d+\)/gi, '$1');

  // Recursive converter for IF, MIN, MAX
  const parseNode = (expr: string): string => {
    expr = expr.trim();

    // Check for IF(...)
    const ifMatch = /^IF\s*\((.*)\)$/is.exec(expr);
    if (ifMatch) {
      const args = splitTopLevelArgs(ifMatch[1]);
      if (args.length >= 2) {
        const cond = parseNode(args[0]);
        const trueBranch = parseNode(args[1]);
        const falseBranch = args.length > 2 ? parseNode(args[2]) : '0';

        // Check if falseBranch is another CASE statement to flatten
        const nestedCaseMatch = /^CASE\s+BEGIN\s+(.*?)\s+END$/is.exec(falseBranch);
        if (nestedCaseMatch) {
          return `CASE BEGIN WHEN ${cond} THEN ${trueBranch} ${nestedCaseMatch[1]} END`;
        }
        return `CASE BEGIN WHEN ${cond} THEN ${trueBranch} ELSE ${falseBranch} END`;
      }
    }

    // Check for MIN(arg1, arg2) -> cap logic
    const minMatch = /^MIN\s*\((.*)\)$/is.exec(expr);
    if (minMatch) {
      const args = splitTopLevelArgs(minMatch[1]);
      if (args.length === 2) {
        const a = parseNode(args[0]);
        const b = parseNode(args[1]);
        // In payroll: MIN(1800, Basic * 0.12) -> cap at 1800
        return `CASE BEGIN WHEN ${b} > ${a} THEN ${a} ELSE ${b} END`;
      }
    }

    // Check for MAX(arg1, arg2) -> floor logic
    const maxMatch = /^MAX\s*\((.*)\)$/is.exec(expr);
    if (maxMatch) {
      const args = splitTopLevelArgs(maxMatch[1]);
      if (args.length === 2) {
        const a = parseNode(args[0]);
        const b = parseNode(args[1]);
        return `CASE BEGIN WHEN ${a} > ${b} THEN ${a} ELSE ${b} END`;
      }
    }

    // Convert AND(c1, c2) -> (c1 And c2)
    expr = expr.replace(/\bAND\s*\(([^)]+)\)/gi, (match, inner) => {
      const subArgs = splitTopLevelArgs(inner);
      return `(${subArgs.map(parseNode).join(' And ')})`;
    });

    // Convert OR(c1, c2) -> (c1 Or c2)
    expr = expr.replace(/\bOR\s*\(([^)]+)\)/gi, (match, inner) => {
      const subArgs = splitTopLevelArgs(inner);
      return `(${subArgs.map(parseNode).join(' Or ')})`;
    });

    // Convert NOT(c) -> (Not c)
    expr = expr.replace(/\bNOT\s*\(([^)]+)\)/gi, (match, inner) => {
      return `(Not ${parseNode(inner)})`;
    });

    return expr;
  };

  const caseConverted = parseNode(f);

  // Substitute cell coordinates with clean Pay Head names
  let readableFormula = replaceCellRefs(caseConverted, false);
  readableFormula = normalizeFormulaSyntax(readableFormula);

  // Generate safe JS expression
  let jsExpr = replaceCellRefs(caseConverted, true);
  
  // Convert CASE BEGIN WHEN ... to ternary for JS
  const caseRegex = /CASE(?:\s+BEGIN)?\s+([\s\S]*?)\s+END/gi;
  jsExpr = jsExpr.replace(caseRegex, (match, inner) => {
    const whenRegex = /WHEN\s+([\s\S]*?)\s+THEN\s+([\s\S]*?)(?=(?:WHEN|\bELSE\b|$))/gi;
    const clauses: { cond: string; val: string }[] = [];
    let elseVal = '0';

    let m;
    while ((m = whenRegex.exec(inner)) !== null) {
      clauses.push({ cond: m[1].trim(), val: m[2].trim() });
    }

    const elseMatch = /ELSE\s+([\s\S]*?)$/i.exec(inner);
    if (elseMatch) elseVal = elseMatch[1].trim();

    if (clauses.length === 0) return inner;

    let res = elseVal;
    for (let i = clauses.length - 1; i >= 0; i--) {
      res = `((${clauses[i].cond}) ? (${clauses[i].val}) : (${res}))`;
    }
    return res;
  });

  // Replace logical operators for JS
  jsExpr = jsExpr
    .replace(/\bAnd\b/gi, ' && ')
    .replace(/\bOr\b/gi, ' || ')
    .replace(/\bNot\b/gi, ' ! ')
    .replace(/<>/g, ' !== ')
    .replace(/(?<![<>=!])=(?![=<>])/g, ' === ');

  // Ensure all pay head names in jsExpr are converted to their jsVariableName
  const sortedHeads = [...allHeads].sort((a, b) => b.payHead.length - a.payHead.length);
  for (const h of sortedHeads) {
    if (!h.payHead) continue;
    const escaped = h.payHead.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const headRegex = new RegExp(`(?<![a-zA-Z0-9_])${escaped}(?![a-zA-Z0-9_])`, 'gi');
    jsExpr = jsExpr.replace(headRegex, h.isBase ? 'Total_CTC' : h.jsVariableName);
  }

  // Also replace standard aliases for Total_CTC
  jsExpr = jsExpr.replace(/\b(Total_CTC|Total CTC|CTC)\b/gi, 'Total_CTC');
  jsExpr = jsExpr.replace(/\s+/g, ' ').trim();

  // Improve operator spacing for readable formula
  readableFormula = readableFormula
    .replace(/(?<=[a-zA-Z0-9_])\s*([*+/])\s*(?=[a-zA-Z0-9_])/g, ' $1 ')
    .replace(/(?<=[a-zA-Z0-9_])\s*-\s*(?=[a-zA-Z0-9_])/g, ' - ')
    .replace(/(?<=[a-zA-Z0-9_])\s*(<=|>=|<>|<|>|=)\s*(?=[a-zA-Z0-9_])/g, ' $1 ')
    .replace(/\s+/g, ' ')
    .trim();

  // Create human-friendly description
  let description = `Calculated using formula: ${readableFormula}`;
  if (readableFormula.includes('CASE BEGIN WHEN')) {
    description = `Conditional calculation with tiered rates/caps based on dependent salary components.`;
  } else if (readableFormula.includes('+')) {
    description = `Sum of component allowances and earnings.`;
  } else if (readableFormula.includes('*')) {
    description = `Calculated as a proportional percentage of base earnings.`;
  }

  return {
    readableFormula,
    jsExpression: jsExpr,
    description
  };
}

/**
 * Topologically sorts pay heads based on their formula dependencies
 * so that sequenceNumber strictly starts at 5 and increments by 5.
 */
export function calculateTopologicalSequences(
  items: Omit<PayHeadFormula, 'sequenceNumber'>[]
): PayHeadFormula[] {
  const itemMap = new Map<string, Omit<PayHeadFormula, 'sequenceNumber'>>();
  const depMap = new Map<string, string[]>();

  for (const item of items) {
    itemMap.set(item.jsVariableName, item);
    const deps: string[] = [];

    if (!item.isFixedValue && !item.isBaseValue && item.formula) {
      for (const other of items) {
        if (other.jsVariableName !== item.jsVariableName) {
          // Check if other's payHead or jsVariableName is referenced
          const escapedHead = other.payHead.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
          const escapedVar = other.jsVariableName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
          const regex = new RegExp(`\\b(?:${escapedHead}|${escapedVar})\\b`, 'i');
          if (regex.test(item.formula)) {
            deps.push(other.jsVariableName);
          }
        }
      }
    }
    depMap.set(item.jsVariableName, deps);
  }

  // Kahn's algorithm or level assignment
  const inDegree = new Map<string, number>();
  const reverseGraph = new Map<string, string[]>(); // who depends on me

  for (const item of items) {
    inDegree.set(item.jsVariableName, (depMap.get(item.jsVariableName) || []).length);
    reverseGraph.set(item.jsVariableName, []);
  }

  for (const [node, deps] of depMap.entries()) {
    for (const dep of deps) {
      const list = reverseGraph.get(dep) || [];
      list.push(node);
      reverseGraph.set(dep, list);
    }
  }

  // Base items and 0-dependency items go first
  const queue: string[] = [];
  // Prioritize base value item first
  const baseItem = items.find(i => i.isBaseValue);
  if (baseItem && inDegree.get(baseItem.jsVariableName) === 0) {
    queue.push(baseItem.jsVariableName);
  }

  for (const item of items) {
    if (item.jsVariableName !== baseItem?.jsVariableName && (inDegree.get(item.jsVariableName) === 0)) {
      queue.push(item.jsVariableName);
    }
  }

  const sortedKeys: string[] = [];
  while (queue.length > 0) {
    const current = queue.shift()!;
    sortedKeys.push(current);

    const dependents = reverseGraph.get(current) || [];
    for (const dep of dependents) {
      const curDegree = (inDegree.get(dep) || 1) - 1;
      inDegree.set(dep, curDegree);
      if (curDegree === 0) {
        queue.push(dep);
      }
    }
  }

  // Add any remaining items (e.g. in case of cycles or unreached nodes)
  for (const item of items) {
    if (!sortedKeys.includes(item.jsVariableName)) {
      sortedKeys.push(item.jsVariableName);
    }
  }

  // Assign sequence numbers: strictly 5, 10, 15, 20...
  const result: PayHeadFormula[] = [];
  sortedKeys.forEach((key, index) => {
    const item = itemMap.get(key);
    if (item) {
      result.push({
        ...item,
        sequenceNumber: (index + 1) * 5
      });
    }
  });

  // Sort by rowNumber for natural sheet order in UI table
  result.sort((a, b) => a.rowNumber - b.rowNumber);

  return result;
}

/**
 * Parses raw text lines from Excel sheet into structured cell records
 */
function parseCellsFromText(sheetData: string): ParsedCell[] {
  const cells: ParsedCell[] = [];
  const lines = sheetData.split('\n');

  // Matches "A2: Basic Salary (Formula: =B1*0.5)" or "B2: 25000" or "C5: HRA"
  const lineRegex = /^\s*([A-Z]+)(\d+):\s*(.*?)(?:\s*\(Formula:\s*=(.*?)\))?\s*$/i;

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('---')) continue;

    const match = lineRegex.exec(trimmed);
    if (match) {
      const col = match[1].toUpperCase();
      const row = parseInt(match[2], 10);
      const valStr = match[3].trim();
      const formula = match[4]?.trim();

      let val: any = valStr;
      if (!isNaN(Number(valStr)) && valStr !== '') {
        val = Number(valStr);
      }

      cells.push({
        col,
        colIndex: col.charCodeAt(0) - 65,
        row,
        address: `${col}${row}`,
        value: val,
        formula: formula ? `=${formula}` : undefined
      });
    }
  }

  return cells;
}

/**
 * 100% Client-Side Local Excel CTC Formula Parser
 * Completely offline, zero network requests, zero API key requirements.
 */
export async function parseCTCFormulasLocally(
  sheetInput: string | { data: string; worksheet?: any; workbook?: any }
): Promise<PayHeadFormula[]> {
  let cells: ParsedCell[] = [];

  // 1. Extract cells either from raw worksheet or text lines
  if (typeof sheetInput === 'object' && sheetInput.worksheet) {
    const ws = sheetInput.worksheet;
    if (ws['!ref']) {
      const range = XLSX.utils.decode_range(ws['!ref']);
      for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const address = XLSX.utils.encode_cell({ r: R, c: C });
          const cell = ws[address];
          if (cell && (cell.v !== undefined || cell.f !== undefined)) {
            const col = XLSX.utils.encode_col(C);
            cells.push({
              col,
              colIndex: C,
              row: R + 1, // 1-indexed Excel row
              address,
              value: cell.v,
              formula: cell.f ? `=${cell.f}` : undefined
            });
          }
        }
      }
    }
  }

  // Fallback to text parsing if no cells loaded from worksheet object
  if (cells.length === 0) {
    const rawText = typeof sheetInput === 'string' ? sheetInput : (sheetInput.data || '');
    cells = parseCellsFromText(rawText);
  }

  if (cells.length === 0) {
    throw new Error(
      "The uploaded sheet contains no readable cells or data. Please ensure your Excel file contains salary components and formulas."
    );
  }

  // Group cells by row
  const rowsMap = new Map<number, ParsedCell[]>();
  for (const cell of cells) {
    const rowList = rowsMap.get(cell.row) || [];
    rowList.push(cell);
    rowsMap.set(cell.row, rowList);
  }

  // Identify Pay Head label column and value column
  // In CTC sheets, Pay Heads are descriptive text strings (e.g. Basic Salary, HRA, Provident Fund, CTC)
  const candidateRows: ExtractedPayHeadRow[] = [];

  const rowNumbers = Array.from(rowsMap.keys()).sort((a, b) => a - b);

  for (const rowNum of rowNumbers) {
    const rowCells = rowsMap.get(rowNum) || [];
    if (rowCells.length === 0) continue;

    // Find candidate label cell: non-empty string, length > 1, not pure number, not a generic table header
    let labelCell: ParsedCell | undefined;
    let valueCell: ParsedCell | undefined;

    // Prioritize cell that has text label
    for (const c of rowCells) {
      if (typeof c.value === 'string' && c.value.trim().length > 1) {
        const str = c.value.trim();
        // Ignore headers like "Sl No", "Particulars", "Amount"
        if (isCommonHeaderLabel(str)) {
          continue;
        }
        // If it looks like a component name
        if (!labelCell && isNaN(Number(str))) {
          labelCell = c;
        }
      }
    }

    // If label found, find corresponding value or formula cell
    if (labelCell) {
      // Find cell in same row with formula or number
      // Prioritize cells that have a formula
      valueCell = rowCells.find(c => c.address !== labelCell!.address && c.formula);
      if (!valueCell) {
        // Then prioritize numeric values
        valueCell = rowCells.find(c => c.address !== labelCell!.address && typeof c.value === 'number');
      }
      if (!valueCell) {
        // Any cell with non-empty value
        valueCell = rowCells.find(c => c.address !== labelCell!.address && c.value !== undefined && c.value !== '');
      }

      const rawHeadName = String(labelCell.value).trim();
      candidateRows.push({
        rowNumber: rowNum,
        payHead: rawHeadName,
        value: valueCell ? valueCell.value : 0,
        formula: valueCell ? valueCell.formula : undefined,
        colIndex: labelCell.colIndex,
        valueColIndex: valueCell ? valueCell.colIndex : labelCell.colIndex + 1,
        cellAddress: valueCell ? valueCell.address : labelCell.address
      });
    }
  }

  if (candidateRows.length === 0) {
    throw new Error(
      "No salary components could be identified in the uploaded sheet.\n\n" +
      "Checklist for your Excel file:\n" +
      "1. Are there labeled rows for components (e.g., 'Basic Salary', 'HRA', 'Total CTC')?\n" +
      "2. Are calculations entered as Excel formulas (starting with '=')?\n" +
      "3. Is the sheet clean without empty top banners?"
    );
  }

  // Create Cell -> Pay Head and Row -> Pay Head maps
  const cellToHead = new Map<string, string>();
  const rowToHead = new Map<number, string>();

  for (const r of candidateRows) {
    rowToHead.set(r.rowNumber, r.payHead);
    // Map all cells in this row to this pay head
    const rowCells = rowsMap.get(r.rowNumber) || [];
    for (const c of rowCells) {
      cellToHead.set(c.address.toUpperCase(), r.payHead);
    }
  }

  // Identify Base Value (Root component like Total CTC or Gross)
  // Look for "Total CTC", "CTC", "Gross CTC", or the component that other components derive from
  let baseRowIndex = candidateRows.findIndex(r => 
    /\b(total ctc|ctc|gross ctc|cost to company|annual ctc)\b/i.test(r.payHead)
  );

  if (baseRowIndex === -1) {
    // Look for Gross Salary
    baseRowIndex = candidateRows.findIndex(r => /\b(gross salary|gross|total earnings)\b/i.test(r.payHead));
  }

  // If still not found, pick the first row with no formula (fixed base input)
  if (baseRowIndex === -1) {
    baseRowIndex = candidateRows.findIndex(r => !r.formula);
  }

  if (baseRowIndex === -1) {
    baseRowIndex = 0;
  }

  const allHeadsInfo = candidateRows.map((r, idx) => ({
    payHead: r.payHead,
    jsVariableName: toJsVariableName(r.payHead),
    isBase: idx === baseRowIndex
  }));

  // Build the preliminary list of PayHeadFormulas
  const preliminaryFormulas: Omit<PayHeadFormula, 'sequenceNumber'>[] = [];

  for (let i = 0; i < candidateRows.length; i++) {
    const row = candidateRows[i];
    const isBase = (i === baseRowIndex);
    const hasFormula = Boolean(row.formula && row.formula.trim().length > 1);

    const jsVarName = toJsVariableName(row.payHead);

    let formulaStr = '';
    let jsFormulaStr = '';
    let description = '';
    const defaultValue = typeof row.value === 'number' ? row.value : 0;
    const isFixed = !hasFormula;

    if (hasFormula && row.formula) {
      const translated = translateExcelFormula(row.formula, cellToHead, rowToHead, allHeadsInfo);
      formulaStr = translated.readableFormula;
      jsFormulaStr = translated.jsExpression;
      description = translated.description;
    } else {
      // Manual / Value-based head
      formulaStr = '';
      jsFormulaStr = `${defaultValue}`;
      description = isBase 
        ? 'Base CTC reference value entered by user upon which all dependent salary components are calculated.'
        : 'Value-based component with fixed/manual input amount.';
    }

    preliminaryFormulas.push({
      payHead: row.payHead,
      formula: formulaStr,
      description,
      jsVariableName: jsVarName,
      jsFormula: jsFormulaStr,
      rowNumber: row.rowNumber,
      isFixedValue: isFixed,
      defaultValue,
      isBaseValue: isBase
    });
  }

  // Calculate dependency sequence numbers (5, 10, 15...) using topological order
  const finalized = calculateTopologicalSequences(preliminaryFormulas);

  return finalized;
}
