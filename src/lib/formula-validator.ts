import { PayHeadFormula } from '../services/gemini';

export interface FormulaIssue {
  severity: 'error' | 'warning';
  type: 'circular_dependency' | 'illegal_function' | 'syntax_error' | 'unknown_reference' | 'evaluation_error';
  message: string;
  details?: string;
  suggestedFix?: string;
}

export interface ComponentValidationResult {
  isValid: boolean;
  hasErrors: boolean;
  hasWarnings: boolean;
  issues: FormulaIssue[];
  dependencies: string[]; // list of jsVariableName this component depends on
  transpiledJsFormula: string;
}

export interface ValidationSummary {
  totalFormulas: number;
  validCount: number;
  errorCount: number;
  warningCount: number;
  allValid: boolean;
  cycles: string[][];
  resultsByVarName: Record<string, ComponentValidationResult>;
}

// Blocklist of forbidden browser APIs, global objects, and dangerous JavaScript keywords
const FORBIDDEN_KEYWORDS = [
  'window', 'document', 'globalthis', 'self', 'top', 'parent', 'frames',
  'eval', 'function', 'settimeout', 'setinterval', 'setimmediate',
  'requestanimationframe', 'fetch', 'xmlhttprequest', 'websocket',
  'worker', 'localstorage', 'sessionstorage', 'indexeddb', 'cookie',
  'location', 'navigator', 'history', 'alert', 'confirm', 'prompt',
  'process', 'require', 'import', 'constructor', 'prototype', '__proto__',
  'object', 'array', 'reflect', 'proxy', 'symbol', 'promise',
  'for', 'while', 'do', 'switch', 'return', 'throw', 'try', 'catch',
  'class', 'delete', 'typeof', 'void', 'debugger', 'await', 'async'
];

/**
 * Normalizes an identifier or pay head name into a clean regex-safe token
 */
function escapeRegex(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Converts a human-readable or CASE-based payroll formula into a safe JavaScript expression.
 */
export function convertFormulaToJs(
  formulaStr: string,
  allItems: PayHeadFormula[],
  currentVarName?: string
): { jsExpression: string; conversionErrors: string[] } {
  const conversionErrors: string[] = [];
  if (!formulaStr || !formulaStr.trim()) {
    return { jsExpression: '0', conversionErrors };
  }

  let expr = formulaStr.trim();

  // Strip leading '=' if present from Excel
  if (expr.startsWith('=')) {
    expr = expr.substring(1).trim();
  }

  // Handle CASE BEGIN WHEN ... THEN ... ELSE ... END structures
  // Pattern supports multiple WHEN ... THEN clauses followed by an optional ELSE and END
  const casePattern = /CASE(?:\s+BEGIN)?\s+([\s\S]*?)\s+END/gi;
  let hasCase = false;

  expr = expr.replace(casePattern, (match, inner) => {
    hasCase = true;
    return convertCaseBlockToTernary(inner);
  });

  // Handle orphaned CASE without END
  if (/CASE\b/i.test(expr) && !hasCase) {
    conversionErrors.push("Formula contains 'CASE' without matching 'END' statement.");
  }

  // Replace logical keywords: AND, OR, NOT
  expr = expr.replace(/\bAND\b/gi, ' && ');
  expr = expr.replace(/\bOR\b/gi, ' || ');
  expr = expr.replace(/\bNOT\b/gi, ' ! ');

  // Replace comparison operators: <> to !==
  expr = expr.replace(/<>/g, ' !== ');

  // Replace single '=' with '===' in comparison contexts, but ignore <=, >=, !==, ===
  expr = expr.replace(/(?<![<>=!])=(?!=)/g, ' === ');

  // Map pay head names to jsVariableName (sorted by length descending to match longest matches first)
  const sortedItems = [...allItems].sort((a, b) => {
    const lenA = Math.max(a.payHead.length, a.jsVariableName.length);
    const lenB = Math.max(b.payHead.length, b.jsVariableName.length);
    return lenB - lenA;
  });

  // Standard aliases for base CTC
  const ctcAliases = ['Total CTC', 'Total_CTC', 'CTC', 'TotalCTC', 'Gross CTC'];

  for (const alias of ctcAliases) {
    const regex = new RegExp(`\\b${escapeRegex(alias)}\\b`, 'gi');
    expr = expr.replace(regex, 'Total_CTC');
  }

  for (const item of sortedItems) {
    if (!item.payHead) continue;

    // Replace full payHead name (e.g. "Basic Salary" -> "Basic_Salary")
    const payHeadRegex = new RegExp(`(?<![a-zA-Z0-9_])${escapeRegex(item.payHead)}(?![a-zA-Z0-9_])`, 'gi');
    expr = expr.replace(payHeadRegex, item.jsVariableName);
  }

  // Clean up any double spaces
  expr = expr.replace(/\s+/g, ' ').trim();

  return { jsExpression: expr, conversionErrors };
}

/**
 * Converts inner content of CASE ... END into nested JS ternary expressions
 */
function convertCaseBlockToTernary(caseInner: string): string {
  // Extract WHEN ... THEN clauses
  // Pattern: WHEN <cond> THEN <val>
  const whenThenRegex = /WHEN\s+([\s\S]*?)\s+THEN\s+([\s\S]*?)(?=(?:WHEN|\bELSE\b|$))/gi;
  const clauses: { condition: string; value: string }[] = [];
  let elseValue = '0';

  let match;
  while ((match = whenThenRegex.exec(caseInner)) !== null) {
    clauses.push({
      condition: match[1].trim(),
      value: match[2].trim()
    });
  }

  const elseMatch = /ELSE\s+([\s\S]*?)$/i.exec(caseInner);
  if (elseMatch) {
    elseValue = elseMatch[1].trim();
  }

  if (clauses.length === 0) {
    return caseInner; // Couldn't parse cleanly, return as-is
  }

  // Build nested ternary from right to left: cond1 ? val1 : (cond2 ? val2 : elseValue)
  let result = elseValue;
  for (let i = clauses.length - 1; i >= 0; i--) {
    const c = clauses[i];
    result = `((${c.condition}) ? (${c.value}) : (${result}))`;
  }

  return result;
}

/**
 * Extracts all component dependencies (jsVariableName) referenced in a formula
 */
export function extractFormulaDependencies(
  formula: string,
  allItems: PayHeadFormula[],
  currentVarName: string
): string[] {
  if (!formula || typeof formula !== 'string') return [];

  const foundDependencies = new Set<string>();
  const normalizedFormula = formula.toLowerCase();

  for (const item of allItems) {
    const varNameLower = item.jsVariableName.toLowerCase();
    const payHeadLower = item.payHead.toLowerCase();

    // Check if payHead or jsVariableName is referenced as a distinct word
    const varRegex = new RegExp(`(?<![a-z0-9_])${escapeRegex(varNameLower)}(?![a-z0-9_])`, 'i');
    const headRegex = new RegExp(`(?<![a-z0-9_])${escapeRegex(payHeadLower)}(?![a-z0-9_])`, 'i');

    if (varRegex.test(normalizedFormula) || headRegex.test(normalizedFormula)) {
      foundDependencies.add(item.jsVariableName);
    }
  }

  // Also check for CTC / Total_CTC references
  if (/\b(total_ctc|total ctc|ctc)\b/i.test(normalizedFormula)) {
    const baseItem = allItems.find(i => i.isBaseValue);
    if (baseItem) {
      foundDependencies.add(baseItem.jsVariableName);
    }
  }

  return Array.from(foundDependencies);
}

/**
 * Detects circular dependencies across all formulas using depth-first search (DFS)
 */
export function detectCircularDependencies(
  items: PayHeadFormula[]
): { cycles: string[][]; cycledVarNames: Set<string> } {
  const dependencyMap = new Map<string, string[]>();
  const itemMap = new Map<string, PayHeadFormula>();

  for (const item of items) {
    itemMap.set(item.jsVariableName, item);
    // Fixed values and base root values don't depend on other components
    if (item.isFixedValue || item.isBaseValue) {
      dependencyMap.set(item.jsVariableName, []);
    } else {
      const deps = extractFormulaDependencies(item.formula || item.jsFormula || '', items, item.jsVariableName);
      dependencyMap.set(item.jsVariableName, deps);
    }
  }

  const cycles: string[][] = [];
  const cycledVarNames = new Set<string>();

  // Tarjan / DFS cycle detection
  const visited = new Set<string>();
  const recStack: string[] = [];
  const recSet = new Set<string>();

  function dfs(node: string) {
    visited.add(node);
    recStack.push(node);
    recSet.add(node);

    const neighbors = dependencyMap.get(node) || [];
    for (const neighbor of neighbors) {
      if (!recSet.has(neighbor)) {
        if (!visited.has(neighbor) && itemMap.has(neighbor)) {
          dfs(neighbor);
        }
      } else {
        // Cycle detected! Extract the cycle path
        const cycleStartIndex = recStack.indexOf(neighbor);
        if (cycleStartIndex !== -1) {
          const cyclePath = [...recStack.slice(cycleStartIndex), neighbor];
          cycles.push(cyclePath);
          for (const varName of cyclePath) {
            cycledVarNames.add(varName);
          }
        }
      }
    }

    recStack.pop();
    recSet.delete(node);
  }

  for (const item of items) {
    if (!visited.has(item.jsVariableName)) {
      dfs(item.jsVariableName);
    }
  }

  return { cycles, cycledVarNames };
}

/**
 * Checks formula for illegal browser functions, security violations, and syntax errors
 */
export function validateFormulaSafetyAndSyntax(
  item: PayHeadFormula,
  allItems: PayHeadFormula[],
  cycleInfo?: { cycles: string[][]; cycledVarNames: Set<string> }
): ComponentValidationResult {
  const issues: FormulaIssue[] = [];
  const formula = (item.formula || '').trim();

  // If fixed value or base value, formula check is relaxed
  if (item.isFixedValue || item.isBaseValue) {
    return {
      isValid: true,
      hasErrors: false,
      hasWarnings: false,
      issues: [],
      dependencies: [],
      transpiledJsFormula: item.jsFormula || `${item.defaultValue || 0}`
    };
  }

  // 1. Check for empty formula
  if (!formula) {
    issues.push({
      severity: 'error',
      type: 'syntax_error',
      message: 'Formula is empty',
      details: 'Formula-calculated components must define a valid calculation formula.',
      suggestedFix: 'Enter an arithmetic expression or mark as Manual / Value-Based Head.'
    });
    return {
      isValid: false,
      hasErrors: true,
      hasWarnings: false,
      issues,
      dependencies: [],
      transpiledJsFormula: '0'
    };
  }

  // 2. Check for Circular Dependency
  const dependencies = extractFormulaDependencies(formula, allItems, item.jsVariableName);
  
  // Direct self-reference check
  if (dependencies.includes(item.jsVariableName)) {
    issues.push({
      severity: 'error',
      type: 'circular_dependency',
      message: `Self-referential formula loop detected: ${item.payHead} references itself.`,
      details: `A salary component cannot use its own calculated value in its formula (${item.payHead} → ${item.payHead}).`,
      suggestedFix: 'Replace the self-reference with another component like Basic Salary or Total CTC.'
    });
  } else if (cycleInfo && cycleInfo.cycledVarNames.has(item.jsVariableName)) {
    // Check if this item is part of a multi-node cycle
    const relevantCycle = cycleInfo.cycles.find(c => c.includes(item.jsVariableName));
    if (relevantCycle) {
      const cycleFormatted = relevantCycle.map(varName => {
        const found = allItems.find(i => i.jsVariableName === varName);
        return found ? found.payHead : varName;
      }).join(' → ');

      issues.push({
        severity: 'error',
        type: 'circular_dependency',
        message: `Circular dependency loop detected: ${cycleFormatted}`,
        details: `This component is locked in a recursive calculation loop. The browser cannot compute this value without infinite recursion.`,
        suggestedFix: 'Edit this formula or one of the dependent components to break the loop.'
      });
    }
  }

  // 3. Security & Forbidden Browser API / Function Check
  // Tokenize and scan for forbidden words
  const formulaLower = formula.toLowerCase();
  
  for (const forbidden of FORBIDDEN_KEYWORDS) {
    const regex = new RegExp(`\\b${escapeRegex(forbidden)}\\b`, 'i');
    if (regex.test(formulaLower)) {
      issues.push({
        severity: 'error',
        type: 'illegal_function',
        message: `Illegal identifier or forbidden keyword '${forbidden}' detected.`,
        details: `Browser execution policy strictly forbids access to '${forbidden}' to prevent unauthorized script execution. Only mathematical expressions and payroll pay heads are permitted.`,
        suggestedFix: `Remove '${forbidden}' and use standard operators or Math.round / Math.max / Math.min.`
      });
    }
  }

  // Check for dangerous operators like semicolons, curly braces, backticks, or script tags
  if (/[;{}`]/.test(formula)) {
    issues.push({
      severity: 'error',
      type: 'illegal_function',
      message: 'Illegal characters (;, {, }, or `) detected in formula.',
      details: 'Formulas must be single continuous mathematical or conditional expressions without multiple statements or code blocks.',
      suggestedFix: 'Remove semicolons, curly brackets, and backticks.'
    });
  }

  // Check for assignment operator (e.g. x = 5 instead of comparison)
  if (/(?<![<>=!])=(?![=<>])/.test(formula)) {
    // If not in a WHEN clause, single '=' could be ambiguous, but in payroll CASE BEGIN WHEN, single '=' is treated as comparison
    if (!/CASE|WHEN/i.test(formula)) {
      issues.push({
        severity: 'warning',
        type: 'syntax_error',
        message: "Single '=' used in mathematical formula.",
        details: "Single '=' is typically an assignment operator. For comparisons, write '===' or use CASE BEGIN WHEN ... THEN ... END.",
        suggestedFix: "Use '===' or '<=' or '>=' for comparisons."
      });
    }
  }

  // 4. Balanced Parentheses Check
  let parenDepth = 0;
  let hasParenError = false;
  for (let i = 0; i < formula.length; i++) {
    if (formula[i] === '(') parenDepth++;
    if (formula[i] === ')') parenDepth--;
    if (parenDepth < 0) {
      hasParenError = true;
      break;
    }
  }
  if (hasParenError || parenDepth !== 0) {
    issues.push({
      severity: 'error',
      type: 'syntax_error',
      message: 'Unbalanced parentheses detected in formula.',
      details: `Count of opening and closing parentheses does not match (${parenDepth > 0 ? `${parenDepth} missing ')'` : 'unexpected extra \')\''}).`,
      suggestedFix: 'Ensure every opening bracket \'(\' has a corresponding closing bracket \')\'.'
    });
  }

  // 5. CASE ... END Structure Check
  const caseMatches = formula.match(/\bCASE\b/gi) || [];
  const endMatches = formula.match(/\bEND\b/gi) || [];
  if (caseMatches.length !== endMatches.length) {
    issues.push({
      severity: 'error',
      type: 'syntax_error',
      message: `Mismatched CASE / END blocks (${caseMatches.length} CASE, ${endMatches.length} END).`,
      details: "Every conditional block starting with 'CASE' must terminate with a matching 'END' keyword.",
      suggestedFix: "Add 'END' to close each 'CASE BEGIN WHEN ...' block."
    });
  }

  if (caseMatches.length > 0) {
    const whenMatches = formula.match(/\bWHEN\b/gi) || [];
    const thenMatches = formula.match(/\bTHEN\b/gi) || [];
    if (whenMatches.length !== thenMatches.length) {
      issues.push({
        severity: 'error',
        type: 'syntax_error',
        message: `Mismatched WHEN / THEN statements (${whenMatches.length} WHEN, ${thenMatches.length} THEN).`,
        details: 'Every WHEN condition must be followed by a THEN expression.',
        suggestedFix: 'Ensure every WHEN condition has an associated THEN value.'
      });
    }
  }

  // 6. Transpile to JS and test run dry-compilation in browser
  const { jsExpression, conversionErrors } = convertFormulaToJs(formula, allItems, item.jsVariableName);
  for (const err of conversionErrors) {
    issues.push({
      severity: 'error',
      type: 'syntax_error',
      message: err,
      details: 'The formula could not be translated into a safe executable expression.'
    });
  }

  // 7. Dry-run Browser Parser Test (Sandbox check)
  if (!issues.some(i => i.severity === 'error')) {
    try {
      // Mock arguments representing all available pay variables with positive test numbers
      const mockArgNames = Array.from(new Set([
        'Total_CTC',
        'CTC',
        'TotalCTC',
        ...allItems.map(i => i.jsVariableName)
      ]));
      const mockArgValues = mockArgNames.map(() => 50000);

      // Create a test function to verify browser JavaScript parser accepts it
      const testFunc = new Function(...mockArgNames, `"use strict"; return (${jsExpression});`);
      const evaluated = testFunc(...mockArgValues);

      if (typeof evaluated !== 'number' && typeof evaluated !== 'boolean') {
        issues.push({
          severity: 'warning',
          type: 'evaluation_error',
          message: `Evaluation resulted in non-numeric value (${typeof evaluated}).`,
          details: 'Payroll calculations must evaluate to a number.',
          suggestedFix: 'Ensure all branches in the formula produce a numeric amount.'
        });
      } else if (Number.isNaN(evaluated)) {
        issues.push({
          severity: 'warning',
          type: 'evaluation_error',
          message: 'Formula evaluates to NaN (Not a Number) with sample values.',
          details: 'Check for division by zero or invalid arithmetic combinations.',
          suggestedFix: 'Guard divisions with CASE WHEN divisor > 0 THEN amount / divisor ELSE 0 END.'
        });
      }
    } catch (e: any) {
      issues.push({
        severity: 'error',
        type: 'syntax_error',
        message: `Browser parsing failed: ${e.message}`,
        details: `The browser could not parse this formula: ${e.name} - ${e.message}.`,
        suggestedFix: 'Review formula syntax, arithmetic operators, or spelling of component names.'
      });
    }
  }

  const hasErrors = issues.some(i => i.severity === 'error');
  const hasWarnings = issues.some(i => i.severity === 'warning');

  return {
    isValid: !hasErrors,
    hasErrors,
    hasWarnings,
    issues,
    dependencies,
    transpiledJsFormula: jsExpression
  };
}

/**
 * Validates all formulas in the dataset, identifying global circular dependencies
 * and individual component errors.
 */
export function validateAllFormulas(items: PayHeadFormula[]): ValidationSummary {
  const cycleInfo = detectCircularDependencies(items);
  const resultsByVarName: Record<string, ComponentValidationResult> = {};

  let validCount = 0;
  let errorCount = 0;
  let warningCount = 0;

  for (const item of items) {
    const res = validateFormulaSafetyAndSyntax(item, items, cycleInfo);
    resultsByVarName[item.jsVariableName] = res;

    if (res.hasErrors) {
      errorCount++;
    } else {
      validCount++;
    }

    if (res.hasWarnings) {
      warningCount++;
    }
  }

  return {
    totalFormulas: items.length,
    validCount,
    errorCount,
    warningCount,
    allValid: errorCount === 0,
    cycles: cycleInfo.cycles,
    resultsByVarName
  };
}
