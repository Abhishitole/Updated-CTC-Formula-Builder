import { parseCTCFormulasLocally } from '../lib/local-formula-parser';

export interface PayHeadFormula {
  sequenceNumber: number;
  payHead: string;
  formula: string;
  description: string;
  jsVariableName: string;
  jsFormula: string;
  rowNumber: number;
  isFixedValue: boolean;
  defaultValue: number;
  isBaseValue: boolean;
}

export function normalizeFormulaSyntax(formula: string): string {
  if (!formula || typeof formula !== 'string') return "";
  
  let formatted = formula.trim();

  // Replace programming operators with payroll keyword operators
  formatted = formatted
    .replace(/&&/g, ' And ')
    .replace(/\|\|/g, ' Or ')
    .replace(/!=/g, ' <> ')
    .replace(/!/g, ' Not ');

  // Standardize core payroll conditional keywords to uppercase: CASE, BEGIN, WHEN, THEN, ELSE, END
  formatted = formatted.replace(/\b(case|begin|when|then|else|end)\b/gi, (match) => match.toUpperCase());

  // Ensure "CASE BEGIN WHEN": If "CASE WHEN" is present without "BEGIN" in between
  formatted = formatted.replace(/\bCASE\s+(?!BEGIN\b)WHEN\b/g, 'CASE BEGIN WHEN');

  // Convert basic IF(condition, trueVal, falseVal) if present
  const ifMatch = formatted.match(/^IF\s*\((.*)\)$/i);
  if (ifMatch) {
    const inner = ifMatch[1];
    const parts: string[] = [];
    let current = '';
    let parenDepth = 0;
    for (let i = 0; i < inner.length; i++) {
      const char = inner[i];
      if (char === '(') parenDepth++;
      else if (char === ')') parenDepth--;
      else if (char === ',' && parenDepth === 0) {
        parts.push(current.trim());
        current = '';
        continue;
      }
      current += char;
    }
    if (current.trim()) parts.push(current.trim());
    if (parts.length === 3) {
      formatted = `CASE BEGIN WHEN ${parts[0]} THEN ${parts[1]} ELSE ${parts[2]} END`;
    }
  }

  // Clean up any double spaces
  formatted = formatted.replace(/\s+/g, ' ').trim();

  return formatted;
}

/**
 * 100% Client-Side Local Excel CTC Formula Parser
 * ZERO external API or AI integration. Runs 100% locally in browser.
 * No API keys, no RPC failures, no network delays.
 */
export async function analyzeCTCFormulas(
  sheetInput: string | { data: string; worksheet?: any; workbook?: any }
): Promise<PayHeadFormula[]> {
  try {
    return await parseCTCFormulasLocally(sheetInput);
  } catch (err: any) {
    console.error("[LocalFormulaParser] Error:", err);
    throw err;
  }
}

/**
 * Kept for backward compatibility if referenced elsewhere.
 * Always returns empty string since no API key is required.
 */
export const getApiKey = (): string => "";
