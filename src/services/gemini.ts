import { GoogleGenAI, Type } from "@google/genai";

export const getApiKey = (): string => {
  // Check for Vite environment variables (used in Netlify/production)
  if (typeof import.meta !== 'undefined' && (import.meta as any).env) {
    if ((import.meta as any).env.VITE_GEMINI_API_KEY) {
      return (import.meta as any).env.VITE_GEMINI_API_KEY;
    }
    if ((import.meta as any).env.GEMINI_API_KEY) {
      return (import.meta as any).env.GEMINI_API_KEY;
    }
  }
  // Check for Node/defined process.env variable
  if (typeof process !== 'undefined' && process.env) {
    if (process.env.VITE_GEMINI_API_KEY) {
      return process.env.VITE_GEMINI_API_KEY;
    }
    if (process.env.GEMINI_API_KEY) {
      return process.env.GEMINI_API_KEY;
    }
  }
  return "";
};

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

  // Convert basic IF(condition, trueVal, falseVal) if AI accidentally outputs Excel IF
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

export async function analyzeCTCFormulas(sheetData: string): Promise<PayHeadFormula[]> {
  const apiKey = getApiKey();
  if (!apiKey) {
    throw new Error(
      "Gemini API key is missing. Please configure VITE_GEMINI_API_KEY in your Netlify Environment Variables (Site Configuration > Environment Variables), then re-trigger a deployment."
    );
  }

  const ai = new GoogleGenAI({ apiKey });

  if (!sheetData || sheetData.trim().length === 0) {
    throw new Error("The Excel sheet appears to be empty or could not be read properly.");
  }

    const prompt = `
    Analyze the following Excel sheet data which contains CTC (Cost to Company) calculations.
    The data includes cell addresses, values, and formulas.
    
    Your task:
    1. EXTRACT EVERY SINGLE ROW: You MUST identify and extract EVERY row from the provided Excel data as a Pay Head. Do not skip any rows. This includes standard components (Basic, HRA), fixed allowances, zero/blank values, and aggregate totals/sub-totals (Gross Salary, Total CTC, Net Pay).
    2. PRESERVE ACTUAL FORMULAS EXACTLY: For each identified item, determine the mathematical formula used to calculate it based on the Excel formulas provided. If the Excel sheet provides an actual formula (e.g., =B2*0.5), you MUST preserve that calculation logic.
    3. IMPORTANT: Convert Excel cell references (like A1, B2, C10) into the actual Pay Head names.
    4. STRICTLY NO INVENTING FORMULAS: You MUST NOT invent formulas or percentages. If a component has NO formula in the Excel sheet (e.g., it is empty, blank, or just a hardcoded number), you MUST treat it as a manual input/fixed value. Do NOT convert it into a percentage of CTC.
    5. Determine the calculation sequence number for each pay head based on their dependencies. The sequence MUST start from 5 and increase by 5. Components that depend on others must have a higher sequence number than their dependencies.
    6. STRICT FORMULA SYNTAX - MANDATORY CASE BEGIN WHEN ... THEN ... ELSE ... END:
       - Where conditional logic, capping (MIN/MAX), tiered brackets, or IF logic is required, you MUST use the exact syntax:
         CASE BEGIN WHEN <Condition> THEN <Expression> ELSE <DefaultExpression> END
       - For multiple conditions or nested IFs:
         CASE BEGIN WHEN <Condition1> THEN <Expression1> WHEN <Condition2> THEN <Expression2> ELSE <DefaultExpression> END
       - You MUST use ONLY these exact uppercase keywords for conditional logic: CASE, BEGIN, WHEN, THEN, ELSE, END.
       - NEVER use "IF", "IFS", "MIN", or "MAX" in the formula field! Always convert them into CASE BEGIN WHEN ... THEN ... ELSE ... END:
         * Excel '=IF(Basic > 20000, Basic * 0.4, Basic * 0.3)' -> 'CASE BEGIN WHEN Basic > 20000 THEN Basic * 0.4 ELSE Basic * 0.3 END'
         * Excel '=MIN(1800, Basic * 0.12)' (PF cap at 1800) -> 'CASE BEGIN WHEN Basic * 0.12 > 1800 THEN 1800 ELSE Basic * 0.12 END'
         * Excel '=MAX(0, Total_CTC - Gross)' (floor at 0) -> 'CASE BEGIN WHEN Total_CTC - Gross > 0 THEN Total_CTC - Gross ELSE 0 END'
         * Excel nested IF '=IF(Basic<=10000, 1000, IF(Basic<=20000, 2000, 3000))' -> 'CASE BEGIN WHEN Basic <= 10000 THEN 1000 WHEN Basic <= 20000 THEN 2000 ELSE 3000 END'
       - For pure arithmetic without conditions (e.g. Basic + HRA or Total_CTC * 0.5), write it directly with Pay Head names.
       - For calculations and logic, use ONLY these characters/words: + - * / And Or Not ( ) = , < > <>.
       - DO NOT use the '%' symbol. Convert percentages to decimals (e.g., 40% becomes 0.4) or division (40 / 100).
       - DO NOT use programming symbols for logic or math (e.g., no &, &&, ||, !, $, !=).
       - The output must be human-readable and use the exact Pay Head names in the formulas.
    7. JAVASCRIPT EVALUATION:
       - Provide a \`jsVariableName\` for each pay head (e.g., "Basic Salary" -> "Basic_Salary"). Must be a valid JS variable name (no spaces, no special chars).
       - Provide a \`jsFormula\` which is a valid JavaScript expression to calculate this value. If there is no formula, leave it empty or "0".
       - In \`jsFormula\`, use \`Total_CTC\` to refer to the total CTC input amount.
       - Use the exact \`jsVariableName\` of other components for dependencies.
       - Use ternary operators (\`condition ? true_val : false_val\`) instead of CASE WHEN.
       - Use standard JS operators (\`&&\`, \`||\`, \`===\`, \`+\`, \`-\`, \`*\`, \`/\`).
       - Use \`Math.round()\`, \`Math.max()\`, \`Math.min()\` if needed.
       - Provide ONLY the expression itself, without \`return\`, without backticks, without semicolons. Example: \`Total_CTC * 0.4\`
    8. ROW NUMBER:
       - Extract the \`rowNumber\` from the cell address where the pay head is defined (e.g., if defined in A10 or B10, rowNumber is 10).
    9. MANUAL INPUTS & VALUE-BASED HEADS (EDITABILITY):
       - If a pay head has NO formula in the Excel sheet (e.g., it is an empty cell, a blank value, or a hardcoded number like 12000), it is a "value-based head". You MUST set \`isFixedValue\` to \`true\`.
       - If it was empty/blank, set \`defaultValue\` to \`0\`. If it had a hardcoded number, set \`defaultValue\` to that number.
       - If it HAS an Excel formula, set \`isFixedValue\` to \`false\` and \`defaultValue\` to \`0\`.
    10. BASE VALUE IDENTIFICATION:
       - Identify the SINGLE pay head that acts as the root of all calculations (typically 'Total CTC', 'CTC', or 'Gross Salary'). This is the value that other formulas depend on.
       - Set \`isBaseValue\` to \`true\` for this single component, and \`false\` for all others.
    
    Examples:
    1. Conditional Formula:
       If cell B2 is "Basic Salary" and B3 is "HRA" with formula "=IF(B2>20000, B2*0.4, B2*0.3)", the output MUST be:
       Sequence Number: 10
       Pay Head: HRA
       Formula: CASE BEGIN WHEN Basic Salary > 20000 THEN Basic Salary * 0.4 ELSE Basic Salary * 0.3 END
       Description: HRA is calculated as 40% of Basic if Basic is above 20000, otherwise 30%.
    
    2. Capping / Limit Formula:
       If cell B4 is "Provident Fund" with formula "=MIN(1800, B2*0.12)", the output MUST be:
       Sequence Number: 15
       Pay Head: Provident Fund
       Formula: CASE BEGIN WHEN Basic Salary * 0.12 > 1800 THEN 1800 ELSE Basic Salary * 0.12 END
       Description: Provident Fund is 12% of Basic, capped at a maximum of 1800.
    
    Data from Excel (Cell Address: Value [Formula]):
    ${sheetData}
    
    Return the result as a JSON array of objects with keys: "sequenceNumber", "payHead", "formula", "description", "jsVariableName", "jsFormula", "rowNumber", "isFixedValue", "defaultValue", and "isBaseValue".
    Ensure EVERY SINGLE ROW from the Excel data is included in the list.
  `;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-3-flash-preview",
      contents: prompt,
      config: {
        systemInstruction: "You are a specialized Payroll Logic Analyst. Your expertise is in reverse-engineering compensation structures from Excel data. You must accurately map cell references to component names. You MUST strictly follow the formula syntax rules: where conditions, caps, or conditional statements are required, use ONLY 'CASE BEGIN WHEN ... THEN ... ELSE ... END' with allowed operators (+ - * / And Or Not ( ) = , < > <>). Never use IF, IFS, MIN, or MAX in the formula field.",
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.ARRAY,
          items: {
            type: Type.OBJECT,
            properties: {
              sequenceNumber: { type: Type.INTEGER, description: "The order in which this component should be calculated based on dependencies" },
              payHead: { type: Type.STRING, description: "The name of the salary component" },
              formula: { type: Type.STRING, description: "The formula using component names. For any condition, capping, or IF statement, you MUST use the exact syntax: CASE BEGIN WHEN condition THEN value ELSE default END. Do NOT use IF, MIN, or MAX." },
              description: { type: Type.STRING, description: "A brief explanation of how this component is derived" },
              jsVariableName: { type: Type.STRING, description: "A valid JavaScript variable name for this component" },
              jsFormula: { type: Type.STRING, description: "A valid JavaScript expression for this formula" },
              rowNumber: { type: Type.INTEGER, description: "The Excel row number where this component is defined (e.g., 2 for cell B2)" },
              isFixedValue: { type: Type.BOOLEAN, description: "True if this is a hardcoded fixed amount that doesn't scale with CTC" },
              defaultValue: { type: Type.NUMBER, description: "The default fixed amount if isFixedValue is true, otherwise 0" },
              isBaseValue: { type: Type.BOOLEAN, description: "True ONLY for the single root component that everything else depends on (e.g. Total CTC)" },
            },
            required: ["sequenceNumber", "payHead", "formula", "description", "jsVariableName", "jsFormula", "rowNumber", "isFixedValue", "defaultValue", "isBaseValue"],
          },
        },
      },
    });

    const text = response.text;
    if (!text || text.trim() === "") {
      throw new Error(
        "The AI model was unable to generate a response. This often happens if the Excel data is too sparse or lacks clear structure. \n\n" +
        "Tip: Ensure your sheet has clear column/row headers (e.g., 'Basic', 'HRA') and that calculations are performed using Excel formulas (e.g., '=A2*0.5') rather than just plain numbers."
      );
    }

    let parsed;
    try {
      parsed = JSON.parse(text);
    } catch (parseError) {
      console.error("JSON Parse Error:", parseError, "Raw Text:", text);
      throw new Error(
        "The analysis result was formatted incorrectly by the AI. \n\n" +
        "Tip: Try simplifying your Excel sheet by removing unrelated data or ensuring that the salary components are clearly labeled in a single table format."
      );
    }

    if (!Array.isArray(parsed) || parsed.length === 0) {
      throw new Error(
        "No salary components or formulas were identified in the uploaded sheet. \n\n" +
        "Checklist for your Excel file:\n" +
        "1. Are there clear labels for components like 'Basic', 'HRA', or 'CTC'?\n" +
        "2. Are you using actual Excel formulas (starting with '=') for calculations?\n" +
        "3. Is the data organized in a readable grid (not scattered randomly)?"
      );
    }

    // Enforce sequence numbers strictly starting from 5 and increasing by 5 based on dependencies
    parsed.sort((a, b) => a.sequenceNumber - b.sequenceNumber);
    parsed = parsed.map((item: any, index: number) => ({
      ...item,
      sequenceNumber: (index + 1) * 5,
      // Normalize formula to strictly enforce CASE BEGIN WHEN ... THEN ... ELSE ... END syntax
      formula: normalizeFormulaSyntax(item.formula || "")
    }));

    // Finally, sort by rowNumber so the UI displays them in the order they appear in the Excel sheet
    parsed.sort((a, b) => a.rowNumber - b.rowNumber);

    return parsed;
  } catch (e: any) {
    console.error("Analysis Error:", e);
    
    // If it's already one of our custom errors, just rethrow it
    if (e.message && (e.message.includes("Tip:") || e.message.includes("Checklist"))) {
      throw e;
    }

    // Handle specific Gemini API errors if they occur
    if (e.message?.includes("safety")) {
      throw new Error("The content of the Excel sheet was flagged by safety filters. Please ensure the file contains only standard payroll data.");
    }

    throw new Error(
      `An unexpected error occurred during analysis: ${e.message || "Unknown error"}. \n\n` +
      "Please try refreshing the page and re-uploading a clean version of your CTC sheet."
    );
  }
}
