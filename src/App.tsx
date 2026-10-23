import React, { useState, useCallback, useMemo } from 'react';
import { useDropzone } from 'react-dropzone';
import { 
  FileSpreadsheet, 
  Upload, 
  Download, 
  Copy, 
  Check, 
  Loader2, 
  AlertCircle,
  Calculator,
  RefreshCw,
  ShieldCheck,
  AlertTriangle,
  Edit3,
  Save,
  X,
  Filter,
  CheckCircle2,
  Sparkles,
  Plus,
  ArrowRight
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { parseExcelFile, exportToExcel } from './lib/excel-parser';
import { analyzeCTCFormulas, PayHeadFormula, normalizeFormulaSyntax } from './services/gemini';
import { 
  validateAllFormulas, 
  validateFormulaSafetyAndSyntax, 
  convertFormulaToJs, 
  detectCircularDependencies 
} from './lib/formula-validator';
import { cn } from './lib/utils';

function renderFormulaTokens(formula: string) {
  if (!formula) return null;
  const parts = formula.split(/(\b(?:CASE|BEGIN|WHEN|THEN|ELSE|END|And|Or|Not)\b)/g);
  return parts.map((part, i) => {
    const upper = part.toUpperCase();
    if (['CASE', 'BEGIN', 'WHEN', 'THEN', 'ELSE', 'END'].includes(upper)) {
      return (
        <span 
          key={i} 
          className="font-bold text-indigo-700 bg-indigo-50 border border-indigo-200/70 px-1.5 py-0.5 rounded text-xs mx-0.5 inline-block shadow-xs"
        >
          {upper}
        </span>
      );
    }
    if (['AND', 'OR', 'NOT'].includes(upper)) {
      return (
        <span 
          key={i} 
          className="font-semibold text-emerald-700 bg-emerald-50 border border-emerald-200/60 px-1 py-0.5 rounded text-xs mx-0.5 inline-block"
        >
          {part}
        </span>
      );
    }
    return <span key={i}>{part}</span>;
  });
}

export default function App() {
  console.log("[App] Rendering...");
  const [file, setFile] = useState<File | null>(null);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [results, setResults] = useState<PayHeadFormula[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [copiedIndex, setCopiedIndex] = useState<number | null>(null);
  const [previewData, setPreviewData] = useState<any[]>([]);
  const [rawData, setRawData] = useState<string>("");
  const [activeSheetName, setActiveSheetName] = useState<string>("");
  const [showRawData, setShowRawData] = useState(false);
  
  const [calculatedValues, setCalculatedValues] = useState<Record<string, number> | null>(null);
  const [fixedValues, setFixedValues] = useState<Record<string, number>>({});

  // Manual editing and validation states
  const [editingIndex, setEditingIndex] = useState<number | null>(null);
  const [editDraftFormula, setEditDraftFormula] = useState<string>("");
  const [filterFlaggedOnly, setFilterFlaggedOnly] = useState<boolean>(false);
  const [saveToast, setSaveToast] = useState<string | null>(null);

  const isLocked = file !== null || isAnalyzing;

  // Validation across all current results
  const validationSummary = useMemo(() => {
    if (!results || results.length === 0) return null;
    return validateAllFormulas(results);
  }, [results]);

  // Live validation for the item currently being edited
  const draftValidation = useMemo(() => {
    if (editingIndex === null || !results[editingIndex]) return null;
    const currentItem = results[editingIndex];
    const draftItem: PayHeadFormula = {
      ...currentItem,
      formula: editDraftFormula
    };
    const tempResults = [...results];
    tempResults[editingIndex] = draftItem;
    const cycleInfo = detectCircularDependencies(tempResults);
    return validateFormulaSafetyAndSyntax(draftItem, tempResults, cycleInfo);
  }, [editingIndex, editDraftFormula, results]);

  const onDrop = useCallback((acceptedFiles: File[]) => {
    if (isLocked) return;
    
    if (acceptedFiles.length > 0) {
      const selectedFile = acceptedFiles[0];
      setFile(selectedFile);
      setError(null);
      setResults([]);
      setPreviewData([]);
      setRawData("");
      setActiveSheetName("");
      setCalculatedValues(null);
      setFixedValues({});
      setEditingIndex(null);
      setEditDraftFormula("");
      setFilterFlaggedOnly(false);
      
      parseExcelFile(selectedFile).then(res => {
        setPreviewData(res.preview);
        setRawData(res.data);
        setActiveSheetName(res.activeSheet);
      }).catch(err => {
        console.error("Preview failed:", err);
        setError('Failed to read the Excel file.');
      });
    }
  }, [isLocked]);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls']
    },
    multiple: false,
    noClick: isLocked,
    noKeyboard: isLocked
  } as any);

  const calculateFitment = (fixedVals: Record<string, number>, currentResults: PayHeadFormula[] = results) => {
    if (currentResults.length === 0) return;
    
    // Find the base item (or fallback to the first item if none is marked)
    const baseItem = currentResults.find(r => r.isBaseValue) || currentResults[0];
    const ctcValue = fixedVals[baseItem.jsVariableName] ?? baseItem.defaultValue ?? 0;

    const values: Record<string, number> = { 
      Total_CTC: ctcValue,
      CTC: ctcValue,
      TotalCTC: ctcValue
    };

    // Evaluate formulas in their dependency sequence order, not display order
    const calculationOrder = [...currentResults].sort((a, b) => a.sequenceNumber - b.sequenceNumber);

    for (const item of calculationOrder) {
      try {
        if (item.isFixedValue || item.jsVariableName === baseItem.jsVariableName) {
          values[item.jsVariableName] = fixedVals[item.jsVariableName] ?? item.defaultValue ?? 0;
        } else {
          if (!item.jsFormula) continue;
          const argNames = Object.keys(values);
          const argValues = Object.values(values);
          const func = new Function(...argNames, `"use strict"; return (${item.jsFormula});`);
          const result = func(...argValues);
          values[item.jsVariableName] = Number(result) || 0;
        }
      } catch (e) {
        console.error("Evaluation error for", item.payHead, e);
        values[item.jsVariableName] = 0;
      }
    }
    setCalculatedValues(values);
  };

  const handleAnalyze = async () => {
    if (!file || isAnalyzing) return;
    
    console.log("[Analysis] Starting analysis for file:", file.name);
    setIsAnalyzing(true);
    setError(null);
    setCalculatedValues(null);
    setFixedValues({});
    setEditingIndex(null);

    try {
      const { data } = await parseExcelFile(file, activeSheetName);
      console.log("[Analysis] Excel parsed successfully, data length:", data.length);
      
      if (!data || data.length < 10) {
        throw new Error("The Excel file seems to be empty or contains no readable data. Please ensure the sheet has content.");
      }
      
      const formulas = await analyzeCTCFormulas(data);
      console.log("[Analysis] Formula extraction complete. Found components:", formulas.length);
      
      if (!formulas || formulas.length === 0) {
        throw new Error("AI was unable to identify any pay head formulas in this sheet. Please check the sheet structure.");
      }
      
      setResults(formulas);
      
      const initialFixedValues: Record<string, number> = {};
      formulas.forEach((f) => {
        if (f.isFixedValue || f.isBaseValue) {
          initialFixedValues[f.jsVariableName] = f.defaultValue || 0;
        }
      });
      setFixedValues(initialFixedValues);
      calculateFitment(initialFixedValues, formulas);
    } catch (err: any) {
      console.error("[Analysis] Error during process:", err);
      setError(err.message || 'An unexpected error occurred during analysis.');
    } finally {
      setIsAnalyzing(false);
      console.log("[Analysis] Process finished.");
    }
  };

  // Editing helpers
  const handleStartEdit = (index: number) => {
    setEditingIndex(index);
    setEditDraftFormula(results[index].formula || "");
  };

  const handleCancelEdit = () => {
    setEditingIndex(null);
    setEditDraftFormula("");
  };

  const handleInsertToken = (token: string) => {
    setEditDraftFormula(prev => {
      if (!prev) return token;
      // Add space if needed
      const needsSpace = !prev.endsWith(" ") && !token.startsWith(" ");
      return prev + (needsSpace ? " " : "") + token;
    });
  };

  const handleSaveFormula = (index: number) => {
    const targetItem = results[index];
    if (!targetItem) return;

    // 1. Normalize formula syntax (CASE BEGIN WHEN ... THEN ... ELSE ... END)
    const normalized = normalizeFormulaSyntax(editDraftFormula);

    // 2. Transpile to safe JS expression
    const { jsExpression } = convertFormulaToJs(normalized, results, targetItem.jsVariableName);

    // 3. Update the item in results array
    const updatedResults = [...results];
    updatedResults[index] = {
      ...targetItem,
      formula: normalized,
      jsFormula: jsExpression
    };

    setResults(updatedResults);
    setEditingIndex(null);
    setEditDraftFormula("");

    // 4. Recalculate fitment with the updated formulas
    calculateFitment(fixedValues, updatedResults);

    // 5. User feedback
    setSaveToast(`Formula for "${targetItem.payHead}" updated and recalculated!`);
    setTimeout(() => setSaveToast(null), 3500);
  };

  const handleCopy = (text: string, index: number) => {
    try {
      navigator.clipboard.writeText(text);
      setCopiedIndex(index);
      setTimeout(() => setCopiedIndex(null), 2000);
    } catch (err) {
      console.error("Failed to copy text:", err);
    }
  };

  const handleExport = () => {
    if (results.length === 0) return;
    try {
      const exportData = results.map(item => {
        const itemVal = validationSummary?.resultsByVarName[item.jsVariableName];
        const statusStr = itemVal?.hasErrors 
          ? `Error: ${itemVal.issues.map(i => i.message).join('; ')}`
          : itemVal?.hasWarnings
          ? `Warning: ${itemVal.issues.map(i => i.message).join('; ')}`
          : 'Safe & Verified';

        return {
          "Display Order": item.rowNumber,
          "Calculation Sequence": item.sequenceNumber,
          "Pay Head": item.payHead,
          "Formula (CASE Syntax)": item.formula,
          "Description": item.description,
          "Validation Status": statusStr,
          "Type": item.isBaseValue ? "Base Value (Root)" : item.isFixedValue ? "Manual / Value-Based" : "Formula Calculated",
          "Calculated / Fixed Value": calculatedValues ? (calculatedValues[item.jsVariableName] ?? item.defaultValue ?? 0) : item.defaultValue ?? 0
        };
      });
      exportToExcel(exportData, `CTC_Formulas_${new Date().getTime()}`);
    } catch (err) {
      console.error("Export failed:", err);
      setError("Failed to export results to Excel.");
    }
  };

  const reset = () => {
    console.log("[App] Resetting state for new analysis");
    setFile(null);
    setResults([]);
    setPreviewData([]);
    setError(null);
    setRawData("");
    setActiveSheetName("");
    setCalculatedValues(null);
    setEditingIndex(null);
    setEditDraftFormula("");
    setFilterFlaggedOnly(false);
  };

  // Filter items if user toggles "Flagged Only"
  const displayedResults = useMemo(() => {
    if (!filterFlaggedOnly || !validationSummary) return results;
    return results.filter(item => {
      const v = validationSummary.resultsByVarName[item.jsVariableName];
      return v && (v.hasErrors || v.hasWarnings);
    });
  }, [results, filterFlaggedOnly, validationSummary]);

  return (
    <div className="min-h-screen bg-background font-sans selection:bg-primary/10 selection:text-primary">
      {/* Toast Notification */}
      <AnimatePresence>
        {saveToast && (
          <motion.div
            initial={{ opacity: 0, y: -20 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -20 }}
            className="fixed top-20 right-6 z-50 bg-emerald-600 text-white px-5 py-3 rounded-2xl shadow-xl flex items-center gap-3 border border-emerald-500 font-medium text-sm backdrop-blur-md"
          >
            <CheckCircle2 size={18} className="text-white shrink-0" />
            <span>{saveToast}</span>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Header */}
      <header className="sticky top-0 z-40 w-full border-b border-border bg-white/80 backdrop-blur-md">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-primary rounded-xl flex items-center justify-center text-white shadow-lg shadow-primary/20 transition-transform hover:scale-105">
              <Calculator size={22} strokeWidth={2.5} />
            </div>
            <div>
              <h1 className="font-bold text-lg tracking-tight text-text-main leading-none">CTC Analyzer</h1>
              <p className="text-[10px] font-medium text-text-muted uppercase tracking-wider mt-1">Professional Formula Extractor & Validator</p>
            </div>
          </div>
          <div className="flex items-center gap-2 sm:gap-4">
            {results.length > 0 && (
              <button onClick={handleExport} className="btn-primary py-1.5 px-3 text-sm">
                <Download size={16} />
                <span className="hidden sm:inline">Export Results</span>
              </button>
            )}
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-10">
        <div className={cn(
          "grid gap-10 transition-all duration-700 ease-in-out",
          file ? "lg:grid-cols-[380px_1fr]" : "grid-cols-1 max-w-3xl mx-auto"
        )}>
          {/* Left Column: Input & Context */}
          <div className="space-y-8">
            <motion.div 
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              className="space-y-6"
            >
              <div
                {...getRootProps()}
                className={cn(
                  "relative rounded-3xl border-2 border-dashed transition-all duration-500 p-10 overflow-hidden",
                  isLocked 
                    ? "border-border bg-slate-50/50 cursor-default" 
                    : "group cursor-pointer border-border bg-white hover:border-primary hover:bg-slate-50",
                  isDragActive && !isLocked && "border-primary bg-primary/5 ring-4 ring-primary/10",
                  file && "p-8"
                )}
              >
                <input {...getInputProps()} />
                
                {!isLocked && (
                  <div className="absolute -right-10 -bottom-10 w-40 h-40 bg-primary/5 rounded-full blur-3xl group-hover:bg-primary/10 transition-colors" />
                )}
                
                <div className="relative flex flex-col items-center justify-center space-y-4">
                  <div className={cn(
                    "w-16 h-16 rounded-2xl flex items-center justify-center transition-all duration-500 shadow-sm",
                    isLocked ? "bg-slate-100 text-slate-400" :
                    isDragActive ? "bg-primary text-white scale-110" : "bg-slate-100 text-text-muted group-hover:bg-primary/10 group-hover:text-primary group-hover:shadow-md"
                  )}>
                    {file ? <FileSpreadsheet size={32} /> : <Upload size={32} />}
                  </div>
                  <div className="text-center">
                    <p className={cn("text-base font-semibold", isLocked ? "text-slate-500" : "text-text-main")}>
                      {file ? file.name : "Drop your CTC sheet here"}
                    </p>
                    <p className="text-xs text-text-muted mt-1">
                      {file ? `${(file.size / 1024).toFixed(1)} KB` : "Supports .xlsx and .xls formats"}
                    </p>
                  </div>
                </div>
              </div>

              {file && !results.length && (
                <motion.button
                  initial={{ opacity: 0, scale: 0.95 }}
                  animate={{ opacity: 1, scale: 1 }}
                  onClick={handleAnalyze}
                  disabled={isAnalyzing}
                  className={cn(
                    "w-full py-4 rounded-2xl font-bold text-white transition-all duration-300 flex items-center justify-center gap-3 shadow-xl shadow-primary/20",
                    isAnalyzing ? "bg-slate-400 cursor-not-allowed" : "bg-primary hover:bg-primary-hover hover:-translate-y-1 active:translate-y-0"
                  )}
                >
                  {isAnalyzing ? <Loader2 className="animate-spin" size={22} /> : <Calculator size={22} />}
                  {isAnalyzing ? "Processing Logic..." : "Start Analysis"}
                </motion.button>
              )}
            </motion.div>

            {/* Sheet Preview */}
            <AnimatePresence>
              {previewData.length > 0 && (
                <motion.div 
                  initial={{ opacity: 0, y: 20 }}
                  animate={{ opacity: 1, y: 0 }}
                  className="glass-card overflow-hidden"
                >
                  <div className="bg-slate-50/50 px-5 py-3 border-b border-border flex items-center justify-between">
                    <div className="flex items-center gap-2">
                      <div className="w-2 h-2 bg-primary rounded-full animate-pulse" />
                      <span className="text-[11px] font-bold text-text-muted uppercase tracking-widest">Data Preview</span>
                    </div>
                    <button 
                      onClick={() => setShowRawData(!showRawData)}
                      className="text-[10px] font-bold text-primary hover:text-primary-hover transition-colors uppercase tracking-wider"
                    >
                      {showRawData ? "Table View" : "Raw View"}
                    </button>
                  </div>
                  
                  <div className="max-h-[350px] overflow-auto">
                    {showRawData ? (
                      <div className="p-4 bg-slate-900 text-slate-300 font-mono text-[10px] leading-relaxed whitespace-pre-wrap">
                        {rawData || "No data extracted."}
                      </div>
                    ) : (
                      <table className="w-full text-[11px] text-left border-collapse">
                        <tbody className="divide-y divide-slate-100">
                          {previewData.map((row, i) => (
                            <tr key={i} className="hover:bg-slate-50/50 transition-colors">
                              {Array.isArray(row) ? row.map((cell, j) => (
                                <td key={j} className="px-4 py-2.5 whitespace-nowrap text-text-muted border-r border-slate-50 last:border-0">
                                  {cell?.toString() || ''}
                                </td>
                              )) : (
                                <td className="px-4 py-2.5 text-text-muted italic opacity-50">Empty row</td>
                              )}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    )}
                  </div>
                </motion.div>
              )}
            </AnimatePresence>

            {error && (
              <motion.div 
                initial={{ opacity: 0, x: -20 }}
                animate={{ opacity: 1, x: 0 }}
                className="bg-red-50 border border-red-100 text-red-700 p-5 rounded-2xl flex items-start gap-4 shadow-sm"
              >
                <AlertCircle className="shrink-0 mt-0.5 text-red-500" size={20} />
                <div className="text-sm font-medium leading-relaxed">
                  {error}
                </div>
              </motion.div>
            )}
          </div>

          {/* Right Column: Results */}
          <div className="space-y-6">
            <AnimatePresence mode="wait">
              {isAnalyzing ? (
                <motion.div
                  key="analyzing"
                  initial={{ opacity: 0, scale: 0.95 }}
                  animate={{ opacity: 1, scale: 1 }}
                  exit={{ opacity: 0, scale: 1.05 }}
                  className="flex flex-col items-center justify-center py-32 space-y-8 text-center glass-card"
                >
                  <div className="relative">
                    <div className="w-20 h-20 border-4 border-slate-100 border-t-primary rounded-full animate-spin" />
                    <div className="absolute inset-0 flex items-center justify-center">
                      <Calculator size={24} className="text-primary animate-pulse" />
                    </div>
                  </div>
                  <div className="space-y-3 px-6">
                    <h3 className="text-2xl font-bold text-text-main tracking-tight">AI is Decoding Formulas</h3>
                    <p className="text-text-muted max-w-sm mx-auto leading-relaxed">
                      Reverse-engineering the payroll logic from your Excel data. This typically takes a few moments.
                    </p>
                  </div>
                </motion.div>
              ) : results.length > 0 ? (
                <motion.div
                  key="results"
                  initial={{ opacity: 0, y: 20 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: -20 }}
                  className="space-y-6"
                >
                  {/* Header & Controls */}
                  <div className="flex flex-col sm:flex-row sm:items-end justify-between gap-4 px-2">
                    <div className="space-y-1">
                      <h2 className="text-3xl font-extrabold text-text-main tracking-tight">Extracted Logic</h2>
                      <p className="text-text-muted font-medium">Verified pay head calculations and safe browser execution</p>
                    </div>
                    <div className="flex items-center gap-3">
                      {validationSummary && (validationSummary.errorCount > 0 || validationSummary.warningCount > 0) && (
                        <button
                          onClick={() => setFilterFlaggedOnly(!filterFlaggedOnly)}
                          className={cn(
                            "py-1.5 px-3 rounded-full text-xs font-bold flex items-center gap-1.5 border transition-all",
                            filterFlaggedOnly
                              ? "bg-red-600 text-white border-red-600 shadow-sm"
                              : "bg-white text-red-700 border-red-200 hover:bg-red-50"
                          )}
                        >
                          <Filter size={13} />
                          <span>{filterFlaggedOnly ? "Showing Flagged Only" : `View Flagged (${validationSummary.errorCount + validationSummary.warningCount})`}</span>
                        </button>
                      )}
                      <div className="flex items-center gap-2 bg-white border border-border px-3 py-1.5 rounded-full shadow-sm">
                        <span className="w-2 h-2 bg-green-500 rounded-full" />
                        <span className="text-xs font-bold text-text-main uppercase tracking-wider">
                          {results.length} Components
                        </span>
                      </div>
                    </div>
                  </div>

                  {/* Top-Level Validation Health Banner */}
                  {validationSummary && (
                    <div className={cn(
                      "rounded-2xl p-5 border transition-all duration-300 shadow-sm flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4",
                      validationSummary.allValid
                        ? "bg-emerald-50/70 border-emerald-200 text-emerald-900"
                        : "bg-red-50/90 border-red-200 text-red-950"
                    )}>
                      <div className="flex items-start sm:items-center gap-3.5">
                        <div className={cn(
                          "w-10 h-10 rounded-xl flex items-center justify-center shrink-0 shadow-sm",
                          validationSummary.allValid
                            ? "bg-emerald-600 text-white"
                            : "bg-red-600 text-white"
                        )}>
                          {validationSummary.allValid ? (
                            <ShieldCheck size={22} strokeWidth={2.5} />
                          ) : (
                            <AlertTriangle size={22} strokeWidth={2.5} />
                          )}
                        </div>
                        <div className="space-y-0.5">
                          <div className="flex items-center gap-2">
                            <h4 className="font-bold text-sm tracking-tight">
                              {validationSummary.allValid
                                ? "Formula Validation Passed: Safe for Browser Execution"
                                : `Validation Issues Detected (${validationSummary.errorCount} Flagged)`}
                            </h4>
                            <span className={cn(
                              "text-[10px] font-bold px-2 py-0.5 rounded-full uppercase tracking-wider",
                              validationSummary.allValid
                                ? "bg-emerald-200/70 text-emerald-800"
                                : "bg-red-200 text-red-800"
                            )}>
                              {validationSummary.allValid ? "100% Verified" : "Action Required"}
                            </span>
                          </div>
                          <p className="text-xs text-text-muted leading-relaxed">
                            {validationSummary.allValid ? (
                              `All ${validationSummary.totalFormulas} extracted formulas checked. Zero circular dependencies, zero illegal functions, syntax balanced.`
                            ) : (
                              `Found ${validationSummary.errorCount} formula error(s) and ${validationSummary.warningCount} warning(s). Review highlighted components below or click 'Edit Formula' to customize.`
                            )}
                          </p>
                          {validationSummary.cycles.length > 0 && (
                            <div className="mt-1.5 inline-flex items-center gap-1.5 text-xs font-semibold text-red-700 bg-red-100/80 px-2.5 py-1 rounded-lg">
                              <RefreshCw size={12} className="animate-spin text-red-600" />
                              <span>
                                Circular Loop: {validationSummary.cycles[0].map(v => {
                                  const c = results.find(r => r.jsVariableName === v);
                                  return c ? c.payHead : v;
                                }).join(' ➔ ')}
                              </span>
                            </div>
                          )}
                        </div>
                      </div>

                      {/* Right action in banner */}
                      <div className="shrink-0 flex items-center gap-2">
                        {validationSummary.errorCount > 0 && !filterFlaggedOnly && (
                          <button
                            onClick={() => setFilterFlaggedOnly(true)}
                            className="bg-red-600 hover:bg-red-700 text-white text-xs font-bold px-3 py-1.5 rounded-xl shadow-xs transition-colors"
                          >
                            Filter Flagged Items
                          </button>
                        )}
                        {filterFlaggedOnly && (
                          <button
                            onClick={() => setFilterFlaggedOnly(false)}
                            className="bg-white hover:bg-slate-100 text-slate-700 text-xs font-bold px-3 py-1.5 rounded-xl border border-slate-200 shadow-xs transition-colors"
                          >
                            Show All Items
                          </button>
                        )}
                      </div>
                    </div>
                  )}

                  {/* Components List */}
                  <div className="grid gap-5">
                    {displayedResults.map((item) => {
                      // Original index in the full results array
                      const originalIndex = results.findIndex(r => r.jsVariableName === item.jsVariableName);
                      const isEditing = editingIndex === originalIndex;
                      const itemVal = validationSummary?.resultsByVarName[item.jsVariableName];
                      const hasErrors = itemVal?.hasErrors ?? false;
                      const hasWarnings = itemVal?.hasWarnings ?? false;

                      return (
                        <motion.div
                          key={item.jsVariableName || originalIndex}
                          initial={{ opacity: 0, y: 15 }}
                          animate={{ opacity: 1, y: 0 }}
                          className={cn(
                            "glass-card p-6 transition-all duration-300 group relative overflow-hidden",
                            hasErrors
                              ? "border-red-300 ring-2 ring-red-400/30 bg-gradient-to-b from-red-50/20 to-white"
                              : hasWarnings
                              ? "border-amber-300 ring-1 ring-amber-300/40"
                              : "hover:shadow-lg hover:shadow-primary/5",
                            isEditing && "ring-2 ring-primary border-primary bg-white shadow-xl"
                          )}
                        >
                          {/* Subtle background sequence number */}
                          <div className="absolute -right-4 -top-4 text-9xl font-black text-slate-50 pointer-events-none select-none group-hover:text-slate-100 transition-colors">
                            {item.sequenceNumber || (originalIndex + 1) * 5}
                          </div>

                          <div className="relative space-y-5">
                            {/* Card Top: Ordering & Component Title & Status Badges */}
                            <div className="flex flex-col sm:flex-row sm:items-start justify-between gap-4">
                              <div className="flex items-start gap-4">
                                <div className="flex items-center gap-2 mt-0.5 shrink-0">
                                  <div className="flex flex-col items-center gap-1">
                                    <span className="text-[9px] font-bold text-text-muted uppercase tracking-wider">Display</span>
                                    <div className="w-9 h-9 bg-slate-700 text-white rounded-xl flex items-center justify-center font-bold text-sm shadow-sm">
                                      {item.rowNumber || originalIndex + 1}
                                    </div>
                                  </div>
                                  <div className="flex flex-col items-center gap-1">
                                    <span className="text-[9px] font-bold text-text-muted uppercase tracking-wider">Sequence</span>
                                    <div className="w-9 h-9 bg-primary text-white rounded-xl flex items-center justify-center font-bold text-sm shadow-sm">
                                      {item.sequenceNumber || (originalIndex + 1) * 5}
                                    </div>
                                  </div>
                                </div>

                                <div className="space-y-1">
                                  <div className="flex flex-wrap items-center gap-2">
                                    <span className="text-[9px] font-black text-primary uppercase tracking-[0.2em] bg-primary/5 px-2 py-0.5 rounded">
                                      {item.isBaseValue ? "Base Anchor" : item.isFixedValue ? "Manual Input" : "Formula Head"}
                                    </span>
                                    {item.formula?.toUpperCase().includes('CASE') && (
                                      <span className="text-[9px] font-bold text-indigo-700 bg-indigo-50 border border-indigo-200/80 px-2 py-0.5 rounded uppercase tracking-wider">
                                        CASE ... END Logic
                                      </span>
                                    )}

                                    {/* Verification Status Pill */}
                                    {hasErrors ? (
                                      <span className="text-[10px] font-bold text-red-700 bg-red-100 border border-red-300 px-2.5 py-0.5 rounded-full flex items-center gap-1">
                                        <AlertCircle size={12} className="text-red-600" />
                                        Validation Issue
                                      </span>
                                    ) : hasWarnings ? (
                                      <span className="text-[10px] font-bold text-amber-700 bg-amber-100 border border-amber-300 px-2.5 py-0.5 rounded-full flex items-center gap-1">
                                        <AlertTriangle size={12} className="text-amber-600" />
                                        Warning
                                      </span>
                                    ) : (
                                      <span className="text-[10px] font-bold text-emerald-700 bg-emerald-50 border border-emerald-200/80 px-2.5 py-0.5 rounded-full flex items-center gap-1">
                                        <ShieldCheck size={12} className="text-emerald-600" />
                                        Safe & Parsed
                                      </span>
                                    )}

                                    <h3 className="text-xl font-bold text-text-main tracking-tight ml-1">{item.payHead}</h3>
                                  </div>
                                  <p className="text-xs font-medium text-text-muted leading-relaxed max-w-xl">
                                    {item.description}
                                  </p>
                                </div>
                              </div>

                              {/* Action Buttons: Edit Formula & Copy */}
                              <div className="flex items-center gap-2 self-start shrink-0">
                                {!isEditing && (
                                  <button
                                    onClick={() => handleStartEdit(originalIndex)}
                                    className="px-3 py-1.5 rounded-xl border border-slate-200 bg-white hover:border-primary hover:text-primary text-slate-700 font-semibold text-xs flex items-center gap-1.5 shadow-xs transition-all active:scale-95"
                                    title="Edit formula string manually"
                                  >
                                    <Edit3 size={13} />
                                    <span>Edit Formula</span>
                                  </button>
                                )}
                              </div>
                            </div>

                            {/* Validation Issue Banner on Card (if flagged) */}
                            {itemVal && itemVal.issues.length > 0 && !isEditing && (
                              <div className={cn(
                                "rounded-2xl p-4 border flex flex-col gap-2.5 transition-all",
                                hasErrors ? "bg-red-50/90 border-red-200 text-red-950" : "bg-amber-50/90 border-amber-200 text-amber-950"
                              )}>
                                <div className="flex items-center justify-between">
                                  <div className="flex items-center gap-2 font-bold text-xs uppercase tracking-wider">
                                    <AlertCircle size={16} className={hasErrors ? "text-red-600 shrink-0" : "text-amber-600 shrink-0"} />
                                    <span>{hasErrors ? "Formula Flagged by Validation Layer" : "Validation Advisory"}</span>
                                  </div>
                                  <button
                                    onClick={() => handleStartEdit(originalIndex)}
                                    className="text-xs font-bold text-primary hover:underline flex items-center gap-1"
                                  >
                                    <span>Edit to Fix</span>
                                    <ArrowRight size={12} />
                                  </button>
                                </div>

                                <div className="space-y-2">
                                  {itemVal.issues.map((issue, idx) => (
                                    <div key={idx} className="bg-white/80 p-3 rounded-xl border border-red-100/60 shadow-2xs space-y-1">
                                      <div className="flex items-center gap-2">
                                        <span className={cn(
                                          "text-[9px] font-black uppercase px-2 py-0.5 rounded",
                                          issue.type === 'circular_dependency' ? "bg-purple-100 text-purple-800" :
                                          issue.type === 'illegal_function' ? "bg-red-100 text-red-800" : "bg-amber-100 text-amber-800"
                                        )}>
                                          {issue.type.replace('_', ' ')}
                                        </span>
                                        <span className="font-bold text-xs text-red-900">{issue.message}</span>
                                      </div>
                                      {issue.details && (
                                        <p className="text-xs text-slate-600 leading-relaxed pl-1">{issue.details}</p>
                                      )}
                                      {issue.suggestedFix && (
                                        <div className="text-xs font-semibold text-emerald-800 bg-emerald-50/60 px-2.5 py-1 rounded-lg mt-1 inline-block">
                                          💡 Suggested Resolution: {issue.suggestedFix}
                                        </div>
                                      )}
                                    </div>
                                  ))}
                                </div>
                              </div>
                            )}

                            {/* Standard Formula Display View */}
                            {!isEditing ? (
                              <div className="relative group/formula">
                                <div className="bg-slate-50 rounded-2xl p-5 font-mono text-sm text-slate-800 border border-slate-100 break-all pr-14 group-hover/formula:bg-white group-hover/formula:border-primary/20 transition-all duration-300 leading-relaxed shadow-2xs">
                                  <span className="text-primary/40 mr-2 font-sans font-bold">ƒ</span>
                                  {renderFormulaTokens(item.formula)}
                                </div>
                                <button
                                  onClick={() => handleCopy(item.formula, originalIndex)}
                                  className="absolute top-1/2 -translate-y-1/2 right-3 p-2.5 bg-white rounded-xl shadow-xs border border-border text-text-muted hover:text-primary hover:border-primary hover:shadow-md transition-all active:scale-90"
                                  title="Copy Formula"
                                >
                                  {copiedIndex === originalIndex ? <Check size={16} className="text-green-500" /> : <Copy size={16} />}
                                </button>
                              </div>
                            ) : (
                              /* Interactive Inline Formula String Editor */
                              <div className="bg-slate-50/80 rounded-2xl p-5 border-2 border-primary/40 space-y-4">
                                <div className="flex items-center justify-between">
                                  <div className="flex items-center gap-2">
                                    <Sparkles size={16} className="text-primary" />
                                    <span className="text-xs font-bold uppercase tracking-wider text-text-main">
                                      Editing Formula String
                                    </span>
                                  </div>
                                  <span className="text-[11px] text-text-muted font-medium">
                                    Uses CASE BEGIN WHEN ... THEN ... ELSE ... END payroll format
                                  </span>
                                </div>

                                {/* Text Area Input */}
                                <div className="relative">
                                  <textarea
                                    rows={3}
                                    value={editDraftFormula}
                                    onChange={(e) => setEditDraftFormula(e.target.value)}
                                    className="w-full font-mono text-sm p-4 rounded-xl border border-slate-200 bg-white focus:border-primary focus:ring-2 focus:ring-primary/20 outline-none transition-all leading-relaxed"
                                    placeholder="e.g. CASE BEGIN WHEN Basic Salary > 20000 THEN Basic Salary * 0.4 ELSE Basic Salary * 0.3 END"
                                  />
                                </div>

                                {/* Quick-Insert Keyword & Component Helper Pills */}
                                <div className="space-y-2">
                                  <div className="flex items-center gap-2">
                                    <span className="text-[10px] font-bold text-text-muted uppercase tracking-wider">Payroll Keywords:</span>
                                    <div className="flex flex-wrap gap-1.5">
                                      {['CASE BEGIN WHEN', 'THEN', 'ELSE', 'END', 'AND', 'OR', 'NOT'].map((kw) => (
                                        <button
                                          key={kw}
                                          type="button"
                                          onClick={() => handleInsertToken(kw)}
                                          className="text-[11px] font-bold px-2 py-0.5 rounded-lg bg-indigo-50 border border-indigo-200/80 text-indigo-700 hover:bg-indigo-100 hover:scale-105 active:scale-95 transition-all shadow-2xs"
                                        >
                                          +{kw}
                                        </button>
                                      ))}
                                      {['Math.min(', 'Math.max(', 'Math.round('].map((fn) => (
                                        <button
                                          key={fn}
                                          type="button"
                                          onClick={() => handleInsertToken(fn)}
                                          className="text-[11px] font-mono font-semibold px-2 py-0.5 rounded-lg bg-slate-100 border border-slate-200 text-slate-700 hover:bg-slate-200 transition-all shadow-2xs"
                                        >
                                          +{fn}
                                        </button>
                                      ))}
                                    </div>
                                  </div>

                                  <div className="flex items-center gap-2 pt-1">
                                    <span className="text-[10px] font-bold text-text-muted uppercase tracking-wider">Insert Components:</span>
                                    <div className="flex flex-wrap gap-1.5 max-h-24 overflow-y-auto">
                                      <button
                                        type="button"
                                        onClick={() => handleInsertToken('Total CTC')}
                                        className="text-[11px] font-semibold px-2 py-0.5 rounded-lg bg-emerald-50 border border-emerald-200 text-emerald-800 hover:bg-emerald-100 transition-all shadow-2xs"
                                      >
                                        + Total CTC
                                      </button>
                                      {results.map((comp) => {
                                        if (comp.jsVariableName === item.jsVariableName) return null; // Avoid self
                                        return (
                                          <button
                                            key={comp.jsVariableName}
                                            type="button"
                                            onClick={() => handleInsertToken(comp.payHead)}
                                            className="text-[11px] font-semibold px-2 py-0.5 rounded-lg bg-white border border-slate-200 text-slate-700 hover:border-primary hover:text-primary transition-all shadow-2xs"
                                          >
                                            + {comp.payHead}
                                          </button>
                                        );
                                      })}
                                    </div>
                                  </div>
                                </div>

                                {/* Real-time Live Validation Preview */}
                                {draftValidation && (
                                  <div className={cn(
                                    "p-3.5 rounded-xl border flex items-start gap-3 transition-all",
                                    draftValidation.hasErrors 
                                      ? "bg-red-50/90 border-red-200 text-red-900" 
                                      : draftValidation.hasWarnings
                                      ? "bg-amber-50/90 border-amber-200 text-amber-900"
                                      : "bg-emerald-50/90 border-emerald-200 text-emerald-900"
                                  )}>
                                    {draftValidation.hasErrors ? (
                                      <AlertCircle size={16} className="text-red-600 shrink-0 mt-0.5" />
                                    ) : draftValidation.hasWarnings ? (
                                      <AlertTriangle size={16} className="text-amber-600 shrink-0 mt-0.5" />
                                    ) : (
                                      <CheckCircle2 size={16} className="text-emerald-600 shrink-0 mt-0.5" />
                                    )}
                                    <div className="text-xs space-y-1">
                                      <div className="font-bold">
                                        {draftValidation.hasErrors
                                          ? `Draft Validation Failed: ${draftValidation.issues.map(i => i.message).join(' ')}`
                                          : draftValidation.hasWarnings
                                          ? `Draft Warning: ${draftValidation.issues.map(i => i.message).join(' ')}`
                                          : "Draft formula is safe and parsed cleanly for browser execution."}
                                      </div>
                                      {draftValidation.transpiledJsFormula && (
                                        <div className="font-mono text-[10px] text-slate-600 break-all bg-white/60 p-1.5 rounded border border-slate-100">
                                          Compiled JS: {draftValidation.transpiledJsFormula}
                                        </div>
                                      )}
                                    </div>
                                  </div>
                                )}

                                {/* Action Buttons */}
                                <div className="flex items-center justify-end gap-3 pt-2">
                                  <button
                                    type="button"
                                    onClick={handleCancelEdit}
                                    className="px-4 py-2 rounded-xl text-xs font-semibold text-slate-600 bg-white border border-slate-200 hover:bg-slate-100 transition-all flex items-center gap-1.5 shadow-2xs"
                                  >
                                    <X size={14} />
                                    Cancel
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => handleSaveFormula(originalIndex)}
                                    className="px-5 py-2 rounded-xl text-xs font-bold text-white bg-primary hover:bg-primary-hover transition-all flex items-center gap-1.5 shadow-md shadow-primary/20 hover:-translate-y-0.5 active:translate-y-0"
                                  >
                                    <Save size={14} />
                                    Save & Recalculate
                                  </button>
                                </div>
                              </div>
                            )}

                            {/* Calculated Value Display or Editable Input */}
                            {item.isFixedValue || item.isBaseValue ? (
                              <div className="mt-6 pt-4 border-t border-slate-100 flex items-center justify-between bg-blue-50/50 -mx-6 -mb-6 px-6 py-4 rounded-b-2xl">
                                <span className="text-sm font-bold text-blue-800 uppercase tracking-wider">
                                  {item.isBaseValue ? "Base Value (Editable)" : "Manual Input / Value Based (Editable)"}
                                </span>
                                <div className="relative w-40">
                                  <span className="absolute left-3 top-1/2 -translate-y-1/2 text-blue-400 font-medium">₹</span>
                                  <input 
                                    type="number" 
                                    value={fixedValues[item.jsVariableName] ?? item.defaultValue ?? 0}
                                    onChange={(e) => {
                                      const val = e.target.value === '' ? 0 : parseFloat(e.target.value);
                                      const newFixedValues = { ...fixedValues, [item.jsVariableName]: val };
                                      setFixedValues(newFixedValues);
                                      if (calculatedValues) {
                                        calculateFitment(newFixedValues);
                                      }
                                    }}
                                    className="w-full pl-8 pr-3 py-2 rounded-xl border border-blue-200 focus:border-blue-500 focus:ring-2 focus:ring-blue-500/20 outline-none transition-all text-right font-bold text-blue-700 bg-white shadow-sm"
                                  />
                                </div>
                              </div>
                            ) : (
                              calculatedValues && calculatedValues[item.jsVariableName] !== undefined && (
                                <div className="mt-6 pt-4 border-t border-slate-100 flex items-center justify-between bg-green-50/50 -mx-6 -mb-6 px-6 py-4 rounded-b-2xl">
                                  <span className="text-sm font-bold text-green-800 uppercase tracking-wider">Calculated Value</span>
                                  <span className="text-2xl font-black text-green-600">
                                    ₹{calculatedValues[item.jsVariableName].toLocaleString('en-IN', { maximumFractionDigits: 2 })}
                                  </span>
                                </div>
                              )
                            )}
                          </div>
                        </motion.div>
                      );
                    })}
                  </div>

                  <div className="pt-8 flex justify-center">
                    <button 
                      onClick={reset} 
                      className="btn-secondary py-2 px-4 text-sm shadow-sm hover:-translate-y-0.5"
                    >
                      <RefreshCw size={16} />
                      New Salary Fitment
                    </button>
                  </div>
                </motion.div>
              ) : (
                <motion.div 
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  className="flex flex-col items-center justify-center py-32 text-center space-y-6 glass-card border-dashed border-2 bg-slate-50/30"
                >
                  <div className="w-20 h-20 bg-white rounded-3xl shadow-sm border border-border flex items-center justify-center text-slate-300">
                    <FileSpreadsheet size={40} strokeWidth={1.5} />
                  </div>
                  <div className="space-y-2">
                    <h3 className="text-xl font-bold text-text-main tracking-tight">Ready for Analysis</h3>
                    <p className="text-sm text-text-muted max-w-[240px] mx-auto leading-relaxed">
                      Upload your Excel sheet to see automated logic extraction and browser safety validation.
                    </p>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>
          </div>
        </div>
      </main>

      <footer className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-12 text-center border-t border-border mt-10">
        <div className="flex flex-col items-center gap-4">
          <div className="flex items-center gap-2 text-text-muted">
            <Calculator size={14} />
            <span className="text-[11px] font-bold uppercase tracking-[0.2em]">CTC Formula Analyzer & Safety Validator</span>
          </div>
          <p className="text-xs text-text-muted font-medium">
            &copy; {new Date().getFullYear()} Professional Payroll Intelligence. Verified with Safe Browser Evaluation & Circular Dependency Detection.
          </p>
        </div>
      </footer>
    </div>
  );
}
