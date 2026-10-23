import * as XLSX from 'xlsx';

export interface ExcelCellData {
  address: string;
  value: any;
  formula?: string;
}

export interface ParsedExcelResult {
  data: string;
  preview: any[];
  sheetNames: string[];
  activeSheet: string;
  worksheet?: any;
}

export async function parseExcelFile(file: File, targetSheetName?: string): Promise<ParsedExcelResult> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array', cellFormula: true });
        console.log("Workbook loaded, sheets:", workbook.SheetNames);
        
        const sheetNames = workbook.SheetNames;
        let activeSheet = targetSheetName;
        
        // If no target sheet specified, try to find the active tab, fallback to first sheet
        if (!activeSheet || !sheetNames.includes(activeSheet)) {
          const activeTabIdx = (workbook.Workbook?.Views?.[0] as any)?.activeTab || 0;
          activeSheet = sheetNames[activeTabIdx] || sheetNames[0];
        }
        
        const worksheet = workbook.Sheets[activeSheet];
        const cells: string[] = [];
        
        // Get preview from the selected sheet
        const preview = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" }).slice(0, 15);

        if (worksheet['!ref']) {
          const range = XLSX.utils.decode_range(worksheet['!ref']);
          cells.push(`\n--- Sheet: ${activeSheet} ---`);
          
          for (let R = range.s.r; R <= range.e.r; ++R) {
            // Safety break for extremely large sheets
            if (cells.length >= 2000) break;
            
            for (let C = range.s.c; C <= range.e.c; ++C) {
              const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
              const cell = worksheet[cellAddress];
              
              // Only include cells that have a value or a formula
              if (cell && (cell.v !== undefined || cell.f !== undefined)) {
                let cellInfo = `${cellAddress}: ${cell.v}`;
                if (cell.f) {
                  cellInfo += ` (Formula: =${cell.f})`;
                }
                cells.push(cellInfo);
                
                // Safety break for extremely large sheets
                if (cells.length >= 2000) break;
              }
            }
          }
        }
        
        // Keep cells
        const limitedCells = cells.slice(0, 1000);
        resolve({ 
          data: limitedCells.join('\n'),
          preview: preview,
          sheetNames: sheetNames,
          activeSheet: activeSheet,
          worksheet: worksheet
        });
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

export function exportToExcel(data: any[], fileName: string) {
  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "CTC Formulas");
  XLSX.writeFile(workbook, `${fileName}.xlsx`);
}
