import * as XLSX from 'xlsx';
import { parseDate, parseAmount, isValidGstin } from './dataSanitizer';

export interface CleaningOptions {
  removeDuplicates: boolean;
  trimWhitespace: boolean;
  standardizeDates: boolean;
  dateFormat: 'DD-MM-YYYY' | 'YYYY-MM-DD' | 'DD/MM/YYYY';
  standardizeNumbers: boolean;
  removeBlankRows: boolean;
  removeBlankCols: boolean;
  cleanGstin: boolean;
  preserveFormulas: boolean;
}

export const DEFAULT_CLEANING_OPTIONS: CleaningOptions = {
  removeDuplicates: false,
  trimWhitespace: true,
  standardizeDates: true,
  dateFormat: 'DD-MM-YYYY',
  standardizeNumbers: true,
  removeBlankRows: false,
  removeBlankCols: false,
  cleanGstin: true,
  preserveFormulas: true,
};

export interface CellChange {
  sheet: string;
  cell: string;
  type: 'whitespace' | 'date' | 'number' | 'gstin';
  oldValue: any;
  newValue: any;
}

export interface SheetAnalysis {
  sheetName: string;
  totalRows: number;
  totalCols: number;
  range: string;
  headers: string[];
  headerRowIndex: number;
  duplicateRowIndices: number[]; // 0-indexed row numbers in worksheet
  blankRowIndices: number[];
  blankColIndices: number[];
  cellsWithWhitespace: { cell: string; val: string; trimmed: string }[];
  unstandardizedDates: { cell: string; val: string; formatted: string }[];
  unstandardizedNumbers: { cell: string; val: any; parsed: number }[];
  formulaCells: string[]; // addresses like 'E10'
  mergedRangesCount: number;
}

export interface ExecutionReport {
  sheetsProcessed: string[];
  totalDuplicatesRemoved: number;
  totalBlankRowsRemoved: number;
  totalBlankColsRemoved: number;
  totalCellsTrimmed: number;
  totalDatesStandardized: number;
  totalNumbersStandardized: number;
  totalGstinsStandardized: number;
  formulasPreserved: number;
  mergedRangesPreserved: number;
  changes: CellChange[];
}

/**
 * Parse an Excel workbook preserving all original formatting, sheet structures,
 * formulas, styles, merged cells, and column metadata.
 */
export function readWorkbookPreserved(data: ArrayBuffer | Uint8Array | string): XLSX.WorkBook {
  return XLSX.read(data, {
    type: typeof data === 'string' ? 'binary' : 'array',
    cellStyles: true,
    cellNF: true,
    cellDates: false, // keep raw serial / text to avoid unwanted timezone conversions
    cellFormula: true,
    cellHTML: true,
  });
}

/**
 * Creates a safe clone of the workbook so pre-processing inspection does not mutate the original.
 */
export function cloneWorkbook(wb: XLSX.WorkBook): XLSX.WorkBook {
  const newWb = XLSX.utils.book_new();
  newWb.Props = wb.Props ? { ...wb.Props } : undefined;
  newWb.Custprops = wb.Custprops ? { ...wb.Custprops } : undefined;

  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    if (!ws) continue;
    const clonedWs: XLSX.WorkSheet = {};

    for (const [key, val] of Object.entries(ws)) {
      if (key.startsWith('!')) {
        if (key === '!merges' && Array.isArray(val)) {
          clonedWs['!merges'] = val.map((m) => ({
            s: { r: m.s.r, c: m.s.c },
            e: { r: m.e.r, c: m.e.c },
          }));
        } else if (key === '!cols' && Array.isArray(val)) {
          clonedWs['!cols'] = val.map((c) => ({ ...c }));
        } else if (key === '!rows' && Array.isArray(val)) {
          clonedWs['!rows'] = val.map((r) => ({ ...r }));
        } else if (key === '!views' && Array.isArray(val)) {
          clonedWs['!views'] = val.map((v) => ({ ...v }));
        } else if (typeof val === 'object' && val !== null) {
          clonedWs[key] = JSON.parse(JSON.stringify(val));
        } else {
          clonedWs[key] = val;
        }
      } else if (typeof val === 'object' && val !== null) {
        clonedWs[key] = { ...(val as XLSX.CellObject) };
      } else {
        clonedWs[key] = val;
      }
    }
    XLSX.utils.book_append_sheet(newWb, clonedWs, name);
  }

  return newWb;
}

/**
 * Inspect a worksheet to identify prospective changes without altering any cell.
 */
export function analyzeWorksheet(
  ws: XLSX.WorkSheet,
  sheetName: string,
  options: CleaningOptions = DEFAULT_CLEANING_OPTIONS
): SheetAnalysis {
  const ref = ws['!ref'];
  if (!ref) {
    return {
      sheetName,
      totalRows: 0,
      totalCols: 0,
      range: 'A1:A1',
      headers: [],
      headerRowIndex: 0,
      duplicateRowIndices: [],
      blankRowIndices: [],
      blankColIndices: [],
      cellsWithWhitespace: [],
      unstandardizedDates: [],
      unstandardizedNumbers: [],
      formulaCells: [],
      mergedRangesCount: 0,
    };
  }

  const range = XLSX.utils.decode_range(ref);
  const totalRows = range.e.r - range.s.r + 1;
  const totalCols = range.e.c - range.s.c + 1;

  // Keywords that strongly indicate an accounting column header
  const headerKeywordsRegex =
    /\b(date|dt|period|inv|bill|vch|voucher|doc|party|vendor|supplier|customer|client|name|particular|description|narration|gstin|gst|pan|code|account|acc|head|amt|amount|taxable|cgst|sgst|igst|tax|total|gross|net|debit|credit|dr|cr|balance|bal|rate|qty|quantity|units|price|serial|s\.no|sl\.no|sr\.no|status|ref|reference|chq|cheque|utr|rem|remark|remarks|hsn|sac)\b/i;

  let headerRowIndex = range.s.r;
  let maxKeywordScore = -1;
  let bestCandidateHeaders: string[] = [];
  const headers: string[] = [];

  const scanLimit = Math.min(range.s.r + 25, range.e.r);
  for (let r = range.s.r; r <= scanLimit; r++) {
    let textCells = 0;
    let score = 0;
    const candidateHeaders: string[] = [];

    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      const valStr = cell && cell.v !== undefined && cell.v !== null ? String(cell.v).trim() : '';
      if (valStr) {
        textCells++;
        candidateHeaders.push(valStr);
        if (headerKeywordsRegex.test(valStr)) {
          score += 2;
        } else if (valStr.length <= 35 && !/^[\d\.,\-\/]+$/.test(valStr)) {
          score += 0.5;
        }
      } else {
        candidateHeaders.push(`Column_${c + 1}`);
      }
    }

    if (textCells >= 2 && score > maxKeywordScore) {
      maxKeywordScore = score;
      headerRowIndex = r;
      bestCandidateHeaders = candidateHeaders;
    }
  }

  // Fallback if no strong header detected
  if (bestCandidateHeaders.length === 0) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r: range.s.r, c });
      const cell = ws[addr];
      bestCandidateHeaders.push(cell && cell.v !== undefined && cell.v !== null ? String(cell.v).trim() : `Column_${c + 1}`);
    }
  }

  headers.push(...bestCandidateHeaders);

  const formulaCells: string[] = [];
  const cellsWithWhitespace: { cell: string; val: string; trimmed: string }[] = [];
  const unstandardizedDates: { cell: string; val: string; formatted: string }[] = [];
  const unstandardizedNumbers: { cell: string; val: any; parsed: number }[] = [];
  const blankRowIndices: number[] = [];
  const blankColIndices: number[] = [];

  // Check columns for completely blank status
  for (let c = range.s.c; c <= range.e.c; c++) {
    let colHasData = false;
    for (let r = range.s.r; r <= range.e.r; r++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '') {
        colHasData = true;
        break;
      }
    }
    if (!colHasData) {
      blankColIndices.push(c);
    }
  }

  // Scan rows
  const rowSignatures = new Map<string, number[]>(); // signature -> row indices

  for (let r = range.s.r; r <= range.e.r; r++) {
    let rowHasData = false;
    const rowValues: string[] = [];

    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];

      if (!cell || cell.v === undefined || cell.v === null || String(cell.v).trim() === '') {
        rowValues.push('');
        continue;
      }

      rowHasData = true;
      const rawVal = cell.v;
      const strVal = String(rawVal);
      rowValues.push(strVal.trim());

      // Check for formulas
      if (cell.f) {
        formulaCells.push(addr);
        continue; // Protect formula cells from text manipulation
      }

      // Check whitespace
      if (cell.t === 's' || typeof rawVal === 'string') {
        const trimmed = strVal.replace(/[\u00A0\u200B]/g, ' ').trim();
        if (strVal !== trimmed && strVal.length > 0) {
          cellsWithWhitespace.push({ cell: addr, val: strVal, trimmed });
        }
      }

      // Check date candidates (only for data rows below header)
      if (r > headerRowIndex && typeof rawVal === 'string') {
        const dateRegex = /\b(\d{1,4}[-/. ]\d{1,2}[-/. ]\d{1,4}|\d{1,2}-[A-Za-z]{3}-\d{2,4})\b/;
        if (dateRegex.test(strVal) && !strVal.includes('₹') && !strVal.match(/^[0-9]+$/)) {
          const parsed = parseDate(strVal, options.dateFormat);
          if (parsed && parsed !== strVal) {
            unstandardizedDates.push({ cell: addr, val: strVal, formatted: parsed });
          }
        }
      }

      // Check number candidates (e.g. ₹ 1,50,000.00, (45,000.00), 1,250.50)
      if (r > headerRowIndex && (cell.t === 's' || typeof rawVal === 'string')) {
        const hasCurrencyOrCommas = /[₹$,]|^\s*\([0-9,.]+\)\s*$/.test(strVal);
        if (hasCurrencyOrCommas && !/[a-zA-Z]{3,}/.test(strVal)) {
          const parsed = parseAmount(strVal);
          if (parsed !== 0 || strVal.includes('0')) {
            unstandardizedNumbers.push({ cell: addr, val: strVal, parsed });
          }
        }
      }
    }

    if (!rowHasData) {
      blankRowIndices.push(r);
    } else if (r > headerRowIndex) {
      const sig = rowValues.join('|~|');
      const existing = rowSignatures.get(sig) || [];
      existing.push(r);
      rowSignatures.set(sig, existing);
    }
  }

  // Duplicate rows are all rows with duplicate signatures except the first occurrence
  const duplicateRowIndices: number[] = [];
  for (const indices of rowSignatures.values()) {
    if (indices.length > 1) {
      // keep first (indices[0]), flag the rest
      for (let i = 1; i < indices.length; i++) {
        duplicateRowIndices.push(indices[i]);
      }
    }
  }
  duplicateRowIndices.sort((a, b) => a - b);

  const mergedRangesCount = Array.isArray(ws['!merges']) ? ws['!merges'].length : 0;

  return {
    sheetName,
    totalRows,
    totalCols,
    range: ref,
    headers,
    headerRowIndex,
    duplicateRowIndices,
    blankRowIndices,
    blankColIndices,
    cellsWithWhitespace,
    unstandardizedDates,
    unstandardizedNumbers,
    formulaCells,
    mergedRangesCount,
  };
}

/**
 * Safely delete a row from a worksheet, shifting cells and adjusting merges and ref.
 */
function deleteWorksheetRow(ws: XLSX.WorkSheet, rowIndex: number, range: XLSX.Range): void {
  // Move all cells below rowIndex up by 1
  for (let r = rowIndex; r < range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const curAddr = XLSX.utils.encode_cell({ r, c });
      const nextAddr = XLSX.utils.encode_cell({ r: r + 1, c });

      if (ws[nextAddr] !== undefined) {
        ws[curAddr] = ws[nextAddr];
      } else {
        delete ws[curAddr];
      }
    }
  }

  // Clear the last row
  for (let c = range.s.c; c <= range.e.c; c++) {
    const lastAddr = XLSX.utils.encode_cell({ r: range.e.r, c });
    delete ws[lastAddr];
  }

  // Adjust merged ranges
  if (Array.isArray(ws['!merges'])) {
    const newMerges: XLSX.Range[] = [];
    for (const m of ws['!merges']) {
      if (m.e.r < rowIndex) {
        // Completely above deleted row -> unchanged
        newMerges.push(m);
      } else if (m.s.r > rowIndex) {
        // Completely below deleted row -> shift up
        newMerges.push({
          s: { r: m.s.r - 1, c: m.s.c },
          e: { r: m.e.r - 1, c: m.e.c },
        });
      } else {
        // Straddles the deleted row -> shrink row count if multi-row
        if (m.e.r > m.s.r) {
          newMerges.push({
            s: { r: m.s.r, c: m.s.c },
            e: { r: m.e.r - 1, c: m.e.c },
          });
        }
      }
    }
    ws['!merges'] = newMerges;
  }

  range.e.r--;
  ws['!ref'] = XLSX.utils.encode_range(range);
}

/**
 * Execute only user-selected cleaning operations in-place on target worksheet(s).
 * Preserves all other sheets, formulas, merges, styles, and unselected data.
 */
export function applyCleaningOperations(
  wb: XLSX.WorkBook,
  targetSheetName: string | 'all',
  options: CleaningOptions
): {
  modifiedWorkbook: XLSX.WorkBook;
  report: ExecutionReport;
} {
  const clonedWb = cloneWorkbook(wb);
  const sheetsToProcess = targetSheetName === 'all'
    ? clonedWb.SheetNames
    : [targetSheetName].filter((s) => clonedWb.SheetNames.includes(s));

  const report: ExecutionReport = {
    sheetsProcessed: sheetsToProcess,
    totalDuplicatesRemoved: 0,
    totalBlankRowsRemoved: 0,
    totalBlankColsRemoved: 0,
    totalCellsTrimmed: 0,
    totalDatesStandardized: 0,
    totalNumbersStandardized: 0,
    totalGstinsStandardized: 0,
    formulasPreserved: 0,
    mergedRangesPreserved: 0,
    changes: [],
  };

  for (const sheetName of sheetsToProcess) {
    const ws = clonedWb.Sheets[sheetName];
    if (!ws || !ws['!ref']) continue;

    const analysis = analyzeWorksheet(ws, sheetName, options);
    report.formulasPreserved += analysis.formulaCells.length;
    report.mergedRangesPreserved += analysis.mergedRangesCount;

    let range = XLSX.utils.decode_range(ws['!ref']);

    // 1. Remove Duplicate Rows (descending order so row indices remain valid during deletion)
    if (options.removeDuplicates && analysis.duplicateRowIndices.length > 0) {
      const sortedDupes = [...analysis.duplicateRowIndices].sort((a, b) => b - a);
      for (const r of sortedDupes) {
        deleteWorksheetRow(ws, r, range);
        report.totalDuplicatesRemoved++;
      }
    }

    // 2. Remove Blank Rows (descending order)
    if (options.removeBlankRows) {
      // Re-scan blank rows if rows were deleted
      const updatedRange = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
      const blankRows: number[] = [];
      for (let r = updatedRange.s.r; r <= updatedRange.e.r; r++) {
        let hasData = false;
        for (let c = updatedRange.s.c; c <= updatedRange.e.c; c++) {
          const addr = XLSX.utils.encode_cell({ r, c });
          const cell = ws[addr];
          if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '') {
            hasData = true;
            break;
          }
        }
        if (!hasData) blankRows.push(r);
      }
      const sortedBlanks = blankRows.sort((a, b) => b - a);
      for (const r of sortedBlanks) {
        deleteWorksheetRow(ws, r, updatedRange);
        report.totalBlankRowsRemoved++;
      }
    }

    // 3. Cell-level operations (Trim whitespace, standardize dates, standardize numbers, clean GSTINs)
    const currentRange = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');

    for (let r = currentRange.s.r; r <= currentRange.e.r; r++) {
      for (let c = currentRange.s.c; c <= currentRange.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        if (!cell || cell.v === undefined || cell.v === null) continue;

        // PROTECT FORMULAS: Never alter cells with formulas
        if (cell.f && options.preserveFormulas) {
          continue;
        }

        const rawVal = cell.v;
        const strVal = String(rawVal);

        // A. Trim Whitespace
        if (options.trimWhitespace && (cell.t === 's' || typeof rawVal === 'string')) {
          const cleaned = strVal.replace(/[\u00A0\u200B]/g, ' ').trim();
          if (cleaned !== strVal) {
            cell.v = cleaned;
            cell.w = cleaned;
            report.totalCellsTrimmed++;
            if (report.changes.length < 50) {
              report.changes.push({
                sheet: sheetName,
                cell: addr,
                type: 'whitespace',
                oldValue: strVal,
                newValue: cleaned,
              });
            }
          }
        }

        // B. Standardize Dates
        if (options.standardizeDates && r > analysis.headerRowIndex && typeof cell.v === 'string') {
          const curVal = String(cell.v);
          const dateRegex = /\b(\d{1,4}[-/. ]\d{1,2}[-/. ]\d{1,4}|\d{1,2}-[A-Za-z]{3}-\d{2,4})\b/;
          if (dateRegex.test(curVal) && !curVal.includes('₹') && !curVal.match(/^[0-9]+$/)) {
            const parsed = parseDate(curVal, options.dateFormat);
            if (parsed && parsed !== curVal) {
              cell.v = parsed;
              cell.w = parsed;
              report.totalDatesStandardized++;
              if (report.changes.length < 50) {
                report.changes.push({
                  sheet: sheetName,
                  cell: addr,
                  type: 'date',
                  oldValue: curVal,
                  newValue: parsed,
                });
              }
            }
          }
        }

        // C. Standardize Numbers / Currency
        if (options.standardizeNumbers && r > analysis.headerRowIndex && typeof cell.v === 'string') {
          const curVal = String(cell.v);
          const hasCurrencyOrCommas = /[₹$,]|^\s*\([0-9,.]+\)\s*$/.test(curVal);
          if (hasCurrencyOrCommas && !/[a-zA-Z]{3,}/.test(curVal)) {
            const parsedNum = parseAmount(curVal);
            if (parsedNum !== 0 || curVal.includes('0')) {
              cell.t = 'n';
              cell.v = parsedNum;
              cell.z = '#,##0.00';
              delete cell.w; // let Excel render format
              report.totalNumbersStandardized++;
              if (report.changes.length < 50) {
                report.changes.push({
                  sheet: sheetName,
                  cell: addr,
                  type: 'number',
                  oldValue: curVal,
                  newValue: parsedNum,
                });
              }
            }
          }
        }

        // D. Clean GSTIN
        if (options.cleanGstin && typeof cell.v === 'string') {
          const candidate = String(cell.v).trim().toUpperCase();
          if (isValidGstin(candidate) && candidate !== cell.v) {
            cell.v = candidate;
            cell.w = candidate;
            report.totalGstinsStandardized++;
            if (report.changes.length < 50) {
              report.changes.push({
                sheet: sheetName,
                cell: addr,
                type: 'gstin',
                oldValue: rawVal,
                newValue: candidate,
              });
            }
          }
        }
      }
    }

    // Ensure column widths are adequate to prevent '###' in Excel
    ensureColumnWidths(ws);
  }

  return {
    modifiedWorkbook: clonedWb,
    report,
  };
}

/**
 * Auto-fits column widths so numbers don't show as ### when opened in Microsoft Excel.
 */
function ensureColumnWidths(ws: XLSX.WorkSheet): void {
  if (!ws['!ref']) return;
  const range = XLSX.utils.decode_range(ws['!ref']);
  const existingCols: XLSX.ColInfo[] = ws['!cols'] || [];
  const colWidths: XLSX.ColInfo[] = [];

  for (let c = range.s.c; c <= range.e.c; c++) {
    const existing = existingCols[c]?.wch || 0;
    let maxLen = existing;

    // Sample rows up to 100 rows for performance
    const sampleLimit = Math.min(range.e.r, range.s.r + 100);
    for (let r = range.s.r; r <= sampleLimit; r++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (!cell || cell.v === undefined || cell.v === null) continue;

      let len = String(cell.w || cell.v).length;
      if (cell.t === 'n') len += 3; // currency padding
      if (len > maxLen) maxLen = len;
    }

    colWidths.push({ wch: Math.max(12, Math.min(maxLen + 3, 50)) });
  }

  ws['!cols'] = colWidths;
}

/**
 * Appends an audit log sheet to the workbook documenting user-selected operations executed.
 */
export function appendAuditLogSheet(wb: XLSX.WorkBook, report: ExecutionReport, sourceFileName: string): void {
  const logRows = [
    { Parameter: 'Source Workbook', Detail: sourceFileName },
    { Parameter: 'Execution Timestamp', Detail: new Date().toLocaleString('en-IN') },
    { Parameter: 'Target Sheets Processed', Detail: report.sheetsProcessed.join(', ') },
    { Parameter: 'Duplicate Records Removed', Detail: report.totalDuplicatesRemoved },
    { Parameter: 'Blank Rows Removed', Detail: report.totalBlankRowsRemoved },
    { Parameter: 'Blank Columns Removed', Detail: report.totalBlankColsRemoved },
    { Parameter: 'Whitespace Cells Trimmed', Detail: report.totalCellsTrimmed },
    { Parameter: 'Date Cells Standardized', Detail: report.totalDatesStandardized },
    { Parameter: 'Numbers / Currency Standardized', Detail: report.totalNumbersStandardized },
    { Parameter: 'GSTINs Standardized', Detail: report.totalGstinsStandardized },
    { Parameter: 'Formula Cells Protected Intact', Detail: report.formulasPreserved },
    { Parameter: 'Merged Header Ranges Preserved', Detail: report.mergedRangesPreserved },
  ];

  const ws = XLSX.utils.json_to_sheet(logRows);
  ws['!cols'] = [{ wch: 32 }, { wch: 38 }];
  XLSX.utils.book_append_sheet(wb, ws, 'Processing_Audit_Log');
}
