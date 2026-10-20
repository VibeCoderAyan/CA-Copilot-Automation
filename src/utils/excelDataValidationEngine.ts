import * as XLSX from 'xlsx';
import { parseDate, parseAmount } from './dataSanitizer';
import { createStyledWorksheet } from './excelStyler';

export type QualityIssueType =
  | 'leading_trailing_whitespace'
  | 'internal_excess_whitespace'
  | 'inconsistent_date_format'
  | 'inconsistent_casing'
  | 'invalid_numeric_format'
  | 'blank_cell'
  | 'empty_row'
  | 'duplicate_row';

export interface ProposedCellChange {
  id: string;
  sheetName: string;
  rowIndex: number; // 1-indexed row number in Excel
  colIndex: number;
  colHeader: string;
  issueType: QualityIssueType;
  description: string;
  originalValue: any;
  proposedValue: any;
  approved: boolean;
}

export interface CleanupOptions {
  trimWhitespace: boolean;
  normalizeWhitespace: boolean;
  standardizeDates: boolean;
  targetDateFormat: 'DD-MM-YYYY' | 'YYYY-MM-DD' | 'DD/MM/YYYY';
  standardizeNumbers: boolean;
  normalizeCasing: boolean;
  casingStyle: 'UPPERCASE_CODES' | 'Title_Case_Names' | 'lowercase';
  flagBlankCells: boolean;
  removeEmptyRows: boolean;
  removeDuplicates: boolean;
}

export const DEFAULT_CLEANUP_OPTIONS: CleanupOptions = {
  trimWhitespace: true,
  normalizeWhitespace: true,
  standardizeDates: true,
  targetDateFormat: 'DD-MM-YYYY',
  standardizeNumbers: true,
  normalizeCasing: true,
  casingStyle: 'UPPERCASE_CODES',
  flagBlankCells: true,
  removeEmptyRows: true,
  removeDuplicates: true,
};

export interface SheetValidationSummary {
  sheetName: string;
  totalRows: number;
  totalRowsScanned?: number;
  totalCols: number;
  headers: string[];
  totalIssuesDetected: number;
  duplicateRowsCount: number;
  emptyRowsCount: number;
  whitespaceCount: number;
  dateIssuesCount: number;
  numberIssuesCount: number;
  casingIssuesCount: number;
  blankCellsCount: number;
  changes: ProposedCellChange[];
  issues?: ProposedCellChange[];
  cleanedRows: Record<string, any>[];
  summary?: SheetValidationSummary;
}

/**
 * Scans raw spreadsheet rows and detects common accounting data quality anomalies.
 * Polymorphic: accepts (sheetName, rows, options) or (rows, options).
 */
export function analyzeAndValidateSheetData(
  sheetNameOrRows: string | Record<string, any>[],
  rowsOrOptions?: Record<string, any>[] | CleanupOptions,
  optionsOrUndefined?: CleanupOptions
): SheetValidationSummary {
  let sheetName = 'Sheet1';
  let rows: Record<string, any>[] = [];
  let options: CleanupOptions = DEFAULT_CLEANUP_OPTIONS;

  if (Array.isArray(sheetNameOrRows)) {
    rows = sheetNameOrRows;
    sheetName = 'Sheet1';
    if (rowsOrOptions && !Array.isArray(rowsOrOptions)) {
      options = { ...DEFAULT_CLEANUP_OPTIONS, ...rowsOrOptions };
    }
  } else {
    sheetName = typeof sheetNameOrRows === 'string' ? sheetNameOrRows : 'Sheet1';
    rows = Array.isArray(rowsOrOptions) ? rowsOrOptions : [];
    if (optionsOrUndefined) {
      options = { ...DEFAULT_CLEANUP_OPTIONS, ...optionsOrUndefined };
    } else if (rowsOrOptions && !Array.isArray(rowsOrOptions)) {
      options = { ...DEFAULT_CLEANUP_OPTIONS, ...rowsOrOptions };
    }
  }

  if (rows.length === 0) {
    const emptyRes: SheetValidationSummary = {
      sheetName,
      totalRows: 0,
      totalRowsScanned: 0,
      totalCols: 0,
      headers: [],
      totalIssuesDetected: 0,
      duplicateRowsCount: 0,
      emptyRowsCount: 0,
      whitespaceCount: 0,
      dateIssuesCount: 0,
      numberIssuesCount: 0,
      casingIssuesCount: 0,
      blankCellsCount: 0,
      changes: [],
      issues: [],
      cleanedRows: [],
    };
    emptyRes.summary = emptyRes;
    return emptyRes;
  }

  const headers = Object.keys(rows[0] || {});
  const changes: ProposedCellChange[] = [];
  const rowHashTracker = new Map<string, number>();

  let duplicateRowsCount = 0;
  let emptyRowsCount = 0;
  let whitespaceCount = 0;
  let dateIssuesCount = 0;
  let numberIssuesCount = 0;
  let casingIssuesCount = 0;
  let blankCellsCount = 0;

  // Process rows
  rows.forEach((row, rowIdx) => {
    const excelRow = rowIdx + 2; // +1 for 0-index, +1 for header row
    const nonBlankValues = headers.map((h) => String(row[h] ?? '').trim()).filter(Boolean);

    // 1. Empty Row Check
    if (nonBlankValues.length === 0) {
      emptyRowsCount++;
      if (options.removeEmptyRows) {
        changes.push({
          id: `empty-row-${excelRow}`,
          sheetName,
          rowIndex: excelRow,
          colIndex: 0,
          colHeader: '(Entire Row)',
          issueType: 'empty_row',
          description: 'Empty row containing no data',
          originalValue: '(Empty)',
          proposedValue: '(Delete Row)',
          approved: true,
        });
      }
      return;
    }

    // 2. Duplicate Row Check
    const rowContentHash = headers.map((h) => String(row[h] ?? '').trim().toLowerCase()).join('||');
    if (rowHashTracker.has(rowContentHash)) {
      duplicateRowsCount++;
      const firstSeenRow = rowHashTracker.get(rowContentHash)!;
      if (options.removeDuplicates) {
        changes.push({
          id: `dup-row-${excelRow}`,
          sheetName,
          rowIndex: excelRow,
          colIndex: 0,
          colHeader: '(Entire Row)',
          issueType: 'duplicate_row',
          description: `Duplicate of Row ${firstSeenRow}`,
          originalValue: '(Duplicate Data)',
          proposedValue: '(Remove Duplicate)',
          approved: true,
        });
      }
    } else {
      rowHashTracker.set(rowContentHash, excelRow);
    }

    // Inspect individual cells in row
    headers.forEach((h, colIdx) => {
      const val = row[h];
      const isHeaderDate = /date|dt|period|inv.*dt/i.test(h);
      const isHeaderNum = /amount|amt|qty|rate|debit|credit|bal|tax|cgst|sgst|igst|price/i.test(h);
      const isHeaderCode = /pan|gstin|tan|hsn|sac|code|ifsc|voucher|vch/i.test(h);
      const isHeaderName = /name|party|customer|vendor|client|particulars|description/i.test(h);

      // A. Blank cell
      if (val === null || val === undefined || (typeof val === 'string' && val.trim() === '')) {
        blankCellsCount++;
        if (options.flagBlankCells) {
          changes.push({
            id: `blank-${excelRow}-${colIdx}`,
            sheetName,
            rowIndex: excelRow,
            colIndex: colIdx + 1,
            colHeader: h,
            issueType: 'blank_cell',
            description: `Blank cell in column '${h}'`,
            originalValue: '(Blank)',
            proposedValue: isHeaderNum ? 0 : '-',
            approved: false, // Default unapproved so user decides whether to fill
          });
        }
        return;
      }

      const strVal = String(val);

      // B. Leading / Trailing / Internal Multiple Whitespaces
      if (options.trimWhitespace && (strVal.startsWith(' ') || strVal.endsWith(' '))) {
        whitespaceCount++;
        const trimmed = strVal.trim();
        changes.push({
          id: `ws-${excelRow}-${colIdx}`,
          sheetName,
          rowIndex: excelRow,
          colIndex: colIdx + 1,
          colHeader: h,
          issueType: 'leading_trailing_whitespace',
          description: 'Leading or trailing whitespace detected',
          originalValue: `"${strVal}"`,
          proposedValue: `"${trimmed}"`,
          approved: true,
        });
      } else if (options.normalizeWhitespace && /\s{2,}/.test(strVal)) {
        whitespaceCount++;
        const normalized = strVal.replace(/\s+/g, ' ').trim();
        changes.push({
          id: `multi-ws-${excelRow}-${colIdx}`,
          sheetName,
          rowIndex: excelRow,
          colIndex: colIdx + 1,
          colHeader: h,
          issueType: 'internal_excess_whitespace',
          description: 'Multiple consecutive spaces detected',
          originalValue: `"${strVal}"`,
          proposedValue: `"${normalized}"`,
          approved: true,
        });
      }

      // C. Date Normalization
      if (options.standardizeDates && isHeaderDate) {
        const parsedDate = parseDate(strVal, options.targetDateFormat);
        if (parsedDate && parsedDate !== strVal && parsedDate !== '') {
          dateIssuesCount++;
          changes.push({
            id: `date-${excelRow}-${colIdx}`,
            sheetName,
            rowIndex: excelRow,
            colIndex: colIdx + 1,
            colHeader: h,
            issueType: 'inconsistent_date_format',
            description: `Inconsistent date format. Standardize to ${options.targetDateFormat}`,
            originalValue: strVal,
            proposedValue: parsedDate,
            approved: true,
          });
        }
      }

      // D. Number normalization (numeric stored as text with symbols or commas)
      if (options.standardizeNumbers && isHeaderNum && typeof val === 'string') {
        if (/₹|,|\s|\$/g.test(val)) {
          const parsed = parseAmount(val);
          numberIssuesCount++;
          changes.push({
            id: `num-${excelRow}-${colIdx}`,
            sheetName,
            rowIndex: excelRow,
            colIndex: colIdx + 1,
            colHeader: h,
            issueType: 'invalid_numeric_format',
            description: 'Numeric value formatted as text with currency symbols/commas',
            originalValue: strVal,
            proposedValue: parsed,
            approved: true,
          });
        }
      }

      // E. Casing Normalization (PAN, GSTIN, IFSC, Codes -> UPPERCASE, Names -> Title Case)
      if (options.normalizeCasing && typeof val === 'string') {
        const trimmed = strVal.trim();
        if (isHeaderCode && trimmed !== trimmed.toUpperCase()) {
          casingIssuesCount++;
          changes.push({
            id: `casing-code-${excelRow}-${colIdx}`,
            sheetName,
            rowIndex: excelRow,
            colIndex: colIdx + 1,
            colHeader: h,
            issueType: 'inconsistent_casing',
            description: `Tax / Identification code should be uppercase`,
            originalValue: trimmed,
            proposedValue: trimmed.toUpperCase(),
            approved: true,
          });
        } else if (isHeaderName && options.casingStyle === 'Title_Case_Names') {
          const titleCased = trimmed
            .toLowerCase()
            .split(' ')
            .map((w) => w.charAt(0).toUpperCase() + w.slice(1))
            .join(' ');
          if (titleCased !== trimmed && trimmed.length > 2) {
            casingIssuesCount++;
            changes.push({
              id: `casing-name-${excelRow}-${colIdx}`,
              sheetName,
              rowIndex: excelRow,
              colIndex: colIdx + 1,
              colHeader: h,
              issueType: 'inconsistent_casing',
              description: 'Party / Account name casing standardization',
              originalValue: trimmed,
              proposedValue: titleCased,
              approved: true,
            });
          }
        }
      }
    });
  });

  // Build Cleaned Rows based on initial approved changes
  const cleanedRows = applyApprovedChanges(rows, changes, headers);

  const res: SheetValidationSummary = {
    sheetName,
    totalRows: rows.length,
    totalRowsScanned: rows.length,
    totalCols: headers.length,
    headers,
    totalIssuesDetected: changes.length,
    duplicateRowsCount,
    emptyRowsCount,
    whitespaceCount,
    dateIssuesCount,
    numberIssuesCount,
    casingIssuesCount,
    blankCellsCount,
    changes,
    issues: changes,
    cleanedRows,
  };
  res.summary = res;
  return res;
}

/**
 * Applies approved changes onto source rows
 */
export function applyApprovedChanges(
  originalRows: Record<string, any>[],
  changes: ProposedCellChange[],
  headers: string[]
): Record<string, any>[] {
  const approvedMap = new Map<string, ProposedCellChange>();
  const deleteRowIndices = new Set<number>();

  changes.forEach((c) => {
    if (c.approved) {
      if (c.issueType === 'empty_row' || c.issueType === 'duplicate_row') {
        deleteRowIndices.add(c.rowIndex);
      } else {
        const key = `${c.rowIndex}::${c.colHeader}`;
        approvedMap.set(key, c);
      }
    }
  });

  const result: Record<string, any>[] = [];

  originalRows.forEach((row, rowIdx) => {
    const excelRow = rowIdx + 2;
    if (deleteRowIndices.has(excelRow)) {
      return; // Omit row
    }

    const cleanRow: Record<string, any> = {};
    headers.forEach((h) => {
      const key = `${excelRow}::${h}`;
      if (approvedMap.has(key)) {
        let proposed = approvedMap.get(key)!.proposedValue;
        if (typeof proposed === 'string' && proposed.startsWith('"') && proposed.endsWith('"')) {
          proposed = proposed.slice(1, -1);
        }
        cleanRow[h] = proposed;
      } else {
        cleanRow[h] = row[h];
      }
    });

    result.push(cleanRow);
  });

  return result;
}

/**
 * Exports cleaned Excel workbook and separate issue report
 */
export function exportCleanedWorkbook(
  cleanedRows: Record<string, any>[],
  changes: ProposedCellChange[],
  sheetName = 'Cleaned_Data'
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Cleaned Data
  const wsCleaned = createStyledWorksheet(cleanedRows);
  XLSX.utils.book_append_sheet(wb, wsCleaned, sheetName.slice(0, 31));

  // Sheet 2: Data Quality Audit Log
  const logRows = changes.map((c, idx) => ({
    'Issue #': idx + 1,
    'Worksheet': c.sheetName,
    'Row #': c.rowIndex,
    'Column': c.colHeader,
    'Issue Type': c.issueType.replace(/_/g, ' ').toUpperCase(),
    'Original Value': String(c.originalValue ?? ''),
    'Cleaned / Proposed Value': String(c.proposedValue ?? ''),
    'Change Status': c.approved ? 'APPROVED & APPLIED' : 'REJECTED (ORIGINAL PRESERVED)',
    'Audit Description': c.description,
  }));

  const wsLog = createStyledWorksheet(logRows, {
    columnWidths: {
      'Issue #': 8,
      'Worksheet': 18,
      'Row #': 8,
      'Column': 20,
      'Issue Type': 26,
      'Original Value': 26,
      'Cleaned / Proposed Value': 26,
      'Change Status': 24,
      'Audit Description': 45,
    },
  });
  XLSX.utils.book_append_sheet(wb, wsLog, 'Data_Quality_Report');

  return wb;
}

/**
 * Realistic uncleaned accounting data sample
 */
export function getSampleUncleanedRows(): Record<string, any>[] {
  return [
    {
      'Voucher No': '  VCH-1001  ', // Whitespace
      'Voucher Date': '2026/08/15', // Non-standard date format
      'Party Name': 'shree ganesh steel traders', // Lowercase
      'GSTIN / PAN': '27aabcg1234f1z5', // Lowercase GSTIN
      'Taxable Value': ' ₹ 1,45,000.00 ', // Text with currency symbol & comma
      'IGST Amount': 26100,
      'Total Amount': '1,71,100.00',
    },
    {
      'Voucher No': 'VCH-1002',
      'Voucher Date': '18-08-2026',
      'Party Name': 'Mahindra Logistics   Depot', // Internal double spaces
      'GSTIN / PAN': '27AAACM9988D1ZQ',
      'Taxable Value': 85000,
      'IGST Amount': 15300,
      'Total Amount': 100300,
    },
    {
      'Voucher No': '', // Empty Row
      'Voucher Date': '',
      'Party Name': '',
      'GSTIN / PAN': '',
      'Taxable Value': '',
      'IGST Amount': '',
      'Total Amount': '',
    },
    {
      'Voucher No': 'VCH-1003',
      'Voucher Date': '09/02/2026', // Slash format
      'Party Name': 'apex industrial spares llp',
      'GSTIN / PAN': '27AAACA4411K1ZX',
      'Taxable Value': 320000,
      'IGST Amount': null, // Blank cell
      'Total Amount': 377600,
    },
    {
      'Voucher No': 'VCH-1002', // Duplicate of Row 2
      'Voucher Date': '18-08-2026',
      'Party Name': 'Mahindra Logistics   Depot',
      'GSTIN / PAN': '27AAACM9988D1ZQ',
      'Taxable Value': 85000,
      'IGST Amount': 15300,
      'Total Amount': 100300,
    },
    {
      'Voucher No': 'VCH-1004',
      'Voucher Date': '25-08-2026',
      'Party Name': 'Reliance Retail Limited',
      'GSTIN / PAN': '27aaacr5055k1zm', // Lowercase GSTIN
      'Taxable Value': '4,50,000',
      'IGST Amount': 81000,
      'Total Amount': 531000,
    },
  ];
}
