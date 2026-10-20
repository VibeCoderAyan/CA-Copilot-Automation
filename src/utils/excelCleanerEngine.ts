import * as XLSX from 'xlsx';
import { parseAmount, parseDate, isValidGstin } from './dataSanitizer';
import { detectColumnCategory, type ColumnCategory } from './columnDetector';
import { createStyledWorksheet } from './excelStyler';

export interface CleanedRow {
  [key: string]: any;
  _rowId: string;
  _isDuplicate: boolean;
  _missingGstin: boolean;
  _missingAmount: boolean;
  _issues: string[];
}

export interface CleaningSummary {
  totalRows: number;
  datesStandardized: number;
  amountsStandardized: number;
  duplicatesFound: number;
  missingGstins: number;
  missingAmounts: number;
  detectedColumns: Record<string, ColumnCategory>;
}

export function cleanSpreadsheetData(rawRows: Record<string, any>[]): {
  cleanedRows: CleanedRow[];
  summary: CleaningSummary;
} {
  if (!rawRows || rawRows.length === 0) {
    return {
      cleanedRows: [],
      summary: {
        totalRows: 0,
        datesStandardized: 0,
        amountsStandardized: 0,
        duplicatesFound: 0,
        missingGstins: 0,
        missingAmounts: 0,
        detectedColumns: {},
      },
    };
  }

  const sampleRow = rawRows[0];
  const headers = Object.keys(sampleRow);
  const detectedColumns: Record<string, ColumnCategory> = {};

  for (const h of headers) {
    const cat = detectColumnCategory(h);
    if (cat) {
      detectedColumns[h] = cat;
    }
  }

  let datesStandardized = 0;
  let amountsStandardized = 0;
  let missingGstins = 0;
  let missingAmounts = 0;

  const seenSignatures = new Map<string, number>();
  const intermediateRows: CleanedRow[] = [];

  for (let idx = 0; idx < rawRows.length; idx++) {
    const raw = rawRows[idx];
    const cleaned: Record<string, any> = {};
    const issues: string[] = [];
    let hasMissingGstin = false;
    let hasMissingAmount = false;

    // Preserve all columns while standardizing detected types
    for (const [col, val] of Object.entries(raw)) {
      const cat = detectedColumns[col];

      if (cat === 'date' || cat === 'dueDate') {
        const parsed = parseDate(val, 'DD-MM-YYYY');
        if (parsed !== val) datesStandardized++;
        cleaned[col] = parsed;
      } else if (
        cat === 'taxableAmount' ||
        cat === 'cgst' ||
        cat === 'sgst' ||
        cat === 'igst' ||
        cat === 'totalAmount' ||
        cat === 'debit' ||
        cat === 'credit' ||
        cat === 'balance'
      ) {
        const num = parseAmount(val);
        cleaned[col] = num;
        if (num !== 0 || val === 0 || val === '0') {
          amountsStandardized++;
        } else if (cat === 'totalAmount' || cat === 'taxableAmount') {
          hasMissingAmount = true;
          issues.push(`Empty or invalid ${col}`);
        }
      } else if (cat === 'gstin') {
        const s = (val || '').toString().trim().toUpperCase();
        cleaned[col] = s;
        if (!s || !isValidGstin(s)) {
          hasMissingGstin = true;
          issues.push(s ? `Invalid GSTIN: ${s}` : 'Missing GSTIN');
        }
      } else {
        // String or other fields: trim cleanly
        cleaned[col] = typeof val === 'string' ? val.trim() : val;
      }
    }

    if (hasMissingGstin) missingGstins++;
    if (hasMissingAmount) missingAmounts++;

    // Generate duplicate detection signature using primary transaction columns
    const sigComponents = headers.map((h) => String(cleaned[h] ?? ''));
    const signature = sigComponents.join('|~|');

    const count = (seenSignatures.get(signature) || 0) + 1;
    seenSignatures.set(signature, count);
    const isDuplicate = count > 1;

    if (isDuplicate) {
      issues.push('Duplicate row identified');
    }

    const rowObj: CleanedRow = {
      ...cleaned,
      _rowId: `row-${idx + 1}`,
      _isDuplicate: isDuplicate,
      _missingGstin: hasMissingGstin,
      _missingAmount: hasMissingAmount,
      _issues: issues,
    };

    intermediateRows.push(rowObj);
  }

  const duplicatesFound = intermediateRows.filter((r) => r._isDuplicate).length;

  return {
    cleanedRows: intermediateRows,
    summary: {
      totalRows: rawRows.length,
      datesStandardized,
      amountsStandardized,
      duplicatesFound,
      missingGstins,
      missingAmounts,
      detectedColumns,
    },
  };
}

/**
 * Creates a complete, formatted multi-sheet Excel file with cleaned data and audit trail
 */
export function createCleanedWorkbook(
  cleanedRows: CleanedRow[],
  originalFileName: string
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Strip internal _fields for the clean export sheet
  const exportData = cleanedRows.map((row) => {
    const copy: Record<string, any> = {};
    for (const [k, v] of Object.entries(row)) {
      if (!k.startsWith('_')) {
        copy[k] = v;
      }
    }
    return copy;
  });

  const wsClean = createStyledWorksheet(exportData);
  XLSX.utils.book_append_sheet(wb, wsClean, 'Cleaned_Records');

  // Add Data Quality Audit Sheet
  const qualityReport = [
    { Parameter: 'Source File', Value: originalFileName },
    { Parameter: 'Total Rows Processed', Value: cleanedRows.length },
    { Parameter: 'Duplicate Records Flagged', Value: cleanedRows.filter((r) => r._isDuplicate).length },
    { Parameter: 'Missing/Invalid GSTIN Rows', Value: cleanedRows.filter((r) => r._missingGstin).length },
    { Parameter: 'Sanitization Status', Value: 'Validated & Standardized by CA Tools' },
    { Parameter: 'Processing Date', Value: new Date().toLocaleString('en-IN') },
  ];
  const wsAudit = createStyledWorksheet(qualityReport, {
    columnWidths: { Parameter: 30, Value: 35 },
  });
  XLSX.utils.book_append_sheet(wb, wsAudit, 'Data_Quality_Summary');

  return wb;
}
