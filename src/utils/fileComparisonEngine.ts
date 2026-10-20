import * as XLSX from 'xlsx';
import { createStyledWorksheet } from './excelStyler';

export interface FieldDifference {
  column: string;
  oldValue: any;
  newValue: any;
}

export interface ModifiedRecord {
  key: string;
  differences: FieldDifference[];
  rowA: Record<string, any>;
  rowB: Record<string, any>;
}

export interface FileComparisonResult {
  fileAName: string;
  fileBName: string;
  keyColumns: string[];
  totalRowsA: number;
  totalRowsB: number;
  addedRecords: Record<string, any>[];
  removedRecords: Record<string, any>[];
  modifiedRecords: ModifiedRecord[];
  matchedRecords: Record<string, any>[];
  columnsA: string[];
  columnsB: string[];
  commonColumns: string[];
  addedColumns: string[];
  removedColumns: string[];
}

/**
 * Compares two tabular datasets by a user-selected key column (or auto-detected key).
 */
export function compareDatasets(
  dataA: Record<string, any>[],
  dataB: Record<string, any>[],
  fileAName: string,
  fileBName: string,
  keyColumns: string[]
): FileComparisonResult {
  const columnsA = dataA.length > 0 ? Object.keys(dataA[0]).filter((k) => !k.startsWith('_')) : [];
  const columnsB = dataB.length > 0 ? Object.keys(dataB[0]).filter((k) => !k.startsWith('_')) : [];

  const commonColumns = columnsA.filter((c) => columnsB.includes(c));
  const addedColumns = columnsB.filter((c) => !columnsA.includes(c));
  const removedColumns = columnsA.filter((c) => !columnsB.includes(c));

  // Determine effective key columns
  let effectiveKeys = keyColumns.filter((k) => commonColumns.includes(k));
  if (effectiveKeys.length === 0) {
    // Auto-detect candidate key column (e.g. Voucher, Invoice, ID, Code, SNo)
    const candidate = commonColumns.find((c) =>
      /\b(voucher|vch|invoice|inv|id|code|bill|doc|sno|serial|ref)\b/i.test(c)
    );
    effectiveKeys = candidate ? [candidate] : (commonColumns.length > 0 ? [commonColumns[0]] : []);
  }

  const makeKey = (row: Record<string, any>): string => {
    if (effectiveKeys.length === 0) {
      return Object.values(row).join('|~|');
    }
    return effectiveKeys.map((k) => String(row[k] ?? '').trim().toLowerCase()).join('|~|');
  };

  const mapA = new Map<string, Record<string, any>>();
  for (const row of dataA) {
    const k = makeKey(row);
    if (!mapA.has(k)) {
      mapA.set(k, row);
    }
  }

  const mapB = new Map<string, Record<string, any>>();
  for (const row of dataB) {
    const k = makeKey(row);
    if (!mapB.has(k)) {
      mapB.set(k, row);
    }
  }

  const addedRecords: Record<string, any>[] = [];
  const removedRecords: Record<string, any>[] = [];
  const modifiedRecords: ModifiedRecord[] = [];
  const matchedRecords: Record<string, any>[] = [];

  // Check rows in B against A
  for (const [key, rowB] of mapB.entries()) {
    if (!mapA.has(key)) {
      addedRecords.push(rowB);
    } else {
      const rowA = mapA.get(key)!;
      const diffs: FieldDifference[] = [];

      for (const col of commonColumns) {
        const valA = rowA[col];
        const valB = rowB[col];

        const strA = valA !== null && valA !== undefined ? String(valA).trim() : '';
        const strB = valB !== null && valB !== undefined ? String(valB).trim() : '';

        // Number comparison with floating-point tolerance
        const numA = Number(strA.replace(/[₹,$\s]/g, ''));
        const numB = Number(strB.replace(/[₹,$\s]/g, ''));

        if (!isNaN(numA) && !isNaN(numB) && strA !== '' && strB !== '') {
          if (Math.abs(numA - numB) > 0.001) {
            diffs.push({ column: col, oldValue: valA, newValue: valB });
          }
        } else if (strA !== strB) {
          diffs.push({ column: col, oldValue: valA, newValue: valB });
        }
      }

      if (diffs.length > 0) {
        modifiedRecords.push({
          key,
          differences: diffs,
          rowA,
          rowB,
        });
      } else {
        matchedRecords.push(rowB);
      }
    }
  }

  // Check rows in A that are missing in B
  for (const [key, rowA] of mapA.entries()) {
    if (!mapB.has(key)) {
      removedRecords.push(rowA);
    }
  }

  return {
    fileAName,
    fileBName,
    keyColumns: effectiveKeys,
    totalRowsA: dataA.length,
    totalRowsB: dataB.length,
    addedRecords,
    removedRecords,
    modifiedRecords,
    matchedRecords,
    columnsA,
    columnsB,
    commonColumns,
    addedColumns,
    removedColumns,
  };
}

/**
 * Creates a professional 5-sheet comparison Excel workbook
 */
export function createComparisonWorkbook(result: FileComparisonResult): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Executive Summary
  const summaryData = [
    { Parameter: 'File A (Baseline / Reference)', Detail: result.fileAName },
    { Parameter: 'File B (Comparison / Revised)', Detail: result.fileBName },
    { Parameter: 'Matching Key Column(s)', Detail: result.keyColumns.join(', ') || 'Composite Row Signature' },
    { Parameter: 'Total Records in File A', Detail: result.totalRowsA },
    { Parameter: 'Total Records in File B', Detail: result.totalRowsB },
    { Parameter: 'Identical / Matched Records', Detail: result.matchedRecords.length },
    { Parameter: 'Modified Records (Field Variances)', Detail: result.modifiedRecords.length },
    { Parameter: 'Added Records (New in File B)', Detail: result.addedRecords.length },
    { Parameter: 'Removed Records (Missing in File B)', Detail: result.removedRecords.length },
    { Parameter: 'Total Columns in File A', Detail: result.columnsA.length },
    { Parameter: 'Total Columns in File B', Detail: result.columnsB.length },
    { Parameter: 'New Columns Added', Detail: result.addedColumns.join(', ') || 'None' },
    { Parameter: 'Columns Dropped', Detail: result.removedColumns.join(', ') || 'None' },
    { Parameter: 'Comparison Timestamp', Detail: new Date().toLocaleString('en-IN') },
  ];
  const wsSummary = createStyledWorksheet(summaryData, {
    columnWidths: { Parameter: 38, Detail: 45 },
  });
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Comparison_Summary');

  // Sheet 2: Modified Records Diff
  const diffRows: Record<string, any>[] = [];
  result.modifiedRecords.forEach((m, idx) => {
    m.differences.forEach((d) => {
      diffRows.push({
        'Diff #': idx + 1,
        'Key Value': m.key,
        'Column Name': d.column,
        'File A (Original)': d.oldValue ?? '(Empty)',
        'File B (Modified)': d.newValue ?? '(Empty)',
      });
    });
  });
  const wsModified = createStyledWorksheet(
    diffRows.length > 0 ? diffRows : [{ Note: 'Zero field-level modifications detected.' }],
    {
      columnWidths: {
        'Diff #': 10,
        'Key Value': 24,
        'Column Name': 24,
        'File A (Original)': 32,
        'File B (Modified)': 32,
      },
    }
  );
  XLSX.utils.book_append_sheet(wb, wsModified, 'Modified_Records_Diff');

  // Sheet 3: Added Records
  const wsAdded = createStyledWorksheet(
    result.addedRecords.length > 0 ? result.addedRecords : [{ Note: 'Zero records added in File B.' }]
  );
  XLSX.utils.book_append_sheet(wb, wsAdded, 'Added_Records');

  // Sheet 4: Removed Records
  const wsRemoved = createStyledWorksheet(
    result.removedRecords.length > 0 ? result.removedRecords : [{ Note: 'Zero records removed from File A.' }]
  );
  XLSX.utils.book_append_sheet(wb, wsRemoved, 'Removed_Records');

  // Sheet 5: Matching Records
  const wsMatched = createStyledWorksheet(
    result.matchedRecords.length > 0 ? result.matchedRecords : [{ Note: 'Zero exact matching records.' }]
  );
  XLSX.utils.book_append_sheet(wb, wsMatched, 'Matching_Records');

  return wb;
}

/**
 * Returns realistic sample datasets for file comparison / revision diffing
 * (includes identical records, field-level modifications, added records, and deleted records)
 */
export function getSampleComparisonDatasets(): {
  dataA: Record<string, any>[];
  dataB: Record<string, any>[];
  fileAName: string;
  fileBName: string;
  keyCol: string;
} {
  const dataA = [
    { 'Voucher No': 'SAL/001', 'Date': '01-10-2026', 'Party Name': 'Apex Logistics Ltd', 'Taxable Amount': 150000, 'Tax Rate': 18, 'Status': 'Approved' },
    { 'Voucher No': 'SAL/002', 'Date': '03-10-2026', 'Party Name': 'Bharat Pharma Supplies', 'Taxable Amount': 85000, 'Tax Rate': 12, 'Status': 'Approved' },
    { 'Voucher No': 'SAL/003', 'Date': '05-10-2026', 'Party Name': 'Delta Heavy Engineering', 'Taxable Amount': 320000, 'Tax Rate': 18, 'Status': 'Pending' },
    { 'Voucher No': 'SAL/004', 'Date': '08-10-2026', 'Party Name': 'Zenith Polymer Works', 'Taxable Amount': 45000, 'Tax Rate': 5, 'Status': 'Approved' },
    { 'Voucher No': 'SAL/005', 'Date': '12-10-2026', 'Party Name': 'Om Metal Fabricators', 'Taxable Amount': 110000, 'Tax Rate': 18, 'Status': 'Approved' },
  ];

  const dataB = [
    // Identical record
    { 'Voucher No': 'SAL/001', 'Date': '01-10-2026', 'Party Name': 'Apex Logistics Ltd', 'Taxable Amount': 150000, 'Tax Rate': 18, 'Status': 'Approved' },
    // Modified record (Amount & Status modified)
    { 'Voucher No': 'SAL/002', 'Date': '03-10-2026', 'Party Name': 'Bharat Pharma Supplies', 'Taxable Amount': 95000, 'Tax Rate': 12, 'Status': 'Audited' },
    // SAL/003 is deleted in B (removed)
    // Identical record
    { 'Voucher No': 'SAL/004', 'Date': '08-10-2026', 'Party Name': 'Zenith Polymer Works', 'Taxable Amount': 45000, 'Tax Rate': 5, 'Status': 'Approved' },
    // Modified record (Tax Rate changed)
    { 'Voucher No': 'SAL/005', 'Date': '12-10-2026', 'Party Name': 'Om Metal Fabricators', 'Taxable Amount': 110000, 'Tax Rate': 12, 'Status': 'Approved' },
    // Added in B
    { 'Voucher No': 'SAL/006', 'Date': '15-10-2026', 'Party Name': 'Shree Chamunda Steels', 'Taxable Amount': 210000, 'Tax Rate': 18, 'Status': 'New' },
  ];

  return {
    dataA,
    dataB,
    fileAName: 'Sales_Register_Baseline_Oct15.xlsx',
    fileBName: 'Sales_Register_Audited_Oct31.xlsx',
    keyCol: 'Voucher No',
  };
}
