import * as XLSX from 'xlsx';
import { createStyledWorksheet } from './excelStyler';
import { cloneWorkbook } from './workbookPreserver';

export type FormulaIssueType =
  | 'ERROR_REF'
  | 'ERROR_DIV_ZERO'
  | 'ERROR_VALUE'
  | 'ERROR_NAME'
  | 'ERROR_NA'
  | 'ERROR_NUM'
  | 'INCONSISTENT_FORMULA'
  | 'HARDCODED_IN_FORMULA_COL'
  | 'TEXT_STORED_AS_NUMBER';

export interface FormulaIssue {
  id: string;
  sheetName: string;
  cellAddress: string;
  issueType: FormulaIssueType;
  severity: 'Critical' | 'Warning' | 'Advisory';
  description: string;
  currentValue: any;
  currentFormula?: string;
  suggestedFix: string;
  proposedFormula?: string;
  proposedValue?: any;
  approved: boolean;
}

export interface FormulaAuditSummary {
  fileName: string;
  sheetsAudited: number;
  totalFormulasScanned: number;
  totalIssuesFound: number;
  criticalErrorsCount: number;
  inconsistenciesCount: number;
  hardcodedConstantsCount: number;
  textNumbersCount: number;
  issues: FormulaIssue[];
}

/**
 * Scans all worksheets in an Excel workbook for formula errors, broken references,
 * hardcoded constants in formula columns, and text-formatted numbers.
 */
export function auditWorkbookFormulas(wb: XLSX.WorkBook, fileName: string): FormulaAuditSummary {
  const issues: FormulaIssue[] = [];
  let totalFormulasScanned = 0;
  let criticalErrorsCount = 0;
  let inconsistenciesCount = 0;
  let hardcodedConstantsCount = 0;
  let textNumbersCount = 0;

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws || !ws['!ref']) continue;

    const range = XLSX.utils.decode_range(ws['!ref']);

    // 1. Scan column by column to detect pattern inconsistencies
    for (let c = range.s.c; c <= range.e.c; c++) {
      const colCells: { r: number; addr: string; cell: XLSX.CellObject }[] = [];
      let formulaCountInCol = 0;

      for (let r = range.s.r; r <= range.e.r; r++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        if (cell && cell.v !== undefined && cell.v !== null) {
          colCells.push({ r, addr, cell });
          if (cell.f) formulaCountInCol++;
        }
      }

      // If a column is predominantly formulas (> 60% of populated cells and at least 3 formulas)
      const isFormulaColumn = colCells.length >= 4 && formulaCountInCol >= 3 && (formulaCountInCol / colCells.length) >= 0.6;

      if (isFormulaColumn) {
        for (const item of colCells) {
          // Skip header row
          if (item.r === range.s.r) continue;

          // If this cell has NO formula while neighboring cells have formulas
          if (!item.cell.f) {
            // Find formula pattern from adjacent rows
            const prevFormula = colCells.find((ci) => ci.r === item.r - 1 && ci.cell.f)?.cell.f;
            const nextFormula = colCells.find((ci) => ci.r === item.r + 1 && ci.cell.f)?.cell.f;
            const templateFormula = prevFormula || nextFormula;

            let proposedFormula = '';
            if (templateFormula) {
              // Adjust row numbers in formula to match current row
              const targetRow = item.r + 1; // 1-indexed in Excel
              proposedFormula = templateFormula.replace(/\d+/g, String(targetRow));
            }

            hardcodedConstantsCount++;
            issues.push({
              id: `issue-${issues.length + 1}`,
              sheetName,
              cellAddress: item.addr,
              issueType: 'HARDCODED_IN_FORMULA_COL',
              severity: 'Warning',
              description: `Hardcoded constant value "${item.cell.v}" in a column where neighboring rows calculate dynamically.`,
              currentValue: item.cell.v,
              suggestedFix: proposedFormula
                ? `Replace hardcoded constant with formula: =${proposedFormula}`
                : 'Replace static value with the column formula pattern.',
              proposedFormula,
              proposedValue: item.cell.v,
              approved: false,
            });
          }
        }
      }
    }

    // 2. Cell-by-cell inspection for standard Excel formula errors
    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        if (!cell) continue;

        if (cell.f) {
          totalFormulasScanned++;
        }

        const rawVal = String(cell.v ?? '');
        const formulaStr = cell.f ? String(cell.f) : '';

        // Check standard Excel formula error codes
        if (rawVal.includes('#REF!') || formulaStr.includes('#REF!')) {
          criticalErrorsCount++;
          issues.push({
            id: `issue-${issues.length + 1}`,
            sheetName,
            cellAddress: addr,
            issueType: 'ERROR_REF',
            severity: 'Critical',
            description: 'Broken cell reference (#REF!). Refers to a deleted cell, deleted row/column, or missing sheet.',
            currentValue: rawVal,
            currentFormula: formulaStr,
            suggestedFix: 'Re-link formula to valid target cell range.',
            approved: false,
          });
        } else if (rawVal.includes('#DIV/0!') || formulaStr.includes('/0')) {
          criticalErrorsCount++;
          issues.push({
            id: `issue-${issues.length + 1}`,
            sheetName,
            cellAddress: addr,
            issueType: 'ERROR_DIV_ZERO',
            severity: 'Critical',
            description: 'Division by zero error (#DIV/0!). Calculation divisor evaluates to 0 or blank.',
            currentValue: rawVal,
            currentFormula: formulaStr,
            suggestedFix: formulaStr ? `Wrap with IFERROR: =IFERROR(${formulaStr}, 0)` : 'Provide non-zero denominator.',
            proposedFormula: formulaStr ? `IFERROR(${formulaStr}, 0)` : undefined,
            approved: false,
          });
        } else if (rawVal.includes('#VALUE!')) {
          criticalErrorsCount++;
          issues.push({
            id: `issue-${issues.length + 1}`,
            sheetName,
            cellAddress: addr,
            issueType: 'ERROR_VALUE',
            severity: 'Critical',
            description: 'Data type mismatch (#VALUE!). Mathematical operation applied to non-numeric text.',
            currentValue: rawVal,
            currentFormula: formulaStr,
            suggestedFix: 'Ensure all input arguments in formula are clean numerical values.',
            approved: false,
          });
        } else if (rawVal.includes('#NAME?')) {
          criticalErrorsCount++;
          issues.push({
            id: `issue-${issues.length + 1}`,
            sheetName,
            cellAddress: addr,
            issueType: 'ERROR_NAME',
            severity: 'Critical',
            description: 'Unrecognized formula function name (#NAME?). Formula contains a typo in function name.',
            currentValue: rawVal,
            currentFormula: formulaStr,
            suggestedFix: 'Verify spelling of Excel formula function (e.g. SUM, VLOOKUP, XLOOKUP, IF).',
            approved: false,
          });
        } else if (rawVal.includes('#N/A')) {
          inconsistenciesCount++;
          issues.push({
            id: `issue-${issues.length + 1}`,
            sheetName,
            cellAddress: addr,
            issueType: 'ERROR_NA',
            severity: 'Warning',
            description: 'Lookup match not found (#N/A). VLOOKUP, MATCH, or INDEX failed to locate key.',
            currentValue: rawVal,
            currentFormula: formulaStr,
            suggestedFix: formulaStr ? `Wrap with IFERROR: =IFERROR(${formulaStr}, "Not Found")` : 'Review lookup key.',
            proposedFormula: formulaStr ? `IFERROR(${formulaStr}, "Not Found")` : undefined,
            approved: false,
          });
        }

        // Check text stored as number in data rows below header
        if (r > range.s.r && cell.t === 's' && typeof cell.v === 'string') {
          const s = cell.v.trim();
          if (/^-?\d+(\.\d+)?$/.test(s) && !s.startsWith('0') && s.length <= 15) {
            textNumbersCount++;
            issues.push({
              id: `issue-${issues.length + 1}`,
              sheetName,
              cellAddress: addr,
              issueType: 'TEXT_STORED_AS_NUMBER',
              severity: 'Advisory',
              description: `Number stored as text ("${s}"). Excel SUM() and AVERAGE() ignore this cell silently.`,
              currentValue: s,
              suggestedFix: 'Convert string to true numeric value.',
              proposedValue: Number(s),
              approved: false,
            });
          }
        }
      }
    }
  }

  return {
    fileName,
    sheetsAudited: wb.SheetNames.length,
    totalFormulasScanned,
    totalIssuesFound: issues.length,
    criticalErrorsCount,
    inconsistenciesCount,
    hardcodedConstantsCount,
    textNumbersCount,
    issues,
  };
}

/**
 * Applies ONLY user-approved corrections to a clone of the workbook, preserving all formatting
 */
export function applyApprovedFormulaFixes(
  wb: XLSX.WorkBook,
  approvedIssues: FormulaIssue[]
): {
  correctedWorkbook: XLSX.WorkBook;
  appliedCount: number;
} {
  const correctedWb = cloneWorkbook(wb);
  let appliedCount = 0;

  for (const issue of approvedIssues) {
    if (!issue.approved) continue;

    const ws = correctedWb.Sheets[issue.sheetName];
    if (!ws) continue;

    const cell = ws[issue.cellAddress];
    if (!cell) continue;

    if (issue.proposedFormula) {
      cell.f = issue.proposedFormula;
      delete cell.w; // force recalculation in Excel
      appliedCount++;
    } else if (issue.proposedValue !== undefined) {
      cell.v = issue.proposedValue;
      cell.t = typeof issue.proposedValue === 'number' ? 'n' : 's';
      delete cell.w;
      appliedCount++;
    }
  }

  return {
    correctedWorkbook: correctedWb,
    appliedCount,
  };
}

/**
 * Creates Excel report detailing detected formula issues and corrections
 */
export function createFormulaAuditWorkbook(summary: FormulaAuditSummary): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Executive Audit Summary
  const summaryRows = [
    { Parameter: 'Workbook Audited', Value: summary.fileName },
    { Parameter: 'Total Worksheets Scanned', Value: summary.sheetsAudited },
    { Parameter: 'Total Formulas Scanned', Value: summary.totalFormulasScanned },
    { Parameter: 'Total Formula & Range Issues Detected', Value: summary.totalIssuesFound },
    { Parameter: 'Critical Formula Errors (#REF!, #DIV/0!, #VALUE!)', Value: summary.criticalErrorsCount },
    { Parameter: 'Hardcoded Constants in Formula Columns', Value: summary.hardcodedConstantsCount },
    { Parameter: 'Lookup / Range Inconsistencies (#N/A)', Value: summary.inconsistenciesCount },
    { Parameter: 'Text Numbers Skipped by Excel SUM()', Value: summary.textNumbersCount },
    { Parameter: 'Audit Audit Time', Value: new Date().toLocaleString('en-IN') },
  ];
  const wsSummary = createStyledWorksheet(summaryRows, {
    columnWidths: { Parameter: 45, Value: 30 },
  });
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Formula_Audit_Summary');

  // Sheet 2: Detected Issues Detail Log
  const issueRows = summary.issues.map((i, idx) => ({
    'Issue #': idx + 1,
    'Sheet Name': i.sheetName,
    'Cell Address': i.cellAddress,
    'Severity': i.severity,
    'Issue Classification': i.issueType,
    'Current Value / Formula': i.currentFormula ? `=${i.currentFormula}` : String(i.currentValue ?? ''),
    'Audit Finding': i.description,
    'Recommended Correction': i.suggestedFix,
    'User Status': i.approved ? 'Approved for Fix' : 'Pending Review',
  }));
  const wsIssues = createStyledWorksheet(
    issueRows.length > 0 ? issueRows : [{ Note: 'Zero formula errors or inconsistencies detected.' }],
    {
      columnWidths: {
        'Issue #': 10,
        'Sheet Name': 18,
        'Cell Address': 14,
        'Severity': 14,
        'Issue Classification': 25,
        'Current Value / Formula': 30,
        'Audit Finding': 45,
        'Recommended Correction': 45,
        'User Status': 18,
      },
    }
  );
  XLSX.utils.book_append_sheet(wb, wsIssues, 'Detected_Formula_Issues');

  return wb;
}

/**
 * Creates an in-memory sample Excel workbook with realistic corporate financial model formulas
 * and deliberate errors (#REF!, #DIV/0!, hardcoded constant in formula column, text number)
 * for instant 1-click auditing and testing.
 */
export function getSampleFormulaAuditWorkbook(): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  const data = [
    { Head: 'Cost of Raw Materials', Oct: 450000, Nov: 520000, Dec: 480000, Q3_Total: 0 },
    { Head: 'Factory Wages & Labor', Oct: 180000, Nov: 195000, Dec: 190000, Q3_Total: 0 },
    { Head: 'Power, Fuel & Utilities', Oct: 75000, Nov: 82000, Dec: 79000, Q3_Total: 0 },
    { Head: 'Freight & Cartage', Oct: 42000, Nov: 46000, Dec: 44000, Q3_Total: 0 },
    { Head: 'Plant Maintenance', Oct: 65000, Nov: 68000, Dec: 67000, Q3_Total: 0 },
  ];

  const ws = XLSX.utils.json_to_sheet(data);
  ws['!cols'] = [{ wch: 32 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 16 }];

  // Column E (Q3_Total) formulas
  ws['E2'] = { t: 'n', f: 'SUM(B2:D2)', v: 1450000 };
  ws['E3'] = { t: 'n', f: 'SUM(B3:D3)', v: 565000 };
  // Intentional trap: Hardcoded static number in a formula column!
  ws['E4'] = { t: 'n', v: 236000 };
  ws['E5'] = { t: 'n', f: 'SUM(B5:D5)', v: 132000 };
  ws['E6'] = { t: 'n', f: 'SUM(B6:D6)', v: 200000 };

  // Row 7: Grand Total formula
  ws['A7'] = { t: 's', v: 'Total Manufacturing Cost' };
  ws['E7'] = { t: 'n', f: 'SUM(E2:E6)', v: 2583000 };

  // Row 8: Broken Excel reference (#REF!)
  ws['A8'] = { t: 's', v: 'Prior Period Variance' };
  ws['E8'] = { t: 'e', f: 'E7+#REF!', v: '#REF!', w: '#REF!' };

  // Row 9: Division by zero error (#DIV/0!)
  ws['A9'] = { t: 's', v: 'Cost Variance Index' };
  ws['E9'] = { t: 'e', f: 'E7/0', v: '#DIV/0!', w: '#DIV/0!' };

  // Row 10: Number stored as text
  ws['A10'] = { t: 's', v: 'Contingency Provision' };
  ws['B10'] = { t: 's', v: '50000' };

  ws['!ref'] = 'A1:E10';

  XLSX.utils.book_append_sheet(wb, ws, 'Manufacturing_Expenses');
  return wb;
}
