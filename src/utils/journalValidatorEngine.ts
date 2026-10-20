import * as XLSX from 'xlsx';
import { parseAmount, parseDate } from './dataSanitizer';
import { createStyledWorksheet } from './excelStyler';

export type JournalIssueSeverity = 'critical' | 'suspicious' | 'warning' | 'info';

export type JournalIssueCategory =
  | 'voucher_imbalance'
  | 'file_imbalance'
  | 'missing_date'
  | 'duplicate_voucher_no'
  | 'blank_account'
  | 'incomplete_entry'
  | 'cash_40a3_violation'
  | 'split_cash_payment'
  | 'unusual_head_posting'
  | 'suspense_entry'
  | 'round_sum_anomaly';

export interface JournalRow {
  id: string;
  rowIndex: number;
  sheetName: string;
  voucherNo: string;
  voucherDate: string;
  accountName: string;
  accountCode: string;
  debit: number;
  credit: number;
  narration: string;
  voucherType: string;
  raw: Record<string, any>;
}

export interface JournalIssue {
  id: string;
  rowIndex: number;
  sheetName: string;
  voucherNo: string;
  category: JournalIssueCategory;
  severity: JournalIssueSeverity;
  title: string;
  description: string;
  suggestedAction: string;
  debit: number;
  credit: number;
  accountName: string;
  reviewStatus: 'pending' | 'reviewed' | 'valid_exception' | 'query_raised';
  auditorNotes?: string;
}

export interface JournalValidationSummary {
  totalVouchers: number;
  totalRows: number;
  totalDebit: number;
  totalCredit: number;
  netFileDifference: number;
  netDifference?: number;
  isFileBalanced: boolean;
  criticalCount: number;
  suspiciousCount: number;
  warningCount: number;
  infoCount: number;
  unbalancedVouchersCount: number;
  cashViolationsCount: number;
  issues: JournalIssue[];
  vouchers: {
    voucherNo: string;
    date: string;
    totalDebit: number;
    totalCredit: number;
    diff: number;
    isBalanced: boolean;
    rowCount: number;
    hasCriticalIssue: boolean;
  }[];
}

export interface JournalColumnMapping {
  voucherNoCol: string;
  dateCol: string;
  accountCol: string;
  accountCodeCol: string;
  debitCol: string;
  creditCol: string;
  narrationCol: string;
  voucherTypeCol: string;
}

import { detectColumnCategory } from './columnDetector';

/**
 * Auto-detects columns from journal entry sheets
 */
export function detectJournalColumns(headers: string[]): JournalColumnMapping {
  const mapping: JournalColumnMapping = {
    voucherNoCol: '',
    dateCol: '',
    accountCol: '',
    accountCodeCol: '',
    debitCol: '',
    creditCol: '',
    narrationCol: '',
    voucherTypeCol: '',
  };

  for (const h of headers) {
    const clean = h.toLowerCase().trim();
    const cat = detectColumnCategory(h);

    if (!mapping.voucherNoCol && (cat === 'voucherNo' || /vch.*no|voucher.*no|vouchernum|doc.*no|txn.*id|entry.*no/i.test(clean))) {
      mapping.voucherNoCol = h;
    } else if (!mapping.dateCol && (cat === 'date' || /date|vch.*date|voucher.*date|posting.*date/i.test(clean))) {
      mapping.dateCol = h;
    } else if (!mapping.accountCol && (cat === 'partyName' || /account.*name|particulars|ledger|head|account.*description/i.test(clean))) {
      mapping.accountCol = h;
    } else if (!mapping.accountCodeCol && (cat === 'accountCode' || /code|account.*code|acct.*no|gl.*code/i.test(clean))) {
      mapping.accountCodeCol = h;
    } else if (!mapping.debitCol && (cat === 'debit' || /debit|dr|dr.*amt|dr.*amount|withdrawal/i.test(clean))) {
      mapping.debitCol = h;
    } else if (!mapping.creditCol && (cat === 'credit' || /credit|cr|cr.*amt|cr.*amount|deposit/i.test(clean))) {
      mapping.creditCol = h;
    } else if (!mapping.narrationCol && (cat === 'narration' || /narration|remark|memo|note|description|particulars.*memo/i.test(clean))) {
      mapping.narrationCol = h;
    } else if (!mapping.voucherTypeCol && /voucher.*type|vch.*type|type|journal.*type/i.test(clean)) {
      mapping.voucherTypeCol = h;
    }
  }

  return mapping;
}

/**
 * Validates and audits journal entry records
 */
export function validateJournalEntries(
  rows: Record<string, any>[],
  mapping: JournalColumnMapping,
  sheetName = 'Journal_Entries'
): { summary: JournalValidationSummary; parsedRows: JournalRow[] } {
  const issues: JournalIssue[] = [];
  const parsedRows: JournalRow[] = [];

  let totalDebit = 0;
  let totalCredit = 0;

  // Group by voucher
  const voucherGroups = new Map<string, JournalRow[]>();
  const voucherDatesMap = new Map<string, Set<string>>();

  rows.forEach((r, idx) => {
    const rowIndex = idx + 2;
    const voucherNo = String(r[mapping.voucherNoCol] || '').trim();
    const rawDate = r[mapping.dateCol];
    const voucherDate = parseDate(rawDate, 'DD-MM-YYYY');
    const accountName = String(r[mapping.accountCol] || '').trim();
    const accountCode = String(r[mapping.accountCodeCol] || '').trim();
    const debit = parseAmount(r[mapping.debitCol]);
    const credit = parseAmount(r[mapping.creditCol]);
    const narration = String(r[mapping.narrationCol] || '').trim();
    const voucherType = String(r[mapping.voucherTypeCol] || '').trim();

    totalDebit += debit;
    totalCredit += credit;

    const parsed: JournalRow = {
      id: `jr-${idx + 1}`,
      rowIndex,
      sheetName,
      voucherNo: voucherNo || `(Row ${rowIndex} Blank Vch)`,
      voucherDate,
      accountName,
      accountCode,
      debit,
      credit,
      narration,
      voucherType,
      raw: r,
    };
    parsedRows.push(parsed);

    // Track voucher groups
    const vKey = voucherNo || `__BLANK_${rowIndex}__`;
    if (!voucherGroups.has(vKey)) {
      voucherGroups.set(vKey, []);
    }
    voucherGroups.get(vKey)!.push(parsed);

    if (voucherNo && voucherDate) {
      if (!voucherDatesMap.has(voucherNo)) {
        voucherDatesMap.set(voucherNo, new Set());
      }
      voucherDatesMap.get(voucherNo)!.add(voucherDate);
    }

    // 1. Missing Date
    if (!voucherDate) {
      issues.push({
        id: `iss-${issues.length + 1}`,
        rowIndex,
        sheetName,
        voucherNo: voucherNo || 'N/A',
        category: 'missing_date',
        severity: 'critical',
        title: 'Missing Transaction Date',
        description: `Row ${rowIndex} in voucher "${voucherNo || 'Blank'}" lacks a valid entry date.`,
        suggestedAction: 'Verify source physical voucher or bank statement and record posting date.',
        debit,
        credit,
        accountName,
        reviewStatus: 'pending',
      });
    }

    // 2. Blank Account Name
    if (!accountName) {
      issues.push({
        id: `iss-${issues.length + 1}`,
        rowIndex,
        sheetName,
        voucherNo: voucherNo || 'N/A',
        category: 'blank_account',
        severity: 'critical',
        title: 'Missing Account / Ledger Name',
        description: `Row ${rowIndex} has financial amounts (Dr: ₹${debit}, Cr: ₹${credit}) without an account head.`,
        suggestedAction: 'Assign specific chart of accounts ledger to prevent orphaned ledger balances.',
        debit,
        credit,
        accountName: '(Blank Account)',
        reviewStatus: 'pending',
      });
    }

    // 3. Incomplete / Nil Entry
    if (debit === 0 && credit === 0) {
      issues.push({
        id: `iss-${issues.length + 1}`,
        rowIndex,
        sheetName,
        voucherNo: voucherNo || 'N/A',
        category: 'incomplete_entry',
        severity: 'warning',
        title: 'Zero Amount Line Item',
        description: `Row ${rowIndex} has both zero debit and zero credit values.`,
        suggestedAction: 'Review whether this row was cancelled or an incomplete draft entry.',
        debit: 0,
        credit: 0,
        accountName,
        reviewStatus: 'pending',
      });
    }

    // 4. Suspense Account Scrutiny
    if (/(suspense|difference|unidentified|unknown|temp|dummy)/i.test(accountName)) {
      issues.push({
        id: `iss-${issues.length + 1}`,
        rowIndex,
        sheetName,
        voucherNo,
        category: 'suspense_entry',
        severity: 'suspicious',
        title: 'Suspense Account Posting',
        description: `Entry posted directly to "${accountName}" (Dr: ₹${debit}, Cr: ₹${credit}).`,
        suggestedAction: 'Clear suspense balances prior to audit sign-off by identifying supporting invoices.',
        debit,
        credit,
        accountName,
        reviewStatus: 'pending',
      });
    }

    // 5. Section 40A(3) Cash Payment Rule (> ₹10,000)
    const isCashAccount = /(cash|petty cash|cash in hand)/i.test(accountName);
    if (isCashAccount && credit > 10000) {
      issues.push({
        id: `iss-${issues.length + 1}`,
        rowIndex,
        sheetName,
        voucherNo,
        category: 'cash_40a3_violation',
        severity: 'critical',
        title: 'Cash Payment Exceeds Sec 40A(3) Limit (₹10,000)',
        description: `Cash disbursement of ₹${credit.toLocaleString('en-IN')} in voucher "${voucherNo}" exceeds statutory limit under Section 40A(3).`,
        suggestedAction: 'Check if covered by Rule 6DD exceptions (e.g. banking holidays, agricultural purchase), else disallow in Tax Audit Form 3CD Clause 21(d).',
        debit,
        credit,
        accountName,
        reviewStatus: 'pending',
      });
    }

    // 6. Round Sum Anomaly Check
    const maxVal = Math.max(debit, credit);
    if (maxVal >= 50000 && maxVal % 50000 === 0) {
      issues.push({
        id: `iss-${issues.length + 1}`,
        rowIndex,
        sheetName,
        voucherNo,
        category: 'round_sum_anomaly',
        severity: 'info',
        title: `Round Sum Transaction (₹${maxVal.toLocaleString('en-IN')})`,
        description: `High-value round sum entry detected. Often indicative of ad-hoc journal provisions or non-invoiced adjustments.`,
        suggestedAction: 'Examine underlying calculation sheet, board resolution, or contract.',
        debit,
        credit,
        accountName,
        reviewStatus: 'pending',
      });
    }

    // 7. Unusual Head Posting (Credit to Direct/Indirect Expense)
    if (credit > 0 && /(salary|wages|rent|depreciation|repairs|power|fuel|consultancy|legal.*fee)/i.test(accountName)) {
      issues.push({
        id: `iss-${issues.length + 1}`,
        rowIndex,
        sheetName,
        voucherNo,
        category: 'unusual_head_posting',
        severity: 'suspicious',
        title: 'Credit to Operating Expense Head',
        description: `Expense account "${accountName}" has a credit posting of ₹${credit.toLocaleString('en-IN')}.`,
        suggestedAction: 'Verify whether this is a genuine reimbursement/recovery or an erroneous expense reversal.',
        debit,
        credit,
        accountName,
        reviewStatus: 'pending',
      });
    }
  });

  // Voucher-level analysis
  const voucherSummaries: JournalValidationSummary['vouchers'] = [];
  let unbalancedVouchersCount = 0;

  voucherGroups.forEach((vRows, vKey) => {
    const isBlankVch = vKey.startsWith('__BLANK_');
    const vNo = isBlankVch ? '(Unassigned Voucher)' : vKey;
    const vDate = vRows[0]?.voucherDate || '';
    const vDebit = vRows.reduce((acc, r) => acc + r.debit, 0);
    const vCredit = vRows.reduce((acc, r) => acc + r.credit, 0);
    const diff = Math.abs(vDebit - vCredit);
    const isBalanced = diff <= 0.01;

    let hasCritical = !isBalanced;

    if (!isBalanced) {
      unbalancedVouchersCount++;
      const firstRow = vRows[0]?.rowIndex || 0;
      issues.push({
        id: `iss-${issues.length + 1}`,
        rowIndex: firstRow,
        sheetName,
        voucherNo: vNo,
        category: 'voucher_imbalance',
        severity: 'critical',
        title: `Voucher Imbalance: Difference ₹${diff.toFixed(2)}`,
        description: `Voucher "${vNo}" does not balance. Total Dr: ₹${vDebit.toFixed(2)} vs Total Cr: ₹${vCredit.toFixed(2)}.`,
        suggestedAction: 'Trace balancing leg or correct posting amounts so debit equals credit.',
        debit: vDebit,
        credit: vCredit,
        accountName: vRows.map((r) => r.accountName).filter(Boolean).join(', '),
        reviewStatus: 'pending',
      });
    }

    // Check duplicate voucher number with conflicting dates
    if (!isBlankVch) {
      const dates = voucherDatesMap.get(vKey);
      if (dates && dates.size > 1) {
        issues.push({
          id: `iss-${issues.length + 1}`,
          rowIndex: vRows[0]?.rowIndex || 0,
          sheetName,
          voucherNo: vNo,
          category: 'duplicate_voucher_no',
          severity: 'warning',
          title: 'Duplicate Voucher Number Across Multiple Dates',
          description: `Voucher number "${vNo}" is used on ${dates.size} distinct dates (${Array.from(dates).join(', ')}).`,
          suggestedAction: 'Ensure voucher numbering is unique and chronological per accounting period.',
          debit: vDebit,
          credit: vCredit,
          accountName: 'Multiple Entries',
          reviewStatus: 'pending',
        });
      }
    }

    voucherSummaries.push({
      voucherNo: vNo,
      date: vDate,
      totalDebit: vDebit,
      totalCredit: vCredit,
      diff,
      isBalanced,
      rowCount: vRows.length,
      hasCriticalIssue: hasCritical,
    });
  });

  // Overall file debit vs credit check
  const netFileDifference = Math.abs(totalDebit - totalCredit);
  const isFileBalanced = netFileDifference <= 0.05;

  if (!isFileBalanced) {
    issues.unshift({
      id: `iss-overall`,
      rowIndex: 1,
      sheetName,
      voucherNo: 'ALL',
      category: 'file_imbalance',
      severity: 'critical',
      title: `Overall Journal Register Imbalance: ₹${netFileDifference.toFixed(2)}`,
      description: `The imported file has a net imbalance. Total Debit: ₹${totalDebit.toFixed(2)} vs Total Credit: ₹${totalCredit.toFixed(2)}.`,
      suggestedAction: 'Identify unbalanced individual vouchers or missing offset accounts.',
      debit: totalDebit,
      credit: totalCredit,
      accountName: 'All Accounts',
      reviewStatus: 'pending',
    });
  }

  const criticalCount = issues.filter((i) => i.severity === 'critical').length;
  const suspiciousCount = issues.filter((i) => i.severity === 'suspicious').length;
  const warningCount = issues.filter((i) => i.severity === 'warning').length;
  const infoCount = issues.filter((i) => i.severity === 'info').length;
  const cashViolationsCount = issues.filter((i) => i.category === 'cash_40a3_violation').length;

  return {
    summary: {
      totalVouchers: voucherGroups.size,
      totalRows: rows.length,
      totalDebit,
      totalCredit,
      netFileDifference,
      netDifference: netFileDifference,
      isFileBalanced,
      criticalCount,
      suspiciousCount,
      warningCount,
      infoCount,
      unbalancedVouchersCount,
      cashViolationsCount,
      issues,
      vouchers: voucherSummaries,
    },
    parsedRows,
  };
}

/**
 * Exports detailed Excel report with Audit Findings & Approved Journal Register
 */
export function exportJournalValidationWorkbook(
  summary: JournalValidationSummary,
  parsedRows: JournalRow[],
  clientName = 'Audit Client',
  period = 'FY 2026-27'
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Audit KPI Summary
  const summaryData = [
    { Parameter: 'Client Entity', Value: clientName },
    { Parameter: 'Financial Period', Value: period },
    { Parameter: 'Total Journal Vouchers', Value: summary.totalVouchers },
    { Parameter: 'Total Line Items', Value: summary.totalRows },
    { Parameter: 'Total Debit (₹)', Value: summary.totalDebit },
    { Parameter: 'Total Credit (₹)', Value: summary.totalCredit },
    { Parameter: 'Net Difference (₹)', Value: summary.netFileDifference },
    { Parameter: 'Register Balance Status', Value: summary.isFileBalanced ? 'BALANCED' : 'IMBALANCED' },
    { Parameter: 'Unbalanced Vouchers Count', Value: summary.unbalancedVouchersCount },
    { Parameter: 'Sec 40A(3) Cash Violations (> ₹10k)', Value: summary.cashViolationsCount },
    { Parameter: 'Critical Issues', Value: summary.criticalCount },
    { Parameter: 'Suspicious Postings Under Scrutiny', Value: summary.suspiciousCount },
    { Parameter: 'Warnings & Inconsistencies', Value: summary.warningCount },
    { Parameter: 'Integrity Rule Note', Value: 'Source register untouched. Corrections require explicit approval.' },
  ];
  const wsSummary = createStyledWorksheet(summaryData, {
    currencyColumns: ['Value'],
    columnWidths: { Parameter: 36, Value: 34 },
  });
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Validation_Summary');

  // Sheet 2: All Findings & Issues
  const issueRows = summary.issues.map((iss, idx) => ({
    'Issue #': idx + 1,
    'Severity': iss.severity.toUpperCase(),
    'Category': iss.category.replace(/_/g, ' ').toUpperCase(),
    'Sheet Name': iss.sheetName,
    'Row #': iss.rowIndex,
    'Voucher #': iss.voucherNo,
    'Account / Particulars': iss.accountName,
    'Debit (₹)': iss.debit,
    'Credit (₹)': iss.credit,
    'Audit Finding': iss.title,
    'Detailed Observation': iss.description,
    'Suggested Auditor Action': iss.suggestedAction,
    'Review Status': iss.reviewStatus.toUpperCase(),
    'Auditor Notes': iss.auditorNotes || '',
  }));
  const wsIssues = createStyledWorksheet(issueRows, {
    currencyColumns: ['Debit (₹)', 'Credit (₹)'],
  });
  XLSX.utils.book_append_sheet(wb, wsIssues, 'Audit_Findings');

  // Sheet 3: Vouchers Overview
  const voucherRows = summary.vouchers.map((v) => ({
    'Voucher #': v.voucherNo,
    'Date': v.date,
    'Lines': v.rowCount,
    'Total Debit (₹)': v.totalDebit,
    'Total Credit (₹)': v.totalCredit,
    'Imbalance (₹)': v.diff,
    'Balance Status': v.isBalanced ? 'BALANCED' : 'UNBALANCED',
    'Auditor Attention': v.hasCriticalIssue ? 'ACTION REQUIRED' : 'OK',
  }));
  const wsVouchers = createStyledWorksheet(voucherRows, {
    currencyColumns: ['Total Debit (₹)', 'Total Credit (₹)', 'Imbalance (₹)'],
  });
  XLSX.utils.book_append_sheet(wb, wsVouchers, 'Vouchers_Overview');

  // Sheet 4: Raw Journal Register
  const registerRows = parsedRows.map((r) => ({
    'Row #': r.rowIndex,
    'Voucher #': r.voucherNo,
    'Date': r.voucherDate,
    'Account Name': r.accountName,
    'Account Code': r.accountCode,
    'Debit (₹)': r.debit,
    'Credit (₹)': r.credit,
    'Narration': r.narration,
    'Voucher Type': r.voucherType,
  }));
  const wsRegister = createStyledWorksheet(registerRows, {
    currencyColumns: ['Debit (₹)', 'Credit (₹)'],
  });
  XLSX.utils.book_append_sheet(wb, wsRegister, 'Journal_Register');

  return wb;
}

/**
 * Realistic sample journal dataset with balanced entries and intentional audit anomalies
 */
export function getSampleJournalRows(): Record<string, any>[] {
  return [
    // Voucher 1: Balanced Normal Depreciation Entry
    { 'Voucher No': 'JV-2026-001', 'Date': '01-09-2026', 'Account Name': 'Depreciation Expense', 'Account Code': '5102', 'Debit': 45000, 'Credit': 0, 'Narration': 'Monthly depreciation on plant', 'Voucher Type': 'Journal' },
    { 'Voucher No': 'JV-2026-001', 'Date': '01-09-2026', 'Account Name': 'Accumulated Depreciation - Plant', 'Account Code': '1202', 'Debit': 0, 'Credit': 45000, 'Narration': 'Monthly depreciation on plant', 'Voucher Type': 'Journal' },

    // Voucher 2: UNBALANCED VOUCHER (Debit != Credit)
    { 'Voucher No': 'JV-2026-002', 'Date': '03-09-2026', 'Account Name': 'Legal & Professional Fees', 'Account Code': '5201', 'Debit': 120000, 'Credit': 0, 'Narration': 'Retainer fee for high court appeal', 'Voucher Type': 'Journal' },
    { 'Voucher No': 'JV-2026-002', 'Date': '03-09-2026', 'Account Name': 'TDS Payable u/s 194J', 'Account Code': '2105', 'Debit': 0, 'Credit': 12000, 'Narration': '10% TDS deducted', 'Voucher Type': 'Journal' },
    { 'Voucher No': 'JV-2026-002', 'Date': '03-09-2026', 'Account Name': 'M/s Singhania & Partners', 'Account Code': '2201', 'Debit': 0, 'Credit': 100000, 'Narration': 'Net payable (short by 8,000!)', 'Voucher Type': 'Journal' },

    // Voucher 3: Sec 40A(3) CASH VIOLATION (> ₹10,000 in cash)
    { 'Voucher No': 'CP-2026-045', 'Date': '05-09-2026', 'Account Name': 'Office Renovation Repairs', 'Account Code': '5304', 'Debit': 28500, 'Credit': 0, 'Narration': 'Plumbing and electrical rewiring cash payment', 'Voucher Type': 'Payment' },
    { 'Voucher No': 'CP-2026-045', 'Date': '05-09-2026', 'Account Name': 'Petty Cash in Hand', 'Account Code': '1101', 'Debit': 0, 'Credit': 28500, 'Narration': 'Cash paid to contractor Sharma (disallowable u/s 40A(3))', 'Voucher Type': 'Payment' },

    // Voucher 4: MISSING DATE & BLANK ACCOUNT
    { 'Voucher No': 'JV-2026-004', 'Date': '', 'Account Name': '', 'Account Code': '9999', 'Debit': 15000, 'Credit': 0, 'Narration': 'Unidentified adjustment line', 'Voucher Type': 'Journal' },
    { 'Voucher No': 'JV-2026-004', 'Date': '', 'Account Name': 'HDFC Bank Current A/c', 'Account Code': '1105', 'Debit': 0, 'Credit': 15000, 'Narration': 'Bank deduction', 'Voucher Type': 'Payment' },

    // Voucher 5: SUSPENSE ACCOUNT ENTRY
    { 'Voucher No': 'JV-2026-005', 'Date': '10-09-2026', 'Account Name': 'Suspense Account', 'Account Code': '2999', 'Debit': 85000, 'Credit': 0, 'Narration': 'Difference in bank reconciliation parked pending confirmation', 'Voucher Type': 'Journal' },
    { 'Voucher No': 'JV-2026-005', 'Date': '10-09-2026', 'Account Name': 'ICICI Bank Operating A/c', 'Account Code': '1106', 'Debit': 0, 'Credit': 85000, 'Narration': 'Unreconciled credit', 'Voucher Type': 'Journal' },

    // Voucher 6: ROUND SUM ANOMALY (₹5,00,000 Director loan)
    { 'Voucher No': 'JV-2026-006', 'Date': '14-09-2026', 'Account Name': 'Axis Bank Escrow A/c', 'Account Code': '1107', 'Debit': 500000, 'Credit': 0, 'Narration': 'Unsecured loan received from Director', 'Voucher Type': 'Receipt' },
    { 'Voucher No': 'JV-2026-006', 'Date': '14-09-2026', 'Account Name': 'Unsecured Loan - Director Rajesh Shah', 'Account Code': '2301', 'Debit': 0, 'Credit': 500000, 'Narration': 'Subject to Sec 269SS verification', 'Voucher Type': 'Receipt' },

    // Voucher 7: UNUSUAL CREDIT TO OPERATING EXPENSE
    { 'Voucher No': 'JV-2026-007', 'Date': '18-09-2026', 'Account Name': 'Sundry Creditors Control A/c', 'Account Code': '2200', 'Debit': 34000, 'Credit': 0, 'Narration': 'Debit note recovery', 'Voucher Type': 'Journal' },
    { 'Voucher No': 'JV-2026-007', 'Date': '18-09-2026', 'Account Name': 'Staff Salary Expense', 'Account Code': '5101', 'Debit': 0, 'Credit': 34000, 'Narration': 'Crediting salary expense instead of recovery ledger', 'Voucher Type': 'Journal' },

    // Voucher 8: BALANCED SALES & GST ENTRY
    { 'Voucher No': 'SAL-2026-108', 'Date': '22-09-2026', 'Account Name': 'Reliance Retail Limited', 'Account Code': '1301', 'Debit': 236000, 'Credit': 0, 'Narration': 'Sale of polymer granules inv 108', 'Voucher Type': 'Sales' },
    { 'Voucher No': 'SAL-2026-108', 'Date': '22-09-2026', 'Account Name': 'Domestic Sales Account', 'Account Code': '4101', 'Debit': 0, 'Credit': 200000, 'Narration': 'Taxable sales value', 'Voucher Type': 'Sales' },
    { 'Voucher No': 'SAL-2026-108', 'Date': '22-09-2026', 'Account Name': 'Output CGST 9%', 'Account Code': '2101', 'Debit': 0, 'Credit': 18000, 'Narration': 'Central GST output', 'Voucher Type': 'Sales' },
    { 'Voucher No': 'SAL-2026-108', 'Date': '22-09-2026', 'Account Name': 'Output SGST 9%', 'Account Code': '2102', 'Debit': 0, 'Credit': 18000, 'Narration': 'State GST output', 'Voucher Type': 'Sales' },

    // Voucher 9: DUPLICATE VOUCHER NUMBER (Vch JV-2026-001 repeated on different date)
    { 'Voucher No': 'JV-2026-001', 'Date': '28-09-2026', 'Account Name': 'Printing & Stationery', 'Account Code': '5301', 'Debit': 4200, 'Credit': 0, 'Narration': 'Office registers and stationery', 'Voucher Type': 'Journal' },
    { 'Voucher No': 'JV-2026-001', 'Date': '28-09-2026', 'Account Name': 'Cash in Hand', 'Account Code': '1101', 'Debit': 0, 'Credit': 4200, 'Narration': 'Reused voucher number JV-2026-001', 'Voucher Type': 'Journal' },
  ];
}
