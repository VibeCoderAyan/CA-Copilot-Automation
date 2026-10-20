import * as XLSX from 'xlsx';
import { createStyledWorksheet } from './excelStyler';

export interface RawTbRow {
  rowIndex: number;
  accountCode: string;
  accountName: string;
  groupCategory: string;
  debit: number;
  credit: number;
  balance: number;
  raw: Record<string, any>;
}

export type TbIssueSeverity = 'critical' | 'warning' | 'info';

export type TbIssueCategory =
  | 'imbalance'
  | 'abnormal_balance'
  | 'duplicate_account'
  | 'missing_code'
  | 'suspense_account'
  | 'round_sum';

export interface TbIssue {
  id: string;
  rowIndex: number;
  accountCode: string;
  accountName: string;
  category: TbIssueCategory;
  severity: TbIssueSeverity;
  title: string;
  description: string;
  suggestedAction: string;
  debit: number;
  credit: number;
}

export interface TbAuditSummary {
  totalAccounts: number;
  totalDebit: number;
  totalCredit: number;
  netDifference: number;
  isBalanced: boolean;
  criticalIssuesCount: number;
  warningIssuesCount: number;
  infoIssuesCount: number;
  issues: TbIssue[];
}

export interface TbColumnMapping {
  accountNameCol: string;
  accountCodeCol: string;
  groupCol: string;
  debitCol: string;
  creditCol: string;
  balanceCol: string;
}

import { detectColumnCategory } from './columnDetector';

/**
 * Auto-detects columns in uploaded trial balance or ledger
 */
export function detectTbColumns(headers: string[]): TbColumnMapping {
  const mapping: TbColumnMapping = {
    accountCodeCol: '',
    accountNameCol: '',
    groupCol: '',
    debitCol: '',
    creditCol: '',
    balanceCol: '',
  };

  for (const h of headers) {
    const clean = h.toLowerCase().trim();
    const cat = detectColumnCategory(h);

    if (!mapping.accountNameCol && (cat === 'partyName' || /account.*name|particulars|ledger|head|description|title/i.test(clean))) {
      mapping.accountNameCol = h;
    } else if (!mapping.accountCodeCol && (cat === 'accountCode' || /code|acc.*no|ledger.*code|acct.*id/i.test(clean))) {
      mapping.accountCodeCol = h;
    } else if (!mapping.groupCol && (cat === 'head' || /group|schedule|type|classification|category|head.*group/i.test(clean))) {
      mapping.groupCol = h;
    } else if (!mapping.debitCol && (cat === 'debit' || /debit|dr|dr.*amt|dr.*amount/i.test(clean))) {
      mapping.debitCol = h;
    } else if (!mapping.creditCol && (cat === 'credit' || /credit|cr|cr.*amt|cr.*amount/i.test(clean))) {
      mapping.creditCol = h;
    } else if (!mapping.balanceCol && (cat === 'balance' || /closing|net|balance|amt|amount/i.test(clean))) {
      mapping.balanceCol = h;
    }
  }

  return mapping;
}

/**
 * Parses numerical values cleanly from messy strings
 */
function parseNum(val: any): number {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  if (!val) return 0;
  const str = String(val).replace(/,/g, '').trim();
  // Handle (100.00) accounting parenthesis for negatives
  if (/^\(.*\)$/.test(str)) {
    const num = parseFloat(str.replace(/[()]/g, ''));
    return isNaN(num) ? 0 : -num;
  }
  const n = parseFloat(str);
  return isNaN(n) ? 0 : n;
}

/**
 * Scans trial balance or ledger records for accounting errors, imbalances, and audit red flags
 */
export function auditTrialBalance(
  rows: Record<string, any>[],
  mapping: TbColumnMapping
): { summary: TbAuditSummary; parsedRows: RawTbRow[] } {
  const issues: TbIssue[] = [];
  const parsedRows: RawTbRow[] = [];
  const seenCodes = new Map<string, number>();
  const seenNames = new Map<string, number>();

  let totalDebit = 0;
  let totalCredit = 0;

  rows.forEach((row, idx) => {
    const name = String(row[mapping.accountNameCol] || '').trim();
    // Skip empty lines or header rows
    if (!name || /total|grand.*total|balance.*sheet/i.test(name)) return;

    const code = String(row[mapping.accountCodeCol] || '').trim();
    const group = String(row[mapping.groupCol] || '').trim();
    let dr = parseNum(row[mapping.debitCol]);
    let cr = parseNum(row[mapping.creditCol]);

    // If single balance column is used with positive/negative
    if (mapping.balanceCol && (!mapping.debitCol || !mapping.creditCol)) {
      const bal = parseNum(row[mapping.balanceCol]);
      if (bal > 0) dr = bal;
      else cr = Math.abs(bal);
    }

    const netBal = dr - cr;
    totalDebit += dr;
    totalCredit += cr;

    const parsed: RawTbRow = {
      rowIndex: idx + 2, // 1-indexed Excel row accounting for header
      accountCode: code,
      accountName: name,
      groupCategory: group,
      debit: dr,
      credit: cr,
      balance: netBal,
      raw: row,
    };
    parsedRows.push(parsed);

    // Rule 1: Duplicate Account Codes
    if (code) {
      const prevRow = seenCodes.get(code.toUpperCase());
      if (prevRow !== undefined) {
        issues.push({
          id: `iss-dup-code-${idx}`,
          rowIndex: idx + 2,
          accountCode: code,
          accountName: name,
          category: 'duplicate_account',
          severity: 'warning',
          title: 'Duplicate Account Code',
          description: `Account Code "${code}" is assigned to multiple ledger heads (Row ${prevRow} and Row ${idx + 2}).`,
          suggestedAction: 'Merge ledgers in ERP or reassign unique Chart of Accounts (COA) code.',
          debit: dr,
          credit: cr,
        });
      } else {
        seenCodes.set(code.toUpperCase(), idx + 2);
      }
    } else {
      // Rule 2: Missing Account Code
      issues.push({
        id: `iss-no-code-${idx}`,
        rowIndex: idx + 2,
        accountCode: '-',
        accountName: name,
        category: 'missing_code',
        severity: 'info',
        title: 'Missing Account Code',
        description: `Ledger "${name}" does not have a formal Chart of Accounts (COA) numeric code.`,
        suggestedAction: 'Assign standardized Schedule III / ERP code to ensure consistent statutory mapping.',
        debit: dr,
        credit: cr,
      });
    }

    // Rule 3: Abnormal / Inverted Balances
    const lowerName = name.toLowerCase();
    const lowerGroup = group.toLowerCase();

    // 3a. Negative Cash in Hand (Cash with Credit Balance)
    if (/cash\b|petty.*cash/i.test(lowerName) && !/bank|credit.*card/i.test(lowerName)) {
      if (cr > dr) {
        issues.push({
          id: `iss-cash-cr-${idx}`,
          rowIndex: idx + 2,
          accountCode: code,
          accountName: name,
          category: 'abnormal_balance',
          severity: 'critical',
          title: 'Negative Cash Balance (Audit Red Flag)',
          description: `Cash account shows a net Credit balance of ₹${(cr - dr).toLocaleString('en-IN')}. Physical cash can never be negative.`,
          suggestedAction: 'Investigate unrecorded cash receipts, delayed deposits, or unvouched cash disbursements immediately.',
          debit: dr,
          credit: cr,
        });
      }
    }

    // 3b. Bank Account showing Credit without OD/CC Classification
    if (/bank|hdfc|icici|sbi|axis|pnb|kotak/i.test(lowerName) && !/od|cc|overdraft|credit.*card|loan/i.test(lowerName)) {
      if (cr > dr) {
        issues.push({
          id: `iss-bank-cr-${idx}`,
          rowIndex: idx + 2,
          accountCode: code,
          accountName: name,
          category: 'abnormal_balance',
          severity: 'warning',
          title: 'Bank Account in Credit (Possible Overdraft)',
          description: `Savings/Current bank account "${name}" shows a credit balance of ₹${(cr - dr).toLocaleString('en-IN')}.`,
          suggestedAction: 'Verify whether branch has sanctioned OD/CC facility; if yes, reclassify under Current Liabilities.',
          debit: dr,
          credit: cr,
        });
      }
    }

    // 3c. Sales / Revenue with Debit Balance
    if (/sales|revenue|turnover|service.*income/i.test(lowerName) || /revenue|income/i.test(lowerGroup)) {
      if (dr > cr && dr > 1000) {
        issues.push({
          id: `iss-sales-dr-${idx}`,
          rowIndex: idx + 2,
          accountCode: code,
          accountName: name,
          category: 'abnormal_balance',
          severity: 'critical',
          title: 'Revenue Account with Net Debit Balance',
          description: `Income/Sales head "${name}" shows a net Debit balance of ₹${(dr - cr).toLocaleString('en-IN')}.`,
          suggestedAction: 'Scrutinize credit notes, sales return reversals, or incorrect debit journal vouchers.',
          debit: dr,
          credit: cr,
        });
      }
    }

    // 3d. Share Capital or Capital with Debit Balance
    if (/share.*capital|equity.*capital|partner.*capital/i.test(lowerName) || /equity/i.test(lowerGroup)) {
      if (dr > cr) {
        issues.push({
          id: `iss-equity-dr-${idx}`,
          rowIndex: idx + 2,
          accountCode: code,
          accountName: name,
          category: 'abnormal_balance',
          severity: 'critical',
          title: 'Capital Account with Debit Balance',
          description: `Equity/Capital account "${name}" has a debit balance. Capital is normally credited.`,
          suggestedAction: 'Verify whether partner drawings were booked directly into capital or capital was erroneously reduced.',
          debit: dr,
          credit: cr,
        });
      }
    }

    // Rule 4: Suspense or Clearing Accounts with Open Balance
    if (/suspense|clearing|diff.*in.*tb|unidentified/i.test(lowerName)) {
      if (Math.abs(netBal) > 0) {
        issues.push({
          id: `iss-suspense-${idx}`,
          rowIndex: idx + 2,
          accountCode: code,
          accountName: name,
          category: 'suspense_account',
          severity: 'critical',
          title: 'Open Suspense Account Balance',
          description: `Suspense / Clearing ledger has an un-cleared open balance of ₹${Math.abs(netBal).toLocaleString('en-IN')}.`,
          suggestedAction: 'Reconcile and clear all suspense items before final financial closure or audit report sign-off.',
          debit: dr,
          credit: cr,
        });
      }
    }

    // Rule 5: Significant Round-Sum Journal Entries
    const maxVal = Math.max(dr, cr);
    if (maxVal >= 500000 && maxVal % 50000 === 0 && !/capital|loan|bank/i.test(lowerName)) {
      issues.push({
        id: `iss-round-${idx}`,
        rowIndex: idx + 2,
        accountCode: code,
        accountName: name,
        category: 'round_sum',
        severity: 'info',
        title: 'Large Round-Figure Balance',
        description: `Account has an exact round figure balance of ₹${maxVal.toLocaleString('en-IN')}.`,
        suggestedAction: 'Select for audit sample testing; verify supporting third-party invoice or authorization.',
        debit: dr,
        credit: cr,
      });
    }
  });

  // Rule 6: Debit-Credit Imbalance
  const netDifference = Math.round(Math.abs(totalDebit - totalCredit) * 100) / 100;
  const isBalanced = netDifference < 0.05;

  if (!isBalanced) {
    issues.unshift({
      id: 'iss-global-imbalance',
      rowIndex: 1,
      accountCode: 'TB-ERR',
      accountName: 'Trial Balance Imbalance',
      category: 'imbalance',
      severity: 'critical',
      title: 'Debit and Credit Imbalance Detected',
      description: `Total Debits (₹${totalDebit.toLocaleString('en-IN', { minimumFractionDigits: 2 })}) do not match Total Credits (₹${totalCredit.toLocaleString('en-IN', { minimumFractionDigits: 2 })}). Exact Discrepancy: ₹${netDifference.toLocaleString('en-IN', { minimumFractionDigits: 2 })}.`,
      suggestedAction: 'Review single-sided journal entries, incomplete ERP migrations, or round-off truncation errors.',
      debit: totalDebit,
      credit: totalCredit,
    });
  }

  const criticalIssuesCount = issues.filter((i) => i.severity === 'critical').length;
  const warningIssuesCount = issues.filter((i) => i.severity === 'warning').length;
  const infoIssuesCount = issues.filter((i) => i.severity === 'info').length;

  return {
    summary: {
      totalAccounts: parsedRows.length,
      totalDebit: Math.round(totalDebit * 100) / 100,
      totalCredit: Math.round(totalCredit * 100) / 100,
      netDifference,
      isBalanced,
      criticalIssuesCount,
      warningIssuesCount,
      infoIssuesCount,
      issues,
    },
    parsedRows,
  };
}

/**
 * Returns realistic sample Trial Balance records with intentional audit test cases
 */
export function getSampleTrialBalanceRows(): Record<string, any>[] {
  return [
    { 'Account Code': '1010', 'Account Name': 'Share Capital (Equity)', 'Group': 'Equity', 'Debit': 0, 'Credit': 5000000 },
    { 'Account Code': '1020', 'Account Name': 'Secured Term Loan - HDFC Bank', 'Group': 'Liabilities', 'Debit': 0, 'Credit': 3200000 },
    { 'Account Code': '2010', 'Account Name': 'Plant & Machinery', 'Group': 'Fixed Assets', 'Debit': 4500000, 'Credit': 0 },
    { 'Account Code': '2020', 'Account Name': 'Office Furniture & Fixtures', 'Group': 'Fixed Assets', 'Debit': 850000, 'Credit': 0 },
    { 'Account Code': '3010', 'Account Name': 'Finished Goods Inventory', 'Group': 'Current Assets', 'Debit': 1850000, 'Credit': 0 },
    { 'Account Code': '3020', 'Account Name': 'Trade Receivables (Debtors)', 'Group': 'Current Assets', 'Debit': 2450000, 'Credit': 0 },
    { 'Account Code': '3030', 'Account Name': 'HDFC Current Bank Account', 'Group': 'Bank', 'Debit': 950000, 'Credit': 0 },
    { 'Account Code': '3040', 'Account Name': 'Cash on Hand (Head Office)', 'Group': 'Cash', 'Debit': 0, 'Credit': 14250 }, // Negative cash audit issue!
    { 'Account Code': '4010', 'Account Name': 'Domestic Sales - Manufacturing', 'Group': 'Revenue', 'Debit': 0, 'Credit': 8200000 },
    { 'Account Code': '4020', 'Account Name': 'Export Services Revenue', 'Group': 'Revenue', 'Debit': 45000, 'Credit': 0 }, // Revenue with debit balance!
    { 'Account Code': '5010', 'Account Name': 'Raw Material Purchases', 'Group': 'Direct Expense', 'Debit': 4200000, 'Credit': 0 },
    { 'Account Code': '5020', 'Account Name': 'Employee Salaries & Wages', 'Group': 'Indirect Expense', 'Debit': 1200000, 'Credit': 0 },
    { 'Account Code': '5030', 'Account Name': 'Factory Rent & Utilities', 'Group': 'Indirect Expense', 'Debit': 450000, 'Credit': 0 },
    { 'Account Code': '5040', 'Account Name': 'Consultancy Charges', 'Group': 'Indirect Expense', 'Debit': 500000, 'Credit': 0 }, // Round sum!
    { 'Account Code': '9999', 'Account Name': 'Suspense Clearing Account', 'Group': 'Suspense', 'Debit': 28500, 'Credit': 0 }, // Suspense balance!
    { 'Account Code': '5010', 'Account Name': 'Raw Material Purchases - Unit 2', 'Group': 'Direct Expense', 'Debit': 380000, 'Credit': 0 }, // Duplicate code 5010!
    { 'Account Code': '', 'Account Name': 'Office Printing & Stationery', 'Group': '', 'Debit': 45750, 'Credit': 0 }, // Missing code!
  ];
}

/**
 * Exports formatted 5-sheet Trial Balance audit report
 */
export function exportTbAuditWorkbook(
  summary: TbAuditSummary,
  parsedRows: RawTbRow[],
  entityName: string = 'Apex Industries Ltd',
  period: string = 'FY 2026-27'
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Audit Executive Summary
  const summaryData = [
    { Parameter: 'Entity Legal Name', Value: entityName },
    { Parameter: 'Accounting Period', Value: period },
    { Parameter: 'Total Ledgers Scanned', Value: summary.totalAccounts },
    { Parameter: 'Total Debit Amount (₹)', Value: summary.totalDebit },
    { Parameter: 'Total Credit Amount (₹)', Value: summary.totalCredit },
    { Parameter: 'Trial Balance Net Discrepancy (₹)', Value: summary.netDifference },
    { Parameter: 'Balance Status', Value: summary.isBalanced ? 'BALANCED (Zero Discrepancy)' : 'IMBALANCED' },
    { Parameter: 'Critical Severity Findings', Value: summary.criticalIssuesCount },
    { Parameter: 'Warnings & Classification Alerts', Value: summary.warningIssuesCount },
    { Parameter: 'COA / Code Recommendations', Value: summary.infoIssuesCount },
    { Parameter: 'Audit Scan Timestamp', Value: new Date().toLocaleDateString('en-IN') },
  ];

  const wsSummary = createStyledWorksheet(summaryData, {
    currencyColumns: ['Value'],
    headers: ['Parameter', 'Value'],
    columnWidths: { Parameter: 35, Value: 40 },
  });
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Audit Findings Summary');

  // Sheet 2: Standardized Trial Balance
  const tbData = parsedRows.map((r, idx) => ({
    'Sl No': idx + 1,
    'Row Ref': r.rowIndex,
    'Account Code': r.accountCode || '-',
    'Account Description': r.accountName,
    'Group / Schedule': r.groupCategory || 'General',
    'Debit (₹)': r.debit,
    'Credit (₹)': r.credit,
    'Net Balance (₹)': r.balance,
  }));

  const wsTb = createStyledWorksheet(tbData, {
    currencyColumns: ['Debit (₹)', 'Credit (₹)', 'Net Balance (₹)'],
    columnWidths: {
      'Sl No': 8,
      'Row Ref': 12,
      'Account Code': 16,
      'Account Description': 35,
      'Group / Schedule': 22,
      'Debit (₹)': 18,
      'Credit (₹)': 18,
      'Net Balance (₹)': 20,
    },
  });
  XLSX.utils.book_append_sheet(wb, wsTb, 'Standardized Trial Balance');

  // Sheet 3: Categorized Issues & Auditor Recommendations
  const issuesData = summary.issues.map((i, idx) => ({
    'Issue #': idx + 1,
    'Severity': i.severity.toUpperCase(),
    'Category': i.category.replace(/_/g, ' ').toUpperCase(),
    'Row Reference': i.rowIndex,
    'Account Code': i.accountCode,
    'Account Name': i.accountName,
    'Issue Title': i.title,
    'Detailed Finding': i.description,
    'Recommended CA Audit Action': i.suggestedAction,
  }));

  const wsIssues = createStyledWorksheet(issuesData, {
    columnWidths: {
      'Issue #': 10,
      'Severity': 14,
      'Category': 22,
      'Row Reference': 14,
      'Account Code': 16,
      'Account Name': 30,
      'Issue Title': 32,
      'Detailed Finding': 45,
      'Recommended CA Audit Action': 45,
    },
  });
  XLSX.utils.book_append_sheet(wb, wsIssues, 'Detailed Findings & Action');

  return wb;
}
