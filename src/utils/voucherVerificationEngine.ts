import * as XLSX from 'xlsx';
import { parseAmount, parseDate, normalizeRefNo } from './dataSanitizer';
import { createStyledWorksheet } from './excelStyler';

export type VerificationStatus =
  | 'matched'
  | 'amount_mismatch'
  | 'date_mismatch'
  | 'missing_in_ledger'
  | 'missing_in_vouchers'
  | 'duplicate_voucher'
  | 'needs_review';

export interface VoucherEntry {
  id: string;
  sourceFile: string;
  sheetName: string;
  rowIndex: number;
  voucherNo: string;
  normalizedVoucherNo: string;
  date: string;
  accountName: string;
  debit: number;
  credit: number;
  amount: number; // Max of debit, credit
  narration: string;
  refNo: string;
  raw: Record<string, any>;
}

export interface LedgerEntry {
  id: string;
  sourceFile: string;
  sheetName: string;
  rowIndex: number;
  voucherNo: string;
  normalizedVoucherNo: string;
  date: string;
  accountName: string;
  debit: number;
  credit: number;
  amount: number;
  narration: string;
  refNo: string;
  raw: Record<string, any>;
}

export interface VerificationItem {
  id: string;
  voucherNo: string;
  status: VerificationStatus;
  severity: 'clean' | 'warning' | 'critical';
  voucherRow?: VoucherEntry;
  ledgerRow?: LedgerEntry;
  voucherAmount: number;
  ledgerAmount: number;
  varianceAmount: number;
  voucherDate: string;
  ledgerDate: string;
  dateDifferenceDays: number;
  accountName: string;
  discrepancyDetails: string;
  reviewStatus: 'unresolved' | 'approved_exception' | 'query_sent';
  auditorNotes?: string;
}

export interface VerificationSummary {
  totalVouchers: number;
  totalVouchersScrutinized?: number;
  totalLedgerLines: number;
  matchedCount: number;
  verifiedCount?: number;
  mismatchCount: number;
  missingInLedgerCount: number;
  missingInVouchersCount: number;
  duplicateCount: number;
  totalVoucherAmount: number;
  totalLedgerAmount: number;
  netVariance: number;
  reconciliationPercentage: number;
  items: VerificationItem[];
  summary?: VerificationSummary;
}

export interface VerificationMatchingConfig {
  dateToleranceDays: number; // default 3 days
  amountTolerance: number; // default 0.50
  fuzzyAccountMatch: boolean;
}

export const DEFAULT_VERIFICATION_CONFIG: VerificationMatchingConfig = {
  dateToleranceDays: 3,
  amountTolerance: 0.5,
  fuzzyAccountMatch: false,
};

export interface VoucherColumnMapping {
  voucherNoCol: string;
  dateCol: string;
  accountCol: string;
  debitCol: string;
  creditCol: string;
  narrationCol: string;
  refCol: string;
}

export interface LedgerColumnMapping {
  voucherNoCol: string;
  dateCol: string;
  accountCol: string;
  debitCol: string;
  creditCol: string;
  narrationCol: string;
  refCol: string;
}

import { detectColumnCategory } from './columnDetector';

/**
 * Auto-detects column mapping for voucher and ledger registers
 */
export function detectVerificationColumns(headers: string[]): VoucherColumnMapping {
  const mapping: VoucherColumnMapping = {
    voucherNoCol: '',
    dateCol: '',
    accountCol: '',
    debitCol: '',
    creditCol: '',
    narrationCol: '',
    refCol: '',
  };

  for (const h of headers) {
    const clean = h.toLowerCase().trim();
    const cat = detectColumnCategory(h);

    if (!mapping.voucherNoCol && (cat === 'voucherNo' || /vch.*no|voucher.*no|doc.*no|trans.*id|entry.*no/i.test(clean))) {
      mapping.voucherNoCol = h;
    } else if (!mapping.dateCol && (cat === 'date' || /date|vch.*date|posting.*date/i.test(clean))) {
      mapping.dateCol = h;
    } else if (!mapping.accountCol && (cat === 'partyName' || /account|particulars|ledger|head/i.test(clean))) {
      mapping.accountCol = h;
    } else if (!mapping.debitCol && (cat === 'debit' || /debit|dr|dr.*amt/i.test(clean))) {
      mapping.debitCol = h;
    } else if (!mapping.creditCol && (cat === 'credit' || /credit|cr|cr.*amt/i.test(clean))) {
      mapping.creditCol = h;
    } else if (!mapping.narrationCol && (cat === 'narration' || /narration|remark|memo|description/i.test(clean))) {
      mapping.narrationCol = h;
    } else if (!mapping.refCol && (cat === 'refNo' || /ref|bill.*no|chq.*no|inv.*no/i.test(clean))) {
      mapping.refCol = h;
    }
  }

  return mapping;
}

export const detectVoucherColumns = detectVerificationColumns;
export const detectLedgerColumns = detectVerificationColumns;
export const verifyVouchersAndLedger = verifyVouchersAgainstLedger;

/**
 * Cross-reconciles vouchers against general ledger records
 */
export function verifyVouchersAgainstLedger(
  voucherRows: Record<string, any>[],
  ledgerRows: Record<string, any>[],
  vMapping: VoucherColumnMapping,
  lMapping: LedgerColumnMapping,
  config?: Partial<VerificationMatchingConfig>,
  vFileName = 'Voucher_Register.xlsx',
  lFileName = 'General_Ledger.xlsx'
): VerificationSummary {
  const activeConfig: VerificationMatchingConfig = {
    ...DEFAULT_VERIFICATION_CONFIG,
    ...(config || {}),
  };
  const safeVoucherRows = Array.isArray(voucherRows) ? voucherRows : [];
  const safeLedgerRows = Array.isArray(ledgerRows) ? ledgerRows : [];

  // Parse Voucher Entries
  const vouchers: VoucherEntry[] = safeVoucherRows.map((r, idx) => {
    const voucherNo = String(r[vMapping.voucherNoCol] || '').trim();
    const date = parseDate(r[vMapping.dateCol], 'DD-MM-YYYY');
    const accountName = String(r[vMapping.accountCol] || '').trim();
    const debit = parseAmount(r[vMapping.debitCol]);
    const credit = parseAmount(r[vMapping.creditCol]);
    const amount = debit > 0 ? debit : credit;
    const narration = String(r[vMapping.narrationCol] || '').trim();
    const refNo = String(r[vMapping.refCol] || '').trim();

    return {
      id: `vch-${idx + 1}`,
      sourceFile: vFileName,
      sheetName: 'Voucher_Register',
      rowIndex: idx + 2,
      voucherNo,
      normalizedVoucherNo: normalizeRefNo(voucherNo),
      date,
      accountName,
      debit,
      credit,
      amount,
      narration,
      refNo,
      raw: r,
    };
  });

  // Parse Ledger Entries
  const ledgers: LedgerEntry[] = safeLedgerRows.map((r, idx) => {
    const voucherNo = String(r[lMapping.voucherNoCol] || '').trim();
    const date = parseDate(r[lMapping.dateCol], 'DD-MM-YYYY');
    const accountName = String(r[lMapping.accountCol] || '').trim();
    const debit = parseAmount(r[lMapping.debitCol]);
    const credit = parseAmount(r[lMapping.creditCol]);
    const amount = debit > 0 ? debit : credit;
    const narration = String(r[lMapping.narrationCol] || '').trim();
    const refNo = String(r[lMapping.refCol] || '').trim();

    return {
      id: `gl-${idx + 1}`,
      sourceFile: lFileName,
      sheetName: 'General_Ledger',
      rowIndex: idx + 2,
      voucherNo,
      normalizedVoucherNo: normalizeRefNo(voucherNo),
      date,
      accountName,
      debit,
      credit,
      amount,
      narration,
      refNo,
      raw: r,
    };
  });

  const items: VerificationItem[] = [];
  const matchedLedgerIds = new Set<string>();

  // Track duplicate vouchers in source
  const vchCounts = new Map<string, number>();
  vouchers.forEach((v) => {
    if (v.normalizedVoucherNo) {
      vchCounts.set(v.normalizedVoucherNo, (vchCounts.get(v.normalizedVoucherNo) || 0) + 1);
    }
  });

  let totalVoucherAmt = 0;
  let totalLedgerAmt = 0;

  vouchers.forEach((v) => {
    totalVoucherAmt += v.amount;
  });
  ledgers.forEach((l) => {
    totalLedgerAmt += l.amount;
  });

  // Step 1: Match Vouchers to Ledger
  vouchers.forEach((v) => {
    const isDupInVouchers = vchCounts.get(v.normalizedVoucherNo)! > 1;

    // Search candidate in Ledger
    let bestMatch: LedgerEntry | null = null;
    let matchType: VerificationStatus = 'missing_in_ledger';
    let discrepancyMsg = 'Voucher exists in daybook register but has no corresponding posting in general ledger.';

    for (const l of ledgers) {
      if (matchedLedgerIds.has(l.id)) continue;

      const normVchMatch = v.normalizedVoucherNo && l.normalizedVoucherNo && v.normalizedVoucherNo === l.normalizedVoucherNo;
      const refMatch = v.refNo && l.refNo && normalizeRefNo(v.refNo) === normalizeRefNo(l.refNo);
      const isKeyMatch = normVchMatch || refMatch;

      if (isKeyMatch) {
        bestMatch = l;
        const amtDiff = Math.abs(v.amount - l.amount);
        const isAmtEqual = amtDiff <= activeConfig.amountTolerance;

        // Date difference
        let dayDiff = 0;
        if (v.date && l.date) {
          const [d1, m1, y1] = v.date.split('-').map(Number);
          const [d2, m2, y2] = l.date.split('-').map(Number);
          const dt1 = new Date(y1, m1 - 1, d1).getTime();
          const dt2 = new Date(y2, m2 - 1, d2).getTime();
          dayDiff = Math.abs(dt1 - dt2) / (1000 * 60 * 60 * 24);
        }

        if (isAmtEqual && dayDiff <= activeConfig.dateToleranceDays) {
          matchType = 'matched';
          discrepancyMsg = 'Voucher fully reconciled with ledger posting.';
        } else if (!isAmtEqual) {
          matchType = 'amount_mismatch';
          discrepancyMsg = `Amount discrepancy: Voucher shows ₹${v.amount.toLocaleString('en-IN')}, Ledger shows ₹${l.amount.toLocaleString('en-IN')} (Variance: ₹${amtDiff.toFixed(2)}).`;
        } else {
          matchType = 'date_mismatch';
          discrepancyMsg = `Date timing difference: Voucher dated ${v.date}, Ledger posted on ${l.date} (${Math.round(dayDiff)} days apart).`;
        }
        break;
      }
    }

    if (bestMatch) {
      matchedLedgerIds.add(bestMatch.id);
    }

    if (isDupInVouchers && matchType === 'matched') {
      matchType = 'duplicate_voucher';
      discrepancyMsg = `Voucher number "${v.voucherNo}" is duplicated across multiple daybook rows.`;
    }

    const varianceAmt = bestMatch ? Math.abs(v.amount - bestMatch.amount) : v.amount;
    const severity = matchType === 'matched' ? 'clean' : matchType === 'date_mismatch' ? 'warning' : 'critical';

    items.push({
      id: `ver-${items.length + 1}`,
      voucherNo: v.voucherNo,
      status: matchType,
      severity,
      voucherRow: v,
      ledgerRow: bestMatch || undefined,
      voucherAmount: v.amount,
      ledgerAmount: bestMatch ? bestMatch.amount : 0,
      varianceAmount: varianceAmt,
      voucherDate: v.date,
      ledgerDate: bestMatch ? bestMatch.date : 'N/A',
      dateDifferenceDays: 0,
      accountName: v.accountName || bestMatch?.accountName || 'Unknown',
      discrepancyDetails: discrepancyMsg,
      reviewStatus: 'unresolved',
    });
  });

  // Step 2: Unmatched Ledger Lines (Missing in Vouchers)
  ledgers.forEach((l) => {
    if (!matchedLedgerIds.has(l.id)) {
      items.push({
        id: `ver-${items.length + 1}`,
        voucherNo: l.voucherNo || '(Missing Vch No)',
        status: 'missing_in_vouchers',
        severity: 'critical',
        ledgerRow: l,
        voucherAmount: 0,
        ledgerAmount: l.amount,
        varianceAmount: l.amount,
        voucherDate: 'N/A',
        ledgerDate: l.date,
        dateDifferenceDays: 0,
        accountName: l.accountName,
        discrepancyDetails: `Ledger entry exists (Row ${l.rowIndex}) with no supporting voucher in voucher register. Potential unrecorded or manual journal entry.`,
        reviewStatus: 'unresolved',
      });
    }
  });

  const matchedCount = items.filter((i) => i.status === 'matched').length;
  const mismatchCount = items.filter((i) => i.status === 'amount_mismatch' || i.status === 'date_mismatch').length;
  const missingInLedgerCount = items.filter((i) => i.status === 'missing_in_ledger').length;
  const missingInVouchersCount = items.filter((i) => i.status === 'missing_in_vouchers').length;
  const duplicateCount = items.filter((i) => i.status === 'duplicate_voucher').length;
  const netVariance = Math.abs(totalVoucherAmt - totalLedgerAmt);
  const reconciliationPercentage = items.length > 0 ? (matchedCount / items.length) * 100 : 0;

  const res: VerificationSummary = {
    totalVouchers: vouchers.length,
    totalVouchersScrutinized: vouchers.length,
    totalLedgerLines: ledgers.length,
    matchedCount,
    verifiedCount: matchedCount,
    mismatchCount,
    missingInLedgerCount,
    missingInVouchersCount,
    duplicateCount,
    totalVoucherAmount: totalVoucherAmt,
    totalLedgerAmount: totalLedgerAmt,
    netVariance,
    reconciliationPercentage,
    items,
  };
  res.summary = res;
  return res;
}

/**
 * Exports detailed Excel report with Matched, Exceptions, and Summary sheets
 */
export function exportVerificationWorkbook(
  summary: VerificationSummary,
  clientName = 'Audit Client',
  period = 'FY 2026-27'
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Executive Verification Summary
  const summaryRows = [
    { Parameter: 'Client Entity Name', Value: clientName },
    { Parameter: 'Audit Financial Period', Value: period },
    { Parameter: 'Total Daybook Vouchers Scanned', Value: summary.totalVouchers },
    { Parameter: 'Total General Ledger Lines Scanned', Value: summary.totalLedgerLines },
    { Parameter: 'Reconciled Matched Records', Value: summary.matchedCount },
    { Parameter: 'Reconciliation Success Rate', Value: `${summary.reconciliationPercentage.toFixed(1)}%` },
    { Parameter: 'Amount / Date Discrepancies', Value: summary.mismatchCount },
    { Parameter: 'Unposted Vouchers (Missing in Ledger)', Value: summary.missingInLedgerCount },
    { Parameter: 'Unsupported Entries (Missing in Vouchers)', Value: summary.missingInVouchersCount },
    { Parameter: 'Duplicate Voucher Numbers', Value: summary.duplicateCount },
    { Parameter: 'Total Voucher Value (₹)', Value: summary.totalVoucherAmount },
    { Parameter: 'Total Ledger Value (₹)', Value: summary.totalLedgerAmount },
    { Parameter: 'Net Reconciliation Variance (₹)', Value: summary.netVariance },
    { Parameter: 'Integrity Rule Note', Value: 'Source files preserved intact. Accounting entries are never altered automatically.' },
  ];
  const wsSummary = createStyledWorksheet(summaryRows, {
    currencyColumns: ['Value'],
    columnWidths: { Parameter: 38, Value: 32 },
  });
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Verification_Summary');

  // Sheet 2: Matched Reconciled Entries
  const matchedRows = summary.items
    .filter((i) => i.status === 'matched')
    .map((m, idx) => ({
      '#': idx + 1,
      'Voucher #': m.voucherNo,
      'Account / Particulars': m.accountName,
      'Voucher Date': m.voucherDate,
      'Ledger Date': m.ledgerDate,
      'Voucher Amount (₹)': m.voucherAmount,
      'Ledger Amount (₹)': m.ledgerAmount,
      'Voucher Source Row': m.voucherRow?.rowIndex || '-',
      'Ledger Source Row': m.ledgerRow?.rowIndex || '-',
      'Reconciliation Status': '100% MATCHED',
    }));
  const wsMatched = createStyledWorksheet(matchedRows, {
    currencyColumns: ['Voucher Amount (₹)', 'Ledger Amount (₹)'],
  });
  XLSX.utils.book_append_sheet(wb, wsMatched, 'Matched_Entries');

  // Sheet 3: Audit Exceptions & Discrepancies
  const exceptionRows = summary.items
    .filter((i) => i.status !== 'matched')
    .map((e, idx) => ({
      'Exception #': idx + 1,
      'Classification': e.status.replace(/_/g, ' ').toUpperCase(),
      'Severity': e.severity.toUpperCase(),
      'Voucher #': e.voucherNo,
      'Account Name': e.accountName,
      'Voucher Amount (₹)': e.voucherAmount,
      'Ledger Amount (₹)': e.ledgerAmount,
      'Variance (₹)': e.varianceAmount,
      'Voucher Date': e.voucherDate,
      'Ledger Date': e.ledgerDate,
      'Voucher File & Row': e.voucherRow ? `${e.voucherRow.sourceFile} (Row ${e.voucherRow.rowIndex})` : 'NOT FOUND',
      'Ledger File & Row': e.ledgerRow ? `${e.ledgerRow.sourceFile} (Row ${e.ledgerRow.rowIndex})` : 'NOT FOUND',
      'Auditor Observation': e.discrepancyDetails,
      'Review Status': e.reviewStatus.toUpperCase(),
    }));
  const wsExceptions = createStyledWorksheet(exceptionRows, {
    currencyColumns: ['Voucher Amount (₹)', 'Ledger Amount (₹)', 'Variance (₹)'],
  });
  XLSX.utils.book_append_sheet(wb, wsExceptions, 'Audit_Exceptions');

  return wb;
}

/**
 * Realistic Dual Sample Dataset (Vouchers vs Ledger)
 */
export function getSampleVerificationData(): { vouchers: Record<string, any>[]; ledgers: Record<string, any>[] } {
  const vouchers = [
    { 'Voucher No': 'VR-2026-001', 'Date': '02-09-2026', 'Account Name': 'Tata Steel Ltd', 'Debit': 450000, 'Credit': 0, 'Narration': 'Raw steel procurement', 'Ref No': 'TS/8901' },
    { 'Voucher No': 'VR-2026-002', 'Date': '04-09-2026', 'Account Name': 'Electricity Expenses', 'Debit': 85000, 'Credit': 0, 'Narration': 'Factory power bill', 'Ref No': 'MSEB-09' },
    { 'Voucher No': 'VR-2026-003', 'Date': '06-09-2026', 'Account Name': 'Reliance Retail Depot', 'Debit': 0, 'Credit': 620000, 'Narration': 'Finished goods sale', 'Ref No': 'RR/4412' },
    { 'Voucher No': 'VR-2026-004', 'Date': '08-09-2026', 'Account Name': 'Office Renovation', 'Debit': 120000, 'Credit': 0, 'Narration': 'Civil repair work', 'Ref No': 'CIV-01' }, // AMOUNT MISMATCH: Voucher shows 120,000, ledger shows 125,000!
    { 'Voucher No': 'VR-2026-005', 'Date': '10-09-2026', 'Account Name': 'Legal & Professional Fees', 'Debit': 50000, 'Credit': 0, 'Narration': 'Retainer fee', 'Ref No': 'LEG-99' }, // DATE MISMATCH: Posted 10 days later in ledger!
    { 'Voucher No': 'VR-2026-006', 'Date': '12-09-2026', 'Account Name': 'Godrej Furniture', 'Debit': 240000, 'Credit': 0, 'Narration': 'Office chairs and tables', 'Ref No': 'GF-102' }, // MISSING IN LEDGER!
    { 'Voucher No': 'VR-2026-007', 'Date': '15-09-2026', 'Account Name': 'HDFC Bank Current A/c', 'Debit': 0, 'Credit': 180000, 'Narration': 'Vendor NEFT payment', 'Ref No': 'CMS8812' },
    { 'Voucher No': 'VR-2026-008', 'Date': '18-09-2026', 'Account Name': 'Sundry Debtors Control', 'Debit': 340000, 'Credit': 0, 'Narration': 'Local trading dispatch', 'Ref No': 'SL-401' },
    { 'Voucher No': 'VR-2026-001', 'Date': '22-09-2026', 'Account Name': 'Stationery Printing', 'Debit': 4200, 'Credit': 0, 'Narration': 'DUPLICATE VOUCHER NO REUSED', 'Ref No': 'ST-01' }, // DUPLICATE VOUCHER NO
  ];

  const ledgers = [
    { 'Voucher No': 'VR-2026-001', 'Posting Date': '02-09-2026', 'Ledger Account': 'Tata Steel Ltd', 'Debit': 450000, 'Credit': 0, 'Narration': 'Raw steel procurement', 'Ref No': 'TS/8901' },
    { 'Voucher No': 'VR-2026-002', 'Posting Date': '04-09-2026', 'Ledger Account': 'Electricity Expenses', 'Debit': 85000, 'Credit': 0, 'Narration': 'Factory power bill', 'Ref No': 'MSEB-09' },
    { 'Voucher No': 'VR-2026-003', 'Posting Date': '06-09-2026', 'Ledger Account': 'Reliance Retail Depot', 'Debit': 0, 'Credit': 620000, 'Narration': 'Finished goods sale', 'Ref No': 'RR/4412' },
    { 'Voucher No': 'VR-2026-004', 'Posting Date': '08-09-2026', 'Ledger Account': 'Office Renovation', 'Debit': 125000, 'Credit': 0, 'Narration': 'Civil repair work', 'Ref No': 'CIV-01' }, // Variance of ₹5,000
    { 'Voucher No': 'VR-2026-005', 'Posting Date': '22-09-2026', 'Ledger Account': 'Legal & Professional Fees', 'Debit': 50000, 'Credit': 0, 'Narration': 'Retainer fee', 'Ref No': 'LEG-99' }, // 12 days difference
    { 'Voucher No': 'VR-2026-007', 'Posting Date': '15-09-2026', 'Ledger Account': 'HDFC Bank Current A/c', 'Debit': 0, 'Credit': 180000, 'Narration': 'Vendor NEFT payment', 'Ref No': 'CMS8812' },
    { 'Voucher No': 'VR-2026-008', 'Posting Date': '18-09-2026', 'Ledger Account': 'Sundry Debtors Control', 'Debit': 340000, 'Credit': 0, 'Narration': 'Local trading dispatch', 'Ref No': 'SL-401' },
    { 'Voucher No': 'JV-MANUAL-99', 'Posting Date': '25-09-2026', 'Ledger Account': 'Director Remuneration', 'Debit': 150000, 'Credit': 0, 'Narration': 'Manual ledger entry with NO daybook voucher!', 'Ref No': 'DIR-01' }, // MISSING IN VOUCHERS!
  ];

  return { vouchers, ledgers };
}
