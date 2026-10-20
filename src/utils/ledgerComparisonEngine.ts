import * as XLSX from 'xlsx';
import { parseAmount, parseDate } from './dataSanitizer';
import { createStyledWorksheet } from './excelStyler';
import { detectColumnCategory } from './columnDetector';

export interface LedgerEntry {
  id: string;
  date: string;
  rawDate: string;
  voucherNo: string;
  particulars: string;
  debit: number;
  credit: number;
  narration?: string;
  [key: string]: any;
}

export interface MatchingCriteria {
  matchInverseDrCr: boolean; // Ledger A Debit matches Ledger B Credit (Inter-company / Party reconciliation)
  amountTolerance: number; // e.g. 0.00 or 1.00 or 10.00
  dateToleranceDays: number; // e.g. 0, 2, 7
  fuzzyVoucherMatch: boolean;
}

export const DEFAULT_MATCHING_CRITERIA: MatchingCriteria = {
  matchInverseDrCr: true,
  amountTolerance: 1.0,
  dateToleranceDays: 3,
  fuzzyVoucherMatch: true,
};

export interface MatchedPair {
  entryA: LedgerEntry;
  entryB: LedgerEntry;
  matchType: 'Confirmed' | 'Probable';
  confidenceScore: number;
  notes: string;
}

export interface DiscrepancyPair {
  entryA: LedgerEntry;
  entryB: LedgerEntry;
  amountDiff: number;
  reason: string;
}

export interface LedgerComparisonResult {
  ledgerAName: string;
  ledgerBName: string;
  totalDebitA: number;
  totalCreditA: number;
  netBalanceA: number;
  totalDebitB: number;
  totalCreditB: number;
  netBalanceB: number;
  balanceDifference: number;
  matchedPairs: MatchedPair[];
  probablePairs: MatchedPair[];
  discrepancies: DiscrepancyPair[];
  unmatchedA: LedgerEntry[];
  unmatchedB: LedgerEntry[];
  matchRate: number;
}

/**
 * Normalizes voucher numbers for fuzzy matching (strips special characters, prefixes like 'SAL/', 'PMT/')
 */
function normalizeVch(vch: string): string {
  return vch.replace(/[^a-zA-Z0-9]/g, '').toLowerCase().trim();
}

/**
 * Calculates day difference between two DD-MM-YYYY dates
 */
function getDayDiff(d1Str: string, d2Str: string): number {
  const p1 = d1Str.split('-');
  const p2 = d2Str.split('-');
  if (p1.length !== 3 || p2.length !== 3) return 999;
  const d1 = new Date(parseInt(p1[2]), parseInt(p1[1]) - 1, parseInt(p1[0]));
  const d2 = new Date(parseInt(p2[2]), parseInt(p2[1]) - 1, parseInt(p2[0]));
  if (isNaN(d1.getTime()) || isNaN(d2.getTime())) return 999;
  return Math.abs(Math.round((d1.getTime() - d2.getTime()) / (1000 * 60 * 60 * 24)));
}

/**
 * Parses raw spreadsheet rows into standardized LedgerEntry objects
 */
export function parseRawLedgerRows(rows: Record<string, any>[], ledgerTag: 'A' | 'B'): LedgerEntry[] {
  if (!rows || rows.length === 0) return [];
  const sample = rows[0];
  const keys = Object.keys(sample);

  const dateCol =
    keys.find((k) => detectColumnCategory(k) === 'date') ||
    keys.find((k) => /\b(date|txn dt|entry dt)\b/i.test(k)) ||
    keys[0];

  const vchCol =
    keys.find((k) => detectColumnCategory(k) === 'voucherNo' || detectColumnCategory(k) === 'refNo') ||
    keys.find((k) => /\b(voucher|vch|inv|invoice|ref|doc|bill)\b/i.test(k)) ||
    '';

  const partCol =
    keys.find((k) => detectColumnCategory(k) === 'narration' || detectColumnCategory(k) === 'partyName') ||
    keys.find((k) => /\b(particulars|party|account|name|desc|narration)\b/i.test(k)) ||
    keys[1];

  const drCol =
    keys.find((k) => detectColumnCategory(k) === 'debit') ||
    keys.find((k) => /\b(debit|dr|withdrawal|payment)\b/i.test(k)) ||
    '';

  const crCol =
    keys.find((k) => detectColumnCategory(k) === 'credit') ||
    keys.find((k) => /\b(credit|cr|deposit|receipt)\b/i.test(k)) ||
    '';

  const amtCol =
    keys.find((k) => detectColumnCategory(k) === 'totalAmount' || detectColumnCategory(k) === 'taxableAmount') ||
    keys.find((k) => /\b(amount|amt|total)\b/i.test(k)) ||
    '';

  return rows.map((r, idx) => {
    const rawD = String(r[dateCol] ?? '').trim();
    const formattedD = rawD ? parseDate(rawD, 'DD-MM-YYYY') : '';
    const vch = vchCol ? String(r[vchCol] ?? '').trim() : `VCH-${idx + 1}`;
    const particulars = partCol ? String(r[partCol] ?? '').trim() : 'Transaction';

    let debit = 0;
    let credit = 0;

    if (drCol && crCol) {
      debit = Math.abs(parseAmount(r[drCol]));
      credit = Math.abs(parseAmount(r[crCol]));
    } else if (amtCol) {
      const amt = parseAmount(r[amtCol]);
      if (amt >= 0) debit = amt;
      else credit = Math.abs(amt);
    }

    return {
      id: `led-${ledgerTag}-${idx + 1}`,
      date: formattedD || rawD,
      rawDate: rawD,
      voucherNo: vch,
      particulars,
      debit,
      credit,
      narration: r.narration || r.Narration || '',
    };
  });
}

/**
 * Reconciles two accounting ledgers with configurable matching rules
 */
export function compareLedgers(
  entriesA: LedgerEntry[],
  entriesB: LedgerEntry[],
  ledgerAName: string,
  ledgerBName: string,
  criteria: MatchingCriteria = DEFAULT_MATCHING_CRITERIA
): LedgerComparisonResult {
  const matchedPairs: MatchedPair[] = [];
  const probablePairs: MatchedPair[] = [];
  const discrepancies: DiscrepancyPair[] = [];

  const matchedBIds = new Set<string>();
  const listA = entriesA.map((a) => ({ ...a }));
  const listB = entriesB.map((b) => ({ ...b }));

  // Financial totals
  const totalDebitA = listA.reduce((sum, a) => sum + a.debit, 0);
  const totalCreditA = listA.reduce((sum, a) => sum + a.credit, 0);
  const netBalanceA = totalDebitA - totalCreditA;

  const totalDebitB = listB.reduce((sum, b) => sum + b.debit, 0);
  const totalCreditB = listB.reduce((sum, b) => sum + b.credit, 0);
  const netBalanceB = totalDebitB - totalCreditB;

  // In reciprocal inter-company / party reconciliation, Party A Dr matches Party B Cr (i.e. NetA + NetB = 0)
  const balanceDifference = criteria.matchInverseDrCr
    ? Math.abs(netBalanceA + netBalanceB)
    : Math.abs(netBalanceA - netBalanceB);

  // Helper to determine target amount in B for comparison
  const getExpectedAmount = (a: LedgerEntry): { isDebit: boolean; amount: number } => {
    if (a.debit > 0) {
      return { isDebit: criteria.matchInverseDrCr ? false : true, amount: a.debit };
    }
    return { isDebit: criteria.matchInverseDrCr ? true : false, amount: a.credit };
  };

  const getBAmount = (b: LedgerEntry, lookForDebit: boolean): number => {
    return lookForDebit ? b.debit : b.credit;
  };

  const matchedAIds = new Set<string>();

  // Pass 1: Confirmed Exact Matches (Exact voucher + Exact amount within tolerance)
  for (const a of listA) {
    if (matchedAIds.has(a.id)) continue;
    const { isDebit, amount } = getExpectedAmount(a);
    const normVchA = normalizeVch(a.voucherNo);

    for (const b of listB) {
      if (matchedBIds.has(b.id)) continue;
      const bAmt = getBAmount(b, isDebit);
      const normVchB = normalizeVch(b.voucherNo);

      const voucherMatches = normVchA && normVchB && normVchA === normVchB;
      const amountMatches = Math.abs(amount - bAmt) <= criteria.amountTolerance;

      if (voucherMatches && amountMatches) {
        matchedPairs.push({
          entryA: a,
          entryB: b,
          matchType: 'Confirmed',
          confidenceScore: 100,
          notes: `Exact Voucher Match (${a.voucherNo}) & Amount (₹${amount.toLocaleString('en-IN')})`,
        });
        matchedAIds.add(a.id);
        matchedBIds.add(b.id);
        break;
      }
    }
  }

  // Pass 2: Amount Discrepancies (Same voucher number, but amounts differ!)
  for (const a of listA) {
    if (matchedAIds.has(a.id)) continue;
    const { isDebit, amount } = getExpectedAmount(a);
    const normVchA = normalizeVch(a.voucherNo);

    if (!normVchA || normVchA.length < 3) continue;

    for (const b of listB) {
      if (matchedBIds.has(b.id)) continue;
      const bAmt = getBAmount(b, isDebit);
      const normVchB = normalizeVch(b.voucherNo);

      if (normVchA === normVchB) {
        const diff = Math.abs(amount - bAmt);
        if (diff > criteria.amountTolerance) {
          discrepancies.push({
            entryA: a,
            entryB: b,
            amountDiff: diff,
            reason: `Voucher "${a.voucherNo}" booked as ₹${amount.toLocaleString('en-IN')} in ${ledgerAName} vs ₹${bAmt.toLocaleString('en-IN')} in ${ledgerBName}`,
          });
          matchedAIds.add(a.id);
          matchedBIds.add(b.id);
          break;
        }
      }
    }
  }

  // Pass 3: Probable Matches (Matching amount + date within tolerance + party/narration heuristic)
  for (const a of listA) {
    if (matchedAIds.has(a.id)) continue;
    const { isDebit, amount } = getExpectedAmount(a);

    for (const b of listB) {
      if (matchedBIds.has(b.id)) continue;
      const bAmt = getBAmount(b, isDebit);
      const amountMatches = Math.abs(amount - bAmt) <= criteria.amountTolerance;
      const dayDiff = getDayDiff(a.date, b.date);

      if (amountMatches && dayDiff <= criteria.dateToleranceDays) {
        probablePairs.push({
          entryA: a,
          entryB: b,
          matchType: 'Probable',
          confidenceScore: Math.max(70, 95 - dayDiff * 5),
          notes: `Amount match (₹${amount.toLocaleString('en-IN')}) with date variance of ${dayDiff} day(s)`,
        });
        matchedAIds.add(a.id);
        matchedBIds.add(b.id);
        break;
      }
    }
  }

  // Unmatched in A and B
  const unmatchedA = listA.filter((a) => !matchedAIds.has(a.id));
  const unmatchedB = listB.filter((b) => !matchedBIds.has(b.id));

  const totalPossible = Math.max(listA.length, listB.length);
  const matchRate = totalPossible > 0 ? Math.round(((matchedPairs.length + probablePairs.length) / totalPossible) * 100) : 0;

  return {
    ledgerAName,
    ledgerBName,
    totalDebitA,
    totalCreditA,
    netBalanceA,
    totalDebitB,
    totalCreditB,
    netBalanceB,
    balanceDifference,
    matchedPairs,
    probablePairs,
    discrepancies,
    unmatchedA,
    unmatchedB,
    matchRate,
  };
}

/**
 * Creates a professional 6-sheet Ledger Reconciliation Excel workbook
 */
export function createLedgerReconciliationWorkbook(result: LedgerComparisonResult): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Reconciliation Summary
  const summaryRows = [
    { Parameter: 'Ledger A Source', Detail: result.ledgerAName },
    { Parameter: 'Ledger B Source', Detail: result.ledgerBName },
    { Parameter: 'Total Debit in Ledger A (₹)', Detail: result.totalDebitA },
    { Parameter: 'Total Credit in Ledger A (₹)', Detail: result.totalCreditA },
    { Parameter: 'Net Closing Balance Ledger A (₹)', Detail: result.netBalanceA },
    { Parameter: 'Total Debit in Ledger B (₹)', Detail: result.totalDebitB },
    { Parameter: 'Total Credit in Ledger B (₹)', Detail: result.totalCreditB },
    { Parameter: 'Net Closing Balance Ledger B (₹)', Detail: result.netBalanceB },
    { Parameter: 'Net Balance Discrepancy (₹)', Detail: result.balanceDifference },
    { Parameter: 'Reconciliation Match Rate', Detail: `${result.matchRate}%` },
    { Parameter: 'Confirmed Matches Count', Detail: result.matchedPairs.length },
    { Parameter: 'Probable Matches Count', Detail: result.probablePairs.length },
    { Parameter: 'Amount Discrepancy Vouchers', Detail: result.discrepancies.length },
    { Parameter: 'Unmatched Entries in Ledger A', Detail: result.unmatchedA.length },
    { Parameter: 'Unmatched Entries in Ledger B', Detail: result.unmatchedB.length },
  ];
  const wsSummary = createStyledWorksheet(summaryRows, {
    currencyColumns: ['Detail'],
    columnWidths: { Parameter: 40, Detail: 35 },
  });
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Reconciliation_Summary');

  // Sheet 2: Confirmed Matches
  const confirmedRows = result.matchedPairs.map((p, idx) => ({
    'Match #': idx + 1,
    'Ledger A Date': p.entryA.date,
    'Ledger A Voucher': p.entryA.voucherNo,
    'Ledger A Particulars': p.entryA.particulars,
    'Ledger A Amount (₹)': p.entryA.debit || p.entryA.credit,
    'Ledger B Date': p.entryB.date,
    'Ledger B Voucher': p.entryB.voucherNo,
    'Ledger B Particulars': p.entryB.particulars,
    'Ledger B Amount (₹)': p.entryB.debit || p.entryB.credit,
    'Match Basis': p.notes,
  }));
  const wsConfirmed = createStyledWorksheet(
    confirmedRows.length > 0 ? confirmedRows : [{ Note: 'Zero confirmed exact matches.' }],
    {
      currencyColumns: ['Ledger A Amount (₹)', 'Ledger B Amount (₹)'],
      dateColumns: ['Ledger A Date', 'Ledger B Date'],
    }
  );
  XLSX.utils.book_append_sheet(wb, wsConfirmed, 'Confirmed_Matches');

  // Sheet 3: Probable Matches
  const probableRows = result.probablePairs.map((p, idx) => ({
    'Probable #': idx + 1,
    'Confidence': `${p.confidenceScore}%`,
    'Ledger A Date': p.entryA.date,
    'Ledger A Voucher': p.entryA.voucherNo,
    'Ledger A Particulars': p.entryA.particulars,
    'Ledger A Amount (₹)': p.entryA.debit || p.entryA.credit,
    'Ledger B Date': p.entryB.date,
    'Ledger B Voucher': p.entryB.voucherNo,
    'Ledger B Particulars': p.entryB.particulars,
    'Ledger B Amount (₹)': p.entryB.debit || p.entryB.credit,
    'Observation': p.notes,
  }));
  const wsProbable = createStyledWorksheet(
    probableRows.length > 0 ? probableRows : [{ Note: 'Zero probable matches.' }],
    {
      currencyColumns: ['Ledger A Amount (₹)', 'Ledger B Amount (₹)'],
      dateColumns: ['Ledger A Date', 'Ledger B Date'],
    }
  );
  XLSX.utils.book_append_sheet(wb, wsProbable, 'Probable_Matches');

  // Sheet 4: Amount Discrepancies
  const discRows = result.discrepancies.map((d, idx) => ({
    'Disc #': idx + 1,
    'Voucher No': d.entryA.voucherNo,
    'Ledger A Amount (₹)': d.entryA.debit || d.entryA.credit,
    'Ledger B Amount (₹)': d.entryB.debit || d.entryB.credit,
    'Variance (₹)': d.amountDiff,
    'Audit Observation': d.reason,
  }));
  const wsDisc = createStyledWorksheet(
    discRows.length > 0 ? discRows : [{ Note: 'Zero voucher amount discrepancies.' }],
    {
      currencyColumns: ['Ledger A Amount (₹)', 'Ledger B Amount (₹)', 'Variance (₹)'],
    }
  );
  XLSX.utils.book_append_sheet(wb, wsDisc, 'Amount_Discrepancies');

  // Sheet 5: Unmatched Ledger A
  const unrecARows = result.unmatchedA.map((a, idx) => ({
    'S.No': idx + 1,
    'Date': a.date,
    'Voucher No': a.voucherNo,
    'Particulars': a.particulars,
    'Debit (₹)': a.debit,
    'Credit (₹)': a.credit,
    'Narration': a.narration || '',
    'Status': 'Missing in Ledger B',
  }));
  const wsUnrecA = createStyledWorksheet(
    unrecARows.length > 0 ? unrecARows : [{ Note: 'All Ledger A transactions reconciled.' }],
    {
      currencyColumns: ['Debit (₹)', 'Credit (₹)'],
      dateColumns: ['Date'],
    }
  );
  XLSX.utils.book_append_sheet(wb, wsUnrecA, 'Unmatched_Ledger_A');

  // Sheet 6: Unmatched Ledger B
  const unrecBRows = result.unmatchedB.map((b, idx) => ({
    'S.No': idx + 1,
    'Date': b.date,
    'Voucher No': b.voucherNo,
    'Particulars': b.particulars,
    'Debit (₹)': b.debit,
    'Credit (₹)': b.credit,
    'Narration': b.narration || '',
    'Status': 'Missing in Ledger A',
  }));
  const wsUnrecB = createStyledWorksheet(
    unrecBRows.length > 0 ? unrecBRows : [{ Note: 'All Ledger B transactions reconciled.' }],
    {
      currencyColumns: ['Debit (₹)', 'Credit (₹)'],
      dateColumns: ['Date'],
    }
  );
  XLSX.utils.book_append_sheet(wb, wsUnrecB, 'Unmatched_Ledger_B');

  return wb;
}

/**
 * Returns a realistic sample reciprocal ledger pair (Our Books vs Counterparty Books)
 * with exact voucher matches, date drift, amount variance, and unmatched items for 1-click demos & testing.
 */
export function getSampleLedgerPair(): {
  ledgerA: Record<string, any>[];
  ledgerB: Record<string, any>[];
  ledgerAName: string;
  ledgerBName: string;
} {
  const ledgerA = [
    { Date: '01-10-2026', 'Voucher No': 'INV/26/101', Particulars: 'Sales - Industrial Motors', Debit: 250000, Credit: 0, Narration: 'Billed 10 units of 15HP induction motors' },
    { Date: '05-10-2026', 'Voucher No': 'INV/26/102', Particulars: 'Sales - Spare Gearboxes', Debit: 120000, Credit: 0, Narration: 'Precision helical gear units' },
    { Date: '12-10-2026', 'Voucher No': 'CMS-88210', Particulars: 'Bank Receipt - HDFC RTGS', Debit: 0, Credit: 250000, Narration: 'Payment received against INV/26/101' },
    { Date: '18-10-2026', 'Voucher No': 'INV/26/103', Particulars: 'Sales - Control Panels', Debit: 180000, Credit: 0, Narration: 'PLC automation panels' },
    { Date: '24-10-2026', 'Voucher No': 'CN/26/015', Particulars: 'Credit Note - Damaged Goods', Debit: 0, Credit: 15000, Narration: 'Return of 1 damaged motor assembly' },
    { Date: '28-10-2026', 'Voucher No': 'INV/26/104', Particulars: 'Sales - Annual Maintenance', Debit: 45000, Credit: 0, Narration: 'Billed Q3 preventive service' },
  ];

  const ledgerB = [
    // Reciprocal of INV/26/101 (Credit in their books)
    { Date: '02-10-2026', 'Voucher No': 'INV/26/101', Particulars: 'Purchase - Industrial Motors', Debit: 0, Credit: 250000, Narration: 'Goods received at central depot' },
    // Reciprocal of INV/26/102 (Credit in their books)
    { Date: '05-10-2026', 'Voucher No': 'INV/26/102', Particulars: 'Purchase - Spare Gearboxes', Debit: 0, Credit: 120000, Narration: 'Helical gear spares' },
    // Reciprocal payment with different bank UTR reference and 1-day date drift for Probable Matching
    { Date: '11-10-2026', 'Voucher No': 'UTR-HDFC-991', Particulars: 'Bank Payment - RTGS HDFC', Debit: 250000, Credit: 0, Narration: 'RTGS remittance advice issued' },
    // Reciprocal of INV/26/103 (Booked with ₹5,000 difference for variance testing)
    { Date: '20-10-2026', 'Voucher No': 'INV/26/103', Particulars: 'Purchase - Control Panels', Debit: 0, Credit: 175000, Narration: 'Booked net of ₹5,000 unapproved freight' },
    // Reciprocal of CN/26/015 (Debit in their books)
    { Date: '25-10-2026', 'Voucher No': 'CN/26/015', Particulars: 'Debit Note - Damaged Goods Return', Debit: 15000, Credit: 0, Narration: 'Rejected motor assembly returned' },
    // Unmatched in B (Transit entry not yet received by A)
    { Date: '30-10-2026', 'Voucher No': 'CMS-99402', Particulars: 'Bank Payment - NEFT In Transit', Debit: 100000, Credit: 0, Narration: 'NEFT transferred on month end' },
  ];

  return {
    ledgerA,
    ledgerB,
    ledgerAName: 'Apex_Motors_Customer_Ledger.xlsx',
    ledgerBName: 'Reliance_Vendor_Statement.xlsx',
  };
}
