import * as XLSX from 'xlsx';
import { parseAmount, parseDate, parseDateObj, normalizeRefNo } from './dataSanitizer';
import { createStyledWorksheet } from './excelStyler';
import { detectColumnCategory } from './columnDetector';

export interface ColumnMapping {
  dateCol: string;
  narrationCol: string;
  refCol: string;
  debitCol: string;
  creditCol: string;
  balanceCol: string;
  singleAmountCol?: string;
  drCrCol?: string;
}

export interface ConvertedBankTransaction {
  id: string;
  date: string;
  rawDate: string;
  narration: string;
  refNo: string;
  withdrawal: number;
  deposit: number;
  balance: number;
  category: string;
  suggestedAccountHead: string;
  suggestedRuleReason: string;
  continuityBreak?: boolean;
  continuityExpectedBalance?: number;
  continuityDiscrepancy?: number;
  matchStatus?: 'Matched' | 'Unmatched' | 'Probable';
  matchedLedgerVoucher?: string;
  matchedLedgerParticulars?: string;
}

export interface ContinuityBreak {
  row: number;
  date: string;
  narration: string;
  prevBalance: number;
  withdrawal: number;
  deposit: number;
  expectedBalance: number;
  actualBalance: number;
  diff: number;
}

export interface ContinuityAuditResult {
  hasBreaks: boolean;
  breakCount: number;
  totalDiscrepancy: number;
  breaks: ContinuityBreak[];
}

export interface DraftJournalEntry {
  id: string;
  transactionId: string;
  date: string;
  particulars: string;
  debitAccount: string;
  creditAccount: string;
  amount: number;
  narration: string;
  status: 'Approved' | 'Pending Review';
  suggestedRule: string;
  isApproved: boolean;
  sourceRef: string;
}

export interface BankStatementConversionResult {
  bankName: string;
  statementPeriod: string;
  totalTransactions: number;
  totalDeposits: number;
  totalWithdrawals: number;
  openingBalance: number;
  closingBalance: number;
  netCashFlow: number;
  transactions: ConvertedBankTransaction[];
  detectedMapping: ColumnMapping;
  continuityAudit: ContinuityAuditResult;
  draftJournals: DraftJournalEntry[];
  ledgerReconciliation?: {
    matchedCount: number;
    unmatchedBankCount: number;
    unmatchedLedgerCount: number;
    variance: number;
    unpresentedCheques: number;
    uncreditedDeposits: number;
  };
}

/**
 * CA-grade rule-based Account Head Auto-Suggester
 */
export function suggestAccountHead(
  narration: string,
  isWithdrawal: boolean,
  bankName: string = 'Bank Account'
): { accountHead: string; rule: string; confidence: number } {
  const n = narration.toLowerCase();

  // 1. Statutory & Tax Payments
  if (n.includes('gst') || n.includes('cpass') || n.includes('gstr') || n.includes('gstpmt')) {
    return {
      accountHead: 'GST Payable / Electronic Cash Ledger',
      rule: "Statutory Tax Rule: Identified GST payment keyword in narration",
      confidence: 96,
    };
  }
  if (n.includes('tds') || n.includes('income tax') || n.includes('advance tax') || n.includes('self assessment') || n.includes('challan 280') || n.includes('challan 281')) {
    return {
      accountHead: 'Direct Taxes / TDS Payable A/c',
      rule: "Statutory Tax Rule: Identified Direct Tax / TDS payment keyword",
      confidence: 96,
    };
  }

  // 2. Payroll & Remuneration
  if (n.includes('salary') || n.includes('payroll') || n.includes('wages') || n.includes('stipend')) {
    return {
      accountHead: 'Salaries & Wages Payable A/c',
      rule: "Payroll Rule: Identified Salary / Wages keyword",
      confidence: 95,
    };
  }
  if (n.includes('director') || n.includes('remuneration') || n.includes('drawings')) {
    return {
      accountHead: 'Director Remuneration / Drawings A/c',
      rule: "Corporate Governance: Identified Director transaction keyword",
      confidence: 92,
    };
  }

  // 3. Rent & Utilities
  if (n.includes('rent') || n.includes('lease') || n.includes('tenancy') || n.includes('office premises')) {
    return {
      accountHead: 'Rent Expense A/c',
      rule: "Facility Rule: Identified Rent / Lease keyword",
      confidence: 92,
    };
  }
  if (n.includes('bescom') || n.includes('msedcl') || n.includes('cesc') || n.includes('electricity') || n.includes('power') || n.includes('water bill')) {
    return {
      accountHead: 'Electricity & Utility Expenses A/c',
      rule: "Utility Rule: Identified electricity / water utility keyword",
      confidence: 94,
    };
  }

  // 4. Banking & Finance
  if (n.includes('bank charge') || n.includes('chg') || n.includes('sms charge') || n.includes('annual fee') || n.includes('folio charge') || n.includes('consolidated charge')) {
    return {
      accountHead: 'Bank Charges & Commission A/c',
      rule: "Financial Cost: Identified Bank Service Fee / Charges keyword",
      confidence: 98,
    };
  }
  if (n.includes('interest') || n.includes('int.') || n.includes('auto sweep')) {
    return {
      accountHead: isWithdrawal ? 'Bank Interest & Finance Charges A/c' : 'Interest Received on Bank Deposits A/c',
      rule: "Finance Rule: Identified Interest keyword",
      confidence: 95,
    };
  }
  if (n.includes('emi') || n.includes('loan') || n.includes('bajaj fin') || n.includes('hdfc cred') || n.includes('repayment')) {
    return {
      accountHead: 'Bank Borrowings / Term Loan Repayment A/c',
      rule: "Debt Rule: Identified EMI / Loan Repayment keyword",
      confidence: 90,
    };
  }

  // 5. Operating & Travel Expenses
  if (n.includes('petrol') || n.includes('diesel') || n.includes('fuel') || n.includes('iocl') || n.includes('bpcl') || n.includes('hpcl')) {
    return {
      accountHead: 'Vehicle & Conveyance Expense A/c',
      rule: "Logistics Rule: Identified Fuel / Petrol Pump keyword",
      confidence: 92,
    };
  }
  if (n.includes('air') || n.includes('irctc') || n.includes('makemytrip') || n.includes('uber') || n.includes('ola') || n.includes('flight') || n.includes('hotel')) {
    return {
      accountHead: 'Travel & Conveyance Expenses A/c',
      rule: "Travel Rule: Identified Travel / Ticketing keyword",
      confidence: 90,
    };
  }
  if (n.includes('audit') || n.includes('legal') || n.includes('consulting') || n.includes('professional') || n.includes('advocate')) {
    return {
      accountHead: 'Legal & Professional Charges A/c',
      rule: "Professional Services: Identified Legal / Audit keyword",
      confidence: 91,
    };
  }
  if (n.includes('insurance') || n.includes('lic') || n.includes('hdfc life') || n.includes('icici pru') || n.includes('bajaj allianz')) {
    return {
      accountHead: 'Insurance Premium Expense A/c',
      rule: "Risk Coverage: Identified Insurance Company keyword",
      confidence: 94,
    };
  }

  // 6. Cash In Hand
  if (n.includes('atm') || n.includes('cash wdl') || n.includes('self wdl') || n.includes('cash deposit')) {
    return {
      accountHead: 'Cash in Hand A/c',
      rule: "Cash Operations: Identified ATM / Cash transaction keyword",
      confidence: 95,
    };
  }

  // 7. Commercial Trade Counterparties
  if (isWithdrawal) {
    return {
      accountHead: 'Sundry Creditors (Vendor Clearing A/c)',
      rule: "Trade Payables: Commercial outward remittance pending specific vendor voucher mapping",
      confidence: 75,
    };
  } else {
    return {
      accountHead: 'Sundry Debtors (Customer Clearing A/c)',
      rule: "Trade Receivables: Commercial inward remittance pending specific customer invoice allocation",
      confidence: 75,
    };
  }
}

/**
 * Validates running balance mathematical continuity across transactions
 */
export function validateBalanceContinuity(
  transactions: ConvertedBankTransaction[]
): ContinuityAuditResult {
  const breaks: ContinuityBreak[] = [];
  let totalDiscrepancy = 0;

  for (let i = 1; i < transactions.length; i++) {
    const prev = transactions[i - 1];
    const curr = transactions[i];

    // Continuity equation: PrevBalance - Withdrawal + Deposit = ExpectedCurrentBalance
    const expected = prev.balance - curr.withdrawal + curr.deposit;
    const diff = Math.abs(curr.balance - expected);

    if (diff > 0.05 && (curr.balance !== 0 || curr.withdrawal > 0 || curr.deposit > 0)) {
      curr.continuityBreak = true;
      curr.continuityExpectedBalance = expected;
      curr.continuityDiscrepancy = curr.balance - expected;

      breaks.push({
        row: i + 1,
        date: curr.date,
        narration: curr.narration,
        prevBalance: prev.balance,
        withdrawal: curr.withdrawal,
        deposit: curr.deposit,
        expectedBalance: expected,
        actualBalance: curr.balance,
        diff: curr.balance - expected,
      });

      totalDiscrepancy += diff;
    }
  }

  return {
    hasBreaks: breaks.length > 0,
    breakCount: breaks.length,
    totalDiscrepancy,
    breaks,
  };
}

/**
 * Automatically suggests column mappings from sample spreadsheet headers
 */
export function detectBankColumnMapping(headers: string[]): ColumnMapping {
  const mapping: ColumnMapping = {
    dateCol: '',
    narrationCol: '',
    refCol: '',
    debitCol: '',
    creditCol: '',
    balanceCol: '',
  };

  for (const h of headers) {
    const cat = detectColumnCategory(h);
    const lower = h.toLowerCase().trim();

    if (!mapping.dateCol && (cat === 'date' || /\b(date|txn dt|value dt|entry dt)\b/i.test(lower))) {
      mapping.dateCol = h;
    } else if (
      !mapping.narrationCol &&
      (cat === 'narration' || /\b(narration|description|particulars|remarks|details)\b/i.test(lower))
    ) {
      mapping.narrationCol = h;
    } else if (
      !mapping.refCol &&
      (cat === 'voucherNo' || cat === 'refNo' || /\b(ref|chq|cheque|utr|trans id|txn id)\b/i.test(lower))
    ) {
      mapping.refCol = h;
    } else if (
      !mapping.debitCol &&
      (cat === 'debit' || /\b(debit|withdrawal|dr|payment)\b/i.test(lower))
    ) {
      mapping.debitCol = h;
    } else if (
      !mapping.creditCol &&
      (cat === 'credit' || /\b(credit|deposit|cr|receipt)\b/i.test(lower))
    ) {
      mapping.creditCol = h;
    } else if (
      !mapping.balanceCol &&
      (cat === 'balance' || /\b(balance|bal|closing bal)\b/i.test(lower))
    ) {
      mapping.balanceCol = h;
    }
  }

  // Fallback: If only one "Amount" column exists
  if (!mapping.debitCol && !mapping.creditCol) {
    const amtCol = headers.find(
      (h) => detectColumnCategory(h) === 'totalAmount' || detectColumnCategory(h) === 'taxableAmount' || /\b(amount|amt|txn amt)\b/i.test(h)
    );
    if (amtCol) mapping.singleAmountCol = amtCol;
  }

  return mapping;
}

/**
 * Generates initial draft journal entries from converted bank transactions
 */
export function generateDraftJournalEntries(
  transactions: ConvertedBankTransaction[],
  bankAccountName: string = 'Bank Account'
): DraftJournalEntry[] {
  return transactions.map((t, idx) => {
    const isWdl = t.withdrawal > 0;
    const amount = isWdl ? t.withdrawal : t.deposit;
    const debitAccount = isWdl ? t.suggestedAccountHead : bankAccountName;
    const creditAccount = isWdl ? bankAccountName : t.suggestedAccountHead;

    return {
      id: `dr-entry-${idx + 1}`,
      transactionId: t.id,
      date: t.date,
      particulars: t.narration,
      debitAccount,
      creditAccount,
      amount,
      narration: `Being ${t.narration} recorded in books via Statement Integration`,
      status: 'Pending Review',
      suggestedRule: t.suggestedRuleReason,
      isApproved: false,
      sourceRef: t.refNo,
    };
  });
}

/**
 * Reconciles parsed bank transactions against an imported client accounting ledger
 */
export function matchBankWithClientLedger(
  bankTransactions: ConvertedBankTransaction[],
  ledgerRows: Record<string, any>[]
): {
  reconciledBank: ConvertedBankTransaction[];
  matchedCount: number;
  unmatchedBankCount: number;
  unmatchedLedgerCount: number;
  variance: number;
  unpresentedCheques: number;
  uncreditedDeposits: number;
} {
  // Normalize ledger rows
  const normLedger = ledgerRows.map((r, idx) => {
    const keys = Object.keys(r);
    const dateK = keys.find((k) => /date|txn dt/i.test(k)) || '';
    const partK = keys.find((k) => /particular|narration|desc/i.test(k)) || '';
    const vchK = keys.find((k) => /vch|voucher|ref|chq/i.test(k)) || '';
    const drK = keys.find((k) => /debit|dr|receipt/i.test(k)) || '';
    const crK = keys.find((k) => /credit|cr|payment/i.test(k)) || '';

    return {
      id: `ldg-${idx + 1}`,
      date: dateK ? parseDate(r[dateK], 'DD-MM-YYYY') : '',
      particulars: partK ? String(r[partK] ?? '').trim() : '',
      voucherNo: vchK ? String(r[vchK] ?? '').trim() : '',
      debit: drK ? Math.abs(parseAmount(r[drK])) : 0, // Book debit = Bank deposit
      credit: crK ? Math.abs(parseAmount(r[crK])) : 0, // Book credit = Bank withdrawal
      matched: false,
    };
  });

  const matchedLedgerIds = new Set<string>();

  // Pass 1: Ref/Cheque match
  for (const b of bankTransactions) {
    const normBRef = normalizeRefNo(b.refNo);
    if (normBRef && normBRef.length >= 3) {
      const match = normLedger.find(
        (l) => !matchedLedgerIds.has(l.id) && normalizeRefNo(l.voucherNo) === normBRef
      );
      if (match) {
        const amtMatch = (b.withdrawal > 0 && Math.abs(b.withdrawal - match.credit) <= 0.05) ||
                         (b.deposit > 0 && Math.abs(b.deposit - match.debit) <= 0.05);
        if (amtMatch) {
          b.matchStatus = 'Matched';
          b.matchedLedgerVoucher = match.voucherNo;
          b.matchedLedgerParticulars = match.particulars;
          match.matched = true;
          matchedLedgerIds.add(match.id);
        }
      }
    }
  }

  // Pass 2: Exact Amount within Date Tolerance
  for (const b of bankTransactions) {
    if (b.matchStatus === 'Matched') continue;

    const bDate = parseDateObj(b.rawDate || b.date);

    for (const l of normLedger) {
      if (matchedLedgerIds.has(l.id)) continue;

      const amtMatch = (b.withdrawal > 0 && Math.abs(b.withdrawal - l.credit) <= 0.05) ||
                       (b.deposit > 0 && Math.abs(b.deposit - l.debit) <= 0.05);

      if (amtMatch) {
        const lDate = parseDateObj(l.date);
        let withinDate = true;
        if (bDate && lDate && !isNaN(bDate.getTime()) && !isNaN(lDate.getTime())) {
          const diffDays = Math.abs(bDate.getTime() - lDate.getTime()) / (1000 * 60 * 60 * 24);
          if (diffDays > 7) withinDate = false;
        }

        if (withinDate) {
          b.matchStatus = 'Matched';
          b.matchedLedgerVoucher = l.voucherNo || 'Ledger Entry';
          b.matchedLedgerParticulars = l.particulars;
          l.matched = true;
          matchedLedgerIds.add(l.id);
          break;
        }
      }
    }
  }

  // Mark remaining
  for (const b of bankTransactions) {
    if (!b.matchStatus) b.matchStatus = 'Unmatched';
  }

  const matchedCount = bankTransactions.filter((b) => b.matchStatus === 'Matched').length;
  const unmatchedBankCount = bankTransactions.filter((b) => b.matchStatus === 'Unmatched').length;
  const unmatchedLedger = normLedger.filter((l) => !l.matched);

  const unpresentedCheques = unmatchedLedger.reduce((sum, l) => sum + l.credit, 0);
  const uncreditedDeposits = unmatchedLedger.reduce((sum, l) => sum + l.debit, 0);
  const variance = unpresentedCheques - uncreditedDeposits;

  return {
    reconciledBank: bankTransactions,
    matchedCount,
    unmatchedBankCount,
    unmatchedLedgerCount: unmatchedLedger.length,
    variance,
    unpresentedCheques,
    uncreditedDeposits,
  };
}

/**
 * Transforms raw bank statement rows according to confirmed column mappings
 */
export function convertBankStatementRows(
  rows: Record<string, any>[],
  mapping: ColumnMapping,
  bankName: string = 'Standard Bank Statement',
  ledgerRows?: Record<string, any>[]
): BankStatementConversionResult {
  if (!rows || rows.length === 0) {
    return {
      bankName,
      statementPeriod: 'N/A',
      totalTransactions: 0,
      totalDeposits: 0,
      totalWithdrawals: 0,
      openingBalance: 0,
      closingBalance: 0,
      netCashFlow: 0,
      transactions: [],
      detectedMapping: mapping,
      continuityAudit: { hasBreaks: false, breakCount: 0, totalDiscrepancy: 0, breaks: [] },
      draftJournals: [],
    };
  }

  let totalDeposits = 0;
  let totalWithdrawals = 0;
  const transactions: ConvertedBankTransaction[] = [];

  for (let idx = 0; idx < rows.length; idx++) {
    const raw = rows[idx];
    const rawDate = mapping.dateCol ? String(raw[mapping.dateCol] ?? '').trim() : '';
    const formattedDate = rawDate ? parseDate(rawDate, 'DD-MM-YYYY') : '';

    const narration = mapping.narrationCol ? String(raw[mapping.narrationCol] ?? '').trim() : '';
    const refNo = mapping.refCol ? String(raw[mapping.refCol] ?? '').trim() : '';

    let withdrawal = 0;
    let deposit = 0;

    if (mapping.debitCol && mapping.creditCol) {
      withdrawal = Math.abs(parseAmount(raw[mapping.debitCol]));
      deposit = Math.abs(parseAmount(raw[mapping.creditCol]));
    } else if (mapping.singleAmountCol) {
      const amt = parseAmount(raw[mapping.singleAmountCol]);
      const drCr = mapping.drCrCol ? String(raw[mapping.drCrCol] ?? '').trim().toLowerCase() : '';
      if (drCr.includes('dr') || amt < 0) {
        withdrawal = Math.abs(amt);
      } else {
        deposit = Math.abs(amt);
      }
    }

    const balance = mapping.balanceCol ? parseAmount(raw[mapping.balanceCol]) : 0;

    // Suggest Account Head
    const isWdl = withdrawal > 0;
    const headSuggestion = suggestAccountHead(narration, isWdl, bankName);

    // Categorization heuristic
    let category = 'Operational';
    const narrLower = narration.toLowerCase();
    if (narrLower.includes('neft') || narrLower.includes('rtgs') || narrLower.includes('imps') || narrLower.includes('upi')) {
      category = 'Electronic Transfer';
    } else if (narrLower.includes('chq') || narrLower.includes('cheque') || narrLower.includes('clearing')) {
      category = 'Cheque Clearing';
    } else if (narrLower.includes('charge') || narrLower.includes('fee') || narrLower.includes('gst')) {
      category = 'Bank Charges';
    } else if (narrLower.includes('interest') || narrLower.includes('int.')) {
      category = 'Bank Interest';
    } else if (narrLower.includes('salary') || narrLower.includes('payroll')) {
      category = 'Payroll';
    } else if (narrLower.includes('atm') || narrLower.includes('cash')) {
      category = 'Cash Operations';
    }

    if (formattedDate || narration || withdrawal > 0 || deposit > 0) {
      totalDeposits += deposit;
      totalWithdrawals += withdrawal;

      transactions.push({
        id: `txn-${idx + 1}`,
        date: formattedDate || rawDate,
        rawDate,
        narration: narration || 'Bank Transaction',
        refNo: refNo || '-',
        withdrawal,
        deposit,
        balance,
        category,
        suggestedAccountHead: headSuggestion.accountHead,
        suggestedRuleReason: headSuggestion.rule,
      });
    }
  }

  // 1. Audit Balance Continuity
  const continuityAudit = validateBalanceContinuity(transactions);

  // 2. Draft Journal Entries
  const draftJournals = generateDraftJournalEntries(transactions, bankName);

  // 3. Ledger Reconciliation if ledger provided
  let ledgerRecResult = undefined;
  if (ledgerRows && ledgerRows.length > 0) {
    const rec = matchBankWithClientLedger(transactions, ledgerRows);
    ledgerRecResult = {
      matchedCount: rec.matchedCount,
      unmatchedBankCount: rec.unmatchedBankCount,
      unmatchedLedgerCount: rec.unmatchedLedgerCount,
      variance: rec.variance,
      unpresentedCheques: rec.unpresentedCheques,
      uncreditedDeposits: rec.uncreditedDeposits,
    };
  }

  const openingBalance = transactions.length > 0 ? transactions[0].balance + (transactions[0].withdrawal - transactions[0].deposit) : 0;
  const closingBalance = transactions.length > 0 ? transactions[transactions.length - 1].balance : 0;
  const netCashFlow = totalDeposits - totalWithdrawals;

  const firstDate = transactions[0]?.date || '';
  const lastDate = transactions[transactions.length - 1]?.date || '';
  const statementPeriod = firstDate && lastDate ? `${firstDate} to ${lastDate}` : 'Current Statement';

  return {
    bankName,
    statementPeriod,
    totalTransactions: transactions.length,
    totalDeposits,
    totalWithdrawals,
    openingBalance,
    closingBalance,
    netCashFlow,
    transactions,
    detectedMapping: mapping,
    continuityAudit,
    draftJournals,
    ledgerReconciliation: ledgerRecResult,
  };
}

/**
 * Generates CSV representation of clean statement & draft entries
 */
export function exportBankStatementCsv(result: BankStatementConversionResult): string {
  const headers = [
    'S.No',
    'Date',
    'Particulars / Narration',
    'Ref / Cheque No',
    'Withdrawal (Debit ₹)',
    'Deposit (Credit ₹)',
    'Running Balance (₹)',
    'Category',
    'Suggested Account Head',
    'Reconciliation Status',
  ];

  const escapeCsv = (val: any) => `"${String(val ?? '').replace(/"/g, '""')}"`;

  const rows = result.transactions.map((t, idx) => [
    idx + 1,
    escapeCsv(t.date),
    escapeCsv(t.narration),
    escapeCsv(t.refNo),
    t.withdrawal.toFixed(2),
    t.deposit.toFixed(2),
    t.balance.toFixed(2),
    escapeCsv(t.category),
    escapeCsv(t.suggestedAccountHead),
    escapeCsv(t.matchStatus || 'Unreconciled'),
  ]);

  return [headers.join(','), ...rows.map((r) => r.join(','))].join('\r\n');
}

/**
 * Creates professional 5-sheet standardized Excel workbook for bank statement & audit reconciliation
 */
export function createConvertedStatementWorkbook(
  result: BankStatementConversionResult,
  rawSourceRows?: Record<string, any>[]
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Standardized Transactions
  const exportRows = result.transactions.map((t, idx) => ({
    'S.No': idx + 1,
    'Transaction Date': t.date,
    'Particulars / Narration': t.narration,
    'Chq / Ref No': t.refNo,
    'Withdrawal (Debit ₹)': t.withdrawal,
    'Deposit (Credit ₹)': t.deposit,
    'Running Balance (₹)': t.balance,
    'Suggested Account Head': t.suggestedAccountHead,
    'Account Head Rule': t.suggestedRuleReason,
    'Reconciliation Status': t.matchStatus || 'Unreconciled',
    'Continuity Status': t.continuityBreak ? 'Discrepancy Detected' : 'Verified Continuous',
  }));

  const wsTxns = createStyledWorksheet(exportRows, {
    currencyColumns: ['Withdrawal (Debit ₹)', 'Deposit (Credit ₹)', 'Running Balance (₹)'],
    dateColumns: ['Transaction Date'],
    columnWidths: {
      'S.No': 8,
      'Transaction Date': 16,
      'Particulars / Narration': 40,
      'Chq / Ref No': 18,
      'Withdrawal (Debit ₹)': 18,
      'Deposit (Credit ₹)': 18,
      'Running Balance (₹)': 20,
      'Suggested Account Head': 32,
      'Account Head Rule': 36,
      'Reconciliation Status': 20,
      'Continuity Status': 20,
    },
  });
  XLSX.utils.book_append_sheet(wb, wsTxns, 'Clean_Statement');

  // Sheet 2: Bank Reconciliation Statement (BRS)
  const brsRows = [
    { Parameter: 'Bank Account Name', Amount: result.bankName },
    { Parameter: 'Statement Period', Amount: result.statementPeriod },
    { Parameter: 'Closing Balance as per Bank Statement (A)', Amount: result.closingBalance },
    { Parameter: 'Add: Cheques Issued in Books but Not Presented in Bank', Amount: result.ledgerReconciliation?.unpresentedCheques || 0 },
    { Parameter: 'Less: Cheques Deposited in Books but Not Credited by Bank', Amount: result.ledgerReconciliation?.uncreditedDeposits || 0 },
    { Parameter: 'Reconciled Adjusted Balance', Amount: result.closingBalance + (result.ledgerReconciliation?.unpresentedCheques || 0) - (result.ledgerReconciliation?.uncreditedDeposits || 0) },
    { Parameter: 'Reconciliation Status', Amount: result.ledgerReconciliation?.variance === 0 ? 'Fully Reconciled' : 'Pending Book Adjustments' },
    { Parameter: 'Total Matched Bank Entries', Amount: result.ledgerReconciliation?.matchedCount || 0 },
    { Parameter: 'Unmatched Bank Transactions', Amount: result.ledgerReconciliation?.unmatchedBankCount || result.totalTransactions },
  ];
  const wsBRS = createStyledWorksheet(brsRows, {
    currencyColumns: ['Amount'],
    columnWidths: { Parameter: 50, Amount: 30 },
  });
  XLSX.utils.book_append_sheet(wb, wsBRS, 'Bank_Reconciliation_BRS');

  // Sheet 3: Unmatched Transactions
  const unmatched = result.transactions.filter((t) => t.matchStatus === 'Unmatched' || !t.matchStatus);
  const unmatchedRows = unmatched.map((u, idx) => ({
    'S.No': idx + 1,
    'Date': u.date,
    'Narration': u.narration,
    'Ref / Cheque No': u.refNo,
    'Debit (₹)': u.withdrawal,
    'Credit (₹)': u.deposit,
    'Proposed Account Head': u.suggestedAccountHead,
    'Audit Action Required': u.withdrawal > 0 ? 'Create Payment Voucher in Books' : 'Create Receipt Voucher in Books',
  }));
  const wsUnmatched = createStyledWorksheet(
    unmatchedRows.length > 0 ? unmatchedRows : [{ Note: 'All bank transactions have been matched against ledger.' }],
    {
      currencyColumns: ['Debit (₹)', 'Credit (₹)'],
      columnWidths: { 'S.No': 8, Date: 16, Narration: 40, 'Ref / Cheque No': 18, 'Debit (₹)': 18, 'Credit (₹)': 18, 'Proposed Account Head': 32, 'Audit Action Required': 35 },
    }
  );
  XLSX.utils.book_append_sheet(wb, wsUnmatched, 'Unmatched_Transactions');

  // Sheet 4: Draft Journal Entries
  const journalRows = result.draftJournals.map((j, idx) => ({
    'Entry #': idx + 1,
    'Date': j.date,
    'Debit Account': j.debitAccount,
    'Credit Account': j.creditAccount,
    'Amount (₹)': j.amount,
    'Narration / Particulars': j.narration,
    'Approval Status': j.isApproved ? 'Approved by CA' : 'Draft Pending Review',
    'Classification Rule': j.suggestedRule,
    'Source Ref': j.sourceRef,
  }));
  const wsJournals = createStyledWorksheet(journalRows, {
    currencyColumns: ['Amount (₹)'],
    dateColumns: ['Date'],
    columnWidths: {
      'Entry #': 10,
      Date: 14,
      'Debit Account': 35,
      'Credit Account': 35,
      'Amount (₹)': 18,
      'Narration / Particulars': 45,
      'Approval Status': 20,
      'Classification Rule': 40,
      'Source Ref': 18,
    },
  });
  XLSX.utils.book_append_sheet(wb, wsJournals, 'Draft_Journal_Entries');

  // Sheet 5: Continuity Audit Report
  const contRows = result.continuityAudit.breaks.map((b, idx) => ({
    'Discrepancy #': idx + 1,
    'Row in Statement': b.row,
    'Date': b.date,
    'Narration': b.narration,
    'Previous Balance (₹)': b.prevBalance,
    'Withdrawal (Debit ₹)': b.withdrawal,
    'Deposit (Credit ₹)': b.deposit,
    'Expected Balance (₹)': b.expectedBalance,
    'Actual Reported Balance (₹)': b.actualBalance,
    'Continuity Break (Variance ₹)': b.diff,
  }));
  const wsCont = createStyledWorksheet(
    contRows.length > 0
      ? contRows
      : [
          {
            'Audit Summary': '100% Mathematical Continuity Verified',
            'Result': `All ${result.totalTransactions} transactions adhere to: Balance[i-1] - Withdrawal[i] + Deposit[i] = Balance[i]`,
          },
        ],
    {
      currencyColumns: ['Previous Balance (₹)', 'Withdrawal (Debit ₹)', 'Deposit (Credit ₹)', 'Expected Balance (₹)', 'Actual Reported Balance (₹)', 'Continuity Break (Variance ₹)'],
      columnWidths: {
        'Discrepancy #': 14,
        'Row in Statement': 16,
        Date: 14,
        Narration: 35,
        'Previous Balance (₹)': 20,
        'Withdrawal (Debit ₹)': 18,
        'Deposit (Credit ₹)': 18,
        'Expected Balance (₹)': 20,
        'Actual Reported Balance (₹)': 22,
        'Continuity Break (Variance ₹)': 25,
      },
    }
  );
  XLSX.utils.book_append_sheet(wb, wsCont, 'Continuity_Audit');

  return wb;
}
