import * as XLSX from 'xlsx';
import { parseAmount, parseDate, parseDateObj, normalizeRefNo } from './dataSanitizer';
import { detectColumnCategory } from './columnDetector';
import { createStyledWorksheet } from './excelStyler';

export interface ParsedBankItem {
  id: string;
  date: string;
  rawDate: any;
  narration: string;
  refNo: string;
  withdrawal: number;
  deposit: number;
  balance: number;
  matchedLedgerId?: string;
  matchScore: number;
  matchStatus: 'Matched' | 'Probable' | 'Unmatched';
  matchReason?: string;
}

export interface ParsedLedgerItem {
  id: string;
  date: string;
  rawDate: any;
  voucherNo: string;
  particulars: string;
  debit: number; // Deposit in books
  credit: number; // Withdrawal / Payment in books
  matchedBankId?: string;
  isMatched: boolean;
}

export interface BrsCalculation {
  bankClosingBalance: number;
  uncreditedDeposits: number; // Cheques deposited in books but not in bank statement
  unpresentedCheques: number; // Cheques issued in books but not in bank statement
  unbookedBankCharges: number; // Debits in bank not in books
  unbookedBankInterest: number; // Credits in bank not in books
  adjustedBankBalance: number;
  booksClosingBalance: number;
  netVariance: number;
  matchedCount: number;
  unreconciledBankCount: number;
  unreconciledLedgerCount: number;
}

/**
 * Normalizes raw sheet JSON into strongly typed Bank Statement rows
 */
export function parseBankStatementRows(rawRows: Record<string, any>[]): ParsedBankItem[] {
  if (!rawRows || rawRows.length === 0) return [];
  const sample = rawRows[0];
  const colMap: Record<string, string> = {};

  for (const key of Object.keys(sample)) {
    const cat = detectColumnCategory(key);
    if (cat) colMap[cat] = key;
  }

  return rawRows.map((row, idx) => {
    const dateVal = colMap.date ? row[colMap.date] : row.date || row.Date;
    const narrationVal = colMap.narration ? row[colMap.narration] : colMap.partyName ? row[colMap.partyName] : row.narration || row.Narration || row.Description || '';
    const refVal = colMap.refNo ? row[colMap.refNo] : colMap.voucherNo ? row[colMap.voucherNo] : row.refNo || row.RefNo || row['Chq No'] || '';
    const withdrawalVal = colMap.debit ? row[colMap.debit] : row.withdrawal || row.Withdrawal || row.Debit || 0;
    const depositVal = colMap.credit ? row[colMap.credit] : row.deposit || row.Deposit || row.Credit || 0;
    const balanceVal = colMap.balance ? row[colMap.balance] : row.balance || row.Balance || 0;

    return {
      id: `bnk-${idx + 1}`,
      date: parseDate(dateVal, 'DD-MM-YYYY'),
      rawDate: dateVal,
      narration: (narrationVal || '').toString().trim(),
      refNo: (refVal || '').toString().trim(),
      withdrawal: parseAmount(withdrawalVal),
      deposit: parseAmount(depositVal),
      balance: parseAmount(balanceVal),
      matchScore: 0,
      matchStatus: 'Unmatched',
    };
  });
}

/**
 * Normalizes raw sheet JSON into strongly typed Ledger (Bank Book) rows
 */
export function parseLedgerRows(rawRows: Record<string, any>[]): ParsedLedgerItem[] {
  if (!rawRows || rawRows.length === 0) return [];
  const sample = rawRows[0];
  const colMap: Record<string, string> = {};

  for (const key of Object.keys(sample)) {
    const cat = detectColumnCategory(key);
    if (cat) colMap[cat] = key;
  }

  return rawRows.map((row, idx) => {
    const dateVal = colMap.date ? row[colMap.date] : row.date || row.Date;
    const vchVal = colMap.voucherNo ? row[colMap.voucherNo] : colMap.refNo ? row[colMap.refNo] : row.voucherNo || row.VoucherNo || row['Voucher No'] || '';
    const partVal = colMap.partyName ? row[colMap.partyName] : colMap.narration ? row[colMap.narration] : row.particulars || row.Particulars || '';
    const debitVal = colMap.debit ? row[colMap.debit] : row.debit || row.Debit || 0;
    const creditVal = colMap.credit ? row[colMap.credit] : row.credit || row.Credit || 0;

    return {
      id: `led-${idx + 1}`,
      date: parseDate(dateVal, 'DD-MM-YYYY'),
      rawDate: dateVal,
      voucherNo: (vchVal || '').toString().trim(),
      particulars: (partVal || '').toString().trim(),
      debit: parseAmount(debitVal),
      credit: parseAmount(creditVal),
      isMatched: false,
    };
  });
}

/**
 * Executes two-way bank reconciliation between parsed bank statement & ledger entries
 */
export function runBankReconciliation(
  bankItems: ParsedBankItem[],
  ledgerItems: ParsedLedgerItem[],
  dateToleranceDays = 5
): {
  reconciledBank: ParsedBankItem[];
  reconciledLedger: ParsedLedgerItem[];
  calculation: BrsCalculation;
} {
  const bList = bankItems.map((b) => ({ ...b }));
  const lList = ledgerItems.map((l) => ({ ...l }));

  const matchedLedgerIds = new Set<string>();

  // Pass 1: Exact Reference / Cheque No Match
  for (const b of bList) {
    if (b.matchStatus === 'Matched') continue;
    const normBRef = normalizeRefNo(b.refNo);
    if (normBRef && normBRef.length >= 4) {
      const match = lList.find(
        (l) => !matchedLedgerIds.has(l.id) && normalizeRefNo(l.voucherNo) === normBRef
      );
      if (match) {
        // Verify amount parity: Withdrawal in Bank matches Credit in Books, OR Deposit in Bank matches Debit in Books
        const isPayment = Math.abs(b.withdrawal - match.credit) <= 0.01 && b.withdrawal > 0;
        const isDeposit = Math.abs(b.deposit - match.debit) <= 0.01 && b.deposit > 0;

        if (isPayment || isDeposit) {
          b.matchedLedgerId = match.id;
          b.matchStatus = 'Matched';
          b.matchScore = 100;
          b.matchReason = `Exact Cheque/Ref No match (${b.refNo})`;
          match.matchedBankId = b.id;
          match.isMatched = true;
          matchedLedgerIds.add(match.id);
        }
      }
    }
  }

  // Pass 2: Exact Amount Match within Date Tolerance
  for (const b of bList) {
    if (b.matchStatus === 'Matched') continue;

    const bDate = parseDateObj(b.rawDate || b.date);

    for (const l of lList) {
      if (matchedLedgerIds.has(l.id)) continue;

      const isPayment = Math.abs(b.withdrawal - l.credit) <= 0.01 && b.withdrawal > 0;
      const isDeposit = Math.abs(b.deposit - l.debit) <= 0.01 && b.deposit > 0;

      if (isPayment || isDeposit) {
        // Enforce date tolerance if dates are available
        const lDate = parseDateObj(l.rawDate || l.date);
        let withinDateWindow = true;
        if (bDate && lDate && !isNaN(bDate.getTime()) && !isNaN(lDate.getTime())) {
          const dayDiff = Math.abs(bDate.getTime() - lDate.getTime()) / (1000 * 60 * 60 * 24);
          if (dayDiff > dateToleranceDays) {
            withinDateWindow = false;
          }
        }

        if (withinDateWindow) {
          b.matchedLedgerId = l.id;
          b.matchStatus = 'Matched';
          b.matchScore = 95;
          b.matchReason = 'Exact Amount Match with Ledger entry within date tolerance';
          l.matchedBankId = b.id;
          l.isMatched = true;
          matchedLedgerIds.add(l.id);
          break;
        }
      }
    }
  }

  // Pass 3: Approximate Narration Match for remainder
  for (const b of bList) {
    if (b.matchStatus === 'Matched') continue;

    const bNarr = b.narration.toLowerCase();
    for (const l of lList) {
      if (matchedLedgerIds.has(l.id)) continue;

      const lPart = l.particulars.toLowerCase();
      const hasPartyKeyword = lPart.split(/\s+/).some((w) => w.length > 4 && bNarr.includes(w));
      const isPayment = Math.abs(b.withdrawal - l.credit) <= 1.0 && b.withdrawal > 0;
      const isDeposit = Math.abs(b.deposit - l.debit) <= 1.0 && b.deposit > 0;

      if ((isPayment || isDeposit) && hasPartyKeyword) {
        b.matchedLedgerId = l.id;
        b.matchStatus = 'Probable';
        b.matchScore = 85;
        b.matchReason = 'Approximate Narration & Amount alignment';
        l.matchedBankId = b.id;
        l.isMatched = true;
        matchedLedgerIds.add(l.id);
        break;
      }
    }
  }

  // Calculate BRS Equation
  const lastBankItem = bList[bList.length - 1];
  const bankClosingBalance = lastBankItem?.balance || 
    bList.reduce((acc, b) => acc + b.deposit - b.withdrawal, 0);

  // Unpresented Cheques: In books (credits/payments) not cleared in bank statement
  const unpresentedCheques = lList
    .filter((l) => !l.isMatched && l.credit > 0)
    .reduce((acc, l) => acc + l.credit, 0);

  // Uncredited Deposits: In books (debits/receipts) not deposited/cleared in bank statement
  const uncreditedDeposits = lList
    .filter((l) => !l.isMatched && l.debit > 0)
    .reduce((acc, l) => acc + l.debit, 0);

  // Unbooked Bank Charges: Debits in bank not in books
  const unbookedBankCharges = bList
    .filter((b) => b.matchStatus === 'Unmatched' && b.withdrawal > 0)
    .reduce((acc, b) => acc + b.withdrawal, 0);

  // Unbooked Bank Interest/Direct receipts: Credits in bank not in books
  const unbookedBankInterest = bList
    .filter((b) => b.matchStatus === 'Unmatched' && b.deposit > 0)
    .reduce((acc, b) => acc + b.deposit, 0);

  const adjustedBankBalance =
    bankClosingBalance + uncreditedDeposits - unpresentedCheques + unbookedBankCharges - unbookedBankInterest;

  const booksClosingBalance = lList.reduce((acc, l) => acc + l.debit - l.credit, 0);
  const netVariance = Math.abs(adjustedBankBalance - booksClosingBalance);

  const matchedCount = bList.filter((b) => b.matchStatus === 'Matched' || b.matchStatus === 'Probable').length;
  const unreconciledBankCount = bList.filter((b) => b.matchStatus === 'Unmatched').length;
  const unreconciledLedgerCount = lList.filter((l) => !l.isMatched).length;

  return {
    reconciledBank: bList,
    reconciledLedger: lList,
    calculation: {
      bankClosingBalance,
      uncreditedDeposits,
      unpresentedCheques,
      unbookedBankCharges,
      unbookedBankInterest,
      adjustedBankBalance,
      booksClosingBalance,
      netVariance,
      matchedCount,
      unreconciledBankCount,
      unreconciledLedgerCount,
    },
  };
}

/**
 * Creates a professional 4-sheet Bank Reconciliation Excel workbook
 */
export function createBrsWorkbook(
  calc: BrsCalculation,
  bankItems: ParsedBankItem[],
  ledgerItems: ParsedLedgerItem[],
  sourceBankRaw?: Record<string, any>[],
  sourceLedgerRaw?: Record<string, any>[]
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: BRS Summary (Starting with Bank Statement Balance to Reconcile to Book Balance)
  const brsSummary = [
    { Particulars: 'Balance as per Bank Statement', 'Amount (₹)': calc.bankClosingBalance },
    { Particulars: 'Add: Cheques deposited but not yet cleared by bank (Uncredited)', 'Amount (₹)': calc.uncreditedDeposits },
    { Particulars: 'Less: Cheques issued but not yet presented for payment (Unpresented)', 'Amount (₹)': -calc.unpresentedCheques },
    { Particulars: 'Add: Bank charges debited by bank not recorded in books', 'Amount (₹)': calc.unbookedBankCharges },
    { Particulars: 'Less: Direct deposits / interest credited by bank not in books', 'Amount (₹)': -calc.unbookedBankInterest },
    { Particulars: 'Reconciled Balance as per Ledger (Cash Book)', 'Amount (₹)': calc.adjustedBankBalance },
    { Particulars: 'Net Reconciliation Variance', 'Amount (₹)': calc.netVariance },
  ];
  const wsSummary = createStyledWorksheet(brsSummary, {
    currencyColumns: ['Amount (₹)'],
    columnWidths: { Particulars: 55, 'Amount (₹)': 22 },
  });
  XLSX.utils.book_append_sheet(wb, wsSummary, 'BRS_Summary');

  // Sheet 2: Matched Items
  const matched = bankItems
    .filter((b) => b.matchStatus !== 'Unmatched')
    .map((b) => {
      const led = ledgerItems.find((l) => l.id === b.matchedLedgerId);
      return {
        'Bank Date': b.date,
        'Bank Narration': b.narration,
        'Bank Ref No': b.refNo,
        'Withdrawal (Debit)': b.withdrawal,
        'Deposit (Credit)': b.deposit,
        'Ledger Date': led?.date || '',
        'Ledger Voucher No': led?.voucherNo || '',
        'Ledger Particulars': led?.particulars || '',
        'Match Type': b.matchStatus,
        'Match Reason': b.matchReason || '',
      };
    });
  const wsMatched = createStyledWorksheet(
    matched.length > 0 ? matched : [{ Note: 'No matched records' }],
    {
      currencyColumns: ['Withdrawal (Debit)', 'Deposit (Credit)'],
      dateColumns: ['Bank Date', 'Ledger Date'],
    }
  );
  XLSX.utils.book_append_sheet(wb, wsMatched, 'Matched_Transactions');

  // Sheet 3: Unreconciled Bank Items
  const unrecBank = bankItems
    .filter((b) => b.matchStatus === 'Unmatched')
    .map((b) => ({
      'Date': b.date,
      'Narration': b.narration,
      'Ref No': b.refNo,
      'Withdrawal (Debit)': b.withdrawal,
      'Deposit (Credit)': b.deposit,
      'Audit Classification': b.withdrawal > 0 ? 'Unbooked Bank Charge / Direct Debit' : 'Unbooked Credit / Auto-Sweep Interest',
    }));
  const wsUnrecBank = createStyledWorksheet(
    unrecBank.length > 0 ? unrecBank : [{ Note: 'Zero unreconciled bank items' }],
    {
      currencyColumns: ['Withdrawal (Debit)', 'Deposit (Credit)'],
      dateColumns: ['Date'],
    }
  );
  XLSX.utils.book_append_sheet(wb, wsUnrecBank, 'Unreconciled_Bank');

  // Sheet 4: Unreconciled Ledger Items
  const unrecLed = ledgerItems
    .filter((l) => !l.isMatched)
    .map((l) => ({
      'Date': l.date,
      'Voucher No': l.voucherNo,
      'Particulars': l.particulars,
      'Debit (Deposit in Books)': l.debit,
      'Credit (Withdrawal in Books)': l.credit,
      'Audit Classification': l.credit > 0 ? 'Cheque Issued to Vendor (Unpresented)' : 'Cheque Deposited in Bank (Uncredited)',
    }));
  const wsUnrecLed = createStyledWorksheet(
    unrecLed.length > 0 ? unrecLed : [{ Note: 'Zero unreconciled ledger items' }],
    {
      currencyColumns: ['Debit (Deposit in Books)', 'Credit (Withdrawal in Books)'],
      dateColumns: ['Date'],
    }
  );
  XLSX.utils.book_append_sheet(wb, wsUnrecLed, 'Unreconciled_Ledger');

  // Sheet 5 & 6: Preserve Original Uploaded Source Data
  if (sourceBankRaw && sourceBankRaw.length > 0) {
    const wsSrcBank = createStyledWorksheet(sourceBankRaw);
    XLSX.utils.book_append_sheet(wb, wsSrcBank, 'Source_Bank_Statement');
  }

  if (sourceLedgerRaw && sourceLedgerRaw.length > 0) {
    const wsSrcLed = createStyledWorksheet(sourceLedgerRaw);
    XLSX.utils.book_append_sheet(wb, wsSrcLed, 'Source_Books_Ledger');
  }

  return wb;
}
