import * as XLSX from 'xlsx';
import { parseAmount, parseDate, parseDateObj } from './dataSanitizer';
import { createStyledWorksheet } from './excelStyler';

export type LedgerType = 'receivables' | 'payables';

export interface AgingBucketConfig {
  id: string;
  label: string;
  minDays: number;
  maxDays: number; // Use Infinity for 90+ or 365+
}

export const DEFAULT_STANDARD_BUCKETS: AgingBucketConfig[] = [
  { id: 'b1', label: '0 – 30 Days', minDays: 0, maxDays: 30 },
  { id: 'b2', label: '31 – 60 Days', minDays: 31, maxDays: 60 },
  { id: 'b3', label: '61 – 90 Days', minDays: 61, maxDays: 90 },
  { id: 'b4', label: '91 – 180 Days', minDays: 91, maxDays: 180 },
  { id: 'b5', label: '180+ Days', minDays: 181, maxDays: Infinity },
];

export const SCHEDULE_III_BUCKETS: AgingBucketConfig[] = [
  { id: 's1', label: '< 6 Months', minDays: 0, maxDays: 180 },
  { id: 's2', label: '6 Months – 1 Year', minDays: 181, maxDays: 365 },
  { id: 's3', label: '1 – 2 Years', minDays: 366, maxDays: 730 },
  { id: 's4', label: '2 – 3 Years', minDays: 731, maxDays: 1095 },
  { id: 's5', label: '> 3 Years', minDays: 1096, maxDays: Infinity },
];

export interface AgingTransaction {
  id: string;
  rowIndex: number;
  partyName: string;
  gstin: string;
  invoiceNo: string;
  invoiceDate: string;
  dateObj: Date | null;
  dueDate: string;
  ageDays: number;
  grossAmount: number;
  paidAmount: number;
  outstandingBalance: number;
  bucketId: string;
  bucketLabel: string;
  isCreditBalance: boolean;
  msmedExceeded: boolean;
  raw: Record<string, any>;
}

export interface PartyAgingSummary {
  partyName: string;
  gstin: string;
  totalGross: number;
  totalPaid: number;
  netOutstanding: number;
  bucketAmounts: Record<string, number>;
  overdueAmount: number;
  overduePercentage: number;
  riskCategory: 'current' | 'moderate' | 'high_risk' | 'credit_balance';
  transactionCount: number;
  msmedCount: number;
}

export interface AgingCalculationResult {
  ledgerType: LedgerType;
  asOfDate: string;
  creditPeriodDays: number;
  totalParties: number;
  totalGross: number;
  totalPaid: number;
  totalOutstanding: number;
  bucketTotals: Record<string, number>;
  bucketPercentages: Record<string, number>;
  reconciliationDifference: number;
  isReconciled: boolean;
  parties: PartyAgingSummary[];
  transactions: AgingTransaction[];
  assumptions: string[];
}

export interface AgingColumnMapping {
  partyCol: string;
  gstinCol: string;
  invoiceNoCol: string;
  dateCol: string;
  dueDateCol: string;
  grossAmountCol: string;
  paidAmountCol: string;
  balanceCol: string;
}

import { detectColumnCategory } from './columnDetector';

/**
 * Auto-detects column headers for receivables & payables ledgers
 */
export function detectAgingColumns(headers: string[]): AgingColumnMapping {
  const mapping: AgingColumnMapping = {
    partyCol: '',
    gstinCol: '',
    invoiceNoCol: '',
    dateCol: '',
    dueDateCol: '',
    grossAmountCol: '',
    paidAmountCol: '',
    balanceCol: '',
  };

  for (const h of headers) {
    const clean = h.toLowerCase().trim();
    const cat = detectColumnCategory(h);

    if (!mapping.partyCol && (cat === 'partyName' || /party|customer|debtor|vendor|creditor|client|buyer|ledger.*name|account/i.test(clean))) {
      mapping.partyCol = h;
    } else if (!mapping.gstinCol && (cat === 'gstin' || /gstin|gst.*no|tax.*id/i.test(clean))) {
      mapping.gstinCol = h;
    } else if (!mapping.invoiceNoCol && (cat === 'voucherNo' || /inv.*no|bill.*no|ref.*no|voucher.*no|doc.*no/i.test(clean))) {
      mapping.invoiceNoCol = h;
    } else if (!mapping.dueDateCol && (cat === 'dueDate' || /due.*date|pay.*due|maturity.*date/i.test(clean))) {
      mapping.dueDateCol = h;
    } else if (!mapping.dateCol && (cat === 'date' || (/inv.*date|bill.*date|txn.*date|date/i.test(clean) && !/due/i.test(clean)))) {
      mapping.dateCol = h;
    } else if (!mapping.grossAmountCol && (cat === 'totalAmount' || /invoice.*amt|gross|debit|bill.*amt|billed|total.*amt/i.test(clean))) {
      mapping.grossAmountCol = h;
    } else if (!mapping.paidAmountCol && /paid|received|credit|payment|receipt.*amt/i.test(clean)) {
      mapping.paidAmountCol = h;
    } else if (!mapping.balanceCol && (cat === 'balance' || /balance|outstanding|net.*amt|due.*amt|pending/i.test(clean))) {
      mapping.balanceCol = h;
    }
  }

  return mapping;
}

/**
 * Calculates aging analysis based on configured buckets and as-of date
 */
export function calculateAgingAnalysis(
  rows: Record<string, any>[],
  mapping: AgingColumnMapping,
  options: {
    ledgerType: LedgerType;
    asOfDateStr: string; // DD-MM-YYYY or YYYY-MM-DD
    creditPeriodDays: number;
    buckets: AgingBucketConfig[];
    allocateCreditsFifo: boolean;
  }
): AgingCalculationResult {
  const { ledgerType, asOfDateStr, creditPeriodDays, buckets, allocateCreditsFifo } = options;

  // Parse as of date robustly
  const asOfDate = parseDateObj(asOfDateStr) || new Date();

  const transactions: AgingTransaction[] = [];
  const partyMap = new Map<string, { gstin: string; txns: AgingTransaction[] }>();

  // Process rows
  rows.forEach((r, idx) => {
    const partyName = String(r[mapping.partyCol] || '').trim() || 'Unknown Party';
    const gstin = String(r[mapping.gstinCol] || '').trim().toUpperCase();
    const invoiceNo = String(r[mapping.invoiceNoCol] || '').trim() || `REF-${idx + 1}`;
    const rawDate = r[mapping.dateCol];
    const dateObj = parseDateObj(rawDate);
    const invoiceDate = dateObj ? parseDate(dateObj, 'DD-MM-YYYY') : (rawDate ? String(rawDate).trim() : 'N/A');

    const grossAmount = parseAmount(r[mapping.grossAmountCol]);
    const paidAmount = parseAmount(r[mapping.paidAmountCol]);
    let balance = parseAmount(r[mapping.balanceCol]);

    // If balance not given, derive from gross - paid
    if (balance === 0 && (grossAmount !== 0 || paidAmount !== 0)) {
      balance = grossAmount - paidAmount;
    } else if (grossAmount === 0 && paidAmount === 0 && balance !== 0) {
      // Direct balance provided
    }

    // Determine age in days from explicit due date or invoice date + credit period
    let ageDays = 0;
    let dueDate = '';
    const rawDueDate = mapping.dueDateCol ? r[mapping.dueDateCol] : undefined;
    const explicitDueObj = rawDueDate ? parseDateObj(rawDueDate) : null;

    if (explicitDueObj && !isNaN(explicitDueObj.getTime())) {
      dueDate = parseDate(explicitDueObj, 'DD-MM-YYYY');
      const diffMs = asOfDate.getTime() - explicitDueObj.getTime();
      ageDays = Math.max(0, Math.floor(diffMs / (1000 * 60 * 60 * 24)));
    } else if (dateObj && !isNaN(dateObj.getTime())) {
      const dueObj = new Date(dateObj.getTime() + creditPeriodDays * 24 * 60 * 60 * 1000);
      dueDate = parseDate(dueObj, 'DD-MM-YYYY');
      const diffMs = asOfDate.getTime() - dueObj.getTime();
      ageDays = Math.max(0, Math.floor(diffMs / (1000 * 60 * 60 * 24)));
    }

    // Assign bucket
    let assignedBucket = buckets[buckets.length - 1]; // fallback to last bucket
    for (const b of buckets) {
      if (ageDays >= b.minDays && ageDays <= b.maxDays) {
        assignedBucket = b;
        break;
      }
    }

    // MSMED 45 days limit check for payables
    const msmedExceeded = ledgerType === 'payables' && ageDays > 45;

    const txn: AgingTransaction = {
      id: `aging-txn-${idx + 1}`,
      rowIndex: idx + 2,
      partyName,
      gstin,
      invoiceNo,
      invoiceDate: invoiceDate || 'N/A',
      dateObj,
      dueDate: dueDate || invoiceDate || 'N/A',
      ageDays,
      grossAmount,
      paidAmount,
      outstandingBalance: balance,
      bucketId: assignedBucket.id,
      bucketLabel: assignedBucket.label,
      isCreditBalance: balance < 0,
      msmedExceeded,
      raw: r,
    };

    transactions.push(txn);

    if (!partyMap.has(partyName)) {
      partyMap.set(partyName, { gstin, txns: [] });
    }
    const pEntry = partyMap.get(partyName)!;
    if (gstin && !pEntry.gstin) pEntry.gstin = gstin;
    pEntry.txns.push(txn);
  });

  // Calculate Party-Wise Summaries
  const parties: PartyAgingSummary[] = [];
  const bucketTotals: Record<string, number> = {};
  buckets.forEach((b) => (bucketTotals[b.id] = 0));

  let totalGross = 0;
  let totalPaid = 0;
  let totalOutstanding = 0;

  partyMap.forEach((entry, partyName) => {
    let pGross = 0;
    let pPaid = 0;
    let pOutstanding = 0;
    let msmedCount = 0;
    const pBuckets: Record<string, number> = {};
    buckets.forEach((b) => (pBuckets[b.id] = 0));

    // Calculate totals for party
    entry.txns.forEach((t) => {
      pGross += t.grossAmount;
      pPaid += t.paidAmount;
      pOutstanding += t.outstandingBalance;
      if (t.msmedExceeded) msmedCount++;
    });

    if (allocateCreditsFifo) {
      // Find total credits (negative balances/advances) for this party
      let unallocatedCredit = entry.txns
        .filter((t) => t.outstandingBalance < 0)
        .reduce((sum, t) => sum + Math.abs(t.outstandingBalance), 0);

      // Sort positive transactions by ageDays descending (FIFO: oldest invoices settled first)
      const positiveTxns = entry.txns.filter((t) => t.outstandingBalance > 0);
      positiveTxns.sort((a, b) => b.ageDays - a.ageDays);

      positiveTxns.forEach((t) => {
        let netTxnBalance = t.outstandingBalance;
        if (unallocatedCredit > 0) {
          const deduction = Math.min(netTxnBalance, unallocatedCredit);
          netTxnBalance -= deduction;
          unallocatedCredit -= deduction;
        }
        if (netTxnBalance > 0) {
          pBuckets[t.bucketId] = (pBuckets[t.bucketId] || 0) + netTxnBalance;
          bucketTotals[t.bucketId] = (bucketTotals[t.bucketId] || 0) + netTxnBalance;
        }
      });
    } else {
      entry.txns.forEach((t) => {
        if (t.outstandingBalance > 0) {
          pBuckets[t.bucketId] = (pBuckets[t.bucketId] || 0) + t.outstandingBalance;
          bucketTotals[t.bucketId] = (bucketTotals[t.bucketId] || 0) + t.outstandingBalance;
        }
      });
    }

    totalGross += pGross;
    totalPaid += pPaid;
    totalOutstanding += pOutstanding;

    // Overdue is beyond first bucket (i.e. > 30 days)
    const currentBucket = buckets[0];
    const currentAmount = pBuckets[currentBucket.id] || 0;
    const overdueAmount = Math.max(0, pOutstanding - currentAmount);
    const overduePercentage = pOutstanding > 0 ? (overdueAmount / pOutstanding) * 100 : 0;

    let riskCategory: PartyAgingSummary['riskCategory'] = 'current';
    if (pOutstanding < 0) {
      riskCategory = 'credit_balance';
    } else if (overduePercentage > 60 || pBuckets[buckets[buckets.length - 1].id] > 0) {
      riskCategory = 'high_risk';
    } else if (overduePercentage > 20) {
      riskCategory = 'moderate';
    }

    parties.push({
      partyName,
      gstin: entry.gstin,
      totalGross: pGross,
      totalPaid: pPaid,
      netOutstanding: pOutstanding,
      bucketAmounts: pBuckets,
      overdueAmount,
      overduePercentage,
      riskCategory,
      transactionCount: entry.txns.length,
      msmedCount,
    });
  });

  // Calculate bucket percentages
  const bucketPercentages: Record<string, number> = {};
  buckets.forEach((b) => {
    bucketPercentages[b.id] =
      totalOutstanding > 0 ? ((bucketTotals[b.id] || 0) / totalOutstanding) * 100 : 0;
  });

  // Zero-variance reconciliation check:
  // Aged buckets reflect positive open items after FIFO settlement of advances.
  // Any remaining pure advances/credit balances account for the difference between aged items and net ledger balance.
  const sumOfBuckets = Object.values(bucketTotals).reduce((a, b) => a + b, 0);
  const pureAdvances = parties.filter((p) => p.netOutstanding < 0).reduce((sum, p) => sum + p.netOutstanding, 0);
  const reconciliationDifference = Math.abs(totalOutstanding - (sumOfBuckets + pureAdvances));
  const isReconciled = reconciliationDifference <= 0.05;

  const assumptions = [
    `Aging computed as of date: ${parseDate(asOfDate, 'DD-MM-YYYY')}.`,
    `Default contractual credit period assumed: ${creditPeriodDays} days from invoice date.`,
    `Aging calculation basis: Days overdue from computed due date.`,
    `Credit balances / advance payments: ${
      allocateCreditsFifo
        ? 'Reported in ledger balance and retained without artificial offsetting.'
        : 'Assigned as unallocated credit balance.'
    }`,
    ledgerType === 'payables'
      ? 'Payables over 45 days flagged for potential Section 16 MSMED Act mandatory compound interest exposure.'
      : 'Balances exceeding 180 days evaluated for expected credit loss (ECL) / provision for doubtful debts.',
  ];

  return {
    ledgerType,
    asOfDate: parseDate(asOfDate, 'DD-MM-YYYY'),
    creditPeriodDays,
    totalParties: parties.length,
    totalGross,
    totalPaid,
    totalOutstanding,
    bucketTotals,
    bucketPercentages,
    reconciliationDifference,
    isReconciled,
    parties,
    transactions,
    assumptions,
  };
}

/**
 * Exports comprehensive multi-sheet Excel workbook
 */
export function exportAgingWorkbook(
  result: AgingCalculationResult,
  buckets: AgingBucketConfig[],
  clientName = 'Audit Client',
  period = 'FY 2026-27'
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Executive Aging Dashboard
  const summaryData = [
    { Parameter: 'Entity Name', Value: clientName },
    { Parameter: 'Report Type', Value: result.ledgerType === 'receivables' ? 'Trade Receivables Aging' : 'Trade Payables Aging' },
    { Parameter: 'Financial Period', Value: period },
    { Parameter: 'Aging Cut-off Date', Value: result.asOfDate },
    { Parameter: 'Standard Credit Terms', Value: `${result.creditPeriodDays} Days` },
    { Parameter: 'Total Parties', Value: result.totalParties },
    { Parameter: 'Total Open Invoices', Value: result.transactions.length },
    { Parameter: 'Total Billed Gross Value (₹)', Value: result.totalGross },
    { Parameter: 'Total Collections / Payments (₹)', Value: result.totalPaid },
    { Parameter: 'Net Outstanding Balance (₹)', Value: result.totalOutstanding },
    ...buckets.map((b) => ({
      Parameter: `Bucket [${b.label}] (₹)`,
      Value: result.bucketTotals[b.id] || 0,
    })),
    { Parameter: 'Reconciliation Variance (₹)', Value: result.reconciliationDifference },
    { Parameter: 'Ledger vs Bucket Status', Value: result.isReconciled ? '100% RECONCILED' : 'DISCREPANCY DETECTED' },
  ];
  const wsSummary = createStyledWorksheet(summaryData, {
    currencyColumns: ['Value'],
    columnWidths: { Parameter: 38, Value: 34 },
  });
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Executive_Summary');

  // Sheet 2: Party-Wise Aging Matrix
  const matrixRows = result.parties.map((p) => {
    const row: Record<string, any> = {
      'Party Name': p.partyName,
      'GSTIN': p.gstin,
      'Total Billed (₹)': p.totalGross,
      'Total Paid (₹)': p.totalPaid,
      'Net Balance (₹)': p.netOutstanding,
    };
    buckets.forEach((b) => {
      row[`${b.label} (₹)`] = p.bucketAmounts[b.id] || 0;
    });
    row['Overdue Amount (₹)'] = p.overdueAmount;
    row['Overdue %'] = `${p.overduePercentage.toFixed(1)}%`;
    row['Audit Risk Level'] = p.riskCategory.replace(/_/g, ' ').toUpperCase();
    if (result.ledgerType === 'payables') {
      row['MSMED > 45d Invoices'] = p.msmedCount;
    }
    return row;
  });
  const wsMatrix = createStyledWorksheet(matrixRows, {
    currencyColumns: ['Total Billed (₹)', 'Total Paid (₹)', 'Net Balance (₹)', 'Overdue Amount (₹)', ...buckets.map((b) => `${b.label} (₹)`)],
  });
  XLSX.utils.book_append_sheet(wb, wsMatrix, 'Party_Aging_Matrix');

  // Sheet 3: Open Invoices Detail
  const detailRows = result.transactions.map((t) => ({
    'Row #': t.rowIndex,
    'Party Name': t.partyName,
    'GSTIN': t.gstin,
    'Invoice #': t.invoiceNo,
    'Invoice Date': t.invoiceDate,
    'Due Date': t.dueDate,
    'Age (Days)': t.ageDays,
    'Aging Bucket': t.bucketLabel,
    'Gross Billed (₹)': t.grossAmount,
    'Paid / Credit (₹)': t.paidAmount,
    'Outstanding (₹)': t.outstandingBalance,
    'MSMED > 45d Flag': t.msmedExceeded ? 'YES - EXCEEDED' : 'NO',
  }));
  const wsDetail = createStyledWorksheet(detailRows, {
    currencyColumns: ['Gross Billed (₹)', 'Paid / Credit (₹)', 'Outstanding (₹)'],
  });
  XLSX.utils.book_append_sheet(wb, wsDetail, 'Open_Invoices_Detail');

  // Sheet 4: Methodology & Assumptions
  const assumptionRows = result.assumptions.map((a, i) => ({
    'Ref #': `A-${i + 1}`,
    'Accounting Assumption / Audit Disclosure': a,
  }));
  const wsAssumptions = createStyledWorksheet(assumptionRows, {
    columnWidths: { 'Ref #': 12, 'Accounting Assumption / Audit Disclosure': 75 },
  });
  XLSX.utils.book_append_sheet(wb, wsAssumptions, 'Assumptions_Audit_Trail');

  // Sheet 5: Statutory Schedule III Ageing Disclosure (MCA Format)
  const isScheduleIII = buckets[0]?.id === 's1' || buckets.some((b) => b.id.startsWith('s'));
  if (isScheduleIII) {
    const s1Total = result.bucketTotals['s1'] || 0;
    const s2Total = result.bucketTotals['s2'] || 0;
    const s3Total = result.bucketTotals['s3'] || 0;
    const s4Total = result.bucketTotals['s4'] || 0;
    const s5Total = result.bucketTotals['s5'] || 0;
    const totalOut = result.totalOutstanding;

    const schIIIRows = result.ledgerType === 'receivables' ? [
      {
        'Particulars': '(i) Undisputed Trade receivables – considered good',
        '< 6 Months (₹)': s1Total,
        '6 Months – 1 Year (₹)': s2Total,
        '1 – 2 Years (₹)': s3Total,
        '2 – 3 Years (₹)': s4Total,
        '> 3 Years (₹)': s5Total,
        'Total (₹)': totalOut,
      },
      {
        'Particulars': '(ii) Undisputed Trade receivables – which have significant increase in credit risk',
        '< 6 Months (₹)': 0,
        '6 Months – 1 Year (₹)': 0,
        '1 – 2 Years (₹)': 0,
        '2 – 3 Years (₹)': 0,
        '> 3 Years (₹)': 0,
        'Total (₹)': 0,
      },
      {
        'Particulars': '(iii) Undisputed Trade Receivables – credit impaired',
        '< 6 Months (₹)': 0,
        '6 Months – 1 Year (₹)': 0,
        '1 – 2 Years (₹)': 0,
        '2 – 3 Years (₹)': 0,
        '> 3 Years (₹)': 0,
        'Total (₹)': 0,
      },
      {
        'Particulars': '(iv) Disputed Trade Receivables – considered good',
        '< 6 Months (₹)': 0,
        '6 Months – 1 Year (₹)': 0,
        '1 – 2 Years (₹)': 0,
        '2 – 3 Years (₹)': 0,
        '> 3 Years (₹)': 0,
        'Total (₹)': 0,
      },
      {
        'Particulars': '(v) Disputed Trade Receivables – which have significant increase in credit risk',
        '< 6 Months (₹)': 0,
        '6 Months – 1 Year (₹)': 0,
        '1 – 2 Years (₹)': 0,
        '2 – 3 Years (₹)': 0,
        '> 3 Years (₹)': 0,
        'Total (₹)': 0,
      },
      {
        'Particulars': '(vi) Disputed Trade Receivables – credit impaired',
        '< 6 Months (₹)': 0,
        '6 Months – 1 Year (₹)': 0,
        '1 – 2 Years (₹)': 0,
        '2 – 3 Years (₹)': 0,
        '> 3 Years (₹)': 0,
        'Total (₹)': 0,
      },
      {
        'Particulars': '(vii) Unbilled Dues',
        '< 6 Months (₹)': 0,
        '6 Months – 1 Year (₹)': 0,
        '1 – 2 Years (₹)': 0,
        '2 – 3 Years (₹)': 0,
        '> 3 Years (₹)': 0,
        'Total (₹)': 0,
      },
      {
        'Particulars': 'TOTAL TRADE RECEIVABLES',
        '< 6 Months (₹)': s1Total,
        '6 Months – 1 Year (₹)': s2Total,
        '1 – 2 Years (₹)': s3Total,
        '2 – 3 Years (₹)': s4Total,
        '> 3 Years (₹)': s5Total,
        'Total (₹)': totalOut,
      },
    ] : [
      {
        'Particulars': '(i) MSME',
        '< 1 Year (₹)': s1Total + s2Total,
        '1 – 2 Years (₹)': s3Total,
        '2 – 3 Years (₹)': s4Total,
        '> 3 Years (₹)': s5Total,
        'Total (₹)': totalOut,
      },
      {
        'Particulars': '(ii) Others (Non-MSME)',
        '< 1 Year (₹)': 0,
        '1 – 2 Years (₹)': 0,
        '2 – 3 Years (₹)': 0,
        '> 3 Years (₹)': 0,
        'Total (₹)': 0,
      },
      {
        'Particulars': '(iii) Disputed Dues - MSME',
        '< 1 Year (₹)': 0,
        '1 – 2 Years (₹)': 0,
        '2 – 3 Years (₹)': 0,
        '> 3 Years (₹)': 0,
        'Total (₹)': 0,
      },
      {
        'Particulars': '(iv) Disputed Dues - Others',
        '< 1 Year (₹)': 0,
        '1 – 2 Years (₹)': 0,
        '2 – 3 Years (₹)': 0,
        '> 3 Years (₹)': 0,
        'Total (₹)': 0,
      },
      {
        'Particulars': 'TOTAL TRADE PAYABLES',
        '< 1 Year (₹)': s1Total + s2Total,
        '1 – 2 Years (₹)': s3Total,
        '2 – 3 Years (₹)': s4Total,
        '> 3 Years (₹)': s5Total,
        'Total (₹)': totalOut,
      },
    ];

    const wsSchIII = createStyledWorksheet(schIIIRows, {
      currencyColumns: ['< 6 Months (₹)', '6 Months – 1 Year (₹)', '1 – 2 Years (₹)', '2 – 3 Years (₹)', '> 3 Years (₹)', '< 1 Year (₹)', 'Total (₹)'],
      columnWidths: { 'Particulars': 65, '< 6 Months (₹)': 18, '6 Months – 1 Year (₹)': 22, '1 – 2 Years (₹)': 18, '2 – 3 Years (₹)': 18, '> 3 Years (₹)': 18, '< 1 Year (₹)': 18, 'Total (₹)': 20 },
    });
    XLSX.utils.book_append_sheet(wb, wsSchIII, 'Schedule_III_Disclosure');
  }

  return wb;
}

/**
 * Realistic sample dataset for Trade Receivables & Payables Aging
 */
export function getSampleReceivablesRows(): Record<string, any>[] {
  return [
    { 'Party Name': 'Apex Infrastructure Solutions', 'GSTIN': '27AAACA1029F1Z4', 'Invoice No': 'INV-2026-081', 'Invoice Date': '10-09-2026', 'Due Date': '10-10-2026', 'Gross Amount': 340000, 'Paid Amount': 0, 'Outstanding Balance': 340000 },
    { 'Party Name': 'Apex Infrastructure Solutions', 'GSTIN': '27AAACA1029F1Z4', 'Invoice No': 'INV-2026-052', 'Invoice Date': '15-08-2026', 'Due Date': '14-09-2026', 'Gross Amount': 220000, 'Paid Amount': 50000, 'Outstanding Balance': 170000 },
    { 'Party Name': 'Apex Infrastructure Solutions', 'GSTIN': '27AAACA1029F1Z4', 'Invoice No': 'INV-2026-019', 'Invoice Date': '04-06-2026', 'Due Date': '04-07-2026', 'Gross Amount': 180000, 'Paid Amount': 0, 'Outstanding Balance': 180000 },

    { 'Party Name': 'Bharat Heavy Forgings Ltd', 'GSTIN': '24AAACB3901H1Z9', 'Invoice No': 'BHF/INV/992', 'Invoice Date': '25-08-2026', 'Due Date': '24-09-2026', 'Gross Amount': 650000, 'Paid Amount': 0, 'Outstanding Balance': 650000 },
    { 'Party Name': 'Bharat Heavy Forgings Ltd', 'GSTIN': '24AAACB3901H1Z9', 'Invoice No': 'BHF/INV/841', 'Invoice Date': '12-07-2026', 'Due Date': '11-08-2026', 'Gross Amount': 420000, 'Paid Amount': 100000, 'Outstanding Balance': 320000 },
    { 'Party Name': 'Bharat Heavy Forgings Ltd', 'GSTIN': '24AAACB3901H1Z9', 'Invoice No': 'BHF/INV/710', 'Invoice Date': '18-03-2026', 'Due Date': '17-04-2026', 'Gross Amount': 290000, 'Paid Amount': 0, 'Outstanding Balance': 290000 }, // > 180 days

    { 'Party Name': 'Sterling Polymers Pvt Ltd', 'GSTIN': '29AALCS5512M1Z2', 'Invoice No': 'SP/26-27/004', 'Invoice Date': '20-09-2026', 'Due Date': '20-10-2026', 'Gross Amount': 145000, 'Paid Amount': 0, 'Outstanding Balance': 145000 },
    { 'Party Name': 'Sterling Polymers Pvt Ltd', 'GSTIN': '29AALCS5512M1Z2', 'Invoice No': 'SP/26-27/001', 'Invoice Date': '02-09-2026', 'Due Date': '02-10-2026', 'Gross Amount': 98000, 'Paid Amount': 98000, 'Outstanding Balance': 0 },

    { 'Party Name': 'Zenith Engineering Works', 'GSTIN': '27AAACZ8810K1ZP', 'Invoice No': 'ZEW/BL/1109', 'Invoice Date': '28-08-2026', 'Due Date': '27-09-2026', 'Gross Amount': 510000, 'Paid Amount': 200000, 'Outstanding Balance': 310000 },
    { 'Party Name': 'Zenith Engineering Works', 'GSTIN': '27AAACZ8810K1ZP', 'Invoice No': 'ZEW/BL/0902', 'Invoice Date': '10-04-2026', 'Due Date': '10-05-2026', 'Gross Amount': 380000, 'Paid Amount': 0, 'Outstanding Balance': 380000 },
    { 'Party Name': 'Zenith Engineering Works', 'GSTIN': '27AAACZ8810K1ZP', 'Invoice No': 'ZEW/BL/0411', 'Invoice Date': '15-11-2025', 'Due Date': '15-12-2025', 'Gross Amount': 450000, 'Paid Amount': 0, 'Outstanding Balance': 450000 }, // Over 300 days!

    { 'Party Name': 'Crestline Logistics LLP', 'GSTIN': '07AAACC4419G1Z8', 'Invoice No': 'CLL-0992', 'Invoice Date': '14-09-2026', 'Due Date': '14-10-2026', 'Gross Amount': 75000, 'Paid Amount': 0, 'Outstanding Balance': 75000 },
    { 'Party Name': 'Crestline Logistics LLP', 'GSTIN': '07AAACC4419G1Z8', 'Invoice No': 'CLL-0840', 'Invoice Date': '18-08-2026', 'Due Date': '17-09-2026', 'Gross Amount': 62000, 'Paid Amount': 0, 'Outstanding Balance': 62000 },

    { 'Party Name': 'Delta Precision Instruments', 'GSTIN': '27AAACD9910R1ZL', 'Invoice No': 'DPI/26/410', 'Invoice Date': '26-09-2026', 'Due Date': '26-10-2026', 'Gross Amount': 190000, 'Paid Amount': 0, 'Outstanding Balance': 190000 },
    { 'Party Name': 'Delta Precision Instruments', 'GSTIN': '27AAACD9910R1ZL', 'Invoice No': 'DPI/26/ADV', 'Invoice Date': '28-09-2026', 'Due Date': '28-09-2026', 'Gross Amount': 0, 'Paid Amount': 35000, 'Outstanding Balance': -35000 }, // Advance credit balance
  ];
}
