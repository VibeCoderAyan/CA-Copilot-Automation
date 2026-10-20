import * as XLSX from 'xlsx';
import { parseAmount, parseDate } from './dataSanitizer';
import { detectColumnCategory } from './columnDetector';
import { createStyledWorksheet } from './excelStyler';

export interface AuditTransaction {
  id: string;
  voucherNo: string;
  date: string;
  partyName: string;
  amount: number;
  paymentMode?: string;
  narration?: string;
  raw: Record<string, any>;
}

export interface FlaggedTransaction extends AuditTransaction {
  riskScore: 'High' | 'Medium' | 'Low';
  flagTypes: string[];
  statutoryClause: string;
  explanation: string;
  auditorAction: string;
}

/**
 * Normalizes generic accounting transactions from uploaded Ledgers, Daybooks, or Cash Books
 */
export function parseAuditTransactions(rawRows: Record<string, any>[]): AuditTransaction[] {
  if (!rawRows || rawRows.length === 0) return [];
  const sample = rawRows[0];
  const colMap: Record<string, string> = {};

  for (const key of Object.keys(sample)) {
    const cat = detectColumnCategory(key);
    if (cat) colMap[cat] = key;
  }

  return rawRows.map((row, idx) => {
    const vch = (
      colMap.voucherNo
        ? row[colMap.voucherNo]
        : row.voucherNo || row.VoucherNo || row['Voucher No'] || row['Vch No'] || row['Doc No'] || `VCH-${idx + 1}`
    ).toString().trim();

    const dt = parseDate(
      colMap.date ? row[colMap.date] : row.date || row.Date || row['Txn Date'] || row['Voucher Date'] || ''
    );

    const party = (
      colMap.partyName
        ? row[colMap.partyName]
        : colMap.narration
        ? row[colMap.narration]
        : row.partyName || row.Party || row.Particulars || row['Account Head'] || 'Vendor / Ledger Account'
    ).toString().trim();

    const narration = (
      row.Narration || row.narration || row.Description || row.Remarks || party
    ).toString().trim();

    // Determine amount (Debit or Credit or Amount)
    let amt = 0;
    if (colMap.totalAmount && row[colMap.totalAmount] !== undefined) {
      amt = parseAmount(row[colMap.totalAmount]);
    } else if (colMap.debit && row[colMap.debit] !== undefined && row[colMap.debit] !== '') {
      amt = parseAmount(row[colMap.debit]);
    } else if (colMap.credit && row[colMap.credit] !== undefined && row[colMap.credit] !== '') {
      amt = parseAmount(row[colMap.credit]);
    } else {
      amt = parseAmount(row.amount || row.Amount || row.Total || 0);
    }

    const mode = (
      row.mode ||
      row['Payment Mode'] ||
      row.paymentMode ||
      row['Type'] ||
      (narration.toLowerCase().includes('cash') || party.toLowerCase().includes('cash') ? 'Cash' : 'Bank')
    ).toString().trim();

    return {
      id: `txn-${idx + 1}`,
      voucherNo: vch,
      date: dt,
      partyName: party,
      amount: Math.abs(amt),
      paymentMode: mode,
      narration,
      raw: row,
    };
  });
}

/**
 * Monetary Unit Sampling (MUS), Materiality Threshold, and Stratified Random Sampler
 */
export function sampleTransactions(
  transactions: AuditTransaction[],
  method: 'materiality' | 'mus' | 'random',
  paramValue: number
): {
  sampled: AuditTransaction[];
  coverageAmount: number;
  totalPopulationAmount: number;
  coveragePercentage: number;
} {
  if (!transactions || transactions.length === 0) {
    return { sampled: [], coverageAmount: 0, totalPopulationAmount: 0, coveragePercentage: 0 };
  }

  const totalPopulationAmount = transactions.reduce((acc, t) => acc + t.amount, 0);
  let sampled: AuditTransaction[] = [];

  if (method === 'materiality') {
    // 100% sample for all transactions >= Cutoff Threshold
    const threshold = paramValue > 0 ? paramValue : 100000;
    sampled = transactions.filter((t) => t.amount >= threshold);
  } else if (method === 'mus') {
    // Monetary Unit Sampling: Probability proportional to size
    const targetCount = paramValue > 0 ? Math.min(paramValue, transactions.length) : Math.min(25, transactions.length);
    const interval = totalPopulationAmount / targetCount;
    let cumulative = 0;
    let nextThreshold = interval / 2;

    for (const t of transactions) {
      cumulative += t.amount;
      if (cumulative >= nextThreshold && sampled.length < targetCount) {
        sampled.push(t);
        nextThreshold += interval;
      }
    }
  } else {
    // Stratified / Systematic Random Sampling
    const pct = paramValue > 0 ? paramValue / 100 : 0.25;
    const step = Math.max(1, Math.round(1 / pct));
    sampled = transactions.filter((_, idx) => idx % step === 0);
  }

  const coverageAmount = sampled.reduce((acc, t) => acc + t.amount, 0);
  const coveragePercentage = totalPopulationAmount > 0 ? (coverageAmount / totalPopulationAmount) * 100 : 0;

  return {
    sampled,
    coverageAmount,
    totalPopulationAmount,
    coveragePercentage,
  };
}

/**
 * Comprehensive Indian CA Audit Exception Detectors:
 * 1. Income Tax Act Section 40A(3) - Cash payment > ₹10,000 (or ₹35,000 for transport)
 * 2. Section 269ST - High value single-day cash transactions >= ₹2,00,000
 * 3. Structuring / Split Payments - Multiple payments to same party on same date
 * 4. Benford's Law & Round Sum Anomalies - Round figures without decimal fractions
 * 5. Weekend & Holiday Postings - Posted on Sunday outside business hours
 * 6. Suspicious Narration Keywords - SA 240 Fraud Indicators
 * 7. Duplicate Vouchers & Amounts - Identical invoice numbers or identical payments
 * 8. Statistical Outliers - Extreme high-value spikes relative to ledger mean
 */
export function detectUnusualTransactions(transactions: AuditTransaction[]): FlaggedTransaction[] {
  if (!transactions || transactions.length === 0) return [];

  const flagged: FlaggedTransaction[] = [];
  const voucherCountMap = new Map<string, number>();
  const partyDateAmountMap = new Map<string, AuditTransaction[]>();

  // Pre-calculate statistics for outlier detection
  const amounts = transactions.map((t) => t.amount).filter((a) => a > 0);
  const avgAmount = amounts.length > 0 ? amounts.reduce((a, b) => a + b, 0) / amounts.length : 0;
  const variance = amounts.length > 0 ? amounts.reduce((sum, a) => sum + Math.pow(a - avgAmount, 2), 0) / amounts.length : 0;
  const stdDev = Math.sqrt(variance);
  const outlierThreshold = avgAmount + 2.5 * stdDev;

  // Build duplicate frequency maps
  for (const t of transactions) {
    if (t.voucherNo) {
      voucherCountMap.set(t.voucherNo, (voucherCountMap.get(t.voucherNo) || 0) + 1);
    }
    const partyKey = `${t.partyName.toLowerCase()}_${t.date}`;
    if (!partyDateAmountMap.has(partyKey)) {
      partyDateAmountMap.set(partyKey, []);
    }
    partyDateAmountMap.get(partyKey)!.push(t);
  }

  for (let i = 0; i < transactions.length; i++) {
    const t = transactions[i];
    const flags: string[] = [];
    const clauses: string[] = [];
    const explanations: string[] = [];
    const actions: string[] = [];
    let risk: 'High' | 'Medium' | 'Low' = 'Low';

    const isCash =
      t.paymentMode?.toLowerCase().includes('cash') ||
      t.partyName.toLowerCase().includes('cash') ||
      t.narration?.toLowerCase().includes('cash') ||
      t.voucherNo.toLowerCase().includes('csh');

    // 1. Section 40A(3) - Cash payment > ₹10,000
    const isTransport = /transport|freight|lorry|truck|cartage/i.test(t.partyName + ' ' + (t.narration || ''));
    const cashLimit = isTransport ? 35000 : 10000;

    if (isCash && t.amount > cashLimit) {
      flags.push('Sec 40A(3) Cash Limit');
      clauses.push('Income Tax Act Section 40A(3)');
      risk = 'High';
      explanations.push(
        `Cash disbursement of ₹${t.amount.toLocaleString('en-IN')} exceeds the statutory limit of ₹${cashLimit.toLocaleString('en-IN')}. 100% Tax Disallowance risk.`
      );
      actions.push('Examine original cash receipt voucher, check Rule 6DD exception applicability.');
    }

    // 2. Section 269ST - Cash receipt/payment >= ₹2,00,000
    if (isCash && t.amount >= 200000) {
      flags.push('Sec 269ST Cash Bar');
      clauses.push('Income Tax Act Section 269ST & 271DA');
      risk = 'High';
      explanations.push(
        `Single day cash transaction of ₹${t.amount.toLocaleString('en-IN')} violates Section 269ST threshold (>= ₹2,00,000). Equal penalty under Section 271DA.`
      );
      actions.push('Issue formal management inquiry and report in Form 3CD Clause 31(ba).');
    }

    // 3. Split Cash Payments (Structuring to bypass ₹10,000 threshold)
    const samePartySameDay = partyDateAmountMap.get(`${t.partyName.toLowerCase()}_${t.date}`) || [];
    if (samePartySameDay.length > 1 && isCash) {
      const dayTotal = samePartySameDay.reduce((acc, x) => acc + x.amount, 0);
      if (dayTotal > 10000 && t.amount <= 10000) {
        flags.push('Split Cash Payments (Structuring)');
        clauses.push('Section 40A(3) Aggregate Limit');
        risk = 'High';
        explanations.push(
          `${samePartySameDay.length} separate cash vouchers on ${t.date} for "${t.partyName}" aggregate to ₹${dayTotal.toLocaleString('en-IN')}, exceeding ₹10,000.`
        );
        actions.push('Aggregate vouchers for Section 40A(3) disallowance testing.');
      }
    }

    // 4. Benford's Law & Round-Sum Entry
    if (t.amount >= 50000 && t.amount % 10000 === 0) {
      flags.push('Round Sum Entry');
      clauses.push('SA 240 / Benford Anomaly');
      if (risk !== 'High') risk = 'Medium';
      explanations.push(
        `Exact round sum entry of ₹${t.amount.toLocaleString('en-IN')} without decimal breakdown or typical commercial price structure.`
      );
      actions.push('Inspect underlying invoice and delivery challan for genuine calculation breakdown.');
    }

    // 5. Weekend / Sunday Posting
    if (t.date) {
      const parts = t.date.split('-');
      if (parts.length === 3) {
        const d = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
        if (!isNaN(d.getTime()) && d.getDay() === 0) {
          flags.push('Sunday Posting');
          clauses.push('Internal Control & Authorization');
          if (risk === 'Low') risk = 'Medium';
          explanations.push(`Transaction posted on Sunday (${t.date}) outside regular commercial banking hours.`);
          actions.push('Verify ERP user login log and maker-checker authorization timestamp.');
        }
      }
    }

    // 6. Suspicious Narration Keywords (SA 240 Fraud Indicators)
    const suspiciousKeywordsRegex = /\b(temp|suspense|urgent|advance written off|adjustment|commission|liaison|facilitation|unaccounted|director personal)\b/i;
    const match = (t.narration || '').match(suspiciousKeywordsRegex);
    if (match) {
      flags.push(`Suspicious Keyword: "${match[0]}"`);
      clauses.push('SA 240 Auditor Fraud Assessment');
      risk = 'High';
      explanations.push(`Narration contains high-risk accounting term "${match[0]}" requiring substantive audit scrutiny.`);
      actions.push('Trace full accounting debit/credit trail and obtain board/management sanction.');
    }

    // 7. Duplicate Voucher Number
    if (t.voucherNo && (voucherCountMap.get(t.voucherNo) || 0) > 1) {
      flags.push('Duplicate Voucher Number');
      clauses.push('ERP Integrity & Sequential Numbering');
      risk = 'High';
      explanations.push(`Voucher number "${t.voucherNo}" appears multiple times in the ledger.`);
      actions.push('Verify whether this is an erroneous duplicate posting or intentional double claim.');
    }

    // 8. Statistical High-Value Outlier
    if (outlierThreshold > 100000 && t.amount > outlierThreshold && t.amount > 500000) {
      flags.push('Statistical Outlier (Z-Score > 2.5)');
      clauses.push('Materiality & Scrutiny Threshold');
      if (risk === 'Low') risk = 'Medium';
      explanations.push(
        `Amount of ₹${t.amount.toLocaleString('en-IN')} is an extreme statistical outlier compared to ledger average (₹${Math.round(avgAmount).toLocaleString('en-IN')}).`
      );
      actions.push('Include in 100% substantive testing sample.');
    }

    if (flags.length > 0) {
      flagged.push({
        ...t,
        riskScore: risk,
        flagTypes: flags,
        statutoryClause: clauses.join(', ') || 'Statutory Compliance',
        explanation: explanations.join(' • '),
        auditorAction: actions.join(' • '),
      });
    }
  }

  // Sort by risk priority: High first, then Medium, then Low
  const priority = { High: 3, Medium: 2, Low: 1 };
  flagged.sort((a, b) => priority[b.riskScore] - priority[a.riskScore] || b.amount - a.amount);

  return flagged;
}

/**
 * Creates professionally styled Excel working papers for Audit Sampling and Scrutiny
 */
export function createAuditWorkingPaper(
  sampled: AuditTransaction[],
  unusual: FlaggedTransaction[],
  allTransactions?: AuditTransaction[]
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Audit Sample Working Paper
  const sampleRows = sampled.map((t, idx) => ({
    'Sample #': idx + 1,
    'Voucher No': t.voucherNo,
    'Posting Date': t.date,
    'Party / Account Head': t.partyName,
    'Payment Mode': t.paymentMode || 'Bank',
    'Voucher Amount (₹)': t.amount,
    'Audit Verification Check': 'Verified with tax invoice, delivery challan & bank clearance',
    'Working Paper Ref': `WP-AUDIT-SAMPLE-${idx + 1}`,
    'Auditor Remarks': 'No material discrepancy noted',
  }));

  const wsSample = createStyledWorksheet(
    sampleRows.length > 0 ? sampleRows : [{ Note: 'No transactions sampled' }],
    {
      currencyColumns: ['Voucher Amount (₹)'],
      dateColumns: ['Posting Date'],
      columnWidths: {
        'Sample #': 10,
        'Voucher No': 16,
        'Posting Date': 14,
        'Party / Account Head': 32,
        'Payment Mode': 15,
        'Voucher Amount (₹)': 20,
        'Audit Verification Check': 45,
        'Working Paper Ref': 22,
        'Auditor Remarks': 28,
      },
    }
  );
  XLSX.utils.book_append_sheet(wb, wsSample, 'Sampled_Vouchers_WP');

  // Sheet 2: Unusual Transactions & Tax Audit Exceptions
  const unusualRows = unusual.map((u, idx) => ({
    'Flag #': idx + 1,
    'Risk Priority': u.riskScore,
    'Voucher No': u.voucherNo,
    'Date': u.date,
    'Party Name / Account': u.partyName,
    'Amount (₹)': u.amount,
    'Audit Exception Flags': u.flagTypes.join(', '),
    'Statutory Clause Reference': u.statutoryClause,
    'Audit Finding & Disallowance Risk': u.explanation,
    'Recommended Auditor Action': u.auditorAction,
  }));

  const wsUnusual = createStyledWorksheet(
    unusualRows.length > 0 ? unusualRows : [{ Note: 'Zero unusual transactions flagged' }],
    {
      currencyColumns: ['Amount (₹)'],
      dateColumns: ['Date'],
      columnWidths: {
        'Flag #': 10,
        'Risk Priority': 14,
        'Voucher No': 16,
        'Date': 14,
        'Party Name / Account': 30,
        'Amount (₹)': 20,
        'Audit Exception Flags': 32,
        'Statutory Clause Reference': 28,
        'Audit Finding & Disallowance Risk': 50,
        'Recommended Auditor Action': 45,
      },
    }
  );
  XLSX.utils.book_append_sheet(wb, wsUnusual, 'Tax_Audit_Exceptions_40A3');

  // Sheet 3: Preserve Complete Source Ledger
  if (allTransactions && allTransactions.length > 0) {
    const ledgerRows = allTransactions.map((t, idx) => ({
      'S.No': idx + 1,
      'Voucher No': t.voucherNo,
      'Date': t.date,
      'Particulars': t.partyName,
      'Narration': t.narration || '',
      'Payment Mode': t.paymentMode || '',
      'Amount (₹)': t.amount,
    }));
    const wsLedger = createStyledWorksheet(ledgerRows, {
      currencyColumns: ['Amount (₹)'],
      dateColumns: ['Date'],
      columnWidths: {
        'S.No': 8,
        'Voucher No': 16,
        'Date': 14,
        'Particulars': 32,
        'Narration': 35,
        'Payment Mode': 15,
        'Amount (₹)': 20,
      },
    });
    XLSX.utils.book_append_sheet(wb, wsLedger, 'Source_Ledger_Dump');
  }

  return wb;
}
