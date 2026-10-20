import * as XLSX from 'xlsx';
import { parseAmount, parseDate, isValidPan } from './dataSanitizer';
import { createStyledWorksheet } from './excelStyler';

/**
 * FORM 3CD TAX AUDIT SCRUTINY & ANNEXURE GENERATION ENGINE (Section 44AB)
 *
 * Implements automated audit scrutiny for key reporting clauses:
 * 1. Clause 21(d) / Section 40A(3) & 40A(3A): Cash payments exceeding ₹10,000 / ₹35,000
 * 2. Clause 44: Break-up of total expenditure into GST registered vs unregistered entities
 * 3. Clause 34: Chapter XVII-B compliance (TDS deduction, short deduction, Section 40(a)(ia) disallowance)
 * 4. Clause 21(a): Inadmissible expenses (penalties, personal expenses, fines)
 * 5. Clause 22: MSME interest disallowance under Section 23 of MSMED Act
 */

export interface VoucherEntry {
  id: string;
  voucherNo: string;
  date: string;
  accountName: string;
  category: string;
  narration: string;
  paymentMode: 'Cash' | 'Bank' | 'Bearer Cheque' | 'Electronic' | 'Journal';
  payeeName: string;
  payeePan?: string;
  payeeGstin?: string;
  amount: number;
  isTransporter?: boolean;
  rule6DdException?: boolean;
  rule6DdReason?: string;
  isGstExempt?: boolean;
  isCompositionSupplier?: boolean;
  tdsSection?: string;
  tdsRateApplied?: number;
  tdsDeducted?: number;
  raw?: Record<string, any>;
}

export interface CashViolationRecord {
  id: string;
  date: string;
  payeeName: string;
  voucherNumbers: string[];
  singlePaymentMax: number;
  dailyAggregateAmount: number;
  thresholdLimit: number; // 10000 or 35000
  disallowableAmount: number;
  isTransporter: boolean;
  isRule6DdExempt: boolean;
  exemptionReason: string;
  vouchers: VoucherEntry[];
}

export interface Clause44Row {
  sNo: number;
  headOfExpenditure: string;
  totalExpenditure: number;
  exemptSupplies: number;
  compositionEntities: number;
  otherRegisteredEntities: number;
  totalRegisteredExpenditure: number;
  unregisteredEntities: number;
  unregisteredPercent: number;
  remarks: string;
}

export interface Clause34SectionSummary {
  section: string;
  sectionName: string;
  totalExpenditureCovered: number;
  amountEligibleForTds: number;
  amountDeductedAtStandardRate: number;
  amountDeductedAtLowerRate: number;
  amountNotDeducted: number;
  tdsAmountRequired: number;
  tdsAmountDeducted: number;
  tdsAmountDeposited: number;
  shortDeductionAmount: number;
  disallowance40aIa: number; // 30% of expenditure where TDS not deducted
  estimatedInterest201: number;
}

export interface InadmissibleExpenseRecord {
  id: string;
  date: string;
  voucherNo: string;
  accountName: string;
  narration: string;
  amount: number;
  disallowanceType: 'Penalty / Fine (Sec 37)' | 'Personal Expenditure' | 'Capital debited to P&L' | 'MSME Interest (Sec 23 MSMED)';
  clauseReference: 'Clause 21(a)' | 'Clause 22';
  confidenceScore: number;
  flaggedKeywords: string[];
}

export interface Form3cdAuditSummary {
  clientName: string;
  assessmentYear: string;
  financialYear: string;
  totalExpenditureScrutinized: number;
  
  // Clause 21(d) / Sec 40A(3)
  clause40a3ViolationsCount: number;
  clause40a3TotalDisallowable: number;
  clause40a3Rule6DdExemptCount: number;
  
  // Clause 44
  totalClause44Expenditure: number;
  totalRegisteredExpenditure: number;
  totalUnregisteredExpenditure: number;
  unregisteredExpenditureRatio: number; // %
  
  // Clause 34
  clause34TotalCoveredExpenditure: number;
  clause34TotalTdsRequired: number;
  clause34TotalTdsDeducted: number;
  clause34ShortDeduction: number;
  clause34Disallowance40aIa: number; // 30% addition
  
  // Clause 21(a) & 22
  clause21aPersonalAndPenalties: number;
  clause22MsmeInterest: number;
  
  // Total additions to taxable income
  totalTaxAdditions: number;
  
  // Scrutinized records
  cashViolations: CashViolationRecord[];
  clause44Breakdown: Clause44Row[];
  clause34Summary: Clause34SectionSummary[];
  inadmissibleExpenses: InadmissibleExpenseRecord[];
}

// Common penalty & personal expense keywords for scanning
const PENALTY_KEYWORDS = ['penalty', 'penal', 'fine', 'traffic', 'compounding', 'late fee gst', 'punitive', 'seizure', 'contravention'];
const PERSONAL_KEYWORDS = ['personal', 'director medical', 'partner personal', 'domestic', 'club membership', 'household', 'private'];
const CAPITAL_KEYWORDS = ['renovation civil', 'air conditioner capital', 'laptop asset', 'building addition', 'plant installation'];

/**
 * Standard TDS Rates for Chapter XVII-B Scrutiny
 */
export const STATUTORY_TDS_RATES: Record<string, { name: string; rate: number; threshold: number }> = {
  '194C': { name: 'Payments to Contractors / Subcontractors', rate: 2.0, threshold: 100000 },
  '194J': { name: 'Fees for Professional / Technical Services', rate: 10.0, threshold: 30000 },
  '194I': { name: 'Rent on Land, Building & Furniture', rate: 10.0, threshold: 240000 },
  '194H': { name: 'Commission & Brokerage', rate: 2.0, threshold: 15000 },
  '194A': { name: 'Interest other than Interest on Securities', rate: 10.0, threshold: 5000 },
  '194Q': { name: 'Purchase of Goods (> ₹50 Lakhs)', rate: 0.1, threshold: 5000000 },
};

/**
 * Auto-detect columns in uploaded expense/ledger registers
 */
export function detectForm3cdColumns(headers: string[]): Record<string, string> {
  const norm = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, '');
  const findCol = (keys: string[]): string => {
    for (const h of headers) {
      const nh = norm(h);
      if (keys.some((k) => nh.includes(norm(k)))) return h;
    }
    return '';
  };

  return {
    voucherNoCol: findCol(['voucherno', 'vchno', 'docno', 'refno', 'invoice', 'transid']),
    dateCol: findCol(['date', 'vchdate', 'txndate', 'bookingdate']),
    accountCol: findCol(['account', 'ledger', 'particulars', 'head', 'expensehead']),
    payeeCol: findCol(['payee', 'vendor', 'party', 'supplier', 'name', 'beneficiary']),
    amountCol: findCol(['amount', 'debit', 'total', 'taxablevalue', 'value']),
    modeCol: findCol(['mode', 'paymentmode', 'type', 'instrument', 'paymenttype']),
    narrationCol: findCol(['narration', 'remarks', 'description', 'particulars']),
    panCol: findCol(['pan', 'panno', 'payeepan']),
    gstinCol: findCol(['gstin', 'gst', 'gstno', 'vendorin']),
  };
}

/**
 * Executes full Form 3CD Audit Scrutiny
 */
export function scrutinizeForm3cdRecords(
  rawRows: Record<string, any>[],
  mapping: Record<string, string>,
  clientName: string = 'Apex Technologies Pvt Ltd',
  financialYear: string = '2025-26',
  assessmentYear: string = '2026-27'
): Form3cdAuditSummary {
  const vouchers: VoucherEntry[] = rawRows.map((r, idx) => {
    const voucherNo = String(r[mapping.voucherNoCol] || `VCH-${idx + 1}`).trim();
    const date = parseDate(String(r[mapping.dateCol] || ''));
    const accountName = String(r[mapping.accountCol] || 'General Expenses').trim();
    const payeeName = String(r[mapping.payeeCol] || 'Undisclosed Payee').trim();
    const amount = Math.abs(parseAmount(r[mapping.amountCol]));
    const narration = String(r[mapping.narrationCol] || '').trim();
    const rawMode = String(r[mapping.modeCol] || '').toLowerCase();
    const payeePan = String(r[mapping.panCol] || '').trim().toUpperCase();
    const payeeGstin = String(r[mapping.gstinCol] || '').trim().toUpperCase();

    // Determine payment mode
    let paymentMode: VoucherEntry['paymentMode'] = 'Electronic';
    if (rawMode.includes('cash') || narration.toLowerCase().includes('cash paid') || narration.toLowerCase().includes('petty cash')) {
      paymentMode = 'Cash';
    } else if (rawMode.includes('bearer') || rawMode.includes('cheque')) {
      paymentMode = 'Bearer Cheque';
    } else if (rawMode.includes('journal') || rawMode.includes('jv')) {
      paymentMode = 'Journal';
    } else if (rawMode.includes('bank') || rawMode.includes('neft') || rawMode.includes('rtgs')) {
      paymentMode = 'Bank';
    }

    // Detect transporter keywords (Eligible for ₹35,000 threshold under Section 40A(3))
    const isTransporter =
      accountName.toLowerCase().includes('freight') ||
      accountName.toLowerCase().includes('transport') ||
      accountName.toLowerCase().includes('carriage') ||
      narration.toLowerCase().includes('truck') ||
      narration.toLowerCase().includes('lorry') ||
      narration.toLowerCase().includes('transporter');

    return {
      id: `vch_${idx}_${voucherNo}`,
      voucherNo,
      date,
      accountName,
      category: accountName,
      narration,
      paymentMode,
      payeeName,
      payeePan: isValidPan(payeePan) ? payeePan : undefined,
      payeeGstin: payeeGstin.length === 15 ? payeeGstin : undefined,
      amount,
      isTransporter,
      raw: r,
    };
  });

  // =========================================================================
  // 1. CLAUSE 21(d) / SECTION 40A(3) CASH PAYMENTS SCRUTINY
  // =========================================================================
  const cashPayments = vouchers.filter((v) => v.paymentMode === 'Cash' || v.paymentMode === 'Bearer Cheque');
  
  // Group by Payee + Date to detect daily aggregate breaches
  const dailyPayeeMap: Record<string, VoucherEntry[]> = {};
  for (const cp of cashPayments) {
    const key = `${cp.date.trim()}||${cp.payeeName.trim().toLowerCase()}`;
    if (!dailyPayeeMap[key]) {
      dailyPayeeMap[key] = [];
    }
    dailyPayeeMap[key].push(cp);
  }

  const cashViolations: CashViolationRecord[] = [];
  let totalDisallowable40a3 = 0;
  let rule6DdExemptCount = 0;

  for (const [key, vList] of Object.entries(dailyPayeeMap)) {
    const [date, payeeLower] = key.split('||');
    const totalDailyAmount = vList.reduce((acc, v) => acc + v.amount, 0);
    const maxSinglePayment = Math.max(...vList.map((v) => v.amount));
    const isTransporter = vList.some((v) => v.isTransporter);
    const thresholdLimit = isTransporter ? 35000 : 10000;

    if (totalDailyAmount > thresholdLimit) {
      // Check for Rule 6DD exception in narration
      const allNarrations = vList.map((v) => v.narration.toLowerCase()).join(' ');
      let isRule6Dd = false;
      let exemptionReason = '';

      if (allNarrations.includes('bank holiday') || allNarrations.includes('holiday')) {
        isRule6Dd = true;
        exemptionReason = 'Rule 6DD(i): Payment made on bank holiday';
      } else if (allNarrations.includes('cultivator') || allNarrations.includes('agricultural') || allNarrations.includes('farm')) {
        isRule6Dd = true;
        exemptionReason = 'Rule 6DD(e): Payment for agricultural produce to cultivator';
      } else if (allNarrations.includes('terminal benefit') || allNarrations.includes('gratuity') || allNarrations.includes('retrenchment')) {
        isRule6Dd = true;
        exemptionReason = 'Rule 6DD(f): Terminal benefit / retrenchment payment to employee';
      } else if (allNarrations.includes('govt') || allNarrations.includes('rbi') || allNarrations.includes('railway')) {
        isRule6Dd = true;
        exemptionReason = 'Rule 6DD(b): Payment to Government / Railway authorities';
      }

      if (isRule6Dd) {
        rule6DdExemptCount++;
      } else {
        totalDisallowable40a3 += totalDailyAmount;
      }

      cashViolations.push({
        id: `viol_${date}_${payeeLower.replace(/\s+/g, '_')}`,
        date,
        payeeName: vList[0].payeeName,
        voucherNumbers: vList.map((v) => v.voucherNo),
        singlePaymentMax: maxSinglePayment,
        dailyAggregateAmount: totalDailyAmount,
        thresholdLimit,
        disallowableAmount: isRule6Dd ? 0 : totalDailyAmount,
        isTransporter,
        isRule6DdExempt: isRule6Dd,
        exemptionReason: exemptionReason || 'Section 40A(3) Violation (No statutory exemption noted)',
        vouchers: vList,
      });
    }
  }

  // =========================================================================
  // 2. CLAUSE 44 BREAK-UP OF TOTAL EXPENDITURE (GST REGISTERED VS UNREGISTERED)
  // =========================================================================
  // Exclude non-expenditure / outside-GST items like Depreciation, Salary, Bad Debts, Direct Taxes
  const nonClause44Accounts = ['depreciation', 'salary', 'salaries', 'wages', 'income tax', 'advance tax', 'interest on income tax'];
  const eligibleVouchers = vouchers.filter((v) => !nonClause44Accounts.some((na) => v.accountName.toLowerCase().includes(na)));

  // Group by expenditure head
  const headMap: Record<string, {
    total: number;
    exempt: number;
    composition: number;
    regularRegistered: number;
    unregistered: number;
  }> = {};

  for (const v of eligibleVouchers) {
    const head = v.accountName;
    if (!headMap[head]) {
      headMap[head] = { total: 0, exempt: 0, composition: 0, regularRegistered: 0, unregistered: 0 };
    }

    headMap[head].total += v.amount;

    const nar = v.narration.toLowerCase();
    const hasGstin = Boolean(v.payeeGstin && v.payeeGstin.length === 15);

    if (hasGstin) {
      if (nar.includes('exempt') || nar.includes('nil rated') || v.isGstExempt) {
        headMap[head].exempt += v.amount;
      } else if (nar.includes('composition') || nar.includes('sec 10') || v.isCompositionSupplier) {
        headMap[head].composition += v.amount;
      } else {
        headMap[head].regularRegistered += v.amount;
      }
    } else {
      // Unregistered vendor
      headMap[head].unregistered += v.amount;
    }
  }

  const clause44Breakdown: Clause44Row[] = Object.entries(headMap).map(([head, data], idx) => {
    const totalReg = data.exempt + data.composition + data.regularRegistered;
    const unregPct = data.total > 0 ? (data.unregistered / data.total) * 100 : 0;
    return {
      sNo: idx + 1,
      headOfExpenditure: head,
      totalExpenditure: data.total,
      exemptSupplies: data.exempt,
      compositionEntities: data.composition,
      otherRegisteredEntities: data.regularRegistered,
      totalRegisteredExpenditure: totalReg,
      unregisteredEntities: data.unregistered,
      unregisteredPercent: Number(unregPct.toFixed(2)),
      remarks: unregPct > 40 ? 'High Unregistered Procurement (> 40%) - Scrutiny Risk' : 'Normal compliance',
    };
  });

  const totalClause44Expenditure = clause44Breakdown.reduce((acc, r) => acc + r.totalExpenditure, 0);
  const totalRegisteredExpenditure = clause44Breakdown.reduce((acc, r) => acc + r.totalRegisteredExpenditure, 0);
  const totalUnregisteredExpenditure = clause44Breakdown.reduce((acc, r) => acc + r.unregisteredEntities, 0);
  const unregisteredExpenditureRatio = totalClause44Expenditure > 0 ? (totalUnregisteredExpenditure / totalClause44Expenditure) * 100 : 0;

  // =========================================================================
  // 3. CLAUSE 34 CHAPTER XVII-B (TDS/TCS COMPLIANCE)
  // =========================================================================
  const clause34Map: Record<string, {
    totalExp: number;
    eligibleForTds: number;
    atStdRate: number;
    atLowRate: number;
    notDeducted: number;
    tdsReq: number;
    tdsDed: number;
    tdsDep: number;
  }> = {};

  // Initialize standard TDS sections
  for (const sec of Object.keys(STATUTORY_TDS_RATES)) {
    clause34Map[sec] = { totalExp: 0, eligibleForTds: 0, atStdRate: 0, atLowRate: 0, notDeducted: 0, tdsReq: 0, tdsDed: 0, tdsDep: 0 };
  }

  for (const v of vouchers) {
    const acc = v.accountName.toLowerCase();
    const nar = v.narration.toLowerCase();
    let sec = '';

    if (acc.includes('contractor') || acc.includes('freight') || acc.includes('transport') || acc.includes('advertising') || acc.includes('sub-contract')) {
      sec = '194C';
    } else if (acc.includes('professional') || acc.includes('legal') || acc.includes('audit fee') || acc.includes('consultancy') || acc.includes('technical')) {
      sec = '194J';
    } else if (acc.includes('rent') || acc.includes('lease rent')) {
      sec = '194I';
    } else if (acc.includes('commission') || acc.includes('brokerage')) {
      sec = '194H';
    } else if (acc.includes('interest paid') || acc.includes('finance charges') || acc.includes('interest on loan')) {
      sec = '194A';
    }

    if (sec && STATUTORY_TDS_RATES[sec]) {
      const rule = STATUTORY_TDS_RATES[sec];
      clause34Map[sec].totalExp += v.amount;

      // Check if exceeds statutory threshold
      if (v.amount >= rule.threshold || clause34Map[sec].totalExp >= rule.threshold) {
        clause34Map[sec].eligibleForTds += v.amount;
        const expectedTds = (v.amount * rule.rate) / 100;
        clause34Map[sec].tdsReq += expectedTds;

        if (nar.includes('tds deducted') || nar.includes('net of tds') || v.tdsDeducted) {
          const actualDed = v.tdsDeducted || expectedTds;
          clause34Map[sec].tdsDed += actualDed;
          clause34Map[sec].tdsDep += actualDed; // Assuming paid unless flagged
          clause34Map[sec].atStdRate += v.amount;
        } else if (nar.includes('form 15g') || nar.includes('form 15h') || nar.includes('nil deduction')) {
          clause34Map[sec].atLowRate += v.amount;
        } else {
          // Omission to deduct
          clause34Map[sec].notDeducted += v.amount;
        }
      }
    }
  }

  const clause34Summary: Clause34SectionSummary[] = Object.entries(clause34Map)
    .filter(([_, data]) => data.totalExp > 0)
    .map(([sec, data]) => {
      const shortDed = Math.max(0, data.tdsReq - data.tdsDed);
      const disallowance = data.notDeducted * 0.3; // 30% addition under Sec 40(a)(ia)
      const interest201 = shortDed * 0.015 * 6; // Est 6 months late deposit @ 1.5% p.m.
      return {
        section: sec,
        sectionName: STATUTORY_TDS_RATES[sec]?.name || `Section ${sec}`,
        totalExpenditureCovered: data.totalExp,
        amountEligibleForTds: data.eligibleForTds,
        amountDeductedAtStandardRate: data.atStdRate,
        amountDeductedAtLowerRate: data.atLowRate,
        amountNotDeducted: data.notDeducted,
        tdsAmountRequired: data.tdsReq,
        tdsAmountDeducted: data.tdsDed,
        tdsAmountDeposited: data.tdsDep,
        shortDeductionAmount: shortDed,
        disallowance40aIa: disallowance,
        estimatedInterest201: interest201,
      };
    });

  const clause34TotalCovered = clause34Summary.reduce((acc, s) => acc + s.totalExpenditureCovered, 0);
  const clause34TotalReq = clause34Summary.reduce((acc, s) => acc + s.tdsAmountRequired, 0);
  const clause34TotalDed = clause34Summary.reduce((acc, s) => acc + s.tdsAmountDeducted, 0);
  const clause34TotalShort = clause34Summary.reduce((acc, s) => acc + s.shortDeductionAmount, 0);
  const clause34TotalDisallowance = clause34Summary.reduce((acc, s) => acc + s.disallowance40aIa, 0);

  // =========================================================================
  // 4. CLAUSE 21(a) & 22 (PENALTIES, PERSONAL, & MSME INTEREST DISALLOWANCES)
  // =========================================================================
  const inadmissibleExpenses: InadmissibleExpenseRecord[] = [];
  let clause21aTotal = 0;
  let clause22MsmeTotal = 0;

  for (const v of vouchers) {
    const text = `${v.accountName} ${v.narration}`.toLowerCase();
    
    // Penalties & Fines
    const foundPenalty = PENALTY_KEYWORDS.filter((k) => text.includes(k));
    if (foundPenalty.length > 0) {
      inadmissibleExpenses.push({
        id: `inad_${v.id}`,
        date: v.date,
        voucherNo: v.voucherNo,
        accountName: v.accountName,
        narration: v.narration,
        amount: v.amount,
        disallowanceType: 'Penalty / Fine (Sec 37)',
        clauseReference: 'Clause 21(a)',
        confidenceScore: 0.95,
        flaggedKeywords: foundPenalty,
      });
      clause21aTotal += v.amount;
      continue;
    }

    // Personal Expenses
    const foundPersonal = PERSONAL_KEYWORDS.filter((k) => text.includes(k));
    if (foundPersonal.length > 0) {
      inadmissibleExpenses.push({
        id: `inad_${v.id}`,
        date: v.date,
        voucherNo: v.voucherNo,
        accountName: v.accountName,
        narration: v.narration,
        amount: v.amount,
        disallowanceType: 'Personal Expenditure',
        clauseReference: 'Clause 21(a)',
        confidenceScore: 0.88,
        flaggedKeywords: foundPersonal,
      });
      clause21aTotal += v.amount;
      continue;
    }

    // Capital items in P&L
    const foundCapital = CAPITAL_KEYWORDS.filter((k) => text.includes(k));
    if (foundCapital.length > 0 && !v.accountName.toLowerCase().includes('fixed asset')) {
      inadmissibleExpenses.push({
        id: `inad_${v.id}`,
        date: v.date,
        voucherNo: v.voucherNo,
        accountName: v.accountName,
        narration: v.narration,
        amount: v.amount,
        disallowanceType: 'Capital debited to P&L',
        clauseReference: 'Clause 21(a)',
        confidenceScore: 0.82,
        flaggedKeywords: foundCapital,
      });
      clause21aTotal += v.amount;
      continue;
    }

    // MSME Interest under Sec 23 MSMED Act
    if (text.includes('msme interest') || text.includes('interest to micro') || text.includes('sec 23 msme')) {
      inadmissibleExpenses.push({
        id: `inad_${v.id}`,
        date: v.date,
        voucherNo: v.voucherNo,
        accountName: v.accountName,
        narration: v.narration,
        amount: v.amount,
        disallowanceType: 'MSME Interest (Sec 23 MSMED)',
        clauseReference: 'Clause 22',
        confidenceScore: 0.98,
        flaggedKeywords: ['msme interest'],
      });
      clause22MsmeTotal += v.amount;
    }
  }

  // Total additions to net profit for Tax Audit Form 3CD
  const totalTaxAdditions = totalDisallowable40a3 + clause34TotalDisallowance + clause21aTotal + clause22MsmeTotal;
  const totalExpenditureScrutinized = vouchers.reduce((acc, v) => acc + v.amount, 0);

  return {
    clientName,
    assessmentYear,
    financialYear,
    totalExpenditureScrutinized,
    clause40a3ViolationsCount: cashViolations.filter((c) => !c.isRule6DdExempt).length,
    clause40a3TotalDisallowable: totalDisallowable40a3,
    clause40a3Rule6DdExemptCount: rule6DdExemptCount,
    totalClause44Expenditure,
    totalRegisteredExpenditure,
    totalUnregisteredExpenditure,
    unregisteredExpenditureRatio: Number(unregisteredExpenditureRatio.toFixed(2)),
    clause34TotalCoveredExpenditure: clause34TotalCovered,
    clause34TotalTdsRequired: clause34TotalReq,
    clause34TotalTdsDeducted: clause34TotalDed,
    clause34ShortDeduction: clause34TotalShort,
    clause34Disallowance40aIa: clause34TotalDisallowance,
    clause21aPersonalAndPenalties: clause21aTotal,
    clause22MsmeInterest: clause22MsmeTotal,
    totalTaxAdditions,
    cashViolations,
    clause44Breakdown,
    clause34Summary,
    inadmissibleExpenses,
  };
}

/**
 * Generates an ICAI-Compliant Multi-Sheet Form 3CD Annexures Excel Workbook
 */
export function exportForm3cdAnnexuresWorkbook(summary: Form3cdAuditSummary): void {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Executive Tax Audit Summary
  const execData = [
    { Parameter: 'Client Legal Entity Name', Value: summary.clientName },
    { Parameter: 'Financial Year (FY)', Value: summary.financialYear },
    { Parameter: 'Assessment Year (AY)', Value: summary.assessmentYear },
    { Parameter: 'Audit Form & Section', Value: 'Form 3CD under Section 44AB of Income Tax Act, 1961' },
    { Parameter: 'Total Scrutinized Expenditure', Value: summary.totalExpenditureScrutinized },
    { Parameter: 'Clause 21(d) / Sec 40A(3) Cash Disallowance', Value: summary.clause40a3TotalDisallowable },
    { Parameter: 'Clause 34 / Sec 40(a)(ia) TDS Disallowance (30%)', Value: summary.clause34Disallowance40aIa },
    { Parameter: 'Clause 21(a) Penalties & Personal Expenses', Value: summary.clause21aPersonalAndPenalties },
    { Parameter: 'Clause 22 MSME Inadmissible Interest (MSMED Act)', Value: summary.clause22MsmeInterest },
    { Parameter: 'TOTAL INADMISSIBLE ADDITIONS TO TAXABLE PROFIT', Value: summary.totalTaxAdditions },
    { Parameter: 'Clause 44 Unregistered Procurement Ratio', Value: `${summary.unregisteredExpenditureRatio}%` },
  ];
  const ws1 = createStyledWorksheet(execData, { currencyColumns: ['Value'] });
  XLSX.utils.book_append_sheet(wb, ws1, 'Form 3CD Summary');

  // Sheet 2: Clause 21(d) / Sec 40A(3) Cash Violations
  const cashData = summary.cashViolations.map((c, i) => ({
    'S.No.': i + 1,
    'Payment Date': c.date,
    'Payee Name': c.payeeName,
    'Vouchers Involved': c.voucherNumbers.join(', '),
    'Nature of Payment': c.isTransporter ? 'Transporter / Goods Carriage' : 'General Vendor',
    'Daily Aggregate Amount': c.dailyAggregateAmount,
    'Statutory Threshold': c.thresholdLimit,
    'Rule 6DD Exemption': c.isRule6DdExempt ? 'YES (Exempt)' : 'NO',
    'Disallowance Sec 40A(3)': c.disallowableAmount,
    'Auditor Remarks': c.exemptionReason,
  }));
  const ws2 = createStyledWorksheet(cashData.length > 0 ? cashData : [{ Note: 'No Clause 40A(3) violations detected' }], {
    currencyColumns: ['Daily Aggregate Amount', 'Statutory Threshold', 'Disallowance Sec 40A(3)'],
    headers: ['S.No.', 'Payment Date', 'Payee Name', 'Vouchers Involved', 'Nature of Payment', 'Daily Aggregate Amount', 'Statutory Threshold', 'Rule 6DD Exemption', 'Disallowance Sec 40A(3)', 'Auditor Remarks'],
  });
  XLSX.utils.book_append_sheet(wb, ws2, 'Clause 21d - Cash 40A(3)');

  // Sheet 3: Clause 44 Breakdown (Statutory CBDT Table)
  const c44Data = summary.clause44Breakdown.map((r) => ({
    'S.No.': r.sNo,
    'Expenditure Head': r.headOfExpenditure,
    'Total Expenditure': r.totalExpenditure,
    'GST Exempt Goods/Services': r.exemptSupplies,
    'Composition Entities (Sec 10)': r.compositionEntities,
    'Other Registered Entities': r.otherRegisteredEntities,
    'Total Registered Spend': r.totalRegisteredExpenditure,
    'Unregistered Entities Spend': r.unregisteredEntities,
    '% Unregistered': `${r.unregisteredPercent}%`,
    'Auditor Remarks': r.remarks,
  }));
  const ws3 = createStyledWorksheet(c44Data.length > 0 ? c44Data : [{ Note: 'No Clause 44 expenditure lines' }], {
    currencyColumns: ['Total Expenditure', 'GST Exempt Goods/Services', 'Composition Entities (Sec 10)', 'Other Registered Entities', 'Total Registered Spend', 'Unregistered Entities Spend'],
  });
  XLSX.utils.book_append_sheet(wb, ws3, 'Clause 44 - GST Spend');

  // Sheet 4: Clause 34 Chapter XVII-B TDS
  const c34Data = summary.clause34Summary.map((s) => ({
    'TDS Section': s.section,
    'Nature of Payment': s.sectionName,
    'Total Expenditure Incurred': s.totalExpenditureCovered,
    'Amount Liable to TDS': s.amountEligibleForTds,
    'Tax Deducted at Standard Rate': s.amountDeductedAtStandardRate,
    'Tax Deducted at Lower/Nil Rate': s.amountDeductedAtLowerRate,
    'Amount Tax NOT Deducted': s.amountNotDeducted,
    'Tax Required to be Deducted': s.tdsAmountRequired,
    'Tax Actually Deducted': s.tdsAmountDeducted,
    'Short Deduction': s.shortDeductionAmount,
    'Disallowance Sec 40(a)(ia) (30%)': s.disallowance40aIa,
    'Est. Interest Sec 201(1A)': s.estimatedInterest201,
  }));
  const ws4 = createStyledWorksheet(c34Data.length > 0 ? c34Data : [{ Note: 'No TDS covered sections' }], {
    currencyColumns: ['Total Expenditure Incurred', 'Amount Liable to TDS', 'Tax Deducted at Standard Rate', 'Amount Tax NOT Deducted', 'Tax Required to be Deducted', 'Tax Actually Deducted', 'Short Deduction', 'Disallowance Sec 40(a)(ia) (30%)', 'Est. Interest Sec 201(1A)'],
  });
  XLSX.utils.book_append_sheet(wb, ws4, 'Clause 34 - Chapter XVII-B');

  // Sheet 5: Clause 21(a) & 22 Inadmissible Expenses
  const inadData = summary.inadmissibleExpenses.map((inad, idx) => ({
    'S.No.': idx + 1,
    'Clause': inad.clauseReference,
    'Disallowance Category': inad.disallowanceType,
    'Date': inad.date,
    'Voucher No': inad.voucherNo,
    'Account Head': inad.accountName,
    'Narration': inad.narration,
    'Disallowable Amount': inad.amount,
    'Keywords Detected': inad.flaggedKeywords.join(', '),
  }));
  const ws5 = createStyledWorksheet(inadData.length > 0 ? inadData : [{ Note: 'No Clause 21(a) / 22 disallowances found' }], {
    currencyColumns: ['Disallowable Amount'],
  });
  XLSX.utils.book_append_sheet(wb, ws5, 'Clause 21a & 22 Disallowance');

  const cleanName = summary.clientName.replace(/[^a-zA-Z0-9]/g, '_');
  XLSX.writeFile(wb, `${cleanName}_Form3CD_Tax_Audit_Annexures.xlsx`);
}

/**
 * Realistic Indian Corporate Audit Sample Dataset for Form 3CD Testing
 */
export function getSampleForm3cdData(): Record<string, any>[] {
  return [
    // 40A(3) Cash Violations
    {
      'Voucher No': 'VR-2025-104',
      'Date': '14-06-2025',
      'Account': 'Site Labor Charges',
      'Payee Name': 'Ramesh Kumar Contractor',
      'Amount': 18500,
      'Payment Mode': 'Cash',
      'Narration': 'Paid cash to Ramesh contractor for site cleaning work',
      'GSTIN': '',
      'PAN': '',
    },
    {
      'Voucher No': 'VR-2025-105',
      'Date': '14-06-2025',
      'Account': 'Site Labor Charges',
      'Payee Name': 'Ramesh Kumar Contractor',
      'Amount': 12000,
      'Payment Mode': 'Cash',
      'Narration': 'Additional cash paid for overtime site labor to Ramesh',
      'GSTIN': '',
      'PAN': '',
    },
    {
      'Voucher No': 'VR-2025-188',
      'Date': '22-07-2025',
      'Account': 'Freight & Cartage Inward',
      'Payee Name': 'Navkar Road Carriers',
      'Amount': 42000,
      'Payment Mode': 'Cash',
      'Narration': 'Cash freight paid for truck transport GJ-01-XX-9901 exceeding 35k limit',
      'GSTIN': '',
      'PAN': 'AABCN7812M',
    },
    {
      'Voucher No': 'VR-2025-230',
      'Date': '15-08-2025',
      'Account': 'Emergency Repairs',
      'Payee Name': 'National Machinery Works',
      'Amount': 14000,
      'Payment Mode': 'Cash',
      'Narration': 'Cash payment for urgent boiler repair on bank holiday',
      'GSTIN': '27AAECN4455P1Z8',
      'PAN': 'AAECN4455P',
    },

    // Clause 21(a) Penalties & Personal expenses
    {
      'Voucher No': 'VR-2025-301',
      'Date': '10-09-2025',
      'Account': 'Rates & Taxes',
      'Payee Name': 'GST Department MH',
      'Amount': 35000,
      'Payment Mode': 'Bank',
      'Narration': 'Paid penalty and late fee under Section 122 of CGST Act for delay',
      'GSTIN': '27AAACG0000A1Z5',
      'PAN': '',
    },
    {
      'Voucher No': 'VR-2025-342',
      'Date': '25-10-2025',
      'Account': 'Directors Travelling Expenses',
      'Payee Name': 'MakeMyTrip Ltd',
      'Amount': 68000,
      'Payment Mode': 'Bank',
      'Narration': 'Personal family vacation flight tickets booked for director',
      'GSTIN': '07AAACM1234F1Z1',
      'PAN': 'AAACM1234F',
    },
    {
      'Voucher No': 'VR-2025-410',
      'Date': '05-11-2025',
      'Account': 'Finance Costs',
      'Payee Name': 'Micro Electro Components LLP',
      'Amount': 24500,
      'Payment Mode': 'Bank',
      'Narration': 'MSME interest paid under Sec 23 of MSMED Act on delayed payment beyond 45 days',
      'GSTIN': '27ABCDE1234F1Z5',
      'PAN': 'ABCDE1234F',
    },

    // Clause 34 TDS Covered Expenses
    {
      'Voucher No': 'VR-2025-501',
      'Date': '18-11-2025',
      'Account': 'Legal & Professional Charges',
      'Payee Name': 'Singhania & Partners Advocates',
      'Amount': 150000,
      'Payment Mode': 'Bank',
      'Narration': 'Professional fee for high court matter - TDS deducted under 194J',
      'GSTIN': '27AAAFS5566K1Z9',
      'PAN': 'AAAFS5566K',
    },
    {
      'Voucher No': 'VR-2025-512',
      'Date': '28-11-2025',
      'Account': 'Legal & Professional Charges',
      'Payee Name': 'Technocraft Consulting',
      'Amount': 85000,
      'Payment Mode': 'Bank',
      'Narration': 'Advisory fees paid without deducting TDS under Sec 194J',
      'GSTIN': '27AAECP1122Q1Z3',
      'PAN': 'AAECP1122Q',
    },
    {
      'Voucher No': 'VR-2025-610',
      'Date': '15-12-2025',
      'Account': 'Office Rent',
      'Payee Name': 'Prestige Commercial Towers',
      'Amount': 360000,
      'Payment Mode': 'Bank',
      'Narration': 'Quarterly commercial rent paid net of TDS under Sec 194I',
      'GSTIN': '27AABCP9988G1Z2',
      'PAN': 'AABCP9988G',
    },
    {
      'Voucher No': 'VR-2025-720',
      'Date': '10-01-2026',
      'Account': 'Advertising & Publicity',
      'Payee Name': 'Apex Media Solutions',
      'Amount': 240000,
      'Payment Mode': 'Bank',
      'Narration': 'Digital billboard marketing contract - TDS deducted under 194C',
      'GSTIN': '27AAACA9911P1Z1',
      'PAN': 'AAACA9911P',
    },

    // Clause 44 Registered vs Unregistered Breakdown
    {
      'Voucher No': 'VR-2026-805',
      'Date': '20-01-2026',
      'Account': 'Printing & Stationery',
      'Payee Name': 'Local Market Stationery Stores',
      'Amount': 42000,
      'Payment Mode': 'Electronic',
      'Narration': 'Office supplies and register books purchased from unregistered dealer',
      'GSTIN': '',
      'PAN': '',
    },
    {
      'Voucher No': 'VR-2026-820',
      'Date': '04-02-2026',
      'Account': 'Printing & Stationery',
      'Payee Name': 'PaperKraft India Pvt Ltd',
      'Amount': 95000,
      'Payment Mode': 'Bank',
      'Narration': 'Corporate letterheads and files - Regular GST registered vendor',
      'GSTIN': '27AAACP4411Q1Z4',
      'PAN': 'AAACP4411Q',
    },
    {
      'Voucher No': 'VR-2026-901',
      'Date': '18-02-2026',
      'Account': 'Office Maintenance & Pantry',
      'Payee Name': 'Fresh Dairy & Provisions (Composition)',
      'Amount': 38000,
      'Payment Mode': 'Bank',
      'Narration': 'Pantry milk and provisions - Supplier under GST Composition Scheme Sec 10',
      'GSTIN': '27AACCP8899K1Z0',
      'PAN': 'AACCP8899K',
    },
    {
      'Voucher No': 'VR-2026-940',
      'Date': '02-03-2026',
      'Account': 'Office Maintenance & Pantry',
      'Payee Name': 'Local Cleaning & Labor Services',
      'Amount': 55000,
      'Payment Mode': 'Electronic',
      'Narration': 'Deep cleaning and maintenance by unregistered service provider',
      'GSTIN': '',
      'PAN': '',
    },
  ];
}
