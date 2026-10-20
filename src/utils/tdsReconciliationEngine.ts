import * as XLSX from 'xlsx';
import { parseAmount, parseDate, parseDateObj, isValidPan, normalizeRefNo } from './dataSanitizer';
import { createStyledWorksheet } from './excelStyler';

export interface TdsSectionRule {
  section: string;
  name: string;
  standardRate: number; // e.g. 10%
  thresholdLimit: number; // annual threshold
  description: string;
}

export const DEFAULT_TDS_SECTIONS: Record<string, TdsSectionRule> = {
  '194C': { section: '194C', name: 'Contractors & Sub-contractors', standardRate: 2.0, thresholdLimit: 100000, description: '1% for Indiv/HUF, 2% for Companies/LLPs' },
  '194J_PROF': { section: '194J', name: 'Fees for Professional Services', standardRate: 10.0, thresholdLimit: 30000, description: '10% for professional services / legal / CA' },
  '194J_TECH': { section: '194J', name: 'Fees for Technical Services', standardRate: 2.0, thresholdLimit: 30000, description: '2% for FTS / BPO / call center' },
  '194I_RENT_BLD': { section: '194I', name: 'Rent on Land & Building', standardRate: 10.0, thresholdLimit: 240000, description: '10% on office rent / factory building' },
  '194I_RENT_PLANT': { section: '194I', name: 'Rent on Plant & Machinery', standardRate: 2.0, thresholdLimit: 240000, description: '2% on equipment hiring' },
  '194H': { section: '194H', name: 'Commission & Brokerage', standardRate: 2.0, thresholdLimit: 15000, description: '2% (amended Budget 2024)' },
  '194Q': { section: '194Q', name: 'Purchase of Goods (> 50 Lakhs)', standardRate: 0.1, thresholdLimit: 5000000, description: '0.1% on purchase of goods' },
  '194A': { section: '194A', name: 'Interest other than Securities', standardRate: 10.0, thresholdLimit: 5000, description: '10% on interest paid to NBFCs/parties' },
};

export type TdsReconciliationStatus =
  | 'reconciled'
  | 'amount_mismatch'
  | 'missing_challan'
  | 'invalid_pan_206aa'
  | 'rate_discrepancy'
  | 'late_deposit_interest'
  | 'duplicate_entry';

export interface TdsRecord {
  id: string;
  rowIndex: number;
  deducteeName: string;
  pan: string;
  isPanValid: boolean;
  section: string;
  bookingDate: string;
  paymentDate: string;
  taxableAmount: number;
  tdsDeducted: number;
  tdsDeposited: number;
  effectiveRatePercent: number;
  challanNo: string;
  bsrCode: string;
  challanDate: string;
  financialYear: string;

  // Audit Calculations
  expectedRatePercent: number;
  expectedTdsAmount: number;
  varianceAmount: number;
  status: TdsReconciliationStatus;
  severity: 'clean' | 'warning' | 'critical';
  isLateDeposit: boolean;
  delayMonths: number;
  estimatedInterest201: number;
  auditReasons: string[];
  raw: Record<string, any>;
}

export interface TdsReconciliationSummary {
  totalRecords: number;
  totalTaxableValue: number;
  totalTdsDeducted: number;
  totalTdsDeposited: number;
  reconciledCount: number;
  unreconciledCount: number;
  missingChallanCount: number;
  invalidPanCount: number;
  lateDepositCount: number;
  totalEstimatedInterest: number;
  netTdsVariance: number;
  isReconciled: boolean;
  records: TdsRecord[];
  sectionBreakdown: {
    section: string;
    count: number;
    taxable: number;
    deducted: number;
    deposited: number;
  }[];
}

export interface TdsColumnMapping {
  deducteeCol: string;
  panCol: string;
  sectionCol: string;
  bookingDateCol: string;
  taxableCol: string;
  tdsDeductedCol: string;
  tdsDepositedCol: string;
  challanNoCol: string;
  bsrCodeCol: string;
  challanDateCol: string;
}

import { detectColumnCategory } from './columnDetector';

/**
 * Auto-detects TDS register columns
 */
export function detectTdsColumns(headers: string[]): TdsColumnMapping {
  const mapping: TdsColumnMapping = {
    deducteeCol: '',
    panCol: '',
    sectionCol: '',
    bookingDateCol: '',
    taxableCol: '',
    tdsDeductedCol: '',
    tdsDepositedCol: '',
    challanNoCol: '',
    bsrCodeCol: '',
    challanDateCol: '',
  };

  for (const h of headers) {
    const clean = h.toLowerCase().trim();
    const cat = detectColumnCategory(h);

    if (!mapping.deducteeCol && (cat === 'partyName' || /deductee|party|vendor|supplier|payee|name/i.test(clean))) {
      mapping.deducteeCol = h;
    } else if (!mapping.panCol && (cat === 'gstin' || /pan|deductee.*pan/i.test(clean))) {
      mapping.panCol = h;
    } else if (!mapping.sectionCol && /section|sec|nature/i.test(clean)) {
      mapping.sectionCol = h;
    } else if (!mapping.challanDateCol && /challan.*date|deposit.*date|tender.*date/i.test(clean)) {
      mapping.challanDateCol = h;
    } else if (!mapping.bookingDateCol && (cat === 'date' || (/date|inv.*date|booking.*date/i.test(clean) && !/challan/i.test(clean)))) {
      mapping.bookingDateCol = h;
    } else if (!mapping.taxableCol && (cat === 'taxableAmount' || /taxable|base|bill.*amt|gross/i.test(clean))) {
      mapping.taxableCol = h;
    } else if (!mapping.tdsDeductedCol && /tds.*deduct|tds.*amt|deducted/i.test(clean)) {
      mapping.tdsDeductedCol = h;
    } else if (!mapping.tdsDepositedCol && /tds.*deposit|tds.*paid|challan.*amt|deposited/i.test(clean)) {
      mapping.tdsDepositedCol = h;
    } else if (!mapping.challanNoCol && (cat === 'refNo' || /challan.*no|cin|challan/i.test(clean))) {
      mapping.challanNoCol = h;
    } else if (!mapping.bsrCodeCol && /bsr|bank.*code/i.test(clean)) {
      mapping.bsrCodeCol = h;
    }
  }

  return mapping;
}

/**
 * Main TDS Reconciliation Engine
 */
export function reconcileTdsRecords(
  rows: Record<string, any>[],
  mapping: TdsColumnMapping,
  fyLabel = 'FY 2026-27'
): TdsReconciliationSummary {
  const records: TdsRecord[] = [];
  const sectionMap = new Map<string, { count: number; taxable: number; deducted: number; deposited: number }>();

  // Track duplicate deductee bills
  const dupCheckMap = new Map<string, number>();

  let totalTaxable = 0;
  let totalDeducted = 0;
  let totalDeposited = 0;
  let missingChallanCount = 0;
  let invalidPanCount = 0;
  let lateDepositCount = 0;
  let totalEstimatedInterest = 0;

  rows.forEach((r, idx) => {
    const rowIndex = idx + 2;
    const deducteeName = String(r[mapping.deducteeCol] || '').trim();
    const pan = String(r[mapping.panCol] || '').trim().toUpperCase();
    const isPanValid = isValidPan(pan);

    const rawSec = String(r[mapping.sectionCol] || '').trim().toUpperCase();
    const cleanSec = rawSec.replace(/[^0-9A-Z]/g, '') || '194C';

    const bookingDate = parseDate(r[mapping.bookingDateCol], 'DD-MM-YYYY');
    const taxableAmount = parseAmount(r[mapping.taxableCol]);
    const tdsDeducted = parseAmount(r[mapping.tdsDeductedCol]);
    let tdsDeposited = parseAmount(r[mapping.tdsDepositedCol]);

    const challanNo = String(r[mapping.challanNoCol] || '').trim();
    const bsrCode = String(r[mapping.bsrCodeCol] || '').trim();
    const challanDate = parseDate(r[mapping.challanDateCol], 'DD-MM-YYYY');

    // If deposited amount not given but challan exists, assume deposited = deducted
    if (tdsDeposited === 0 && challanNo && challanDate) {
      tdsDeposited = tdsDeducted;
    }

    const effectiveRate = taxableAmount > 0 ? (tdsDeducted / taxableAmount) * 100 : 0;

    // Determine expected statutory rate
    let expectedRate = 2.0; // default
    if (cleanSec.includes('194J')) {
      expectedRate = 10.0;
    } else if (cleanSec.includes('194I')) {
      expectedRate = 10.0;
    } else if (cleanSec.includes('194H')) {
      expectedRate = 2.0;
    } else if (cleanSec.includes('194Q')) {
      expectedRate = 0.1;
    } else if (cleanSec.includes('194C')) {
      expectedRate = 2.0;
    }

    // Section 206AA: Higher 20% rate mandatory if PAN is missing or invalid
    if (!isPanValid && taxableAmount > 0) {
      expectedRate = 20.0;
      invalidPanCount++;
    }

    const expectedTds = Math.round(taxableAmount * (expectedRate / 100));
    const varianceAmount = Math.abs(tdsDeducted - tdsDeposited);

    const bDateObj = parseDateObj(r[mapping.bookingDateCol]);
    const cDateObj = parseDateObj(r[mapping.challanDateCol]);

    // Due date check: 7th of following month (or 30th April for March)
    let isLateDeposit = false;
    let delayMonths = 0;
    let estimatedInterest201 = 0;

    if (bDateObj && cDateObj && !isNaN(bDateObj.getTime()) && !isNaN(cDateObj.getTime())) {
      const bMonth = bDateObj.getMonth(); // 0 = Jan, 2 = Mar, 11 = Dec
      const bYear = bDateObj.getFullYear();

      // Statutory Due Date: 7th of following month (or April 30 for March)
      let dueDate: Date;
      if (bMonth === 2) {
        // March deduction -> Due date is April 30 of same calendar year
        dueDate = new Date(bYear, 3, 30, 23, 59, 59);
      } else {
        const nextMonthYear = bMonth === 11 ? bYear + 1 : bYear;
        const nextMonth = bMonth === 11 ? 0 : bMonth + 1;
        dueDate = new Date(nextMonthYear, nextMonth, 7, 23, 59, 59);
      }

      if (cDateObj > dueDate) {
        isLateDeposit = true;
        lateDepositCount++;
        // Section 201(1A)(ii) Income Tax Act: 1.5% interest per month or part of a month from deduction to payment
        const monthDiff = (cDateObj.getFullYear() - bDateObj.getFullYear()) * 12 + (cDateObj.getMonth() - bDateObj.getMonth());
        delayMonths = Math.max(1, monthDiff + 1);
        estimatedInterest201 = Math.round(tdsDeducted * (0.015 * delayMonths));
        if (isNaN(estimatedInterest201) || !isFinite(estimatedInterest201)) {
          estimatedInterest201 = 0;
        }
        totalEstimatedInterest += estimatedInterest201;
      }
    }

    const auditReasons: string[] = [];
    let status: TdsReconciliationStatus = 'reconciled';
    let severity: 'clean' | 'warning' | 'critical' = 'clean';

    if (!isPanValid) {
      status = 'invalid_pan_206aa';
      severity = 'critical';
      auditReasons.push('Invalid/Missing PAN. Section 206AA requires 20% TDS withholding.');
    } else if (!challanNo || !challanDate) {
      status = 'missing_challan';
      severity = 'critical';
      missingChallanCount++;
      auditReasons.push('TDS deducted in books but no payment challan / BSR code recorded.');
    } else if (varianceAmount > 10) {
      status = 'amount_mismatch';
      severity = 'critical';
      auditReasons.push(`Variance of ₹${varianceAmount.toFixed(2)} between deducted (₹${tdsDeducted}) and deposited (₹${tdsDeposited}).`);
    } else if (Math.abs(effectiveRate - expectedRate) > 0.5 && expectedRate !== 20) {
      status = 'rate_discrepancy';
      severity = 'warning';
      auditReasons.push(`Effective rate ${effectiveRate.toFixed(1)}% departs from standard section rate ${expectedRate}%.`);
    } else if (isLateDeposit) {
      status = 'late_deposit_interest';
      severity = 'warning';
      auditReasons.push(`Late deposit. Deposited on ${challanDate} after statutory due date. Sec 201(1A) interest: ₹${estimatedInterest201}.`);
    } else {
      auditReasons.push('Fully reconciled with tax payment challan.');
    }

    // Duplicate check key
    const dupKey = `${deducteeName}_${taxableAmount}_${bookingDate}`;
    dupCheckMap.set(dupKey, (dupCheckMap.get(dupKey) || 0) + 1);
    if (dupCheckMap.get(dupKey)! > 1) {
      if (severity !== 'critical') {
        status = 'duplicate_entry';
        severity = 'warning';
      }
      auditReasons.push('Potential duplicate TDS entry with identical party, amount, and date.');
    }

    totalTaxable += taxableAmount;
    totalDeducted += tdsDeducted;
    totalDeposited += tdsDeposited;

    // Track Section totals
    if (!sectionMap.has(cleanSec)) {
      sectionMap.set(cleanSec, { count: 0, taxable: 0, deducted: 0, deposited: 0 });
    }
    const secEntry = sectionMap.get(cleanSec)!;
    secEntry.count++;
    secEntry.taxable += taxableAmount;
    secEntry.deducted += tdsDeducted;
    secEntry.deposited += tdsDeposited;

    records.push({
      id: `tds-${idx + 1}`,
      rowIndex,
      deducteeName: deducteeName || `Deductee ${idx + 1}`,
      pan,
      isPanValid,
      section: cleanSec,
      bookingDate: bookingDate || 'N/A',
      paymentDate: challanDate || bookingDate || 'N/A',
      taxableAmount,
      tdsDeducted,
      tdsDeposited,
      effectiveRatePercent: effectiveRate,
      challanNo,
      bsrCode,
      challanDate: challanDate || 'Pending',
      financialYear: fyLabel,
      expectedRatePercent: expectedRate,
      expectedTdsAmount: expectedTds,
      varianceAmount,
      status,
      severity,
      isLateDeposit,
      delayMonths,
      estimatedInterest201,
      auditReasons,
      raw: r,
    });
  });

  const reconciledCount = records.filter((r) => r.status === 'reconciled').length;
  const unreconciledCount = records.length - reconciledCount;
  const netTdsVariance = Math.abs(totalDeducted - totalDeposited);
  const isReconciled = netTdsVariance <= 5 && unreconciledCount === 0;

  const sectionBreakdown = Array.from(sectionMap.entries()).map(([sec, val]) => ({
    section: `Sec ${sec}`,
    count: val.count,
    taxable: val.taxable,
    deducted: val.deducted,
    deposited: val.deposited,
  }));

  return {
    totalRecords: records.length,
    totalTaxableValue: totalTaxable,
    totalTdsDeducted: totalDeducted,
    totalTdsDeposited: totalDeposited,
    reconciledCount,
    unreconciledCount,
    missingChallanCount,
    invalidPanCount,
    lateDepositCount,
    totalEstimatedInterest,
    netTdsVariance,
    isReconciled,
    records,
    sectionBreakdown,
  };
}

/**
 * Exports complete multi-sheet Excel TDS Reconciliation workbook
 */
export function exportTdsWorkbook(
  summary: TdsReconciliationSummary,
  clientName = 'Audit Client',
  fy = 'FY 2026-27'
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Executive TDS Summary
  const summaryRows = [
    { Parameter: 'Client Entity', Value: clientName },
    { Parameter: 'Financial Year', Value: fy },
    { Parameter: 'Total TDS Transactions Scanned', Value: summary.totalRecords },
    { Parameter: 'Total Taxable Base Amount (₹)', Value: summary.totalTaxableValue },
    { Parameter: 'Total TDS Deducted in Books (₹)', Value: summary.totalTdsDeducted },
    { Parameter: 'Total TDS Deposited via Challans (₹)', Value: summary.totalTdsDeposited },
    { Parameter: 'Net Reconciliation Variance (₹)', Value: summary.netTdsVariance },
    { Parameter: 'Reconciled Records', Value: summary.reconciledCount },
    { Parameter: 'Unreconciled / Flagged Exceptions', Value: summary.unreconciledCount },
    { Parameter: 'Missing Challan / Deposit Details', Value: summary.missingChallanCount },
    { Parameter: 'Invalid PAN (Sec 206AA 20% Risk)', Value: summary.invalidPanCount },
    { Parameter: 'Late Deposit Instances (Sec 201(1A))', Value: summary.lateDepositCount },
    { Parameter: 'Total Estimated Interest Exposure (₹)', Value: summary.totalEstimatedInterest },
    { Parameter: 'Statutory Disclaimer', Value: 'Data reconciliation only. Does not constitute filing confirmation of Form 24Q/26Q.' },
  ];
  const wsSummary = createStyledWorksheet(summaryRows, {
    currencyColumns: ['Value'],
    columnWidths: { Parameter: 40, Value: 34 },
  });
  XLSX.utils.book_append_sheet(wb, wsSummary, 'TDS_Summary');

  // Sheet 2: Deductee-Wise Status
  const deducteeRows = summary.records.map((r) => ({
    'Row #': r.rowIndex,
    'Deductee Name': r.deducteeName,
    'PAN': r.pan,
    'PAN Valid': r.isPanValid ? 'VALID' : 'INVALID (206AA)',
    'Section': r.section,
    'Booking Date': r.bookingDate,
    'Taxable Base (₹)': r.taxableAmount,
    'TDS Deducted (₹)': r.tdsDeducted,
    'TDS Deposited (₹)': r.tdsDeposited,
    'Variance (₹)': r.varianceAmount,
    'Effective Rate': `${r.effectiveRatePercent.toFixed(1)}%`,
    'Challan No': r.challanNo || 'MISSING',
    'BSR Code': r.bsrCode || 'MISSING',
    'Challan Date': r.challanDate,
    'Reconciliation Status': r.status.replace(/_/g, ' ').toUpperCase(),
    'Audit Observation': r.auditReasons.join(' | '),
  }));
  const wsDeductee = createStyledWorksheet(deducteeRows, {
    currencyColumns: ['Taxable Base (₹)', 'TDS Deducted (₹)', 'TDS Deposited (₹)', 'Variance (₹)'],
  });
  XLSX.utils.book_append_sheet(wb, wsDeductee, 'Deductee_Reconciliation');

  // Sheet 3: Section-Wise Analysis
  const wsSection = createStyledWorksheet(summary.sectionBreakdown, {
    currencyColumns: ['taxable', 'deducted', 'deposited'],
  });
  XLSX.utils.book_append_sheet(wb, wsSection, 'Section_Breakdown');

  return wb;
}

/**
 * Realistic Indian TDS Sample Dataset
 */
export function getSampleTdsRows(): Record<string, any>[] {
  return [
    {
      'Deductee Name': 'Tata Consultancy Services Ltd',
      'PAN': 'AAACT2727Q',
      'Section': '194J',
      'Booking Date': '15-07-2026',
      'Taxable Amount': 250000,
      'TDS Deducted': 25000, // 10%
      'TDS Deposited': 25000,
      'Challan No': '04128',
      'BSR Code': '0510012',
      'Challan Date': '04-08-2026', // Deposited before 7th Aug - Compliant
    },
    {
      'Deductee Name': 'Shree Logistics Transporters',
      'PAN': 'ABCPS8812K',
      'Section': '194C',
      'Booking Date': '20-07-2026',
      'Taxable Amount': 180000,
      'TDS Deducted': 3600, // 2%
      'TDS Deposited': 3600,
      'Challan No': '04129',
      'BSR Code': '0510012',
      'Challan Date': '05-08-2026',
    },
    {
      'Deductee Name': 'Apex Commercial Space (Rent)',
      'PAN': 'AAACA1029F',
      'Section': '194I',
      'Booking Date': '01-08-2026',
      'Taxable Amount': 150000,
      'TDS Deducted': 15000, // 10%
      'TDS Deposited': 15000,
      'Challan No': '05991',
      'BSR Code': '0510012',
      'Challan Date': '25-09-2026', // LATE DEPOSIT (Due date 7th Sept, paid 25th Sept!)
    },
    {
      'Deductee Name': 'Kunal Singhania (Design Consultant)',
      'PAN': 'INVALID_PAN_88', // INVALID PAN EXCEPTION (Sec 206AA 20% required!)
      'Section': '194J',
      'Booking Date': '10-08-2026',
      'Taxable Amount': 80000,
      'TDS Deducted': 8000, // Deducted only 10% instead of 20%!
      'TDS Deposited': 8000,
      'Challan No': '05992',
      'BSR Code': '0510012',
      'Challan Date': '06-09-2026',
    },
    {
      'Deductee Name': 'Godrej Security Systems',
      'PAN': 'AAACG0549H',
      'Section': '194C',
      'Booking Date': '18-08-2026',
      'Taxable Amount': 320000,
      'TDS Deducted': 6400, // 2%
      'TDS Deposited': 0, // MISSING CHALLAN EXCEPTION!
      'Challan No': '',
      'BSR Code': '',
      'Challan Date': '',
    },
    {
      'Deductee Name': 'Reliance Retail Limited',
      'PAN': 'AAACR5055K',
      'Section': '194Q',
      'Booking Date': '25-08-2026',
      'Taxable Amount': 6500000,
      'TDS Deducted': 6500, // 0.1%
      'TDS Deposited': 6500,
      'Challan No': '06102',
      'BSR Code': '0510012',
      'Challan Date': '07-09-2026',
    },
    {
      'Deductee Name': 'Delta Instruments Maintenance',
      'PAN': 'AAACD9910R',
      'Section': '194C',
      'Booking Date': '28-08-2026',
      'Taxable Amount': 95000,
      'TDS Deducted': 1900,
      'TDS Deposited': 1500, // AMOUNT MISMATCH (Variance ₹400 short deposited!)
      'Challan No': '06103',
      'BSR Code': '0510012',
      'Challan Date': '07-09-2026',
    },
  ];
}
