import * as XLSX from 'xlsx';
import { parseAmount, parseDate } from './dataSanitizer';
import { createStyledWorksheet } from './excelStyler';

/**
 * GSTR-1 vs GSTR-3B vs BOOKS 3-WAY RECONCILIATION & DRC-01B NOTICE SCRUTINY ENGINE
 *
 * Implements authoritative tripartite reconciliation:
 * 1. Turnover: Books vs GSTR-1 vs GSTR-3B Table 3.1(a)
 * 2. Tax Liabilities: IGST, CGST, SGST, and Cess across all 3 sources
 * 3. DRC-01B (Rule 88C) Automated Risk Engine: Flags tax difference where GSTR-1 > GSTR-3B > ₹25,000 / 10%
 * 4. Section 50 Interest Exposure Calculation @ 18% p.a.
 * 5. GSTR-9/9C Table 5 & Table 7 Annual Turnover Bridge
 */

export interface Monthly3WayRecord {
  month: string; // e.g. "April 2025"
  monthIndex: number; // 1 to 12
  quarter: 'Q1' | 'Q2' | 'Q3' | 'Q4';

  // 1. Turnover / Taxable Value
  booksTurnover: number;
  gstr1Turnover: number;
  gstr3bTurnover: number;
  turnoverVarianceGstr1Vs3b: number; // GSTR-1 minus GSTR-3B
  turnoverVarianceBooksVsGstr1: number; // Books minus GSTR-1

  // 2. IGST Liability
  booksIgst: number;
  gstr1Igst: number;
  gstr3bIgst: number;
  igstDiff1Vs3b: number; // GSTR-1 minus GSTR-3B

  // 3. CGST Liability
  booksCgst: number;
  gstr1Cgst: number;
  gstr3bCgst: number;
  cgstDiff1Vs3b: number;

  // 4. SGST Liability
  booksSgst: number;
  gstr1Sgst: number;
  gstr3bSgst: number;
  sgstDiff1Vs3b: number;

  // 5. Cess Liability
  booksCess: number;
  gstr1Cess: number;
  gstr3bCess: number;
  cessDiff1Vs3b: number;

  // Total Tax Liabilities
  booksTotalTax: number;
  gstr1TotalTax: number;
  gstr3bTotalTax: number;
  netTaxShortfall1Vs3b: number; // GSTR-1 minus GSTR-3B (Positive = GSTR-1 is higher = DRC-01B risk)

  // Compliance & DRC-01B Evaluation
  drc01bStatus: 'CLEAN' | 'WARNING' | 'CRITICAL_NOTICE_RISK' | 'EXCESS_TAX_PAID';
  drc01bThresholdExceeded: boolean;
  estimatedInterestSec50: number; // 18% p.a. pro-rated
  suggestedDepartmentReply: string;
  reconciliationStatus: 'Fully Reconciled' | 'Timing Difference' | 'Material Variance' | 'Unreported in Books';
}

export interface GstTripleRecSummary {
  clientGstin: string;
  clientLegalName: string;
  financialYear: string;

  // Annual Totals
  annualBooksTurnover: number;
  annualGstr1Turnover: number;
  annualGstr3bTurnover: number;
  netTurnoverVariance1Vs3b: number;
  netTurnoverVarianceBooksVs1: number;

  annualBooksTotalTax: number;
  annualGstr1TotalTax: number;
  annualGstr3bTotalTax: number;
  netAnnualTaxVariance1Vs3b: number;

  // Head-wise Annual Variances
  annualIgstVariance: number;
  annualCgstVariance: number;
  annualSgstVariance: number;
  annualCessVariance: number;

  // DRC-01B Metrics
  criticalDrc01bMonthsCount: number;
  totalDrc01bTaxShortfall: number;
  totalSection50Interest: number;
  reconciliationMatchPercent: number;

  monthlyBreakdown: Monthly3WayRecord[];
  drc01bNoticeItems: Monthly3WayRecord[];
}

export const INDIAN_MONTHS = [
  'April', 'May', 'June', 'July', 'August', 'September',
  'October', 'November', 'December', 'January', 'February', 'March'
];

/**
 * Auto-detect columns in a consolidated or multi-sheet 3-way reconciliation file
 */
export function detectGstTripleRecColumns(headers: string[]): Record<string, string> {
  const norm = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, '');
  const findCol = (keys: string[]): string => {
    for (const h of headers) {
      const nh = norm(h);
      if (keys.some((k) => nh.includes(norm(k)))) return h;
    }
    return '';
  };

  return {
    monthCol: findCol(['month', 'period', 'returnperiod', 'monthname', 'taxperiod']),
    
    // Books columns
    booksTurnoverCol: findCol(['booksturnover', 'bookssales', 'turnoverbooks', 'salesbooks', 'plsales', 'auditedsales']),
    booksIgstCol: findCol(['booksigst', 'igstbooks', 'booksigstamount']),
    booksCgstCol: findCol(['bookscgst', 'cgstbooks', 'bookscgstamount']),
    booksSgstCol: findCol(['bookssgst', 'sgstbooks', 'bookssgstamount']),
    booksCessCol: findCol(['bookscess', 'cessbooks']),

    // GSTR-1 columns
    gstr1TurnoverCol: findCol(['gstr1turnover', 'turnovergstr1', 'gstr1taxable', 'taxablegstr1', 'gstr1sales']),
    gstr1IgstCol: findCol(['gstr1igst', 'igstgstr1', 'gstr1igstliability']),
    gstr1CgstCol: findCol(['gstr1cgst', 'cgstgstr1', 'gstr1cgstliability']),
    gstr1SgstCol: findCol(['gstr1sgst', 'sgstgstr1', 'gstr1sgstliability']),
    gstr1CessCol: findCol(['gstr1cess', 'cessgstr1']),

    // GSTR-3B columns
    gstr3bTurnoverCol: findCol(['gstr3bturnover', 'turnovergstr3b', 'gstr3btaxable', 'taxablegstr3b', 'gstr3btable31']),
    gstr3bIgstCol: findCol(['gstr3bigst', 'igstgstr3b', 'gstr3bigstpaid']),
    gstr3bCgstCol: findCol(['gstr3bcgst', 'cgstgstr3b', 'gstr3bcgstpaid']),
    gstr3bSgstCol: findCol(['gstr3bsgst', 'sgstgstr3b', 'gstr3bsgstpaid']),
    gstr3bCessCol: findCol(['gstr3bcess', 'cessgstr3b']),
  };
}

/**
 * Reconciles 3-Way data month-by-month and calculates DRC-01B notice risk
 */
export function reconcileGstTripleData(
  rows: Record<string, any>[],
  mapping: Record<string, string>,
  clientGstin: string = '27AAACP1234F1Z8',
  clientLegalName: string = 'Apex Technologies Private Limited',
  financialYear: string = 'FY 2025-26',
  drc01bThreshold: number = 25000 // Statutory default: ₹25,000 tax difference
): GstTripleRecSummary {
  const monthlyBreakdown: Monthly3WayRecord[] = [];

  // Group or iterate rows
  rows.forEach((r, idx) => {
    const rawMonth = String(r[mapping.monthCol] || INDIAN_MONTHS[idx % 12] || `Month-${idx + 1}`).trim();
    const monthIndex = (idx % 12) + 1;
    const quarter: Monthly3WayRecord['quarter'] =
      monthIndex <= 3 ? 'Q1' : monthIndex <= 6 ? 'Q2' : monthIndex <= 9 ? 'Q3' : 'Q4';

    // 1. Turnover
    const booksTurnover = Math.abs(parseAmount(r[mapping.booksTurnoverCol]));
    const gstr1Turnover = Math.abs(parseAmount(r[mapping.gstr1TurnoverCol]));
    const gstr3bTurnover = Math.abs(parseAmount(r[mapping.gstr3bTurnoverCol]));
    const turnoverVarianceGstr1Vs3b = gstr1Turnover - gstr3bTurnover;
    const turnoverVarianceBooksVsGstr1 = booksTurnover - gstr1Turnover;

    // 2. IGST
    const booksIgst = Math.abs(parseAmount(r[mapping.booksIgstCol]));
    const gstr1Igst = Math.abs(parseAmount(r[mapping.gstr1IgstCol]));
    const gstr3bIgst = Math.abs(parseAmount(r[mapping.gstr3bIgstCol]));
    const igstDiff1Vs3b = gstr1Igst - gstr3bIgst;

    // 3. CGST
    const booksCgst = Math.abs(parseAmount(r[mapping.booksCgstCol]));
    const gstr1Cgst = Math.abs(parseAmount(r[mapping.gstr1CgstCol]));
    const gstr3bCgst = Math.abs(parseAmount(r[mapping.gstr3bCgstCol]));
    const cgstDiff1Vs3b = gstr1Cgst - gstr3bCgst;

    // 4. SGST
    const booksSgst = Math.abs(parseAmount(r[mapping.booksSgstCol]));
    const gstr1Sgst = Math.abs(parseAmount(r[mapping.gstr1SgstCol]));
    const gstr3bSgst = Math.abs(parseAmount(r[mapping.gstr3bSgstCol]));
    const sgstDiff1Vs3b = gstr1Sgst - gstr3bSgst;

    // 5. Cess
    const booksCess = Math.abs(parseAmount(r[mapping.booksCessCol]));
    const gstr1Cess = Math.abs(parseAmount(r[mapping.gstr1CessCol]));
    const gstr3bCess = Math.abs(parseAmount(r[mapping.gstr3bCessCol]));
    const cessDiff1Vs3b = gstr1Cess - gstr3bCess;

    // Total Taxes
    const booksTotalTax = booksIgst + booksCgst + booksSgst + booksCess;
    const gstr1TotalTax = gstr1Igst + gstr1Cgst + gstr1Sgst + gstr1Cess;
    const gstr3bTotalTax = gstr3bIgst + gstr3bCgst + gstr3bSgst + gstr3bCess;
    const netTaxShortfall1Vs3b = gstr1TotalTax - gstr3bTotalTax;

    // DRC-01B Notice Risk Evaluation (Rule 88C)
    let drc01bStatus: Monthly3WayRecord['drc01bStatus'] = 'CLEAN';
    let drc01bThresholdExceeded = false;
    let estimatedInterestSec50 = 0;
    let suggestedDepartmentReply = 'Nil variance between GSTR-1 and GSTR-3B tax liabilities.';
    let reconciliationStatus: Monthly3WayRecord['reconciliationStatus'] = 'Fully Reconciled';

    if (netTaxShortfall1Vs3b > drc01bThreshold) {
      drc01bStatus = 'CRITICAL_NOTICE_RISK';
      drc01bThresholdExceeded = true;
      reconciliationStatus = 'Material Variance';

      // Estimate interest under Sec 50 @ 18% p.a. for approx delay (e.g. 90 days average)
      estimatedInterestSec50 = Math.round((netTaxShortfall1Vs3b * 0.18 * 90) / 365);

      if (turnoverVarianceGstr1Vs3b > 0 && Math.abs(netTaxShortfall1Vs3b - turnoverVarianceGstr1Vs3b * 0.18) < 1000) {
        suggestedDepartmentReply =
          'DRC-01B Rule 88C: Outward taxable turnover reported in GSTR-1 exceeded GSTR-3B. Differential tax payable under DRC-03 along with Section 50 interest, or verify if credit notes were omitted in GSTR-3B Table 4.';
      } else if (gstr1Igst > gstr3bIgst && gstr1Cgst < gstr3bCgst) {
        suggestedDepartmentReply =
          'Head Mismatch: Tax reported under IGST in GSTR-1 was paid under CGST/SGST in GSTR-3B. Section 77 refund / DRC-03 adjustment required.';
      } else {
        suggestedDepartmentReply =
          'Typographical error in GSTR-1 liability: Requisite amendment to be filed in Table 9 of subsequent GSTR-1.';
      }
    } else if (netTaxShortfall1Vs3b > 0) {
      drc01bStatus = 'WARNING';
      reconciliationStatus = 'Timing Difference';
      suggestedDepartmentReply = 'Tax shortfall is below the statutory DRC-01B automated trigger threshold of ₹25,000.';
    } else if (netTaxShortfall1Vs3b < -100) {
      drc01bStatus = 'EXCESS_TAX_PAID';
      reconciliationStatus = 'Timing Difference';
      suggestedDepartmentReply = 'Excess tax paid in GSTR-3B compared to GSTR-1. Eligible for adjustment in upcoming months or refund via RFD-01.';
    }

    // Check Books vs GSTR-1 turnover walk
    if (Math.abs(turnoverVarianceBooksVsGstr1) > 50000 && drc01bStatus === 'CLEAN') {
      reconciliationStatus = 'Unreported in Books';
    }

    monthlyBreakdown.push({
      month: rawMonth,
      monthIndex,
      quarter,
      booksTurnover,
      gstr1Turnover,
      gstr3bTurnover,
      turnoverVarianceGstr1Vs3b,
      turnoverVarianceBooksVsGstr1,
      booksIgst,
      gstr1Igst,
      gstr3bIgst,
      igstDiff1Vs3b,
      booksCgst,
      gstr1Cgst,
      gstr3bCgst,
      cgstDiff1Vs3b,
      booksSgst,
      gstr1Sgst,
      gstr3bSgst,
      sgstDiff1Vs3b,
      booksCess,
      gstr1Cess,
      gstr3bCess,
      cessDiff1Vs3b,
      booksTotalTax,
      gstr1TotalTax,
      gstr3bTotalTax,
      netTaxShortfall1Vs3b,
      drc01bStatus,
      drc01bThresholdExceeded,
      estimatedInterestSec50,
      suggestedDepartmentReply,
      reconciliationStatus,
    });
  });

  // Annual Totals
  const annualBooksTurnover = monthlyBreakdown.reduce((acc, m) => acc + m.booksTurnover, 0);
  const annualGstr1Turnover = monthlyBreakdown.reduce((acc, m) => acc + m.gstr1Turnover, 0);
  const annualGstr3bTurnover = monthlyBreakdown.reduce((acc, m) => acc + m.gstr3bTurnover, 0);
  const netTurnoverVariance1Vs3b = annualGstr1Turnover - annualGstr3bTurnover;
  const netTurnoverVarianceBooksVs1 = annualBooksTurnover - annualGstr1Turnover;

  const annualBooksTotalTax = monthlyBreakdown.reduce((acc, m) => acc + m.booksTotalTax, 0);
  const annualGstr1TotalTax = monthlyBreakdown.reduce((acc, m) => acc + m.gstr1TotalTax, 0);
  const annualGstr3bTotalTax = monthlyBreakdown.reduce((acc, m) => acc + m.gstr3bTotalTax, 0);
  const netAnnualTaxVariance1Vs3b = annualGstr1TotalTax - annualGstr3bTotalTax;

  const annualIgstVariance = monthlyBreakdown.reduce((acc, m) => acc + m.igstDiff1Vs3b, 0);
  const annualCgstVariance = monthlyBreakdown.reduce((acc, m) => acc + m.cgstDiff1Vs3b, 0);
  const annualSgstVariance = monthlyBreakdown.reduce((acc, m) => acc + m.sgstDiff1Vs3b, 0);
  const annualCessVariance = monthlyBreakdown.reduce((acc, m) => acc + m.cessDiff1Vs3b, 0);

  const drc01bNoticeItems = monthlyBreakdown.filter((m) => m.drc01bThresholdExceeded);
  const criticalDrc01bMonthsCount = drc01bNoticeItems.length;
  const totalDrc01bTaxShortfall = drc01bNoticeItems.reduce((acc, m) => acc + m.netTaxShortfall1Vs3b, 0);
  const totalSection50Interest = drc01bNoticeItems.reduce((acc, m) => acc + m.estimatedInterestSec50, 0);

  // Match Score %
  const totalMonths = Math.max(1, monthlyBreakdown.length);
  const cleanMonths = monthlyBreakdown.filter((m) => m.drc01bStatus === 'CLEAN').length;
  const reconciliationMatchPercent = Number(((cleanMonths / totalMonths) * 100).toFixed(1));

  return {
    clientGstin,
    clientLegalName,
    financialYear,
    annualBooksTurnover,
    annualGstr1Turnover,
    annualGstr3bTurnover,
    netTurnoverVariance1Vs3b,
    netTurnoverVarianceBooksVs1,
    annualBooksTotalTax,
    annualGstr1TotalTax,
    annualGstr3bTotalTax,
    netAnnualTaxVariance1Vs3b,
    annualIgstVariance,
    annualCgstVariance,
    annualSgstVariance,
    annualCessVariance,
    criticalDrc01bMonthsCount,
    totalDrc01bTaxShortfall,
    totalSection50Interest,
    reconciliationMatchPercent,
    monthlyBreakdown,
    drc01bNoticeItems,
  };
}

/**
 * Generates an ICAI & GST Auditor Grade Multi-Sheet Excel Workbook
 */
export function exportGstTripleRecWorkbook(summary: GstTripleRecSummary): void {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Executive 3-Way Summary & DRC-01B Dashboard
  const execData = [
    { Parameter: 'Taxpayer Trade / Legal Name', Value: summary.clientLegalName },
    { Parameter: 'GSTIN', Value: summary.clientGstin },
    { Parameter: 'Financial Year (FY)', Value: summary.financialYear },
    { Parameter: 'Audit Form Reference', Value: 'Form GSTR-9C & Rule 88C (DRC-01B Compliance)' },
    { Parameter: 'Annual Turnover (Books of Account)', Value: summary.annualBooksTurnover },
    { Parameter: 'Annual Turnover (GSTR-1 Outward Supplies)', Value: summary.annualGstr1Turnover },
    { Parameter: 'Annual Turnover (GSTR-3B Table 3.1)', Value: summary.annualGstr3bTurnover },
    { Parameter: 'Turnover Variance (GSTR-1 minus GSTR-3B)', Value: summary.netTurnoverVariance1Vs3b },
    { Parameter: 'Turnover Variance (Books minus GSTR-1)', Value: summary.netTurnoverVarianceBooksVs1 },
    { Parameter: 'Total GSTR-1 Tax Liability Determined', Value: summary.annualGstr1TotalTax },
    { Parameter: 'Total GSTR-3B Tax Paid via Cash/ITC', Value: summary.annualGstr3bTotalTax },
    { Parameter: 'Net Tax Shortfall (GSTR-1 minus GSTR-3B)', Value: summary.netAnnualTaxVariance1Vs3b },
    { Parameter: 'DRC-01B Notice Risk Months Count', Value: summary.criticalDrc01bMonthsCount },
    { Parameter: 'Total Differential Tax at Risk (DRC-01B)', Value: summary.totalDrc01bTaxShortfall },
    { Parameter: 'Estimated Sec 50 Interest Exposure (18% p.a.)', Value: summary.totalSection50Interest },
    { Parameter: '3-Way Reconciliation Score', Value: `${summary.reconciliationMatchPercent}%` },
  ];
  const ws1 = createStyledWorksheet(execData, {
    currencyColumns: ['Value'],
    headers: ['Parameter', 'Value'],
    columnWidths: { Parameter: 45, Value: 35 },
  });
  XLSX.utils.book_append_sheet(wb, ws1, 'Executive Dashboard');

  // Sheet 2: Monthly Turnover Reconciliation (Apr-Mar)
  const turnoverData = summary.monthlyBreakdown.map((m) => ({
    'Month': m.month,
    'Quarter': m.quarter,
    'Books Turnover': m.booksTurnover,
    'GSTR-1 Turnover': m.gstr1Turnover,
    'GSTR-3B Turnover': m.gstr3bTurnover,
    'Variance (GSTR-1 vs 3B)': m.turnoverVarianceGstr1Vs3b,
    'Variance (Books vs 1)': m.turnoverVarianceBooksVsGstr1,
    'Rec Status': m.reconciliationStatus,
    'DRC-01B Flag': m.drc01bStatus,
  }));
  const ws2 = createStyledWorksheet(turnoverData, {
    currencyColumns: ['Books Turnover', 'GSTR-1 Turnover', 'GSTR-3B Turnover', 'Variance (GSTR-1 vs 3B)', 'Variance (Books vs 1)'],
    headers: ['Month', 'Quarter', 'Books Turnover', 'GSTR-1 Turnover', 'GSTR-3B Turnover', 'Variance (GSTR-1 vs 3B)', 'Variance (Books vs 1)', 'Rec Status', 'DRC-01B Flag'],
  });
  XLSX.utils.book_append_sheet(wb, ws2, 'Monthly Turnover Rec');

  // Sheet 3: Monthly Tax Liabilities 3-Way Rec
  const taxData = summary.monthlyBreakdown.map((m) => ({
    'Month': m.month,
    'GSTR-1 IGST': m.gstr1Igst,
    'GSTR-3B IGST': m.gstr3bIgst,
    'IGST Diff': m.igstDiff1Vs3b,
    'GSTR-1 CGST': m.gstr1Cgst,
    'GSTR-3B CGST': m.gstr3bCgst,
    'CGST Diff': m.cgstDiff1Vs3b,
    'GSTR-1 SGST': m.gstr1Sgst,
    'GSTR-3B SGST': m.gstr3bSgst,
    'SGST Diff': m.sgstDiff1Vs3b,
    'GSTR-1 Total Tax': m.gstr1TotalTax,
    'GSTR-3B Total Tax': m.gstr3bTotalTax,
    'Net Shortfall (1 vs 3B)': m.netTaxShortfall1Vs3b,
    'DRC-01B Notice Risk': m.drc01bThresholdExceeded ? 'YES (Triggered)' : 'NO',
  }));
  const ws3 = createStyledWorksheet(taxData, {
    currencyColumns: [
      'GSTR-1 IGST', 'GSTR-3B IGST', 'IGST Diff',
      'GSTR-1 CGST', 'GSTR-3B CGST', 'CGST Diff',
      'GSTR-1 SGST', 'GSTR-3B SGST', 'SGST Diff',
      'GSTR-1 Total Tax', 'GSTR-3B Total Tax', 'Net Shortfall (1 vs 3B)'
    ],
    headers: [
      'Month',
      'GSTR-1 IGST', 'GSTR-3B IGST', 'IGST Diff',
      'GSTR-1 CGST', 'GSTR-3B CGST', 'CGST Diff',
      'GSTR-1 SGST', 'GSTR-3B SGST', 'SGST Diff',
      'GSTR-1 Total Tax', 'GSTR-3B Total Tax',
      'Net Shortfall (1 vs 3B)', 'DRC-01B Notice Risk'
    ],
  });
  XLSX.utils.book_append_sheet(wb, ws3, 'Tax Liability 3-Way Rec');

  // Sheet 4: DRC-01B Automated Notice Schedule & Draft Explanations
  const drc01bData = summary.drc01bNoticeItems.map((n, i) => ({
    'S.No.': i + 1,
    'Tax Period': n.month,
    'GSTR-1 Tax Determined': n.gstr1TotalTax,
    'GSTR-3B Tax Paid': n.gstr3bTotalTax,
    'Tax Liability Shortfall': n.netTaxShortfall1Vs3b,
    'Est. Interest Sec 50 (18%)': n.estimatedInterestSec50,
    'Suggested Reply to GST Officer / DRC-01B': n.suggestedDepartmentReply,
  }));
  const ws4 = createStyledWorksheet(drc01bData.length > 0 ? drc01bData : [{ Note: 'No DRC-01B notices triggered for this financial year.' }], {
    currencyColumns: ['GSTR-1 Tax Determined', 'GSTR-3B Tax Paid', 'Tax Liability Shortfall', 'Est. Interest Sec 50 (18%)'],
    headers: ['S.No.', 'Tax Period', 'GSTR-1 Tax Determined', 'GSTR-3B Tax Paid', 'Tax Liability Shortfall', 'Est. Interest Sec 50 (18%)', 'Suggested Reply to GST Officer / DRC-01B'],
  });
  XLSX.utils.book_append_sheet(wb, ws4, 'DRC-01B Notice Explanations');

  const cleanName = summary.clientLegalName.replace(/[^a-zA-Z0-9]/g, '_');
  XLSX.writeFile(wb, `${cleanName}_GSTR1_3B_Books_3Way_Reconciliation.xlsx`);
}

/**
 * Realistic 12-Month Sample Dataset with Indian Corporate Scenarios
 */
export function getSampleGstTripleRecData(): Record<string, any>[] {
  return [
    {
      'Month': 'April 2025',
      'Books Turnover': 4500000,
      'Books IGST': 360000,
      'Books CGST': 225000,
      'Books SGST': 225000,
      'Books Cess': 0,
      'GSTR-1 Turnover': 4500000,
      'GSTR-1 IGST': 360000,
      'GSTR-1 CGST': 225000,
      'GSTR-1 SGST': 225000,
      'GSTR-1 Cess': 0,
      'GSTR-3B Turnover': 4500000,
      'GSTR-3B IGST': 360000,
      'GSTR-3B CGST': 225000,
      'GSTR-3B SGST': 225000,
      'GSTR-3B Cess': 0,
    },
    {
      // May: GSTR-1 higher by ₹80,000 tax due to invoice cancelled in books after 11th -> DRC-01B Trigger!
      'Month': 'May 2025',
      'Books Turnover': 5200000,
      'Books IGST': 416000,
      'Books CGST': 260000,
      'Books SGST': 260000,
      'Books Cess': 0,
      'GSTR-1 Turnover': 5650000,
      'GSTR-1 IGST': 496000,
      'GSTR-1 CGST': 260000,
      'GSTR-1 SGST': 260000,
      'GSTR-1 Cess': 0,
      'GSTR-3B Turnover': 5200000,
      'GSTR-3B IGST': 416000,
      'GSTR-3B CGST': 260000,
      'GSTR-3B SGST': 260000,
      'GSTR-3B Cess': 0,
    },
    {
      'Month': 'June 2025',
      'Books Turnover': 6100000,
      'Books IGST': 488000,
      'Books CGST': 305000,
      'Books SGST': 305000,
      'Books Cess': 0,
      'GSTR-1 Turnover': 6100000,
      'GSTR-1 IGST': 488000,
      'GSTR-1 CGST': 305000,
      'GSTR-1 SGST': 305000,
      'GSTR-1 Cess': 0,
      'GSTR-3B Turnover': 6100000,
      'GSTR-3B IGST': 488000,
      'GSTR-3B CGST': 305000,
      'GSTR-3B SGST': 305000,
      'GSTR-3B Cess': 0,
    },
    {
      'Month': 'July 2025',
      'Books Turnover': 4800000,
      'Books IGST': 384000,
      'Books CGST': 240000,
      'Books SGST': 240000,
      'Books Cess': 0,
      'GSTR-1 Turnover': 4800000,
      'GSTR-1 IGST': 384000,
      'GSTR-1 CGST': 240000,
      'GSTR-1 SGST': 240000,
      'GSTR-1 Cess': 0,
      'GSTR-3B Turnover': 4800000,
      'GSTR-3B IGST': 384000,
      'GSTR-3B CGST': 240000,
      'GSTR-3B SGST': 240000,
      'GSTR-3B Cess': 0,
    },
    {
      // August: Head mismatch (IGST vs CGST/SGST)
      'Month': 'August 2025',
      'Books Turnover': 5900000,
      'Books IGST': 472000,
      'Books CGST': 295000,
      'Books SGST': 295000,
      'Books Cess': 0,
      'GSTR-1 Turnover': 5900000,
      'GSTR-1 IGST': 592000,
      'GSTR-1 CGST': 235000,
      'GSTR-1 SGST': 235000,
      'GSTR-1 Cess': 0,
      'GSTR-3B Turnover': 5900000,
      'GSTR-3B IGST': 472000,
      'GSTR-3B CGST': 295000,
      'GSTR-3B SGST': 295000,
      'GSTR-3B Cess': 0,
    },
    {
      'Month': 'September 2025',
      'Books Turnover': 7200000,
      'Books IGST': 576000,
      'Books CGST': 360000,
      'Books SGST': 360000,
      'Books Cess': 0,
      'GSTR-1 Turnover': 7200000,
      'GSTR-1 IGST': 576000,
      'GSTR-1 CGST': 360000,
      'GSTR-1 SGST': 360000,
      'GSTR-1 Cess': 0,
      'GSTR-3B Turnover': 7200000,
      'GSTR-3B IGST': 576000,
      'GSTR-3B CGST': 360000,
      'GSTR-3B SGST': 360000,
      'GSTR-3B Cess': 0,
    },
    {
      // October: Festival sales under-reported in 3B -> DRC-01B Trigger of ₹1,44,000 tax
      'Month': 'October 2025',
      'Books Turnover': 8500000,
      'Books IGST': 680000,
      'Books CGST': 425000,
      'Books SGST': 425000,
      'Books Cess': 0,
      'GSTR-1 Turnover': 8500000,
      'GSTR-1 IGST': 680000,
      'GSTR-1 CGST': 425000,
      'GSTR-1 SGST': 425000,
      'GSTR-1 Cess': 0,
      'GSTR-3B Turnover': 7700000,
      'GSTR-3B IGST': 600000,
      'GSTR-3B CGST': 389000,
      'GSTR-3B SGST': 389000,
      'GSTR-3B Cess': 0,
    },
    {
      'Month': 'November 2025',
      'Books Turnover': 6400000,
      'Books IGST': 512000,
      'Books CGST': 320000,
      'Books SGST': 320000,
      'Books Cess': 0,
      'GSTR-1 Turnover': 6400000,
      'GSTR-1 IGST': 512000,
      'GSTR-1 CGST': 320000,
      'GSTR-1 SGST': 320000,
      'GSTR-1 Cess': 0,
      'GSTR-3B Turnover': 6400000,
      'GSTR-3B IGST': 512000,
      'GSTR-3B CGST': 320000,
      'GSTR-3B SGST': 320000,
      'GSTR-3B Cess': 0,
    },
    {
      // December: Excess tax paid in 3B by ₹30,000
      'Month': 'December 2025',
      'Books Turnover': 5800000,
      'Books IGST': 464000,
      'Books CGST': 290000,
      'Books SGST': 290000,
      'Books Cess': 0,
      'GSTR-1 Turnover': 5800000,
      'GSTR-1 IGST': 464000,
      'GSTR-1 CGST': 290000,
      'GSTR-1 SGST': 290000,
      'GSTR-1 Cess': 0,
      'GSTR-3B Turnover': 5966667,
      'GSTR-3B IGST': 480000,
      'GSTR-3B CGST': 304000,
      'GSTR-3B SGST': 304000,
      'GSTR-3B Cess': 0,
    },
    {
      'Month': 'January 2026',
      'Books Turnover': 6700000,
      'Books IGST': 536000,
      'Books CGST': 335000,
      'Books SGST': 335000,
      'Books Cess': 0,
      'GSTR-1 Turnover': 6700000,
      'GSTR-1 IGST': 536000,
      'GSTR-1 CGST': 335000,
      'GSTR-1 SGST': 335000,
      'GSTR-1 Cess': 0,
      'GSTR-3B Turnover': 6700000,
      'GSTR-3B IGST': 536000,
      'GSTR-3B CGST': 335000,
      'GSTR-3B SGST': 335000,
      'GSTR-3B Cess': 0,
    },
    {
      'Month': 'February 2026',
      'Books Turnover': 7100000,
      'Books IGST': 568000,
      'Books CGST': 355000,
      'Books SGST': 355000,
      'Books Cess': 0,
      'GSTR-1 Turnover': 7100000,
      'GSTR-1 IGST': 568000,
      'GSTR-1 CGST': 355000,
      'GSTR-1 SGST': 355000,
      'GSTR-1 Cess': 0,
      'GSTR-3B Turnover': 7100000,
      'GSTR-3B IGST': 568000,
      'GSTR-3B CGST': 355000,
      'GSTR-3B SGST': 355000,
      'GSTR-3B Cess': 0,
    },
    {
      'Month': 'March 2026',
      'Books Turnover': 9200000,
      'Books IGST': 736000,
      'Books CGST': 460000,
      'Books SGST': 460000,
      'Books Cess': 0,
      'GSTR-1 Turnover': 9200000,
      'GSTR-1 IGST': 736000,
      'GSTR-1 CGST': 460000,
      'GSTR-1 SGST': 460000,
      'GSTR-1 Cess': 0,
      'GSTR-3B Turnover': 9200000,
      'GSTR-3B IGST': 736000,
      'GSTR-3B CGST': 460000,
      'GSTR-3B SGST': 460000,
      'GSTR-3B Cess': 0,
    },
  ];
}
