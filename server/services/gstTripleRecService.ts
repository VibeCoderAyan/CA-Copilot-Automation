/**
 * Backend Service: GSTR-1 vs GSTR-3B vs Books 3-Way Reconciliation & DRC-01B Engine
 */
import { BackendStore } from '../db/store.ts';
import { ExportService } from './exportService.ts';
import type { FirmContext, ExportArtifact } from '../types/index.ts';

export interface Monthly3WayInput {
  month: string;
  booksTurnover: number;
  booksIgst?: number;
  booksCgst?: number;
  booksSgst?: number;
  gstr1Turnover: number;
  gstr1Igst?: number;
  gstr1Cgst?: number;
  gstr1Sgst?: number;
  gstr3bTurnover: number;
  gstr3bIgst?: number;
  gstr3bCgst?: number;
  gstr3bSgst?: number;
}

export class GstTripleRecService {
  private exportService: ExportService;

  constructor(private store: BackendStore) {
    this.exportService = new ExportService(store);
  }

  public reconcile(
    firmId: string,
    clientId: string,
    data: Monthly3WayInput[],
    financialYear: string = '2025-26',
    drc01bThreshold: number = 25000
  ) {
    const records = data.map((d) => {
      const gstr1TotalTax = (d.gstr1Igst || 0) + (d.gstr1Cgst || 0) + (d.gstr1Sgst || 0);
      const gstr3bTotalTax = (d.gstr3bIgst || 0) + (d.gstr3bCgst || 0) + (d.gstr3bSgst || 0);
      const netTaxShortfall = gstr1TotalTax - gstr3bTotalTax;
      const isDrc01b = netTaxShortfall > drc01bThreshold;

      // Estimate interest @ 18% p.a.
      const estimatedInterest = isDrc01b ? Math.round((netTaxShortfall * 0.18 * 90) / 365) : 0;

      return {
        month: d.month,
        booksTurnover: d.booksTurnover,
        gstr1Turnover: d.gstr1Turnover,
        gstr3bTurnover: d.gstr3bTurnover,
        turnoverDiff1Vs3b: d.gstr1Turnover - d.gstr3bTurnover,
        turnoverDiffBooksVs1: d.booksTurnover - d.gstr1Turnover,
        gstr1TotalTax,
        gstr3bTotalTax,
        netTaxShortfall,
        isDrc01bTriggered: isDrc01b,
        estimatedInterest,
      };
    });

    const drc01bList = records.filter((r) => r.isDrc01bTriggered);

    return {
      firmId,
      clientId,
      financialYear,
      annualBooksTurnover: records.reduce((acc, r) => acc + r.booksTurnover, 0),
      annualGstr1Turnover: records.reduce((acc, r) => acc + r.gstr1Turnover, 0),
      annualGstr3bTurnover: records.reduce((acc, r) => acc + r.gstr3bTurnover, 0),
      annualGstr1Tax: records.reduce((acc, r) => acc + r.gstr1TotalTax, 0),
      annualGstr3bTax: records.reduce((acc, r) => acc + r.gstr3bTotalTax, 0),
      netAnnualTaxVariance: records.reduce((acc, r) => acc + r.netTaxShortfall, 0),
      drc01bNoticeMonthsCount: drc01bList.length,
      totalDrc01bTaxShortfall: drc01bList.reduce((acc, r) => acc + r.netTaxShortfall, 0),
      totalEstimatedInterest: drc01bList.reduce((acc, r) => acc + r.estimatedInterest, 0),
      records,
      drc01bList,
      generatedAt: new Date().toISOString(),
    };
  }

  /**
   * Generates a structured multi-sheet GSTR-1 vs 3B vs Books 3-Way Reconciliation Excel report
   */
  public exportTripleRecWorkbook(
    context: FirmContext,
    data: Monthly3WayInput[],
    financialYear: string = '2025-26',
    drc01bThreshold: number = 25000,
    clientLegalName: string = 'Apex Technologies Pvt Ltd',
    clientGstin: string = '27AABCU9603R1ZM'
  ): ExportArtifact {
    const summary = this.reconcile(context.firmId, context.clientId, data, financialYear, drc01bThreshold);

    // Sheet 1: Executive Dashboard
    const execRows = [
      { Parameter: 'Taxpayer Trade / Legal Name', Value: clientLegalName },
      { Parameter: 'GSTIN', Value: clientGstin },
      { Parameter: 'Financial Year (FY)', Value: financialYear },
      { Parameter: 'Statutory Standard', Value: 'Form GSTR-9C & Rule 88C (DRC-01B Compliance)' },
      { Parameter: 'Annual Turnover (Books of Account)', Value: summary.annualBooksTurnover },
      { Parameter: 'Annual Turnover (GSTR-1 Outward Supplies)', Value: summary.annualGstr1Turnover },
      { Parameter: 'Annual Turnover (GSTR-3B Table 3.1)', Value: summary.annualGstr3bTurnover },
      { Parameter: 'Turnover Variance (GSTR-1 minus GSTR-3B)', Value: summary.annualGstr1Turnover - summary.annualGstr3bTurnover },
      { Parameter: 'Total GSTR-1 Tax Liability Determined', Value: summary.annualGstr1Tax },
      { Parameter: 'Total GSTR-3B Tax Paid via Cash/ITC', Value: summary.annualGstr3bTax },
      { Parameter: 'Net Tax Shortfall (1 vs 3B)', Value: summary.netAnnualTaxVariance },
      { Parameter: 'DRC-01B Notice Risk Months Count', Value: summary.drc01bNoticeMonthsCount },
      { Parameter: 'Total Differential Tax at Risk (DRC-01B)', Value: summary.totalDrc01bTaxShortfall },
      { Parameter: 'Estimated Sec 50 Interest Exposure (18% p.a.)', Value: summary.totalEstimatedInterest },
    ];

    // Sheet 2: Monthly Turnover Reconciliation
    const turnoverRows = summary.records.map((r, idx) => ({
      'S.No': idx + 1,
      'Month': r.month,
      'Books Turnover': r.booksTurnover,
      'GSTR-1 Turnover': r.gstr1Turnover,
      'GSTR-3B Turnover': r.gstr3bTurnover,
      'Variance (1 vs 3B)': r.turnoverDiff1Vs3b,
      'Variance (Books vs 1)': r.turnoverDiffBooksVs1,
      'Status': Math.abs(r.turnoverDiff1Vs3b) < 1 ? 'Balanced' : 'Turnover Mismatch',
      'DRC-01B Notice Trigger': r.isDrc01bTriggered ? 'Triggered (Over Limit)' : 'Safe',
    }));

    // Sheet 3: Monthly Tax Liability Reconciliation
    const taxRows = summary.records.map((r, idx) => ({
      'S.No': idx + 1,
      'Month': r.month,
      'GSTR-1 Tax Determined': r.gstr1TotalTax,
      'GSTR-3B Tax Paid': r.gstr3bTotalTax,
      'Net Tax Shortfall (1 vs 3B)': r.netTaxShortfall,
      'DRC-01B Trigger': r.isDrc01bTriggered ? 'YES' : 'NO',
      'Est. Sec 50 Interest': r.estimatedInterest,
    }));

    // Sheet 4: DRC-01B Notice Schedule
    const drcRows = summary.drc01bList.map((d, idx) => ({
      'Notice #': idx + 1,
      'Tax Period': d.month,
      'GSTR-1 Liability': d.gstr1TotalTax,
      'GSTR-3B Paid': d.gstr3bTotalTax,
      'Shortfall at Risk': d.netTaxShortfall,
      'Estimated Interest (18%)': d.estimatedInterest,
      'Statutory Action Required': 'File DRC-03 or submit formal reply in Part B of Form DRC-01B within 7 days',
    }));

    const cleanClient = clientLegalName.replace(/[^a-zA-Z0-9]/g, '_');
    return this.exportService.generateExcelWorkbook(
      context,
      `gst-3way-${Date.now()}`,
      `${cleanClient}_GSTR1_3B_Books_3Way_Reconciliation.xlsx`,
      [
        { sheetName: 'Executive_Summary', rows: execRows, currencyColumns: ['Value'] },
        {
          sheetName: 'Monthly_Turnover_Rec',
          rows: turnoverRows,
          currencyColumns: ['Books Turnover', 'GSTR-1 Turnover', 'GSTR-3B Turnover', 'Variance (1 vs 3B)', 'Variance (Books vs 1)'],
        },
        {
          sheetName: 'Monthly_Tax_Liability_Rec',
          rows: taxRows,
          currencyColumns: ['GSTR-1 Tax Determined', 'GSTR-3B Tax Paid', 'Net Tax Shortfall (1 vs 3B)', 'Est. Sec 50 Interest'],
        },
        {
          sheetName: 'DRC01B_Notice_Schedule',
          rows: drcRows.length > 0 ? drcRows : [{ Notice: 'No DRC-01B notices triggered for this financial year.' }],
          currencyColumns: ['GSTR-1 Liability', 'GSTR-3B Paid', 'Shortfall at Risk', 'Estimated Interest (18%)'],
        },
      ]
    );
  }
}
