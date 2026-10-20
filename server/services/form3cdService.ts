/**
 * Backend Service: Form 3CD Tax Audit Scrutiny & Clause Generator (Section 44AB)
 */
import { BackendStore } from '../db/store.ts';
import { ExportService } from './exportService.ts';
import type { FirmContext, ExportArtifact } from '../types/index.ts';

export interface Form3cdVoucherPayload {
  voucherNo: string;
  date: string;
  accountName: string;
  payeeName: string;
  amount: number;
  paymentMode: string;
  narration?: string;
  pan?: string;
  gstin?: string;
}

export class Form3cdService {
  private exportService: ExportService;

  constructor(private store: BackendStore) {
    this.exportService = new ExportService(store);
  }

  public scrutinizeLedger(
    firmId: string,
    clientId: string,
    vouchers: Form3cdVoucherPayload[],
    financialYear: string = '2025-26',
    assessmentYear: string = '2026-27'
  ) {
    // 1. Clause 21(d) / Sec 40A(3) Cash Payments Scrutiny
    const cashVouchers = vouchers.filter((v) =>
      v.paymentMode?.toLowerCase().includes('cash') ||
      v.paymentMode?.toLowerCase().includes('bearer') ||
      v.narration?.toLowerCase().includes('cash paid')
    );

    const dailyPayeeMap: Record<string, Form3cdVoucherPayload[]> = {};
    for (const cv of cashVouchers) {
      const key = `${cv.date.trim()}||${cv.payeeName.trim().toLowerCase()}`;
      if (!dailyPayeeMap[key]) dailyPayeeMap[key] = [];
      dailyPayeeMap[key].push(cv);
    }

    const cashViolations: any[] = [];
    let totalDisallowable40a3 = 0;

    for (const [key, vList] of Object.entries(dailyPayeeMap)) {
      const [date, payee] = key.split('||');
      const dailyTotal = vList.reduce((acc, v) => acc + v.amount, 0);
      const isTransporter = vList.some((v) =>
        v.accountName?.toLowerCase().includes('freight') ||
        v.accountName?.toLowerCase().includes('transport') ||
        v.narration?.toLowerCase().includes('truck')
      );
      const threshold = isTransporter ? 35000 : 10000;

      if (dailyTotal > threshold) {
        // Check Rule 6DD
        const allNarrations = vList.map((v) => (v.narration || '').toLowerCase()).join(' ');
        const isRule6Dd =
          allNarrations.includes('bank holiday') ||
          allNarrations.includes('agricultural') ||
          allNarrations.includes('terminal benefit') ||
          allNarrations.includes('govt');

        if (!isRule6Dd) {
          totalDisallowable40a3 += dailyTotal;
        }

        cashViolations.push({
          date,
          payeeName: vList[0].payeeName,
          voucherNumbers: vList.map((v) => v.voucherNo),
          dailyAggregateAmount: dailyTotal,
          thresholdLimit: threshold,
          disallowableAmount: isRule6Dd ? 0 : dailyTotal,
          isTransporter,
          isRule6DdExempt: isRule6Dd,
          exemptionReason: isRule6Dd ? 'Rule 6DD exemption noted' : 'Section 40A(3) cash breach',
        });
      }
    }

    // 2. Clause 44 Breakdown (Registered vs Unregistered)
    const headMap: Record<string, { total: number; registered: number; unregistered: number }> = {};
    for (const v of vouchers) {
      const head = v.accountName || 'General Expenses';
      if (!headMap[head]) headMap[head] = { total: 0, registered: 0, unregistered: 0 };
      headMap[head].total += v.amount;

      if (v.gstin && v.gstin.length === 15) {
        headMap[head].registered += v.amount;
      } else {
        headMap[head].unregistered += v.amount;
      }
    }

    const clause44Breakdown = Object.entries(headMap).map(([head, d], i) => ({
      sNo: i + 1,
      headOfExpenditure: head,
      totalExpenditure: d.total,
      registeredExpenditure: d.registered,
      unregisteredExpenditure: d.unregistered,
      unregisteredPercent: d.total > 0 ? Number(((d.unregistered / d.total) * 100).toFixed(2)) : 0,
    }));

    const totalClause44 = clause44Breakdown.reduce((acc, r) => acc + r.totalExpenditure, 0);
    const totalUnreg = clause44Breakdown.reduce((acc, r) => acc + r.unregisteredExpenditure, 0);
    const unregRatio = totalClause44 > 0 ? Number(((totalUnreg / totalClause44) * 100).toFixed(2)) : 0;

    // 3. Clause 21(a) Penalties & Fines
    const penalties = vouchers.filter((v) => {
      const text = `${v.accountName} ${v.narration}`.toLowerCase();
      return text.includes('penalty') || text.includes('fine') || text.includes('compounding') || text.includes('personal');
    });
    const totalPenalties = penalties.reduce((acc, v) => acc + v.amount, 0);

    const totalTaxAdditions = totalDisallowable40a3 + totalPenalties;

    return {
      firmId,
      clientId,
      financialYear,
      assessmentYear,
      totalScrutinizedAmount: vouchers.reduce((acc, v) => acc + v.amount, 0),
      clause40a3TotalDisallowable: totalDisallowable40a3,
      clause40a3ViolationsCount: cashViolations.filter((c) => !c.isRule6DdExempt).length,
      cashViolations,
      clause44Breakdown,
      totalClause44Expenditure: totalClause44,
      totalUnregisteredExpenditure: totalUnreg,
      unregisteredExpenditureRatio: unregRatio,
      clause21aPenaltiesTotal: totalPenalties,
      totalTaxAdditions,
      generatedAt: new Date().toISOString(),
    };
  }

  /**
   * Generates a structured multi-sheet Form 3CD Annexures Excel report
   */
  public exportForm3cdWorkbook(
    context: FirmContext,
    vouchers: Form3cdVoucherPayload[],
    financialYear: string = '2025-26',
    assessmentYear: string = '2026-27',
    clientLegalName: string = 'Apex Technologies Pvt Ltd'
  ): ExportArtifact {
    const summary = this.scrutinizeLedger(context.firmId, context.clientId, vouchers, financialYear, assessmentYear);

    // Sheet 1: Executive Tax Audit Summary
    const execSummaryRows = [
      { Parameter: 'Client Legal Entity Name', Value: clientLegalName },
      { Parameter: 'Client Identification Code', Value: context.clientId },
      { Parameter: 'Financial Year (FY)', Value: financialYear },
      { Parameter: 'Assessment Year (AY)', Value: assessmentYear },
      { Parameter: 'Statutory Form & Scope', Value: 'Form 3CD Tax Audit Scrutiny u/s 44AB' },
      { Parameter: 'Total Scrutinized Expenditure', Value: summary.totalScrutinizedAmount },
      { Parameter: 'Clause 21(d) / Sec 40A(3) Disallowable Cash Additions', Value: summary.clause40a3TotalDisallowable },
      { Parameter: 'Clause 21(a) Penalties, Fines & Personal Expenses', Value: summary.clause21aPenaltiesTotal },
      { Parameter: 'TOTAL INADMISSIBLE ADDITIONS TO TAXABLE PROFIT', Value: summary.totalTaxAdditions },
      { Parameter: 'Clause 44 Unregistered Procurement Ratio', Value: `${summary.unregisteredExpenditureRatio}%` },
    ];

    // Sheet 2: Clause 21(d) / Sec 40A(3) Cash Violations
    const cashRows = summary.cashViolations.map((c: any, idx: number) => ({
      'S.No': idx + 1,
      'Payment Date': c.date,
      'Payee Name': c.payeeName,
      'Vouchers Involved': c.voucherNumbers?.join(', ') || `VCH-${idx + 1}`,
      'Nature of Payment': c.isTransporter ? 'Transporter (Limit ₹35,000)' : 'General Payee (Limit ₹10,000)',
      'Daily Aggregate Amount': c.dailyAggregateAmount,
      'Statutory Threshold': c.thresholdLimit,
      'Rule 6DD Exemption': c.isRule6DdExempt ? 'YES (Exempt)' : 'NO',
      'Disallowance Sec 40A(3)': c.disallowableAmount,
      'Auditor Remarks': c.exemptionReason || (c.isRule6DdExempt ? 'Exempt under Rule 6DD' : 'Disallowable addition under Section 40A(3)'),
    }));

    // Sheet 3: Clause 44 GST Breakdown
    const c44Rows = summary.clause44Breakdown.map((r: any) => ({
      'S.No': r.sNo,
      'Expenditure Head': r.headOfExpenditure,
      'Total Expenditure': r.totalExpenditure,
      'Registered Suppliers Spend': r.registeredExpenditure,
      'Unregistered Entities Spend': r.unregisteredExpenditure,
      '% Unregistered': `${r.unregisteredPercent}%`,
    }));

    // Sheet 4: Clause 21(a) Penalties & Fines
    const penalties = vouchers.filter((v) => {
      const text = `${v.accountName} ${v.narration}`.toLowerCase();
      return text.includes('penalty') || text.includes('fine') || text.includes('compounding') || text.includes('personal');
    });
    const penaltyRows = penalties.map((p, idx) => ({
      'S.No': idx + 1,
      'Date': p.date,
      'Voucher No': p.voucherNo,
      'Account Head': p.accountName,
      'Narration': p.narration || '',
      'Disallowable Amount': p.amount,
      'Auditor Defense': 'Statutory penalty/fine inadmissible under Section 37(1) Explanation 1',
    }));

    const cleanClient = clientLegalName.replace(/[^a-zA-Z0-9]/g, '_');
    return this.exportService.generateExcelWorkbook(
      context,
      `tax-audit-3cd-${Date.now()}`,
      `${cleanClient}_Form3CD_Tax_Audit_Report.xlsx`,
      [
        { sheetName: 'Executive_Tax_Audit_Summary', rows: execSummaryRows, currencyColumns: ['Value'] },
        {
          sheetName: 'Clause_21d_Cash_40A3',
          rows: cashRows.length > 0 ? cashRows : [{ Notice: 'No Section 40A(3) cash breaches detected' }],
          currencyColumns: ['Daily Aggregate Amount', 'Statutory Threshold', 'Disallowance Sec 40A(3)'],
        },
        {
          sheetName: 'Clause_44_GST_Spend',
          rows: c44Rows.length > 0 ? c44Rows : [{ Notice: 'No Clause 44 expenditure lines' }],
          currencyColumns: ['Total Expenditure', 'Registered Suppliers Spend', 'Unregistered Entities Spend'],
        },
        {
          sheetName: 'Clause_21a_Penalties',
          rows: penaltyRows.length > 0 ? penaltyRows : [{ Notice: 'No Section 37(1) penalties or personal expenses detected' }],
          currencyColumns: ['Disallowable Amount'],
        },
      ]
    );
  }
}
