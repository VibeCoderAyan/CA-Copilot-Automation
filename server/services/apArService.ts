/**
 * Module 2: Accounts Payable & Accounts Receivable Automation Service
 * Monitors supplier payments, customer collections, MSMED compliance, and duplicate payment risks
 */

import type { AgingRecord, PaymentRiskAlert, FirmContext } from '../types';
import { BackendStore } from '../db/store';
import { ExportService } from './exportService';
import { parseAmount, parseDate, parseDateObj } from '../../src/utils/dataSanitizer';

export class ApArAutomationService {
  private store: BackendStore;
  private exportService: ExportService;

  constructor(store?: BackendStore) {
    this.store = store || BackendStore.getInstance();
    this.exportService = new ExportService(this.store);
  }

  /**
   * Processes raw AP/AR records, computes aging buckets, and flags duplicate payment risks
   */
  public processLedgerRecords(
    context: FirmContext,
    records: Record<string, any>[],
    type: 'payable' | 'receivable',
    asOfDateStr: string = '2026-09-30'
  ): {
    agingRecords: AgingRecord[];
    riskAlerts: PaymentRiskAlert[];
    summary: {
      totalGross: number;
      totalPaid: number;
      totalOutstanding: number;
      overdueMsmeCount: number;
      duplicateRiskCount: number;
    };
  } {
    const asOfDateObj = parseDateObj(asOfDateStr) || new Date();
    const asOfTime = asOfDateObj.getTime();
    const agingRecords: AgingRecord[] = [];
    const riskAlerts: PaymentRiskAlert[] = [];

    const seenInvoices = new Map<string, { party: string; amount: number; id: string }>();

    for (let idx = 0; idx < records.length; idx++) {
      const row = records[idx];
      const party = (
        row['Party Name'] ||
        row.Vendor ||
        row.Customer ||
        row.Debtor ||
        row.Creditor ||
        row['Sundry Debtors'] ||
        row['Sundry Creditors'] ||
        row['Debtor Name'] ||
        row['Creditor Name'] ||
        row['Client Name'] ||
        row.Client ||
        row.Party ||
        'Party'
      ).toString().trim();
      const gstin = (
        row.GSTIN ||
        row.gstin ||
        row['Party GSTIN'] ||
        row['Supplier GSTIN'] ||
        row['Customer GSTIN'] ||
        row['GST No'] ||
        ''
      ).toString().trim().toUpperCase();
      const invNo = (
        row['Invoice No'] ||
        row.BillNo ||
        row.InvoiceNumber ||
        row['Invoice Number'] ||
        row['Bill No'] ||
        row['Voucher No'] ||
        row['Ref No'] ||
        `INV-${1000 + idx}`
      ).toString().trim();
      const invDateObj = parseDateObj(row['Invoice Date'] || row.Date || row['Bill Date'] || row['Txn Date'] || '2026-09-01');
      const invTime = invDateObj ? invDateObj.getTime() : asOfTime;
      const invDate = parseDate(invDateObj || new Date(invTime), 'YYYY-MM-DD');
      const termsDays = parseInt(row['Credit Days'] || row.Terms || row['Credit Period'] || '30', 10) || 30;

      const gross = parseAmount(row['Gross Amount'] || row.Total || row.Amount || row['Bill Amount'] || row['Invoice Amount'] || row.Debit || 0);
      const paid = parseAmount(row['Paid Amount'] || row.Paid || row.Settled || row['Received Amount'] || row.Credit || 0);
      const balance = parseAmount(row['Outstanding Balance'] || row.Balance || row['Outstanding Amount'] || row['Net Balance'] || row['Due Amount'] || (gross - paid));

      // Calculate Due Date & Age
      const rawDueDate = row['Due Date'] || row.DueDate || row.dueDate || row['Due date'] || row['Payment Due Date'];
      const explicitDueDateObj = rawDueDate ? parseDateObj(rawDueDate) : null;
      let dueDate = '';
      if (explicitDueDateObj && !isNaN(explicitDueDateObj.getTime())) {
        dueDate = parseDate(explicitDueDateObj, 'YYYY-MM-DD');
      } else {
        const dueTime = invTime + termsDays * 24 * 60 * 60 * 1000;
        dueDate = new Date(dueTime).toISOString().split('T')[0];
      }
      const ageDays = Math.max(0, Math.round((asOfTime - invTime) / (1000 * 60 * 60 * 24)));


      // Aging Bucket
      let bucketLabel = '0-30 Days';
      if (ageDays > 180) bucketLabel = '180+ Days';
      else if (ageDays > 90) bucketLabel = '91-180 Days';
      else if (ageDays > 60) bucketLabel = '61-90 Days';
      else if (ageDays > 30) bucketLabel = '31-60 Days';

      // MSME Overdue (> 45 days under Section 15 of MSMED Act)
      const isMsmeOverdue = type === 'payable' && ageDays > 45 && balance > 0;

      // Duplicate Payment Risk Check
      const dedupKey = `${party.toLowerCase()}|${invNo.toUpperCase()}`;
      let duplicateRisk = false;
      if (type === 'payable' && seenInvoices.has(dedupKey)) {
        duplicateRisk = true;
        riskAlerts.push({
          id: `risk-dup-${idx}`,
          firmId: context.firmId,
          clientId: context.clientId,
          vendorName: party,
          invoiceNo: invNo,
          amount: gross,
          riskType: 'duplicate_invoice_number',
          description: `Duplicate invoice "${invNo}" detected for vendor "${party}" with amount ₹${gross.toLocaleString('en-IN')}`,
          severity: 'critical',
          suggestedAction: 'Hold payment release. Verify physical voucher to prevent double-payment.',
        });
      } else {
        seenInvoices.set(dedupKey, { party, amount: gross, id: `aging-${idx}` });
      }

      if (isMsmeOverdue) {
        riskAlerts.push({
          id: `risk-msme-${idx}`,
          firmId: context.firmId,
          clientId: context.clientId,
          vendorName: party,
          invoiceNo: invNo,
          amount: balance,
          riskType: 'msme_45_day_violation',
          description: `MSME supplier invoice "${invNo}" is overdue by ${ageDays} days (exceeds 45-day statutory limit under Section 15 of MSMED Act)`,
          severity: 'warning',
          suggestedAction: 'Prioritize payment to avoid statutory penal interest under Section 16 MSMED Act.',
        });
      }

      agingRecords.push({
        id: `aging-${type}-${Date.now()}-${idx + 1}`,
        firmId: context.firmId,
        clientId: context.clientId,
        type,
        partyName: party,
        gstin,
        invoiceNo: invNo,
        invoiceDate: invDate || '2026-09-01',
        dueDate,
        ageDays,
        grossAmount: gross,
        paidAmount: paid,
        outstandingBalance: balance,
        bucketLabel,
        isCreditBalance: balance < 0,
        isMsmeOverdue,
        duplicateRisk,
      });
    }

    this.store.saveAgingRecords(context.clientId, agingRecords);
    this.store.savePaymentRiskAlerts(context.clientId, riskAlerts);

    const totalGross = agingRecords.reduce((acc, r) => acc + r.grossAmount, 0);
    const totalPaid = agingRecords.reduce((acc, r) => acc + r.paidAmount, 0);
    const totalOutstanding = agingRecords.reduce((acc, r) => acc + r.outstandingBalance, 0);

    return {
      agingRecords,
      riskAlerts,
      summary: {
        totalGross,
        totalPaid,
        totalOutstanding,
        overdueMsmeCount: agingRecords.filter((r) => r.isMsmeOverdue).length,
        duplicateRiskCount: riskAlerts.filter((r) => r.riskType === 'duplicate_invoice_number').length,
      },
    };
  }

  /**
   * Generates Payment Review Worksheet without releasing funds
   */
  public generatePaymentReviewList(context: FirmContext, maxPaymentAmount?: number) {
    const payables = this.store.getAgingRecords(context.clientId).filter((r) => r.type === 'payable');
    const alerts = this.store.getPaymentRiskAlerts(context.clientId);

    const reviewRows = payables.map((p) => {
      const hasDuplicateRisk = alerts.some((a) => a.invoiceNo === p.invoiceNo && a.riskType === 'duplicate_invoice_number');
      const isHold = hasDuplicateRisk || p.outstandingBalance <= 0;

      return {
        'Vendor Name': p.partyName,
        'Invoice No': p.invoiceNo,
        'Invoice Date': p.invoiceDate,
        'Due Date': p.dueDate,
        'Age (Days)': p.ageDays,
        'Outstanding (₹)': p.outstandingBalance,
        'MSME Flag': p.isMsmeOverdue ? 'Overdue > 45 Days' : 'Compliant',
        'Duplicate Risk': hasDuplicateRisk ? 'CRITICAL RISK - DO NOT PAY' : 'Clean',
        'Payment Decision': isHold ? 'ON HOLD' : 'APPROVED FOR MANUAL RUN',
        'Recommended Payment': isHold ? 0 : Math.min(p.outstandingBalance, maxPaymentAmount || p.outstandingBalance),
      };
    });

    return this.exportService.generateExcelWorkbook(
      context,
      `ap-review-${context.clientId}`,
      `AP_Payment_Review_Worksheet_${context.clientId}.xlsx`,
      [{ sheetName: 'Payment_Review', rows: reviewRows }]
    );
  }
}
