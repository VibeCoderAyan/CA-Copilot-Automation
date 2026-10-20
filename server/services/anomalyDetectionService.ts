/**
 * Module 4: Accounting Anomaly Detection Engine
 * Configurable, deterministic rules detecting suspicious journal entries, round sums, cash limits, and velocity spikes
 */

import type { AnomalyAlert, FirmContext, ExtractedRecord } from '../types';
import { BackendStore } from '../db/store';
import { ExportService } from './exportService';
import { parseDateObj } from '../../src/utils/dataSanitizer';

export class AnomalyDetectionService {
  private store: BackendStore;
  private exportService: ExportService;

  constructor(store?: BackendStore) {
    this.store = store || BackendStore.getInstance();
    this.exportService = new ExportService(this.store);
  }

  /**
   * Scans a batch of extracted financial records against all active anomaly rules
   */
  public scanTransactions(
    context: FirmContext,
    records: ExtractedRecord[]
  ): { alerts: AnomalyAlert[]; newAlertsCount: number } {
    const rules = this.store.listAnomalyRules(context.firmId).filter((r) => r.enabled);
    const existingAlerts = this.store.listAnomalyAlerts(context.clientId);
    const existingKeys = new Set(
      existingAlerts.map((a) => `${a.ruleCode}|${a.voucherNo}|${a.amount}`)
    );

    const generatedAlerts: AnomalyAlert[] = [];

    // Helper: robustly detect Sunday using standard accounting date parser
    const isSunday = (dStr?: string): boolean => {
      if (!dStr) return false;
      const d = parseDateObj(dStr);
      return d ? d.getDay() === 0 : false;
    };

    // 1. Rule Evaluation Loop
    for (let i = 0; i < records.length; i++) {
      const rec = records[i];
      const fields = rec.normalizedFields;
      const amount = fields.amount || Math.max(fields.debit || 0, fields.credit || 0);
      const vch = fields.invoiceNo || fields.refNo || `VCH-${rec.rowIndex}`;
      const narration = (fields.narration || '').toLowerCase();
      const account = fields.accountName || fields.partyName || '';

      for (const rule of rules) {
        // A. Duplicate Journal Check
        if (rule.category === 'journal_duplicate') {
          for (let j = i + 1; j < Math.min(records.length, i + 15); j++) {
            const nextRec = records[j];
            const nextAmt = nextRec.normalizedFields.amount || Math.max(nextRec.normalizedFields.debit || 0, nextRec.normalizedFields.credit || 0);
            if (
              amount > 0 &&
              amount === nextAmt &&
              fields.narration &&
              fields.narration.toLowerCase() === (nextRec.normalizedFields.narration || '').toLowerCase()
            ) {
              const dedupKey = `${rule.ruleCode}|${vch}|${amount}`;
              if (!existingKeys.has(dedupKey)) {
                generatedAlerts.push({
                  id: `alert-${Date.now()}-${generatedAlerts.length + 1}`,
                  firmId: context.firmId,
                  clientId: context.clientId,
                  ruleCode: rule.ruleCode,
                  ruleName: rule.name,
                  severity: rule.severity,
                  voucherNo: vch,
                  accountName: account,
                  amount,
                  date: fields.date,
                  explanation: `Duplicate-looking journal voucher: Same amount ₹${amount.toLocaleString('en-IN')} and identical narration "${fields.narration}" matches Row ${nextRec.rowIndex}`,
                  reviewStatus: 'open',
                  createdAt: new Date().toISOString(),
                });
                existingKeys.add(dedupKey);
              }
            }
          }
        }

        // B. Unusual Round Sum Expense
        if (rule.category === 'round_sum' && amount >= (rule.thresholdValue || 50000)) {
          if (amount % 10000 === 0 && (fields.debit || 0) > 0) {
            const dedupKey = `${rule.ruleCode}|${vch}|${amount}`;
            if (!existingKeys.has(dedupKey)) {
              generatedAlerts.push({
                id: `alert-${Date.now()}-${generatedAlerts.length + 1}`,
                firmId: context.firmId,
                clientId: context.clientId,
                ruleCode: rule.ruleCode,
                ruleName: rule.name,
                severity: rule.severity,
                voucherNo: vch,
                accountName: account,
                amount,
                date: fields.date,
                explanation: `Round-sum expenditure of exactly ₹${amount.toLocaleString('en-IN')} in "${account}". Common audit indicator of lump-sum provisions or un-vouched advances.`,
                reviewStatus: 'open',
                createdAt: new Date().toISOString(),
              });
              existingKeys.add(dedupKey);
            }
          }
        }

        // C. Section 40A(3) Cash Limit
        if (rule.category === 'cash_limit' && amount > (rule.thresholdValue || 10000)) {
          const isCash = narration.includes('cash') || (fields.refNo === 'CASH' && !narration.includes('chq'));
          if (isCash) {
            const dedupKey = `${rule.ruleCode}|${vch}|${amount}`;
            if (!existingKeys.has(dedupKey)) {
              generatedAlerts.push({
                id: `alert-${Date.now()}-${generatedAlerts.length + 1}`,
                firmId: context.firmId,
                clientId: context.clientId,
                ruleCode: rule.ruleCode,
                ruleName: rule.name,
                severity: rule.severity,
                voucherNo: vch,
                accountName: account,
                amount,
                date: fields.date,
                explanation: `Statutory Cash Limit Violation: Cash payment of ₹${amount.toLocaleString('en-IN')} exceeds threshold of ₹10,000 under Section 40A(3). Risk of 100% expenditure disallowance.`,
                reviewStatus: 'open',
                createdAt: new Date().toISOString(),
              });
              existingKeys.add(dedupKey);
            }
          }
        }

        // D. Weekend / Off-Hours Postings
        if (rule.category === 'weekend_posting' && amount >= (rule.thresholdValue || 100000)) {
          if (isSunday(fields.date)) {
            const dedupKey = `${rule.ruleCode}|${vch}|${amount}`;
            if (!existingKeys.has(dedupKey)) {
              generatedAlerts.push({
                id: `alert-${Date.now()}-${generatedAlerts.length + 1}`,
                firmId: context.firmId,
                clientId: context.clientId,
                ruleCode: rule.ruleCode,
                ruleName: rule.name,
                severity: rule.severity,
                voucherNo: vch,
                accountName: account,
                amount,
                date: fields.date,
                explanation: `High-value voucher of ₹${amount.toLocaleString('en-IN')} dated on Sunday (${fields.date}). Potential back-dated entry.`,
                reviewStatus: 'open',
                createdAt: new Date().toISOString(),
              });
              existingKeys.add(dedupKey);
            }
          }
        }

        // E. Polarity Inversion (Negative Cash, Debit Creditors, Credit Debtors)
        if (rule.category === 'abnormal_polarity') {
          const accLower = account.toLowerCase();
          const debit = fields.debit || 0;
          const credit = fields.credit || 0;

          // 1. Negative Cash / Cash In Hand Credit Movement
          if (accLower.includes('cash in hand') && credit > debit) {
            const dedupKey = `${rule.ruleCode}|${vch}|cash`;
            if (!existingKeys.has(dedupKey)) {
              generatedAlerts.push({
                id: `alert-${Date.now()}-${generatedAlerts.length + 1}`,
                firmId: context.firmId,
                clientId: context.clientId,
                ruleCode: rule.ruleCode,
                ruleName: rule.name,
                severity: 'critical',
                voucherNo: vch,
                accountName: account,
                amount,
                date: fields.date,
                explanation: `Critical Accounting Anomaly: Cash in Hand shows credit movement leading to negative cash balance.`,
                reviewStatus: 'open',
                createdAt: new Date().toISOString(),
              });
              existingKeys.add(dedupKey);
            }
          }

          // 2. Abnormal Polarity in Sundry Creditors (Debit balance without payment or return)
          if ((accLower.includes('creditor') || accLower.includes('payable')) && debit > 0 && credit === 0 && !narration.includes('payment') && !narration.includes('return') && !narration.includes('chq') && !narration.includes('neft')) {
            const dedupKey = `${rule.ruleCode}|${vch}|creditor_debit`;
            if (!existingKeys.has(dedupKey)) {
              generatedAlerts.push({
                id: `alert-${Date.now()}-${generatedAlerts.length + 1}`,
                firmId: context.firmId,
                clientId: context.clientId,
                ruleCode: rule.ruleCode,
                ruleName: rule.name,
                severity: 'warning',
                voucherNo: vch,
                accountName: account,
                amount,
                date: fields.date,
                explanation: `Abnormal Polarity: Sundry Creditor "${account}" debited without standard vendor payment or purchase return voucher.`,
                reviewStatus: 'open',
                createdAt: new Date().toISOString(),
              });
              existingKeys.add(dedupKey);
            }
          }

          // 3. Abnormal Polarity in Sundry Debtors (Credit balance without receipt or credit note)
          if ((accLower.includes('debtor') || accLower.includes('receivable')) && credit > 0 && debit === 0 && !narration.includes('receipt') && !narration.includes('received') && !narration.includes('credit note')) {
            const dedupKey = `${rule.ruleCode}|${vch}|debtor_credit`;
            if (!existingKeys.has(dedupKey)) {
              generatedAlerts.push({
                id: `alert-${Date.now()}-${generatedAlerts.length + 1}`,
                firmId: context.firmId,
                clientId: context.clientId,
                ruleCode: rule.ruleCode,
                ruleName: rule.name,
                severity: 'warning',
                voucherNo: vch,
                accountName: account,
                amount,
                date: fields.date,
                explanation: `Abnormal Polarity: Sundry Debtor "${account}" credited without standard customer collection receipt or credit note.`,
                reviewStatus: 'open',
                createdAt: new Date().toISOString(),
              });
              existingKeys.add(dedupKey);
            }
          }
        }
      }
    }

    const mergedAlerts = [...existingAlerts, ...generatedAlerts];
    this.store.saveAnomalyAlerts(context.clientId, mergedAlerts);

    return {
      alerts: mergedAlerts,
      newAlertsCount: generatedAlerts.length,
    };
  }

  /**
   * Reviews an anomaly alert with investigator feedback or false-positive tuning
   */
  public reviewAlert(
    context: FirmContext,
    alertId: string,
    action: 'investigating' | 'dismissed_false_positive' | 'marked_for_adjustment',
    resolutionNotes: string
  ): { success: boolean; alert?: AnomalyAlert; error?: string } {
    const alerts = this.store.listAnomalyAlerts(context.clientId);
    const target = alerts.find((a) => a.id === alertId);
    if (!target) return { success: false, error: 'Alert not found' };

    const updated = this.store.updateAnomalyAlert(context.clientId, alertId, {
      reviewStatus: action,
      resolutionNotes,
      reviewedBy: `${context.userId} (${context.userRole})`,
      reviewedAt: new Date().toISOString(),
    });

    this.store.logAuditEvent({
      id: `aud-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`,
      firmId: context.firmId,
      clientId: context.clientId,
      userId: context.userId,
      userRole: context.userRole,
      action: `ANOMALY_ALERT_${action.toUpperCase()}`,
      entityType: 'approval',
      entityId: alertId,
      details: `Anomaly [${target.ruleCode}] review updated to "${action}". Reason: ${resolutionNotes}`,
      timestamp: new Date().toISOString(),
    });

    return { success: true, alert: updated };
  }

  /**
   * Exports Anomaly Investigation Report
   */
  public exportAnomalyReport(context: FirmContext) {
    const alerts = this.store.listAnomalyAlerts(context.clientId);
    const rows = alerts.map((a) => ({
      'Alert ID': a.id,
      Severity: a.severity.toUpperCase(),
      'Rule Code': a.ruleCode,
      'Rule Name': a.ruleName,
      Voucher: a.voucherNo || 'N/A',
      Account: a.accountName || 'N/A',
      'Amount (₹)': a.amount || 0,
      Date: a.date || 'N/A',
      Explanation: a.explanation,
      'Review Status': a.reviewStatus,
      'Resolution Notes': a.resolutionNotes || 'Pending Review',
      'Reviewed By': a.reviewedBy || 'Unreviewed',
    }));

    return this.exportService.generateExcelWorkbook(
      context,
      `anomaly-${context.clientId}`,
      `Anomaly_Investigation_Report_${context.clientId}.xlsx`,
      [{ sheetName: 'Anomaly_Alerts', rows }]
    );
  }
}
