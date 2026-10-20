/**
 * Configurable Accounting Rules Engine
 * Deterministic Financial & Statutory Validation
 */

import type { ExtractedRecord, ValidationIssue, AccountingRuleConfig } from '../types';

export class AccountingRulesEngine {
  /**
   * Validates bank statement running balance continuity
   */
  public static validateRunningBalanceContinuity(
    records: ExtractedRecord[],
    sourceFileId?: string
  ): { isValid: boolean; issues: ValidationIssue[] } {
    const issues: ValidationIssue[] = [];

    for (let i = 1; i < records.length; i++) {
      const prev = records[i - 1];
      const curr = records[i];

      const prevBal = prev.normalizedFields.balance;
      const currBal = curr.normalizedFields.balance;
      const wdr = curr.normalizedFields.debit || 0;
      const dep = curr.normalizedFields.credit || 0;

      if (prevBal !== undefined && currBal !== undefined && (wdr > 0 || dep > 0)) {
        const expectedBal = prevBal - wdr + dep;
        const diff = Math.abs(expectedBal - currBal);

        if (diff > 0.05) {
          issues.push({
            id: `issue-cont-${i}`,
            ruleCode: 'RULE_BANK_CONTINUITY_BREAK',
            ruleName: 'Running Balance Discontinuity',
            severity: 'critical',
            message: `Discontinuity at Row ${curr.rowIndex}: Expected balance ₹${expectedBal.toFixed(2)} (Prev ₹${prevBal.toFixed(2)} - Wdr ₹${wdr.toFixed(2)} + Dep ₹${dep.toFixed(2)}) but statement shows ₹${currBal.toFixed(2)} (Break of ₹${diff.toFixed(2)})`,
            field: 'balance',
            actualValue: currBal,
            expectedValue: expectedBal,
            recordId: curr.id,
            sourceFileId,
            suggestedCorrection: { balance: expectedBal },
          });
        }
      }
    }

    return {
      isValid: issues.length === 0,
      issues,
    };
  }

  /**
   * Validates statutory tax arithmetic for an invoice
   */
  public static validateInvoiceArithmetic(
    record: ExtractedRecord,
    tolerance: number = 1.0
  ): ValidationIssue[] {
    const issues: ValidationIssue[] = [];
    const fields = record.normalizedFields;

    const taxable = fields.taxableAmount || 0;
    const cgst = fields.cgst || 0;
    const sgst = fields.sgst || 0;
    const igst = fields.igst || 0;
    const cess = fields.cess || 0;
    const total = fields.totalAmount || 0;

    const calculatedTotal = taxable + cgst + sgst + igst + cess;
    const diff = Math.abs(calculatedTotal - total);

    if (total > 0 && taxable > 0 && diff > tolerance) {
      issues.push({
        id: `issue-tax-math-${record.id}`,
        ruleCode: 'RULE_INVOICE_TAX_ARITHMETIC',
        ruleName: 'Statutory Tax Math Discrepancy',
        severity: 'error',
        message: `Invoice ${fields.invoiceNo || 'Unknown'}: Sum of components (Taxable ₹${taxable.toFixed(2)} + CGST ₹${cgst.toFixed(2)} + SGST ₹${sgst.toFixed(2)} + IGST ₹${igst.toFixed(2)} + Cess ₹${cess.toFixed(2)} = ₹${calculatedTotal.toFixed(2)}) deviates from stated Total ₹${total.toFixed(2)} by ₹${diff.toFixed(2)}`,
        field: 'totalAmount',
        actualValue: total,
        expectedValue: calculatedTotal,
        recordId: record.id,
        suggestedCorrection: { totalAmount: calculatedTotal },
      });
    }

    // Intra-state symmetry check: CGST should equal SGST
    if (cgst > 0 && sgst > 0 && Math.abs(cgst - sgst) > 1.0) {
      issues.push({
        id: `issue-intra-gst-${record.id}`,
        ruleCode: 'RULE_CGST_SGST_ASYMMETRY',
        ruleName: 'Intra-State CGST/SGST Asymmetry',
        severity: 'warning',
        message: `Intra-state GST asymmetry: CGST (₹${cgst.toFixed(2)}) does not equal SGST (₹${sgst.toFixed(2)})`,
        field: 'cgst',
        actualValue: cgst,
        expectedValue: sgst,
        recordId: record.id,
      });
    }

    return issues;
  }

  /**
   * Validates Section 40A(3) Income Tax Act Cash Limit (₹10,000 threshold)
   */
  public static validateCashPaymentLimit(
    record: ExtractedRecord,
    maxLimit: number = 10000
  ): ValidationIssue | undefined {
    const narration = (record.normalizedFields.narration || '').toLowerCase();
    const amount = record.normalizedFields.amount || record.normalizedFields.debit || 0;

    // Section 40A(3) applies strictly to expenditures/payments made, NOT cash deposits into bank
    const isDeposit = narration.includes('deposit') || narration.includes('dep ') || ((record.normalizedFields.credit || 0) > 0 && (record.normalizedFields.debit || 0) === 0);
    if (isDeposit) return undefined;

    const hasElectronicIndicator = narration.includes('neft') || narration.includes('rtgs') || narration.includes('upi') || narration.includes('imps') || narration.includes('chq') || narration.includes('cheque');
    const isExplicitCash = narration.includes('cash paid') || narration.includes('cash purchase') || narration.includes('cash payment') || record.normalizedFields.refNo === 'CASH' || narration.startsWith('self');
    const isCashTxn = isExplicitCash || (narration.includes('cash') && !hasElectronicIndicator);

    if (isCashTxn && amount > maxLimit) {
      return {
        id: `issue-40a3-${record.id}`,
        ruleCode: 'RULE_SECTION_40A_3_VIOLATION',
        ruleName: 'Income Tax Sec 40A(3) Cash Disallowance Risk',
        severity: 'error',
        message: `Cash payment of ₹${amount.toLocaleString('en-IN')} exceeds statutory limit of ₹${maxLimit.toLocaleString('en-IN')} under Section 40A(3). Risk of 100% expenditure disallowance.`,
        field: 'amount',
        actualValue: amount,
        expectedValue: maxLimit,
        recordId: record.id,
      };
    }
    return undefined;
  }

  /**
   * Validates Trial Balance Double-Entry Balance
   */
  public static validateTrialBalance(records: ExtractedRecord[]): {
    isBalanced: boolean;
    totalDebits: number;
    totalCredits: number;
    difference: number;
    issues: ValidationIssue[];
  } {
    const issues: ValidationIssue[] = [];
    let totalDebits = 0;
    let totalCredits = 0;

    for (const rec of records) {
      const dr = rec.normalizedFields.debit || 0;
      const cr = rec.normalizedFields.credit || 0;
      totalDebits += dr;
      totalCredits += cr;

      // Check abnormal balance
      const name = (rec.normalizedFields.accountName || '').toLowerCase();
      if ((name.includes('cash in hand') || name.includes('petty cash')) && cr > dr) {
        issues.push({
          id: `issue-abn-cash-${rec.id}`,
          ruleCode: 'RULE_ABNORMAL_NEGATIVE_CASH',
          ruleName: 'Abnormal Credit Balance in Cash',
          severity: 'critical',
          message: `Account "${rec.normalizedFields.accountName}" has credit balance of ₹${(cr - dr).toFixed(2)}. Cash balance cannot be negative.`,
          field: 'balance',
          recordId: rec.id,
        });
      }
    }

    const diff = Math.abs(totalDebits - totalCredits);
    const isBalanced = diff < 0.05;

    if (!isBalanced) {
      issues.push({
        id: `issue-tb-unbalanced`,
        ruleCode: 'RULE_TRIAL_BALANCE_UNBALANCED',
        ruleName: 'Debit and Credit Imbalance',
        severity: 'critical',
        message: `Trial Balance does not balance: Total Debits ₹${totalDebits.toLocaleString('en-IN')} vs Total Credits ₹${totalCredits.toLocaleString('en-IN')} (Variance ₹${diff.toLocaleString('en-IN')})`,
        actualValue: totalDebits,
        expectedValue: totalCredits,
      });
    }

    return {
      isBalanced,
      totalDebits,
      totalCredits,
      difference: diff,
      issues,
    };
  }

  /**
   * Master validator running all active rules on a dataset
   */
  public static validateDataset(
    records: ExtractedRecord[],
    config: AccountingRuleConfig,
    docType: string
  ): ValidationIssue[] {
    const issues: ValidationIssue[] = [];

    if (docType === 'bank_statement') {
      const cont = this.validateRunningBalanceContinuity(records);
      issues.push(...cont.issues);
    }

    if (docType === 'purchase_invoice' || docType === 'purchase_register') {
      for (const rec of records) {
        const arith = this.validateInvoiceArithmetic(rec, config.gstTolerance);
        issues.push(...arith);
      }
    }

    if (docType === 'trial_balance') {
      const tb = this.validateTrialBalance(records);
      issues.push(...tb.issues);
    }

    // Generic cash limits & missing date checks
    for (const rec of records) {
      const cashIssue = this.validateCashPaymentLimit(rec, config.maxCashVoucherAmount);
      if (cashIssue) issues.push(cashIssue);

      if (!rec.normalizedFields.date && (rec.normalizedFields.amount || 0) > 0) {
        issues.push({
          id: `issue-missing-date-${rec.id}`,
          ruleCode: 'RULE_MISSING_TRANSACTION_DATE',
          ruleName: 'Missing Transaction Date',
          severity: 'warning',
          message: `Row ${rec.rowIndex}: Transaction has value ₹${(rec.normalizedFields.amount || 0).toFixed(2)} but date is missing or unparseable`,
          field: 'date',
          recordId: rec.id,
        });
      }
    }

    return issues;
  }
}
