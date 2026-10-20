/**
 * Module 3: Audit Evidence Management & Traceability Service
 * Connects transaction vouchers to working paper evidence indexes and flags missing documentation
 */

import type { AuditEvidence, EvidenceLink, FirmContext, ExtractedRecord } from '../types';
import { BackendStore } from '../db/store';
import { ExportService } from './exportService';

export class AuditEvidenceService {
  private store: BackendStore;
  private exportService: ExportService;

  constructor(store?: BackendStore) {
    this.store = store || BackendStore.getInstance();
    this.exportService = new ExportService(this.store);
  }

  /**
   * Initializes working paper evidence slots for an engagement period
   */
  public initializeEvidenceSlots(
    context: FirmContext,
    period: string = 'September 2026'
  ): AuditEvidence[] {
    const defaultEvidenceSlots: Omit<AuditEvidence, 'id'>[] = [
      {
        firmId: context.firmId,
        clientId: context.clientId,
        period,
        auditArea: 'Cash & Bank',
        wpIndex: 'WP-BNK-001',
        documentTitle: 'Bank Statement & BRS Reconciliation Certificate',
        status: 'Requested',
        sourceReference: 'Main Current Account',
      },
      {
        firmId: context.firmId,
        clientId: context.clientId,
        period,
        auditArea: 'Purchases',
        wpIndex: 'WP-PUR-001',
        documentTitle: 'Purchase Register & GSTR-2B Matching Schedule',
        status: 'Requested',
        sourceReference: 'Inward Invoices',
      },
      {
        firmId: context.firmId,
        clientId: context.clientId,
        period,
        auditArea: 'Revenue',
        wpIndex: 'WP-SAL-001',
        documentTitle: 'Sales Register, GSTR-1, and E-Way Bill Reconciliation',
        status: 'Requested',
        sourceReference: 'Outward Turnover',
      },
      {
        firmId: context.firmId,
        clientId: context.clientId,
        period,
        auditArea: 'Taxes',
        wpIndex: 'WP-TAX-001',
        documentTitle: 'Form 26AS, AIS, and TIS Annual Tax Statements',
        status: 'Requested',
        sourceReference: 'ITD Portal Download',
      },
      {
        firmId: context.firmId,
        clientId: context.clientId,
        period,
        auditArea: 'Fixed Assets',
        wpIndex: 'WP-FAR-001',
        documentTitle: 'Fixed Asset Register (FAR) & Depreciation Schedule II',
        status: 'Requested',
        sourceReference: 'Property, Plant & Equipment',
      },
    ];

    const evidenceList: AuditEvidence[] = defaultEvidenceSlots.map((slot, idx) => ({
      ...slot,
      id: `ev-${Date.now()}-${idx + 1}`,
    }));

    this.store.saveAuditEvidence(context.clientId, evidenceList);
    return evidenceList;
  }

  /**
   * Links a ledger transaction to an evidence working paper slot
   */
  public linkTransactionToEvidence(
    context: FirmContext,
    evidenceId: string,
    transaction: {
      transactionRef: string;
      voucherNo: string;
      amount: number;
      date: string;
      accountHead: string;
      verificationStatus?: 'Satisfactory' | 'Exceptions Found' | 'Pending Inquiry';
    }
  ): EvidenceLink {
    const link: EvidenceLink = {
      id: `link-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`,
      evidenceId,
      transactionRef: transaction.transactionRef,
      voucherNo: transaction.voucherNo,
      amount: transaction.amount,
      date: transaction.date,
      accountHead: transaction.accountHead,
      verificationStatus: transaction.verificationStatus || 'Satisfactory',
    };

    this.store.saveEvidenceLink(link);

    // Update parent evidence status to 'Linked'
    const list = this.store.getAuditEvidence(context.clientId);
    const target = list.find((e) => e.id === evidenceId);
    if (target && target.status === 'Requested') {
      target.status = 'Linked';
      this.store.saveAuditEvidence(context.clientId, list);
    }

    // Audit Event
    this.store.logAuditEvent({
      id: `aud-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`,
      firmId: context.firmId,
      clientId: context.clientId,
      userId: context.userId,
      userRole: context.userRole,
      action: 'EVIDENCE_LINKED',
      entityType: 'approval',
      entityId: link.id,
      details: `Linked Voucher ${link.voucherNo} (₹${link.amount.toLocaleString('en-IN')}) to Evidence ${target?.wpIndex || evidenceId}`,
      timestamp: new Date().toISOString(),
    });

    return link;
  }

  /**
   * Scans transactions and identifies items exceeding materiality threshold lacking supporting evidence
   */
  public detectMissingEvidence(
    context: FirmContext,
    transactions: ExtractedRecord[],
    materialityThreshold: number = 100000
  ): {
    missingCount: number;
    flaggedTransactions: {
      voucherNo: string;
      date: string;
      amount: number;
      account: string;
      narration: string;
      reason: string;
    }[];
  } {
    const existingList = this.store.getAuditEvidence(context.clientId);
    const linkedLinks: EvidenceLink[] = [];
    for (const ev of existingList) {
      linkedLinks.push(...this.store.getEvidenceLinks(ev.id));
    }
    const linkedVouchers = new Set(linkedLinks.map((l) => l.voucherNo));

    const flagged: any[] = [];
    for (const txn of transactions) {
      const amt = txn.normalizedFields.amount || Math.max(txn.normalizedFields.debit || 0, txn.normalizedFields.credit || 0);
      const vch = txn.normalizedFields.invoiceNo || txn.normalizedFields.refNo || `VCH-${txn.rowIndex}`;

      if (amt >= materialityThreshold && !linkedVouchers.has(vch)) {
        flagged.push({
          voucherNo: vch,
          date: txn.normalizedFields.date || 'Undated',
          amount: amt,
          account: txn.normalizedFields.accountName || txn.normalizedFields.partyName || 'General Account',
          narration: txn.normalizedFields.narration || '',
          reason: `Material transaction of ₹${amt.toLocaleString('en-IN')} exceeds materiality threshold of ₹${materialityThreshold.toLocaleString('en-IN')} and lacks linked working paper evidence`,
        });
      }
    }

    return {
      missingCount: flagged.length,
      flaggedTransactions: flagged,
    };
  }

  /**
   * Exports Audit Evidence Register & Index
   */
  public exportEvidenceRegister(context: FirmContext, period: string) {
    const list = this.store.getAuditEvidence(context.clientId);
    const rows = list.map((e) => {
      const links = this.store.getEvidenceLinks(e.id);
      return {
        'WP Index': e.wpIndex,
        'Audit Area': e.auditArea,
        'Document Title': e.documentTitle,
        Status: e.status,
        'Linked Transactions': links.length,
        'Total Verified Amount': links.reduce((acc, l) => acc + l.amount, 0),
        'Reviewed By': e.reviewedBy || 'Pending',
        'Reviewer Notes': e.reviewerNotes || '',
      };
    });

    return this.exportService.generateExcelWorkbook(
      context,
      `evidence-${context.clientId}`,
      `Audit_Evidence_Register_${context.clientId}_${period.replace(/\s+/g, '_')}.xlsx`,
      [{ sheetName: 'Evidence_Index', rows }]
    );
  }
}
