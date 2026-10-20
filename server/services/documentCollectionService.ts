/**
 * Module 5: Client Document Collection & Follow-up Automation Service
 * Tracks client submission checklists, receives files, and drafts statutory reminder notices
 */

import type { ClientDocumentRequest, DocumentReminderDraft, FirmContext } from '../types';
import { BackendStore } from '../db/store';
import { ExportService } from './exportService';

export class DocumentCollectionService {
  private store: BackendStore;
  private exportService: ExportService;

  constructor(store?: BackendStore) {
    this.store = store || BackendStore.getInstance();
    this.exportService = new ExportService(this.store);
  }

  /**
   * Initializes standard engagement document requests for a client
   */
  public createEngagementBundle(
    context: FirmContext,
    financialPeriod: string = 'FY 2026-27',
    engagementType: ClientDocumentRequest['engagementType'] = 'Statutory Audit',
    dueDate: string = '15-10-2026'
  ): ClientDocumentRequest[] {
    const standardItems: Omit<ClientDocumentRequest, 'id' | 'createdAt' | 'updatedAt'>[] = [
      {
        firmId: context.firmId,
        clientId: context.clientId,
        financialPeriod,
        engagementType,
        category: 'Banking & Cash',
        title: 'Full Financial Year Bank Statements with Certificate',
        description: 'Signed PDF bank statement downloads covering 01-04-2026 to 31-03-2027 with year-end interest certificates',
        requiredDueDate: dueDate,
        status: 'requested',
        reminderCount: 0,
        receivedFiles: [],
      },
      {
        firmId: context.firmId,
        clientId: context.clientId,
        financialPeriod,
        engagementType,
        category: 'Direct Taxes',
        title: 'Form 26AS, Annual Information Statement (AIS) & TIS',
        description: 'Complete TRACES 26AS and ITD portal AIS/TIS annual JSON/PDF downloads for tax credit verification',
        requiredDueDate: dueDate,
        status: 'requested',
        reminderCount: 0,
        receivedFiles: [],
      },
      {
        firmId: context.firmId,
        clientId: context.clientId,
        financialPeriod,
        engagementType,
        category: 'GST & Indirect Taxes',
        title: 'Annual GSTR-1, GSTR-3B & GSTR-2B Summary Worksheets',
        description: 'Monthly filed GSTR-3B acknowledgments and GSTR-2B inward ITC portal registers',
        requiredDueDate: dueDate,
        status: 'requested',
        reminderCount: 0,
        receivedFiles: [],
      },
      {
        firmId: context.firmId,
        clientId: context.clientId,
        financialPeriod,
        engagementType,
        category: 'Fixed Assets',
        title: 'Fixed Asset Register (FAR) & Purchase Invoices',
        description: 'Detailed FAR indicating asset descriptions, gross block, additions, deletions, and put-to-use dates',
        requiredDueDate: dueDate,
        status: 'requested',
        reminderCount: 0,
        receivedFiles: [],
      },
      {
        firmId: context.firmId,
        clientId: context.clientId,
        financialPeriod,
        engagementType,
        category: 'Inventory & Stock',
        title: 'Physical Stock Inventory Valuation Certificate',
        description: 'Management representation certificate for physical stock count, valuation methodology (AS-2 / Ind AS 2)',
        requiredDueDate: dueDate,
        status: 'requested',
        reminderCount: 0,
        receivedFiles: [],
      },
    ];

    const createdRequests: ClientDocumentRequest[] = [];
    for (const item of standardItems) {
      const created = this.store.createDocumentRequest({
        ...item,
        id: `dreq-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      });
      createdRequests.push(created);
    }

    return createdRequests;
  }

  /**
   * Registers a file uploaded by the client against a specific request
   */
  public registerClientFile(
    context: FirmContext,
    requestId: string,
    fileName: string,
    fileSizeBytes: number,
    fileHash: string
  ): ClientDocumentRequest {
    const list = this.store.listDocumentRequests(context.clientId);
    const req = list.find((r) => r.id === requestId);
    if (!req) throw new Error('Document request not found');

    req.receivedFiles.push({
      fileName,
      fileSizeBytes,
      uploadedAt: new Date().toISOString(),
      fileHash,
      accepted: true,
    });
    req.status = 'received';

    this.store.updateDocumentRequest(context.clientId, requestId, {
      receivedFiles: req.receivedFiles,
      status: 'received',
    });

    this.store.logAuditEvent({
      id: `aud-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`,
      firmId: context.firmId,
      clientId: context.clientId,
      userId: context.userId,
      userRole: context.userRole,
      action: 'CLIENT_DOCUMENT_RECEIVED',
      entityType: 'file',
      entityId: requestId,
      details: `Received file "${fileName}" (${(fileSizeBytes / 1024).toFixed(1)} KB) for request "${req.title}"`,
      timestamp: new Date().toISOString(),
    });

    return req;
  }

  /**
   * Drafts an executive reminder letter/email for outstanding documents
   */
  public draftReminderNotice(
    context: FirmContext,
    clientName: string = 'Valued Client',
    contactEmail: string = 'finance@client.com'
  ): DocumentReminderDraft {
    const requests = this.store.listDocumentRequests(context.clientId);
    const pending = requests.filter((r) => r.status === 'requested' || r.status === 'partially_received');

    const pendingBullets = pending
      .map((p, i) => `${i + 1}. [${p.category}] ${p.title} (Due: ${p.requiredDueDate})`)
      .join('\n');

    const subject = `Urgent: Pending Audit & Compliance Documentation Request - ${context.clientId.toUpperCase()}`;
    const body = `Dear ${clientName} Accounts Team,

We hope this message finds you well.

In connection with our ongoing ${pending[0]?.engagementType || 'Audit & Accounting Engagement'} for the period ${pending[0]?.financialPeriod || 'FY 2026-27'}, we have completed our preliminary audit scrutiny.

To maintain statutory compliance timelines and finalize our review without statutory filing delays, please provide the following ${pending.length} pending document(s) at your earliest convenience:

${pendingBullets}

Please upload these documents securely through the CA Tools Client Portal or reply directly with the requested attachments.

Kind regards,
Audit & Assurance Team
${context.firmName}
`;

    return {
      clientId: context.clientId,
      clientName,
      contactEmail,
      subject,
      body,
      pendingItemsCount: pending.length,
      earliestDueDate: pending[0]?.requiredDueDate || 'Immediate',
    };
  }

  /**
   * Exports Client-wise Outstanding Document Checklist
   */
  public exportOutstandingDocumentReport(context: FirmContext) {
    const list = this.store.listDocumentRequests(context.clientId);
    const rows = list.map((r) => ({
      Category: r.category,
      'Document Title': r.title,
      'Engagement Type': r.engagementType,
      'Due Date': r.requiredDueDate,
      Status: r.status.toUpperCase(),
      'Files Received': r.receivedFiles.length,
      'Reminders Dispatched': r.reminderCount,
      'Last Follow-up': r.lastReminderSentAt || 'None',
    }));

    return this.exportService.generateExcelWorkbook(
      context,
      `doc-req-${context.clientId}`,
      `Outstanding_Document_Request_${context.clientId}.xlsx`,
      [{ sheetName: 'Document_Checklist', rows }]
    );
  }
}
