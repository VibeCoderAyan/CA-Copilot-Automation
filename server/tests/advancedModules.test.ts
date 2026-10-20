/**
 * Automated Test Suite for 7 Advanced CA Firm Automation Modules
 * Verifies real end-to-end services, statutory rules, prerequisite graphs, and SheetJS exports
 */

import path from 'path';
import fs from 'fs';
import { BackendStore } from '../db/store.ts';
import { MonthEndClosingService } from '../services/monthEndClosingService.ts';
import { ApArAutomationService } from '../services/apArService.ts';
import { AuditEvidenceService } from '../services/auditEvidenceService.ts';
import { AnomalyDetectionService } from '../services/anomalyDetectionService.ts';
import { DocumentCollectionService } from '../services/documentCollectionService.ts';
import { RecurringComplianceService } from '../services/recurringComplianceService.ts';
import { MultiClientBatchService } from '../services/multiClientBatchService.ts';
import type { FirmContext, ExtractedRecord } from '../types/index.ts';

function assert(condition: boolean, message: string): void {
  if (!condition) {
    console.error(`❌ FAIL: ${message}`);
    process.exit(1);
  } else {
    console.log(`✓ PASS: ${message}`);
  }
}

async function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function runAdvancedModulesTests() {
  console.log('================================================================');
  console.log('🧪 RUNNING ADVANCED CA FIRM AUTOMATION SUITE TEST SUITE (7 MODULES)');
  console.log('================================================================\n');

  // Isolated test store
  const testStorePath = path.resolve(process.cwd(), 'data', 'test-advanced-modules-store.json');
  if (fs.existsSync(testStorePath)) fs.unlinkSync(testStorePath);
  const store = BackendStore.getInstance(testStorePath);
  store.clearAll();

  const firmA: FirmContext = {
    firmId: 'firm_apex_ca',
    firmName: 'Apex Chartered Accountants & Co.',
    userId: 'ca_partner_01',
    userRole: 'Partner',
    clientId: 'client_apex_technologies',
  };

  const firmB: FirmContext = {
    firmId: 'firm_apex_ca',
    firmName: 'Apex Chartered Accountants & Co.',
    userId: 'ca_article_02',
    userRole: 'Article',
    clientId: 'client_bharat_logistics',
  };

  // Instantiate services
  const closingService = new MonthEndClosingService(store);
  const apArService = new ApArAutomationService(store);
  const evidenceService = new AuditEvidenceService(store);
  const anomalyService = new AnomalyDetectionService(store);
  const docService = new DocumentCollectionService(store);
  const complianceService = new RecurringComplianceService(store);
  const batchService = new MultiClientBatchService(store);

  // =========================================================================
  // MODULE 1: MONTH-END CLOSING AUTOMATION
  // =========================================================================
  console.log('--- MODULE 1: MONTH-END CLOSING WORKFLOW ---');
  const createResult = closingService.createPeriod(firmA, 'September 2026', 'FY 2026-27', '10-10-2026');
  const period = createResult.period;
  assert(period.status === 'in_progress', 'Closing period created in in_progress state');
  assert(createResult.tasks.length === 5, 'Checklist initialized with 5 prerequisite steps');

  // Evaluate initial status
  const initialEval = closingService.evaluateClosingStatus(firmA, period.id);
  assert(initialEval.metrics.canClose === false, 'Cannot close period initially with pending checklist');

  // Mark all checklist steps completed
  for (const item of createResult.tasks) {
    store.updateClosingTask(period.id, item.id, {
      status: 'completed',
      completionNotes: `Verified by ${firmA.userId}`,
      completedBy: firmA.userId,
      completedAt: new Date().toISOString(),
    });
  }

  // Attempt partner approval
  const approveResult = closingService.approveClosing(firmA, period.id, 'Partner sign-off verified all checks');
  assert(approveResult.success === true, 'Partner approval succeeds when prerequisites are satisfied');
  assert(approveResult.period?.status === 'closed', 'Period status is locked as closed');
  assert(approveResult.period?.approvedBy?.includes(firmA.userId) === true, 'Audit log records partner userId');

  // Export closing summary
  const closingExport = closingService.exportClosingSummary(firmA, period.id);
  assert(Boolean(closingExport.id), 'Closing export artifact generated');
  assert(closingExport.fileType === 'xlsx', 'Closing export is valid xlsx format');
  assert(Boolean(closingExport.contentBase64), 'Closing export contains SheetJS Base64 payload');

  // =========================================================================
  // MODULE 2: AP & AR AUTOMATION
  // =========================================================================
  console.log('\n--- MODULE 2: ACCOUNTS PAYABLE & RECEIVABLE AUTOMATION ---');
  const sampleInvoices = [
    // 0-30 days
    { 'Party Name': 'Tata Steel', 'Invoice No': 'INV-101', 'Invoice Date': '2026-09-20', 'Total': 100000, 'Paid': 0, 'Terms': 15 },
    // 61-90 days, overdue > 45 days
    { 'Party Name': 'Shree Cement', 'Invoice No': 'INV-102', 'Invoice Date': '2026-07-15', 'Total': 250000, 'Paid': 0, 'Terms': 30 },
    // Duplicate of INV-101 to trigger duplicate risk
    { 'Party Name': 'Tata Steel', 'Invoice No': 'INV-101', 'Invoice Date': '2026-09-20', 'Total': 100000, 'Paid': 0, 'Terms': 15 },
  ];

  const aparResult = apArService.processLedgerRecords(firmA, sampleInvoices, 'payable', '2026-10-15');
  assert(aparResult.agingRecords.length === 3, 'Processed 3 open ledger records');
  assert(aparResult.summary.totalOutstanding === 450000, 'Calculated correct total outstanding liability of 4,50,000');

  // Verify MSMED 45-day alert
  const msmeAlert = aparResult.riskAlerts.find((a) => a.riskType === 'msme_45_day_violation');
  assert(Boolean(msmeAlert), 'Flagged Section 15 MSMED 45-day overdue liability violation');

  // Verify duplicate invoice/payment alert
  const dupAlert = aparResult.riskAlerts.find((a) => a.riskType === 'duplicate_invoice_number');
  assert(Boolean(dupAlert), 'Flagged potential duplicate invoice payment risk');

  // Generate payment review list
  const paymentReviewExport = apArService.generatePaymentReviewList(firmA);
  assert(paymentReviewExport.fileType === 'xlsx', 'Generated Excel Payment Review List without automated payments');

  // =========================================================================
  // MODULE 3: AUDIT EVIDENCE AUTOMATION
  // =========================================================================
  console.log('\n--- MODULE 3: AUDIT EVIDENCE & TRACEABILITY ---');
  const evidenceSlots = evidenceService.initializeEvidenceSlots(firmA, 'September 2026');
  assert(evidenceSlots.length >= 5, 'Initialized standard working paper slots (WP-BNK, WP-PUR, WP-REV, etc.)');

  // Link transaction to WP-BNK-001
  const bnkSlot = evidenceSlots.find((s) => s.wpIndex === 'WP-BNK-001')!;
  const linkResult = evidenceService.linkTransactionToEvidence(firmA, bnkSlot.id, {
    transactionRef: 'TXN-BNK-001',
    voucherNo: 'VR-BNK-001',
    date: '2026-09-15',
    amount: 1500000,
    accountHead: 'Bank Balances',
  });
  assert(linkResult.voucherNo === 'VR-BNK-001', 'Transaction linked to working paper slot');

  // Check parent slot updated to Linked
  const updatedSlots = store.getAuditEvidence(firmA.clientId);
  const updatedBnk = updatedSlots.find((s) => s.id === bnkSlot.id)!;
  assert(updatedBnk.status === 'Linked', 'Slot status updated to Linked');

  // Materiality scan for missing evidence
  const highValueTxns: ExtractedRecord[] = [
    {
      id: 'rec-1',
      sourceFileId: 'src-1',
      rowIndex: 1,
      rawFields: {},
      normalizedFields: { invoiceNo: 'VR-CAP-001', date: '2026-09-10', amount: 500000, accountName: 'Capital Assets' },
      status: 'valid',
    },
    {
      id: 'rec-2',
      sourceFileId: 'src-1',
      rowIndex: 2,
      rawFields: {},
      normalizedFields: { invoiceNo: 'VR-OFF-002', date: '2026-09-12', amount: 12000, accountName: 'Office Expense' },
      status: 'valid',
    },
  ];
  const missingEvidenceScan = evidenceService.detectMissingEvidence(firmA, highValueTxns, 100000);
  assert(missingEvidenceScan.missingCount === 1, 'Identified 1 material transaction lacking supporting evidence');
  assert(missingEvidenceScan.flaggedTransactions[0].voucherNo === 'VR-CAP-001', 'Correctly flagged 5,00,000 voucher');

  // Export evidence register
  const evidenceExport = evidenceService.exportEvidenceRegister(firmA, 'September 2026');
  assert(Boolean(evidenceExport.id) && evidenceExport.fileType === 'xlsx', 'Exported Audit Evidence Register');

  // =========================================================================
  // MODULE 4: ACCOUNTING ANOMALY DETECTION ENGINE
  // =========================================================================
  console.log('\n--- MODULE 4: ACCOUNTING ANOMALY DETECTION ENGINE ---');
  const sampleVouchers: ExtractedRecord[] = [
    // Duplicate journal
    {
      id: 'rec-j1',
      sourceFileId: 'src-j1',
      rowIndex: 1,
      rawFields: {},
      normalizedFields: { invoiceNo: 'JV-201', refNo: 'JV-201', date: '15-09-2026', narration: 'September retainership', debit: 45000, credit: 0, accountName: 'Legal Retainer' },
      status: 'valid',
    },
    {
      id: 'rec-j2',
      sourceFileId: 'src-j1',
      rowIndex: 2,
      rawFields: {},
      normalizedFields: { invoiceNo: 'JV-202', refNo: 'JV-202', date: '15-09-2026', narration: 'September retainership', debit: 45000, credit: 0, accountName: 'Legal Retainer' },
      status: 'valid',
    },
    // Suspicious round sum expense >= 50,000
    {
      id: 'rec-j3',
      sourceFileId: 'src-j1',
      rowIndex: 3,
      rawFields: {},
      normalizedFields: { invoiceNo: 'VR-501', refNo: 'VR-501', date: '16-09-2026', narration: 'Consulting fees', debit: 100000, credit: 0, amount: 100000, accountName: 'General Consultancy' },
      status: 'valid',
    },
    // Section 40A(3) cash violation > 10,000
    {
      id: 'rec-j4',
      sourceFileId: 'src-j1',
      rowIndex: 4,
      rawFields: {},
      normalizedFields: { invoiceNo: 'VR-CASH-01', refNo: 'VR-CASH-01', date: '17-09-2026', narration: 'Cash paid to contractor exceeding 40A(3)', debit: 25000, credit: 0, amount: 25000, accountName: 'Cash in Hand' },
      status: 'valid',
    },
    // Weekend posting > 1,00,000 (2026-09-20 was Sunday)
    {
      id: 'rec-j5',
      sourceFileId: 'src-j1',
      rowIndex: 5,
      rawFields: {},
      normalizedFields: { invoiceNo: 'BNK-SUN-01', refNo: 'BNK-SUN-01', date: '2026-09-20', narration: 'Sunday posting release', debit: 500000, credit: 0, amount: 500000, accountName: 'Director Remuneration' },
      status: 'valid',
    },
    // Polarity inversion: Negative cash balance
    {
      id: 'rec-j6',
      sourceFileId: 'src-j1',
      rowIndex: 6,
      rawFields: {},
      normalizedFields: { invoiceNo: 'CB-NEG-01', refNo: 'CB-NEG-01', date: '25-09-2026', narration: 'Excess withdrawal', debit: 0, credit: 15000, balance: -15000, accountName: 'Cash in Hand' },
      status: 'valid',
    },
  ];

  const scanResult = anomalyService.scanTransactions(firmA, sampleVouchers);
  assert(scanResult.newAlertsCount >= 4, 'Detected deterministic anomaly alerts');

  // Review an alert as dismissed false-positive with reasons
  const alerts = store.listAnomalyAlerts(firmA.clientId);
  const targetAlert = alerts[0];
  const reviewResult = anomalyService.reviewAlert(firmA, targetAlert.id, 'dismissed_false_positive', 'Verified authorized advance as per contract clause 4');
  assert(reviewResult.success === true, 'Auditor review recorded successfully');
  assert(reviewResult.alert?.reviewStatus === 'dismissed_false_positive', 'Alert status updated to dismissed_false_positive');
  assert(Boolean(reviewResult.alert?.resolutionNotes?.includes('clause 4')), 'Audit notes preserved');

  // Export anomaly report
  const anomalyExport = anomalyService.exportAnomalyReport(firmA);
  assert(anomalyExport.fileType === 'xlsx', 'Exported deterministic anomaly scrutiny register');

  // =========================================================================
  // MODULE 5: CLIENT DOCUMENT COLLECTION & FOLLOW-UP
  // =========================================================================
  console.log('\n--- MODULE 5: CLIENT DOCUMENT COLLECTION & FOLLOW-UP ---');
  const bundle = docService.createEngagementBundle(firmA, 'FY 2026-27', 'Statutory Audit', '15-10-2026');
  assert(bundle.length === 5, 'Generated 5 statutory requisition items');

  // Register client file with SHA-256 hash
  const bankReq = bundle.find((r) => r.category === 'Banking & Cash')!;
  const registeredDoc = docService.registerClientFile(firmA, bankReq.id, 'hdfc_statement_sep2026.pdf', 345000, 'sha256_e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855');
  assert(registeredDoc.status === 'received', 'Document status updated to received upon verified file upload');
  assert(registeredDoc.receivedFiles[0].fileHash.startsWith('sha256_'), 'File integrity hash stamped');

  // Draft polite reminder notice for remaining items
  const reminderDraft = docService.draftReminderNotice(firmA, 'Apex Technologies Pvt Ltd', 'cfo@apextech.in');
  assert(reminderDraft.pendingItemsCount === 4, 'Reminder includes 4 remaining pending requisition items');
  assert(reminderDraft.body.includes('Apex Technologies Pvt Ltd'), 'Personalized reminder drafted with client name');

  // Export outstanding documents
  const docExport = docService.exportOutstandingDocumentReport(firmA);
  assert(docExport.fileType === 'xlsx', 'Exported Client Outstanding Requisition Letter');

  // =========================================================================
  // MODULE 6: RECURRING COMPLIANCE & TASK AUTOMATION
  // =========================================================================
  console.log('\n--- MODULE 6: RECURRING COMPLIANCE & TASK AUTOMATION ---');
  const complianceResult = complianceService.generatePeriodInstances(firmA, 'September 2026', 2026, 9);
  assert(complianceResult.createdCount >= 4, 'Generated statutory tasks: GSTR-1, GSTR-3B, TDS, Advance Tax');

  // Re-run generation (Idempotency test)
  const rerunResult = complianceService.generatePeriodInstances(firmA, 'September 2026', 2026, 9);
  assert(rerunResult.createdCount === 0, 'Schedule rerun is idempotent (0 duplicate tasks created)');

  // Evaluate dynamic status
  const evaluatedTasks = complianceService.evaluateTaskStatuses(firmA, '2026-10-12');
  const gstr1 = evaluatedTasks.find((t) => t.title.includes('GSTR-1'))!;
  assert(gstr1.status === 'overdue', 'Evaluated GSTR-1 due on 11th as overdue on 12th');

  // Complete task with filing proof
  const completedTask = complianceService.completeTask(firmA, gstr1.id, 'ARN_AA270926012485C');
  assert(completedTask.success === true, 'Task completed with filing proof');
  assert(completedTask.instance?.status === 'completed', 'Task status is completed');
  assert(completedTask.instance?.completionProofRef === 'ARN_AA270926012485C', 'Filing proof recorded');

  // Export compliance calendar
  const compExport = complianceService.exportComplianceCalendar(firmA);
  assert(compExport.fileType === 'xlsx', 'Exported Statutory Compliance Calendar');

  // =========================================================================
  // MODULE 7: MULTI-CLIENT BATCH AUTOMATION
  // =========================================================================
  console.log('\n--- MODULE 7: MULTI-CLIENT BATCH AUTOMATION ---');
  const clientItems = [
    {
      clientId: firmA.clientId,
      clientName: 'Apex Technologies Pvt Ltd',
      files: [
        { fileName: 'bank_statement.csv', buffer: Buffer.from('Date,Particulars,Debit,Credit,Balance\n01-09-2026,Opening,0,0,100000\n') },
        { fileName: 'general_ledger.csv', buffer: Buffer.from('Date,Particulars,Debit,Credit,Balance\n01-09-2026,Opening,0,0,100000\n') },
      ],
    },
    {
      clientId: firmB.clientId,
      clientName: 'Bharat Logistics Hub LLP',
      files: [
        { fileName: 'bank_statement.csv', buffer: Buffer.from('Date,Particulars,Debit,Credit,Balance\n01-09-2026,Opening,0,0,250000\n') },
        { fileName: 'general_ledger.csv', buffer: Buffer.from('Date,Particulars,Debit,Credit,Balance\n01-09-2026,Opening,0,0,250000\n') },
      ],
    },
  ];

  const batchRun = await batchService.startBatchRun(firmA, 'wf_bank_rec', clientItems);
  assert(batchRun.totalClients === 2, 'Batch queued 2 clients');

  // Wait for background batch item processing
  let pollAttempts = 0;
  let finalBatch = store.getBatchRun(firmA.firmId, batchRun.id);
  while (finalBatch && finalBatch.status === 'running' && pollAttempts < 20) {
    await sleep(200);
    finalBatch = store.getBatchRun(firmA.firmId, batchRun.id);
    pollAttempts++;
  }

  assert(Boolean(finalBatch), 'Batch run found in store');
  assert(finalBatch!.completedClients === 2, 'Batch completed processing for both clients');
  assert(finalBatch!.status === 'completed', 'Batch overall status is completed');

  // Strict isolation test
  const clientAItems = finalBatch!.items.find((c) => c.clientId === firmA.clientId)!;
  const clientBItems = finalBatch!.items.find((c) => c.clientId === firmB.clientId)!;
  assert(clientAItems.jobId !== clientBItems.jobId, 'Each client executed in completely independent background jobs');

  // Export batch summary
  const batchExport = batchService.exportBatchSummary(firmA, batchRun.id);
  assert(batchExport.fileType === 'xlsx', 'Exported Consolidated Multi-Client Batch Summary');

  console.log('\n================================================================');
  console.log('🎉 ALL 7 ADVANCED AUTOMATION MODULE TESTS PASSED PERFECTLY!');
  console.log('================================================================\n');

  // Clean up test store
  if (fs.existsSync(testStorePath)) fs.unlinkSync(testStorePath);
}

runAdvancedModulesTests().catch((err) => {
  console.error('Unhandled test failure:', err);
  process.exit(1);
});
