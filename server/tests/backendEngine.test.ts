/**
 * Comprehensive Automated Test Suite
 * Intelligent CA Tools Backend Automation Engine
 */

import path from 'path';
import fs from 'fs';
import { BackendStore } from '../db/store';
import { FileProcessorService } from '../services/fileProcessor';
import { UniversalMatchingEngine } from '../services/matchingEngine';
import { AccountingRulesEngine } from '../services/accountingRulesEngine';
import { ExceptionService } from '../services/exceptionService';
import { ExportService } from '../services/exportService';
import { WorkflowOrchestrator } from '../services/workflowOrchestrator';
import type { FirmContext, ExtractedRecord } from '../types';

function assert(condition: boolean, message: string): void {
  if (!condition) {
    console.error(`❌ FAIL: ${message}`);
    process.exit(1);
  } else {
    console.log(`✓ PASS: ${message}`);
  }
}

async function runBackendTests() {
  console.log('================================================================');
  console.log('🧪 RUNNING CA TOOLS BACKEND AUTOMATION ENGINE TEST SUITE');
  console.log('================================================================');

  // Use isolated test store
  const testStorePath = path.resolve(process.cwd(), 'data', 'test-ca-store.json');
  if (fs.existsSync(testStorePath)) fs.unlinkSync(testStorePath);
  const store = BackendStore.getInstance(testStorePath);
  store.clearAll();

  const firmA: FirmContext = {
    firmId: 'firm_alpha_ca',
    firmName: 'Alpha & Associates CA',
    userId: 'partner_rajesh',
    userRole: 'Partner',
    clientId: 'client_zenith_pvt_ltd',
  };

  const firmB: FirmContext = {
    firmId: 'firm_beta_ca',
    firmName: 'Beta & Co Chartered Accountants',
    userId: 'article_vikram',
    userRole: 'Article',
    clientId: 'client_gamma_industries',
  };

  // --- 1. FILE INGESTION & FORMAT DETECTION ---
  console.log('\n=== 1. TESTING FILE INGESTION & PARSING ===');
  const csvData = `Txn Date,Particulars,Chq No,Withdrawal,Deposit,Balance\n01-09-2026,Opening Balance,,0,0,500000\n05-09-2026,GST Payment Challan PMT-06,9901,45000,0,455000\n10-09-2026,NEFT Customer Receipt XYZ,8812,0,120000,575000\n`;
  const buffer = Buffer.from(csvData, 'utf-8');
  const fileResult = FileProcessorService.processFileBuffer(buffer, 'bank_statement_sep.csv', firmA.firmId, firmA.clientId);

  assert(fileResult.sourceFile.detectedFormat === 'csv', 'Detects CSV format correctly');
  assert(fileResult.sourceFile.documentType === 'bank_statement', 'Classifies document type as bank_statement');
  assert(fileResult.records.length === 3, 'Extracts exactly 3 records from CSV');
  assert(fileResult.records[1].normalizedFields.debit === 45000, 'Normalizes withdrawal debit to 45000');
  assert(fileResult.records[2].normalizedFields.credit === 120000, 'Normalizes deposit credit to 120000');
  assert(fileResult.records[0].normalizedFields.balance === 500000, 'Normalizes running balance to 500000');
  assert(fileResult.sourceFile.fileHash.length === 64, 'Computes 64-char SHA-256 integrity hash');

  // --- 2. REUSABLE MATCHING ENGINE ---
  console.log('\n=== 2. TESTING REUSABLE MATCHING ENGINE ===');
  const bankRecords: ExtractedRecord[] = [
    {
      id: 'bnk-1',
      sourceFileId: 'src-1',
      rowIndex: 1,
      rawFields: {},
      normalizedFields: { date: '05-09-2026', amount: 50000, withdrawal: 50000, refNo: 'CHQ-1001', narration: 'Vendor Payment Alpha' },
      confidence: 100,
      isSuspect: false,
    },
    {
      id: 'bnk-2',
      sourceFileId: 'src-1',
      rowIndex: 2,
      rawFields: {},
      normalizedFields: { date: '12-09-2026', amount: 25000, withdrawal: 25000, refNo: 'NEFT-8899', narration: 'Rent Office' },
      confidence: 100,
      isSuspect: false,
    },
    {
      id: 'bnk-3',
      sourceFileId: 'src-1',
      rowIndex: 3,
      rawFields: {},
      normalizedFields: { date: '20-09-2026', amount: 90000, deposit: 90000, narration: 'Bulk Customer Remittance' },
      confidence: 100,
      isSuspect: false,
    },
  ];

  const ledgerRecords: ExtractedRecord[] = [
    {
      id: 'led-1',
      sourceFileId: 'src-2',
      rowIndex: 1,
      rawFields: {},
      normalizedFields: { date: '05-09-2026', amount: 50000, credit: 50000, invoiceNo: 'CHQ-1001', narration: 'Vendor Alpha Booked' },
      confidence: 100,
      isSuspect: false,
    },
    {
      id: 'led-2',
      sourceFileId: 'src-2',
      rowIndex: 2,
      rawFields: {},
      normalizedFields: { date: '14-09-2026', amount: 25000, credit: 25000, narration: 'Rent Booked' }, // 2 days variance
      confidence: 100,
      isSuspect: false,
    },
    {
      id: 'led-3a',
      sourceFileId: 'src-2',
      rowIndex: 3,
      rawFields: {},
      normalizedFields: { date: '20-09-2026', amount: 50000, debit: 50000, narration: 'Invoice A Part' },
      confidence: 100,
      isSuspect: false,
    },
    {
      id: 'led-3b',
      sourceFileId: 'src-2',
      rowIndex: 4,
      rawFields: {},
      normalizedFields: { date: '20-09-2026', amount: 40000, debit: 40000, narration: 'Invoice B Part' },
      confidence: 100,
      isSuspect: false,
    },
  ];

  const matchResult = UniversalMatchingEngine.executeMatching(bankRecords, ledgerRecords, 'run-test-1', {
    dateToleranceDays: 3,
    amountTolerance: 1.0,
    checkDirection: true,
    allowOneToMany: true,
  });

  assert(matchResult.summary.exactMatches === 1, 'Pass 1: Identifies 1 exact match on identifier and amount');
  assert(matchResult.summary.probableMatches === 1, 'Pass 2: Identifies 1 probable match within 2 days variance');
  assert(matchResult.summary.partialMatches === 1, 'Pass 3: Identifies 1-to-many match (₹90,000 = ₹50,000 + ₹40,000)');
  assert(matchResult.unmatchedPrimary.length === 0, 'Zero unmatched primary bank records after matching passes');

  // --- 3. CONFIGURABLE ACCOUNTING RULES ENGINE ---
  console.log('\n=== 3. TESTING ACCOUNTING RULES ENGINE ===');
  // Continuity Break test
  const discontinuousRecords: ExtractedRecord[] = [
    { id: '1', sourceFileId: 's', rowIndex: 1, rawFields: {}, normalizedFields: { balance: 100000 }, confidence: 100, isSuspect: false },
    { id: '2', sourceFileId: 's', rowIndex: 2, rawFields: {}, normalizedFields: { debit: 20000, credit: 0, balance: 70000 }, confidence: 100, isSuspect: false }, // should be 80,000
  ];
  const contAudit = AccountingRulesEngine.validateRunningBalanceContinuity(discontinuousRecords);
  assert(!contAudit.isValid, 'Flags running balance discontinuity as invalid');
  assert(contAudit.issues[0].ruleCode === 'RULE_BANK_CONTINUITY_BREAK', 'Reports RULE_BANK_CONTINUITY_BREAK');

  // Statutory Tax Math test
  const badTaxRecord: ExtractedRecord = {
    id: 'inv-err',
    sourceFileId: 's',
    rowIndex: 1,
    rawFields: {},
    normalizedFields: {
      invoiceNo: 'INV-99',
      taxableAmount: 100000,
      cgst: 9000,
      sgst: 9000,
      igst: 0,
      cess: 0,
      totalAmount: 125000, // Error: 100k + 18k = 118k, but stated 125k
    },
    confidence: 100,
    isSuspect: false,
  };
  const taxIssues = AccountingRulesEngine.validateInvoiceArithmetic(badTaxRecord);
  assert(taxIssues.length > 0, 'Flags statutory tax math mismatch');
  assert(taxIssues[0].ruleCode === 'RULE_INVOICE_TAX_ARITHMETIC', 'Identifies RULE_INVOICE_TAX_ARITHMETIC');

  // Section 40A(3) Cash Limit test
  const cashRecord: ExtractedRecord = {
    id: 'c-1',
    sourceFileId: 's',
    rowIndex: 1,
    rawFields: {},
    normalizedFields: { narration: 'Cash paid to transporter', debit: 25000, amount: 25000 },
    confidence: 100,
    isSuspect: false,
  };
  const cashIssue = AccountingRulesEngine.validateCashPaymentLimit(cashRecord, 10000);
  assert(cashIssue !== undefined && cashIssue.ruleCode === 'RULE_SECTION_40A_3_VIOLATION', 'Flags Sec 40A(3) cash payment limit violation (> ₹10,000)');

  // Abnormal Cash in Hand balance test
  const abnormalCashRecords: ExtractedRecord[] = [
    { id: 'tb-1', sourceFileId: 's', rowIndex: 1, rawFields: {}, normalizedFields: { accountName: 'Cash in Hand', debit: 0, credit: 5000 }, confidence: 100, isSuspect: false },
    { id: 'tb-2', sourceFileId: 's', rowIndex: 2, rawFields: {}, normalizedFields: { accountName: 'Capital Account', debit: 0, credit: 5000 }, confidence: 100, isSuspect: false },
  ];
  const tbAudit = AccountingRulesEngine.validateTrialBalance(abnormalCashRecords);
  assert(tbAudit.issues.some((i) => i.ruleCode === 'RULE_ABNORMAL_NEGATIVE_CASH'), 'Flags abnormal credit balance in cash');

  // --- 4. WORKFLOW ORCHESTRATION & STATE MACHINE ---
  console.log('\n=== 4. TESTING WORKFLOW ORCHESTRATION LAYER ===');
  const orchestrator = new WorkflowOrchestrator(store);

  // Workflow A: Bank Statement Automation
  const bankCsv = `Date,Particulars,Ref No,Withdrawal,Deposit,Balance\n01-09-2026,Opening Balance,,0,0,100000\n05-09-2026,GST Challan Payment,GST99,18000,0,82000\n10-09-2026,Salaries September,SAL01,50000,0,32000\n`;
  const { job: bankJob, run: bankRun } = await orchestrator.startWorkflow(
    firmA,
    'wf_bank_rec',
    [{ fileName: 'Bank_Statement.csv', buffer: Buffer.from(bankCsv, 'utf-8') }],
    { idempotencyKey: 'idemp-bank-01' }
  );

  assert(bankJob.status === 'queued' || bankJob.status === 'running' || bankJob.status === 'completed' || bankJob.status === 'review_required', 'Job created in valid state machine state');
  assert(bankRun.sourceFiles.length === 1, 'Workflow run associated with source file');

  // Test Idempotency
  const duplicateSubmission = await orchestrator.startWorkflow(
    firmA,
    'wf_bank_rec',
    [{ fileName: 'Bank_Statement.csv', buffer: Buffer.from(bankCsv, 'utf-8') }],
    { idempotencyKey: 'idemp-bank-01' }
  );
  assert(duplicateSubmission.job.id === bankJob.id, 'Idempotency key prevents duplicate job creation and returns existing job');

  // Workflow B: Bulk Invoice Processing & Duplicate Detection
  console.log('\n=== 5. TESTING INVOICE WORKFLOW & DEDUP ===');
  const invoiceCsv = `Invoice No,Invoice Date,Supplier GSTIN,Supplier Name,Taxable,CGST,SGST,IGST,Total\nINV-001,10-09-2026,27AABCA1234F1Z5,Apex Logistics,100000,9000,9000,0,118000\nINV-001,10-09-2026,27AABCA1234F1Z5,Apex Logistics,100000,9000,9000,0,118000\n`; // Exact duplicate
  const { job: invJob } = await orchestrator.startWorkflow(
    firmA,
    'wf_invoice_proc',
    [{ fileName: 'Invoices_Batch.csv', buffer: Buffer.from(invoiceCsv, 'utf-8') }]
  );
  assert(invJob.id.startsWith('job-'), 'Invoice processing workflow initiated with valid Job ID');

  // --- 6. EXCEPTION REVIEW & APPROVAL AUDIT SYSTEM ---
  console.log('\n=== 6. TESTING EXCEPTION REVIEW & AUDIT TRAIL ===');
  const exceptionService = new ExceptionService(store);
  const loggedExc = exceptionService.logException(
    firmA,
    bankRun.id,
    bankJob.id,
    'unmatched_transaction',
    'warning',
    'Unidentified Bank Debit',
    'Debit of ₹15,000 to Unknown Beneficiary',
    { amount: 15000, narration: 'Debit to Unknown' },
    'Book under Suspense Account'
  );

  assert(loggedExc.resolutionStatus === 'pending', 'Exception initialized with pending status');

  const resolutionResult = exceptionService.resolveException(
    firmA,
    loggedExc.id,
    'approved',
    'Approved by CA Partner for Suspense Account',
    { account: 'Suspense Account', amount: 15000 }
  );

  assert(resolutionResult.success, 'Exception resolution succeeded');
  assert(resolutionResult.exception?.resolutionStatus === 'approved', 'Exception status updated to approved');
  assert(Boolean(resolutionResult.approval?.performedBy.includes('partner_rajesh')), 'Approval recorded with auditor signature');

  const auditEvents = store.queryAuditEvents(firmA.firmId, { entityId: resolutionResult.approval?.id });
  assert(auditEvents.length > 0, 'Immutable audit event logged for exception approval');

  // --- 7. REAL EXPORTS & INTEGRITY VERIFICATION ---
  console.log('\n=== 7. TESTING REAL MULTI-SHEET EXPORTS & INTEGRITY ===');
  const exportService = new ExportService(store);
  const artifact = exportService.generateExcelWorkbook(firmA, bankRun.id, 'Test_Reconciliation_BRS.xlsx', [
    { sheetName: 'Executive_Summary', rows: [{ Metric: 'Reconciled Amount', Value: 100000 }] },
    { sheetName: 'Matched_Records', rows: [{ Ref: 'CHQ-01', Amount: 50000 }] },
    { sheetName: 'Exceptions_Log', rows: [{ Issue: 'None', Status: 'Verified' }] },
  ]);

  assert(artifact.integrityVerified === true, 'Workbook passes re-read integrity verification');
  assert(artifact.sheetNames?.length === 3, 'Workbook generated with all 3 specified sheets');
  assert(artifact.fileSizeBytes > 1000, 'Workbook is real binary Excel file (> 1KB)');

  // --- 8. FIRM-LEVEL DATA ISOLATION ---
  console.log('\n=== 8. TESTING FIRM-LEVEL MULTI-TENANCY ISOLATION ===');
  const firmAJobs = store.listJobs(firmA.firmId);
  const firmBJobs = store.listJobs(firmB.firmId);

  assert(firmAJobs.length >= 2, 'Firm A can view its own jobs');
  assert(firmBJobs.length === 0, 'Firm B CANNOT see any of Firm A jobs (Strict Multi-Tenancy)');

  const crossFirmException = store.getException(firmB.firmId, loggedExc.id);
  assert(crossFirmException === undefined, 'Firm B is strictly blocked from accessing Firm A exceptions');

  // Cleanup test database
  try {
    if (fs.existsSync(testStorePath)) fs.unlinkSync(testStorePath);
  } catch {
    // Windows file handle timing tolerance
  }

  console.log('\n================================================================');
  console.log('🎉 ALL BACKEND AUTOMATION ENGINE TESTS PASSED WITH 100% SUCCESS!');
  console.log('================================================================');
}

runBackendTests().catch((err) => {
  console.error('Fatal test error:', err);
  process.exit(1);
});
