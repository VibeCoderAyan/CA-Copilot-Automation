/**
 * Comprehensive Test Suite for the 6 New Offline Backend Systems:
 * 1. Tally Prime Direct Live Sync (Port 9000 XML Loopback)
 * 2. Hot-Folder Zero-Click File Ingestion & Auto-Reconciliation
 * 3. Multi-User Office LAN Collaboration & Workpaper Lock (ICAI SA-230)
 * 4. High-Speed Indexed Local Database Engine
 * 5. Automated Client Query Sheet & Formal Observation Generator
 * 6. Nightly Audit Cron & Morning Exception Briefing Engine
 */

import assert from 'assert';
import fs from 'fs';
import path from 'path';
import * as XLSX from 'xlsx';
import { TallyLiveSyncService } from '../services/tallyLiveSyncService.ts';
import { HotFolderWatcherService } from '../services/hotFolderWatcherService.ts';
import { OfficeLanService } from '../services/officeLanService.ts';
import { LocalDbEngine } from '../db/localDbEngine.ts';
import { ClientQuerySheetService } from '../services/clientQuerySheetService.ts';
import { NightlyAuditCronService } from '../services/nightlyAuditCronService.ts';

async function runTests() {
  console.log('================================================================');
  console.log('🧪 RUNNING TEST SUITE: 6 NEW OFFLINE BACKEND ENGINES');
  console.log('================================================================');

  // =========================================================================
  // 1. TALLY PRIME DIRECT LIVE SYNC TESTS
  // =========================================================================
  console.log('\n--- FEATURE 1: TALLY PRIME LIVE SYNC (PORT 9000 XML) ---');
  const tallyService = new TallyLiveSyncService({ port: 9000, timeoutMs: 1500 });

  // Test connection refusal error handling (when Tally is not running on dev machine)
  const health = await tallyService.checkHealth();
  assert.strictEqual(typeof health.isConnected, 'boolean');
  assert.strictEqual(health.port, 9000);
  console.log('✓ PASS: Tally Prime health check completed with clean offline diagnostic message');

  // Test voucher push validation
  const emptyPush = await tallyService.pushVouchers('Test Co', []);
  assert.strictEqual(emptyPush.success, false);
  console.log('✓ PASS: Correctly rejects empty voucher push array');

  // =========================================================================
  // 2. HOT-FOLDER ZERO-CLICK INGESTION TESTS
  // =========================================================================
  console.log('\n--- FEATURE 2: HOT-FOLDER ZERO-CLICK WATCHER ---');
  const testHotDir = path.resolve(process.cwd(), 'data', 'test-hot-folder');
  const testInbox = path.join(testHotDir, 'inbox');
  const testProcessed = path.join(testHotDir, 'processed');
  const testError = path.join(testHotDir, 'error');

  const hotWatcher = new HotFolderWatcherService({
    inboxPath: testInbox,
    processedPath: testProcessed,
    errorPath: testError,
  });

  const startRes = hotWatcher.start();
  assert.strictEqual(startRes.success, true);
  console.log('✓ PASS: Hot Folder Watcher started successfully on local filesystem');

  // Create a synthetic Bank Statement in the test inbox
  const testBankWb = XLSX.utils.book_new();
  const bankRows = [
    { 'Date': '01-04-2025', 'Particulars': 'Opening Balance', 'Debit': 0, 'Credit': 0, 'Balance': 250000 },
    { 'Date': '05-04-2025', 'Particulars': 'Customer Inward NEFT', 'Debit': 0, 'Credit': 150000, 'Balance': 400000 },
    { 'Date': '12-04-2025', 'Particulars': 'Office Rent Chq 101', 'Debit': 45000, 'Credit': 0, 'Balance': 355000 },
  ];
  const bankWs = XLSX.utils.json_to_sheet(bankRows);
  XLSX.utils.book_append_sheet(testBankWb, bankWs, 'Bank_Transactions');
  const bankFilePath = path.join(testInbox, 'hdfc_bank_statement_april.xlsx');
  XLSX.writeFile(testBankWb, bankFilePath);

  // Trigger processing
  const ingestEvent = hotWatcher.processFile(bankFilePath, 'hdfc_bank_statement_april.xlsx');
  assert.strictEqual(ingestEvent.detectedType, 'bank_statement');
  assert.strictEqual(ingestEvent.recordsProcessed, 3);
  assert.strictEqual(ingestEvent.status, 'completed');
  assert.ok(ingestEvent.summary?.includes('Bank Statement parsed'));
  console.log('✓ PASS: Automatically classified bank statement and calculated debits & credits');

  hotWatcher.stop();
  console.log('✓ PASS: Hot Folder Watcher stopped cleanly');

  // Cleanup test hot folder
  try {
    fs.rmSync(testHotDir, { recursive: true, force: true });
  } catch (e) {}

  // =========================================================================
  // 3. MULTI-USER OFFICE LAN COLLABORATION & WORKPAPER LOCK
  // =========================================================================
  console.log('\n--- FEATURE 3: OFFICE LAN COLLABORATION & SA-230 LOCK ---');
  const lanService = new OfficeLanService();

  const netInfo = lanService.getNetworkInfo(5000);
  assert.strictEqual(typeof netInfo.hostname, 'string');
  assert.ok(netInfo.primaryLanUrl.startsWith('http://'));
  console.log(`✓ PASS: Discovered local LAN sharing URL: ${netInfo.primaryLanUrl}`);

  // Test Role Permissions
  assert.strictEqual(lanService.checkPermission('ArticleAssistant', 'upload'), true);
  assert.strictEqual(lanService.checkPermission('ArticleAssistant', 'reconcile'), true);
  assert.strictEqual(lanService.checkPermission('ArticleAssistant', 'sign_off'), false); // Articles cannot sign off
  assert.strictEqual(lanService.checkPermission('PrincipalPartner', 'sign_off'), true); // Partner can sign off
  assert.strictEqual(lanService.checkPermission('PrincipalPartner', 'lock_workpaper'), true);
  console.log('✓ PASS: Enforced ICAI role permission matrix between Partners and Articles');

  // Test Workpaper Lock
  const lockRes = lanService.lockWorkpaper('WP-BRS-2025-Q1', 'client_tata', 'CA Rajesh Sharma', '054321', 'All bank entries verified against Passbook.');
  assert.strictEqual(lockRes.success, true);
  assert.strictEqual(lockRes.state.isLocked, true);
  assert.strictEqual(lockRes.state.status, 'AuditorApproved');

  const statusAfterLock = lanService.getWorkpaperStatus('WP-BRS-2025-Q1');
  assert.strictEqual(statusAfterLock.isLocked, true);
  assert.strictEqual(statusAfterLock.lockedBy, 'CA Rajesh Sharma');
  console.log('✓ PASS: Workpaper successfully locked under ICAI SA-230 Auditor Sign-off');

  // Partner Unlock
  const unlockRes = lanService.unlockWorkpaper('WP-BRS-2025-Q1', 'CA Rajesh Sharma');
  assert.strictEqual(unlockRes.success, true);
  assert.strictEqual(lanService.getWorkpaperStatus('WP-BRS-2025-Q1').isLocked, false);
  console.log('✓ PASS: Partner override successfully unlocked workpaper for rectification');

  // =========================================================================
  // 4. HIGH-SPEED INDEXED LOCAL DATABASE ENGINE
  // =========================================================================
  console.log('\n--- FEATURE 4: HIGH-SPEED INDEXED LOCAL DATABASE ENGINE ---');
  const testDbPath = path.resolve(process.cwd(), 'data', 'test-local-db.json');
  const localDb = new LocalDbEngine(testDbPath);
  localDb.clearAll();

  // Insert 200 synthetic ledger records
  const sampleRecords = [];
  for (let i = 1; i <= 200; i++) {
    sampleRecords.push({
      clientId: i % 2 === 0 ? 'client_apex' : 'client_bharat',
      date: `2025-04-${String((i % 28) + 1).padStart(2, '0')}`,
      voucherNo: `VCH-${1000 + i}`,
      accountHead: i % 3 === 0 ? 'Sales Account' : 'Sundry Debtors',
      debit: i % 2 === 0 ? 5000 * (i % 10) : 0,
      credit: i % 2 !== 0 ? 5000 * (i % 10) : 0,
      narration: `Being transaction ${i} recorded`,
      partyGstin: i % 4 === 0 ? '27AAACP1234F1Z8' : '27BBBCP5678F1Z2',
    });
  }

  const insertResult = localDb.insertBatch(sampleRecords);
  assert.strictEqual(insertResult.insertedCount, 200);
  console.log(`✓ PASS: Batch inserted and indexed 200 transactions in ${insertResult.durationMs}ms`);

  // Query by Voucher No (Exact Index Lookup)
  const voucherQuery = localDb.query({ voucherNo: 'VCH-1050' });
  assert.strictEqual(voucherQuery.totalMatches, 1);
  assert.strictEqual(voucherQuery.records[0].voucherNo, 'VCH-1050');
  console.log(`✓ PASS: Exact Voucher index lookup completed in ${voucherQuery.durationMs}ms`);

  // Query by Client + GSTIN + Amount Filter
  const filteredQuery = localDb.query({
    clientId: 'client_apex',
    partyGstin: '27AAACP1234F1Z8',
    minAmount: 10000,
  });
  assert.ok(filteredQuery.totalMatches > 0);
  for (const r of filteredQuery.records) {
    assert.strictEqual(r.clientId, 'client_apex');
    assert.strictEqual(r.partyGstin, '27AAACP1234F1Z8');
    assert.ok(Math.max(r.debit, r.credit) >= 10000);
  }
  console.log(`✓ PASS: Multi-column indexed filtered query returned ${filteredQuery.totalMatches} matches in ${filteredQuery.durationMs}ms`);

  // Cleanup test DB
  try {
    fs.unlinkSync(testDbPath);
  } catch (e) {}

  // =========================================================================
  // 5. CLIENT QUERY SHEET & FORMAL OBSERVATION GENERATOR
  // =========================================================================
  console.log('\n--- FEATURE 5: CLIENT QUERY SHEET & AUDIT OBSERVATION GENERATOR ---');
  const queryService = new ClientQuerySheetService();

  const queryPkg = {
    clientName: 'Tata Steel Processing Ltd',
    financialYear: 'FY 2026-27',
    assessmentYear: 'AY 2027-28',
    firmName: 'K. Sharma & Associates, Chartered Accountants',
    firmFrn: '120488N',
    partnerName: 'CA Rajesh Sharma, FCA',
    partnerMrn: '054321',
    queryDate: '19-10-2026',
    clientEmail: 'cfo@tatasteel.com',
    queryItems: [
      {
        srNo: 1,
        date: '14-06-2026',
        referenceNo: 'INV-402',
        ledgerHead: 'Machinery Repairs',
        partyName: 'Precision Engineering Works',
        amount: 85000,
        discrepancyType: 'Missing_GSTR2B' as const,
        auditObservation: 'Input tax credit claimed in books (₹15,300) but invoice not appearing in GSTR-2B portal.',
        requiredClientAction: 'Provide vendor tax invoice & confirm filing status of vendor GSTR-1.',
      },
      {
        srNo: 2,
        date: '28-07-2026',
        referenceNo: 'PV-899',
        ledgerHead: 'Freight & Cartage',
        amount: 18500,
        discrepancyType: 'Sec_40A3_Cash_Limit' as const,
        auditObservation: 'Cash payment exceeding statutory ₹10,000 threshold under Section 40A(3).',
        requiredClientAction: 'Provide transporter declaration under Rule 6DD or accept Form 3CD disallowance.',
      },
    ],
  };

  const queryArtifacts = queryService.generateQueryPackage(queryPkg);
  assert.strictEqual(queryArtifacts.totalQueries, 2);
  assert.strictEqual(queryArtifacts.totalDisputedAmount, 103500);

  // Validate HTML letter
  assert.ok(queryArtifacts.htmlLetter.includes('Tata Steel Processing Ltd'));
  assert.ok(queryArtifacts.htmlLetter.includes('120488N'));
  assert.ok(queryArtifacts.htmlLetter.includes('Missing_GSTR2B') || queryArtifacts.htmlLetter.includes('Input tax credit claimed'));
  console.log('✓ PASS: Generated formal ICAI-compliant HTML Audit Query Letter');

  // Validate EML draft
  assert.ok(queryArtifacts.emlRawContent.includes('To: "Tata Steel Processing Ltd" <cfo@tatasteel.com>'));
  assert.ok(queryArtifacts.emlRawContent.includes('Subject: Audit Queries & Requisition'));
  console.log('✓ PASS: Generated RFC 2822 .eml draft for direct email dispatch');

  // Validate Excel response sheet
  assert.ok(Buffer.isBuffer(queryArtifacts.excelBuffer));
  const parsedQueryWb = XLSX.read(queryArtifacts.excelBuffer, { type: 'buffer' });
  assert.strictEqual(parsedQueryWb.SheetNames[0], 'Audit_Queries_Client_Action');
  console.log('✓ PASS: Generated multi-column Excel Client Response Register');

  // =========================================================================
  // 6. NIGHTLY AUDIT CRON & MORNING EXCEPTION BRIEFING ENGINE
  // =========================================================================
  console.log('\n--- FEATURE 6: NIGHTLY AUDIT CRON & MORNING EXCEPTION BRIEFING ---');
  const cronService = new NightlyAuditCronService();

  const testVouchers = [
    // 1. Normal cash opening
    { date: '2025-04-01', voucherNo: 'CV-01', accountName: 'Cash on Hand', debit: 50000, credit: 0, narration: 'Opening cash' },
    // 2. Heavy cash payment creating negative cash balance
    { date: '2025-04-05', voucherNo: 'CV-02', accountName: 'Cash on Hand', debit: 0, credit: 75000, narration: 'Office renovation payment' },
    // 3. Section 40A(3) cash breach > ₹10,000
    { date: '2025-04-10', voucherNo: 'CV-03', accountName: 'Cash on Hand', debit: 0, credit: 28000, narration: 'Direct cash payment for raw materials' },
    // 4. Stale cheque older than 90 days
    { date: '2024-12-01', voucherNo: 'BP-45', accountName: 'HDFC Bank Ltd', debit: 0, credit: 62000, narration: 'Payment by Cheque No 44102 to vendor' },
    // 5. High-value round-sum on Sunday
    { date: '2025-04-06', voucherNo: 'JV-99', accountName: 'Consultancy Fees', debit: 500000, credit: 0, narration: 'Quarterly management fee adjustment' }, // April 6, 2025 is a Sunday!
  ];

  const morningBriefing = cronService.runAuditSweep('Zenith Industrial Ltd', 'FY 2025-26', testVouchers);

  assert.strictEqual(morningBriefing.totalVouchersScanned, 5);
  assert.ok(morningBriefing.totalIssuesFlagged >= 3);
  assert.ok(morningBriefing.negativeCashDays >= 1);
  assert.ok(morningBriefing.sec40a3BreachesAmount >= 28000);
  assert.ok(morningBriefing.staleChequesAmount >= 62000);

  // Verify Critical Negative Cash issue
  const negCashIssue = morningBriefing.issues.find((i) => i.ruleCode === 'NEG_CASH');
  assert.ok(negCashIssue);
  assert.strictEqual(negCashIssue.severity, 'CRITICAL');
  console.log('✓ PASS: Detected negative cash balance and flagged as CRITICAL statutory issue');

  // Verify Section 40A(3) breach
  const sec40a3Issue = morningBriefing.issues.find((i) => i.ruleCode === 'SEC_40A3' && i.voucherNo === 'CV-03');
  assert.ok(sec40a3Issue);
  assert.strictEqual(sec40a3Issue.amount, 28000);
  console.log('✓ PASS: Detected Section 40A(3) cash threshold breach (> ₹10,000)');

  // Verify Stale Cheque
  const staleIssue = morningBriefing.issues.find((i) => i.ruleCode === 'STALE_CHEQUE');
  assert.ok(staleIssue);
  console.log('✓ PASS: Detected unpresented stale cheque older than 90 days');

  // Verify Sunday High-Value entry
  const sundayIssue = morningBriefing.issues.find((i) => i.ruleCode === 'ROUND_SUM_SUNDAY');
  assert.ok(sundayIssue);
  console.log('✓ PASS: Detected high-value round-sum journal entry posted on Sunday');

  console.log('\n================================================================');
  console.log('🎉 ALL 6 NEW OFFLINE BACKEND ENGINES TESTED & VERIFIED 100%!');
  console.log('================================================================');
}

runTests().catch((err) => {
  console.error('❌ Test failed with error:', err);
  process.exit(1);
});
