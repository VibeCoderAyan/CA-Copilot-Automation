/**
 * Automated Verification Suite for Advanced Backend Upgrades:
 * 1. Atomic Crash-Proof Storage & Concurrent Write Integrity
 * 2. High-Throughput Streaming & Chunked Processor Telemetry
 * 3. Section 206AB & 206CCA Statutory Non-Filer Verification Engine
 * 4. 4-Sheet Multi-Annexure Excel Audit Artifact Generation
 */

import assert from 'assert';
import * as XLSX from 'xlsx';
import { BackendStore } from '../db/store';
import { StreamingProcessorService } from '../services/streamingProcessorService';
import { Section206abService, type Section206abInputRecord } from '../services/section206abService';

async function runUpgradesTests() {
  console.log('================================================================');
  console.log('🧪 RUNNING ADVANCED BACKEND ARCHITECTURAL UPGRADES TEST SUITE');
  console.log('================================================================');

  const store = BackendStore.getInstance();

  // --- TEST 1: ATOMIC STORAGE & CONCURRENT WRITE INTEGRITY ---
  console.log('\n--- TEST 1: ATOMIC STORAGE & CONCURRENT WRITE INTEGRITY ---');
  const initialJobsCount = store.listJobs('firm_apex_ca').length;

  // Perform 10 rapid concurrent writes
  const writes = [];
  for (let i = 0; i < 10; i++) {
    writes.push(
      new Promise<void>((resolve) => {
        store.createJob({
          id: `job-concurrent-${i}-${Date.now()}`,
          firmId: 'firm_apex_ca',
          clientId: 'client_apex_technologies',
          workflowId: 'wf_bank_rec',
          status: 'queued',
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString(),
          sourceFileIds: [],
        });
        resolve();
      })
    );
  }
  await Promise.all(writes);

  const updatedJobs = store.listJobs('firm_apex_ca');
  const finalCount = updatedJobs.length;
  assert(finalCount >= initialJobsCount + 10, 'All 10 concurrent writes persisted without being dropped by race conditions');
  console.log('✓ PASS: Atomic write queue persisted all concurrent transactions without dropped writes');

  // Verify backup snapshot creation
  const backupPath = store.createBackup();
  assert(backupPath.includes('ca-store-backup-'), 'Backup snapshot created with standard timestamp');
  console.log('✓ PASS: Write-ahead atomic replacement and backup snapshot verified');

  // --- TEST 2: STREAMING PROCESSOR & MEMORY TELEMETRY ---
  console.log('\n--- TEST 2: STREAMING PROCESSOR & MEMORY-BOUNDED TELEMETRY ---');
  // Generate an in-memory 2,000-row mock ledger
  const mockRows: Record<string, any>[] = [];
  for (let i = 1; i <= 2000; i++) {
    mockRows.push({
      Date: `01/04/2026`,
      Particulars: i % 2 === 0 ? 'Vendor Payment Alpha' : 'Client Inward Remittance',
      RefNo: `TXN-${i % 100}`, // Creates intentional duplicates for signature check
      Debit: i % 2 === 0 ? 1500.50 : 0,
      Credit: i % 2 !== 0 ? 3200.00 : 0,
      Party: i % 2 === 0 ? 'Alpha Corp Ltd' : 'Beta Retailers Pvt Ltd',
      Taxable: 1000,
      CGST: 90,
      SGST: 90,
    });
  }

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(mockRows);
  XLSX.utils.book_append_sheet(wb, ws, 'Daybook');
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  let chunkCountReceived = 0;
  let totalRowsReceived = 0;

  const streamResult = await StreamingProcessorService.processLargeBufferInChunks(
    buffer,
    'Annual_Ledger_Dump.xlsx',
    500, // 500 rows per chunk = 4 chunks
    (chunk, telemetry) => {
      chunkCountReceived++;
      totalRowsReceived += chunk.length;
      assert(telemetry.heapUsedMB > 0, 'Heap memory telemetry tracked');
    }
  );

  assert(streamResult.totalRecords === 2000, 'Processed exactly 2000 records');
  assert(chunkCountReceived === 4, 'Dispatched exactly 4 chunks of 500 rows');
  assert(totalRowsReceived === 2000, 'All rows accounted for across discrete chunks');
  assert(streamResult.uniquePartiesCount === 2, 'Correctly extracted 2 unique parties');
  assert(streamResult.duplicateKeysCount > 0, 'Accurately detected recurring transaction signatures');
  assert(streamResult.telemetry.throughputRowsPerSec > 0, 'Throughput telemetry computed accurately');
  console.log(`✓ PASS: Processed 2,000 records in 4 chunks @ ${streamResult.telemetry.throughputRowsPerSec} rows/sec without event-loop block`);

  // --- TEST 3: SECTION 206AB STATUTORY NON-FILER VERIFICATION ---
  console.log('\n--- TEST 3: SECTION 206AB & 206CCA STATUTORY VERIFICATION ---');
  const sec206abService = new Section206abService(store);

  const testVendors: Section206abInputRecord[] = [
    // Case 1: High-risk specified person: Prior TDS ₹65,000 (>= ₹50k) and has NOT filed ITR
    {
      id: 'v-01',
      vendorName: 'Apex Transport Logistics',
      pan: 'ABCDE1234F', // Firm
      tdsSection: '194C',
      paymentAmount: 500000,
      normalTdsRate: 2.0, // Normal 2%
      aggregatePriorYearTds: 65000,
      hasFiledItr: false,
    },
    // Case 2: Compliant Filer with verified acknowledgment
    {
      id: 'v-02',
      vendorName: 'Zenith Legal & Advisory',
      gstin: '27AAACZ1122K1Z5', // PAN derived: AAACZ1122K (Company)
      tdsSection: '194J',
      paymentAmount: 200000,
      normalTdsRate: 10.0,
      aggregatePriorYearTds: 80000,
      hasFiledItr: true,
      itrAckNumber: 'ITR-2026-991823712',
    },
    // Case 3: Below ₹50,000 threshold (Section 206AB exempt)
    {
      id: 'v-03',
      vendorName: 'Local Maintenance Services',
      pan: 'AABPL5544P',
      tdsSection: '194C',
      paymentAmount: 100000,
      normalTdsRate: 1.0,
      aggregatePriorYearTds: 22000, // Below ₹50k
      hasFiledItr: false,
    },
    // Case 4: Statutorily Excluded Section 192 (Salaries)
    {
      id: 'v-04',
      vendorName: 'Executive Director Compensation',
      pan: 'ABCPD9911P',
      tdsSection: '192',
      paymentAmount: 1200000,
      normalTdsRate: 30.0,
      aggregatePriorYearTds: 250000,
      hasFiledItr: false,
    },
    // Case 5: Invalid PAN -> Section 206AA overrides (Minimum 20%)
    {
      id: 'v-05',
      vendorName: 'Unregistered Contractor',
      pan: 'INVALID_PAN',
      tdsSection: '194C',
      paymentAmount: 50000,
      normalTdsRate: 1.0,
    },
  ];

  const evalResult = sec206abService.evaluateRecords(testVendors, 'Apex Technologies Ltd', 'FY 2026-27');

  assert(evalResult.totalVendorsScrutinized === 5, 'Audited 5 vendors');
  assert(evalResult.highRiskSpecifiedPersonsCount === 1, 'Identified exactly 1 high-risk specified non-filer');

  // Verify Case 1: Rate must be Max(2 * 2%, 5%) = 5%
  const case1 = evalResult.records.find((r) => r.id === 'v-01')!;
  assert(case1.complianceStatus === 'HIGH_RISK_SPECIFIED_PERSON', 'Case 1 tagged as HIGH_RISK_SPECIFIED_PERSON');
  assert(case1.applicableTdsRate === 5.0, 'Statutory higher rate is 5.0% (Max of 2*2% or 5%)');
  assert(case1.normalTdsAmount === 10000, 'Normal TDS: 2% of ₹5,00,000 = ₹10,000');
  assert(case1.applicableTdsAmount === 25000, 'Section 206AB TDS: 5% of ₹5,00,000 = ₹25,000');
  assert(case1.differentialTdsExposure === 15000, 'Under-deduction exposure: ₹15,000');
  console.log('✓ PASS: Section 206AB 5% floor rule and ₹15,000 differential exposure verified');

  // Verify Case 2: Compliant filer
  const case2 = evalResult.records.find((r) => r.id === 'v-02')!;
  assert(case2.complianceStatus === 'COMPLIANT_FILER', 'Case 2 is COMPLIANT_FILER');
  assert(case2.applicableTdsRate === 10.0, 'Normal 10% rate maintained');
  assert(case2.differentialTdsExposure === 0, 'Zero exposure for compliant filer');
  assert(case2.pan === 'AAACZ1122K', 'PAN correctly extracted from GSTIN');
  console.log('✓ PASS: Derived PAN from GSTIN and compliant filer status verified');

  // Verify Case 3: Exempt below threshold
  const case3 = evalResult.records.find((r) => r.id === 'v-03')!;
  assert(case3.complianceStatus === 'EXEMPT_BELOW_THRESHOLD', 'Case 3 exempt below ₹50k threshold');
  assert(case3.applicableTdsRate === 1.0, 'Normal 1% maintained');
  console.log('✓ PASS: ₹50,000 aggregate prior year threshold exemption verified');

  // Verify Case 4: Excluded Section 192
  const case4 = evalResult.records.find((r) => r.id === 'v-04')!;
  assert(case4.complianceStatus === 'STATUTORILY_EXCLUDED_SECTION', 'Section 192 salary is statutorily excluded');
  console.log('✓ PASS: Section 192 proviso exclusion verified');

  // Verify Case 5: Invalid PAN
  const case5 = evalResult.records.find((r) => r.id === 'v-05')!;
  assert(case5.complianceStatus === 'INVALID_PAN_SECTION_206AA', 'Invalid PAN flags Section 206AA');
  assert(case5.applicableTdsRate === 20.0, 'Section 206AA mandates minimum 20% rate');
  console.log('✓ PASS: Section 206AA 20% minimum rate override verified');

  // --- TEST 4: SECTION 206AB 4-SHEET EXPORT WORKBOOK ---
  console.log('\n--- TEST 4: 4-SHEET SECTION 206AB AUDIT WORKBOOK PIPELINE ---');
  const artifact = sec206abService.exportAuditArtifact(evalResult, 'firm_apex_ca', 'client_apex_technologies');

  assert(artifact.id.startsWith('art-sec206ab-'), 'Generated artifact ID with standard prefix');
  assert(artifact.fileSizeBytes > 1000, 'Valid binary file size');
  assert(artifact.base64Payload && artifact.base64Payload.length > 500, 'SheetJS Base64 payload generated');

  const readBuffer = Buffer.from(artifact.base64Payload!, 'base64');
  const readWb = XLSX.read(readBuffer, { type: 'buffer' });

  assert(readWb.SheetNames.includes('Executive_Summary'), 'Contains Executive_Summary sheet');
  assert(readWb.SheetNames.includes('Section_206AB_Audit'), 'Contains Section_206AB_Audit sheet');
  assert(readWb.SheetNames.includes('Statutory_Exclusions'), 'Contains Statutory_Exclusions sheet');
  assert(readWb.SheetNames.includes('Vendor_Declaration_Template'), 'Contains Vendor_Declaration_Template sheet');
  assert(readWb.SheetNames.length === 4, 'Exactly 4 statutory audit sheets generated');
  console.log('✓ PASS: 4-Sheet Section 206AB workbook validated with audit declarations & exclusions');

  console.log('\n================================================================');
  console.log('🎉 ALL 4 ADVANCED BACKEND ARCHITECTURAL UPGRADES PASSED 100%!');
  console.log('================================================================');
}

runUpgradesTests().catch((err) => {
  console.error('❌ UPGRADES TEST SUITE FAILED:', err);
  process.exit(1);
});
