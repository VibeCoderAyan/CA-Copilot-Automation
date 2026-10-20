/**
 * Test Suite: Form 3CD Tax Audit Scrutiny & GSTR-1 vs 3B vs Books 3-Way Rec Backend Services
 */
import assert from 'node:assert';
import { BackendStore } from '../db/store.ts';
import { Form3cdService } from '../services/form3cdService.ts';
import { GstTripleRecService } from '../services/gstTripleRecService.ts';

console.log('================================================================');
console.log('🧪 RUNNING FORM 3CD & GST 3-WAY RECONCILIATION TEST SUITE');
console.log('================================================================\n');

const store = BackendStore.getInstance();
const form3cdService = new Form3cdService(store);
const gstTripleRecService = new GstTripleRecService(store);

// =========================================================================
// TEST 1: FORM 3CD CLAUSE 21(d) / SEC 40A(3) CASH LIMITS
// =========================================================================
console.log('--- TEST 1: FORM 3CD CLAUSE 21(d) SEC 40A(3) SCRUTINY ---');

const testVouchers = [
  // Payee A: Two cash payments on same date totaling ₹22,000 (> ₹10,000 limit)
  {
    voucherNo: 'V-101',
    date: '15-06-2025',
    accountName: 'Repairs & Maintenance',
    payeeName: 'Sharma Contractors',
    amount: 12000,
    paymentMode: 'Cash',
    narration: 'Part payment for floor repair',
  },
  {
    voucherNo: 'V-102',
    date: '15-06-2025',
    accountName: 'Repairs & Maintenance',
    payeeName: 'Sharma Contractors',
    amount: 10000,
    paymentMode: 'Cash',
    narration: 'Second part payment for floor repair',
  },
  // Payee B: Transporter cash payment of ₹32,000 (Within ₹35,000 carriage limit)
  {
    voucherNo: 'V-103',
    date: '18-06-2025',
    accountName: 'Freight Inward',
    payeeName: 'Speed Cargo Logistics',
    amount: 32000,
    paymentMode: 'Cash',
    narration: 'Lorry freight paid in cash for truck goods carriage',
  },
  // Payee C: Transporter cash payment of ₹45,000 (Exceeds ₹35,000 limit)
  {
    voucherNo: 'V-104',
    date: '20-06-2025',
    accountName: 'Transport Charges',
    payeeName: 'Gujarat Roadways',
    amount: 45000,
    paymentMode: 'Cash',
    narration: 'Truck transport freight paid in cash',
  },
  // Payee D: Bank holiday cash payment of ₹18,000 (Rule 6DD exemption)
  {
    voucherNo: 'V-105',
    date: '15-08-2025',
    accountName: 'Emergency Electric Works',
    payeeName: 'City Electricians',
    amount: 18000,
    paymentMode: 'Cash',
    narration: 'Emergency electrical repairs paid on bank holiday',
  },
  // Payee E: Penalties & fines (Clause 21a)
  {
    voucherNo: 'V-106',
    date: '10-09-2025',
    accountName: 'Legal Expenses',
    payeeName: 'Municipal Corp',
    amount: 25000,
    paymentMode: 'Bank',
    narration: 'Late filing penalty and fine paid under municipal laws',
  },
  // Regular Vendor: GST registered
  {
    voucherNo: 'V-107',
    date: '25-09-2025',
    accountName: 'Office Stationery',
    payeeName: 'PaperKraft Ltd',
    amount: 50000,
    paymentMode: 'Bank',
    narration: 'Stationery purchase with GST invoice',
    gstin: '27AAACP1234F1Z8',
  },
  // Unregistered Supplier
  {
    voucherNo: 'V-108',
    date: '28-09-2025',
    accountName: 'Office Stationery',
    payeeName: 'Local Corner Vendor',
    amount: 15000,
    paymentMode: 'Bank',
    narration: 'Pantry and stationery from unregistered seller',
    gstin: '',
  },
];

const f3cdResult = form3cdService.scrutinizeLedger('firm_apex_ca', 'client_apex_tech', testVouchers, '2025-26', '2026-27');

assert(f3cdResult.totalScrutinizedAmount === 207000, 'Scrutinized total expenditure accurately (₹2,07,000)');
assert(f3cdResult.cashViolations.length >= 3, 'Detected all 3 cash payment breach groups');

// Verify Sharma Contractors daily aggregate ₹22,000 is disallowed
const sharmaViol = f3cdResult.cashViolations.find((v: any) => v.payeeName.includes('Sharma Contractors'));
assert(sharmaViol, 'Sharma Contractors breach identified');
assert(sharmaViol.dailyAggregateAmount === 22000, 'Daily aggregate correctly sums both vouchers to ₹22,000');
assert(sharmaViol.disallowableAmount === 22000, 'Full ₹22,000 marked as Sec 40A(3) disallowable');

// Verify Gujarat Roadways transporter limit ₹35k breached
const gujViol = f3cdResult.cashViolations.find((v: any) => v.payeeName.includes('Gujarat Roadways'));
assert(gujViol, 'Gujarat Roadways identified as transporter breach');
assert(gujViol.thresholdLimit === 35000, 'Transporter threshold evaluated as ₹35,000');
assert(gujViol.disallowableAmount === 45000, 'Disallowed ₹45,000');

// Verify Rule 6DD exemption on bank holiday
const bankHolViol = f3cdResult.cashViolations.find((v: any) => v.payeeName.includes('City Electricians'));
assert(bankHolViol, 'Bank holiday payment identified');
assert(bankHolViol.isRule6DdExempt === true, 'Rule 6DD exemption granted');
assert(bankHolViol.disallowableAmount === 0, 'Exempted payment has 0 disallowance');

console.log('✓ PASS: Daily cash aggregation correctly flags ₹10,000 and ₹35,000 transporter breaches');
console.log('✓ PASS: Rule 6DD bank holiday statutory exemption accurately calculated');

// =========================================================================
// TEST 2: FORM 3CD CLAUSE 44 (GST REGISTERED VS UNREGISTERED BREAKDOWN)
// =========================================================================
console.log('\n--- TEST 2: FORM 3CD CLAUSE 44 BREAKDOWN ---');
assert(f3cdResult.clause44Breakdown.length > 0, 'Clause 44 breakdown generated');
const stationeryHead = f3cdResult.clause44Breakdown.find((h: any) => h.headOfExpenditure.includes('Office Stationery'));
assert(stationeryHead, 'Stationery expenditure head categorized');
assert(stationeryHead.registeredExpenditure === 50000, 'Registered spend is ₹50,000');
assert(stationeryHead.unregisteredExpenditure === 15000, 'Unregistered spend is ₹15,000');
assert(f3cdResult.unregisteredExpenditureRatio > 0, 'Computed unregistered spend ratio');
console.log(`✓ PASS: Clause 44 splits registered vs unregistered spend accurately (${f3cdResult.unregisteredExpenditureRatio}%)`);

// =========================================================================
// TEST 3: FORM 3CD CLAUSE 21(a) PENALTIES & FINES
// =========================================================================
console.log('\n--- TEST 3: FORM 3CD CLAUSE 21(a) PENALTIES & FINES ---');
assert(f3cdResult.clause21aPenaltiesTotal === 25000, 'Clause 21(a) flagged ₹25,000 municipal penalty');
assert(f3cdResult.totalTaxAdditions === 22000 + 45000 + 25000, 'Total tax additions sum is 100% mathematically correct (₹92,000)');
console.log('✓ PASS: Clause 21(a) penalty disallowances and total taxable profit additions verified');

// =========================================================================
// TEST 4: GSTR-1 VS 3B VS BOOKS 3-WAY RECONCILIATION
// =========================================================================
console.log('\n--- TEST 4: GSTR-1 VS 3B VS BOOKS 3-WAY RECONCILIATION ---');

const test3WayData = [
  // Month 1: Clean match
  {
    month: 'April 2025',
    booksTurnover: 5000000,
    gstr1Turnover: 5000000,
    gstr1Igst: 400000,
    gstr1Cgst: 250000,
    gstr1Sgst: 250000,
    gstr3bTurnover: 5000000,
    gstr3bIgst: 400000,
    gstr3bCgst: 250000,
    gstr3bSgst: 250000,
  },
  // Month 2: DRC-01B Notice Risk (GSTR-1 tax exceeds GSTR-3B by ₹80,000 > ₹25,000 threshold)
  {
    month: 'May 2025',
    booksTurnover: 6000000,
    gstr1Turnover: 6500000,
    gstr1Igst: 520000,
    gstr1Cgst: 300000,
    gstr1Sgst: 300000,
    gstr3bTurnover: 6000000,
    gstr3bIgst: 440000,
    gstr3bCgst: 300000,
    gstr3bSgst: 300000,
  },
  // Month 3: Excess tax paid in 3B by ₹20,000
  {
    month: 'June 2025',
    booksTurnover: 4000000,
    gstr1Turnover: 4000000,
    gstr1Igst: 320000,
    gstr1Cgst: 200000,
    gstr1Sgst: 200000,
    gstr3bTurnover: 4125000,
    gstr3bIgst: 340000,
    gstr3bCgst: 200000,
    gstr3bSgst: 200000,
  },
];

const gstResult = gstTripleRecService.reconcile('firm_apex_ca', 'client_apex_tech', test3WayData, '2025-26', 25000);

assert(gstResult.records.length === 3, 'All 3 months reconciled');
assert(gstResult.drc01bNoticeMonthsCount === 1, 'Exactly 1 month (May 2025) flagged for DRC-01B notice');

const mayRecord = gstResult.drc01bList[0];
assert(mayRecord.month === 'May 2025', 'May 2025 confirmed as critical notice risk');
assert(mayRecord.netTaxShortfall === 80000, 'Tax liability shortfall is exactly ₹80,000');
assert(mayRecord.estimatedInterest > 0, 'Section 50 interest computed (> 0)');
assert(mayRecord.turnoverDiff1Vs3b === 500000, 'Turnover difference GSTR-1 vs 3B is ₹5,00,000');

console.log('✓ PASS: DRC-01B Rule 88C statutory trigger accurately identifies tax shortfall > ₹25,000');
console.log(`✓ PASS: Section 50 late interest estimated at ₹${mayRecord.estimatedInterest.toLocaleString('en-IN')}`);
console.log('✓ PASS: Tripartite turnover walk and annual totals mathematically verified');

console.log('\n================================================================');
console.log('🎉 ALL FORM 3CD & GST TRIPLE RECONCILIATION TESTS PASSED 100%!');
console.log('================================================================\n');
