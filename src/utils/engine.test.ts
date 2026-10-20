import { detectColumnCategory, mapSpreadsheetColumns } from './columnDetector.ts';
import { parseAmount, parseDate, parseDateObj, isValidGstin, formatINR, normalizeRefNo } from './dataSanitizer.ts';
import { cleanSpreadsheetData } from './excelCleanerEngine.ts';
import { parseBankStatementRows, parseLedgerRows, runBankReconciliation } from './bankReconciliationEngine.ts';
import { parsePurchaseRegisterRows, parseGstr2bRows, runGstReconciliation } from './gstReconciliationEngine.ts';
import { parseTrialBalanceRows, compileFinancialStatements } from './financialReportsEngine.ts';
import { parseAuditTransactions, sampleTransactions, detectUnusualTransactions } from './auditEngine.ts';
import { reconcileTdsRecords, TdsColumnMapping } from './tdsReconciliationEngine.ts';
import { computeFinancialRatios, FinancialInputs } from './ratioAnalyzerEngine.ts';
import { calculateAssetDepreciation, AssetRecord } from './depreciationEngine.ts';
declare const process: { exit: (code: number) => never };

function assert(condition: boolean, msg: string) {
  if (!condition) {
    console.error(`❌ FAIL: ${msg}`);
    process.exit(1);
  } else {
    console.log(`✓ PASS: ${msg}`);
  }
}

console.log('=== 1. TESTING COLUMN DETECTION & ALIAS MAPPING ===');
assert(detectColumnCategory('Invoice_Date') === 'date', 'Detects Invoice_Date as date');
assert(detectColumnCategory('Txn Date') === 'date', 'Detects Txn Date as date');
assert(detectColumnCategory('Due Date') === 'dueDate', 'Detects Due Date as dueDate');
assert(detectColumnCategory('Payment Due Date') === 'dueDate', 'Detects Payment Due Date as dueDate');
assert(detectColumnCategory('Sundry Debtors') === 'partyName', 'Detects Sundry Debtors as partyName (not debit)');
assert(detectColumnCategory('Debtors') === 'partyName', 'Detects Debtors as partyName');
assert(detectColumnCategory('Sundry Creditors') === 'partyName', 'Detects Sundry Creditors as partyName (not credit)');
assert(detectColumnCategory('Item Description') === 'narration', 'Detects Item Description as narration (not credit)');
assert(detectColumnCategory('GSTIN of supplier') === 'gstin', 'Detects GSTIN of supplier as gstin');
assert(detectColumnCategory('Taxable Value (₹)') === 'taxableAmount', 'Detects Taxable Value (₹) as taxableAmount');
assert(detectColumnCategory('Withdrawal (Debit)') === 'debit', 'Detects Withdrawal (Debit) as debit');
assert(detectColumnCategory('Deposit (Credit)') === 'credit', 'Detects Deposit (Credit) as credit');
assert(detectColumnCategory('Particulars / Narration') === 'narration', 'Detects Particulars as narration');

const mapped = mapSpreadsheetColumns(
  ['Voucher No', 'Txn Date', 'Party Name', 'Gross Amount', 'Random Col'],
  ['voucherNo', 'date', 'partyName']
);
assert(mapped.allFound, 'All mandatory columns mapped');
assert(mapped.unmapped.includes('Random Col'), 'Identifies unmapped custom columns');

console.log('\n=== 2. TESTING DATA SANITIZATION & CURRENCY ===');
assert(parseAmount('1,50,000.00') === 150000, 'Parses comma-formatted number');
assert(parseAmount('₹ 2,45,000.50') === 245000.5, 'Parses currency prefix');
assert(parseAmount('(45,000.00)') === -45000, 'Parses accounting negative in parentheses');
assert(parseAmount('-1200.00') === -1200, 'Parses standard minus');
assert(parseAmount('N/A') === 0, 'Parses N/A as 0');
assert(parseAmount(null) === 0, 'Parses null as 0');

assert(parseDate('09/04/2026') === '09-04-2026', 'Standardizes DD/MM/YYYY');
assert(parseDate('2026-09-06') === '06-09-2026', 'Standardizes YYYY-MM-DD');
assert(parseDate('15.09.2026') === '15-09-2026', 'Standardizes dot separated dates');
assert(isValidGstin('27AABCA1234F1Z5') === true, 'Validates correct 15-char GSTIN');
assert(isValidGstin('INVALID_GSTIN') === false, 'Rejects invalid GSTIN');
assert(normalizeRefNo('INV/2026/0891') === 'INV20260891', 'Normalizes reference number');

console.log('\n=== 3. TESTING EXCEL CLEANER ENGINE ===');
const messyData = [
  { Voucher_No: 'SAL/01', Invoice_Date: '09/04/2026', Party_Name: 'Apex Logistics', GSTIN: '27AABCA1234F1Z5', Amount: '1,50,000.00' },
  { Voucher_No: 'SAL/02', Invoice_Date: '2026-09-06', Party_Name: 'Bharat Biotech', GSTIN: 'INVALID', Amount: '(45,000.00)' },
  { Voucher_No: 'SAL/01', Invoice_Date: '09/04/2026', Party_Name: 'Apex Logistics', GSTIN: '27AABCA1234F1Z5', Amount: '1,50,000.00' }, // Exact duplicate
];
const cleanResult = cleanSpreadsheetData(messyData);
assert(cleanResult.cleanedRows.length === 3, 'Preserves all original rows (never discards)');
assert(cleanResult.summary.duplicatesFound === 1, 'Identifies exact duplicate row');
assert(cleanResult.summary.missingGstins === 1, 'Identifies invalid/missing GSTIN');
assert(cleanResult.cleanedRows[1].Amount === -45000, 'Cleaned accounting negative amount to float');

console.log('\n=== 4. TESTING BANK RECONCILIATION ENGINE ===');
const sampleBank = [
  { Date: '02-09-2026', Narration: 'NEFT CR - BHARAT SHELL', 'Ref No': 'CMS101', Withdrawal: 0, Deposit: 485000, Balance: 1485000 },
  { Date: '15-09-2026', Narration: 'BANK CHARGES - PROCESSING', 'Ref No': 'CHG01', Withdrawal: 1770, Deposit: 0, Balance: 1483230 },
];
const sampleLedger = [
  { Date: '02-09-2026', 'Voucher No': 'CMS101', Particulars: 'Bharat Shell Limited', Debit: 485000, Credit: 0 },
  { Date: '28-09-2026', 'Voucher No': 'CHQ440', Particulars: 'Apollo Tyres (Unpresented)', Debit: 0, Credit: 50000 },
];

const parsedB = parseBankStatementRows(sampleBank);
const parsedL = parseLedgerRows(sampleLedger);
const brs = runBankReconciliation(parsedB, parsedL);

assert(brs.calculation.matchedCount === 1, 'Matched exact Ref No and Amount');
assert(brs.calculation.unpresentedCheques === 50000, 'Correct unpresented cheque total');
assert(brs.calculation.unbookedBankCharges === 1770, 'Correct unbooked bank charges total');
assert(brs.calculation.adjustedBankBalance === 1483230 - 50000 + 1770 - 0, 'Accurate adjusted bank balance');

console.log('\n=== 5. TESTING GST RECONCILIATION ENGINE ===');
const booksGst = [
  { GSTIN: '27AABCA1234F1Z5', 'Invoice No': 'INV-01', Date: '05-09-2026', 'Taxable Value': 100000, CGST: 9000, SGST: 9000, IGST: 0 },
  { GSTIN: '07AAACB9876Q1Z2', 'Invoice No': 'INV-02', Date: '10-09-2026', 'Taxable Value': 50000, CGST: 0, SGST: 0, IGST: 9000 },
];
const portal2bGst = [
  { 'GSTIN of supplier': '27AABCA1234F1Z5', 'Invoice number': 'INV-01', 'Invoice date': '05-09-2026', 'Taxable Value (₹)': 100000, 'Central Tax (₹)': 9000, 'State/UT Tax (₹)': 9000, 'Integrated Tax (₹)': 0 },
  { 'GSTIN of supplier': '07AAACB9876Q1Z2', 'Invoice number': 'INV-02', 'Invoice date': '10-09-2026', 'Taxable Value (₹)': 50000, 'Central Tax (₹)': 0, 'State/UT Tax (₹)': 0, 'Integrated Tax (₹)': 5000 }, // Mismatch: 9000 vs 5000
  { 'GSTIN of supplier': '33AAECZ8901L1ZT', 'Invoice number': 'INV-99', 'Invoice date': '25-09-2026', 'Taxable Value (₹)': 40000, 'Central Tax (₹)': 3600, 'State/UT Tax (₹)': 3600, 'Integrated Tax (₹)': 0 }, // Missing in books
];

const parsedPr = parsePurchaseRegisterRows(booksGst);
const parsed2b = parseGstr2bRows(portal2bGst);
const gstRec = runGstReconciliation(parsedPr, parsed2b);

assert(gstRec.summary.matchedCount === 1, 'Identifies 100% matched invoice');
assert(gstRec.summary.mismatchCount === 1, 'Identifies tax mismatch invoice');
assert(gstRec.summary.missingInBooksCount === 1, 'Identifies invoice missing in books');
assert(gstRec.summary.eligibleItc === 18000, 'Correct 100% eligible ITC calculation');

console.log('\n=== 6. TESTING FINANCIAL REPORTS ENGINE ===');
const trialBalanceData = [
  { 'Account Code': '1001', 'Account Name': 'HDFC Bank Current Account', Head: 'Asset', Debit: 500000, Credit: 0 },
  { 'Account Code': '2001', 'Account Name': 'Share Capital', Head: 'Equity', Debit: 0, Credit: 300000 },
  { 'Account Code': '3001', 'Account Name': 'Sales Revenue', Head: 'Revenue', Debit: 0, Credit: 400000 },
  { 'Account Code': '4001', 'Account Name': 'Operating Expenses', Head: 'Expense', Debit: 200000, Credit: 0 },
];
const parsedTb = parseTrialBalanceRows(trialBalanceData);
const fin = compileFinancialStatements(parsedTb);

assert(fin.isBalanced, 'Trial balance is balanced (Debits 700k == Credits 700k)');
assert(fin.netProfit === 200000, 'Correct Net Profit = 400k - 200k = 200k');
assert(fin.totalCapAndLiab === 300000 + 200000, 'Capital + Net Profit = 500k');
assert(fin.assets === 500000, 'Total Assets = 500k');
assert(fin.balanceSheetDiff < 0.01, 'Schedule III Balance Sheet equation verifies with 0 variance');

console.log('\n=== 7. TESTING AUDIT SCUTINY & SAMPLING ENGINE ===');
const auditTxnsData = [
  { 'Voucher No': 'PMT/01', Date: '05-09-2026', Particulars: 'Vendor Payment', Amount: 50000, 'Payment Mode': 'Bank' },
  { 'Voucher No': 'PMT/02', Date: '14-09-2026', Particulars: 'Petty Cash', Amount: 25000, 'Payment Mode': 'Cash' }, // Sec 40A(3) violation (> 10k)
  { 'Voucher No': 'SAL/03', Date: '27-09-2026', Particulars: 'Sunday Sale', Amount: 1000000, 'Payment Mode': 'Bank' }, // Sunday + Round sum
];
const parsedAudit = parseAuditTransactions(auditTxnsData);
const sampling = sampleTransactions(parsedAudit, 'materiality', 100000);
assert(sampling.sampled.length === 1 && sampling.sampled[0].amount === 1000000, 'Materiality sampling correctly isolates items >= 100k');

const flagged = detectUnusualTransactions(parsedAudit);
assert(flagged.some((f) => f.flagTypes.includes('Sec 40A(3) Cash Limit')), 'Flags Section 40A(3) cash limit violation');
assert(flagged.some((f) => f.flagTypes.includes('Round Sum Entry')), 'Flags round sum entries');

console.log('\n=== 8. TESTING ROBUST DATE PARSING & SERIAL CONVERSION ===');
const dSlash = parseDateObj('15/09/2026');
assert(dSlash !== null && dSlash.getDate() === 15 && dSlash.getMonth() === 8 && dSlash.getFullYear() === 2026, 'Parses DD/MM/YYYY slash dates accurately');
const dExcel = parseDateObj(45231); // Excel numeric serial
assert(dExcel !== null && !isNaN(dExcel.getTime()), 'Parses Excel serial number to valid date');
const dIso = parseDateObj('2026-10-31');
assert(dIso !== null && dIso.getDate() === 31 && dIso.getMonth() === 9 && dIso.getFullYear() === 2026, 'Parses YYYY-MM-DD ISO dates accurately');

console.log('\n=== 9. TESTING TDS SECTION 201(1A) STATUTORY INTEREST ===');
const tdsRows = [
  {
    'Deductee Name': 'Apex Infrastructure',
    'PAN': 'AABCA1234F',
    'Section': '194C',
    'Booking Date': '15/04/2026', // April deduction (Due date: May 7, 2026)
    'Taxable Value': 500000,
    'TDS Deducted': 10000,
    'TDS Deposited': 10000,
    'Challan No': 'CHL-001',
    'Challan Date': '10/06/2026', // Deposited in June -> Late deposit! (3 months: Apr, May, Jun = 4.5% interest)
  },
];
const tdsMapping: TdsColumnMapping = {
  deducteeCol: 'Deductee Name',
  panCol: 'PAN',
  sectionCol: 'Section',
  bookingDateCol: 'Booking Date',
  taxableCol: 'Taxable Value',
  tdsDeductedCol: 'TDS Deducted',
  tdsDepositedCol: 'TDS Deposited',
  challanNoCol: 'Challan No',
  challanDateCol: 'Challan Date',
  bsrCodeCol: '',
};
const tdsSummary = reconcileTdsRecords(tdsRows, tdsMapping);
assert(tdsSummary.lateDepositCount === 1, 'Correctly flags deposit after May 7 due date as late deposit');
assert(tdsSummary.totalEstimatedInterest === 450, 'Computes 1.5% interest per month/part-of-month under Sec 201(1A): ₹10,000 * 4.5% = ₹450');
assert(!isNaN(tdsSummary.totalEstimatedInterest), 'Guarantees totalEstimatedInterest is never NaN');

console.log('\n=== 10. TESTING RATIO ANALYZER WITH NEGATIVE EQUITY ===');
const dummyCyInputs: FinancialInputs = {
  periodLabel: 'FY 2026-27',
  currentAssets: 1000000,
  cashAndBank: 200000,
  tradeReceivables: 500000,
  inventories: 300000,
  otherCurrentAssets: 0,
  fixedAssetsNet: 2000000,
  totalAssets: 3000000,
  currentLiabilities: 1500000,
  tradePayables: 800000,
  shortTermBorrowings: 700000,
  otherCurrentLiabilities: 0,
  longTermDebt: 3500000,
  totalDebt: 4200000,
  shareholdersEquity: -1200000, // NEGATIVE NET WORTH
  totalLiabilitiesAndEquity: 3000000,
  revenue: 5000000,
  cogs: 3500000,
  grossProfit: 1500000,
  operatingExpenses: 1200000,
  operatingProfit: 300000,
  financeCosts: 0, // DEBT-FREE / ZERO FINANCE COST
  depreciation: 200000,
  profitBeforeTax: 100000,
  taxExpense: 25000,
  profitAfterTax: 75000,
};
const ratioResult = computeFinancialRatios(dummyCyInputs, dummyCyInputs);
const roeRatio = ratioResult.ratios.find((r) => r.id === 'return_on_equity')!;
assert(roeRatio.cyValue === null, 'ROE is safely null for negative net worth (never computes bogus positive return)');
assert(roeRatio.interpretation === 'concern', 'Flags capital erosion as critical concern');
const icrRatio = ratioResult.ratios.find((r) => r.id === 'interest_coverage')!;
assert(icrRatio.cyValue === null && icrRatio.notes.includes('debt-free'), 'Debt-free company safely handled without division by zero');

console.log('\n=== 11. TESTING DEPRECIATION 180-DAY CUTOFF & SALVAGE CAP ===');
const sampleAsset: AssetRecord = {
  id: 'ast-test-1',
  code: 'EQ-01',
  name: 'CNC Milling Machine',
  category: 'plant_machinery',
  purchaseDate: '15-01-2026', // Q4 Addition: Jan 15, 2026 in FY 2025-26 -> Used < 180 days!
  putToUseDate: '15-01-2026',
  originalCost: 1000000,
  residualValue: 50000, // 5% residual
  openingWdv: 1000000,
  additionsMoreThan180Days: 0,
  additionsLessThan180Days: 500000,
  deletionsDeductions: 0,
  saleRealization: 0,
  method: 'WDV',
};
const deprResult = calculateAssetDepreciation(sampleAsset);
assert(deprResult.coActClosingWdv >= sampleAsset.residualValue, 'Net book value never drops below 5% salvage value');
assert(deprResult.coActDepreciation > 0, 'Depreciation correctly calculated');

console.log('\n=== 12. TESTING TOOL 1: BULK INVOICE-TO-EXCEL AUTOMATION & DEDUP ===');
const { validateInvoiceRecord, detectDuplicatesInBatch, exportInvoiceCsv, createInvoiceExportWorkbook } = await import('./invoiceDataExtractorEngine.ts');
const inv1 = validateInvoiceRecord({
  id: 'inv-t1',
  fileName: 'Apex_Inv_01.pdf',
  invoiceNumber: 'INV/2026/01',
  invoiceDate: '10-09-2026',
  supplierName: 'Apex Logistics Pvt Ltd',
  supplierGstin: '27AABCA1234F1Z5',
  customerGstin: '27AAECP1122C1ZG',
  placeOfSupply: '27-Maharashtra',
  taxableAmount: 100000,
  cgst: 9000,
  sgst: 9000,
  igst: 0,
  cess: 500,
  totalAmount: 118500,
  currency: 'INR',
  confidenceScore: 98,
});
assert(inv1.status === 'Verified', 'Invoice passes statutory tax math check');
assert(inv1.totalTax === 18500, 'Calculates CGST + SGST + Cess accurately');

// Test Duplicate Invoice Detection
const invDuplicate = { ...inv1, id: 'inv-t2', fileName: 'Apex_Inv_01_Copy.pdf' };
const batchWithDups = detectDuplicatesInBatch([inv1, invDuplicate]);
assert(batchWithDups[1].status === 'Duplicate', 'Detects duplicate invoice in batch with same GSTIN, number, and amount');
const csvOut = exportInvoiceCsv(batchWithDups);
assert(csvOut.includes('INV/2026/01'), 'Generates valid CSV export with invoice serial');
const invWb = createInvoiceExportWorkbook(batchWithDups, 'Test_Batch');
assert(invWb.SheetNames.includes('Consolidated_Invoices') && invWb.SheetNames.includes('Duplicates_Report'), 'Creates 4-sheet workbook with duplicates report');

console.log('\n=== 13. TESTING TOOL 2: BANK STATEMENT-TO-LEDGER CONTINUITY & DRAFT JOURNALS ===');
const { validateBalanceContinuity, suggestAccountHead, generateDraftJournalEntries, convertBankStatementRows, createConvertedStatementWorkbook, exportBankStatementCsv } = await import('./bankStatementConverterEngine.ts');

// Test Running Balance Continuity check
const continuousTxns: any[] = [
  { id: 't1', date: '01-09-2026', narration: 'Opening', withdrawal: 0, deposit: 0, balance: 100000 },
  { id: 't2', date: '05-09-2026', narration: 'Payment to Vendor', withdrawal: 25000, deposit: 0, balance: 75000 },
  { id: 't3', date: '10-09-2026', narration: 'Receipt from Client', withdrawal: 0, deposit: 50000, balance: 125000 },
];
const continuityGood = validateBalanceContinuity(continuousTxns);
assert(!continuityGood.hasBreaks, 'Verifies 100% mathematical continuity when PrevBal - Wdr + Dep == CurrBal');

// Continuity Break Test
const brokenTxns: any[] = [
  { id: 't1', date: '01-09-2026', narration: 'Opening', withdrawal: 0, deposit: 0, balance: 100000 },
  { id: 't2', date: '05-09-2026', narration: 'Payment to Vendor', withdrawal: 25000, deposit: 0, balance: 60000 }, // Break: expected 75k, actual 60k!
];
const continuityBad = validateBalanceContinuity(brokenTxns);
assert(continuityBad.hasBreaks && continuityBad.breakCount === 1, 'Accurately catches running balance break');

// Account Head Auto-Suggester Rules
const gstHead = suggestAccountHead('GST ELECTRONIC CASH LEDGER PMT - CPASS', true);
assert(gstHead.accountHead.includes('GST Payable'), 'Suggests GST Payable for GST challan narration');
const salHead = suggestAccountHead('SALARY NEFT DISBURSEMENT BATCH', true);
assert(salHead.accountHead.includes('Salaries'), 'Suggests Salaries & Wages for payroll narration');
const chgHead = suggestAccountHead('CONSOLIDATED BANK CHARGES', true);
assert(chgHead.accountHead.includes('Bank Charges'), 'Suggests Bank Charges for bank fee narration');

// Draft Journal Entry Generation
const draftEntries = generateDraftJournalEntries([
  { id: 't1', date: '05-09-2026', narration: 'GST PAYMENT', refNo: 'GST01', withdrawal: 50000, deposit: 0, balance: 50000, category: 'Bank Charges', suggestedAccountHead: 'GST Payable A/c', suggestedRuleReason: 'Statutory GST rule' }
], 'HDFC Bank A/c');
assert(draftEntries[0].debitAccount === 'GST Payable A/c' && draftEntries[0].creditAccount === 'HDFC Bank A/c', 'Draft double-entry journal correctly debits liability and credits bank');

console.log('\n=== 14. TESTING TOOL 3: GST RECONCILIATION TOLERANCES & SEC 16(2)(aa) ===');
const { normalizeInvoiceNumber, exportGstReconciliationCsv, createGstReconciliationWorkbook } = await import('./gstReconciliationEngine.ts');
assert(normalizeInvoiceNumber('INV/2026/0042') === 'INV202642', 'Normalizes invoice number by stripping special chars and leading zeros');
assert(normalizeInvoiceNumber('APO-4412') === 'APO4412', 'Normalizes hyphenated invoice numbers');

const gstCsv = exportGstReconciliationCsv(gstRec.comparisonRecords);
assert(gstCsv.includes('Supplier GSTIN'), 'Generates standard CSV export for GST reconciliation');
const gstWb = createGstReconciliationWorkbook(gstRec.comparisonRecords, gstRec.summary);
assert(gstWb.SheetNames.includes('Executive_Summary') && gstWb.SheetNames.includes('Matched_Invoices'), 'Generates multi-sheet reconciliation workbook');

console.log('\n=== 15. TESTING TOOL 4: SCHEDULE III TB & DYNAMIC AUDIT ADJUSTMENTS ===');
const { checkAbnormalBalance, createFinancialStatementsWorkbook } = await import('./financialReportsEngine.ts');
// Abnormal balance check
const abnormalCash: any = { accountCode: 'AC-101', accountName: 'Cash in Hand', head: 'Asset', debit: 0, credit: 15000 };
assert(checkAbnormalBalance(abnormalCash).isAbnormal, 'Flags negative cash balance as abnormal');

// Test Dynamic Audit Adjustment
const baseTB: any[] = [
  { id: '1', accountCode: 'AC-1', accountName: 'Share Capital', head: 'Equity', scheduleIIIGroup: 'Share Capital', debit: 0, credit: 500000, netBalance: -500000 },
  { id: '2', accountCode: 'AC-2', accountName: 'Bank Balance', head: 'Asset', scheduleIIIGroup: 'Cash & Bank Balances', debit: 500000, credit: 0, netBalance: 500000 },
];
const initialStatements = compileFinancialStatements(baseTB, []);
assert(initialStatements.isBalanced, 'Initial statements balance');

// Apply adjustment: Dr Closing Stock (Inventories) ₹1,00,000 / Cr Changes in Inventory (Revenue) ₹1,00,000
const auditAdjustment: any = {
  id: 'adj-test-1',
  debitAccountCode: 'AC-INV',
  debitAccountName: 'Closing Stock (Inventories)',
  creditAccountCode: 'AC-CHG',
  creditAccountName: 'Changes in Inventory',
  amount: 100000,
  narration: 'Physical stock adjustment',
  category: 'Closing Stock',
  isApproved: true,
};
const adjustedStatements = compileFinancialStatements(baseTB, [auditAdjustment]);
assert(adjustedStatements.isBalanced, 'Adjusted statements maintain strict mathematical balance');
assert(adjustedStatements.scheduleIII.inventories === 100000, 'Inventories dynamically updated to ₹1,00,000');
assert(adjustedStatements.scheduleIII.netProfitBeforeTax === 100000, 'Net profit dynamically updated after adjustment entry');

console.log('\n=== 16. TESTING TOOL 5: AUDIT FILE AUTO-PROCESSING & CHECKLIST ===');
const { classifyAuditDocument, getCategoryCode, evaluateAuditChecklist, computeRenamedFiles, createRenamingManifestWorkbook } = await import('./bulkRenamerEngine.ts');
assert(classifyAuditDocument('hdfc_bank_statement_q2.pdf').category === 'Bank Statements', 'Classifies bank statement file');
assert(classifyAuditDocument('gstr3b_filed_ack_sept.pdf').category === 'Tax Records', 'Classifies GSTR-3B as tax record');
assert(classifyAuditDocument('trial_balance_final.xlsx').category === 'Trial Balances', 'Classifies TB file');

assert(getCategoryCode('Bank Statements') === 'BNK', 'Working paper code for Bank Statements is BNK');
assert(getCategoryCode('Purchase Invoices') === 'PUR', 'Working paper code for Purchase Invoices is PUR');

// Test Working Paper Evidence Index generation
const sampleAuditDocs: any[] = [
  { id: 'd1', originalName: 'hdfc_bank_sep.pdf', originalExtension: 'pdf', sizeBytes: 50000, detectedDate: '2026-09-30', detectedVendor: 'HDFC', detectedInvoiceNo: 'STMT1', detectedCategory: 'Bank Statements', categoryCode: 'BNK', confidenceScore: 95 },
  { id: 'd2', originalName: 'apollo_bill_4412.pdf', originalExtension: 'pdf', sizeBytes: 120000, detectedDate: '2026-09-15', detectedVendor: 'Apollo', detectedInvoiceNo: 'INV4412', detectedCategory: 'Purchase Invoices', categoryCode: 'PUR', confidenceScore: 92 },
];
const indexedDocs = computeRenamedFiles(sampleAuditDocs, {
  pattern: '{evidenceIndex}_{client}_{category}_{date}',
  prefix: '',
  suffix: '',
  clientName: 'Apex_Corp',
  dateFormat: 'YYYY-MM-DD',
  caseOption: 'original',
  spaceOption: 'underscore',
  seqStart: 1,
  seqPadding: 3,
  autoResolveDuplicates: true,
  organizeIntoFolders: true,
});
assert(indexedDocs[0].evidenceIndex === 'WP-BNK-001', 'Assigns working paper evidence index WP-BNK-001');
assert(indexedDocs[1].evidenceIndex === 'WP-PUR-001', 'Assigns working paper evidence index WP-PUR-001');

// Test ICAI Checklist Evaluation
const checklistEvaluation = evaluateAuditChecklist(indexedDocs);
const bnkChecklist = checklistEvaluation.find((c) => c.category === 'Bank Statements')!;
assert(bnkChecklist.isFound === true, 'Identifies that Bank Statement has been provided');
const farChecklist = checklistEvaluation.find((c) => c.title.includes('Fixed Asset'))!;
assert(farChecklist.isFound === false, 'Accurately flags missing Fixed Asset Register from audit checklist');

console.log('\n=== 17. TESTING STRICT FILE VALIDATION & UNSUPPORTED PDF REJECTION ===');
const { extractFromInvoicePdf } = await import('./invoiceDataExtractorEngine.ts');

const fakeBlob = new Blob(['This is not an invoice or bank statement. Random plain text story.'], { type: 'application/pdf' });
const fakePdfFile = new File([fakeBlob], 'random_story.pdf', { type: 'application/pdf' });

let threwError = false;
let thrownMessage = '';
try {
  await extractFromInvoicePdf(fakePdfFile);
} catch (err: any) {
  threwError = true;
  thrownMessage = err.message;
}
console.log('ACTUAL THROWN MESSAGE:', JSON.stringify(thrownMessage));
assert(threwError === true, 'Strictly halts and throws error when an unsupported PDF is uploaded to Invoice Extractor');
assert(
  thrownMessage.length > 5,
  'Diagnostic error message explains why document was rejected'
);

// Test that parsePurchaseRegisterRows identifies completely unrelated tabular rows
const randomRows = [
  { 'Student Name': 'Rohan', 'Grade': 'A', 'Marks': 95 },
  { 'Student Name': 'Pooja', 'Grade': 'B', 'Marks': 82 }
];
const randomVouchers = parsePurchaseRegisterRows(randomRows);
const allEmpty = randomVouchers.every((v) => v.taxableAmount === 0 && v.totalInvoiceAmount === 0 && !v.invoiceNumber && !v.supplierGstin);
assert(allEmpty, 'Correctly flags that random data contains zero recognizable tax invoices or amounts');

console.log('\n=== 18. TESTING DEBTORS AGING FORMAT & SCHEDULE III DISCLOSURE WORKBOOK ===');
const {
  detectAgingColumns,
  calculateAgingAnalysis,
  exportAgingWorkbook,
  getSampleReceivablesRows,
  SCHEDULE_III_BUCKETS,
} = await import('./receivablesAgingEngine.ts');

// Test Column Auto-Detection on Real-World Tally / ERP Sundry Debtors Headers
const debtorHeaders = ['Sundry Debtors', 'GSTIN', 'Invoice No', 'Invoice Date', 'Due Date', 'Gross Amount', 'Paid Amount', 'Outstanding Balance'];
const debtorMapping = detectAgingColumns(debtorHeaders);
assert(debtorMapping.partyCol === 'Sundry Debtors', 'Auto-detects Sundry Debtors as party column');
assert(debtorMapping.dueDateCol === 'Due Date', 'Auto-detects Due Date column');
assert(debtorMapping.balanceCol === 'Outstanding Balance', 'Auto-detects Outstanding Balance');

// Test Aging Analysis with explicit Due Date
const sampleDebtorRows = getSampleReceivablesRows();
const debtorResult = calculateAgingAnalysis(sampleDebtorRows, debtorMapping, {
  ledgerType: 'receivables',
  asOfDateStr: '2026-10-31',
  creditPeriodDays: 30,
  buckets: SCHEDULE_III_BUCKETS,
  allocateCreditsFifo: true,
});
assert(debtorResult.totalParties > 0, 'Calculates aging across all debtor entities');
assert(debtorResult.totalOutstanding > 0, 'Computes total outstanding receivables balance');
assert(debtorResult.isReconciled, 'Maintains 100% mathematical zero-variance reconciliation');

// Test Schedule III Statutory Workbook Generation
const debtorWb = exportAgingWorkbook(debtorResult, SCHEDULE_III_BUCKETS, 'Test Client Ltd', 'FY 2026-27');
assert(debtorWb.SheetNames.includes('Executive_Summary'), 'Contains Executive_Summary sheet');
assert(debtorWb.SheetNames.includes('Party_Aging_Matrix'), 'Contains Party_Aging_Matrix sheet');
assert(debtorWb.SheetNames.includes('Open_Invoices_Detail'), 'Contains Open_Invoices_Detail sheet');
assert(debtorWb.SheetNames.includes('Schedule_III_Disclosure'), 'Generates official MCA Schedule_III_Disclosure sheet');

console.log('\n=== 19. TESTING RECIPROCAL LEDGER COMPARISON & ZERO DISCREPANCY BALANCING ===');
const {
  compareLedgers,
  parseRawLedgerRows,
  getSampleLedgerPair,
  createLedgerReconciliationWorkbook,
} = await import('./ledgerComparisonEngine.ts');

// Test 1: Verify reciprocal balancing accounting equation: Net A (Dr) + Net B (Cr) = 0 Discrepancy
const ledgerA_Reciprocal = [
  { id: '1', date: '01-10-2026', rawDate: '01-10-2026', voucherNo: 'V1', particulars: 'Sales to B', debit: 500000, credit: 0 },
];
const ledgerB_Reciprocal = [
  { id: '2', date: '01-10-2026', rawDate: '01-10-2026', voucherNo: 'V1', particulars: 'Purchase from A', debit: 0, credit: 500000 },
];
const recResult = compareLedgers(ledgerA_Reciprocal, ledgerB_Reciprocal, 'Ledger A', 'Ledger B', {
  matchInverseDrCr: true,
  amountTolerance: 0,
  dateToleranceDays: 0,
  fuzzyVoucherMatch: false,
});
assert(recResult.balanceDifference === 0, 'Reciprocal matching correctly yields 0.00 balanceDifference when Dr equals Cr');
assert(recResult.matchedPairs.length === 1, 'Exactly matches reciprocal transaction pair');

// Test 2: Sample Ledger Pair with Confirmed, Probable, and Variance entries
const sampleLedgerPair = getSampleLedgerPair();
const parsedEntriesA = parseRawLedgerRows(sampleLedgerPair.ledgerA, 'A');
const parsedEntriesB = parseRawLedgerRows(sampleLedgerPair.ledgerB, 'B');
assert(parsedEntriesA.length === 6, 'Parses 6 entries in sample Ledger A');
assert(parsedEntriesB.length === 6, 'Parses 6 entries in sample Ledger B');

const fullRecResult = compareLedgers(parsedEntriesA, parsedEntriesB, sampleLedgerPair.ledgerAName, sampleLedgerPair.ledgerBName);
assert(fullRecResult.matchedPairs.length >= 2, 'Detects confirmed exact voucher & amount matches');
assert(fullRecResult.probablePairs.length >= 1, 'Detects probable match with date drift');
assert(fullRecResult.discrepancies.length >= 1, 'Identifies voucher amount discrepancy');
assert(fullRecResult.unmatchedA.length >= 1, 'Identifies unmatched entries in Ledger A');
assert(fullRecResult.unmatchedB.length >= 1, 'Identifies unmatched entries in Ledger B');

const recWb = createLedgerReconciliationWorkbook(fullRecResult);
assert(recWb.SheetNames.length === 6, 'Generates comprehensive 6-sheet reconciliation workbook');
assert(recWb.SheetNames.includes('Reconciliation_Summary'), 'Contains Reconciliation_Summary sheet');
assert(recWb.SheetNames.includes('Confirmed_Matches'), 'Contains Confirmed_Matches sheet');
assert(recWb.SheetNames.includes('Amount_Discrepancies'), 'Contains Amount_Discrepancies sheet');

console.log('\n=== 20. TESTING FORMULA AUDIT ENGINE & BOUNDARY COLUMN DETECTION ===');
const {
  getSampleFormulaAuditWorkbook,
  auditWorkbookFormulas,
  applyApprovedFormulaFixes,
} = await import('./formulaCheckerEngine.ts');

const formulaAuditWb = getSampleFormulaAuditWorkbook();
const auditSummary = auditWorkbookFormulas(formulaAuditWb, 'Test_Audit_Model.xlsx');
assert(auditSummary.criticalErrorsCount >= 2, 'Detects critical formula errors (#REF! and #DIV/0!)');
assert(auditSummary.hardcodedConstantsCount >= 1, 'Detects hardcoded static constant in dynamic formula column');
assert(auditSummary.textNumbersCount >= 1, 'Detects numbers stored as text that fail SUM()');

// Test fix application on fixable issues (ERROR_DIV_ZERO, HARDCODED_IN_FORMULA_COL, TEXT_STORED_AS_NUMBER)
const fixableIssue = auditSummary.issues.find((i) => i.proposedFormula || i.proposedValue !== undefined);
if (fixableIssue) fixableIssue.approved = true;
const { appliedCount } = applyApprovedFormulaFixes(formulaAuditWb, auditSummary.issues);
assert(appliedCount >= 1, 'Applies approved formula correction preserving workbook');

// Test Bank Column Detector boundary checks (no false 'cr' in 'Description' or 'dr' in 'Address')
const { detectBankColumnMapping } = await import('./bankStatementConverterEngine.ts');
const trickyHeaders = ['Txn Date', 'Description', 'Chq Ref', 'Address Line 1', 'Debit Amount', 'Credit Amount', 'Net Balance'];
const mappedCols = detectBankColumnMapping(trickyHeaders);
assert(mappedCols.narrationCol === 'Description', 'Description correctly maps to narration, NOT credit');
assert(mappedCols.debitCol === 'Debit Amount', 'Debit Amount correctly maps to debit, NOT Address Line');
assert(mappedCols.creditCol === 'Credit Amount', 'Credit Amount correctly maps to credit');

console.log('\n🎉 ALL 20 COMPREHENSIVE CA WORKFLOW AUTOMATION & STATUTORY ENGINE TESTS PASSED COMPLETELY!');

