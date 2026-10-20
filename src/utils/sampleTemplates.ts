import { exportStyledWorkbook } from './excelStyler';

/**
 * Generates and downloads a pre-formatted, professionally styled Excel template for any CA tool
 */
export function downloadTemplate(
  tool: 'excel-cleaner' | 'bank-rec' | 'gst-rec' | 'trial-balance' | 'audit'
) {
  if (tool === 'excel-cleaner') {
    const data = [
      {
        Voucher_No: 'SAL/26/0891',
        Invoice_Date: '09/04/2026',
        Party_Name: 'Apex Logistics Solutions Pvt Ltd',
        GSTIN: '27AABCA1234F1Z5',
        Taxable_Amt: 150000.0,
        CGST: 13500.0,
        SGST: 13500.0,
        IGST: 0.0,
        Total_Amt: 177000.0,
      },
      {
        Voucher_No: 'SAL/26/0892',
        Invoice_Date: '2026-09-06',
        Party_Name: 'Bharat Biotech Ancillaries LLP',
        GSTIN: '07AAACB9876Q1Z2',
        Taxable_Amt: -45000.0,
        CGST: -4050.0,
        SGST: -4050.0,
        IGST: 0.0,
        Total_Amt: -53100.0,
      },
      {
        Voucher_No: 'SAL/26/0893',
        Invoice_Date: '08-Sep-26',
        Party_Name: 'Sterling Retail Corp',
        GSTIN: '29AALCS4567M1Z9',
        Taxable_Amt: 320000.0,
        CGST: 0.0,
        SGST: 0.0,
        IGST: 57600.0,
        Total_Amt: 377600.0,
      },
    ];
    exportStyledWorkbook(
      [{ name: 'Sales_Register_Template', data }],
      'Excel_Cleaner_Template.xlsx'
    );
  } else if (tool === 'bank-rec') {
    const bankData = [
      { Date: '02-09-2026', Narration: 'NEFT CR - BHARAT SHELL - CMS289104', 'Ref No': 'CMS289104', Withdrawal: 0, Deposit: 485000, Balance: 1485000 },
      { Date: '04-09-2026', Narration: 'CHQ WDL - 004128 - SHREE LOGISTICS', 'Ref No': '004128', Withdrawal: 94250, Deposit: 0, Balance: 1390750 },
      { Date: '08-09-2026', Narration: 'RTGS INW - RELIANCE RETAIL - AXIS9921', 'Ref No': 'AXIS9921', Withdrawal: 0, Deposit: 1250000, Balance: 2640750 },
      { Date: '12-09-2026', Narration: 'ACH DR - TATA MOTORS EMI 8820', 'Ref No': 'ACH8820', Withdrawal: 184500, Deposit: 0, Balance: 2456250 },
      { Date: '15-09-2026', Narration: 'BANK CHARGES - FOREX PROCESSING', 'Ref No': 'CHG901', Withdrawal: 1770, Deposit: 0, Balance: 2454480 },
    ];
    const ledgerData = [
      { Date: '02-09-2026', 'Voucher No': 'CMS289104', Particulars: 'Bharat Shell Limited', Debit: 485000, Credit: 0 },
      { Date: '04-09-2026', 'Voucher No': '004128', Particulars: 'Shree Logistics Spares', Debit: 0, Credit: 94250 },
      { Date: '08-09-2026', 'Voucher No': 'AXIS9921', Particulars: 'Reliance Retail Depot', Debit: 1250000, Credit: 0 },
      { Date: '12-09-2026', 'Voucher No': 'ACH8820', Particulars: 'Tata Motors Commercial Loan', Debit: 0, Credit: 184500 },
      { Date: '28-09-2026', 'Voucher No': '004130', Particulars: 'Apollo Tyres & Rubber Corp', Debit: 0, Credit: 50000 },
    ];
    exportStyledWorkbook(
      [
        { name: 'Bank_Statement_Template', data: bankData },
        { name: 'Tally_Ledger_Template', data: ledgerData },
      ],
      'Bank_Reconciliation_Templates.xlsx'
    );
  } else if (tool === 'gst-rec') {
    const prData = [
      { GSTIN: '27AABCA1234F1Z5', 'Supplier Name': 'Apex Logistics Solutions', 'Invoice No': 'INV-2026-01', Date: '05-09-2026', 'Taxable Value': 150000, CGST: 13500, SGST: 13500, IGST: 0, 'Total Value': 177000 },
      { GSTIN: '07AAACB9876Q1Z2', 'Supplier Name': 'Bharat Biotech Ancillaries', 'Invoice No': 'BB-9921', Date: '12-09-2026', 'Taxable Value': 200000, CGST: 0, SGST: 0, IGST: 36000, 'Total Value': 236000 },
      { GSTIN: '29AALCS4567M1Z9', 'Supplier Name': 'Sterling Retail Corp', 'Invoice No': 'ST-8810', Date: '15-09-2026', 'Taxable Value': 80000, CGST: 7200, SGST: 7200, IGST: 0, 'Total Value': 94400 },
    ];
    const gstrData = [
      { 'GSTIN of supplier': '27AABCA1234F1Z5', 'Trade/Legal name': 'Apex Logistics Solutions', 'Invoice number': 'INV-2026-01', 'Invoice date': '05-09-2026', 'Taxable Value (₹)': 150000, 'Central Tax (₹)': 13500, 'State/UT Tax (₹)': 13500, 'Integrated Tax (₹)': 0, 'Invoice Value (₹)': 177000 },
      { 'GSTIN of supplier': '07AAACB9876Q1Z2', 'Trade/Legal name': 'Bharat Biotech Ancillaries', 'Invoice number': 'BB-9921', 'Invoice date': '12-09-2026', 'Taxable Value (₹)': 200000, 'Central Tax (₹)': 0, 'State/UT Tax (₹)': 0, 'Integrated Tax (₹)': 32000, 'Invoice Value (₹)': 232000 },
    ];
    exportStyledWorkbook(
      [
        { name: 'Purchase_Register_Template', data: prData },
        { name: 'GSTR2B_Portal_Template', data: gstrData },
      ],
      'GST_Reconciliation_Templates.xlsx'
    );
  } else if (tool === 'trial-balance') {
    const tbData = [
      { 'Account Code': '1001', 'Account Name': 'Cash in Hand', Head: 'Asset', Debit: 45200, Credit: 0 },
      { 'Account Code': '1010', 'Account Name': 'HDFC Bank Current Account', Head: 'Asset', Debit: 2730230, Credit: 0 },
      { 'Account Code': '1050', 'Account Name': 'Sundry Debtors', Head: 'Asset', Debit: 8450000, Credit: 0 },
      { 'Account Code': '2001', 'Account Name': 'Share Capital', Head: 'Equity', Debit: 0, Credit: 10000000 },
      { 'Account Code': '2200', 'Account Name': 'Sundry Creditors', Head: 'Liability', Debit: 0, Credit: 3840000 },
      { 'Account Code': '3001', 'Account Name': 'Revenue from Operations', Head: 'Revenue', Debit: 0, Credit: 32400000 },
      { 'Account Code': '4001', 'Account Name': 'Operating Direct Expenses', Head: 'Expense', Debit: 18680000, Credit: 0 },
      { 'Account Code': '4100', 'Account Name': 'Employee Benefit Expenses', Head: 'Expense', Debit: 7000000, Credit: 0 },
      { 'Account Code': '4300', 'Account Name': 'Administrative & Legal Expenses', Head: 'Expense', Debit: 350000, Credit: 0 },
      { 'Account Code': '2050', 'Account Name': 'Retained Earnings', Head: 'Equity', Debit: 0, Credit: 8965430 },
    ];
    exportStyledWorkbook(
      [{ name: 'Trial_Balance_Template', data: tbData }],
      'Trial_Balance_Template.xlsx'
    );
  } else if (tool === 'audit') {
    const auditData = [
      { 'Voucher No': 'PMT/26/0101', Date: '05-09-2026', Particulars: 'Apex Logistics Solutions', Amount: 177000, 'Payment Mode': 'Bank NEFT' },
      { 'Voucher No': 'PMT/26/0102', Date: '14-09-2026', Particulars: 'Petty Cash Imprest Withdrawal', Amount: 500000, 'Payment Mode': 'Cash' },
      { 'Voucher No': 'PMT/26/0103', Date: '20-09-2026', Particulars: 'Universal Auto Body Repairs', Amount: 28500, 'Payment Mode': 'Cash' },
      { 'Voucher No': 'SAL/26/0104', Date: '27-09-2026', Particulars: 'Sunday Depot Clearance Sale', Amount: 1450000, 'Payment Mode': 'Bank RTGS' },
      { 'Voucher No': 'PMT/26/0105', Date: '28-09-2026', Particulars: 'Fast Tyre Retreading Hub', Amount: 49000, 'Payment Mode': 'Bank IMPS' },
    ];
    exportStyledWorkbook(
      [{ name: 'Ledger_Dump_Template', data: auditData }],
      'Audit_Sampling_Ledger_Template.xlsx'
    );
  }
}
