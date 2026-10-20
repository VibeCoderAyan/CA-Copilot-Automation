import * as XLSX from 'xlsx';
import { createStyledWorksheet } from './excelStyler';

export type TemplateCategory =
  | 'Reconciliation'
  | 'Tax Audit (Form 3CD)'
  | 'Statutory Audit'
  | 'Taxation & TDS'
  | 'Corporate Accounting';

export type TemplateId =
  | 'bank_rec'
  | 'gst_rec'
  | 'fixed_asset_register'
  | 'msme_debtors_creditors'
  | 'form3cd_clause21d'
  | 'tds_quarterly_register'
  | 'stock_verification'
  | 'expense_register'
  | 'general_ledger'
  | 'trial_balance'
  | 'audit_working_paper'
  | 'tax_invoice_register';

export interface TemplateDefinition {
  id: TemplateId;
  name: string;
  category: TemplateCategory;
  statutoryStandard: string;
  description: string;
  sheetName: string;
  instructions: string[];
  columns: string[];
  blankData: Record<string, any>[];
  sampleData: Record<string, any>[];
  currencyColumns?: string[];
  dateColumns?: string[];
  columnWidths?: Record<string, number>;
}

export const TEMPLATES: Record<TemplateId, TemplateDefinition> = {
  bank_rec: {
    id: 'bank_rec',
    name: 'Bank Reconciliation Statement (BRS)',
    category: 'Reconciliation',
    statutoryStandard: 'ICAI SA 505 & Banking Guidance Note',
    description: 'Reconcile ledger balance with bank statements, tracking unpresented cheques and uncredited deposits.',
    sheetName: 'Bank_Reconciliation',
    instructions: [
      'Enter the opening closing balance as per company books in line 1.',
      'List unpresented cheques (issued but not cleared) under Additions.',
      'List direct bank credits and unbooked interest under Additions.',
      'List uncredited deposits and bank charges/penalties under Deductions.',
      'Verify that the final computed balance matches the official bank statement exactly.',
    ],
    columns: ['Particulars / Transaction Description', 'Reference / Chq No', 'Date', 'Debit (₹)', 'Credit (₹)', 'Net Adjustment (₹)', 'Adjustment Type', 'Clearance Date', 'Audit Verification Note'],
    currencyColumns: ['Debit (₹)', 'Credit (₹)', 'Net Adjustment (₹)'],
    dateColumns: ['Date', 'Clearance Date'],
    columnWidths: {
      'Particulars / Transaction Description': 44,
      'Reference / Chq No': 18,
      'Date': 14,
      'Debit (₹)': 16,
      'Credit (₹)': 16,
      'Net Adjustment (₹)': 18,
      'Adjustment Type': 22,
      'Clearance Date': 16,
      'Audit Verification Note': 36,
    },
    blankData: [
      { 'Particulars / Transaction Description': '1. Balance as per Company Books (GL-1010)', 'Reference / Chq No': 'GL-1010', 'Date': '31-03-2026', 'Debit (₹)': 0, 'Credit (₹)': 0, 'Net Adjustment (₹)': 0, 'Adjustment Type': 'Base Book Balance', 'Clearance Date': '-', 'Audit Verification Note': 'As per general ledger trial balance' },
      { 'Particulars / Transaction Description': 'Add: Cheque issued but not presented for payment', 'Reference / Chq No': 'CHQ-000001', 'Date': '28-03-2026', 'Debit (₹)': 0, 'Credit (₹)': 0, 'Net Adjustment (₹)': 0, 'Adjustment Type': 'Add to Books', 'Clearance Date': 'Pending', 'Audit Verification Note': 'Vendor payment outstanding' },
      { 'Particulars / Transaction Description': 'Add: Direct customer remittance not posted in books', 'Reference / Chq No': 'NEFT-000001', 'Date': '30-03-2026', 'Debit (₹)': 0, 'Credit (₹)': 0, 'Net Adjustment (₹)': 0, 'Adjustment Type': 'Add to Books', 'Clearance Date': '30-03-2026', 'Audit Verification Note': 'Customer NEFT credit' },
      { 'Particulars / Transaction Description': 'Less: Cheque deposited but not cleared by bank', 'Reference / Chq No': 'DEP-000001', 'Date': '30-03-2026', 'Debit (₹)': 0, 'Credit (₹)': 0, 'Net Adjustment (₹)': 0, 'Adjustment Type': 'Deduct from Books', 'Clearance Date': 'Pending', 'Audit Verification Note': 'Under transit clearing' },
      { 'Particulars / Transaction Description': 'Less: Bank service charges / auto-debit interest', 'Reference / Chq No': 'DR-CHG-01', 'Date': '31-03-2026', 'Debit (₹)': 0, 'Credit (₹)': 0, 'Net Adjustment (₹)': 0, 'Adjustment Type': 'Deduct from Books', 'Clearance Date': '31-03-2026', 'Audit Verification Note': 'Bank statement debit' },
      { 'Particulars / Transaction Description': 'Balance as per Official Bank Statement', 'Reference / Chq No': 'STMT-MAR26', 'Date': '31-03-2026', 'Debit (₹)': 0, 'Credit (₹)': 0, 'Net Adjustment (₹)': 0, 'Adjustment Type': 'Reconciled Bank Balance', 'Clearance Date': '31-03-2026', 'Audit Verification Note': 'Verified against bank confirmation letter' },
    ],
    sampleData: [
      { 'Particulars / Transaction Description': '1. Balance as per Company Books (HDFC Current A/c)', 'Reference / Chq No': 'GL-1010', 'Date': '31-03-2026', 'Debit (₹)': 1450000, 'Credit (₹)': 0, 'Net Adjustment (₹)': 1450000, 'Adjustment Type': 'Base Book Balance', 'Clearance Date': '31-03-2026', 'Audit Verification Note': 'Agreed with Tally General Ledger' },
      { 'Particulars / Transaction Description': 'Add: Cheque issued to Tata Steel Limited', 'Reference / Chq No': 'CHQ-882101', 'Date': '28-03-2026', 'Debit (₹)': 0, 'Credit (₹)': 285000, 'Net Adjustment (₹)': 285000, 'Adjustment Type': 'Add to Books', 'Clearance Date': '04-04-2026', 'Audit Verification Note': 'Cleared subsequently on 4th April' },
      { 'Particulars / Transaction Description': 'Add: Direct NEFT from Infosys BPO Ltd', 'Reference / Chq No': 'NEFT-551209', 'Date': '30-03-2026', 'Debit (₹)': 420000, 'Credit (₹)': 0, 'Net Adjustment (₹)': 420000, 'Adjustment Type': 'Add to Books', 'Clearance Date': '30-03-2026', 'Audit Verification Note': 'Customer direct deposit (Receipt to be posted)' },
      { 'Particulars / Transaction Description': 'Less: Outstation cheque deposited (Advance from L&T)', 'Reference / Chq No': 'DEP-441092', 'Date': '29-03-2026', 'Debit (₹)': 310000, 'Credit (₹)': 0, 'Net Adjustment (₹)': -310000, 'Adjustment Type': 'Deduct from Books', 'Clearance Date': '03-04-2026', 'Audit Verification Note': 'Cleared on 3rd April in bank statement' },
      { 'Particulars / Transaction Description': 'Less: Annual Bank Service Charges & AMC', 'Reference / Chq No': 'DR-CHG-99', 'Date': '31-03-2026', 'Debit (₹)': 0, 'Credit (₹)': 4500, 'Net Adjustment (₹)': -4500, 'Adjustment Type': 'Deduct from Books', 'Clearance Date': '31-03-2026', 'Audit Verification Note': 'Adjustment journal posted in books' },
      { 'Particulars / Transaction Description': 'Balance as per Official Bank Statement', 'Reference / Chq No': 'STMT-MAR26', 'Date': '31-03-2026', 'Debit (₹)': 1840500, 'Credit (₹)': 0, 'Net Adjustment (₹)': 1840500, 'Adjustment Type': 'Reconciled Bank Balance', 'Clearance Date': '31-03-2026', 'Audit Verification Note': '100% matched with HDFC NetBanking confirmation' },
    ],
  },

  gst_rec: {
    id: 'gst_rec',
    name: 'GSTR-2B vs Books ITC Reconciliation',
    category: 'Reconciliation',
    statutoryStandard: 'CGST Act Sec 16(2)(aa) & Rule 88C / 88D',
    description: 'Match purchase register with auto-drafted GSTR-2B to verify ITC eligibility and flag blocked credits.',
    sheetName: 'GSTR2B_vs_Books_ITC',
    instructions: [
      'Enter supplier GSTIN, invoice details, and tax split from Purchase Register.',
      'Cross-reference with GSTR-2B portal data to verify availability of ITC.',
      'Flag blocked ITC under Section 17(5) (e.g. food, motor vehicles, corporate gifts).',
      'Generate vendor follow-up notices for invoices missing in GSTR-2B.',
    ],
    columns: ['S.No', 'Supplier GSTIN', 'Supplier Legal Trade Name', 'Invoice Number', 'Invoice Date', 'HSN / SAC', 'Taxable Value (₹)', 'IGST (₹)', 'CGST (₹)', 'SGST (₹)', 'Total Invoice (₹)', '2B ITC Status', 'ITC Eligibility', 'Matching Discrepancy', 'CA Action Note'],
    currencyColumns: ['Taxable Value (₹)', 'IGST (₹)', 'CGST (₹)', 'SGST (₹)', 'Total Invoice (₹)'],
    dateColumns: ['Invoice Date'],
    columnWidths: {
      'S.No': 8,
      'Supplier GSTIN': 18,
      'Supplier Legal Trade Name': 28,
      'Invoice Number': 18,
      'Invoice Date': 14,
      'HSN / SAC': 14,
      'Taxable Value (₹)': 18,
      'IGST (₹)': 16,
      'CGST (₹)': 16,
      'SGST (₹)': 16,
      'Total Invoice (₹)': 20,
      '2B ITC Status': 18,
      'ITC Eligibility': 16,
      'Matching Discrepancy': 22,
      'CA Action Note': 34,
    },
    blankData: [
      { 'S.No': 1, 'Supplier GSTIN': '27AAACA0000A1Z5', 'Supplier Legal Trade Name': 'Sample Supplier Pvt Ltd', 'Invoice Number': 'INV-0001', 'Invoice Date': '15-09-2026', 'HSN / SAC': '8471', 'Taxable Value (₹)': 0, 'IGST (₹)': 0, 'CGST (₹)': 0, 'SGST (₹)': 0, 'Total Invoice (₹)': 0, '2B ITC Status': 'Available in 2B', 'ITC Eligibility': 'Eligible', 'Matching Discrepancy': 'None (Matched)', 'CA Action Note': 'Claim 100% ITC in GSTR-3B Table 4(A)(5)' },
    ],
    sampleData: [
      { 'S.No': 1, 'Supplier GSTIN': '27AAACA1234A1Z5', 'Supplier Legal Trade Name': 'Apollo Tyres Limited', 'Invoice Number': 'INV-4412', 'Invoice Date': '15-09-2026', 'HSN / SAC': '401110', 'Taxable Value (₹)': 125000, 'IGST (₹)': 0, 'CGST (₹)': 11250, 'SGST (₹)': 11250, 'Total Invoice (₹)': 147500, '2B ITC Status': 'Available in 2B', 'ITC Eligibility': 'Eligible', 'Matching Discrepancy': 'None (Matched)', 'CA Action Note': 'Claim 100% ITC in GSTR-3B Table 4(A)(5)' },
      { 'S.No': 2, 'Supplier GSTIN': '29BBBCB5678B2Z1', 'Supplier Legal Trade Name': 'Total Lubricants India Ltd', 'Invoice Number': 'TL-8821', 'Invoice Date': '20-09-2026', 'HSN / SAC': '271019', 'Taxable Value (₹)': 240000, 'IGST (₹)': 43200, 'CGST (₹)': 0, 'SGST (₹)': 0, 'Total Invoice (₹)': 283200, '2B ITC Status': 'Available in 2B', 'ITC Eligibility': 'Eligible', 'Matching Discrepancy': 'None (Matched)', 'CA Action Note': 'Inter-state IGST verified with E-Way Bill' },
      { 'S.No': 3, 'Supplier GSTIN': '07CCCC8901C3Z8', 'Supplier Legal Trade Name': 'Modern Corporate Gifting Ltd', 'Invoice Number': 'MCG-102', 'Invoice Date': '22-09-2026', 'HSN / SAC': '998399', 'Taxable Value (₹)': 45000, 'IGST (₹)': 0, 'CGST (₹)': 4050, 'SGST (₹)': 4050, 'Total Invoice (₹)': 53100, '2B ITC Status': 'Available in 2B', 'ITC Eligibility': 'Ineligible', 'Matching Discrepancy': 'Blocked Sec 17(5)', 'CA Action Note': 'Report in Table 4(B)(1) as Ineligible' },
      { 'S.No': 4, 'Supplier GSTIN': '27DDDCD3344D1Z2', 'Supplier Legal Trade Name': 'Express Cloud Solutions LLP', 'Invoice Number': 'ECS-902', 'Invoice Date': '26-09-2026', 'HSN / SAC': '998313', 'Taxable Value (₹)': 80000, 'IGST (₹)': 0, 'CGST (₹)': 7200, 'SGST (₹)': 7200, 'Total Invoice (₹)': 94400, '2B ITC Status': 'Missing in 2B', 'ITC Eligibility': 'Ineligible', 'Matching Discrepancy': 'Supplier GSTR-1 Pending', 'CA Action Note': 'Withhold tax payment under Sec 16(2)(aa)' },
      { 'S.No': 'TOTAL', 'Supplier GSTIN': '-', 'Supplier Legal Trade Name': 'Grand Total (₹)', 'Invoice Number': '-', 'Invoice Date': '-', 'HSN / SAC': '-', 'Taxable Value (₹)': 490000, 'IGST (₹)': 43200, 'CGST (₹)': 22500, 'SGST (₹)': 22500, 'Total Invoice (₹)': 578200, '2B ITC Status': '-', 'ITC Eligibility': '-', 'Matching Discrepancy': '-', 'CA Action Note': 'Net Eligible ITC: ₹65,700' },
    ],
  },

  fixed_asset_register: {
    id: 'fixed_asset_register',
    name: 'Fixed Asset Register & Depreciation Schedule',
    category: 'Statutory Audit',
    statutoryStandard: 'Companies Act Schedule II & IT Act Sec 32',
    description: 'Block of assets register tracking additions, 180-day put-to-use cutoff, WDV depreciation, and net book values.',
    sheetName: 'Fixed_Asset_Register',
    instructions: [
      'Maintain asset identity codes, physical locations, and asset block classifications.',
      'Record exact date of acquisition and date put to use.',
      'Assets put to use for less than 180 days in the year are restricted to 50% statutory depreciation under IT Act.',
      'Ensure net book value never drops below 5% salvage value under Companies Act 2013.',
    ],
    columns: ['Asset Code', 'Asset Description', 'Block of Assets / Class', 'Department / Location', 'Acquisition Date', 'Put to Use Date', '>180 Days (Y/N)', 'Opening Gross Block (₹)', 'Additions during Year (₹)', 'Deletions / Sales (₹)', 'Total Gross Block (₹)', 'Depreciation Rate (%)', 'Depreciation for Year (₹)', 'Closing WDV (₹)', 'Physical Verification Status'],
    currencyColumns: ['Opening Gross Block (₹)', 'Additions during Year (₹)', 'Deletions / Sales (₹)', 'Total Gross Block (₹)', 'Depreciation for Year (₹)', 'Closing WDV (₹)'],
    dateColumns: ['Acquisition Date', 'Put to Use Date'],
    columnWidths: {
      'Asset Code': 14,
      'Asset Description': 30,
      'Block of Assets / Class': 24,
      'Department / Location': 20,
      'Acquisition Date': 16,
      'Put to Use Date': 16,
      '>180 Days (Y/N)': 16,
      'Opening Gross Block (₹)': 22,
      'Additions during Year (₹)': 22,
      'Deletions / Sales (₹)': 20,
      'Total Gross Block (₹)': 22,
      'Depreciation Rate (%)': 18,
      'Depreciation for Year (₹)': 22,
      'Closing WDV (₹)': 20,
      'Physical Verification Status': 26,
    },
    blankData: [
      { 'Asset Code': 'FA-001', 'Asset Description': 'Plant & Machinery - CNC Machine', 'Block of Assets / Class': 'Plant & Machinery (15%)', 'Department / Location': 'Shop Floor 1', 'Acquisition Date': '15-05-2025', 'Put to Use Date': '01-06-2025', '>180 Days (Y/N)': 'YES', 'Opening Gross Block (₹)': 0, 'Additions during Year (₹)': 0, 'Deletions / Sales (₹)': 0, 'Total Gross Block (₹)': 0, 'Depreciation Rate (%)': 15, 'Depreciation for Year (₹)': 0, 'Closing WDV (₹)': 0, 'Physical Verification Status': 'Verified' },
    ],
    sampleData: [
      { 'Asset Code': 'FA-PM-01', 'Asset Description': 'Automatic CNC Lathe Unit', 'Block of Assets / Class': 'Plant & Machinery (15%)', 'Department / Location': 'Factory Unit A', 'Acquisition Date': '10-04-2025', 'Put to Use Date': '25-04-2025', '>180 Days (Y/N)': 'YES', 'Opening Gross Block (₹)': 4500000, 'Additions during Year (₹)': 0, 'Deletions / Sales (₹)': 0, 'Total Gross Block (₹)': 4500000, 'Depreciation Rate (%)': 15, 'Depreciation for Year (₹)': 675000, 'Closing WDV (₹)': 3825000, 'Physical Verification Status': 'Verified - Tagged PM-101' },
      { 'Asset Code': 'FA-PM-02', 'Asset Description': 'High-Speed Packaging Line', 'Block of Assets / Class': 'Plant & Machinery (15%)', 'Department / Location': 'Packaging Bay', 'Acquisition Date': '15-11-2025', 'Put to Use Date': '01-12-2025', '>180 Days (Y/N)': 'NO', 'Opening Gross Block (₹)': 0, 'Additions during Year (₹)': 1800000, 'Deletions / Sales (₹)': 0, 'Total Gross Block (₹)': 1800000, 'Depreciation Rate (%)': 7.5, 'Depreciation for Year (₹)': 135000, 'Closing WDV (₹)': 1665000, 'Physical Verification Status': 'Verified - Tagged PM-102 (<180D half rate)' },
      { 'Asset Code': 'FA-IT-01', 'Asset Description': 'Dell PowerEdge Enterprise Server', 'Block of Assets / Class': 'Computers & Servers (40%)', 'Department / Location': 'Data Center', 'Acquisition Date': '01-05-2025', 'Put to Use Date': '10-05-2025', '>180 Days (Y/N)': 'YES', 'Opening Gross Block (₹)': 850000, 'Additions during Year (₹)': 0, 'Deletions / Sales (₹)': 0, 'Total Gross Block (₹)': 850000, 'Depreciation Rate (%)': 40, 'Depreciation for Year (₹)': 340000, 'Closing WDV (₹)': 510000, 'Physical Verification Status': 'Verified - Asset Tag IT-01' },
      { 'Asset Code': 'FA-VH-01', 'Asset Description': 'Commercial Delivery Van', 'Block of Assets / Class': 'Motor Vehicles (15%)', 'Department / Location': 'Logistics Hub', 'Acquisition Date': '12-06-2024', 'Put to Use Date': '15-06-2024', '>180 Days (Y/N)': 'YES', 'Opening Gross Block (₹)': 1200000, 'Additions during Year (₹)': 0, 'Deletions / Sales (₹)': 300000, 'Total Gross Block (₹)': 900000, 'Depreciation Rate (%)': 15, 'Depreciation for Year (₹)': 135000, 'Closing WDV (₹)': 765000, 'Physical Verification Status': 'Verified - Reg MH12AB1234' },
      { 'Asset Code': 'TOTAL', 'Asset Description': 'Grand Total Fixed Assets (₹)', 'Block of Assets / Class': '-', 'Department / Location': '-', 'Acquisition Date': '-', 'Put to Use Date': '-', '>180 Days (Y/N)': '-', 'Opening Gross Block (₹)': 6550000, 'Additions during Year (₹)': 1800000, 'Deletions / Sales (₹)': 300000, 'Total Gross Block (₹)': 8050000, 'Depreciation Rate (%)': '-', 'Depreciation for Year (₹)': 1285000, 'Closing WDV (₹)': 6765000, 'Physical Verification Status': '100% Verified by Audit Team' },
    ],
  },

  msme_debtors_creditors: {
    id: 'msme_debtors_creditors',
    name: 'Debtors & Creditors Ageing with MSME Sec 43B(h)',
    category: 'Statutory Audit',
    statutoryStandard: 'IT Act Sec 43B(h) & MSMED Act 2006',
    description: 'Ageing analysis of payables and receivables, isolating overdue MSME dues liable to statutory disallowance.',
    sheetName: 'MSME_Ageing_Scrutiny',
    instructions: [
      'Enter party details, GSTIN, and MSME Udyam Registration number.',
      'Ageing must be computed from the date of acceptance or delivery.',
      'Under Section 43B(h), any sum payable to Micro or Small enterprise beyond 45 days (or 15 days if no agreement) is strictly DISALLOWED in taxable income.',
      'Obtain written balance confirmations (SA 505) for all material accounts exceeding audit threshold.',
    ],
    columns: ['Party Code', 'Party Trade Legal Name', 'GSTIN', 'MSME Enterprise Category', 'MSME Udyam Reg No', 'Agreed Credit Term (Days)', 'Due Date', 'Total Outstanding (₹)', '0-45 Days (₹)', '46-90 Days (₹)', '91-180 Days (₹)', '>180 Days (₹)', 'Sec 43B(h) Disallowance Exposure (₹)', 'Balance Confirmation Status'],
    currencyColumns: ['Total Outstanding (₹)', '0-45 Days (₹)', '46-90 Days (₹)', '91-180 Days (₹)', '>180 Days (₹)', 'Sec 43B(h) Disallowance Exposure (₹)'],
    dateColumns: ['Due Date'],
    columnWidths: {
      'Party Code': 14,
      'Party Trade Legal Name': 30,
      'GSTIN': 18,
      'MSME Enterprise Category': 24,
      'MSME Udyam Reg No': 22,
      'Agreed Credit Term (Days)': 22,
      'Due Date': 14,
      'Total Outstanding (₹)': 20,
      '0-45 Days (₹)': 18,
      '46-90 Days (₹)': 18,
      '91-180 Days (₹)': 18,
      '>180 Days (₹)': 18,
      'Sec 43B(h) Disallowance Exposure (₹)': 28,
      'Balance Confirmation Status': 26,
    },
    blankData: [
      { 'Party Code': 'CR-001', 'Party Trade Legal Name': 'Precision Engineering Tools', 'GSTIN': '27AAACP0000A1Z1', 'MSME Enterprise Category': 'Micro Enterprise', 'MSME Udyam Reg No': 'UDYAM-MH-01-00001', 'Agreed Credit Term (Days)': 45, 'Due Date': '15-03-2026', 'Total Outstanding (₹)': 0, '0-45 Days (₹)': 0, '46-90 Days (₹)': 0, '91-180 Days (₹)': 0, '>180 Days (₹)': 0, 'Sec 43B(h) Disallowance Exposure (₹)': 0, 'Balance Confirmation Status': 'Pending' },
    ],
    sampleData: [
      { 'Party Code': 'CR-101', 'Party Trade Legal Name': 'Apex Tooling Solutions LLP', 'GSTIN': '27AAACA1234A1Z5', 'MSME Enterprise Category': 'Micro Enterprise', 'MSME Udyam Reg No': 'UDYAM-MH-01-12345', 'Agreed Credit Term (Days)': 45, 'Due Date': '15-01-2026', 'Total Outstanding (₹)': 350000, '0-45 Days (₹)': 0, '46-90 Days (₹)': 350000, '91-180 Days (₹)': 0, '>180 Days (₹)': 0, 'Sec 43B(h) Disallowance Exposure (₹)': 350000, 'Balance Confirmation Status': 'Confirmed - ₹3,50,000 (Overdue >45D)' },
      { 'Party Code': 'CR-102', 'Party Trade Legal Name': 'Bharat Precision Castings', 'GSTIN': '29BBBCB5678B2Z1', 'MSME Enterprise Category': 'Small Enterprise', 'MSME Udyam Reg No': 'UDYAM-KR-03-98765', 'Agreed Credit Term (Days)': 30, 'Due Date': '20-03-2026', 'Total Outstanding (₹)': 480000, '0-45 Days (₹)': 480000, '46-90 Days (₹)': 0, '91-180 Days (₹)': 0, '>180 Days (₹)': 0, 'Sec 43B(h) Disallowance Exposure (₹)': 0, 'Balance Confirmation Status': 'Confirmed - Within statutory limit' },
      { 'Party Code': 'CR-103', 'Party Trade Legal Name': 'Tata Steel Downstream Products', 'GSTIN': '07CCCC8901C3Z8', 'MSME Enterprise Category': 'Large Enterprise (Non-MSME)', 'MSME Udyam Reg No': 'Not Applicable', 'Agreed Credit Term (Days)': 60, 'Due Date': '10-12-2025', 'Total Outstanding (₹)': 1250000, '0-45 Days (₹)': 0, '46-90 Days (₹)': 0, '91-180 Days (₹)': 1250000, '>180 Days (₹)': 0, 'Sec 43B(h) Disallowance Exposure (₹)': 0, 'Balance Confirmation Status': 'Confirmed - Non-MSME (No disallowance)' },
      { 'Party Code': 'CR-104', 'Party Trade Legal Name': 'Kaveri Industrial Packaging', 'GSTIN': '27DDDCD3344D1Z2', 'MSME Enterprise Category': 'Micro Enterprise', 'MSME Udyam Reg No': 'UDYAM-MH-02-55443', 'Agreed Credit Term (Days)': 45, 'Due Date': '05-11-2025', 'Total Outstanding (₹)': 210000, '0-45 Days (₹)': 0, '46-90 Days (₹)': 0, '91-180 Days (₹)': 210000, '>180 Days (₹)': 0, 'Sec 43B(h) Disallowance Exposure (₹)': 210000, 'Balance Confirmation Status': 'Disputed - Overdue >45D (Liable to tax)' },
      { 'Party Code': 'TOTAL', 'Party Trade Legal Name': 'Grand Total Outstanding (₹)', 'GSTIN': '-', 'MSME Enterprise Category': '-', 'MSME Udyam Reg No': '-', 'Agreed Credit Term (Days)': '-', 'Due Date': '-', 'Total Outstanding (₹)': 2290000, '0-45 Days (₹)': 480000, '46-90 Days (₹)': 350000, '91-180 Days (₹)': 1460000, '>180 Days (₹)': 0, 'Sec 43B(h) Disallowance Exposure (₹)': 560000, 'Balance Confirmation Status': 'Total Disallowance Exposure: ₹5,60,000' },
    ],
  },

  form3cd_clause21d: {
    id: 'form3cd_clause21d',
    name: 'Form 3CD Clause 21(d) Sec 40A(3) Scrutiny',
    category: 'Tax Audit (Form 3CD)',
    statutoryStandard: 'IT Act Sec 40A(3) & IT Rules Rule 6DD',
    description: 'Tax audit worksheet tracking cash payments exceeding statutory limits (₹10,000 / ₹35,000) for Form 3CD.',
    sheetName: 'Form3CD_Sec40A3_Cash',
    instructions: [
      'Scrutinize cash payments made in a single day to a single person/beneficiary.',
      'Statutory limits: ₹10,000 general limit; ₹35,000 for plying, hiring or leasing goods carriages (transporters).',
      'If payment falls under Rule 6DD exceptions (bank holidays, agricultural produce, NEFT breakdown), record proof.',
      'All unexempted breaches must be reported in Form 3CD Clause 21(d) as direct additions to business income.',
    ],
    columns: ['Voucher No', 'Transaction Date', 'Payee / Beneficiary Name', 'PAN / Aadhaar', 'Account Head Debited', 'Payment Mode', 'Aggregate Daily Cash Paid (₹)', 'Statutory Limit (₹)', 'Breach Amount (₹)', 'Rule 6DD Statutory Exemption Claimed', 'Tax Auditor Verification Finding'],
    currencyColumns: ['Aggregate Daily Cash Paid (₹)', 'Statutory Limit (₹)', 'Breach Amount (₹)'],
    dateColumns: ['Transaction Date'],
    columnWidths: {
      'Voucher No': 16,
      'Transaction Date': 16,
      'Payee / Beneficiary Name': 30,
      'PAN / Aadhaar': 18,
      'Account Head Debited': 24,
      'Payment Mode': 16,
      'Aggregate Daily Cash Paid (₹)': 24,
      'Statutory Limit (₹)': 18,
      'Breach Amount (₹)': 18,
      'Rule 6DD Statutory Exemption Claimed': 32,
      'Tax Auditor Verification Finding': 36,
    },
    blankData: [
      { 'Voucher No': 'PV-001', 'Transaction Date': '15-06-2025', 'Payee / Beneficiary Name': 'Ramesh Kumar Contractor', 'PAN / Aadhaar': 'ABCDE1234F', 'Account Head Debited': 'Repairs & Maintenance', 'Payment Mode': 'Bearer Cheque / Cash', 'Aggregate Daily Cash Paid (₹)': 0, 'Statutory Limit (₹)': 10000, 'Breach Amount (₹)': 0, 'Rule 6DD Statutory Exemption Claimed': 'None', 'Tax Auditor Verification Finding': 'Compliant' },
    ],
    sampleData: [
      { 'Voucher No': 'PV-104', 'Transaction Date': '14-06-2025', 'Payee / Beneficiary Name': 'Shree Ram Transport Roadways', 'PAN / Aadhaar': 'AABCS1122D', 'Account Head Debited': 'Freight & Cartage Inward', 'Payment Mode': 'Cash Payment', 'Aggregate Daily Cash Paid (₹)': 42000, 'Statutory Limit (₹)': 35000, 'Breach Amount (₹)': 42000, 'Rule 6DD Statutory Exemption Claimed': 'None', 'Tax Auditor Verification Finding': 'Disallowed u/s 40A(3) - Exceeds ₹35,000 limit' },
      { 'Voucher No': 'PV-212', 'Transaction Date': '15-08-2025', 'Payee / Beneficiary Name': 'Modern Electric Supplies', 'PAN / Aadhaar': 'BBBCD5566K', 'Account Head Debited': 'Factory Electrical Repairs', 'Payment Mode': 'Cash Payment', 'Aggregate Daily Cash Paid (₹)': 18500, 'Statutory Limit (₹)': 10000, 'Breach Amount (₹)': 0, 'Rule 6DD Statutory Exemption Claimed': 'Rule 6DD(j) - National Bank Holiday', 'Tax Auditor Verification Finding': 'Exempt under Rule 6DD(j) - Independence Day closure' },
      { 'Voucher No': 'PV-308', 'Transaction Date': '10-10-2025', 'Payee / Beneficiary Name': 'Kisan Agro Producers Trust', 'PAN / Aadhaar': 'AAACT9900L', 'Account Head Debited': 'Direct Raw Jute Purchases', 'Payment Mode': 'Cash Payment', 'Aggregate Daily Cash Paid (₹)': 85000, 'Statutory Limit (₹)': 10000, 'Breach Amount (₹)': 0, 'Rule 6DD Statutory Exemption Claimed': 'Rule 6DD(e) - Direct Agricultural Produce', 'Tax Auditor Verification Finding': 'Exempt under Rule 6DD(e) - Producer direct purchase' },
      { 'Voucher No': 'PV-415', 'Transaction Date': '22-01-2026', 'Payee / Beneficiary Name': 'Quick Courier Services', 'PAN / Aadhaar': 'CCCDP4455N', 'Account Head Debited': 'Postage & Courier Expenses', 'Payment Mode': 'Cash Payment', 'Aggregate Daily Cash Paid (₹)': 14000, 'Statutory Limit (₹)': 10000, 'Breach Amount (₹)': 14000, 'Rule 6DD Statutory Exemption Claimed': 'None', 'Tax Auditor Verification Finding': 'Disallowed u/s 40A(3) - Cash split in 2 vouchers on same day' },
      { 'Voucher No': 'TOTAL', 'Transaction Date': '-', 'Payee / Beneficiary Name': 'Grand Total Scrutinized (₹)', 'PAN / Aadhaar': '-', 'Account Head Debited': '-', 'Payment Mode': '-', 'Aggregate Daily Cash Paid (₹)': 159500, 'Statutory Limit (₹)': '-', 'Breach Amount (₹)': 56000, 'Rule 6DD Statutory Exemption Claimed': '-', 'Tax Auditor Verification Finding': 'Total Form 3CD Addition to Profit: ₹56,000' },
    ],
  },

  tds_quarterly_register: {
    id: 'tds_quarterly_register',
    name: 'TDS Quarterly Deduction & Form 26AS Matching',
    category: 'Taxation & TDS',
    statutoryStandard: 'IT Act Chapter XVII-B & Rule 31A',
    description: 'Quarterly TDS register with section mapping (194C, 194J, 194Q), Challan 281 details, and Form 26AS matching.',
    sheetName: 'TDS_Quarterly_Register',
    instructions: [
      'Enter deductee PAN, section code (e.g., 194C, 194J, 194Q, 194H, 194I).',
      'Verify 4th character of PAN: "P" for Individual/HUF (1% under 194C), "C"/"F" for Company/Firm (2% under 194C).',
      'Verify Section 206AA rate (20%) if valid PAN is missing.',
      'Record Challan 281 BSR code, deposit date, and challan sequence number for Form 24Q/26Q filing.',
    ],
    columns: ['Deductee PAN', 'Deductee Trade Name', 'PAN Status', 'TDS Section', 'Invoice / Bill No', 'Transaction Date', 'Gross Taxable Amount (₹)', 'TDS Rate (%)', 'TDS Deducted (₹)', 'Challan BSR Code', 'Challan Deposit Date', 'Challan Serial No', '26AS Matched (Y/N)', 'Sec 206AB Non-Filer Status'],
    currencyColumns: ['Gross Taxable Amount (₹)', 'TDS Deducted (₹)'],
    dateColumns: ['Transaction Date', 'Challan Deposit Date'],
    columnWidths: {
      'Deductee PAN': 16,
      'Deductee Trade Name': 28,
      'PAN Status': 14,
      'TDS Section': 14,
      'Invoice / Bill No': 18,
      'Transaction Date': 16,
      'Gross Taxable Amount (₹)': 22,
      'TDS Rate (%)': 14,
      'TDS Deducted (₹)': 18,
      'Challan BSR Code': 18,
      'Challan Deposit Date': 18,
      'Challan Serial No': 18,
      '26AS Matched (Y/N)': 16,
      'Sec 206AB Non-Filer Status': 24,
    },
    blankData: [
      { 'Deductee PAN': 'ABCDE1234F', 'Deductee Trade Name': 'Sample Contractor', 'PAN Status': 'Valid', 'TDS Section': '194C', 'Invoice / Bill No': 'BILL-001', 'Transaction Date': '10-06-2025', 'Gross Taxable Amount (₹)': 0, 'TDS Rate (%)': 2, 'TDS Deducted (₹)': 0, 'Challan BSR Code': '0210001', 'Challan Deposit Date': '07-07-2025', 'Challan Serial No': '00121', '26AS Matched (Y/N)': 'YES', 'Sec 206AB Non-Filer Status': 'Compliant' },
    ],
    sampleData: [
      { 'Deductee PAN': 'AAACV1234B', 'Deductee Trade Name': 'Verma & Associates Legal LLP', 'PAN Status': 'Valid', 'TDS Section': '194J', 'Invoice / Bill No': 'VA-2025-01', 'Transaction Date': '15-06-2025', 'Gross Taxable Amount (₹)': 350000, 'TDS Rate (%)': 10, 'TDS Deducted (₹)': 35000, 'Challan BSR Code': '0210452', 'Challan Deposit Date': '05-07-2025', 'Challan Serial No': '00014', '26AS Matched (Y/N)': 'YES', 'Sec 206AB Non-Filer Status': 'Compliant Filer' },
      { 'Deductee PAN': 'BBBCP5678K', 'Deductee Trade Name': 'Rajesh Sharma Civil Contractor', 'PAN Status': 'Valid', 'TDS Section': '194C', 'Invoice / Bill No': 'RS-889', 'Transaction Date': '20-06-2025', 'Gross Taxable Amount (₹)': 450000, 'TDS Rate (%)': 1, 'TDS Deducted (₹)': 4500, 'Challan BSR Code': '0210452', 'Challan Deposit Date': '05-07-2025', 'Challan Serial No': '00015', '26AS Matched (Y/N)': 'YES', 'Sec 206AB Non-Filer Status': 'Compliant Filer' },
      { 'Deductee PAN': 'APPLIED_FOR', 'Deductee Trade Name': 'Local Labour Supplier Unit', 'PAN Status': 'Invalid / Missing', 'TDS Section': '194C', 'Invoice / Bill No': 'LL-102', 'Transaction Date': '28-06-2025', 'Gross Taxable Amount (₹)': 100000, 'TDS Rate (%)': 20, 'TDS Deducted (₹)': 20000, 'Challan BSR Code': '0210452', 'Challan Deposit Date': '05-07-2025', 'Challan Serial No': '00016', '26AS Matched (Y/N)': 'YES', 'Sec 206AB Non-Filer Status': 'Sec 206AA 20% Applied' },
      { 'Deductee PAN': 'AAACT9900L', 'Deductee Trade Name': 'Titan Industrial Warehouse Ltd', 'PAN Status': 'Valid', 'TDS Section': '194I', 'Invoice / Bill No': 'RENT-Q1', 'Transaction Date': '30-06-2025', 'Gross Taxable Amount (₹)': 600000, 'TDS Rate (%)': 10, 'TDS Deducted (₹)': 60000, 'Challan BSR Code': '0210452', 'Challan Deposit Date': '05-07-2025', 'Challan Serial No': '00017', '26AS Matched (Y/N)': 'YES', 'Sec 206AB Non-Filer Status': 'Compliant Filer' },
      { 'Deductee PAN': 'TOTAL', 'Deductee Trade Name': 'Grand Total Deductions (₹)', 'PAN Status': '-', 'TDS Section': '-', 'Invoice / Bill No': '-', 'Transaction Date': '-', 'Gross Taxable Amount (₹)': 1500000, 'TDS Rate (%)': '-', 'TDS Deducted (₹)': 119500, 'Challan BSR Code': '-', 'Challan Deposit Date': '-', 'Challan Serial No': '-', '26AS Matched (Y/N)': '100% Matched', 'Sec 206AB Non-Filer Status': 'Form 26Q Q1 Ready' },
    ],
  },

  stock_verification: {
    id: 'stock_verification',
    name: 'Physical Stock Verification & AS-2 Valuation',
    category: 'Statutory Audit',
    statutoryStandard: 'ICAI AS-2 (Inventories) & CARO 2020 Cl. 3(ii)',
    description: 'Physical inventory count sheet with batch reconciliation, stock variance analysis, and lower-of-cost-or-NRV valuation.',
    sheetName: 'Stock_Verification_AS2',
    instructions: [
      'Record unique Item Code, Batch No, and Warehouse Location.',
      'Record count as per perpetual inventory records and actual physical count.',
      'Quantify physical variance (Shortage / Excess) and compute financial impact.',
      'Under AS-2 / Ind AS 2, inventory must be valued at the LOWER of Cost or Net Realisable Value (NRV).',
    ],
    columns: ['Item Code', 'Item Description / Grade', 'UOM', 'Warehouse Location', 'Book Quantity', 'Physical Quantity', 'Variance Quantity', 'Unit Cost (₹)', 'Estimated NRV (₹)', 'Valuation Basis', 'Total Valuation (₹)', 'Variance Impact (₹)', 'CARO 2020 Discrepancy >10% (Y/N)', 'Auditor Observation'],
    currencyColumns: ['Unit Cost (₹)', 'Estimated NRV (₹)', 'Total Valuation (₹)', 'Variance Impact (₹)'],
    columnWidths: {
      'Item Code': 14,
      'Item Description / Grade': 30,
      'UOM': 10,
      'Warehouse Location': 20,
      'Book Quantity': 16,
      'Physical Quantity': 16,
      'Variance Quantity': 16,
      'Unit Cost (₹)': 16,
      'Estimated NRV (₹)': 16,
      'Valuation Basis': 22,
      'Total Valuation (₹)': 20,
      'Variance Impact (₹)': 18,
      'CARO 2020 Discrepancy >10% (Y/N)': 26,
      'Auditor Observation': 34,
    },
    blankData: [
      { 'Item Code': 'RAW-001', 'Item Description / Grade': 'Cold Rolled Steel Coil 1.2mm', 'UOM': 'MT', 'Warehouse Location': 'Bay 1 / Bin A', 'Book Quantity': 0, 'Physical Quantity': 0, 'Variance Quantity': 0, 'Unit Cost (₹)': 0, 'Estimated NRV (₹)': 0, 'Valuation Basis': 'Lower of Cost or NRV', 'Total Valuation (₹)': 0, 'Variance Impact (₹)': 0, 'CARO 2020 Discrepancy >10% (Y/N)': 'NO', 'Auditor Observation': 'Physically counted and verified' },
    ],
    sampleData: [
      { 'Item Code': 'RAW-ST-01', 'Item Description / Grade': 'Cold Rolled Steel Sheets 1.2mm', 'UOM': 'MT', 'Warehouse Location': 'Warehouse 1 - Bay A', 'Book Quantity': 250, 'Physical Quantity': 250, 'Variance Quantity': 0, 'Unit Cost (₹)': 62000, 'Estimated NRV (₹)': 68000, 'Valuation Basis': 'At Cost (Cost < NRV)', 'Total Valuation (₹)': 15500000, 'Variance Impact (₹)': 0, 'CARO 2020 Discrepancy >10% (Y/N)': 'NO', 'Auditor Observation': 'Test count agreed with stock card' },
      { 'Item Code': 'RAW-AL-02', 'Item Description / Grade': 'Aluminum Extrusion Profiles 6063', 'UOM': 'KGS', 'Warehouse Location': 'Warehouse 1 - Bay B', 'Book Quantity': 15000, 'Physical Quantity': 14850, 'Variance Quantity': -150, 'Unit Cost (₹)': 240, 'Estimated NRV (₹)': 260, 'Valuation Basis': 'At Cost (Cost < NRV)', 'Total Valuation (₹)': 3564000, 'Variance Impact (₹)': -36000, 'CARO 2020 Discrepancy >10% (Y/N)': 'NO', 'Auditor Observation': '1% handling loss - within normal tolerance' },
      { 'Item Code': 'FG-BL-09', 'Item Description / Grade': 'Slow-Moving Hydraulic Valves', 'UOM': 'NOS', 'Warehouse Location': 'Finished Goods Bay C', 'Book Quantity': 400, 'Physical Quantity': 400, 'Variance Quantity': 0, 'Unit Cost (₹)': 4500, 'Estimated NRV (₹)': 3200, 'Valuation Basis': 'At NRV (NRV < Cost)', 'Total Valuation (₹)': 1280000, 'Variance Impact (₹)': 0, 'CARO 2020 Discrepancy >10% (Y/N)': 'NO', 'Auditor Observation': 'AS-2 write-down to NRV: ₹5,20,000 written off' },
      { 'Item Code': 'PKG-BX-04', 'Item Description / Grade': 'Corrugated Export Cartons', 'UOM': 'NOS', 'Warehouse Location': 'Packaging Store', 'Book Quantity': 12000, 'Physical Quantity': 10200, 'Variance Quantity': -1800, 'Unit Cost (₹)': 65, 'Estimated NRV (₹)': 70, 'Valuation Basis': 'At Cost (Cost < NRV)', 'Total Valuation (₹)': 663000, 'Variance Impact (₹)': -117000, 'CARO 2020 Discrepancy >10% (Y/N)': 'YES', 'Auditor Observation': '15% variance >10% limit - Reported under CARO 3(ii)' },
      { 'Item Code': 'TOTAL', 'Item Description / Grade': 'Grand Total Inventory Valuation (₹)', 'UOM': '-', 'Warehouse Location': '-', 'Book Quantity': 0, 'Physical Quantity': 0, 'Variance Quantity': 0, 'Unit Cost (₹)': 0, 'Estimated NRV (₹)': 0, 'Valuation Basis': '-', 'Total Valuation (₹)': 21007000, 'Variance Impact (₹)': -153000, 'CARO 2020 Discrepancy >10% (Y/N)': '-', 'Auditor Observation': 'Net Physical Variance: -₹1,53,000' },
    ],
  },

  expense_register: {
    id: 'expense_register',
    name: 'Petty Cash & Office Expense Register',
    category: 'Corporate Accounting',
    statutoryStandard: 'ICAI Guidance Note on ICFR',
    description: 'Office expense and petty cash voucher register tracking expense heads, GST input credits, and daily cash balances.',
    sheetName: 'Petty_Cash_Register',
    instructions: [
      'Record each petty cash/expense voucher with date, beneficiary, and approved head.',
      'Assign correct expense head for automated P&L grouping.',
      'Maintain running cash balance to avoid negative cash red flags in audit.',
      'Verify supporting cash memos/tax invoices for every disbursement exceeding ₹500.',
    ],
    columns: ['Voucher No', 'Payment Date', 'Paid To / Vendor', 'Expense Category', 'Gross Amount (₹)', 'GST Component (₹)', 'TDS Deducted (₹)', 'Net Paid (₹)', 'Approved By', 'Running Cash Balance (₹)', 'Supporting Bill Attached (Y/N)'],
    currencyColumns: ['Gross Amount (₹)', 'GST Component (₹)', 'TDS Deducted (₹)', 'Net Paid (₹)', 'Running Cash Balance (₹)'],
    dateColumns: ['Payment Date'],
    columnWidths: {
      'Voucher No': 16,
      'Payment Date': 14,
      'Paid To / Vendor': 28,
      'Expense Category': 24,
      'Gross Amount (₹)': 18,
      'GST Component (₹)': 18,
      'TDS Deducted (₹)': 16,
      'Net Paid (₹)': 18,
      'Approved By': 20,
      'Running Cash Balance (₹)': 24,
      'Supporting Bill Attached (Y/N)': 24,
    },
    blankData: [
      { 'Voucher No': 'VCH-001', 'Payment Date': '01-10-2026', 'Paid To / Vendor': 'Opening Cash Imprest', 'Expense Category': 'Cash Transfer', 'Gross Amount (₹)': 50000, 'GST Component (₹)': 0, 'TDS Deducted (₹)': 0, 'Net Paid (₹)': 50000, 'Approved By': 'CFO', 'Running Cash Balance (₹)': 50000, 'Supporting Bill Attached (Y/N)': 'YES' },
    ],
    sampleData: [
      { 'Voucher No': 'VCH-101', 'Payment Date': '01-10-2026', 'Paid To / Vendor': 'Opening Cash Imprest', 'Expense Category': 'Cash Transfer', 'Gross Amount (₹)': 50000, 'GST Component (₹)': 0, 'TDS Deducted (₹)': 0, 'Net Paid (₹)': 50000, 'Approved By': 'CFO', 'Running Cash Balance (₹)': 50000, 'Supporting Bill Attached (Y/N)': 'YES' },
      { 'Voucher No': 'VCH-102', 'Payment Date': '03-10-2026', 'Paid To / Vendor': 'Standard Stationery Mart', 'Expense Category': 'Printing & Stationery', 'Gross Amount (₹)': 4500, 'GST Component (₹)': 686, 'TDS Deducted (₹)': 0, 'Net Paid (₹)': 4500, 'Approved By': 'Admin Lead', 'Running Cash Balance (₹)': 45500, 'Supporting Bill Attached (Y/N)': 'YES' },
      { 'Voucher No': 'VCH-103', 'Payment Date': '05-10-2026', 'Paid To / Vendor': 'BlueDart Express Courier', 'Expense Category': 'Courier & Postage', 'Gross Amount (₹)': 1850, 'GST Component (₹)': 282, 'TDS Deducted (₹)': 0, 'Net Paid (₹)': 1850, 'Approved By': 'Admin Lead', 'Running Cash Balance (₹)': 43650, 'Supporting Bill Attached (Y/N)': 'YES' },
      { 'Voucher No': 'VCH-104', 'Payment Date': '08-10-2026', 'Paid To / Vendor': 'Quick Clean Facility Services', 'Expense Category': 'Office Maintenance', 'Gross Amount (₹)': 3200, 'GST Component (₹)': 488, 'TDS Deducted (₹)': 0, 'Net Paid (₹)': 3200, 'Approved By': 'HR Manager', 'Running Cash Balance (₹)': 40450, 'Supporting Bill Attached (Y/N)': 'YES' },
      { 'Voucher No': 'TOTAL', 'Payment Date': '-', 'Paid To / Vendor': 'Grand Total Disbursements (₹)', 'Expense Category': '-', 'Gross Amount (₹)': 9550, 'GST Component (₹)': 1456, 'TDS Deducted (₹)': 0, 'Net Paid (₹)': 9550, 'Approved By': '-', 'Running Cash Balance (₹)': 40450, 'Supporting Bill Attached (Y/N)': 'All Attached' },
    ],
  },

  general_ledger: {
    id: 'general_ledger',
    name: 'General Ledger Working Paper',
    category: 'Corporate Accounting',
    statutoryStandard: 'ICAI AS-1 & General Accounting Principles',
    description: 'Account-wise general ledger transactions tracking debits, credits, running balances, and working paper cross-references.',
    sheetName: 'General_Ledger',
    instructions: [
      'Chronological posting of all debits and credits for an account.',
      'Running balance calculates automatically: Previous Balance + Debit - Credit.',
      'Tick Reconciled column once verified against bank statement or supplier balance confirmation.',
      'Cross-reference audit working papers (e.g., WP-BNK-01, WP-PUR-02) for audit paper trail.',
    ],
    columns: ['Posting Date', 'Voucher Type', 'Voucher No', 'Cheque / Ref No', 'Particulars & Narration', 'Debit (₹)', 'Credit (₹)', 'Running Balance (₹)', 'Reconciled (Y/N)', 'Working Paper Cross-Ref'],
    currencyColumns: ['Debit (₹)', 'Credit (₹)', 'Running Balance (₹)'],
    dateColumns: ['Posting Date'],
    columnWidths: {
      'Posting Date': 14,
      'Voucher Type': 16,
      'Voucher No': 16,
      'Cheque / Ref No': 18,
      'Particulars & Narration': 42,
      'Debit (₹)': 18,
      'Credit (₹)': 18,
      'Running Balance (₹)': 20,
      'Reconciled (Y/N)': 16,
      'Working Paper Cross-Ref': 24,
    },
    blankData: [
      { 'Posting Date': '01-04-2025', 'Voucher Type': 'Journal', 'Voucher No': 'OB-01', 'Cheque / Ref No': '-', 'Particulars & Narration': 'Opening Balance B/F', 'Debit (₹)': 0, 'Credit (₹)': 0, 'Running Balance (₹)': 0, 'Reconciled (Y/N)': 'Y', 'Working Paper Cross-Ref': 'WP-TB-01' },
    ],
    sampleData: [
      { 'Posting Date': '01-09-2025', 'Voucher Type': 'Journal', 'Voucher No': 'OB-01', 'Cheque / Ref No': '-', 'Particulars & Narration': 'Opening Balance B/F from previous month', 'Debit (₹)': 450000, 'Credit (₹)': 0, 'Running Balance (₹)': 450000, 'Reconciled (Y/N)': 'Y', 'Working Paper Cross-Ref': 'WP-TB-01' },
      { 'Posting Date': '10-09-2025', 'Voucher Type': 'Payment', 'Voucher No': 'PV-401', 'Cheque / Ref No': 'NEFT-8821', 'Particulars & Narration': 'Payment to Apollo Tyres Ltd for Raw Materials', 'Debit (₹)': 0, 'Credit (₹)': 125000, 'Running Balance (₹)': 325000, 'Reconciled (Y/N)': 'Y', 'Working Paper Cross-Ref': 'WP-BNK-01' },
      { 'Posting Date': '18-09-2025', 'Voucher Type': 'Receipt', 'Voucher No': 'RV-202', 'Cheque / Ref No': 'RTGS-1102', 'Particulars & Narration': 'Receipt from Customer Apex Auto Pvt Ltd', 'Debit (₹)': 250000, 'Credit (₹)': 0, 'Running Balance (₹)': 575000, 'Reconciled (Y/N)': 'Y', 'Working Paper Cross-Ref': 'WP-BNK-01' },
      { 'Posting Date': '28-09-2025', 'Voucher Type': 'Payment', 'Voucher No': 'PV-422', 'Cheque / Ref No': 'CHQ-5511', 'Particulars & Narration': 'Quarterly Advance Tax Payment (AY 2026-27)', 'Debit (₹)': 0, 'Credit (₹)': 150000, 'Running Balance (₹)': 425000, 'Reconciled (Y/N)': 'Y', 'Working Paper Cross-Ref': 'WP-TAX-01' },
      { 'Posting Date': 'TOTAL', 'Voucher Type': '-', 'Voucher No': '-', 'Cheque / Ref No': '-', 'Particulars & Narration': 'Total Period Movements & Closing Balance (₹)', 'Debit (₹)': 700000, 'Credit (₹)': 275000, 'Running Balance (₹)': 425000, 'Reconciled (Y/N)': 'Y', 'Working Paper Cross-Ref': 'Agreed with Trial Balance' },
    ],
  },

  trial_balance: {
    id: 'trial_balance',
    name: 'Trial Balance & Schedule III Mapping',
    category: 'Corporate Accounting',
    statutoryStandard: 'Companies Act 2013 Schedule III (Div I & II)',
    description: 'Schedule III trial balance mapping ledger balances into Balance Sheet and P&L line items with audit adjustments.',
    sheetName: 'Trial_Balance_ScheduleIII',
    instructions: [
      'Enter opening balances, period movements, and closing balances from General Ledger.',
      'Classify each account under official Schedule III head (PPE, Trade Receivables, Borrowings, etc.).',
      'Ensure total Debits match total Credits with zero discrepancy.',
      'Record proposed audit adjustment entries in dedicated columns before drafting final balance sheet.',
    ],
    columns: ['Major Schedule', 'Account Code', 'Account Head / Description', 'Opening Dr (₹)', 'Opening Cr (₹)', 'Period Debit (₹)', 'Period Credit (₹)', 'Closing Dr (₹)', 'Closing Cr (₹)', 'Schedule III Mapping Head', 'Audit Adjustment Dr/Cr (₹)', 'Final Adjusted Balance (₹)'],
    currencyColumns: ['Opening Dr (₹)', 'Opening Cr (₹)', 'Period Debit (₹)', 'Period Credit (₹)', 'Closing Dr (₹)', 'Closing Cr (₹)', 'Audit Adjustment Dr/Cr (₹)', 'Final Adjusted Balance (₹)'],
    columnWidths: {
      'Major Schedule': 20,
      'Account Code': 14,
      'Account Head / Description': 34,
      'Opening Dr (₹)': 18,
      'Opening Cr (₹)': 18,
      'Period Debit (₹)': 18,
      'Period Credit (₹)': 18,
      'Closing Dr (₹)': 18,
      'Closing Cr (₹)': 18,
      'Schedule III Mapping Head': 28,
      'Audit Adjustment Dr/Cr (₹)': 24,
      'Final Adjusted Balance (₹)': 22,
    },
    blankData: [
      { 'Major Schedule': 'Assets', 'Account Code': '101', 'Account Head / Description': 'Cash & Bank Balances', 'Opening Dr (₹)': 0, 'Opening Cr (₹)': 0, 'Period Debit (₹)': 0, 'Period Credit (₹)': 0, 'Closing Dr (₹)': 0, 'Closing Cr (₹)': 0, 'Schedule III Mapping Head': 'Cash & Cash Equivalents', 'Audit Adjustment Dr/Cr (₹)': 0, 'Final Adjusted Balance (₹)': 0 },
    ],
    sampleData: [
      { 'Major Schedule': 'Fixed Assets', 'Account Code': '2010', 'Account Head / Description': 'Plant & Machinery - Cost Block', 'Opening Dr (₹)': 4500000, 'Opening Cr (₹)': 0, 'Period Debit (₹)': 450000, 'Period Credit (₹)': 0, 'Closing Dr (₹)': 4950000, 'Closing Cr (₹)': 0, 'Schedule III Mapping Head': 'Property, Plant & Equipment', 'Audit Adjustment Dr/Cr (₹)': 0, 'Final Adjusted Balance (₹)': 4950000 },
      { 'Major Schedule': 'Fixed Assets', 'Account Code': '2019', 'Account Head / Description': 'Accumulated Depreciation - P&M', 'Opening Dr (₹)': 0, 'Opening Cr (₹)': 1200000, 'Period Debit (₹)': 0, 'Period Credit (₹)': 450000, 'Closing Dr (₹)': 0, 'Closing Cr (₹)': 1650000, 'Schedule III Mapping Head': 'Deduction from PPE', 'Audit Adjustment Dr/Cr (₹)': 0, 'Final Adjusted Balance (₹)': -1650000 },
      { 'Major Schedule': 'Current Assets', 'Account Code': '3020', 'Account Head / Description': 'Trade Receivables (Sundry Debtors)', 'Opening Dr (₹)': 2200000, 'Opening Cr (₹)': 0, 'Period Debit (₹)': 8200000, 'Period Credit (₹)': 7950000, 'Closing Dr (₹)': 2450000, 'Closing Cr (₹)': 0, 'Schedule III Mapping Head': 'Trade Receivables', 'Audit Adjustment Dr/Cr (₹)': -50000, 'Final Adjusted Balance (₹)': 2400000 },
      { 'Major Schedule': 'Equity', 'Account Code': '1010', 'Account Head / Description': 'Equity Share Capital', 'Opening Dr (₹)': 0, 'Opening Cr (₹)': 5000000, 'Period Debit (₹)': 0, 'Period Credit (₹)': 0, 'Closing Dr (₹)': 0, 'Closing Cr (₹)': 5000000, 'Schedule III Mapping Head': 'Share Capital', 'Audit Adjustment Dr/Cr (₹)': 0, 'Final Adjusted Balance (₹)': -5000000 },
      { 'Major Schedule': 'Revenue', 'Account Code': '5001', 'Account Head / Description': 'Revenue from Operations (Gross)', 'Opening Dr (₹)': 0, 'Opening Cr (₹)': 0, 'Period Debit (₹)': 0, 'Period Credit (₹)': 12500000, 'Closing Dr (₹)': 0, 'Closing Cr (₹)': 12500000, 'Schedule III Mapping Head': 'Revenue from Operations', 'Audit Adjustment Dr/Cr (₹)': 0, 'Final Adjusted Balance (₹)': -12500000 },
      { 'Major Schedule': 'TOTAL', 'Account Code': '-', 'Account Head / Description': 'Grand Total Debits & Credits (₹)', 'Opening Dr (₹)': 6700000, 'Opening Cr (₹)': 6200000, 'Period Debit (₹)': 8650000, 'Period Credit (₹)': 20900000, 'Closing Dr (₹)': 7400000, 'Closing Cr (₹)': 19150000, 'Schedule III Mapping Head': 'Strict Mathematical Balance', 'Audit Adjustment Dr/Cr (₹)': -50000, 'Final Adjusted Balance (₹)': 0 },
    ],
  },

  audit_working_paper: {
    id: 'audit_working_paper',
    name: 'Statutory Audit Working Paper & Sampling Sheet',
    category: 'Statutory Audit',
    statutoryStandard: 'ICAI SA 230, SA 500 & SA 530 Standards',
    description: 'Statutory audit documentation covering materiality thresholds, 3-way matching verifications, and auditor conclusions.',
    sheetName: 'Audit_Working_Paper_SA500',
    instructions: [
      'Document benchmark planning materiality and performance materiality.',
      'Record sample voucher details, vendor names, and monetary values.',
      'Verify 3-way match: Purchase Order (PO) + Goods Receipt Note (GRN) + Tax Invoice.',
      'Verify statutory tax compliance (TDS deduction under 194C/J/Q and GST ITC eligibility).',
      'Record auditor conclusion and review sign-off with clear working paper cross-references.',
    ],
    columns: ['Working Paper Ref', 'Audit Section / Head', 'Sample Voucher No', 'Voucher Date', 'Party / Entity Name', 'Transaction Amount (₹)', 'Materiality Threshold (₹)', 'Audit Check: 3-Way Match', 'Audit Check: Tax Compliance', 'Audit Check: Authorization', 'Auditor Verification Finding', 'Audit Sign-off'],
    currencyColumns: ['Transaction Amount (₹)', 'Materiality Threshold (₹)'],
    dateColumns: ['Voucher Date'],
    columnWidths: {
      'Working Paper Ref': 18,
      'Audit Section / Head': 24,
      'Sample Voucher No': 18,
      'Voucher Date': 14,
      'Party / Entity Name': 28,
      'Transaction Amount (₹)': 22,
      'Materiality Threshold (₹)': 22,
      'Audit Check: 3-Way Match': 24,
      'Audit Check: Tax Compliance': 24,
      'Audit Check: Authorization': 22,
      'Auditor Verification Finding': 38,
      'Audit Sign-off': 20,
    },
    blankData: [
      { 'Working Paper Ref': 'WP-PUR-01', 'Audit Section / Head': 'Purchases Scrutiny', 'Sample Voucher No': 'PV-001', 'Voucher Date': '15-05-2025', 'Party / Entity Name': 'Sample Supplier Ltd', 'Transaction Amount (₹)': 0, 'Materiality Threshold (₹)': 100000, 'Audit Check: 3-Way Match': 'Verified', 'Audit Check: Tax Compliance': 'TDS & GST Compliant', 'Audit Check: Authorization': 'Approved by GM', 'Auditor Verification Finding': 'Satisfactory - Clean audit evidence', 'Audit Sign-off': 'Senior Review OK' },
    ],
    sampleData: [
      { 'Working Paper Ref': 'WP-CAP-01', 'Audit Section / Head': 'Capital Expenditure', 'Sample Voucher No': 'CAP-01', 'Voucher Date': '15-06-2025', 'Party / Entity Name': 'Tata Machinery Ltd', 'Transaction Amount (₹)': 1850000, 'Materiality Threshold (₹)': 500000, 'Audit Check: 3-Way Match': 'PO, GRN & Tax Invoice Verified', 'Audit Check: Tax Compliance': 'TDS Sec 194Q & GST Compliant', 'Audit Check: Authorization': 'Board Resolution Sighted', 'Auditor Verification Finding': 'Installation report and physical verification sighted. Capitalized properly.', 'Audit Sign-off': 'CA Senior Sign-off' },
      { 'Working Paper Ref': 'WP-LEG-02', 'Audit Section / Head': 'Legal & Professional', 'Sample Voucher No': 'PV-205', 'Voucher Date': '10-08-2025', 'Party / Entity Name': 'Verma & Associates Legal', 'Transaction Amount (₹)': 350000, 'Materiality Threshold (₹)': 100000, 'Audit Check: 3-Way Match': 'Engagement Letter Verified', 'Audit Check: Tax Compliance': 'TDS Sec 194J Deducted @ 10%', 'Audit Check: Authorization': 'Managing Director Signed', 'Auditor Verification Finding': 'Verified pending litigation disclosure for contingent liabilities in notes.', 'Audit Sign-off': 'CA Senior Sign-off' },
      { 'Working Paper Ref': 'WP-ADV-03', 'Audit Section / Head': 'Marketing & Brand Expense', 'Sample Voucher No': 'PV-312', 'Voucher Date': '14-11-2025', 'Party / Entity Name': 'Ogilvy Communications Ltd', 'Transaction Amount (₹)': 820000, 'Materiality Threshold (₹)': 200000, 'Audit Check: 3-Way Match': 'Campaign Proof Attached', 'Audit Check: Tax Compliance': 'TDS Sec 194C Deducted @ 2%', 'Audit Check: Authorization': 'VP Marketing Approved', 'Auditor Verification Finding': 'GSTR-2B reflected on portal. Certified for commercial substance.', 'Audit Sign-off': 'CA Senior Sign-off' },
      { 'Working Paper Ref': 'TOTAL', 'Audit Section / Head': 'Total Sample Tested (₹)', 'Sample Voucher No': '-', 'Voucher Date': '-', 'Party / Entity Name': 'Coverage: 68.4% of Target Ledger', 'Transaction Amount (₹)': 3020000, 'Materiality Threshold (₹)': '-', 'Audit Check: 3-Way Match': '100% Passed', 'Audit Check: Tax Compliance': '100% Passed', 'Audit Check: Authorization': '100% Passed', 'Auditor Verification Finding': 'No material audit misstatement detected. Unqualified opinion supported.', 'Audit Sign-off': 'Partner Concurrence' },
    ],
  },

  tax_invoice_register: {
    id: 'tax_invoice_register',
    name: 'Sales & Purchase Statutory Tax Invoice Register',
    category: 'Taxation & TDS',
    statutoryStandard: 'CGST Rules Rule 46 & Rule 48(4) (E-Invoice)',
    description: 'Statutory GST sales and purchase register capturing HSN/SAC codes, Place of Supply, tax splits, and IRN details.',
    sheetName: 'Tax_Invoice_Register',
    instructions: [
      'Maintain statutory invoice register compliant with Rule 46 of CGST Rules.',
      'Record mandatory 6 or 8 digit HSN/SAC code and Place of Supply (POS) for every B2B transaction.',
      'Separate intra-state (CGST + SGST) and inter-state (IGST) supplies.',
      'Capture 64-character Invoice Reference Number (IRN) and E-Way Bill number for turnover > ₹5 Crores.',
    ],
    columns: ['Invoice No', 'Invoice Date', 'Party Trade Name', 'Party GSTIN', 'Place of Supply (POS)', 'HSN / SAC Code', 'Taxable Amount (₹)', 'GST Rate (%)', 'CGST (₹)', 'SGST (₹)', 'IGST (₹)', 'Total Invoice Value (₹)', 'RCM Applicable', 'IRN Generated / Ack No', 'E-Way Bill No'],
    currencyColumns: ['Taxable Amount (₹)', 'CGST (₹)', 'SGST (₹)', 'IGST (₹)', 'Total Invoice Value (₹)'],
    dateColumns: ['Invoice Date'],
    columnWidths: {
      'Invoice No': 16,
      'Invoice Date': 14,
      'Party Trade Name': 28,
      'Party GSTIN': 18,
      'Place of Supply (POS)': 20,
      'HSN / SAC Code': 16,
      'Taxable Amount (₹)': 18,
      'GST Rate (%)': 14,
      'CGST (₹)': 16,
      'SGST (₹)': 16,
      'IGST (₹)': 16,
      'Total Invoice Value (₹)': 20,
      'RCM Applicable': 16,
      'IRN Generated / Ack No': 24,
      'E-Way Bill No': 20,
    },
    blankData: [
      { 'Invoice No': 'INV-001', 'Invoice Date': '01-10-2026', 'Party Trade Name': 'Sample Customer Pvt Ltd', 'Party GSTIN': '27AAACA0000A1Z5', 'Place of Supply (POS)': '27-Maharashtra', 'HSN / SAC Code': '84713010', 'Taxable Amount (₹)': 0, 'GST Rate (%)': 18, 'CGST (₹)': 0, 'SGST (₹)': 0, 'IGST (₹)': 0, 'Total Invoice Value (₹)': 0, 'RCM Applicable': 'No', 'IRN Generated / Ack No': 'IRN-PENDING', 'E-Way Bill No': '-' },
    ],
    sampleData: [
      { 'Invoice No': 'INV-1021', 'Invoice Date': '10-09-2025', 'Party Trade Name': 'Apex Electronics Private Limited', 'Party GSTIN': '27AAACA1234A1Z5', 'Place of Supply (POS)': '27-Maharashtra', 'HSN / SAC Code': '84713010', 'Taxable Amount (₹)': 450000, 'GST Rate (%)': 18, 'CGST (₹)': 40500, 'SGST (₹)': 40500, 'IGST (₹)': 0, 'Total Invoice Value (₹)': 531000, 'RCM Applicable': 'No', 'IRN Generated / Ack No': '1234567890abcdef...', 'E-Way Bill No': '221456789012' },
      { 'Invoice No': 'INV-1022', 'Invoice Date': '15-09-2025', 'Party Trade Name': 'Bangalore IT Logistics Ltd', 'Party GSTIN': '29BBBCB5678B2Z1', 'Place of Supply (POS)': '29-Karnataka', 'HSN / SAC Code': '998313', 'Taxable Amount (₹)': 620000, 'GST Rate (%)': 18, 'CGST (₹)': 0, 'SGST (₹)': 0, 'IGST (₹)': 111600, 'Total Invoice Value (₹)': 731600, 'RCM Applicable': 'No', 'IRN Generated / Ack No': '9876543210fedcba...', 'E-Way Bill No': '221456789034' },
      { 'Invoice No': 'INV-1023', 'Invoice Date': '22-09-2025', 'Party Trade Name': 'Gujarat Polymer Industries', 'Party GSTIN': '24DDDDD9988D1Z4', 'Place of Supply (POS)': '24-Gujarat', 'HSN / SAC Code': '390110', 'Taxable Amount (₹)': 890000, 'GST Rate (%)': 18, 'CGST (₹)': 0, 'SGST (₹)': 0, 'IGST (₹)': 160200, 'Total Invoice Value (₹)': 1050200, 'RCM Applicable': 'No', 'IRN Generated / Ack No': '4561237890abcdef...', 'E-Way Bill No': '221456789056' },
      { 'Invoice No': 'TOTAL', 'Invoice Date': '-', 'Party Trade Name': 'Grand Total Invoiced (₹)', 'Party GSTIN': '-', 'Place of Supply (POS)': '-', 'HSN / SAC Code': '-', 'Taxable Amount (₹)': 1960000, 'GST Rate (%)': '-', 'CGST (₹)': 40500, 'SGST (₹)': 40500, 'IGST (₹)': 271800, 'Total Invoice Value (₹)': 2312800, 'RCM Applicable': '-', 'IRN Generated / Ack No': 'All E-Invoiced', 'E-Way Bill No': 'All Valid' },
    ],
  },
};

export interface GenerateTemplateOptions {
  templateId: TemplateId;
  companyName: string;
  financialYear: string;
  includeSampleData: boolean;
  includeInstructionsSheet: boolean;
}

/**
 * Generates properly structured Excel workbook (.xlsx) for selected accounting template
 */
export function generateAccountingTemplateWorkbook(options: GenerateTemplateOptions): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();
  const tmpl = TEMPLATES[options.templateId];
  if (!tmpl) throw new Error(`Unknown template: ${options.templateId}`);

  // Sheet 1: Readme & Instructions Sheet (if enabled)
  if (options.includeInstructionsSheet) {
    const readmeData = [
      { Parameter: 'Accounting Working Paper Template', Value: tmpl.name },
      { Parameter: 'Statutory Standard & Compliance Benchmark', Value: tmpl.statutoryStandard },
      { Parameter: 'Category / Engagement Scope', Value: tmpl.category },
      { Parameter: 'Client Entity / Assessee Name', Value: options.companyName || 'Apex Industries Ltd' },
      { Parameter: 'Financial Year / Assessment Year', Value: options.financialYear || 'FY 2026-27' },
      { Parameter: 'Data Mode Selected', Value: options.includeSampleData ? 'Pre-filled Sample CA Practice Data' : 'Clean Formatted Blank Working Paper' },
      { Parameter: 'Software Compatibility Verified', Value: 'Microsoft Excel, Google Sheets, LibreOffice Calc, Apple Numbers' },
      { Parameter: 'Standard Operating Procedure 1', Value: tmpl.instructions[0] || 'Fill required accounting and audit data.' },
      { Parameter: 'Standard Operating Procedure 2', Value: tmpl.instructions[1] || 'Preserve standardized column headers to ensure macro and formula compatibility.' },
      { Parameter: 'Standard Operating Procedure 3', Value: tmpl.instructions[2] || 'Format numerical fields using standard Indian commas (Lakhs & Crores).' },
      { Parameter: 'Standard Operating Procedure 4', Value: tmpl.instructions[3] || 'Review and sign off with ICAI Working Paper reference before audit file archiving.' },
      { Parameter: 'Auditor Documentation Standard', Value: 'Engineered as per ICAI SA 230 (Audit Documentation) & SQC 1' },
      { Parameter: 'Generated On', Value: new Date().toLocaleString('en-IN') },
    ];

    const wsReadme = createStyledWorksheet(readmeData, {
      columnWidths: { Parameter: 34, Value: 56 },
    });
    XLSX.utils.book_append_sheet(wb, wsReadme, 'ICAI_Audit_Guidance');
  }

  // Sheet 2: Main Template Data Sheet
  const sourceData = options.includeSampleData ? tmpl.sampleData : tmpl.blankData;
  const wsMain = createStyledWorksheet(sourceData, {
    currencyColumns: tmpl.currencyColumns,
    dateColumns: tmpl.dateColumns,
    columnWidths: tmpl.columnWidths,
  });

  XLSX.utils.book_append_sheet(wb, wsMain, tmpl.sheetName);

  return wb;
}
