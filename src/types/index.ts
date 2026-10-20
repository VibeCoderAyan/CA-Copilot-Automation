// Type definitions for LedgerFlow CA Automation Platform

export type ActiveTab = 
  | 'overview'
  | 'dashboard'
  | 'automation'
  | 'reconciliation'
  | 'audit'
  | 'reports'
  | 'clients'
  | 'settings'
  | 'excel'
  | 'bank-rec'
  | 'gst-rec'
  | 'pdf-excel'
  | 'invoices'
  | 'deadlines'
  | 'history'
  | 'file-compare'
  | 'bank-converter'
  | 'formula-checker'
  | 'ledger-compare'
  | 'invoice-extractor'
  | 'bulk-renamer'
  | 'audit-checklist'
  | 'depreciation'
  | 'tb-checker'
  | 'excel-templates'
  | 'duplicate-invoice'
  | 'journal-validator'
  | 'receivables-aging'
  | 'far-checker'
  | 'audit-workpapers'
  | 'payroll-calculator'
  | 'ratio-analyzer'
  | 'voucher-verification'
  | 'tds-reconciliation'
  | 'cashflow-generator'
  | 'excel-data-cleanup'
  | 'purchase-invoice-matching'
  | 'month-end-closing'
  | 'audit-sampling'
  | 'balance-confirmation';

export interface ToastMessage {
  id: string;
  type: 'success' | 'error' | 'info';
  title: string;
  description?: string;
}

export interface BankTransaction {
  id: string;
  date: string;
  valueDate?: string;
  narration: string;
  refNo?: string;
  chqNo?: string;
  withdrawal: number;
  deposit: number;
  balance: number;
  matchStatus: 'Matched' | 'Probable' | 'Unmatched';
  matchedLedgerId?: string;
  difference?: number;
  suggestedScore?: number;
  [key: string]: any;
}

export interface LedgerTransaction {
  id: string;
  date: string;
  voucherNo: string;
  voucherType?: string;
  particulars: string;
  debit: number;
  credit: number;
  [key: string]: any;
}

export interface GstRecord {
  id: string;
  supplierGstin: string;
  supplierName: string;
  invoiceNumber: string;
  invoiceDate: string;
  invoiceType?: string;
  booksTaxable: number;
  booksTotalTax: number;
  gstr2bTaxable: number;
  gstr2bTotalTax: number;
  status: 'Matched' | 'Tax Mismatch' | 'Missing in 2B' | 'Missing in Books';
  itcEligibility: 'Eligible' | 'Ineligible' | 'Under Review' | 'Vendor Follow-up';
  diffAmount: number;
  [key: string]: any;
}

export interface InvoiceExtracted {
  id: string;
  fileName: string;
  supplierName: string;
  supplierGstin: string;
  invoiceNumber: string;
  invoiceDate: string;
  taxableAmount: number;
  cgst: number;
  sgst: number;
  igst: number;
  totalAmount: number;
  confidence: any;
  status: 'Approved' | 'Review Required' | 'Needs Review';
  notes?: string;
  [key: string]: any;
}

export interface TrialBalanceItem {
  accountCode: string;
  accountName: string;
  head: 'Asset' | 'Liability' | 'Equity' | 'Revenue' | 'Expense';
  subHead: string;
  debit: number;
  credit: number;
  previousYearDebit?: number;
  previousYearCredit?: number;
  [key: string]: any;
}

export interface ChatMessage {
  id: string;
  sender: 'user' | 'assistant';
  text: string;
  timestamp: string;
  actionSuggestion?: {
    label: string;
    action: () => void;
  };
  actionInvoked?: any;
  [key: string]: any;
}

export interface Client {
  id: string;
  name: string;
  pan: string;
  gstin: string;
  entityType?: string;
  [key: string]: any;
}

export type UserRole = 'Partner' | 'Senior' | 'Article' | 'Administrator' | 'Reviewer' | 'Auditor' | string;

export interface AuditLog {
  id: string;
  timestamp: string;
  user: string;
  action: string;
  details: string;
  [key: string]: any;
}

export interface AuditChecklistItem {
  id: string;
  [key: string]: any;
}

export interface DocumentRequestItem {
  id: string;
  [key: string]: any;
}

export interface UnusualTransaction {
  id: string;
  [key: string]: any;
}

export interface StatutoryDeadline {
  id: string;
  [key: string]: any;
}

export interface ActivityItem {
  id: string;
  [key: string]: any;
}
