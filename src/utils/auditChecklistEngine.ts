import * as XLSX from 'xlsx';
import { createStyledWorksheet } from './excelStyler';

export type AuditType =
  | 'statutory_audit'
  | 'tax_audit_3cd'
  | 'gst_audit_9c'
  | 'ifc_review'
  | 'bank_concurrent';

export type ItemStatus = 'received' | 'pending' | 'na';
export type ItemPriority = 'mandatory' | 'recommended' | 'optional';

export interface AuditChecklistItem {
  id: string;
  code: string;
  category: string;
  particulars: string;
  requirement: string;
  priority: ItemPriority;
  status: ItemStatus;
  receivedDate?: string;
  documentRef?: string;
  assignedTo?: string;
  notes?: string;
}

export interface ClientAuditProject {
  id: string;
  clientName: string;
  financialYear: string;
  auditType: AuditType;
  engagementPartner: string;
  leadSenior: string;
  targetCompletionDate: string;
  items: AuditChecklistItem[];
  updatedAt: string;
}

// Pre-defined templates for Indian CA audit engagements
export const AUDIT_TEMPLATES: Record<
  AuditType,
  { name: string; description: string; defaultItems: Omit<AuditChecklistItem, 'id' | 'status'>[] }
> = {
  statutory_audit: {
    name: 'Statutory Audit (Companies Act 2013)',
    description: 'Comprehensive year-end audit checklist under Companies Act 2013 and CARO 2020.',
    defaultItems: [
      {
        code: 'SA-PPE-01',
        category: 'Property, Plant & Equipment',
        particulars: 'Fixed Asset Register (FAR)',
        requirement: 'Updated FAR with location, original cost, depreciation, and physical verification tag.',
        priority: 'mandatory',
      },
      {
        code: 'SA-PPE-02',
        category: 'Property, Plant & Equipment',
        particulars: 'Physical Verification Report',
        requirement: 'Management physical verification report of PPE and material discrepancies reconciliation.',
        priority: 'mandatory',
      },
      {
        code: 'SA-INV-01',
        category: 'Inventories',
        particulars: 'Year-End Stock Count Sheets',
        requirement: 'Inventory valuation report, slow-moving/obsolete inventory provision policy and sheets.',
        priority: 'mandatory',
      },
      {
        code: 'SA-BNK-01',
        category: 'Cash & Bank Balances',
        particulars: 'Direct Bank Balance Confirmations',
        requirement: 'Direct external bank balance confirmations for all operational, EEFC, and fixed deposit accounts.',
        priority: 'mandatory',
      },
      {
        code: 'SA-BNK-02',
        category: 'Cash & Bank Balances',
        particulars: 'Bank Reconciliation Statements (BRS)',
        requirement: 'BRS for all accounts as of year-end with subsequent clearing dates of unpresented cheques.',
        priority: 'mandatory',
      },
      {
        code: 'SA-REC-01',
        category: 'Trade Receivables',
        particulars: 'Debtors Ageing & External Confirmations',
        requirement: 'Ageing analysis (MSME and non-MSME) and sample direct balance confirmations from customers.',
        priority: 'mandatory',
      },
      {
        code: 'SA-CARO-01',
        category: 'CARO 2020 & Compliances',
        particulars: 'Working Capital Borrowings & Quarterly Statements',
        requirement: 'Quarterly stock statements submitted to banks vs books of account reconciliation (Clause 3(ii)(b)).',
        priority: 'mandatory',
      },
      {
        code: 'SA-CARO-02',
        category: 'CARO 2020 & Compliances',
        particulars: 'Title Deeds of Immovable Properties',
        requirement: 'Title deeds of all immovable property held in the name of the company (Clause 3(i)(c)).',
        priority: 'mandatory',
      },
      {
        code: 'SA-TRAIL-01',
        category: 'Statutory Compliances',
        particulars: 'Audit Trail (Edit Log) Compliance Report',
        requirement: 'Report on feature of recording audit trail enabled throughout the year without tampering.',
        priority: 'mandatory',
      },
      {
        code: 'SA-TAX-01',
        category: 'Taxation & Dues',
        particulars: 'Direct & Indirect Tax Assessment Orders',
        requirement: 'Copies of assessment orders, appeals filed, and disputed tax liabilities schedule.',
        priority: 'recommended',
      },
      {
        code: 'SA-REP-01',
        category: 'Management Representations',
        particulars: 'Management Representation Letter (MRL)',
        requirement: 'Signed MRL covering fraud, related parties, subsequent events, and going concern.',
        priority: 'mandatory',
      },
    ],
  },

  tax_audit_3cd: {
    name: 'Tax Audit (Form 3CD - Sec 44AB)',
    description: 'Mandatory clauses verification and supporting documentation under Income Tax Act.',
    defaultItems: [
      {
        code: 'TA-CL11-01',
        category: 'Books & Accounts',
        particulars: 'Clause 11: List of Books of Account',
        requirement: 'Complete list of books of account maintained and address of locations where books are kept.',
        priority: 'mandatory',
      },
      {
        code: 'TA-CL17-01',
        category: 'Capital Assets & Property',
        particulars: 'Clause 17: Sec 43CA / 50C Property Transfers',
        requirement: 'Details of land/building transferred during the year and stamp duty valuation comparison.',
        priority: 'recommended',
      },
      {
        code: 'TA-CL21-01',
        category: 'Expenses & Disallowances',
        particulars: 'Clause 21(a): Disallowance under Sec 40(a)',
        requirement: 'Details of TDS non-deduction or non-payment on expenses paid to residents and non-residents.',
        priority: 'mandatory',
      },
      {
        code: 'TA-CL21-02',
        category: 'Expenses & Disallowances',
        particulars: 'Clause 21(d): Sec 40A(3) Cash Payments > ₹10,000',
        requirement: 'Scrutiny of cash ledger and payments exceeding ₹10,000 made otherwise than by account payee cheque/ECS.',
        priority: 'mandatory',
      },
      {
        code: 'TA-CL26-01',
        category: 'Statutory Dues',
        particulars: 'Clause 26: Sec 43B Payments',
        requirement: 'Challans and proof of payment of GST, PF, ESI, Bonus, and interest paid on or before ITR due date.',
        priority: 'mandatory',
      },
      {
        code: 'TA-CL31-01',
        category: 'Loans & Deposits',
        particulars: 'Clause 31: Sec 269SS & 269T Cash Loans',
        requirement: 'Verification of loans, deposits, or specified sums accepted or repaid in cash exceeding ₹20,000.',
        priority: 'mandatory',
      },
      {
        code: 'TA-CL34-01',
        category: 'TDS / TCS Compliance',
        particulars: 'Clause 34: Quarterly TDS Returns & Reconciliations',
        requirement: 'Quarterly 24Q, 26Q, 27Q forms, Form 26AS matching, and interest on delayed deduction/deposit calculations.',
        priority: 'mandatory',
      },
      {
        code: 'TA-CL44-01',
        category: 'GST Expenditure',
        particulars: 'Clause 44: Break-up of Total Expenditure under GST',
        requirement: 'Break-up of expenditure relating to GST-registered entities vs exempt vs non-registered entities.',
        priority: 'mandatory',
      },
    ],
  },

  gst_audit_9c: {
    name: 'GST Annual Reconciliation (GSTR-9 & 9C)',
    description: 'Turnover, rate-wise outward supply, and input tax credit reconciliation.',
    defaultItems: [
      {
        code: 'GST-TO-01',
        category: 'Turnover Reconciliation',
        particulars: 'GSTR-1 vs GSTR-3B vs Audited Financials',
        requirement: 'Reconciliation of gross turnover as per audited P&L with aggregate turnover declared in GST returns.',
        priority: 'mandatory',
      },
      {
        code: 'GST-ITC-01',
        category: 'Input Tax Credit',
        particulars: 'GSTR-2B vs Purchase Register Reconciliation',
        requirement: 'Invoice-wise reconciliation of ITC availed in GSTR-3B with GSTR-2B auto-drafted statement.',
        priority: 'mandatory',
      },
      {
        code: 'GST-ITC-02',
        category: 'Input Tax Credit',
        particulars: 'Ineligible ITC Reversals (Rule 38, 42, 43, Sec 17(5))',
        requirement: 'Computation sheets of common ITC reversal for exempt supplies and blocked credits under Sec 17(5).',
        priority: 'mandatory',
      },
      {
        code: 'GST-RCM-01',
        category: 'Reverse Charge Mechanism',
        particulars: 'RCM Inward Supplies Liability (Sec 9(3) / 9(4))',
        requirement: 'GTA consignment notes, advocate fee invoices, security service payments, and RCM cash payment proof.',
        priority: 'mandatory',
      },
      {
        code: 'GST-EINV-01',
        category: 'E-Invoicing & E-Way Bills',
        particulars: 'IRN & QR Code Sample Verification',
        requirement: 'Sample test of B2B tax invoices for valid IRN generation from the IRP portal.',
        priority: 'recommended',
      },
    ],
  },

  ifc_review: {
    name: 'Internal Financial Controls (IFC / ICOFR)',
    description: 'Internal control testing across key business cycles and process flows.',
    defaultItems: [
      {
        code: 'IFC-P2P-01',
        category: 'Procure to Pay (P2P)',
        particulars: 'Three-Way Match Verification',
        requirement: 'Purchase Order (PO), Goods Receipt Note (GRN), and Vendor Tax Invoice three-way matching audit.',
        priority: 'mandatory',
      },
      {
        code: 'IFC-O2C-01',
        category: 'Order to Cash (O2C)',
        particulars: 'Customer Credit Limit Approval & Invoicing',
        requirement: 'Approval evidence for credit limits exceeding thresholds and authorization for billing rate changes.',
        priority: 'mandatory',
      },
      {
        code: 'IFC-H2R-01',
        category: 'Hire to Retire (Payroll)',
        particulars: 'Monthly Attendance, Overtime & Salary Register',
        requirement: 'Biometric attendance logs, HR salary sign-off, and statutory PF/ESI deduction authorization.',
        priority: 'mandatory',
      },
      {
        code: 'IFC-ITGC-01',
        category: 'IT General Controls',
        particulars: 'ERP User Access & Segregation of Duties (SoD)',
        requirement: 'ERP user access matrices review, generic admin login restrictions, and password change policies.',
        priority: 'mandatory',
      },
    ],
  },

  bank_concurrent: {
    name: 'Bank Concurrent Audit',
    description: 'Daily cash, credit appraisal, KYC, and NPA verification for bank branches.',
    defaultItems: [
      {
        code: 'BCA-CSH-01',
        category: 'Cash & Vault',
        particulars: 'Surprise Cash Verification & Overnight Retention',
        requirement: 'Physical verification of branch cash balance and verification against the insured retention limit.',
        priority: 'mandatory',
      },
      {
        code: 'BCA-KYC-01',
        category: 'Account Opening & KYC',
        particulars: 'High Value & Sensitive Accounts KYC',
        requirement: 'Verification of beneficial ownership, CKYC records, and PAN verification for new accounts.',
        priority: 'mandatory',
      },
      {
        code: 'BCA-LOAN-01',
        category: 'Credit Appraisal',
        particulars: 'Loan Sanction Terms & Security Creation',
        requirement: 'Verification of hypothecation deeds, ROC charge registration (Form CHG-1), and insurance policies.',
        priority: 'mandatory',
      },
      {
        code: 'BCA-DP-01',
        category: 'Working Capital Advances',
        particulars: 'Drawing Power (DP) Calculation & Stock Statements',
        requirement: 'Monthly stock and debtors statements scrutiny, paid stock deduction, and DP calculation review.',
        priority: 'mandatory',
      },
      {
        code: 'BCA-NPA-01',
        category: 'Asset Classification',
        particulars: 'SMA-0, SMA-1, SMA-2 & Potential NPA Identification',
        requirement: 'Review of accounts with interest/installment overdue > 60 days and compliance with IRAC norms.',
        priority: 'mandatory',
      },
    ],
  },
};

/**
 * Creates default items for a selected audit type
 */
export function createDefaultChecklistItems(auditType: AuditType): AuditChecklistItem[] {
  const template = AUDIT_TEMPLATES[auditType];
  return template.defaultItems.map((item, index) => ({
    ...item,
    id: `item-${auditType}-${index + 1}`,
    status: 'pending',
    notes: '',
    documentRef: '',
    receivedDate: undefined,
  }));
}

/**
 * Generates initial project template (returns empty array so zero fake data is loaded by default)
 */
export function getInitialClientProjects(): ClientAuditProject[] {
  return [];
}

/**
 * Calculates completion metrics
 */
export function calculateChecklistMetrics(items: AuditChecklistItem[]) {
  const total = items.length;
  const received = items.filter((i) => i.status === 'received').length;
  const pending = items.filter((i) => i.status === 'pending').length;
  const na = items.filter((i) => i.status === 'na').length;
  const applicable = total - na;
  const percentage = applicable > 0 ? Math.round((received / applicable) * 100) : 0;

  return { total, received, pending, na, percentage };
}

/**
 * Generates formatted multi-sheet Excel audit report
 */
export function exportAuditChecklistWorkbook(project: ClientAuditProject): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();
  const metrics = calculateChecklistMetrics(project.items);

  // Sheet 1: Audit Engagement Summary
  const summaryData = [
    { Parameter: 'Client Legal Entity', Value: project.clientName },
    { Parameter: 'Financial Year / Period', Value: project.financialYear },
    { Parameter: 'Audit Engagement Type', Value: AUDIT_TEMPLATES[project.auditType]?.name || project.auditType },
    { Parameter: 'Engagement Partner', Value: project.engagementPartner },
    { Parameter: 'Lead Audit Senior', Value: project.leadSenior },
    { Parameter: 'Target Completion Date', Value: project.targetCompletionDate },
    { Parameter: 'Total Checklist Documents', Value: metrics.total },
    { Parameter: 'Documents Received & Verified', Value: metrics.received },
    { Parameter: 'Pending Client Requirements', Value: metrics.pending },
    { Parameter: 'Not Applicable Items', Value: metrics.na },
    { Parameter: 'Audit Document Completion Rate', Value: `${metrics.percentage}%` },
    { Parameter: 'Report Generation Timestamp', Value: new Date().toLocaleString('en-IN') },
  ];

  const wsSummary = createStyledWorksheet(summaryData, {
    columnWidths: { Parameter: 35, Value: 45 },
  });
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Audit Engagement Summary');

  // Sheet 2: Full Audit Checklist Working Paper
  const fullData = project.items.map((item, idx) => ({
    'Sl No': idx + 1,
    'Item Code': item.code,
    'Audit Section / Head': item.category,
    'Particulars': item.particulars,
    'Detailed Audit Requirement': item.requirement,
    'Priority': item.priority.toUpperCase(),
    'Status': item.status === 'received' ? 'RECEIVED' : item.status === 'pending' ? 'PENDING' : 'NOT APPLICABLE',
    'Received Date': item.receivedDate || '-',
    'Working Paper Ref': item.documentRef || '-',
    'Assigned Team Member': item.assignedTo || '-',
    'Audit Notes & Observations': item.notes || '-',
  }));

  const wsFull = createStyledWorksheet(fullData, {
    columnWidths: {
      'Sl No': 8,
      'Item Code': 14,
      'Audit Section / Head': 24,
      'Particulars': 30,
      'Detailed Audit Requirement': 45,
      'Priority': 14,
      'Status': 16,
      'Received Date': 14,
      'Working Paper Ref': 20,
      'Assigned Team Member': 22,
      'Audit Notes & Observations': 35,
    },
  });
  XLSX.utils.book_append_sheet(wb, wsFull, 'Full Audit Checklist');

  // Sheet 3: Formal Pending Document Requisition for Client
  const pendingItems = project.items.filter((i) => i.status === 'pending');
  const pendingData = pendingItems.map((item, idx) => ({
    'Requirement No': idx + 1,
    'Reference Code': item.code,
    'Audit Category': item.category,
    'Document / Record Description': item.particulars,
    'Statutory / Audit Purpose': item.requirement,
    'Urgency': item.priority === 'mandatory' ? 'CRITICAL / MANDATORY' : 'STANDARD',
    'Auditor Remarks': item.notes || 'Please provide copy with management sign-off.',
    'Client Reply / Annexure': '',
  }));

  const wsPending = createStyledWorksheet(pendingData, {
    columnWidths: {
      'Requirement No': 14,
      'Reference Code': 16,
      'Audit Category': 22,
      'Document / Record Description': 32,
      'Statutory / Audit Purpose': 45,
      'Urgency': 22,
      'Auditor Remarks': 35,
      'Client Reply / Annexure': 25,
    },
  });
  XLSX.utils.book_append_sheet(wb, wsPending, 'Pending Client Requisition');

  return wb;
}
