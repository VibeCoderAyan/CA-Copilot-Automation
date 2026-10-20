import * as XLSX from 'xlsx';
import { parseAmount, parseDate } from './dataSanitizer';
import { createStyledWorksheet } from './excelStyler';

export type ConfirmationType = 'Receivable' | 'Payable' | 'Bank Balance' | 'Loan / Borrowing';

export type ConfirmationStatus =
  | 'not_prepared'
  | 'ready_to_send'
  | 'sent'
  | 'response_received'
  | 'difference_reported'
  | 'followup_required'
  | 'closed';

export interface ConfirmationParty {
  id: string;
  partyName: string;
  accountCode: string;
  confirmationType: ConfirmationType;
  contactPerson: string;
  emailOrPhone: string;
  balanceDate: string;
  expectedBalance: number; // Books balance
  confirmedBalance?: number | null; // Third party confirmed balance
  differenceAmount?: number | null; // Calculated only if confirmedBalance exists
  status: ConfirmationStatus;
  sentDate?: string;
  responseDate?: string;
  followupDate?: string;
  isDirectResponse: boolean; // SA 505: Received directly from third party vs internal
  notes: string;
  isOverdueFollowup?: boolean;
  raw: Record<string, any>;
}

export interface ConfirmationSummary {
  clientName: string;
  balanceAsOnDate: string;
  totalParties: number;
  totalExpectedValue: number;
  totalBalance?: number;
  totalConfirmedValue: number;
  netDifferenceValue: number;
  readyToSendCount: number;
  sentCount: number;
  responseReceivedCount: number;
  differenceReportedCount: number;
  followupRequiredCount: number;
  closedCount: number;
  parties: ConfirmationParty[];
  summary?: ConfirmationSummary;
}

export interface ConfirmationColumnMapping {
  partyCol: string;
  codeCol: string;
  typeCol: string;
  contactCol: string;
  expectedBalCol: string;
  confirmedBalCol: string;
  statusCol: string;
  dateCol: string;
}

import { detectColumnCategory } from './columnDetector';

export function detectConfirmationColumns(headers: string[]): ConfirmationColumnMapping {
  const m: ConfirmationColumnMapping = {
    partyCol: '',
    codeCol: '',
    typeCol: '',
    contactCol: '',
    expectedBalCol: '',
    confirmedBalCol: '',
    statusCol: '',
    dateCol: '',
  };

  for (const h of headers) {
    const cat = detectColumnCategory(h);
    if (!m.partyCol && cat === 'partyName') m.partyCol = h;
    else if (!m.codeCol && cat === 'accountCode') m.codeCol = h;
    else if (!m.dateCol && cat === 'date') m.dateCol = h;
    else if (!m.expectedBalCol && (cat === 'balance' || cat === 'totalAmount')) m.expectedBalCol = h;
  }

  for (const h of headers) {
    const clean = h.toLowerCase().trim();
    if (!m.partyCol && /party|customer|vendor|bank|name|debtor|creditor/i.test(clean)) m.partyCol = h;
    else if (!m.codeCol && /code|account.*no|acc.*no|ledger.*code/i.test(clean)) m.codeCol = h;
    else if (!m.typeCol && /type|class|group|nature/i.test(clean)) m.typeCol = h;
    else if (!m.contactCol && /contact|email|phone|person/i.test(clean)) m.contactCol = h;
    else if (!m.expectedBalCol && /expected|book.*bal|balance|amount|closing/i.test(clean) && !/confirm/i.test(clean)) m.expectedBalCol = h;
    else if (!m.confirmedBalCol && /confirm.*bal|party.*bal|external/i.test(clean)) m.confirmedBalCol = h;
    else if (!m.statusCol && /status|stage|state/i.test(clean)) m.statusCol = h;
    else if (!m.dateCol && /date|as.*on/i.test(clean)) m.dateCol = h;
  }

  return m;
}

/**
 * Calculates confirmation tracker metrics and flags
 */
export function evaluateConfirmationSummary(
  parties: ConfirmationParty[],
  clientName = 'Corporate Client',
  balanceAsOnDate = '30-09-2026'
): ConfirmationSummary {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let totalExpected = 0;
  let totalConfirmed = 0;
  let netDiff = 0;
  let readyToSendCount = 0;
  let sentCount = 0;
  let responseReceivedCount = 0;
  let differenceReportedCount = 0;
  let followupRequiredCount = 0;
  let closedCount = 0;

  parties.forEach((p) => {
    totalExpected += p.expectedBalance;

    // Difference computation only if confirmed balance is entered
    if (p.confirmedBalance !== undefined && p.confirmedBalance !== null) {
      p.differenceAmount = p.expectedBalance - p.confirmedBalance;
      totalConfirmed += p.confirmedBalance;
      netDiff += Math.abs(p.differenceAmount);
    } else {
      p.differenceAmount = null;
    }

    // Overdue follow-up check
    let isOverdue = false;
    if (p.followupDate && p.status !== 'closed' && p.status !== 'response_received') {
      const [d, m, y] = p.followupDate.split('-').map(Number);
      const fDate = new Date(y, m - 1, d);
      if (fDate < today) isOverdue = true;
    }
    p.isOverdueFollowup = isOverdue;

    // Status tallies
    if (p.status === 'ready_to_send') readyToSendCount++;
    else if (p.status === 'sent') sentCount++;
    else if (p.status === 'response_received') responseReceivedCount++;
    else if (p.status === 'difference_reported') differenceReportedCount++;
    else if (p.status === 'followup_required') followupRequiredCount++;
    else if (p.status === 'closed') closedCount++;
  });

  const res: ConfirmationSummary = {
    clientName,
    balanceAsOnDate,
    totalParties: parties.length,
    totalExpectedValue: totalExpected,
    totalBalance: totalExpected,
    totalConfirmedValue: totalConfirmed,
    netDifferenceValue: netDiff,
    readyToSendCount,
    sentCount,
    responseReceivedCount,
    differenceReportedCount,
    followupRequiredCount,
    closedCount,
    parties,
  };
  res.summary = res;
  return res;
}

/**
 * Generates official SA 505 External Balance Confirmation Request Letter text
 */
export function generateConfirmationLetterText(
  party: ConfirmationParty,
  clientName: string,
  auditorFirmName = 'M/s S. K. Agarwal & Associates, Chartered Accountants',
  auditorEmail = 'audit@skagarwal.in'
): string {
  return `CONFIRMATION OF ACCOUNT BALANCE
(In accordance with Standard on Auditing 505 - External Confirmations)

Date: ${new Date().toLocaleDateString('en-IN')}

To,
${party.partyName}
${party.contactPerson ? `Attn: ${party.contactPerson}` : ''}
${party.emailOrPhone ? `Contact: ${party.emailOrPhone}` : ''}

Dear Sir / Madam,

Sub: Balance Confirmation as on ${party.balanceDate} for Statutory Audit

In connection with the statutory audit of our books of account for the period ended ${party.balanceDate}, our Statutory Auditors, ${auditorFirmName}, desire to verify the balance of your account.

According to our books of account, the amount ${party.confirmationType === 'Receivable' ? 'receivable from you' : party.confirmationType === 'Payable' ? 'payable to you' : 'balance'} as on ${party.balanceDate} was:

    INR ₹ ${party.expectedBalance.toLocaleString('en-IN')} 
    (${party.expectedBalance >= 0 ? 'Debit / Owed to us' : 'Credit / Advance'})

Please examine your records and either:
1. Confirm that the above balance is in agreement with your books, OR
2. If the balance does not agree, state the balance according to your books and provide a reconciliation statement showing invoices, payments, or credit notes in transit.

Please send your confirmation DIRECTLY to our auditors:
    ${auditorFirmName}
    Email: ${auditorEmail}

A stamped, addressed envelope / return confirmation slip is enclosed below for your convenience.

Thanking you,
Yours faithfully,
For ${clientName}

(Authorized Signatory)

-----------------------------------------------------------------------------------------
AUDIT CONFIRMATION SLIP (To be returned directly to the Auditors)

To: ${auditorFirmName}
Email: ${auditorEmail}

Ref: ${party.partyName} (Account Code: ${party.accountCode || 'N/A'})

(  ) The balance of ₹ ${party.expectedBalance.toLocaleString('en-IN')} as on ${party.balanceDate} is CORRECT.
(  ) The balance is NOT IN AGREEMENT. According to our books, the balance is ₹ ______________.
     (Reconciliation statement / statement of account attached).

Authorized Signature: _______________________
Name & Title: ______________________________
Date: ______________________________________
Company Stamp:`;
}

/**
 * Exports complete multi-sheet Excel balance confirmation workbook
 */
export function exportConfirmationWorkbook(summary: ConfirmationSummary): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Executive Summary
  const summaryRows = [
    { Parameter: 'Auditee Client Name', Value: summary.clientName },
    { Parameter: 'Balance Confirmation As On Date', Value: summary.balanceAsOnDate },
    { Parameter: 'Total Third Parties In Scope', Value: summary.totalParties },
    { Parameter: 'Total Expected Book Balance (₹)', Value: summary.totalExpectedValue },
    { Parameter: 'Total Third-Party Confirmed Balance (₹)', Value: summary.totalConfirmedValue },
    { Parameter: 'Net Reconciled Differences (₹)', Value: summary.netDifferenceValue },
    { Parameter: 'Confirmations Sent / In Flight', Value: summary.sentCount },
    { Parameter: 'Direct Responses Received', Value: summary.responseReceivedCount },
    { Parameter: 'Differences Reported', Value: summary.differenceReportedCount },
    { Parameter: 'Follow-ups Required', Value: summary.followupRequiredCount },
    { Parameter: 'Reconciled & Closed', Value: summary.closedCount },
    { Parameter: 'Standard on Auditing Reference', Value: 'SA 505 (External Confirmations) & CARO 2020' },
  ];
  const wsSummary = createStyledWorksheet(summaryRows, {
    currencyColumns: ['Value'],
    columnWidths: { Parameter: 38, Value: 34 },
  });
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');

  // Sheet 2: Master Confirmation Register
  const registerRows = summary.parties.map((p, idx) => ({
    '#': idx + 1,
    'Party / Account Name': p.partyName,
    'Account Code': p.accountCode || '-',
    'Category': p.confirmationType,
    'Contact Person': p.contactPerson || '-',
    'Email / Phone': p.emailOrPhone || '-',
    'Books Balance (₹)': p.expectedBalance,
    'Party Confirmed Balance (₹)': p.confirmedBalance ?? 'Awaiting Response',
    'Variance (₹)': p.differenceAmount ?? '-',
    'Current Status': p.status.replace(/_/g, ' ').toUpperCase(),
    'Sent Date': p.sentDate || '-',
    'Response Date': p.responseDate || '-',
    'Follow-up Date': p.followupDate || '-',
    'Response Type': p.isDirectResponse ? 'DIRECT TO AUDITOR' : 'INTERNAL FORWARD',
    'Auditor Notes / Discrepancy': p.notes || '-',
  }));
  const wsRegister = createStyledWorksheet(registerRows, {
    currencyColumns: ['Books Balance (₹)', 'Party Confirmed Balance (₹)', 'Variance (₹)'],
  });
  XLSX.utils.book_append_sheet(wb, wsRegister, 'Confirmation_Register');

  // Sheet 3: Differences & Exceptions
  const diffRows = summary.parties
    .filter((p) => p.status === 'difference_reported' || (p.differenceAmount !== null && p.differenceAmount !== 0 && p.differenceAmount !== undefined))
    .map((d, idx) => ({
      '#': idx + 1,
      'Party Name': d.partyName,
      'Type': d.confirmationType,
      'Books Balance (₹)': d.expectedBalance,
      'Party Confirmed Balance (₹)': d.confirmedBalance || 0,
      'Difference (₹)': d.differenceAmount || 0,
      'Response Date': d.responseDate || '-',
      'Reported Discrepancy': d.notes || 'In-transit cheques or billing differences',
      'Action Plan': 'Obtain itemized party ledger statement and pass year-end reconciliation JV',
    }));
  const wsDiff = createStyledWorksheet(diffRows, {
    currencyColumns: ['Books Balance (₹)', 'Party Confirmed Balance (₹)', 'Difference (₹)'],
  });
  XLSX.utils.book_append_sheet(wb, wsDiff, 'Differences_Report');

  return wb;
}

/**
 * Realistic party sample dataset for Balance Confirmations
 */
export function getSampleConfirmationParties(): ConfirmationParty[] {
  return [
    {
      id: 'conf-1',
      partyName: 'Reliance Retail Depot Ltd',
      accountCode: 'DR-1002',
      confirmationType: 'Receivable',
      contactPerson: 'Karan Mehra (Accounts Head)',
      emailOrPhone: 'karan.m@relianceretail.com',
      balanceDate: '30-09-2026',
      expectedBalance: 620000,
      confirmedBalance: 620000,
      differenceAmount: 0,
      status: 'response_received',
      sentDate: '01-10-2026',
      responseDate: '05-10-2026',
      isDirectResponse: true,
      notes: '100% matched. Direct confirmation letter received via auditor email.',
      raw: {},
    },
    {
      id: 'conf-2',
      partyName: 'Tata Steel Downstream Ltd',
      accountCode: 'CR-4001',
      confirmationType: 'Payable',
      contactPerson: 'Suresh Patil',
      emailOrPhone: 'spatil@tatasteel.com',
      balanceDate: '30-09-2026',
      expectedBalance: 450000,
      confirmedBalance: 485000,
      differenceAmount: -35000,
      status: 'difference_reported',
      sentDate: '01-10-2026',
      responseDate: '06-10-2026',
      isDirectResponse: true,
      notes: 'Difference of ₹35,000 reported. September transport freight invoice in transit.',
      raw: {},
    },
    {
      id: 'conf-3',
      partyName: 'HDFC Bank Ltd (Current A/c)',
      accountCode: 'BANK-01',
      confirmationType: 'Bank Balance',
      contactPerson: 'Branch Manager',
      emailOrPhone: 'operations@hdfcbank.com',
      balanceDate: '30-09-2026',
      expectedBalance: 2450000,
      confirmedBalance: 2450000,
      differenceAmount: 0,
      status: 'closed',
      sentDate: '01-10-2026',
      responseDate: '04-10-2026',
      isDirectResponse: true,
      notes: 'Standard bank confirmation certificate received with bank seal.',
      raw: {},
    },
    {
      id: 'conf-4',
      partyName: 'Godrej Consumer Goods',
      accountCode: 'DR-1005',
      confirmationType: 'Receivable',
      contactPerson: 'Rameshwar Vyas',
      emailOrPhone: 'rvyas@godrej.com',
      balanceDate: '30-09-2026',
      expectedBalance: 340000,
      confirmedBalance: null,
      status: 'sent',
      sentDate: '02-10-2026',
      followupDate: '09-10-2026',
      isDirectResponse: false,
      notes: 'First reminder sent on Oct 5. Awaiting formal confirmation.',
      raw: {},
    },
    {
      id: 'conf-5',
      partyName: 'Apex Industrial Packaging',
      accountCode: 'CR-4009',
      confirmationType: 'Payable',
      contactPerson: 'Amitabh Joshi',
      emailOrPhone: 'finance@apexpack.in',
      balanceDate: '30-09-2026',
      expectedBalance: 125000,
      confirmedBalance: null,
      status: 'followup_required',
      sentDate: '01-10-2026',
      followupDate: '04-10-2026', // Overdue follow-up!
      isDirectResponse: false,
      notes: 'Follow-up overdue by 2 days. Telephone reminder required.',
      raw: {},
    },
    {
      id: 'conf-6',
      partyName: 'State Bank of India (Term Loan)',
      accountCode: 'LOAN-01',
      confirmationType: 'Loan / Borrowing',
      contactPerson: 'Chief Manager, SME Branch',
      emailOrPhone: 'sme.branch@sbi.co.in',
      balanceDate: '30-09-2026',
      expectedBalance: 6500000,
      confirmedBalance: 6500000,
      differenceAmount: 0,
      status: 'response_received',
      sentDate: '01-10-2026',
      responseDate: '05-10-2026',
      isDirectResponse: true,
      notes: 'Term loan outstanding confirmed. Principal ₹65,00,000, accrued interest paid.',
      raw: {},
    },
    {
      id: 'conf-7',
      partyName: 'Mahindra Logistics Ltd',
      accountCode: 'CR-4015',
      confirmationType: 'Payable',
      contactPerson: 'Nitin Kulkarni',
      emailOrPhone: 'nitin.k@mahindra.com',
      balanceDate: '30-09-2026',
      expectedBalance: 85000,
      confirmedBalance: null,
      status: 'ready_to_send',
      isDirectResponse: false,
      notes: 'Draft letter prepared. Awaiting CA partner sign-off.',
      raw: {},
    },
  ];
}
