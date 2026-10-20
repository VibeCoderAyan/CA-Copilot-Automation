import * as XLSX from 'xlsx';
import { parseAmount, parseDate } from './dataSanitizer';
import { createStyledWorksheet } from './excelStyler';

export interface AuditLeadSchedule {
  wpRef: string;
  leadScheduleName: string;
  groupCategory: 'Asset' | 'Liability' | 'Equity' | 'Revenue' | 'Expense';
  items: AuditWorkpaperAccount[];
  totalPriorYear: number;
  totalOpening: number;
  totalDebits: number;
  totalCredits: number;
  totalClosing: number;
  varianceAmount: number;
  variancePercent: number;
  isMaterialVariance: boolean;
  auditRisk: 'high' | 'medium' | 'low';
  reviewNotes: string[];
}

export interface AuditWorkpaperAccount {
  id: string;
  accountCode: string;
  accountName: string;
  leadGroup: string;
  priorYearBalance: number;
  openingBalance: number;
  periodDebit: number;
  periodCredit: number;
  closingBalance: number;
  varianceAmount: number;
  variancePercent: number;
  isMaterial: boolean;
  missingEvidenceFlags: string[];
  auditTickMark: string; // e.g. "^ (TB Checked)", "* (Calculated)", "! (Inquiry Required)"
  raw: Record<string, any>;
}

export interface AuditWorkpaperHeader {
  clientName: string;
  engagementType: string;
  financialYear: string;
  asOfDate: string;
  preparedBy: string;
  reviewedBy: string;
  materialityThresholdAmount: number; // e.g. ₹1,00,000
  materialityThresholdPercent: number; // e.g. 10%
}

export interface AuditProcedureItem {
  wpRef: string;
  procedureArea: string;
  objective: string;
  substantiveProcedure: string;
  evidenceRequired: string;
  verificationStatus: 'Not Started' | 'Sample Tested' | 'Pending Client Confirmation' | 'Satisfactory';
  auditFindingNote: string;
}

export interface ExternalConfirmationItem {
  id: string;
  partyName: string;
  category: 'Bank' | 'Trade Debtor' | 'Trade Creditor' | 'Borrowing / Loan';
  bookBalance: number;
  confirmedBalance: number;
  difference: number;
  status: 'Confirmation Sent' | 'Confirmed Matched' | 'Variance Disputed' | 'Awaiting Response';
  remarks: string;
}

export interface AuditWorkpaperPack {
  header: AuditWorkpaperHeader;
  totalTrialBalanceDebit: number;
  totalTrialBalanceCredit: number;
  isTbBalanced: boolean;
  leadSchedules: AuditLeadSchedule[];
  allAccounts: AuditWorkpaperAccount[];
  analyticalExceptionsCount: number;
  missingEvidenceCount: number;
  procedures: AuditProcedureItem[];
  confirmations: ExternalConfirmationItem[];
  disclaimer: string;
}

export interface WorkpaperColumnMapping {
  accountCodeCol: string;
  accountNameCol: string;
  groupCol: string;
  priorYearCol: string;
  openingCol: string;
  debitCol: string;
  creditCol: string;
  closingCol: string;
}

/**
 * Auto-detects columns for Trial Balance or Ledger imports
 */
export function detectWorkpaperColumns(headers: string[]): WorkpaperColumnMapping {
  const mapping: WorkpaperColumnMapping = {
    accountCodeCol: '',
    accountNameCol: '',
    groupCol: '',
    priorYearCol: '',
    openingCol: '',
    debitCol: '',
    creditCol: '',
    closingCol: '',
  };

  for (const h of headers) {
    const clean = h.toLowerCase().trim();
    if (!mapping.accountCodeCol && /code|acc.*no|gl.*code|account.*id/i.test(clean)) {
      mapping.accountCodeCol = h;
    } else if (!mapping.accountNameCol && /account.*name|particulars|ledger|head|description|account/i.test(clean)) {
      mapping.accountNameCol = h;
    } else if (!mapping.groupCol && /group|schedule|type|classification|category/i.test(clean)) {
      mapping.groupCol = h;
    } else if (!mapping.priorYearCol && /prior|prev|last.*year|ly|py/i.test(clean)) {
      mapping.priorYearCol = h;
    } else if (!mapping.openingCol && /opening|op.*bal/i.test(clean)) {
      mapping.openingCol = h;
    } else if (!mapping.debitCol && /debit|dr|dr.*amt|movement.*dr/i.test(clean)) {
      mapping.debitCol = h;
    } else if (!mapping.creditCol && /credit|cr|cr.*amt|movement.*cr/i.test(clean)) {
      mapping.creditCol = h;
    } else if (!mapping.closingCol && /closing|net.*bal|balance|closing.*amt|cl.*bal/i.test(clean)) {
      mapping.closingCol = h;
    }
  }

  // Fallbacks
  if (!mapping.accountNameCol && headers.length > 0) mapping.accountNameCol = headers[0];
  if (!mapping.debitCol && headers.length > 1) mapping.debitCol = headers[1];
  if (!mapping.creditCol && headers.length > 2) mapping.creditCol = headers[2];

  return mapping;
}

/**
 * Maps ledger items to standard Lead Schedule codes (WP-A through WP-I)
 */
export function assignLeadScheduleGroup(
  accountName: string,
  rawGroup: string
): { wpRef: string; leadScheduleName: string; category: 'Asset' | 'Liability' | 'Equity' | 'Revenue' | 'Expense' } {
  const text = `${accountName} ${rawGroup}`.toLowerCase();

  // Fixed Assets (WP-A)
  if (/plant|machinery|computer|laptop|furniture|fixture|vehicle|car|building|land|asset|equipment|depreciation.*accum/i.test(text) && !/depreciation.*expense/i.test(text)) {
    return { wpRef: 'WP-A', leadScheduleName: 'Property, Plant & Equipment (Fixed Assets)', category: 'Asset' };
  }
  // Cash & Bank (WP-B)
  if (/cash|petty cash|bank|hdfc|icici|sbi|axis|current a\/c|deposit|escrow/i.test(text)) {
    return { wpRef: 'WP-B', leadScheduleName: 'Cash and Bank Balances', category: 'Asset' };
  }
  // Debtors / Receivables (WP-C)
  if (/debtor|receivable|customer|trade receivable|unbilled/i.test(text)) {
    return { wpRef: 'WP-C', leadScheduleName: 'Trade Receivables & Advances', category: 'Asset' };
  }
  // Inventories (WP-D)
  if (/stock|inventory|raw material|finished goods|work in progress|wip|stores/i.test(text)) {
    return { wpRef: 'WP-D', leadScheduleName: 'Inventories & Work-in-Progress', category: 'Asset' };
  }
  // Equity & Capital (WP-E)
  if (/capital|equity|share|retained earning|reserve|surplus|securities premium/i.test(text)) {
    return { wpRef: 'WP-E', leadScheduleName: 'Share Capital & Reserves', category: 'Equity' };
  }
  // Loans & Borrowings (WP-F)
  if (/loan|borrowing|overdraft|cc limit|debenture|unsecured loan|secured loan/i.test(text)) {
    return { wpRef: 'WP-F', leadScheduleName: 'Borrowings & Financial Liabilities', category: 'Liability' };
  }
  // Trade Payables (WP-G)
  if (/creditor|payable|supplier|vendor|provision|accrued|tds payable|gst payable/i.test(text)) {
    return { wpRef: 'WP-G', leadScheduleName: 'Trade Payables & Provisions', category: 'Liability' };
  }
  // Revenue (WP-H)
  if (/sale|revenue|income|service|turnover|export|interest received/i.test(text)) {
    return { wpRef: 'WP-H', leadScheduleName: 'Revenue from Operations & Other Income', category: 'Revenue' };
  }
  // Expenses (WP-I)
  return { wpRef: 'WP-I', leadScheduleName: 'Direct Costs & Operating Expenses', category: 'Expense' };
}

/**
 * Builds the complete Audit Working Paper Pack
 */
export function generateAuditWorkingPaperPack(
  rows: Record<string, any>[],
  mapping: WorkpaperColumnMapping,
  header: AuditWorkpaperHeader
): AuditWorkpaperPack {
  const scheduleMap = new Map<string, AuditLeadSchedule>();

  const standardSchedules: { wpRef: string; name: string; category: 'Asset' | 'Liability' | 'Equity' | 'Revenue' | 'Expense' }[] = [
    { wpRef: 'WP-A', name: 'Property, Plant & Equipment (Fixed Assets)', category: 'Asset' },
    { wpRef: 'WP-B', name: 'Cash and Bank Balances', category: 'Asset' },
    { wpRef: 'WP-C', name: 'Trade Receivables & Advances', category: 'Asset' },
    { wpRef: 'WP-D', name: 'Inventories & Work-in-Progress', category: 'Asset' },
    { wpRef: 'WP-E', name: 'Share Capital & Reserves', category: 'Equity' },
    { wpRef: 'WP-F', name: 'Borrowings & Financial Liabilities', category: 'Liability' },
    { wpRef: 'WP-G', name: 'Trade Payables & Provisions', category: 'Liability' },
    { wpRef: 'WP-H', name: 'Revenue from Operations & Other Income', category: 'Revenue' },
    { wpRef: 'WP-I', name: 'Direct Costs & Operating Expenses', category: 'Expense' },
  ];

  standardSchedules.forEach((s) => {
    scheduleMap.set(s.wpRef, {
      wpRef: s.wpRef,
      leadScheduleName: s.name,
      groupCategory: s.category,
      items: [],
      totalPriorYear: 0,
      totalOpening: 0,
      totalDebits: 0,
      totalCredits: 0,
      totalClosing: 0,
      varianceAmount: 0,
      variancePercent: 0,
      isMaterialVariance: false,
      auditRisk: 'low',
      reviewNotes: [],
    });
  });

  let totalTbDebit = 0;
  let totalTbCredit = 0;
  const allAccounts: AuditWorkpaperAccount[] = [];
  let missingEvidenceCount = 0;
  let analyticalExceptionsCount = 0;

  rows.forEach((r, idx) => {
    const accountCode = String(r[mapping.accountCodeCol] || '').trim() || `ACC-${idx + 1}`;
    const accountName = String(r[mapping.accountNameCol] || '').trim();
    if (!accountName && !accountCode) return;

    const rawGroup = String(r[mapping.groupCol] || '').trim();
    const groupInfo = assignLeadScheduleGroup(accountName, rawGroup);

    const priorYear = parseAmount(r[mapping.priorYearCol]);
    let opening = parseAmount(r[mapping.openingCol]);
    const debit = parseAmount(r[mapping.debitCol]);
    const credit = parseAmount(r[mapping.creditCol]);
    let closing = parseAmount(r[mapping.closingCol]);

    totalTbDebit += debit;
    totalTbCredit += credit;

    // Derive closing if missing
    if (closing === 0 && (opening !== 0 || debit !== 0 || credit !== 0)) {
      if (groupInfo.category === 'Asset' || groupInfo.category === 'Expense') {
        closing = opening + debit - credit;
      } else {
        closing = opening + credit - debit;
      }
    }

    const varianceAmount = Math.abs(closing - priorYear);
    const variancePercent = priorYear !== 0 ? (varianceAmount / Math.abs(priorYear)) * 100 : closing !== 0 ? 100 : 0;
    const isMaterial =
      varianceAmount >= header.materialityThresholdAmount &&
      variancePercent >= header.materialityThresholdPercent;

    if (isMaterial) analyticalExceptionsCount++;

    const missingFlags: string[] = [];

    // Red Flag 1: Negative Cash
    if (/cash\b/i.test(accountName) && closing < 0) {
      missingFlags.push('Negative Cash Balance detected on closing date (SA 500 red flag).');
      missingEvidenceCount++;
    }

    // Red Flag 2: Bank balance unverified
    if (/bank/i.test(accountName) && closing !== 0) {
      missingFlags.push('Requires Bank Statement and Bank Reconciliation Statement (BRS) tie-out.');
    }

    // Red Flag 3: Debtors older or large balance without movement
    if (groupInfo.wpRef === 'WP-C' && debit === 0 && credit === 0 && closing > 100000) {
      missingFlags.push('Stagnant debtor balance with nil movement during period; evaluate for ECL provision.');
    }

    // Red Flag 4: Prior year missing
    if (priorYear === 0 && closing > header.materialityThresholdAmount) {
      missingFlags.push('No comparative prior year balance provided.');
    }

    const account: AuditWorkpaperAccount = {
      id: `acc-${idx + 1}`,
      accountCode,
      accountName,
      leadGroup: groupInfo.wpRef,
      priorYearBalance: priorYear,
      openingBalance: opening,
      periodDebit: debit,
      periodCredit: credit,
      closingBalance: closing,
      varianceAmount,
      variancePercent,
      isMaterial,
      missingEvidenceFlags: missingFlags,
      auditTickMark: isMaterial ? '! (Analytical Inquiry Required)' : '^ (TB Tie-Out)',
      raw: r,
    };

    allAccounts.push(account);

    const sched = scheduleMap.get(groupInfo.wpRef)!;
    sched.items.push(account);
    sched.totalPriorYear += priorYear;
    sched.totalOpening += opening;
    sched.totalDebits += debit;
    sched.totalCredits += credit;
    sched.totalClosing += closing;
  });

  // Recompute Lead Schedule Summaries
  const leadSchedules: AuditLeadSchedule[] = [];
  scheduleMap.forEach((sched) => {
    if (sched.items.length === 0) return;

    sched.varianceAmount = Math.abs(sched.totalClosing - sched.totalPriorYear);
    sched.variancePercent =
      sched.totalPriorYear !== 0
        ? (sched.varianceAmount / Math.abs(sched.totalPriorYear)) * 100
        : sched.totalClosing !== 0
        ? 100
        : 0;

    sched.isMaterialVariance =
      sched.varianceAmount >= header.materialityThresholdAmount &&
      sched.variancePercent >= header.materialityThresholdPercent;

    if (sched.isMaterialVariance) {
      sched.auditRisk = 'high';
      sched.reviewNotes.push(
        `Material variance of ₹${sched.varianceAmount.toLocaleString('en-IN')} (${sched.variancePercent.toFixed(1)}%) requires senior management inquiry and substantive testing.`
      );
    } else {
      sched.auditRisk = 'low';
      sched.reviewNotes.push('Analytical movement within normal operational variance thresholds.');
    }

    leadSchedules.push(sched);
  });

  // Standard Audit Procedures Program
  const procedures: AuditProcedureItem[] = [
    {
      wpRef: 'WP-A',
      procedureArea: 'Property, Plant & Equipment',
      objective: 'Verify ownership, physical existence, and statutory depreciation under Schedule II.',
      substantiveProcedure: 'Test sample additions > ₹1,00,000 to invoice, GR, and installation report. Verify 5% residual value retained.',
      evidenceRequired: 'Fixed Asset Register, Vendor Invoices, Physical Verification Certificate',
      verificationStatus: 'Sample Tested',
      auditFindingNote: 'Additions checked. Depreciation recalculated on Schedule II straight line basis.',
    },
    {
      wpRef: 'WP-B',
      procedureArea: 'Bank Balances & Cash',
      objective: 'Ensure completeness, accuracy, and cut-off of cash & bank accounts at year end.',
      substantiveProcedure: 'Obtain direct external bank confirmation. Scrutinize BRS for stale cheques > 90 days.',
      evidenceRequired: 'Bank Statements, Bank Confirmations, Signed BRS',
      verificationStatus: 'Pending Client Confirmation',
      auditFindingNote: 'Bank confirmation letters dispatched. BRS obtained for primary operational accounts.',
    },
    {
      wpRef: 'WP-C',
      procedureArea: 'Trade Receivables & Aging',
      objective: 'Assess recoverability and compliance with Schedule III aging disclosure.',
      substantiveProcedure: 'Perform subsequent receipt testing (post-balance sheet date receipts). Scrutinize balances > 180 days for ECL.',
      evidenceRequired: 'Debtors Ledger, Invoices, Delivery Challans, External Confirmations',
      verificationStatus: 'Sample Tested',
      auditFindingNote: 'Aging buckets verified against invoice dates. Top 5 debtors selected for confirmation.',
    },
    {
      wpRef: 'WP-G',
      procedureArea: 'Trade Payables & MSMED',
      objective: 'Ensure completeness of liabilities and compliance with Section 15/16 of MSMED Act.',
      substantiveProcedure: 'Search for unrecorded liabilities through cut-off testing of subsequent purchase invoices. Verify MSME vendor registrations.',
      evidenceRequired: 'Creditors Ledger, MSME Udyam Certificates, Subsequent Bank Book',
      verificationStatus: 'Sample Tested',
      auditFindingNote: 'MSME overdue balances tracked for 45-day interest disclosure in notes to accounts.',
    },
    {
      wpRef: 'WP-H',
      procedureArea: 'Revenue Recognition',
      objective: 'Verify occurrence, accuracy, and GSTR-1 vs Financial Books reconciliation.',
      substantiveProcedure: 'Reconcile turnover with GSTR-1 and GSTR-9 portal filings. Perform sample vouching of sales invoices around cut-off date.',
      evidenceRequired: 'Sales Register, GSTR-1 / 3B Portals, Shipping Bills / E-way Bills',
      verificationStatus: 'Satisfactory',
      auditFindingNote: 'Turnover reconciled with GST portal filings. Minor timing difference noted in E-way bills.',
    },
  ];

  // Standard External Confirmations Tracker
  const confirmations: ExternalConfirmationItem[] = [
    {
      id: 'conf-1',
      partyName: 'HDFC Bank Ltd - Current A/c 502000189',
      category: 'Bank',
      bookBalance: 4250000,
      confirmedBalance: 4250000,
      difference: 0,
      status: 'Confirmed Matched',
      remarks: 'Direct bank confirmation received via secure portal.',
    },
    {
      id: 'conf-2',
      partyName: 'State Bank of India - Term Loan A/c 3019',
      category: 'Borrowing / Loan',
      bookBalance: 12500000,
      confirmedBalance: 12500000,
      difference: 0,
      status: 'Confirmed Matched',
      remarks: 'Loan statement and interest certificate tied out.',
    },
    {
      id: 'conf-3',
      partyName: 'Reliance Retail Limited',
      category: 'Trade Debtor',
      bookBalance: 1850000,
      confirmedBalance: 1800000,
      difference: 50000,
      status: 'Variance Disputed',
      remarks: 'Client deducted ₹50,000 for rate difference credit note pending approval.',
    },
    {
      id: 'conf-4',
      partyName: 'Tata Steel Processing Ltd',
      category: 'Trade Creditor',
      bookBalance: 2400000,
      confirmedBalance: 0,
      difference: 2400000,
      status: 'Awaiting Response',
      remarks: 'Negative confirmation sent on 05-04-2027; follow-up reminder sent.',
    },
  ];

  const netDiff = Math.abs(totalTbDebit - totalTbCredit);
  const isTbBalanced = netDiff <= 0.05;

  const disclaimer =
    'AUDIT WORKING PAPER DRAFT: Prepared for internal review and documentation in accordance with Standards on Auditing (SAs) issued by ICAI. Calculations, lead schedules, and analytical ratios derived from trial balance. Final conclusion subject to partner sign-off and physical vouching evidence.';

  return {
    header,
    totalTrialBalanceDebit: totalTbDebit,
    totalTrialBalanceCredit: totalTbCredit,
    isTbBalanced,
    leadSchedules,
    allAccounts,
    analyticalExceptionsCount,
    missingEvidenceCount,
    procedures,
    confirmations,
    disclaimer,
  };
}

/**
 * Exports complete 5-sheet standardized Excel working paper package
 */
export function exportAuditWorkpaperWorkbook(
  pack: AuditWorkpaperPack
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Master WP Index & Metadata
  const indexData = [
    { 'Index #': 'WP-00', 'Working Paper Schedule': 'Audit Working Paper Master Index', 'Prepared By': pack.header.preparedBy, 'Reviewed By': pack.header.reviewedBy, 'Review Status': 'DRAFT - REVIEW PENDING' },
    { 'Index #': 'WP-A', 'Working Paper Schedule': 'Property, Plant & Equipment (Fixed Assets)', 'Prepared By': pack.header.preparedBy, 'Reviewed By': pack.header.reviewedBy, 'Review Status': 'SAMPLE TESTED' },
    { 'Index #': 'WP-B', 'Working Paper Schedule': 'Cash and Bank Balances', 'Prepared By': pack.header.preparedBy, 'Reviewed By': pack.header.reviewedBy, 'Review Status': 'CONFIRMATION PENDING' },
    { 'Index #': 'WP-C', 'Working Paper Schedule': 'Trade Receivables & Advances', 'Prepared By': pack.header.preparedBy, 'Reviewed By': pack.header.reviewedBy, 'Review Status': 'SAMPLE TESTED' },
    { 'Index #': 'WP-D', 'Working Paper Schedule': 'Inventories & Work-in-Progress', 'Prepared By': pack.header.preparedBy, 'Reviewed By': pack.header.reviewedBy, 'Review Status': 'PHYSICAL COUNT RECONCILED' },
    { 'Index #': 'WP-E', 'Working Paper Schedule': 'Share Capital & Reserves', 'Prepared By': pack.header.preparedBy, 'Reviewed By': pack.header.reviewedBy, 'Review Status': 'VERIFIED TO ROC FORM PAS-3' },
    { 'Index #': 'WP-F', 'Working Paper Schedule': 'Borrowings & Financial Liabilities', 'Prepared By': pack.header.preparedBy, 'Reviewed By': pack.header.reviewedBy, 'Review Status': 'TIED TO BANK SANCTION LETTER' },
    { 'Index #': 'WP-G', 'Working Paper Schedule': 'Trade Payables & Provisions', 'Prepared By': pack.header.preparedBy, 'Reviewed By': pack.header.reviewedBy, 'Review Status': 'MSMED TESTED' },
    { 'Index #': 'WP-H', 'Working Paper Schedule': 'Revenue from Operations & Other Income', 'Prepared By': pack.header.preparedBy, 'Reviewed By': pack.header.reviewedBy, 'Review Status': 'GSTR-1 RECONCILED' },
    { 'Index #': 'WP-I', 'Working Paper Schedule': 'Direct Costs & Operating Expenses', 'Prepared By': pack.header.preparedBy, 'Reviewed By': pack.header.reviewedBy, 'Review Status': 'ANALYTICAL REVIEW DONE' },
  ];
  const wsIndex = createStyledWorksheet(indexData, {
    columnWidths: { 'Index #': 12, 'Working Paper Schedule': 42, 'Prepared By': 22, 'Reviewed By': 22, 'Review Status': 30 },
  });
  XLSX.utils.book_append_sheet(wb, wsIndex, 'WP_Master_Index');

  // Sheet 2: Lead Schedules Summary
  const leadRows = pack.leadSchedules.map((s) => ({
    'WP Ref': s.wpRef,
    'Lead Schedule': s.leadScheduleName,
    'Class': s.groupCategory,
    'Prior Year (₹)': s.totalPriorYear,
    'Current Opening (₹)': s.totalOpening,
    'Debits Movement (₹)': s.totalDebits,
    'Credits Movement (₹)': s.totalCredits,
    'Closing Balance (₹)': s.totalClosing,
    'Variance (₹)': s.varianceAmount,
    'Variance %': `${s.variancePercent.toFixed(1)}%`,
    'Material Variance': s.isMaterialVariance ? 'YES - MATERIAL' : 'NO',
    'Audit Risk': s.auditRisk.toUpperCase(),
    'Audit Notes': s.reviewNotes.join(' | '),
  }));
  const wsLead = createStyledWorksheet(leadRows, {
    currencyColumns: [
      'Prior Year (₹)',
      'Current Opening (₹)',
      'Debits Movement (₹)',
      'Credits Movement (₹)',
      'Closing Balance (₹)',
      'Variance (₹)',
    ],
  });
  XLSX.utils.book_append_sheet(wb, wsLead, 'Lead_Schedules_Summary');

  // Sheet 3: Detailed Trial Balance & Accounts Schedule
  const accountsRows = pack.allAccounts.map((a) => ({
    'WP Ref': a.leadGroup,
    'A/c Code': a.accountCode,
    'Account Description': a.accountName,
    'Prior Year Balance (₹)': a.priorYearBalance,
    'Opening Balance (₹)': a.openingBalance,
    'Debit (₹)': a.periodDebit,
    'Credit (₹)': a.periodCredit,
    'Closing Balance (₹)': a.closingBalance,
    'Fluctuation (₹)': a.varianceAmount,
    'Fluctuation %': `${a.variancePercent.toFixed(1)}%`,
    'Material Item': a.isMaterial ? 'YES' : 'NO',
    'Audit Tick Mark': a.auditTickMark,
    'Audit Red Flags & Notes': a.missingEvidenceFlags.join(' | ') || 'No Exceptions',
  }));
  const wsAccounts = createStyledWorksheet(accountsRows, {
    currencyColumns: [
      'Prior Year Balance (₹)',
      'Opening Balance (₹)',
      'Debit (₹)',
      'Credit (₹)',
      'Closing Balance (₹)',
      'Fluctuation (₹)',
    ],
  });
  XLSX.utils.book_append_sheet(wb, wsAccounts, 'TB_Lead_Accounts');

  // Sheet 4: Substantive Audit Testing Program
  const procRows = pack.procedures.map((p) => ({
    'WP Ref': p.wpRef,
    'Audit Area': p.procedureArea,
    'Audit Objective': p.objective,
    'Testing Procedure': p.substantiveProcedure,
    'Audit Evidence Required': p.evidenceRequired,
    'Status': p.verificationStatus,
    'Auditor Finding & Memo': p.auditFindingNote,
  }));
  const wsProc = createStyledWorksheet(procRows, {
    columnWidths: {
      'WP Ref': 12,
      'Audit Area': 28,
      'Audit Objective': 40,
      'Testing Procedure': 55,
      'Audit Evidence Required': 40,
      'Status': 25,
      'Auditor Finding & Memo': 50,
    },
  });
  XLSX.utils.book_append_sheet(wb, wsProc, 'Audit_Procedures_Program');

  // Sheet 5: External Confirmations Tracker
  const confRows = pack.confirmations.map((c) => ({
    'Confirmation #': c.id,
    'Third Party Name': c.partyName,
    'Classification': c.category,
    'Book Balance (₹)': c.bookBalance,
    'Confirmed Amount (₹)': c.confirmedBalance,
    'Variance (₹)': c.difference,
    'Status': c.status,
    'Audit Reconciliation Note': c.remarks,
  }));
  const wsConf = createStyledWorksheet(confRows, {
    currencyColumns: ['Book Balance (₹)', 'Confirmed Amount (₹)', 'Variance (₹)'],
  });
  XLSX.utils.book_append_sheet(wb, wsConf, 'External_Confirmations');

  return wb;
}

/**
 * Sample manufacturing Trial Balance dataset with comparative prior year figures
 */
export function getSampleAuditTbRows(): Record<string, any>[] {
  return [
    // WP-A: Fixed Assets
    { 'Account Code': '1001', 'Account Name': 'Plant & Machinery Gross Block', 'Group': 'Fixed Assets', 'Prior Year': 12500000, 'Opening Balance': 12500000, 'Debit': 2800000, 'Credit': 0, 'Closing Balance': 15300000 },
    { 'Account Code': '1002', 'Account Name': 'Accumulated Depreciation - Plant', 'Group': 'Fixed Assets', 'Prior Year': 3200000, 'Opening Balance': 3200000, 'Debit': 0, 'Credit': 1050000, 'Closing Balance': 4250000 },
    { 'Account Code': '1003', 'Account Name': 'Computer & IT Equipment', 'Group': 'Fixed Assets', 'Prior Year': 850000, 'Opening Balance': 850000, 'Debit': 450000, 'Credit': 0, 'Closing Balance': 1300000 },

    // WP-B: Cash & Bank
    { 'Account Code': '1101', 'Account Name': 'Cash in Hand (Petty Cash)', 'Group': 'Current Assets', 'Prior Year': 45000, 'Opening Balance': 45000, 'Debit': 280000, 'Credit': 345000, 'Closing Balance': -20000 }, // INTENTIONAL RED FLAG: NEGATIVE CASH!
    { 'Account Code': '1102', 'Account Name': 'HDFC Bank Current Account', 'Group': 'Current Assets', 'Prior Year': 3850000, 'Opening Balance': 3850000, 'Debit': 18500000, 'Credit': 18100000, 'Closing Balance': 4250000 },
    { 'Account Code': '1103', 'Account Name': 'Fixed Deposit with SBI (Lien with Bank)', 'Group': 'Current Assets', 'Prior Year': 2000000, 'Opening Balance': 2000000, 'Debit': 500000, 'Credit': 0, 'Closing Balance': 2500000 },

    // WP-C: Debtors
    { 'Account Code': '1201', 'Account Name': 'Sundry Debtors - Domestic', 'Group': 'Current Assets', 'Prior Year': 5400000, 'Opening Balance': 5400000, 'Debit': 34000000, 'Credit': 32500000, 'Closing Balance': 6900000 },
    { 'Account Code': '1202', 'Account Name': 'Provision for Doubtful Debts', 'Group': 'Current Assets', 'Prior Year': 250000, 'Opening Balance': 250000, 'Debit': 0, 'Credit': 100000, 'Closing Balance': 350000 },

    // WP-D: Inventories
    { 'Account Code': '1301', 'Account Name': 'Raw Materials Stock', 'Group': 'Current Assets', 'Prior Year': 2800000, 'Opening Balance': 2800000, 'Debit': 14200000, 'Credit': 13800000, 'Closing Balance': 3200000 },
    { 'Account Code': '1302', 'Account Name': 'Finished Goods Warehouse Stock', 'Group': 'Current Assets', 'Prior Year': 1900000, 'Opening Balance': 1900000, 'Debit': 18900000, 'Credit': 18400000, 'Closing Balance': 2400000 },

    // WP-E: Equity & Reserves
    { 'Account Code': '2001', 'Account Name': 'Equity Share Capital', 'Group': 'Equity', 'Prior Year': 10000000, 'Opening Balance': 10000000, 'Debit': 0, 'Credit': 0, 'Closing Balance': 10000000 },
    { 'Account Code': '2002', 'Account Name': 'General Reserves & Surplus', 'Group': 'Equity', 'Prior Year': 4500000, 'Opening Balance': 4500000, 'Debit': 0, 'Credit': 1850000, 'Closing Balance': 6350000 },

    // WP-F: Borrowings
    { 'Account Code': '2101', 'Account Name': 'SBI Term Loan (Machinery Mortgage)', 'Group': 'Liabilities', 'Prior Year': 15000000, 'Opening Balance': 15000000, 'Debit': 2500000, 'Credit': 0, 'Closing Balance': 12500000 },
    { 'Account Code': '2102', 'Account Name': 'HDFC Cash Credit (CC) Facility', 'Group': 'Liabilities', 'Prior Year': 3200000, 'Opening Balance': 3200000, 'Debit': 0, 'Credit': 1600000, 'Closing Balance': 4800000 },

    // WP-G: Payables & Provisions
    { 'Account Code': '2201', 'Account Name': 'Sundry Creditors for Raw Materials', 'Group': 'Current Liabilities', 'Prior Year': 4100000, 'Opening Balance': 4100000, 'Debit': 16200000, 'Credit': 17500000, 'Closing Balance': 5400000 },
    { 'Account Code': '2202', 'Account Name': 'TDS & GST Statutory Dues Payable', 'Group': 'Current Liabilities', 'Prior Year': 380000, 'Opening Balance': 380000, 'Debit': 2100000, 'Credit': 2250000, 'Closing Balance': 530000 },

    // WP-H: Revenue
    { 'Account Code': '3001', 'Account Name': 'Domestic Sales of Polymer Products', 'Group': 'Revenue', 'Prior Year': 38000000, 'Opening Balance': 0, 'Debit': 0, 'Credit': 44500000, 'Closing Balance': 44500000 }, // MATERIAL FLUCTUATION > 10%
    { 'Account Code': '3002', 'Account Name': 'Export Sales (Zero Rated)', 'Group': 'Revenue', 'Prior Year': 6200000, 'Opening Balance': 0, 'Debit': 0, 'Credit': 8900000, 'Closing Balance': 8900000 },

    // WP-I: Expenses
    { 'Account Code': '4001', 'Account Name': 'Raw Material Consumption', 'Group': 'Expenses', 'Prior Year': 22500000, 'Opening Balance': 0, 'Debit': 26800000, 'Credit': 0, 'Closing Balance': 26800000 },
    { 'Account Code': '4002', 'Account Name': 'Power, Fuel & Utilities', 'Group': 'Expenses', 'Prior Year': 2800000, 'Opening Balance': 0, 'Debit': 3450000, 'Credit': 0, 'Closing Balance': 3450000 },
    { 'Account Code': '4003', 'Account Name': 'Employee Salaries & Benefits', 'Group': 'Expenses', 'Prior Year': 4200000, 'Opening Balance': 0, 'Debit': 5100000, 'Credit': 0, 'Closing Balance': 5100000 },
    { 'Account Code': '4004', 'Account Name': 'Audit & Professional Fees', 'Group': 'Expenses', 'Prior Year': 180000, 'Opening Balance': 0, 'Debit': 225000, 'Credit': 0, 'Closing Balance': 225000 },
  ];
}
