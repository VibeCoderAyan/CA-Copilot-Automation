import * as XLSX from 'xlsx';
import { parseAmount } from './dataSanitizer';
import { createStyledWorksheet } from './excelStyler';

export interface CashFlowLineItem {
  id: string;
  category: 'operating' | 'investing' | 'financing' | 'cash_reconciliation';
  description: string;
  cyAmount: number;
  pyAmount: number;
  sourceAccounts: string;
  isSubtotal?: boolean;
  isTotal?: boolean;
  notes?: string;
}

export interface CashFlowStatementResult {
  clientName: string;
  cyPeriod: string;
  pyPeriod: string;
  operatingItems: CashFlowLineItem[];
  netOperatingCash: number;
  investingItems: CashFlowLineItem[];
  netInvestingCash: number;
  financingItems: CashFlowLineItem[];
  netFinancingCash: number;
  netChangeInCash: number;
  openingCash: number;
  computedClosingCash: number;
  actualClosingCash: number;
  reconciliationDifference: number;
  isReconciled: boolean;
  unmappedAccounts: string[];
}

export interface CashFlowColumnMapping {
  accountNameCol: string;
  categoryCol: string;
  cyAmountCol: string;
  pyAmountCol: string;
}

/**
 * Auto-detects columns for Cash Flow Statement Generator
 */
export function detectCashFlowColumns(headers: string[]): CashFlowColumnMapping {
  const mapping: CashFlowColumnMapping = {
    accountNameCol: '',
    categoryCol: '',
    cyAmountCol: '',
    pyAmountCol: '',
  };

  for (const h of headers) {
    const clean = h.toLowerCase().trim();
    if (!mapping.accountNameCol && /particulars|account|head|line.*item|description/i.test(clean)) {
      mapping.accountNameCol = h;
    } else if (!mapping.categoryCol && /category|group|type|schedule|class/i.test(clean)) {
      mapping.categoryCol = h;
    } else if (!mapping.cyAmountCol && /current|cy|fy.*27|fy.*26-27|amount|2027|2026/i.test(clean) && !/prev|prior|ly/i.test(clean)) {
      mapping.cyAmountCol = h;
    } else if (!mapping.pyAmountCol && /prev|prior|py|ly|fy.*26|fy.*25-26|last.*year/i.test(clean)) {
      mapping.pyAmountCol = h;
    }
  }

  // Fallbacks
  if (!mapping.accountNameCol && headers.length > 0) mapping.accountNameCol = headers[0];
  if (!mapping.cyAmountCol && headers.length > 1) mapping.cyAmountCol = headers[1];
  if (!mapping.pyAmountCol && headers.length > 2) mapping.pyAmountCol = headers[2];

  return mapping;
}

/**
 * Builds standard AS-3 / Ind AS 7 Indirect Method Cash Flow Statement
 */
export function generateCashFlowStatement(
  rows: Record<string, any>[],
  mapping: CashFlowColumnMapping,
  clientName = 'Corporate Client',
  cyPeriod = 'FY 2026-27',
  pyPeriod = 'FY 2025-26'
): CashFlowStatementResult {
  let netProfitBeforeTax = 0;
  let depreciation = 0;
  let financeCosts = 0;
  let interestIncome = 0;
  let taxesPaid = 0;

  // Working Capital Balances
  let cyDebtors = 0;
  let pyDebtors = 0;
  let cyInventory = 0;
  let pyInventory = 0;
  let cyOtherCurrentAssets = 0;
  let pyOtherCurrentAssets = 0;
  let cyCreditors = 0;
  let pyCreditors = 0;
  let cyOtherCurrentLiabilities = 0;
  let pyOtherCurrentLiabilities = 0;

  // Investing Balances
  let cyFixedAssetsGross = 0;
  let pyFixedAssetsGross = 0;
  let proceedsSaleAssets = 0;

  // Financing Balances
  let cyShareCapital = 0;
  let pyShareCapital = 0;
  let cyBorrowings = 0;
  let pyBorrowings = 0;
  let dividendPaid = 0;

  // Cash Equivalents
  let cyCash = 0;
  let pyCash = 0;

  const safeRows = Array.isArray(rows) ? rows : [];
  const safeMapping: CashFlowColumnMapping = mapping || {
    accountNameCol: '',
    categoryCol: '',
    cyAmountCol: '',
    pyAmountCol: '',
  };

  const unmapped: string[] = [];

  safeRows.forEach((r) => {
    if (!r || typeof r !== 'object') return;
    const name = String(r[safeMapping.accountNameCol] || '').trim();
    const cat = String(r[safeMapping.categoryCol] || '').trim();
    const text = `${name} ${cat}`.toLowerCase();
    const cyVal = parseAmount(r[safeMapping.cyAmountCol]);
    const pyVal = parseAmount(r[safeMapping.pyAmountCol]);

    if (!name && cyVal === 0 && pyVal === 0) return;

    if (/profit.*before.*tax|pbt/i.test(text)) {
      netProfitBeforeTax = cyVal;
    } else if (/net.*profit.*after.*tax|pat/i.test(text) && netProfitBeforeTax === 0) {
      netProfitBeforeTax = cyVal;
    } else if (/depreciation|amortisation|amortization/i.test(text)) {
      depreciation += cyVal;
    } else if (/finance.*cost|interest.*expense|interest.*paid/i.test(text)) {
      financeCosts += cyVal;
    } else if (/interest.*income|interest.*received|dividend.*income/i.test(text)) {
      interestIncome += cyVal;
    } else if (/income.*tax.*paid|taxes.*paid|current.*tax/i.test(text)) {
      taxesPaid += cyVal;
    } else if (/debtor|receivable|trade.*receivable/i.test(text)) {
      cyDebtors += cyVal;
      pyDebtors += pyVal;
    } else if (/inventory|stock|finished.*goods|raw.*material.*stock/i.test(text)) {
      cyInventory += cyVal;
      pyInventory += pyVal;
    } else if (/other.*current.*asset|loans.*advances|prepaid/i.test(text)) {
      cyOtherCurrentAssets += cyVal;
      pyOtherCurrentAssets += pyVal;
    } else if (/creditor|payable|trade.*payable/i.test(text)) {
      cyCreditors += cyVal;
      pyCreditors += pyVal;
    } else if (/other.*current.*liability|provisions|accrued/i.test(text)) {
      cyOtherCurrentLiabilities += cyVal;
      pyOtherCurrentLiabilities += pyVal;
    } else if (/property.*plant|fixed.*asset|machinery|building|gross.*block/i.test(text)) {
      cyFixedAssetsGross += cyVal;
      pyFixedAssetsGross += pyVal;
    } else if (/sale.*asset|proceed.*asset/i.test(text)) {
      proceedsSaleAssets += cyVal;
    } else if (/share.*capital|equity.*capital|securities.*premium/i.test(text)) {
      cyShareCapital += cyVal;
      pyShareCapital += pyVal;
    } else if (/borrowing|term.*loan|overdraft|debenture|loan/i.test(text)) {
      cyBorrowings += cyVal;
      pyBorrowings += pyVal;
    } else if (/dividend.*paid|dividend/i.test(text)) {
      dividendPaid += cyVal;
    } else if (/cash|bank.*balance|cash.*equivalent/i.test(text)) {
      cyCash += cyVal;
      pyCash += pyVal;
    } else {
      unmapped.push(name);
    }
  });

  // If net profit before tax was not explicitly mapped, check if profit after tax was mapped
  if (netProfitBeforeTax === 0 && rows.length > 0) {
    // Retain 0 and notify auditor in unmapped accounts
    unmapped.push('Profit Before Tax (PBT) not identified; verify Statement of Profit & Loss mapping');
  }

  // 1. Operating Activities
  const operatingItems: CashFlowLineItem[] = [];

  operatingItems.push({
    id: 'op-1',
    category: 'operating',
    description: 'Net Profit Before Tax (from Statement of Profit & Loss)',
    cyAmount: netProfitBeforeTax,
    pyAmount: 0,
    sourceAccounts: 'P&L Statement - Profit Before Tax',
  });

  operatingItems.push({
    id: 'op-2',
    category: 'operating',
    description: 'Adjustment for: Depreciation and Amortisation Expense',
    cyAmount: depreciation,
    pyAmount: 0,
    sourceAccounts: 'Depreciation Schedule (Schedule II)',
    notes: 'Non-cash charge added back',
  });

  operatingItems.push({
    id: 'op-3',
    category: 'operating',
    description: 'Adjustment for: Finance Costs / Interest Expense',
    cyAmount: financeCosts,
    pyAmount: 0,
    sourceAccounts: 'Finance Costs Schedule',
    notes: 'Financing item added back to operating flow',
  });

  if (interestIncome > 0) {
    operatingItems.push({
      id: 'op-4',
      category: 'operating',
      description: 'Adjustment for: Interest / Dividend Income',
      cyAmount: -interestIncome,
      pyAmount: 0,
      sourceAccounts: 'Other Income Schedule',
      notes: 'Investing income deducted from operating flow',
    });
  }

  const operatingProfitBeforeWc = netProfitBeforeTax + depreciation + financeCosts - interestIncome;

  operatingItems.push({
    id: 'op-subtotal',
    category: 'operating',
    description: 'Operating Profit Before Working Capital Changes',
    cyAmount: operatingProfitBeforeWc,
    pyAmount: 0,
    sourceAccounts: 'Subtotal',
    isSubtotal: true,
  });

  // Working Capital Adjustments
  const diffDebtors = cyDebtors - pyDebtors;
  operatingItems.push({
    id: 'op-wc-1',
    category: 'operating',
    description: `${diffDebtors >= 0 ? '(Increase)' : 'Decrease'} in Trade Receivables`,
    cyAmount: -diffDebtors,
    pyAmount: 0,
    sourceAccounts: `Trade Receivables (CY ₹${cyDebtors.toLocaleString('en-IN')} vs PY ₹${pyDebtors.toLocaleString('en-IN')})`,
    notes: diffDebtors >= 0 ? 'Cash tied up in credit sales' : 'Cash collected from debtors',
  });

  const diffInventory = cyInventory - pyInventory;
  operatingItems.push({
    id: 'op-wc-2',
    category: 'operating',
    description: `${diffInventory >= 0 ? '(Increase)' : 'Decrease'} in Inventories`,
    cyAmount: -diffInventory,
    pyAmount: 0,
    sourceAccounts: `Inventories Schedule (CY ₹${cyInventory.toLocaleString('en-IN')} vs PY ₹${pyInventory.toLocaleString('en-IN')})`,
    notes: diffInventory >= 0 ? 'Cash outflow for stock buildup' : 'Cash released from inventory',
  });

  const diffCreditors = cyCreditors - pyCreditors;
  operatingItems.push({
    id: 'op-wc-3',
    category: 'operating',
    description: `${diffCreditors >= 0 ? 'Increase' : '(Decrease)'} in Trade Payables`,
    cyAmount: diffCreditors,
    pyAmount: 0,
    sourceAccounts: `Trade Payables (CY ₹${cyCreditors.toLocaleString('en-IN')} vs PY ₹${pyCreditors.toLocaleString('en-IN')})`,
    notes: diffCreditors >= 0 ? 'Credit extended by suppliers' : 'Cash paid to settle suppliers',
  });

  const diffOtherAssets = cyOtherCurrentAssets - pyOtherCurrentAssets;
  if (diffOtherAssets !== 0) {
    operatingItems.push({
      id: 'op-wc-4',
      category: 'operating',
      description: `${diffOtherAssets >= 0 ? '(Increase)' : 'Decrease'} in Other Current Assets & Advances`,
      cyAmount: -diffOtherAssets,
      pyAmount: 0,
      sourceAccounts: 'Other Current Assets',
    });
  }

  const diffOtherLiab = cyOtherCurrentLiabilities - pyOtherCurrentLiabilities;
  if (diffOtherLiab !== 0) {
    operatingItems.push({
      id: 'op-wc-5',
      category: 'operating',
      description: `${diffOtherLiab >= 0 ? 'Increase' : '(Decrease)'} in Other Current Liabilities & Provisions`,
      cyAmount: diffOtherLiab,
      pyAmount: 0,
      sourceAccounts: 'Current Liabilities & Provisions',
    });
  }

  const cashGeneratedFromOps =
    operatingProfitBeforeWc - diffDebtors - diffInventory + diffCreditors - diffOtherAssets + diffOtherLiab;

  if (taxesPaid > 0) {
    operatingItems.push({
      id: 'op-tax',
      category: 'operating',
      description: 'Less: Direct Taxes Paid (Net of Refund)',
      cyAmount: -taxesPaid,
      pyAmount: 0,
      sourceAccounts: 'Advance Tax & TDS Payments',
    });
  }

  const netOperatingCash = cashGeneratedFromOps - taxesPaid;
  operatingItems.push({
    id: 'op-total',
    category: 'operating',
    description: 'Net Cash Flow from Operating Activities (A)',
    cyAmount: netOperatingCash,
    pyAmount: 0,
    sourceAccounts: 'Operating Total',
    isTotal: true,
  });

  // 2. Investing Activities
  const investingItems: CashFlowLineItem[] = [];

  // Fixed Asset additions (if net gross cost increased)
  const diffFixedAssets = cyFixedAssetsGross - pyFixedAssetsGross;
  const capexAmount = Math.max(0, diffFixedAssets);
  investingItems.push({
    id: 'inv-1',
    category: 'investing',
    description: 'Purchase of Property, Plant & Equipment (Capital Expenditure)',
    cyAmount: -capexAmount,
    pyAmount: 0,
    sourceAccounts: 'Fixed Asset Additions Schedule',
    notes: 'Outflow for plant machinery & infrastructure',
  });

  if (proceedsSaleAssets > 0) {
    investingItems.push({
      id: 'inv-2',
      category: 'investing',
      description: 'Proceeds from Sale of Fixed Assets',
      cyAmount: proceedsSaleAssets,
      pyAmount: 0,
      sourceAccounts: 'Asset Disposals Register',
    });
  }

  if (interestIncome > 0) {
    investingItems.push({
      id: 'inv-3',
      category: 'investing',
      description: 'Interest & Dividends Received',
      cyAmount: interestIncome,
      pyAmount: 0,
      sourceAccounts: 'Bank & Investment Statements',
    });
  }

  const netInvestingCash = -capexAmount + proceedsSaleAssets + interestIncome;
  investingItems.push({
    id: 'inv-total',
    category: 'investing',
    description: 'Net Cash Flow used in Investing Activities (B)',
    cyAmount: netInvestingCash,
    pyAmount: 0,
    sourceAccounts: 'Investing Total',
    isTotal: true,
  });

  // 3. Financing Activities
  const financingItems: CashFlowLineItem[] = [];

  const diffShareCapital = cyShareCapital - pyShareCapital;
  if (diffShareCapital > 0) {
    financingItems.push({
      id: 'fin-1',
      category: 'financing',
      description: 'Proceeds from Issue of Equity Share Capital',
      cyAmount: diffShareCapital,
      pyAmount: 0,
      sourceAccounts: 'Share Capital Schedule',
    });
  }

  const diffBorrowings = cyBorrowings - pyBorrowings;
  if (diffBorrowings !== 0) {
    financingItems.push({
      id: 'fin-2',
      category: 'financing',
      description: `${diffBorrowings > 0 ? 'Proceeds from' : 'Repayment of'} Long-Term & Short-Term Borrowings`,
      cyAmount: diffBorrowings,
      pyAmount: 0,
      sourceAccounts: 'Borrowings Schedule',
      notes: diffBorrowings > 0 ? 'Loan disbursement from banks' : 'Principal debt repayment',
    });
  }

  if (financeCosts > 0) {
    financingItems.push({
      id: 'fin-3',
      category: 'financing',
      description: 'Finance Costs / Interest Paid on Borrowings',
      cyAmount: -financeCosts,
      pyAmount: 0,
      sourceAccounts: 'Bank Interest Advice',
      notes: 'Cash outflow for servicing debt',
    });
  }

  if (dividendPaid > 0) {
    financingItems.push({
      id: 'fin-4',
      category: 'financing',
      description: 'Dividends Paid to Shareholders',
      cyAmount: -dividendPaid,
      pyAmount: 0,
      sourceAccounts: 'Dividend Bank Account',
    });
  }

  const netFinancingCash = diffShareCapital + diffBorrowings - financeCosts - dividendPaid;
  financingItems.push({
    id: 'fin-total',
    category: 'financing',
    description: 'Net Cash Flow from / (used in) Financing Activities (C)',
    cyAmount: netFinancingCash,
    pyAmount: 0,
    sourceAccounts: 'Financing Total',
    isTotal: true,
  });

  // Net Change in Cash
  const netChangeInCash = netOperatingCash + netInvestingCash + netFinancingCash;

  // Reconcile with actual Balance Sheet Cash
  const openingCash = pyCash;
  const actualClosingCash = cyCash;

  if (openingCash === 0 && actualClosingCash === 0 && rows.length > 0) {
    unmapped.push('Cash and Bank Balances (Opening/Closing) not mapped. Please check Cash & Cash Equivalents line items.');
  }

  const computedClosingCash = openingCash + netChangeInCash;
  const reconciliationDifference = Math.abs(computedClosingCash - actualClosingCash);
  const isReconciled = reconciliationDifference <= 1.0;

  return {
    clientName,
    cyPeriod,
    pyPeriod,
    operatingItems,
    netOperatingCash,
    investingItems,
    netInvestingCash,
    financingItems,
    netFinancingCash,
    netChangeInCash,
    openingCash,
    computedClosingCash,
    actualClosingCash,
    reconciliationDifference,
    isReconciled,
    unmappedAccounts: unmapped,
  };
}

/**
 * Exports complete multi-sheet Excel Cash Flow Statement workbook
 */
export function exportCashFlowWorkbook(
  result: CashFlowStatementResult
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Formal Cash Flow Statement (AS-3 Format)
  const statementRows: Record<string, any>[] = [
    { 'Particulars / Cash Flow Items': 'A. CASH FLOW FROM OPERATING ACTIVITIES', 'Notes': '', 'Amount (₹)': '' },
    ...result.operatingItems.map((item) => ({
      'Particulars / Cash Flow Items': item.isTotal || item.isSubtotal ? item.description : `  ${item.description}`,
      'Notes': item.notes || item.sourceAccounts,
      'Amount (₹)': item.cyAmount,
    })),
    { 'Particulars / Cash Flow Items': '', 'Notes': '', 'Amount (₹)': '' },
    { 'Particulars / Cash Flow Items': 'B. CASH FLOW FROM INVESTING ACTIVITIES', 'Notes': '', 'Amount (₹)': '' },
    ...result.investingItems.map((item) => ({
      'Particulars / Cash Flow Items': item.isTotal ? item.description : `  ${item.description}`,
      'Notes': item.notes || item.sourceAccounts,
      'Amount (₹)': item.cyAmount,
    })),
    { 'Particulars / Cash Flow Items': '', 'Notes': '', 'Amount (₹)': '' },
    { 'Particulars / Cash Flow Items': 'C. CASH FLOW FROM FINANCING ACTIVITIES', 'Notes': '', 'Amount (₹)': '' },
    ...result.financingItems.map((item) => ({
      'Particulars / Cash Flow Items': item.isTotal ? item.description : `  ${item.description}`,
      'Notes': item.notes || item.sourceAccounts,
      'Amount (₹)': item.cyAmount,
    })),
    { 'Particulars / Cash Flow Items': '', 'Notes': '', 'Amount (₹)': '' },
    { 'Particulars / Cash Flow Items': 'NET INCREASE / (DECREASE) IN CASH & CASH EQUIVALENTS (A + B + C)', 'Notes': 'Net Movement', 'Amount (₹)': result.netChangeInCash },
    { 'Particulars / Cash Flow Items': 'Cash and Cash Equivalents at the Beginning of the Year', 'Notes': 'Opening Balance Sheet', 'Amount (₹)': result.openingCash },
    { 'Particulars / Cash Flow Items': 'Cash and Cash Equivalents at the End of the Year', 'Notes': 'Computed Closing', 'Amount (₹)': result.computedClosingCash },
    { 'Particulars / Cash Flow Items': 'Balance Sheet Closing Cash Tie-Out Check', 'Notes': result.isReconciled ? '100% RECONCILED' : 'DISCREPANCY DETECTED', 'Amount (₹)': result.actualClosingCash },
  ];

  const wsStatement = createStyledWorksheet(statementRows, {
    currencyColumns: ['Amount (₹)'],
    columnWidths: {
      'Particulars / Cash Flow Items': 65,
      'Notes': 45,
      'Amount (₹)': 22,
    },
  });
  XLSX.utils.book_append_sheet(wb, wsStatement, 'Cash_Flow_Statement');

  // Sheet 2: Reconciliation & Working Schedules
  const reconRows = [
    { Parameter: 'Client Entity Name', Value: result.clientName },
    { Parameter: 'Financial Period', Value: result.cyPeriod },
    { Parameter: 'Accounting Standard Applied', Value: 'AS-3 / Ind AS 7 (Indirect Method)' },
    { Parameter: 'Net Cash from Operating Activities (₹)', Value: result.netOperatingCash },
    { Parameter: 'Net Cash used in Investing Activities (₹)', Value: result.netInvestingCash },
    { Parameter: 'Net Cash from Financing Activities (₹)', Value: result.netFinancingCash },
    { Parameter: 'Net Increase / (Decrease) in Cash (₹)', Value: result.netChangeInCash },
    { Parameter: 'Opening Cash and Cash Equivalents (₹)', Value: result.openingCash },
    { Parameter: 'Computed Closing Cash and Cash Equivalents (₹)', Value: result.computedClosingCash },
    { Parameter: 'Actual Balance Sheet Cash (₹)', Value: result.actualClosingCash },
    { Parameter: 'Reconciliation Variance (₹)', Value: result.reconciliationDifference },
    { Parameter: 'Reconciliation Tie-Out Status', Value: result.isReconciled ? 'BALANCED (ZERO VARIANCE)' : 'VARIANCE DETECTED' },
  ];
  const wsRecon = createStyledWorksheet(reconRows, {
    currencyColumns: ['Value'],
    columnWidths: { Parameter: 42, Value: 34 },
  });
  XLSX.utils.book_append_sheet(wb, wsRecon, 'Reconciliation_Tie_Out');

  return wb;
}

/**
 * Realistic Indian Balance Sheet & P&L Sample Dataset for Cash Flow
 */
export function getSampleCashFlowRows(): Record<string, any>[] {
  return [
    { 'Particulars': 'Profit Before Tax (from Statement of P&L)', 'Category': 'Operating PBT', 'FY 2026-27 (Current)': 5400000, 'FY 2025-26 (Prior)': 4200000 },
    { 'Particulars': 'Depreciation on Property, Plant & Equipment', 'Category': 'Non-Cash Adjustment', 'FY 2026-27 (Current)': 1650000, 'FY 2025-26 (Prior)': 1500000 },
    { 'Particulars': 'Finance Costs (Bank Loan Interest)', 'Category': 'Operating Adjustment', 'FY 2026-27 (Current)': 1250000, 'FY 2025-26 (Prior)': 1400000 },
    { 'Particulars': 'Interest Income on Fixed Deposits', 'Category': 'Operating Adjustment', 'FY 2026-27 (Current)': 180000, 'FY 2025-26 (Prior)': 150000 },
    { 'Particulars': 'Direct Income Taxes Paid', 'Category': 'Operating Tax', 'FY 2026-27 (Current)': 1350000, 'FY 2025-26 (Prior)': 1050000 },

    // Working Capital Balances
    { 'Particulars': 'Trade Receivables (Sundry Debtors)', 'Category': 'Current Asset', 'FY 2026-27 (Current)': 7100000, 'FY 2025-26 (Prior)': 6200000 }, // +9,00,000 Outflow
    { 'Particulars': 'Inventories / Stock in Trade', 'Category': 'Current Asset', 'FY 2026-27 (Current)': 6200000, 'FY 2025-26 (Prior)': 5400000 }, // +8,00,000 Outflow
    { 'Particulars': 'Other Current Assets & Advances', 'Category': 'Current Asset', 'FY 2026-27 (Current)': 1450000, 'FY 2025-26 (Prior)': 1200000 }, // +2,50,000 Outflow
    { 'Particulars': 'Trade Payables (Sundry Creditors)', 'Category': 'Current Liability', 'FY 2026-27 (Current)': 5400000, 'FY 2025-26 (Prior)': 4800000 }, // +6,00,000 Inflow
    { 'Particulars': 'Other Current Liabilities & Provisions', 'Category': 'Current Liability', 'FY 2026-27 (Current)': 1850000, 'FY 2025-26 (Prior)': 1550000 }, // +3,00,000 Inflow

    // Investing
    { 'Particulars': 'Property, Plant & Equipment (Gross Capex)', 'Category': 'Fixed Asset', 'FY 2026-27 (Current)': 22500000, 'FY 2025-26 (Prior)': 19800000 }, // Capex 27,00,000 Outflow
    { 'Particulars': 'Proceeds from Sale of Old Machinery', 'Category': 'Asset Sale', 'FY 2026-27 (Current)': 350000, 'FY 2025-26 (Prior)': 0 },

    // Financing
    { 'Particulars': 'Equity Share Capital', 'Category': 'Equity', 'FY 2026-27 (Current)': 10000000, 'FY 2025-26 (Prior)': 10000000 }, // Nil change
    { 'Particulars': 'Long-Term Term Loans (Bank Borrowings)', 'Category': 'Borrowing', 'FY 2026-27 (Current)': 6500000, 'FY 2025-26 (Prior)': 7800000 }, // -13,00,000 Repayment
    { 'Particulars': 'Dividends Paid to Shareholders', 'Category': 'Financing Dividend', 'FY 2026-27 (Current)': 500000, 'FY 2025-26 (Prior)': 400000 },

    // Cash
    { 'Particulars': 'Cash and Bank Balances', 'Category': 'Cash Equivalent', 'FY 2026-27 (Current)': 2820000, 'FY 2025-26 (Prior)': 2100000 }, // Net change +7,20,000
  ];
}
