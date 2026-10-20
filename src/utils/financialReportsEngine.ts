import * as XLSX from 'xlsx';
import { parseAmount } from './dataSanitizer';
import { detectColumnCategory } from './columnDetector';
import { createStyledWorksheet } from './excelStyler';

export type AccountHead = 'Asset' | 'Liability' | 'Equity' | 'Revenue' | 'Expense';

export type ScheduleIIIGroup =
  | 'Share Capital'
  | 'Reserves & Surplus'
  | 'Non-Current Liabilities'
  | 'Trade Payables'
  | 'Other Current Liabilities'
  | 'Property, Plant & Equipment'
  | 'Non-Current Investments'
  | 'Inventories'
  | 'Trade Receivables'
  | 'Cash & Bank Balances'
  | 'Other Current Assets'
  | 'Revenue from Operations'
  | 'Other Income'
  | 'Cost of Materials Consumed'
  | 'Employee Benefit Expenses'
  | 'Finance Costs'
  | 'Depreciation & Amortization'
  | 'Other Expenses';

export interface TrialBalanceRecord {
  id: string;
  accountCode: string;
  accountName: string;
  head: AccountHead;
  subHead: string;
  scheduleIIIGroup: ScheduleIIIGroup;
  openingDebit?: number;
  openingCredit?: number;
  periodDebit?: number;
  periodCredit?: number;
  debit: number; // Closing Debit
  credit: number; // Closing Credit
  netBalance: number; // Positive = Debit, Negative = Credit
  isAbnormal?: boolean;
  abnormalReason?: string;
}

export interface AuditAdjustmentEntry {
  id: string;
  debitAccountCode: string;
  debitAccountName: string;
  creditAccountCode: string;
  creditAccountName: string;
  amount: number;
  narration: string;
  category: 'Closing Stock' | 'Depreciation' | 'Provision' | 'Accrual' | 'Prepaid' | 'Reclassification' | 'Other';
  isApproved: boolean;
}

export interface ScheduleIIIBreakdown {
  shareCapital: number;
  reservesAndSurplus: number;
  nonCurrentLiabilities: number;
  tradePayables: number;
  otherCurrentLiabilities: number;
  totalLiabilitiesAndEquity: number;

  ppe: number;
  nonCurrentInvestments: number;
  inventories: number;
  tradeReceivables: number;
  cashAndBank: number;
  otherCurrentAssets: number;
  totalAssets: number;

  revenueFromOperations: number;
  otherIncome: number;
  totalRevenue: number;

  costOfMaterials: number;
  employeeBenefits: number;
  financeCosts: number;
  depreciation: number;
  otherExpenses: number;
  totalExpenses: number;

  netProfitBeforeTax: number;
}

export interface FinancialStatementsResult {
  isBalanced: boolean;
  totalDebits: number;
  totalCredits: number;
  trialBalanceDiff: number;
  revenue: number;
  expenses: number;
  netProfit: number;
  assets: number;
  liabilities: number;
  equity: number;
  totalCapAndLiab: number;
  balanceSheetDiff: number;
  trialBalance: TrialBalanceRecord[];
  scheduleIII: ScheduleIIIBreakdown;
  adjustments: AuditAdjustmentEntry[];
  abnormalAccounts: TrialBalanceRecord[];
}

/**
 * Maps account name and nature to Companies Act Schedule III Group
 */
export function mapToScheduleIII(name: string, head: AccountHead, netBalance: number): ScheduleIIIGroup {
  const n = name.toLowerCase();

  // Revenue & Income
  if (head === 'Revenue') {
    if (n.includes('interest') || n.includes('dividend') || n.includes('rent received') || n.includes('scrap') || n.includes('discount received') || n.includes('other income')) {
      return 'Other Income';
    }
    return 'Revenue from Operations';
  }

  // Expenses
  if (head === 'Expense') {
    if (n.includes('purchase') || n.includes('raw material') || n.includes('stock in trade') || n.includes('freight in')) {
      return 'Cost of Materials Consumed';
    }
    if (n.includes('salary') || n.includes('wage') || n.includes('bonus') || n.includes('pf') || n.includes('esi') || n.includes('staff welfare') || n.includes('gratuity')) {
      return 'Employee Benefit Expenses';
    }
    if (n.includes('interest paid') || n.includes('bank charge') || n.includes('finance cost') || n.includes('processing fee') || n.includes('loan interest')) {
      return 'Finance Costs';
    }
    if (n.includes('depreciation') || n.includes('amortization') || n.includes('written off')) {
      return 'Depreciation & Amortization';
    }
    return 'Other Expenses';
  }

  // Equity
  if (head === 'Equity') {
    if (n.includes('share capital') || n.includes('equity capital') || n.includes('partners capital')) {
      return 'Share Capital';
    }
    return 'Reserves & Surplus';
  }

  // Liabilities
  if (head === 'Liability') {
    if (n.includes('creditor') || n.includes('trade payable') || n.includes('bills payable')) {
      return 'Trade Payables';
    }
    if (n.includes('term loan') || n.includes('borrowing') || n.includes('debenture') || n.includes('car loan') || n.includes('machinery loan')) {
      return 'Non-Current Liabilities';
    }
    return 'Other Current Liabilities';
  }

  // Assets
  if (head === 'Asset') {
    if (n.includes('cash') || n.includes('bank') || n.includes('petty cash') || n.includes('fixed deposit') || n.includes('fd ')) {
      return 'Cash & Bank Balances';
    }
    if (n.includes('debtor') || n.includes('trade receivable') || n.includes('bills receivable')) {
      return 'Trade Receivables';
    }
    if (n.includes('inventory') || n.includes('stock') || n.includes('raw material') || n.includes('finished goods') || n.includes('wip')) {
      return 'Inventories';
    }
    if (n.includes('investment') || n.includes('mutual fund') || n.includes('shares')) {
      return 'Non-Current Investments';
    }
    if (n.includes('plant') || n.includes('machinery') || n.includes('vehicle') || n.includes('computer') || n.includes('furniture') || n.includes('building') || n.includes('land') || n.includes('equipment') || n.includes('fixed asset')) {
      return 'Property, Plant & Equipment';
    }
    return 'Other Current Assets';
  }

  return 'Other Current Assets';
}

/**
 * Infers accounting head if not explicitly specified
 */
export function inferAccountHead(accountName: string, debit: number, credit: number): { head: AccountHead; subHead: string } {
  const name = (accountName || '').toLowerCase();

  if (
    name.includes('sale') ||
    name.includes('revenue') ||
    name.includes('turnover') ||
    name.includes('fees received') ||
    name.includes('interest received') ||
    name.includes('discount received') ||
    name.includes('other income') ||
    name.includes('changes in inventor') ||
    name.includes('change in inventory') ||
    name.includes('closing stock (p&l)') ||
    name.includes('closing stock - p&l')
  ) {
    return { head: 'Revenue', subHead: 'Operating & Other Revenue' };
  }

  if (
    name.includes('purchase') ||
    name.includes('salary') ||
    name.includes('wage') ||
    name.includes('freight') ||
    name.includes('fuel') ||
    name.includes('toll') ||
    name.includes('rent') ||
    name.includes('depreciation') ||
    name.includes('repair') ||
    name.includes('interest paid') ||
    name.includes('finance cost') ||
    name.includes('audit fee') ||
    name.includes('printing') ||
    name.includes('legal') ||
    name.includes('expense')
  ) {
    return { head: 'Expense', subHead: 'Operating & Administrative Expenses' };
  }

  if (
    name.includes('share capital') ||
    name.includes('capital account') ||
    name.includes('equity') ||
    name.includes('reserves') ||
    name.includes('surplus') ||
    name.includes('retained earnings')
  ) {
    return { head: 'Equity', subHead: 'Shareholders Funds' };
  }

  if (
    name.includes('creditor') ||
    name.includes('trade payable') ||
    name.includes('loan') ||
    name.includes('borrowing') ||
    name.includes('overdraft') ||
    name.includes('duties & taxes') ||
    name.includes('gst output') ||
    name.includes('tds payable') ||
    name.includes('provision') ||
    name.includes('payable')
  ) {
    return { head: 'Liability', subHead: 'Current & Non-Current Liabilities' };
  }

  if (
    name.includes('bank') ||
    name.includes('cash') ||
    name.includes('debtor') ||
    name.includes('trade receivable') ||
    name.includes('fixed asset') ||
    name.includes('vehicle') ||
    name.includes('plant') ||
    name.includes('machinery') ||
    name.includes('equipment') ||
    name.includes('gst input') ||
    name.includes('itc') ||
    name.includes('tds receivable') ||
    name.includes('prepaid') ||
    name.includes('deposit') ||
    name.includes('investment') ||
    name.includes('inventory') ||
    name.includes('stock')
  ) {
    return { head: 'Asset', subHead: 'Current & Non-Current Assets' };
  }

  if (debit > credit) {
    return { head: 'Asset', subHead: 'Sundry Debit Balances' };
  }
  return { head: 'Liability', subHead: 'Sundry Credit Balances' };
}

/**
 * Checks for abnormal account balances (e.g. negative cash, credit debtors, debit creditors)
 */
export function checkAbnormalBalance(record: TrialBalanceRecord): { isAbnormal: boolean; reason?: string } {
  const n = record.accountName.toLowerCase();
  const net = record.debit - record.credit;

  // 1. Cash in Hand cannot have credit balance (overdrawn cash)
  if (n.includes('cash in hand') || n.includes('petty cash')) {
    if (net < 0) {
      return { isAbnormal: true, reason: 'Abnormal Credit Balance in Cash in Hand (Cash cannot be negative)' };
    }
  }

  // 2. Sundry Debtors with Credit balance (Advance from Customers)
  if ((n.includes('debtor') || n.includes('trade receivable')) && net < 0) {
    return { isAbnormal: true, reason: 'Credit balance in Sundry Debtors (Customer advances should be reclassified to Current Liabilities)' };
  }

  // 3. Sundry Creditors with Debit balance (Advance to Suppliers)
  if ((n.includes('creditor') || n.includes('trade payable')) && net > 0) {
    return { isAbnormal: true, reason: 'Debit balance in Sundry Creditors (Supplier advances should be reclassified to Current Loans & Advances)' };
  }

  // 4. Sales/Revenue with Debit balance
  if (record.head === 'Revenue' && net > 0) {
    return { isAbnormal: true, reason: 'Debit balance in Revenue Account (Abnormal sales returns or debit posting)' };
  }

  // 5. Purchases/Expenses with Credit balance
  if (record.head === 'Expense' && net < 0) {
    return { isAbnormal: true, reason: 'Credit balance in Expense Account (Abnormal purchase returns or credit posting)' };
  }

  return { isAbnormal: false };
}

/**
 * Parses raw Trial Balance spreadsheet rows into standardized TrialBalanceRecords
 */
export function parseTrialBalanceRows(rawRows: Record<string, any>[]): TrialBalanceRecord[] {
  if (!rawRows || rawRows.length === 0) return [];
  const sample = rawRows[0];
  const colMap: Record<string, string> = {};

  for (const key of Object.keys(sample)) {
    const cat = detectColumnCategory(key);
    if (cat) colMap[cat] = key;
  }

  return rawRows.map((row, idx) => {
    const code = (colMap.accountCode ? row[colMap.accountCode] : row.accountCode || row['Account Code'] || row.Code || `AC-${1000 + idx}`).toString().trim();
    const name = (colMap.partyName ? row[colMap.partyName] : row.accountName || row['Account Name'] || row.Particulars || 'Account').toString().trim();

    // Support opening, period, and closing debits/credits
    const opDr = parseAmount(row['Opening Debit'] || row['Op Debit'] || 0);
    const opCr = parseAmount(row['Opening Credit'] || row['Op Credit'] || 0);
    const perDr = parseAmount(row['Period Debit'] || row['Dr Movement'] || 0);
    const perCr = parseAmount(row['Period Credit'] || row['Cr Movement'] || 0);

    let debit = parseAmount(colMap.debit ? row[colMap.debit] : row.debit || row.Debit || row['Closing Debit'] || 0);
    let credit = parseAmount(colMap.credit ? row[colMap.credit] : row.credit || row.Credit || row['Closing Credit'] || 0);

    // If only net balance column exists
    if (debit === 0 && credit === 0) {
      const net = parseAmount(row.Balance || row['Net Balance'] || row.Amount || 0);
      if (net >= 0) debit = net;
      else credit = Math.abs(net);
    }

    let explicitHead = (colMap.head ? row[colMap.head] : row.head || row.Head || '').toString().trim();
    let head: AccountHead;
    let subHead = 'General';

    if (['Asset', 'Liability', 'Equity', 'Revenue', 'Expense'].includes(explicitHead)) {
      head = explicitHead as AccountHead;
    } else {
      const inferred = inferAccountHead(name, debit, credit);
      head = inferred.head;
      subHead = inferred.subHead;
    }

    const netBal = debit - credit;
    const scheduleIIIGroup = mapToScheduleIII(name, head, netBal);

    const record: TrialBalanceRecord = {
      id: `tb-${idx + 1}`,
      accountCode: code,
      accountName: name,
      head,
      subHead,
      scheduleIIIGroup,
      openingDebit: opDr,
      openingCredit: opCr,
      periodDebit: perDr,
      periodCredit: perCr,
      debit,
      credit,
      netBalance: netBal,
    };

    const abnormal = checkAbnormalBalance(record);
    record.isAbnormal = abnormal.isAbnormal;
    record.abnormalReason = abnormal.reason;

    return record;
  });
}

/**
 * Compiles Companies Act Schedule III Balance Sheet and Profit & Loss from Trial Balance and user adjustments
 */
export function compileFinancialStatements(
  records: TrialBalanceRecord[],
  adjustments: AuditAdjustmentEntry[] = []
): FinancialStatementsResult {
  // Apply approved audit adjustments
  const adjustedRecords: TrialBalanceRecord[] = records.map((r) => ({ ...r }));

  for (const adj of adjustments) {
    if (!adj.isApproved) continue;

    // Apply Debit
    let drAccount = adjustedRecords.find(
      (r) => r.accountCode === adj.debitAccountCode || r.accountName.toLowerCase() === adj.debitAccountName.toLowerCase()
    );
    if (!drAccount) {
      const inferred = inferAccountHead(adj.debitAccountName, adj.amount, 0);
      drAccount = {
        id: `adj-dr-${adj.id}`,
        accountCode: adj.debitAccountCode || 'ADJ-DR',
        accountName: adj.debitAccountName,
        head: inferred.head,
        subHead: inferred.subHead,
        scheduleIIIGroup: mapToScheduleIII(adj.debitAccountName, inferred.head, adj.amount),
        debit: adj.amount,
        credit: 0,
        netBalance: adj.amount,
      };
      adjustedRecords.push(drAccount);
    } else {
      drAccount.debit += adj.amount;
      drAccount.netBalance = drAccount.debit - drAccount.credit;
    }

    // Apply Credit
    let crAccount = adjustedRecords.find(
      (r) => r.accountCode === adj.creditAccountCode || r.accountName.toLowerCase() === adj.creditAccountName.toLowerCase()
    );
    if (!crAccount) {
      const inferred = inferAccountHead(adj.creditAccountName, 0, adj.amount);
      crAccount = {
        id: `adj-cr-${adj.id}`,
        accountCode: adj.creditAccountCode || 'ADJ-CR',
        accountName: adj.creditAccountName,
        head: inferred.head,
        subHead: inferred.subHead,
        scheduleIIIGroup: mapToScheduleIII(adj.creditAccountName, inferred.head, -adj.amount),
        debit: 0,
        credit: adj.amount,
        netBalance: -adj.amount,
      };
      adjustedRecords.push(crAccount);
    } else {
      crAccount.credit += adj.amount;
      crAccount.netBalance = crAccount.debit - crAccount.credit;
    }
  }

  const totalDebits = adjustedRecords.reduce((acc, r) => acc + r.debit, 0);
  const totalCredits = adjustedRecords.reduce((acc, r) => acc + r.credit, 0);
  const trialBalanceDiff = Math.abs(totalDebits - totalCredits);
  const isBalanced = trialBalanceDiff < 0.05;

  // Trading & P&L
  const revenue = adjustedRecords.filter((r) => r.head === 'Revenue').reduce((acc, r) => acc + (r.credit - r.debit), 0);
  const expenses = adjustedRecords.filter((r) => r.head === 'Expense').reduce((acc, r) => acc + (r.debit - r.credit), 0);
  const netProfit = revenue - expenses;

  // Schedule III Breakdowns
  const sumGroup = (grp: ScheduleIIIGroup, isDebitSide: boolean) =>
    adjustedRecords
      .filter((r) => r.scheduleIIIGroup === grp)
      .reduce((sum, r) => sum + (isDebitSide ? r.debit - r.credit : r.credit - r.debit), 0);

  const shareCapital = sumGroup('Share Capital', false);
  const reservesAndSurplus = sumGroup('Reserves & Surplus', false) + netProfit;
  const nonCurrentLiabilities = sumGroup('Non-Current Liabilities', false);
  const tradePayables = sumGroup('Trade Payables', false);
  const otherCurrentLiabilities = sumGroup('Other Current Liabilities', false);
  const totalLiabilitiesAndEquity = shareCapital + reservesAndSurplus + nonCurrentLiabilities + tradePayables + otherCurrentLiabilities;

  const ppe = sumGroup('Property, Plant & Equipment', true);
  const nonCurrentInvestments = sumGroup('Non-Current Investments', true);
  const inventories = sumGroup('Inventories', true);
  const tradeReceivables = sumGroup('Trade Receivables', true);
  const cashAndBank = sumGroup('Cash & Bank Balances', true);
  const otherCurrentAssets = sumGroup('Other Current Assets', true);
  const totalAssets = ppe + nonCurrentInvestments + inventories + tradeReceivables + cashAndBank + otherCurrentAssets;

  const revenueFromOperations = sumGroup('Revenue from Operations', false);
  const otherIncome = sumGroup('Other Income', false);
  const totalRevenue = revenueFromOperations + otherIncome;

  const costOfMaterials = sumGroup('Cost of Materials Consumed', true);
  const employeeBenefits = sumGroup('Employee Benefit Expenses', true);
  const financeCosts = sumGroup('Finance Costs', true);
  const depreciation = sumGroup('Depreciation & Amortization', true);
  const otherExpenses = sumGroup('Other Expenses', true);
  const totalExpenses = costOfMaterials + employeeBenefits + financeCosts + depreciation + otherExpenses;

  const netProfitBeforeTax = totalRevenue - totalExpenses;

  // General Balance Sheet Head Totals
  const assets = adjustedRecords.filter((r) => r.head === 'Asset').reduce((acc, r) => acc + (r.debit - r.credit), 0);
  const liabilities = adjustedRecords.filter((r) => r.head === 'Liability').reduce((acc, r) => acc + (r.credit - r.debit), 0);
  const equity = adjustedRecords.filter((r) => r.head === 'Equity').reduce((acc, r) => acc + (r.credit - r.debit), 0);
  const totalCapAndLiab = equity + liabilities + netProfit;
  const balanceSheetDiff = Math.abs(assets - totalCapAndLiab);

  const abnormalAccounts = adjustedRecords.filter((r) => r.isAbnormal);

  return {
    isBalanced,
    totalDebits,
    totalCredits,
    trialBalanceDiff,
    revenue,
    expenses,
    netProfit,
    assets,
    liabilities,
    equity,
    totalCapAndLiab,
    balanceSheetDiff,
    trialBalance: adjustedRecords,
    scheduleIII: {
      shareCapital,
      reservesAndSurplus,
      nonCurrentLiabilities,
      tradePayables,
      otherCurrentLiabilities,
      totalLiabilitiesAndEquity,
      ppe,
      nonCurrentInvestments,
      inventories,
      tradeReceivables,
      cashAndBank,
      otherCurrentAssets,
      totalAssets,
      revenueFromOperations,
      otherIncome,
      totalRevenue,
      costOfMaterials,
      employeeBenefits,
      financeCosts,
      depreciation,
      otherExpenses,
      totalExpenses,
      netProfitBeforeTax,
    },
    adjustments,
    abnormalAccounts,
  };
}

/**
 * Creates formatted Companies Act Schedule III Balance Sheet, P&L, Adjusted TB, and Adjustments Excel workbook
 */
export function createFinancialStatementsWorkbook(
  result: FinancialStatementsResult,
  financialYear: string = 'FY 2025-26'
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();
  const s3 = result.scheduleIII;

  // Sheet 1: Schedule III Balance Sheet
  const bsRows = [
    { 'Schedule III Part I': 'I. EQUITY AND LIABILITIES', Note: '', Particulars: '', 'Amount (₹)': '' },
    { 'Schedule III Part I': '(1) Shareholders Funds', Note: 'Note 1', Particulars: 'Equity Share Capital', 'Amount (₹)': s3.shareCapital },
    { 'Schedule III Part I': '', Note: 'Note 2', Particulars: 'Reserves & Surplus (incl. Current Period Profit)', 'Amount (₹)': s3.reservesAndSurplus },
    { 'Schedule III Part I': '(2) Non-Current Liabilities', Note: 'Note 3', Particulars: 'Long-Term Borrowings & Provisions', 'Amount (₹)': s3.nonCurrentLiabilities },
    { 'Schedule III Part I': '(3) Current Liabilities', Note: 'Note 4', Particulars: 'Trade Payables (Sundry Creditors)', 'Amount (₹)': s3.tradePayables },
    { 'Schedule III Part I': '', Note: 'Note 5', Particulars: 'Other Current Liabilities & Provisions', 'Amount (₹)': s3.otherCurrentLiabilities },
    { 'Schedule III Part I': 'TOTAL EQUITY & LIABILITIES', Note: '', Particulars: 'TOTAL EQUITY & LIABILITIES', 'Amount (₹)': s3.totalLiabilitiesAndEquity },
    { 'Schedule III Part I': '', Note: '', Particulars: '', 'Amount (₹)': '' },
    { 'Schedule III Part I': 'II. ASSETS', Note: '', Particulars: '', 'Amount (₹)': '' },
    { 'Schedule III Part I': '(1) Non-Current Assets', Note: 'Note 6', Particulars: 'Property, Plant & Equipment (PPE)', 'Amount (₹)': s3.ppe },
    { 'Schedule III Part I': '', Note: 'Note 7', Particulars: 'Non-Current Investments', 'Amount (₹)': s3.nonCurrentInvestments },
    { 'Schedule III Part I': '(2) Current Assets', Note: 'Note 8', Particulars: 'Inventories (Closing Stock)', 'Amount (₹)': s3.inventories },
    { 'Schedule III Part I': '', Note: 'Note 9', Particulars: 'Trade Receivables (Sundry Debtors)', 'Amount (₹)': s3.tradeReceivables },
    { 'Schedule III Part I': '', Note: 'Note 10', Particulars: 'Cash & Cash Equivalents (Bank & Cash)', 'Amount (₹)': s3.cashAndBank },
    { 'Schedule III Part I': '', Note: 'Note 11', Particulars: 'Short-Term Loans, Advances & Other Current Assets', 'Amount (₹)': s3.otherCurrentAssets },
    { 'Schedule III Part I': 'TOTAL ASSETS', Note: '', Particulars: 'TOTAL ASSETS', 'Amount (₹)': s3.totalAssets },
    { 'Schedule III Part I': '', Note: '', Particulars: '', 'Amount (₹)': '' },
    { 'Schedule III Part I': 'Mathematical Reconciliation', Note: '', Particulars: Math.abs(s3.totalAssets - s3.totalLiabilitiesAndEquity) < 1 ? 'Balanced (Zero Discrepancy)' : 'Imbalance Detected', 'Amount (₹)': Math.abs(s3.totalAssets - s3.totalLiabilitiesAndEquity) },
  ];
  const wsBs = createStyledWorksheet(bsRows, {
    currencyColumns: ['Amount (₹)'],
    columnWidths: { 'Schedule III Part I': 38, Note: 12, Particulars: 48, 'Amount (₹)': 22 },
  });
  XLSX.utils.book_append_sheet(wb, wsBs, 'Schedule_III_Balance_Sheet');

  // Sheet 2: Statement of Profit & Loss
  const plRows = [
    { 'Schedule III Part II': 'I. Revenue from Operations', Head: 'Revenue', 'Amount (₹)': s3.revenueFromOperations },
    { 'Schedule III Part II': 'II. Other Income', Head: 'Revenue', 'Amount (₹)': s3.otherIncome },
    { 'Schedule III Part II': 'III. Total Income (I + II)', Head: 'Total Income', 'Amount (₹)': s3.totalRevenue },
    { 'Schedule III Part II': '', Head: '', 'Amount (₹)': '' },
    { 'Schedule III Part II': 'IV. Expenses:', Head: '', 'Amount (₹)': '' },
    { 'Schedule III Part II': 'Cost of Materials Consumed', Head: 'Expense', 'Amount (₹)': s3.costOfMaterials },
    { 'Schedule III Part II': 'Employee Benefits Expense', Head: 'Expense', 'Amount (₹)': s3.employeeBenefits },
    { 'Schedule III Part II': 'Finance Costs', Head: 'Expense', 'Amount (₹)': s3.financeCosts },
    { 'Schedule III Part II': 'Depreciation and Amortization Expense', Head: 'Expense', 'Amount (₹)': s3.depreciation },
    { 'Schedule III Part II': 'Other Expenses', Head: 'Expense', 'Amount (₹)': s3.otherExpenses },
    { 'Schedule III Part II': 'Total Expenses', Head: 'Total Expense', 'Amount (₹)': s3.totalExpenses },
    { 'Schedule III Part II': '', Head: '', 'Amount (₹)': '' },
    { 'Schedule III Part II': 'V. Profit Before Exceptional Items & Tax (III - IV)', Head: 'Profit', 'Amount (₹)': s3.netProfitBeforeTax },
  ];
  const wsPl = createStyledWorksheet(plRows, {
    currencyColumns: ['Amount (₹)'],
    columnWidths: { 'Schedule III Part II': 48, Head: 18, 'Amount (₹)': 22 },
  });
  XLSX.utils.book_append_sheet(wb, wsPl, 'Profit_and_Loss');

  // Sheet 3: Adjusted Trial Balance
  const tbRows = result.trialBalance.map((r, idx) => ({
    'S.No': idx + 1,
    'Account Code': r.accountCode,
    'Account Name': r.accountName,
    'Head': r.head,
    'Schedule III Group': r.scheduleIIIGroup,
    'Debit (₹)': r.debit,
    'Credit (₹)': r.credit,
    'Net Balance (₹)': r.netBalance,
    'Audit Status': r.isAbnormal ? `Abnormal: ${r.abnormalReason}` : 'Normal',
  }));
  const wsTb = createStyledWorksheet(tbRows, {
    currencyColumns: ['Debit (₹)', 'Credit (₹)', 'Net Balance (₹)'],
    columnWidths: {
      'S.No': 8,
      'Account Code': 15,
      'Account Name': 36,
      'Head': 14,
      'Schedule III Group': 30,
      'Debit (₹)': 18,
      'Credit (₹)': 18,
      'Net Balance (₹)': 18,
      'Audit Status': 35,
    },
  });
  XLSX.utils.book_append_sheet(wb, wsTb, 'Adjusted_Trial_Balance');

  // Sheet 4: Audit Adjustments Log
  const adjRows = result.adjustments.map((a, idx) => ({
    'Adj #': idx + 1,
    'Category': a.category,
    'Debit Account': a.debitAccountName,
    'Credit Account': a.creditAccountName,
    'Adjustment Amount (₹)': a.amount,
    'Narration / Audit Rationale': a.narration,
    'Status': a.isApproved ? 'Approved & Included in Statements' : 'Draft Pending Review',
  }));
  const wsAdj = createStyledWorksheet(
    adjRows.length > 0 ? adjRows : [{ Note: 'Zero audit adjustments posted.' }],
    {
      currencyColumns: ['Adjustment Amount (₹)'],
      columnWidths: { 'Adj #': 10, Category: 18, 'Debit Account': 30, 'Credit Account': 30, 'Adjustment Amount (₹)': 22, 'Narration / Audit Rationale': 45, Status: 32 },
    }
  );
  XLSX.utils.book_append_sheet(wb, wsAdj, 'Audit_Adjustments_Log');

  // Sheet 5: Abnormal Balances Report
  const abnRows = result.abnormalAccounts.map((a, idx) => ({
    'Issue #': idx + 1,
    'Account Code': a.accountCode,
    'Account Name': a.accountName,
    'Head': a.head,
    'Debit (₹)': a.debit,
    'Credit (₹)': a.credit,
    'Discrepancy Detail': a.abnormalReason,
    'Recommended Audit Rectification': 'Inspect underlying ledger vouchers and pass reclassification adjustment entry.',
  }));
  const wsAbn = createStyledWorksheet(
    abnRows.length > 0 ? abnRows : [{ Note: 'Zero abnormal balances found in Trial Balance.' }],
    {
      currencyColumns: ['Debit (₹)', 'Credit (₹)'],
      columnWidths: { 'Issue #': 10, 'Account Code': 15, 'Account Name': 32, Head: 14, 'Debit (₹)': 18, 'Credit (₹)': 18, 'Discrepancy Detail': 45, 'Recommended Audit Rectification': 45 },
    }
  );
  XLSX.utils.book_append_sheet(wb, wsAbn, 'Abnormal_Balances');

  return wb;
}
