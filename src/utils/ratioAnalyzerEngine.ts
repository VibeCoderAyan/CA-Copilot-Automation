import * as XLSX from 'xlsx';
import { parseAmount } from './dataSanitizer';
import { createStyledWorksheet } from './excelStyler';

export interface FinancialInputs {
  periodLabel: string;
  // Balance Sheet
  currentAssets: number;
  cashAndBank: number;
  tradeReceivables: number;
  inventories: number;
  otherCurrentAssets: number;
  fixedAssetsNet: number;
  totalAssets: number;

  currentLiabilities: number;
  tradePayables: number;
  shortTermBorrowings: number;
  otherCurrentLiabilities: number;
  longTermDebt: number;
  totalDebt: number;
  shareholdersEquity: number;
  totalLiabilitiesAndEquity: number;

  // P&L
  revenue: number;
  cogs: number;
  grossProfit: number;
  operatingExpenses: number;
  operatingProfit: number; // EBIT
  financeCosts: number; // Interest
  depreciation: number;
  profitBeforeTax: number;
  taxExpense: number;
  profitAfterTax: number; // PAT
}

export interface RatioResult {
  id: string;
  category: 'Liquidity' | 'Solvency' | 'Profitability' | 'Activity / Turnover';
  name: string;
  formulaDescription: string;
  cyValue: number | null;
  pyValue: number | null;
  unit: 'ratio' | 'percentage' | 'days' | 'times';
  benchmark: string;
  interpretation: 'healthy' | 'neutral' | 'concern' | 'not_applicable';
  notes: string;
  cyFormulaExpanded: string;
  pyFormulaExpanded: string;
}

export interface RatioAnalysisResult {
  clientName: string;
  cyPeriod: string;
  pyPeriod: string;
  cyInputs: FinancialInputs;
  pyInputs: FinancialInputs;
  ratios: RatioResult[];
  alerts: string[];
}

export interface RatioColumnMapping {
  accountNameCol: string;
  categoryCol: string;
  cyAmountCol: string;
  pyAmountCol: string;
}

/**
 * Auto-detects financial statement mapping columns
 */
export function detectRatioColumns(headers: string[]): RatioColumnMapping {
  const mapping: RatioColumnMapping = {
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
 * Extracts and maps FinancialInputs from raw uploaded spreadsheet rows
 */
export function parseFinancialStatements(
  rows: Record<string, any>[],
  mapping: RatioColumnMapping,
  cyLabel = 'FY 2026-27',
  pyLabel = 'FY 2025-26'
): { cy: FinancialInputs; py: FinancialInputs; unmapped: string[] } {
  const cy: FinancialInputs = {
    periodLabel: cyLabel,
    currentAssets: 0,
    cashAndBank: 0,
    tradeReceivables: 0,
    inventories: 0,
    otherCurrentAssets: 0,
    fixedAssetsNet: 0,
    totalAssets: 0,
    currentLiabilities: 0,
    tradePayables: 0,
    shortTermBorrowings: 0,
    otherCurrentLiabilities: 0,
    longTermDebt: 0,
    totalDebt: 0,
    shareholdersEquity: 0,
    totalLiabilitiesAndEquity: 0,
    revenue: 0,
    cogs: 0,
    grossProfit: 0,
    operatingExpenses: 0,
    operatingProfit: 0,
    financeCosts: 0,
    depreciation: 0,
    profitBeforeTax: 0,
    taxExpense: 0,
    profitAfterTax: 0,
  };

  const py: FinancialInputs = {
    ...cy,
    periodLabel: pyLabel,
  };

  const unmapped: string[] = [];

  rows.forEach((r) => {
    const name = String(r[mapping.accountNameCol] || '').trim();
    const cat = String(r[mapping.categoryCol] || '').trim();
    const text = `${name} ${cat}`.toLowerCase();
    const cyVal = parseAmount(r[mapping.cyAmountCol]);
    const pyVal = parseAmount(r[mapping.pyAmountCol]);

    if (!name && cyVal === 0 && pyVal === 0) return;

    let matched = true;

    if (/revenue|sales|turnover|income.*operation/i.test(text) && !/other.*income/i.test(text)) {
      cy.revenue += cyVal;
      py.revenue += pyVal;
    } else if (/cost.*goods|cogs|raw.*material.*consumed|direct.*expense/i.test(text)) {
      cy.cogs += cyVal;
      py.cogs += pyVal;
    } else if (/gross.*profit/i.test(text)) {
      cy.grossProfit = cyVal;
      py.grossProfit = pyVal;
    } else if (/depreciation|amortisation|amortization/i.test(text)) {
      cy.depreciation += cyVal;
      py.depreciation += pyVal;
    } else if (/finance.*cost|interest.*expense|interest.*paid/i.test(text)) {
      cy.financeCosts += cyVal;
      py.financeCosts += pyVal;
    } else if (/tax.*expense|current.*tax|provision.*tax/i.test(text)) {
      cy.taxExpense += cyVal;
      py.taxExpense += pyVal;
    } else if (/profit.*after.*tax|net.*profit/i.test(text)) {
      cy.profitAfterTax = cyVal;
      py.profitAfterTax = pyVal;
    } else if (/cash|bank.*balance|bank.*current/i.test(text)) {
      cy.cashAndBank += cyVal;
      py.cashAndBank += pyVal;
      cy.currentAssets += cyVal;
      py.currentAssets += pyVal;
    } else if (/debtor|receivable|trade.*receivable/i.test(text)) {
      cy.tradeReceivables += cyVal;
      py.tradeReceivables += pyVal;
      cy.currentAssets += cyVal;
      py.currentAssets += pyVal;
    } else if (/inventory|stock|finished.*goods|work.*in.*progress/i.test(text)) {
      cy.inventories += cyVal;
      py.inventories += pyVal;
      cy.currentAssets += cyVal;
      py.currentAssets += pyVal;
    } else if (/current.*asset|loans.*advances/i.test(text)) {
      cy.otherCurrentAssets += cyVal;
      py.otherCurrentAssets += pyVal;
      cy.currentAssets += cyVal;
      py.currentAssets += pyVal;
    } else if (/plant|machinery|equipment|fixed.*asset|property|ppe|building/i.test(text)) {
      cy.fixedAssetsNet += cyVal;
      py.fixedAssetsNet += pyVal;
    } else if (/creditor|payable|trade.*payable/i.test(text)) {
      cy.tradePayables += cyVal;
      py.tradePayables += pyVal;
      cy.currentLiabilities += cyVal;
      py.currentLiabilities += pyVal;
    } else if (/short.*term.*borrow|overdraft|cc.*limit/i.test(text)) {
      cy.shortTermBorrowings += cyVal;
      py.shortTermBorrowings += pyVal;
      cy.currentLiabilities += cyVal;
      py.currentLiabilities += pyVal;
      cy.totalDebt += cyVal;
      py.totalDebt += pyVal;
    } else if (/term.*loan|long.*term.*borrow|debenture/i.test(text)) {
      cy.longTermDebt += cyVal;
      py.longTermDebt += pyVal;
      cy.totalDebt += cyVal;
      py.totalDebt += pyVal;
    } else if (/equity|share.*capital|reserves|surplus|net.*worth/i.test(text)) {
      cy.shareholdersEquity += cyVal;
      py.shareholdersEquity += pyVal;
    } else if (/operating.*expense|admin|selling|employee/i.test(text)) {
      cy.operatingExpenses += cyVal;
      py.operatingExpenses += pyVal;
    } else {
      matched = false;
      unmapped.push(name);
    }
  });

  // Auto-compute derived items if not explicitly provided
  if (cy.grossProfit === 0 && cy.revenue > 0) cy.grossProfit = cy.revenue - cy.cogs;
  if (py.grossProfit === 0 && py.revenue > 0) py.grossProfit = py.revenue - py.cogs;

  if (cy.operatingProfit === 0 && cy.grossProfit > 0) cy.operatingProfit = cy.grossProfit - cy.operatingExpenses;
  if (py.operatingProfit === 0 && py.grossProfit > 0) py.operatingProfit = py.grossProfit - py.operatingExpenses;

  if (cy.profitBeforeTax === 0 && cy.operatingProfit !== 0) cy.profitBeforeTax = cy.operatingProfit - cy.financeCosts;
  if (py.profitBeforeTax === 0 && py.operatingProfit !== 0) py.profitBeforeTax = py.operatingProfit - py.financeCosts;

  if (cy.profitAfterTax === 0 && cy.profitBeforeTax !== 0) cy.profitAfterTax = cy.profitBeforeTax - cy.taxExpense;
  if (py.profitAfterTax === 0 && py.profitBeforeTax !== 0) py.profitAfterTax = py.profitBeforeTax - py.taxExpense;

  cy.totalAssets = cy.currentAssets + cy.fixedAssetsNet;
  py.totalAssets = py.currentAssets + py.fixedAssetsNet;

  return { cy, py, unmapped };
}

/**
 * Calculates 10 standard CA financial ratios with full formula transparency
 */
export function computeFinancialRatios(
  cy: FinancialInputs,
  py: FinancialInputs,
  clientName = 'Corporate Client'
): RatioAnalysisResult {
  const alerts: string[] = [];
  const ratios: RatioResult[] = [];

  // Helper safe divide
  const safeDivide = (num: number, den: number, label: string): number | null => {
    if (den === 0 || isNaN(den)) {
      alerts.push(`${label}: Denominator is zero or missing. Calculation skipped.`);
      return null;
    }
    return num / den;
  };

  // 1. Current Ratio
  const cyCr = safeDivide(cy.currentAssets, cy.currentLiabilities, 'Current Ratio (CY)');
  const pyCr = safeDivide(py.currentAssets, py.currentLiabilities, 'Current Ratio (PY)');
  ratios.push({
    id: 'current_ratio',
    category: 'Liquidity',
    name: 'Current Ratio',
    formulaDescription: 'Current Assets / Current Liabilities',
    cyValue: cyCr,
    pyValue: pyCr,
    unit: 'ratio',
    benchmark: '2.0 : 1',
    interpretation: cyCr !== null && cyCr >= 1.33 ? (cyCr >= 2.0 ? 'healthy' : 'neutral') : 'concern',
    notes: 'Measures ability to pay short-term obligations. Benchmark 2:1 is standard under Tandon Committee / Bank lending norms.',
    cyFormulaExpanded: `₹${cy.currentAssets.toLocaleString('en-IN')} / ₹${cy.currentLiabilities.toLocaleString('en-IN')}`,
    pyFormulaExpanded: `₹${py.currentAssets.toLocaleString('en-IN')} / ₹${py.currentLiabilities.toLocaleString('en-IN')}`,
  });

  // 2. Quick Ratio (Acid Test)
  const cyQa = cy.currentAssets - cy.inventories;
  const pyQa = py.currentAssets - py.inventories;
  const cyQr = safeDivide(cyQa, cy.currentLiabilities, 'Quick Ratio (CY)');
  const pyQr = safeDivide(pyQa, py.currentLiabilities, 'Quick Ratio (PY)');
  ratios.push({
    id: 'quick_ratio',
    category: 'Liquidity',
    name: 'Quick Ratio (Acid Test)',
    formulaDescription: '(Current Assets - Inventory) / Current Liabilities',
    cyValue: cyQr,
    pyValue: pyQr,
    unit: 'ratio',
    benchmark: '1.0 : 1',
    interpretation: cyQr !== null && cyQr >= 1.0 ? 'healthy' : cyQr !== null && cyQr >= 0.8 ? 'neutral' : 'concern',
    notes: 'Excludes illiquid inventory to test immediate liquidity. Standard benchmark is 1:1.',
    cyFormulaExpanded: `(₹${cy.currentAssets.toLocaleString('en-IN')} - ₹${cy.inventories.toLocaleString('en-IN')}) / ₹${cy.currentLiabilities.toLocaleString('en-IN')}`,
    pyFormulaExpanded: `(₹${py.currentAssets.toLocaleString('en-IN')} - ₹${py.inventories.toLocaleString('en-IN')}) / ₹${py.currentLiabilities.toLocaleString('en-IN')}`,
  });

  // 3. Debt-to-Equity Ratio
  let cyDe: number | null = null;
  let pyDe: number | null = null;
  let cyDeFormula = `₹${cy.totalDebt.toLocaleString('en-IN')} / ₹${cy.shareholdersEquity.toLocaleString('en-IN')}`;
  let pyDeFormula = `₹${py.totalDebt.toLocaleString('en-IN')} / ₹${py.shareholdersEquity.toLocaleString('en-IN')}`;

  if (cy.shareholdersEquity <= 0) {
    alerts.push('Debt-to-Equity: Shareholders equity is negative or eroded, indicating capital impairment.');
    cyDeFormula = `Negative Net Worth (₹${cy.shareholdersEquity.toLocaleString('en-IN')}) - D/E not meaningful`;
  } else {
    cyDe = safeDivide(cy.totalDebt, cy.shareholdersEquity, 'Debt-Equity (CY)');
  }

  if (py.shareholdersEquity <= 0) {
    pyDeFormula = `Negative Net Worth (₹${py.shareholdersEquity.toLocaleString('en-IN')}) - D/E not meaningful`;
  } else {
    pyDe = safeDivide(py.totalDebt, py.shareholdersEquity, 'Debt-Equity (PY)');
  }

  ratios.push({
    id: 'debt_to_equity',
    category: 'Solvency',
    name: 'Debt-to-Equity Ratio',
    formulaDescription: 'Total Debt / Shareholders Equity (Net Worth)',
    cyValue: cyDe,
    pyValue: pyDe,
    unit: 'ratio',
    benchmark: '< 2.0 : 1',
    interpretation: cy.shareholdersEquity <= 0 ? 'concern' : cyDe !== null && cyDe <= 2.0 && cyDe >= 0 ? 'healthy' : 'concern',
    notes: cy.shareholdersEquity <= 0
      ? 'Critical Concern: Net worth is completely eroded by accumulated losses.'
      : 'Assesses financial leverage. A ratio exceeding 2:1 indicates high debt gearing and interest burden.',
    cyFormulaExpanded: cyDeFormula,
    pyFormulaExpanded: pyDeFormula,
  });

  // 4. Gross Profit Ratio
  const cyGpRatio = safeDivide(cy.grossProfit * 100, cy.revenue, 'Gross Profit Ratio (CY)');
  const pyGpRatio = safeDivide(py.grossProfit * 100, py.revenue, 'Gross Profit Ratio (PY)');
  ratios.push({
    id: 'gross_profit_ratio',
    category: 'Profitability',
    name: 'Gross Profit Ratio',
    formulaDescription: '(Gross Profit / Net Revenue) * 100',
    cyValue: cyGpRatio,
    pyValue: pyGpRatio,
    unit: 'percentage',
    benchmark: '> 20.0%',
    interpretation: cyGpRatio !== null && cyGpRatio >= 20 ? 'healthy' : cyGpRatio !== null && cyGpRatio >= 10 ? 'neutral' : 'concern',
    notes: 'Reflects core manufacturing and procurement markup efficiency before operating expenses.',
    cyFormulaExpanded: `(₹${cy.grossProfit.toLocaleString('en-IN')} / ₹${cy.revenue.toLocaleString('en-IN')}) * 100`,
    pyFormulaExpanded: `(₹${py.grossProfit.toLocaleString('en-IN')} / ₹${py.revenue.toLocaleString('en-IN')}) * 100`,
  });

  // 5. Net Profit Ratio
  const cyNpRatio = safeDivide(cy.profitAfterTax * 100, cy.revenue, 'Net Profit Ratio (CY)');
  const pyNpRatio = safeDivide(py.profitAfterTax * 100, py.revenue, 'Net Profit Ratio (PY)');
  ratios.push({
    id: 'net_profit_ratio',
    category: 'Profitability',
    name: 'Net Profit Ratio',
    formulaDescription: '(Profit After Tax / Net Revenue) * 100',
    cyValue: cyNpRatio,
    pyValue: pyNpRatio,
    unit: 'percentage',
    benchmark: '> 5.0%',
    interpretation: cyNpRatio !== null && cyNpRatio >= 8 ? 'healthy' : cyNpRatio !== null && cyNpRatio > 0 ? 'neutral' : 'concern',
    notes: 'Indicates bottom-line profitability after all operating costs, interest, depreciation, and tax.',
    cyFormulaExpanded: `(₹${cy.profitAfterTax.toLocaleString('en-IN')} / ₹${cy.revenue.toLocaleString('en-IN')}) * 100`,
    pyFormulaExpanded: `(₹${py.profitAfterTax.toLocaleString('en-IN')} / ₹${py.revenue.toLocaleString('en-IN')}) * 100`,
  });

  // 6. Operating Profit Ratio (EBIT Margin)
  const cyOpRatio = safeDivide(cy.operatingProfit * 100, cy.revenue, 'Operating Profit Ratio (CY)');
  const pyOpRatio = safeDivide(py.operatingProfit * 100, py.revenue, 'Operating Profit Ratio (PY)');
  ratios.push({
    id: 'operating_profit_ratio',
    category: 'Profitability',
    name: 'Operating Profit Ratio (EBIT Margin)',
    formulaDescription: '(Operating Profit [EBIT] / Net Revenue) * 100',
    cyValue: cyOpRatio,
    pyValue: pyOpRatio,
    unit: 'percentage',
    benchmark: '> 10.0%',
    interpretation: cyOpRatio !== null && cyOpRatio >= 10 ? 'healthy' : cyOpRatio !== null && cyOpRatio > 0 ? 'neutral' : 'concern',
    notes: 'Measures operational profitability independent of financing structure and income taxes.',
    cyFormulaExpanded: `(₹${cy.operatingProfit.toLocaleString('en-IN')} / ₹${cy.revenue.toLocaleString('en-IN')}) * 100`,
    pyFormulaExpanded: `(₹${py.operatingProfit.toLocaleString('en-IN')} / ₹${py.revenue.toLocaleString('en-IN')}) * 100`,
  });

  // 7. Inventory Turnover Ratio
  const cyItr = safeDivide(cy.cogs > 0 ? cy.cogs : cy.revenue * 0.7, cy.inventories, 'Inventory Turnover (CY)');
  const pyItr = safeDivide(py.cogs > 0 ? py.cogs : py.revenue * 0.7, py.inventories, 'Inventory Turnover (PY)');
  ratios.push({
    id: 'inventory_turnover',
    category: 'Activity / Turnover',
    name: 'Inventory Turnover Ratio',
    formulaDescription: 'Cost of Goods Sold / Closing Inventory',
    cyValue: cyItr,
    pyValue: pyItr,
    unit: 'times',
    benchmark: '> 4.0 Times',
    interpretation: cyItr !== null && cyItr >= 4 ? 'healthy' : cyItr !== null && cyItr >= 2 ? 'neutral' : 'concern',
    notes: `Velocity at which inventory is sold. Equivalent to ${cyItr && cyItr > 0 ? Math.round(365 / cyItr) : 0} days of inventory holding.`,
    cyFormulaExpanded: `₹${(cy.cogs || cy.revenue * 0.7).toLocaleString('en-IN')} / ₹${cy.inventories.toLocaleString('en-IN')}`,
    pyFormulaExpanded: `₹${(py.cogs || py.revenue * 0.7).toLocaleString('en-IN')} / ₹${py.inventories.toLocaleString('en-IN')}`,
  });

  // 8. Receivables (Debtors) Turnover Ratio
  const cyRtr = safeDivide(cy.revenue, cy.tradeReceivables, 'Receivables Turnover (CY)');
  const pyRtr = safeDivide(py.revenue, py.tradeReceivables, 'Receivables Turnover (PY)');
  ratios.push({
    id: 'receivables_turnover',
    category: 'Activity / Turnover',
    name: 'Receivables Turnover Ratio',
    formulaDescription: 'Net Credit Revenue / Closing Trade Receivables',
    cyValue: cyRtr,
    pyValue: pyRtr,
    unit: 'times',
    benchmark: '> 6.0 Times',
    interpretation: cyRtr !== null && cyRtr >= 6 ? 'healthy' : cyRtr !== null && cyRtr >= 3 ? 'neutral' : 'concern',
    notes: `Frequency of debtor collections. Equivalent to average collection period of ${cyRtr && cyRtr > 0 ? Math.round(365 / cyRtr) : 0} days.`,
    cyFormulaExpanded: `₹${cy.revenue.toLocaleString('en-IN')} / ₹${cy.tradeReceivables.toLocaleString('en-IN')}`,
    pyFormulaExpanded: `₹${py.revenue.toLocaleString('en-IN')} / ₹${py.tradeReceivables.toLocaleString('en-IN')}`,
  });

  // 9. Interest Coverage Ratio
  const isCyDebtFree = cy.financeCosts <= 0;
  const isPyDebtFree = py.financeCosts <= 0;
  const cyIcr = isCyDebtFree ? null : safeDivide(cy.operatingProfit, cy.financeCosts, 'Interest Coverage (CY)');
  const pyIcr = isPyDebtFree ? null : safeDivide(py.operatingProfit, py.financeCosts, 'Interest Coverage (PY)');
  ratios.push({
    id: 'interest_coverage',
    category: 'Solvency',
    name: 'Interest Coverage Ratio',
    formulaDescription: 'Operating Profit (EBIT) / Finance Costs',
    cyValue: cyIcr,
    pyValue: pyIcr,
    unit: 'times',
    benchmark: '> 2.5 Times',
    interpretation: isCyDebtFree ? 'healthy' : cyIcr !== null && cyIcr >= 2.5 ? 'healthy' : cyIcr !== null && cyIcr >= 1.5 ? 'neutral' : 'concern',
    notes: isCyDebtFree
      ? 'Company is debt-free with zero interest servicing obligations.'
      : 'Crucial metric for bank loan sanctions. Indicates ease of servicing loan interest from operating profits.',
    cyFormulaExpanded: isCyDebtFree ? 'Debt-Free (Nil Finance Costs)' : `₹${cy.operatingProfit.toLocaleString('en-IN')} / ₹${cy.financeCosts.toLocaleString('en-IN')}`,
    pyFormulaExpanded: isPyDebtFree ? 'Debt-Free (Nil Finance Costs)' : `₹${py.operatingProfit.toLocaleString('en-IN')} / ₹${py.financeCosts.toLocaleString('en-IN')}`,
  });

  // 10. Return on Equity (ROE)
  let cyRoe: number | null = null;
  let pyRoe: number | null = null;
  let cyRoeFormula = `(₹${cy.profitAfterTax.toLocaleString('en-IN')} / ₹${cy.shareholdersEquity.toLocaleString('en-IN')}) * 100`;
  let pyRoeFormula = `(₹${py.profitAfterTax.toLocaleString('en-IN')} / ₹${py.shareholdersEquity.toLocaleString('en-IN')}) * 100`;

  if (cy.shareholdersEquity <= 0) {
    alerts.push('Return on Equity: Shareholders equity is negative. ROE is not mathematically applicable.');
    cyRoeFormula = `Negative Net Worth (₹${cy.shareholdersEquity.toLocaleString('en-IN')}) - ROE not applicable`;
  } else {
    cyRoe = safeDivide(cy.profitAfterTax * 100, cy.shareholdersEquity, 'Return on Equity (CY)');
  }

  if (py.shareholdersEquity <= 0) {
    pyRoeFormula = `Negative Net Worth (₹${py.shareholdersEquity.toLocaleString('en-IN')}) - ROE not applicable`;
  } else {
    pyRoe = safeDivide(py.profitAfterTax * 100, py.shareholdersEquity, 'Return on Equity (PY)');
  }

  ratios.push({
    id: 'return_on_equity',
    category: 'Profitability',
    name: 'Return on Equity (ROE)',
    formulaDescription: '(Profit After Tax / Shareholders Equity) * 100',
    cyValue: cyRoe,
    pyValue: pyRoe,
    unit: 'percentage',
    benchmark: '> 15.0%',
    interpretation: cy.shareholdersEquity <= 0 ? 'concern' : cyRoe !== null && cyRoe >= 15 ? 'healthy' : cyRoe !== null && cyRoe >= 8 ? 'neutral' : 'concern',
    notes: cy.shareholdersEquity <= 0
      ? 'Critical Concern: Net worth is negative due to accumulated losses. ROE cannot be computed.'
      : 'Indicates the return generated on shareholders net capital investment.',
    cyFormulaExpanded: cyRoeFormula,
    pyFormulaExpanded: pyRoeFormula,
  });

  return {
    clientName,
    cyPeriod: cy.periodLabel,
    pyPeriod: py.periodLabel,
    cyInputs: cy,
    pyInputs: py,
    ratios,
    alerts,
  };
}

/**
 * Exports complete multi-sheet Excel financial ratio analysis workbook
 */
export function exportRatioWorkbook(
  result: RatioAnalysisResult
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Executive Ratio Dashboard
  const summaryRows = result.ratios.map((r, idx) => ({
    '#': idx + 1,
    'Category': r.category,
    'Financial Ratio': r.name,
    'Formula Description': r.formulaDescription,
    [`${result.cyPeriod} (CY)`]: r.cyValue !== null ? `${r.cyValue.toFixed(2)}${r.unit === 'percentage' ? '%' : ''}` : 'N/A',
    [`${result.pyPeriod} (PY)`]: r.pyValue !== null ? `${r.pyValue.toFixed(2)}${r.unit === 'percentage' ? '%' : ''}` : 'N/A',
    'Standard Benchmark': r.benchmark,
    'Auditor Rating': r.interpretation.toUpperCase(),
    'Analytical Significance': r.notes,
  }));
  const wsSummary = createStyledWorksheet(summaryRows, {
    columnWidths: {
      '#': 6,
      'Category': 16,
      'Financial Ratio': 28,
      'Formula Description': 36,
      [`${result.cyPeriod} (CY)`]: 16,
      [`${result.pyPeriod} (PY)`]: 16,
      'Standard Benchmark': 18,
      'Auditor Rating': 16,
      'Analytical Significance': 55,
    },
  });
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Executive_Ratio_Summary');

  // Sheet 2: Formula & Calculation Transparency
  const formulaRows = result.ratios.map((r) => ({
    'Ratio Name': r.name,
    'CY Substituted Formula': r.cyFormulaExpanded,
    'CY Result': r.cyValue !== null ? r.cyValue.toFixed(2) : 'N/A',
    'PY Substituted Formula': r.pyFormulaExpanded,
    'PY Result': r.pyValue !== null ? r.pyValue.toFixed(2) : 'N/A',
  }));
  const wsFormula = createStyledWorksheet(formulaRows, {
    columnWidths: {
      'Ratio Name': 28,
      'CY Substituted Formula': 45,
      'CY Result': 14,
      'PY Substituted Formula': 45,
      'PY Result': 14,
    },
  });
  XLSX.utils.book_append_sheet(wb, wsFormula, 'Formula_Calculations');

  // Sheet 3: Financial Statement Inputs
  const inputRows = [
    { 'Financial Component': 'Revenue from Operations (Sales)', [`${result.cyPeriod} (₹)`]: result.cyInputs.revenue, [`${result.pyPeriod} (₹)`]: result.pyInputs.revenue },
    { 'Financial Component': 'Cost of Goods Sold (COGS)', [`${result.cyPeriod} (₹)`]: result.cyInputs.cogs, [`${result.pyPeriod} (₹)`]: result.pyInputs.cogs },
    { 'Financial Component': 'Gross Profit', [`${result.cyPeriod} (₹)`]: result.cyInputs.grossProfit, [`${result.pyPeriod} (₹)`]: result.pyInputs.grossProfit },
    { 'Financial Component': 'Operating Expenses (Admin & Selling)', [`${result.cyPeriod} (₹)`]: result.cyInputs.operatingExpenses, [`${result.pyPeriod} (₹)`]: result.pyInputs.operatingExpenses },
    { 'Financial Component': 'Operating Profit (EBIT)', [`${result.cyPeriod} (₹)`]: result.cyInputs.operatingProfit, [`${result.pyPeriod} (₹)`]: result.pyInputs.operatingProfit },
    { 'Financial Component': 'Finance Costs / Interest', [`${result.cyPeriod} (₹)`]: result.cyInputs.financeCosts, [`${result.pyPeriod} (₹)`]: result.pyInputs.financeCosts },
    { 'Financial Component': 'Depreciation & Amortisation', [`${result.cyPeriod} (₹)`]: result.cyInputs.depreciation, [`${result.pyPeriod} (₹)`]: result.pyInputs.depreciation },
    { 'Financial Component': 'Profit After Tax (Net Profit)', [`${result.cyPeriod} (₹)`]: result.cyInputs.profitAfterTax, [`${result.pyPeriod} (₹)`]: result.pyInputs.profitAfterTax },
    { 'Financial Component': 'Current Assets', [`${result.cyPeriod} (₹)`]: result.cyInputs.currentAssets, [`${result.pyPeriod} (₹)`]: result.pyInputs.currentAssets },
    { 'Financial Component': 'Inventories / Stock in Trade', [`${result.cyPeriod} (₹)`]: result.cyInputs.inventories, [`${result.pyPeriod} (₹)`]: result.pyInputs.inventories },
    { 'Financial Component': 'Trade Receivables / Debtors', [`${result.cyPeriod} (₹)`]: result.cyInputs.tradeReceivables, [`${result.pyPeriod} (₹)`]: result.pyInputs.tradeReceivables },
    { 'Financial Component': 'Cash and Bank Balances', [`${result.cyPeriod} (₹)`]: result.cyInputs.cashAndBank, [`${result.pyPeriod} (₹)`]: result.pyInputs.cashAndBank },
    { 'Financial Component': 'Current Liabilities & Provisions', [`${result.cyPeriod} (₹)`]: result.cyInputs.currentLiabilities, [`${result.pyPeriod} (₹)`]: result.pyInputs.currentLiabilities },
    { 'Financial Component': 'Trade Payables / Creditors', [`${result.cyPeriod} (₹)`]: result.cyInputs.tradePayables, [`${result.pyPeriod} (₹)`]: result.pyInputs.tradePayables },
    { 'Financial Component': 'Total Debt (Borrowings)', [`${result.cyPeriod} (₹)`]: result.cyInputs.totalDebt, [`${result.pyPeriod} (₹)`]: result.pyInputs.totalDebt },
    { 'Financial Component': 'Shareholders Equity (Net Worth)', [`${result.cyPeriod} (₹)`]: result.cyInputs.shareholdersEquity, [`${result.pyPeriod} (₹)`]: result.pyInputs.shareholdersEquity },
  ];
  const wsInputs = createStyledWorksheet(inputRows, {
    currencyColumns: [`${result.cyPeriod} (₹)`, `${result.pyPeriod} (₹)`],
  });
  XLSX.utils.book_append_sheet(wb, wsInputs, 'Statement_Inputs');

  return wb;
}

/**
 * High-quality comparative sample financial statements
 */
export function getSampleRatioRows(): Record<string, any>[] {
  return [
    { 'Particulars': 'Revenue from Operations (Domestic Sales)', 'Category': 'Revenue', 'FY 2026-27 (Current)': 48500000, 'FY 2025-26 (Prior)': 42000000 },
    { 'Particulars': 'Cost of Materials Consumed (COGS)', 'Category': 'Direct Expense', 'FY 2026-27 (Current)': 31200000, 'FY 2025-26 (Prior)': 27800000 },
    { 'Particulars': 'Employee Benefit Expenses', 'Category': 'Operating Expense', 'FY 2026-27 (Current)': 4800000, 'FY 2025-26 (Prior)': 4100000 },
    { 'Particulars': 'Other Operating & Administrative Expenses', 'Category': 'Operating Expense', 'FY 2026-27 (Current)': 3400000, 'FY 2025-26 (Prior)': 3100000 },
    { 'Particulars': 'Finance Costs (Bank Loan Interest)', 'Category': 'Finance Cost', 'FY 2026-27 (Current)': 1250000, 'FY 2025-26 (Prior)': 1400000 },
    { 'Particulars': 'Depreciation and Amortisation Expense', 'Category': 'Depreciation', 'FY 2026-27 (Current)': 1650000, 'FY 2025-26 (Prior)': 1500000 },
    { 'Particulars': 'Income Tax Expense', 'Category': 'Tax Expense', 'FY 2026-27 (Current)': 1550000, 'FY 2025-26 (Prior)': 1025000 },
    { 'Particulars': 'Inventories / Stock in Trade', 'Category': 'Current Asset', 'FY 2026-27 (Current)': 6200000, 'FY 2025-26 (Prior)': 5400000 },
    { 'Particulars': 'Trade Receivables / Debtors', 'Category': 'Current Asset', 'FY 2026-27 (Current)': 7100000, 'FY 2025-26 (Prior)': 6800000 },
    { 'Particulars': 'Cash and Bank Balances', 'Category': 'Current Asset', 'FY 2026-27 (Current)': 2850000, 'FY 2025-26 (Prior)': 2100000 },
    { 'Particulars': 'Other Current Assets & GST Input Credits', 'Category': 'Current Asset', 'FY 2026-27 (Current)': 1450000, 'FY 2025-26 (Prior)': 1200000 },
    { 'Particulars': 'Property, Plant & Equipment (Net PPE)', 'Category': 'Fixed Asset', 'FY 2026-27 (Current)': 18500000, 'FY 2025-26 (Prior)': 17800000 },
    { 'Particulars': 'Trade Payables / Sundry Creditors', 'Category': 'Current Liability', 'FY 2026-27 (Current)': 5400000, 'FY 2025-26 (Prior)': 5100000 },
    { 'Particulars': 'Short-Term Borrowings (Working Capital CC)', 'Category': 'Current Liability', 'FY 2026-27 (Current)': 3200000, 'FY 2025-26 (Prior)': 3600000 },
    { 'Particulars': 'Long-Term Borrowings (Term Loans)', 'Category': 'Non-Current Liability', 'FY 2026-27 (Current)': 6500000, 'FY 2025-26 (Prior)': 7800000 },
    { 'Particulars': 'Shareholders Equity (Capital & Reserves)', 'Category': 'Equity', 'FY 2026-27 (Current)': 21100000, 'FY 2025-26 (Prior)': 16400000 },
  ];
}
