import { UploadIllustration } from '../common/UploadIllustration';
import { InnerToolBadge } from '../common/InnerToolBadge';
import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  Download,
  UploadCloud,
  CheckCircle2,
  AlertTriangle,
  FileBarChart,
  RefreshCw,
  Plus,
  Trash2,
  Check,
  X,
  FileText,
  Sliders,
  Scale,
  ListOrdered,
  AlertCircle,
} from 'lucide-react';
import {
  parseTrialBalanceRows,
  compileFinancialStatements,
  createFinancialStatementsWorkbook,
  type FinancialStatementsResult,
  type TrialBalanceRecord,
  type AuditAdjustmentEntry,
} from '../../utils/financialReportsEngine';
import { ToolIcon } from '../common/ToolIcon';
import { DataFormatGuide } from '../common/DataFormatGuide';
import { formatINR, extractSmartTableFromSheet } from '../../utils/dataSanitizer';

interface FinancialReportsToolProps {
  onBack: () => void;
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
}

export const FinancialReportsTool: React.FC<FinancialReportsToolProps> = ({ onBack, onShowToast }) => {
  const [file, setFile] = useState<{ name: string; size: string } | null>(null);
  const [tbRecords, setTbRecords] = useState<TrialBalanceRecord[]>([]);
  const [statementsResult, setStatementsResult] = useState<FinancialStatementsResult | null>(null);
  const [reportType, setReportType] = useState<'bs' | 'pl' | 'tb' | 'adj' | 'abnormal'>('bs');
  const [financialYear, setFinancialYear] = useState('FY 2026-27');
  const [isGenerated, setIsGenerated] = useState(false);
  const [validationWarning, setValidationWarning] = useState<string | null>(null);

  // Dynamic Audit Adjustments
  const [adjustments, setAdjustments] = useState<AuditAdjustmentEntry[]>([]);

  // Form for adding new adjustment
  const [newDrAccount, setNewDrAccount] = useState('');
  const [newCrAccount, setNewCrAccount] = useState('');
  const [newAmount, setNewAmount] = useState('');
  const [newNarration, setNewNarration] = useState('');
  const [newCategory, setNewCategory] = useState<AuditAdjustmentEntry['category']>('Accrual');

  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;
    setValidationWarning(null);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const buffer = evt.target?.result as ArrayBuffer;
        const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
        const smartTable = extractSmartTableFromSheet(wb.Sheets[wb.SheetNames[0]]);
        const json = smartTable.rows;

        if (!json || json.length === 0) {
          setValidationWarning('Trial balance spreadsheet contains zero account rows.');
          onShowToast('Empty File', 'No rows found in uploaded Trial Balance.', 'error');
          return;
        }

        const parsed = parseTrialBalanceRows(json);
        setFile({ name: f.name, size: `${Math.round(f.size / 1024)} KB` });
        setTbRecords(parsed);

        const res = compileFinancialStatements(parsed, adjustments);
        setStatementsResult(res);
        setIsGenerated(true);
      } catch (err: any) {
        setValidationWarning(`Failed to parse file: ${err.message}`);
        onShowToast('Parse Error', 'Could not read Trial Balance spreadsheet.', 'error');
      }
    };
    reader.readAsArrayBuffer(f);
  };

  const handleGenerate = () => {
    if (tbRecords.length === 0) {
      return;
    }
    const records = tbRecords;

    const res = compileFinancialStatements(records, adjustments);
    setStatementsResult(res);
    setIsGenerated(true);

    if (!res.isBalanced) {
      setValidationWarning(
        `Trial Balance Imbalance: Total Debits (${formatINR(res.totalDebits)}) differ from Total Credits (${formatINR(res.totalCredits)}) by ${formatINR(res.trialBalanceDiff)}.`
      );
    } else {
      setValidationWarning(null);
    }

    onShowToast(
      'Statements Compiled',
      `Schedule III Net Profit: ${formatINR(res.scheduleIII.netProfitBeforeTax)} • Total Assets: ${formatINR(res.scheduleIII.totalAssets)}.`,
      'success'
    );
  };

  const handleAddAdjustment = (e: React.FormEvent) => {
    e.preventDefault();
    const amt = parseFloat(newAmount);
    if (!newDrAccount || !newCrAccount || isNaN(amt) || amt <= 0) {
      onShowToast('Invalid Adjustment', 'Please provide valid Dr account, Cr account, and positive amount.', 'error');
      return;
    }

    const newAdj: AuditAdjustmentEntry = {
      id: `adj-${Date.now()}`,
      debitAccountCode: `ADJ-${adjustments.length + 1}D`,
      debitAccountName: newDrAccount.trim(),
      creditAccountCode: `ADJ-${adjustments.length + 1}C`,
      creditAccountName: newCrAccount.trim(),
      amount: amt,
      narration: newNarration.trim() || 'Auditor adjustment entry',
      category: newCategory,
      isApproved: true,
    };

    const updated = [...adjustments, newAdj];
    setAdjustments(updated);

    if (tbRecords.length > 0) {
      const res = compileFinancialStatements(tbRecords, updated);
      setStatementsResult(res);
    }

    setNewDrAccount('');
    setNewCrAccount('');
    setNewAmount('');
    setNewNarration('');
    onShowToast('Adjustment Added & Applied', `Recalculated financial statements with ₹${amt} adjustment.`, 'success');
  };

  const handleToggleAdjustment = (id: string) => {
    const updated = adjustments.map((a) => (a.id === id ? { ...a, isApproved: !a.isApproved } : a));
    setAdjustments(updated);
    if (tbRecords.length > 0) {
      const res = compileFinancialStatements(tbRecords, updated);
      setStatementsResult(res);
    }
  };

  const handleDeleteAdjustment = (id: string) => {
    const updated = adjustments.filter((a) => a.id !== id);
    setAdjustments(updated);
    if (tbRecords.length > 0) {
      const res = compileFinancialStatements(tbRecords, updated);
      setStatementsResult(res);
      onShowToast('Adjustment Removed', 'Recalculated statements without the adjustment.', 'info');
    }
  };

  const handleDownload = () => {
    if (!statementsResult) {
      return;
    }
    const wb = createFinancialStatementsWorkbook(statementsResult, financialYear);
    const outName = `Financial_Statements_Schedule_III_${financialYear.replace(/\s+/g, '_')}.xlsx`;
    XLSX.writeFile(wb, outName);
    onShowToast('Excel Workbook Downloaded', `Saved 5-sheet workbook: Balance Sheet, P&L, Adjusted TB, Adjustments Log, and Abnormal Balances.`, 'success');
  };

  const currentResult = statementsResult;
  const s3 = statementsResult?.scheduleIII;

  return (
    <div className="space-y-6">
      {/* 1. Header */}
      <div>
        <button
          onClick={onBack}
          className="inline-flex items-center gap-2 px-3 py-1.5 rounded-lg text-xs font-semibold text-gray-700 bg-white border border-gray-200 shadow-2xs hover:bg-gray-50 hover:text-gray-900 hover:border-gray-300 transition-all cursor-pointer mb-5 group"
        >
          <ArrowLeft className="w-3.5 h-3.5 text-gray-400 group-hover:text-gray-900 group-hover:-translate-x-0.5 transition-transform" />
          <span>Back to All Tools</span>
        </button>

        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-5 border-b border-gray-200">
          <div className="flex items-center gap-3.5">
            <InnerToolBadge toolId="financial-reports" />
            <div>
              <h1 className="text-xl font-bold text-gray-900 tracking-tight">Trial Balance-to-Financial Statements Automation</h1>
              <p className="text-xs text-gray-500 mt-0.5">
                Generate Companies Act Schedule III Balance Sheet, Profit & Loss, audit adjustments worksheet, and abnormal balance detection.
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <button
              onClick={handleDownload}
              className="inline-flex items-center gap-2 px-4 py-2 bg-gray-900 hover:bg-black text-white rounded-lg text-xs font-semibold transition-colors shadow-sm cursor-pointer"
            >
              <Download className="w-4 h-4" />
              <span>Export 5-Sheet Financials (.xlsx)</span>
            </button>
          </div>
        </div>
      </div>

      {/* 2. Format Requirements Guide */}
      <DataFormatGuide
        toolType="trial-balance"
        acceptedFormats=".xlsx, .xls, .csv"
        requiredFields={[
          { name: 'Account Code', description: 'General ledger account code', aliases: ['Account_Code', 'Code', 'GL_Code'] },
          { name: 'Account / Particulars', description: 'Ledger account title', aliases: ['Account_Name', 'Particulars', 'Account'] },
          { name: 'Debit Amount', description: 'Debit closing balance (₹)', aliases: ['Debit', 'Dr', 'Debit_Amt'] },
          { name: 'Credit Amount', description: 'Credit closing balance (₹)', aliases: ['Credit', 'Cr', 'Credit_Amt'] },
          { name: 'Accounting Head (Optional)', description: 'Asset, Liability, Equity, Revenue, Expense', aliases: ['Head', 'Group', 'Category'] },
        ]}
        notes={[
          'Schedule III Automation: Intelligently classifies ledgers into Non-Current Assets, Current Assets, Shareholders Funds, and Liabilities.',
          'Audit Adjustment Engine: Add closing stock, depreciation, or accrual entries and watch the Balance Sheet & P&L recalculate dynamically.',
          'Abnormal Polarity Detection: Flags unexpected sign conventions (negative cash, credit debtors, debit creditors).',
          'Strict Financial Integrity: Never silently fabricates numbers or forces balance sheets to balance artificially.',
        ]}
      />

      {/* Warning banner */}
      {validationWarning && (
        <div className="p-4 rounded-xl bg-[#FEF9EE] border border-[#FDE68A] text-xs text-[#92400E] flex items-center justify-between">
          <div className="flex items-center gap-2">
            <AlertTriangle className="w-4 h-4 shrink-0" />
            <span>{validationWarning}</span>
          </div>
          <button onClick={() => setValidationWarning(null)} className="underline hover:no-underline font-semibold">
            Dismiss
          </button>
        </div>
      )}

      {/* Abnormal balances alert */}
      {currentResult && currentResult.abnormalAccounts.length > 0 && (
        <div className="p-3.5 rounded-xl bg-amber-50 border border-amber-200 text-xs text-amber-900 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <AlertCircle className="w-4 h-4 text-amber-600 shrink-0" />
            <span>
              <strong>{currentResult.abnormalAccounts.length} Abnormal Balances Identified:</strong> Detected credit balances in cash/debtors or abnormal expense polarity. Inspect in Abnormal Balances tab.
            </span>
          </div>
          <button
            onClick={() => setReportType('abnormal')}
            className="px-2.5 py-1 text-xs font-semibold bg-white border border-amber-300 rounded text-amber-800 hover:bg-amber-100 transition-colors"
          >
            Review Issues
          </button>
        </div>
      )}

      {/* 3. Upload & Parameter Controls */}
      <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-sm space-y-4">
        <input
          type="file"
          ref={fileInputRef}
          onChange={handleFileUpload}
          accept=".xlsx,.xls,.csv"
          className="hidden"
        />

        <div className="grid grid-cols-1 sm:grid-cols-4 gap-4">
          <div>
            <label className="text-xs font-semibold text-gray-900 block mb-1">Financial Period</label>
            <select
              value={financialYear}
              onChange={(e) => setFinancialYear(e.target.value)}
              className="w-full px-3 py-2 bg-gray-50 border border-gray-200 rounded-lg text-xs text-gray-900 font-medium focus:border-gray-300 cursor-pointer"
            >
              <option value="FY 2030-31">FY 2030-31</option>
              <option value="FY 2029-30">FY 2029-30</option>
              <option value="FY 2028-29">FY 2028-29</option>
              <option value="FY 2027-28">FY 2027-28</option>
              <option value="FY 2026-27">FY 2026-27 (Current FY)</option>
              <option value="FY 2025-26">FY 2025-26 (Previous Year)</option>
              <option value="FY 2024-25">FY 2024-25</option>
              <option value="FY 2023-24">FY 2023-24</option>
              <option value="FY 2022-23">FY 2022-23</option>
            </select>
          </div>

          <div className="sm:col-span-3 flex items-end gap-2 justify-end">
            <button
              onClick={() => fileInputRef.current?.click()}
              className="px-3.5 py-2 text-xs text-gray-900 bg-gray-50 hover:bg-[#EAEAEA] rounded-lg border border-gray-200 flex items-center gap-1.5 transition-colors font-medium"
            >
              <UploadCloud className="w-3.5 h-3.5 text-gray-600" />
              <span>{file ? 'Change TB' : 'Upload Client TB'}</span>
            </button>
            <button
              onClick={handleGenerate}
              className="px-4 py-2 text-xs font-semibold text-white bg-gray-900 hover:bg-black rounded-lg transition-colors shadow-sm flex items-center gap-1.5"
            >
              <RefreshCw className="w-3.5 h-3.5" />
              <span>Compile & Recalculate</span>
            </button>
          </div>
        </div>

        {file && (
          <div className="pt-3 border-t border-gray-200 flex items-center justify-between text-xs text-gray-600">
            <span>Active Trial Balance: <strong className="text-gray-900">{file.name}</strong> ({file.size}) • {tbRecords.length} accounts mapped</span>
          </div>
        )}
      </div>

      {!currentResult || !s3 ? (
        <div
          onClick={() => fileInputRef.current?.click()}
          className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-8 sm:p-10 text-center cursor-pointer transition-colors flex flex-col items-center justify-center shadow-2xs"
        >
          <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mx-auto mb-3 shadow-2xs">
            <UploadCloud className="w-6 h-6" />
          </div>
          <h3 className="text-sm font-semibold text-gray-900">Upload Client Trial Balance (.xlsx, .xls, .csv)</h3>
          <p className="text-xs text-gray-500 mt-1 max-w-md mx-auto leading-relaxed">
            Upload your general ledger Trial Balance with Account Code, Account Name, Debit, and Credit columns. The engine compiles Companies Act Schedule III Balance Sheet and Profit & Loss statements automatically.
          </p>
          <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
            <UploadCloud className="w-4 h-4" />
            <span>Upload Trial Balance File</span>
          </div>
          <span className="text-[11px] text-gray-400 font-medium mt-3 block">
            Supports .xlsx, .xls, .csv · Automatic grouping into Schedule III assets, liabilities, revenue & expenses
          </span>
        </div>
      ) : (
        <>
          {/* 4. Navigation View Tabs */}
          <div className="flex items-center gap-2 border-b border-gray-200 pb-1 text-xs">
            <button
              onClick={() => setReportType('bs')}
              className={`px-3.5 py-1.5 font-semibold rounded-t-lg transition-colors flex items-center gap-1.5 ${
                reportType === 'bs'
                  ? 'bg-white text-gray-900 border-t-2 border-gray-300 shadow-xs'
                  : 'text-gray-600 hover:text-gray-900'
              }`}
            >
              <Scale className="w-3.5 h-3.5" />
              <span>Schedule III Balance Sheet</span>
            </button>

        <button
          onClick={() => setReportType('pl')}
          className={`px-3.5 py-1.5 font-semibold rounded-t-lg transition-colors flex items-center gap-1.5 ${
            reportType === 'pl'
              ? 'bg-white text-gray-900 border-t-2 border-gray-300 shadow-xs'
              : 'text-gray-600 hover:text-gray-900'
          }`}
        >
          <FileBarChart className="w-3.5 h-3.5" />
          <span>Profit & Loss Statement</span>
        </button>

        <button
          onClick={() => setReportType('adj')}
          className={`px-3.5 py-1.5 font-semibold rounded-t-lg transition-colors flex items-center gap-1.5 ${
            reportType === 'adj'
              ? 'bg-white text-gray-900 border-t-2 border-gray-300 shadow-xs'
              : 'text-gray-600 hover:text-gray-900'
          }`}
        >
          <Sliders className="w-3.5 h-3.5" />
          <span>Audit Adjustments ({adjustments.length})</span>
        </button>

        <button
          onClick={() => setReportType('tb')}
          className={`px-3.5 py-1.5 font-semibold rounded-t-lg transition-colors flex items-center gap-1.5 ${
            reportType === 'tb'
              ? 'bg-white text-gray-900 border-t-2 border-gray-300 shadow-xs'
              : 'text-gray-600 hover:text-gray-900'
          }`}
        >
          <ListOrdered className="w-3.5 h-3.5" />
          <span>Adjusted Trial Balance</span>
        </button>

        <button
          onClick={() => setReportType('abnormal')}
          className={`px-3.5 py-1.5 font-semibold rounded-t-lg transition-colors flex items-center gap-1.5 ${
            reportType === 'abnormal'
              ? 'bg-white text-gray-900 border-t-2 border-gray-300 shadow-xs'
              : 'text-gray-600 hover:text-gray-900'
          }`}
        >
          <AlertCircle className={`w-3.5 h-3.5 ${currentResult.abnormalAccounts.length > 0 ? 'text-amber-600' : 'text-gray-400'}`} />
          <span>Abnormal Balances ({currentResult.abnormalAccounts.length})</span>
        </button>
      </div>

      {/* 5. Statement Render Workspace */}
      <div className="bg-white rounded-xl border border-gray-200 overflow-hidden shadow-sm">
        {/* VIEW 1: Schedule III Balance Sheet */}
        {reportType === 'bs' && (
          <div className="divide-y divide-[#E5E7EB] text-xs">
            <div className="p-4 bg-gray-50 font-bold text-gray-900 flex justify-between">
              <span>I. EQUITY AND LIABILITIES</span>
              <span>Amount (₹)</span>
            </div>
            <div className="p-3.5 flex justify-between items-center pl-8">
              <span>(1) Shareholders Funds: Equity Share Capital</span>
              <span className="font-mono font-medium">{formatINR(s3.shareCapital)}</span>
            </div>
            <div className="p-3.5 flex justify-between items-center pl-8 bg-gray-50/20">
              <span className="text-gray-900 font-semibold">
                (2) Reserves & Surplus (including Net Profit for the Period)
              </span>
              <span className="font-mono font-bold text-gray-900">
                {formatINR(s3.reservesAndSurplus)}
              </span>
            </div>
            <div className="p-3.5 flex justify-between items-center pl-8">
              <span>(3) Non-Current Liabilities (Long-Term Borrowings & Provisions)</span>
              <span className="font-mono font-medium">{formatINR(s3.nonCurrentLiabilities)}</span>
            </div>
            <div className="p-3.5 flex justify-between items-center pl-8">
              <span>(4) Trade Payables (Sundry Creditors for Goods & Services)</span>
              <span className="font-mono font-medium">{formatINR(s3.tradePayables)}</span>
            </div>
            <div className="p-3.5 flex justify-between items-center pl-8">
              <span>(5) Other Current Liabilities & Statutory Dues</span>
              <span className="font-mono font-medium">{formatINR(s3.otherCurrentLiabilities)}</span>
            </div>
            <div className="p-4 bg-gray-50 font-bold text-gray-900 flex justify-between text-sm">
              <span>TOTAL EQUITY AND LIABILITIES</span>
              <span className="font-mono text-gray-900">{formatINR(s3.totalLiabilitiesAndEquity)}</span>
            </div>

            <div className="p-4 bg-gray-50 font-bold text-gray-900 flex justify-between mt-4">
              <span>II. ASSETS</span>
              <span>Amount (₹)</span>
            </div>
            <div className="p-3.5 flex justify-between items-center pl-8">
              <span>(1) Property, Plant and Equipment (PPE & Capital Assets)</span>
              <span className="font-mono font-medium">{formatINR(s3.ppe)}</span>
            </div>
            <div className="p-3.5 flex justify-between items-center pl-8">
              <span>(2) Non-Current Investments</span>
              <span className="font-mono font-medium">{formatINR(s3.nonCurrentInvestments)}</span>
            </div>
            <div className="p-3.5 flex justify-between items-center pl-8">
              <span>(3) Inventories (Closing Stock verified / adjusted)</span>
              <span className="font-mono font-medium">{formatINR(s3.inventories)}</span>
            </div>
            <div className="p-3.5 flex justify-between items-center pl-8">
              <span>(4) Trade Receivables (Sundry Debtors)</span>
              <span className="font-mono font-medium">{formatINR(s3.tradeReceivables)}</span>
            </div>
            <div className="p-3.5 flex justify-between items-center pl-8">
              <span>(5) Cash and Bank Balances</span>
              <span className="font-mono font-medium">{formatINR(s3.cashAndBank)}</span>
            </div>
            <div className="p-3.5 flex justify-between items-center pl-8">
              <span>(6) Short-Term Loans, Advances & Other Current Assets</span>
              <span className="font-mono font-medium">{formatINR(s3.otherCurrentAssets)}</span>
            </div>
            <div className="p-4 bg-gray-50 font-bold text-gray-900 flex justify-between text-sm">
              <span>TOTAL ASSETS</span>
              <span className="font-mono text-gray-900">{formatINR(s3.totalAssets)}</span>
            </div>

            <div className={`p-4 flex justify-between items-center text-xs font-semibold ${
              currentResult.balanceSheetDiff < 1 ? 'bg-slate-100 text-gray-900' : 'bg-[#FDF2F2] text-[#C43D43]'
            }`}>
              <span>Schedule III Balance Sheet Verification (Total Assets - Total Liabilities & Equity)</span>
              <span>{currentResult.balanceSheetDiff < 1 ? '✓ Balanced (Zero Discrepancy)' : `Imbalance: ${formatINR(currentResult.balanceSheetDiff)}`}</span>
            </div>
          </div>
        )}

        {/* VIEW 2: Statement of Profit & Loss */}
        {reportType === 'pl' && (
          <div className="divide-y divide-[#E5E7EB] text-xs">
            <div className="p-4 bg-gray-50 font-bold text-gray-900 flex justify-between">
              <span>I. Revenue from Operations (Gross Operating Inflow)</span>
              <span className="font-mono font-bold text-gray-900">{formatINR(s3.revenueFromOperations)}</span>
            </div>
            <div className="p-3.5 flex justify-between items-center pl-8">
              <span>II. Other Income (Interest, Discounts, Dividends)</span>
              <span className="font-mono font-medium">{formatINR(s3.otherIncome)}</span>
            </div>
            <div className="p-4 bg-gray-50 font-bold text-gray-900 flex justify-between">
              <span>III. Total Income (I + II)</span>
              <span className="font-mono font-bold">{formatINR(s3.totalRevenue)}</span>
            </div>

            <div className="p-4 bg-gray-50 font-bold text-gray-900 flex justify-between mt-2">
              <span>IV. Expenses:</span>
              <span>Amount (₹)</span>
            </div>
            <div className="p-3.5 flex justify-between items-center pl-8">
              <span>Cost of Materials Consumed / Purchases of Stock-in-Trade</span>
              <span className="font-mono">{formatINR(s3.costOfMaterials)}</span>
            </div>
            <div className="p-3.5 flex justify-between items-center pl-8">
              <span>Employee Benefits Expense (Salaries, Wages, Bonuses)</span>
              <span className="font-mono">{formatINR(s3.employeeBenefits)}</span>
            </div>
            <div className="p-3.5 flex justify-between items-center pl-8">
              <span>Finance Costs (Bank Charges, Loan Interest)</span>
              <span className="font-mono">{formatINR(s3.financeCosts)}</span>
            </div>
            <div className="p-3.5 flex justify-between items-center pl-8">
              <span>Depreciation and Amortization Expense</span>
              <span className="font-mono">{formatINR(s3.depreciation)}</span>
            </div>
            <div className="p-3.5 flex justify-between items-center pl-8">
              <span>Other Administrative & Operational Expenses</span>
              <span className="font-mono">{formatINR(s3.otherExpenses)}</span>
            </div>
            <div className="p-4 bg-gray-50 font-bold text-gray-900 flex justify-between">
              <span>Total Expenses</span>
              <span className="font-mono font-bold text-[#C43D43]">{formatINR(s3.totalExpenses)}</span>
            </div>

            <div className="p-4 bg-slate-100 font-bold text-gray-900 flex justify-between text-sm">
              <span>V. Profit Before Exceptional Items and Tax (III - IV)</span>
              <span className="font-mono font-bold">{formatINR(s3.netProfitBeforeTax)}</span>
            </div>
          </div>
        )}

        {/* VIEW 3: Audit Adjustments Worksheet */}
        {reportType === 'adj' && (
          <div className="p-5 space-y-6">
            <div>
              <h3 className="text-sm font-bold text-gray-900">Interactive Audit Adjustments Worksheet</h3>
              <p className="text-xs text-gray-600 mt-0.5">
                Pass year-end closing entries (Depreciation, Closing Stock, Tax Provision, Accruals). Statements recalculate in real time.
              </p>
            </div>

            {/* Add Adjustment Entry Form */}
            <form onSubmit={handleAddAdjustment} className="p-4 rounded-xl bg-gray-50 border border-gray-200 space-y-3">
              <span className="text-xs font-bold text-gray-800 block">Post New Audit Adjustment Entry</span>
              <div className="grid grid-cols-1 sm:grid-cols-5 gap-3 text-xs">
                <div>
                  <label className="block text-[11px] font-medium text-gray-600 mb-1">Debit Account (Dr)</label>
                  <input
                    type="text"
                    placeholder="e.g. Closing Stock"
                    value={newDrAccount}
                    onChange={(e) => setNewDrAccount(e.target.value)}
                    className="w-full px-2.5 py-1.5 bg-white border border-gray-300 rounded text-xs"
                    required
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-medium text-gray-600 mb-1">Credit Account (Cr)</label>
                  <input
                    type="text"
                    placeholder="e.g. Changes in Inventory"
                    value={newCrAccount}
                    onChange={(e) => setNewCrAccount(e.target.value)}
                    className="w-full px-2.5 py-1.5 bg-white border border-gray-300 rounded text-xs"
                    required
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-medium text-gray-600 mb-1">Amount (₹)</label>
                  <input
                    type="number"
                    step="0.01"
                    placeholder="e.g. 50000"
                    value={newAmount}
                    onChange={(e) => setNewAmount(e.target.value)}
                    className="w-full px-2.5 py-1.5 bg-white border border-gray-300 rounded text-xs font-mono"
                    required
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-medium text-gray-600 mb-1">Category</label>
                  <select
                    value={newCategory}
                    onChange={(e) => setNewCategory(e.target.value as any)}
                    className="w-full px-2.5 py-1.5 bg-white border border-gray-300 rounded text-xs"
                  >
                    <option value="Closing Stock">Closing Stock</option>
                    <option value="Depreciation">Depreciation</option>
                    <option value="Provision">Provision</option>
                    <option value="Accrual">Accrual</option>
                    <option value="Prepaid">Prepaid</option>
                    <option value="Reclassification">Reclassification</option>
                    <option value="Other">Other</option>
                  </select>
                </div>
                <div className="flex items-end">
                  <button
                    type="submit"
                    className="w-full py-1.5 px-3 bg-gray-900 hover:bg-black text-white rounded text-xs font-semibold flex items-center justify-center gap-1 shadow-xs"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    <span>Apply Adjustment</span>
                  </button>
                </div>
              </div>
              <div>
                <label className="block text-[11px] font-medium text-gray-600 mb-1">Audit Rationale / Narration</label>
                <input
                  type="text"
                  placeholder="e.g. Physical inventory valuation verified as at 31st March 2027"
                  value={newNarration}
                  onChange={(e) => setNewNarration(e.target.value)}
                  className="w-full px-2.5 py-1 bg-white border border-gray-300 rounded text-xs"
                />
              </div>
            </form>

            {/* Adjustments Table */}
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs border-collapse">
                <thead className="bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold">
                  <tr>
                    <th className="py-2.5 px-3 border-r border-gray-200">Category</th>
                    <th className="py-2.5 px-3 border-r border-gray-200">Debit Account (Dr)</th>
                    <th className="py-2.5 px-3 border-r border-gray-200">Credit Account (Cr)</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 text-right">Amount (₹)</th>
                    <th className="py-2.5 px-3 border-r border-gray-200">Audit Rationale</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 text-center">Status</th>
                    <th className="py-2.5 px-3 text-center">Delete</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#E5E7EB] text-gray-900">
                  {adjustments.map((a) => (
                    <tr key={a.id} className={a.isApproved ? 'bg-gray-50/20' : 'bg-gray-50/50'}>
                      <td className="py-2 px-3 border-r border-gray-200 font-semibold">{a.category}</td>
                      <td className="py-2 px-3 border-r border-gray-200 font-mono text-gray-900">{a.debitAccountName}</td>
                      <td className="py-2 px-3 border-r border-gray-200 font-mono text-gray-800">{a.creditAccountName}</td>
                      <td className="py-2 px-3 border-r border-gray-200 text-right font-mono font-bold">{formatINR(a.amount)}</td>
                      <td className="py-2 px-3 border-r border-gray-200 text-[11px] text-gray-600 max-w-[200px] truncate">{a.narration}</td>
                      <td className="py-2 px-3 border-r border-gray-200 text-center">
                        <button
                          onClick={() => handleToggleAdjustment(a.id)}
                          className={`px-2 py-0.5 rounded text-[10px] font-semibold transition-colors ${
                            a.isApproved ? 'bg-slate-100 text-gray-900 border border-slate-200' : 'bg-gray-100 text-gray-600'
                          }`}
                        >
                          {a.isApproved ? 'Applied ✓' : 'Excluded'}
                        </button>
                      </td>
                      <td className="py-2 px-3 text-center">
                        <button
                          onClick={() => handleDeleteAdjustment(a.id)}
                          className="p-1 text-gray-400 hover:text-red-600 transition-colors"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* VIEW 4: Adjusted Trial Balance */}
        {reportType === 'tb' && (
          <div className="overflow-x-auto max-h-[500px]">
            <table className="w-full text-left text-xs border-collapse">
              <thead className="sticky top-0 bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold z-10">
                <tr>
                  <th className="py-2.5 px-3 border-r border-gray-200">Code</th>
                  <th className="py-2.5 px-3 border-r border-gray-200">Account Title</th>
                  <th className="py-2.5 px-3 border-r border-gray-200">Classified Head</th>
                  <th className="py-2.5 px-3 border-r border-gray-200">Schedule III Group</th>
                  <th className="py-2.5 px-3 border-r border-gray-200 text-right">Debit (₹)</th>
                  <th className="py-2.5 px-3 border-r border-gray-200 text-right">Credit (₹)</th>
                  <th className="py-2.5 px-3 text-right">Net Balance (₹)</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#E5E7EB] text-gray-900">
                {currentResult.trialBalance.map((item, idx) => (
                  <tr key={idx} className={`hover:bg-gray-50 ${item.isAbnormal ? 'bg-amber-50/50' : ''}`}>
                    <td className="py-2 px-3 border-r border-gray-200 font-mono text-[11px] text-gray-600">
                      {item.accountCode}
                    </td>
                    <td className="py-2 px-3 border-r border-gray-200 font-medium text-gray-900">
                      {item.accountName}
                    </td>
                    <td className="py-2 px-3 border-r border-gray-200">
                      <span className={`px-2 py-0.5 rounded text-[10px] font-semibold ${
                        item.head === 'Asset' ? 'bg-gray-100 text-gray-700' :
                        item.head === 'Liability' ? 'bg-[#FEF9EE] text-[#B45309]' :
                        item.head === 'Equity' ? 'bg-[#EFF6FF] text-[#1D4ED8]' :
                        item.head === 'Revenue' ? 'bg-gray-50 text-[#111827]' :
                        'bg-[#FDF2F2] text-[#C43D43]'
                      }`}>
                        {item.head}
                      </span>
                    </td>
                    <td className="py-2 px-3 border-r border-gray-200 text-[11px] text-gray-600">
                      {item.scheduleIIIGroup}
                    </td>
                    <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">
                      {item.debit > 0 ? formatINR(item.debit) : '-'}
                    </td>
                    <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">
                      {item.credit > 0 ? formatINR(item.credit) : '-'}
                    </td>
                    <td className={`py-2 px-3 text-right font-mono font-semibold ${item.netBalance >= 0 ? 'text-gray-900' : 'text-gray-900'}`}>
                      {formatINR(item.netBalance)}
                    </td>
                  </tr>
                ))}
              </tbody>
              <tfoot className="sticky bottom-0 bg-gray-50 border-t-2 border-gray-200 font-bold text-gray-900">
                <tr>
                  <td colSpan={4} className="py-2.5 px-3 border-r border-gray-200">
                    Total Trial Balance Volume
                  </td>
                  <td className="py-2.5 px-3 border-r border-gray-200 text-right font-mono text-gray-900">
                    {formatINR(currentResult.totalDebits)}
                  </td>
                  <td className="py-2.5 px-3 border-r border-gray-200 text-right font-mono text-gray-900">
                    {formatINR(currentResult.totalCredits)}
                  </td>
                  <td className="py-2.5 px-3 text-right font-mono text-gray-900">
                    {currentResult.trialBalanceDiff < 0.05 ? '0.00 (Balanced)' : `Diff: ${formatINR(currentResult.trialBalanceDiff)}`}
                  </td>
                </tr>
              </tfoot>
            </table>
          </div>
        )}

        {/* VIEW 5: Abnormal Balances */}
        {reportType === 'abnormal' && (
          <div className="p-5 space-y-4">
            <div>
              <h3 className="text-sm font-bold text-gray-900">Abnormal Account Balances Report</h3>
              <p className="text-xs text-gray-600 mt-0.5">
                Accounts with unusual signs (e.g. credit balances in cash/debtors or negative revenue/expenses).
              </p>
            </div>

            {currentResult.abnormalAccounts.length > 0 ? (
              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold">
                    <tr>
                      <th className="py-2.5 px-3 border-r border-gray-200">Code</th>
                      <th className="py-2.5 px-3 border-r border-gray-200">Account Title</th>
                      <th className="py-2.5 px-3 border-r border-gray-200">Head</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 text-right">Debit (₹)</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 text-right">Credit (₹)</th>
                      <th className="py-2.5 px-3">Discrepancy Detail & Audit Action</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#E5E7EB]">
                    {currentResult.abnormalAccounts.map((a, i) => (
                      <tr key={i} className="bg-amber-50/60">
                        <td className="py-2 px-3 border-r border-gray-200 font-mono text-[11px]">{a.accountCode}</td>
                        <td className="py-2 px-3 border-r border-gray-200 font-semibold text-gray-900">{a.accountName}</td>
                        <td className="py-2 px-3 border-r border-gray-200">{a.head}</td>
                        <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">{formatINR(a.debit)}</td>
                        <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">{formatINR(a.credit)}</td>
                        <td className="py-2 px-3 text-amber-900 font-medium">{a.abnormalReason}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <div className="p-8 text-center bg-gray-50 rounded-xl border border-gray-200 space-y-2">
                <CheckCircle2 className="w-8 h-8 text-gray-900 mx-auto" />
                <h4 className="text-sm font-bold text-gray-900">Zero Abnormal Balances Detected</h4>
                <p className="text-xs text-gray-500 max-w-md mx-auto">
                  All asset, liability, equity, revenue, and expense accounts have normal mathematical signs.
                </p>
              </div>
            )}
          </div>
        )}
      </div>
    </>
  )}
</div>
);
};
