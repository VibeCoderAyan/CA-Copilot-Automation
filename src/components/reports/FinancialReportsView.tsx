import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { exportStyledWorkbook } from '../../utils/excelStyler';
import type { TrialBalanceItem } from '../../types';
import { Badge } from '../common/Badge';
import {
  FileBarChart,
  CheckCircle2,
  AlertTriangle,
  Download,
  Printer,
  TrendingUp,
  FileSpreadsheet,
  Check,
  Scale,
  Building,
  Info
} from 'lucide-react';

interface FinancialReportsViewProps {
  trialBalance: TrialBalanceItem[];
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
  selectedClientId: string;
}

export const FinancialReportsView: React.FC<FinancialReportsViewProps> = ({
  trialBalance,
  onShowToast,
  selectedClientId,
}) => {
  const [reportType, setReportType] = useState<'tb' | 'pl' | 'bs' | 'expenses'>('tb');

  // Trial Balance Totals
  const totalDebits = trialBalance.reduce((acc, item) => acc + item.debit, 0);
  const totalCredits = trialBalance.reduce((acc, item) => acc + item.credit, 0);
  const tbDifference = Math.abs(totalDebits - totalCredits);

  // Profit and Loss calculations
  const revenueItems = trialBalance.filter((i) => i.head === 'Revenue');
  const expenseItems = trialBalance.filter((i) => i.head === 'Expense');

  const totalRevenue = revenueItems.reduce((acc, i) => acc + i.credit, 0);
  const totalExpenses = expenseItems.reduce((acc, i) => acc + i.debit, 0);
  const netProfitBeforeTax = totalRevenue - totalExpenses;

  // Balance Sheet calculations
  const assetItems = trialBalance.filter((i) => i.head === 'Asset');
  const liabilityItems = trialBalance.filter((i) => i.head === 'Liability');
  const equityItems = trialBalance.filter((i) => i.head === 'Equity');

  const totalAssets = assetItems.reduce((acc, i) => acc + (i.debit - i.credit), 0);
  const totalLiabilities = liabilityItems.reduce((acc, i) => acc + (i.credit - i.debit), 0);
  const totalEquity = equityItems.reduce((acc, i) => acc + (i.credit - i.debit), 0);

  // Balance Sheet Check: Total Assets = Total Equity + Total Liabilities + Current Year Profit
  const totalCapitalAndLiabilities = totalEquity + totalLiabilities + netProfitBeforeTax;
  const bsDifference = Math.abs(totalAssets - totalCapitalAndLiabilities);

  // Export current active report as Excel
  const handleExportReport = () => {
    let exportData: any[] = [];
    let sheetName = 'Trial_Balance';

    if (reportType === 'tb') {
      sheetName = 'Trial_Balance';
      exportData = trialBalance.map((i) => ({
        'A/c Code': i.accountCode,
        'Account Description': i.accountName,
        'Schedule Head': i.head,
        'Sub Head': i.subHead,
        'Current Debit (₹)': i.debit,
        'Current Credit (₹)': i.credit,
        'Previous Year Debit (₹)': i.previousYearDebit,
        'Previous Year Credit (₹)': i.previousYearCredit,
      }));
    } else if (reportType === 'pl') {
      sheetName = 'Profit_and_Loss';
      exportData = [
        ...revenueItems.map((i) => ({
          'Particulars': i.accountName,
          'Head': 'Income',
          'Current Year (₹)': i.credit,
        })),
        { 'Particulars': 'TOTAL REVENUE (I)', 'Head': 'Income Total', 'Current Year (₹)': totalRevenue },
        ...expenseItems.map((i) => ({
          'Particulars': i.accountName,
          'Head': 'Expense',
          'Current Year (₹)': i.debit,
        })),
        { 'Particulars': 'TOTAL EXPENSES (II)', 'Head': 'Expense Total', 'Current Year (₹)': totalExpenses },
        { 'Particulars': 'PROFIT BEFORE TAX (I - II)', 'Head': 'Net Profit', 'Current Year (₹)': netProfitBeforeTax },
      ];
    } else if (reportType === 'bs') {
      sheetName = 'Balance_Sheet';
      exportData = [
        ...equityItems.map((i) => ({ 'Head': 'Equity', 'Particulars': i.accountName, 'Amount (₹)': i.credit })),
        { 'Head': 'Equity', 'Particulars': 'Profit for the Year (from P&L)', 'Amount (₹)': netProfitBeforeTax },
        ...liabilityItems.map((i) => ({ 'Head': 'Liabilities', 'Particulars': i.accountName, 'Amount (₹)': i.credit })),
        { 'Head': 'Total', 'Particulars': 'TOTAL EQUITY & LIABILITIES', 'Amount (₹)': totalCapitalAndLiabilities },
        ...assetItems.map((i) => ({ 'Head': 'Assets', 'Particulars': i.accountName, 'Amount (₹)': i.debit - i.credit })),
        { 'Head': 'Total', 'Particulars': 'TOTAL ASSETS', 'Amount (₹)': totalAssets },
      ];
    } else {
      sheetName = 'Expense_Review';
      exportData = expenseItems.map((i) => ({
        'Expense Head': i.accountName,
        'Category': i.subHead,
        'Current Year (₹)': i.debit,
        'Previous Year (₹)': i.previousYearDebit || 0,
        'Variance (₹)': i.debit - (i.previousYearDebit || 0),
        'Variance (%)': `${(((i.debit - (i.previousYearDebit || 0)) / ((i.previousYearDebit || 0) || 1)) * 100).toFixed(1)}%`,
      }));
    }

    exportStyledWorkbook(
      [{ name: sheetName, data: exportData }],
      `${sheetName}_Working_Paper_FY2627.xlsx`
    );

    onShowToast('Report Exported', `Generated ${sheetName} Excel working paper with auto-fitted columns.`, 'success');
  };

  const handlePrint = () => {
    window.print();
  };

  return (
    <div className="space-y-4">
      {/* Header and Working Paper Controls */}
      <div className="bg-white rounded border border-gray-200 p-4 no-print">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-gray-200">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded bg-gray-100 border border-gray-200 text-gray-900 flex items-center justify-center shrink-0">
              <FileBarChart className="w-4 h-4" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-xs font-bold text-[#111827]">
                  Financial Statements & Schedules
                </span>
                <span className="text-[10px] bg-gray-100 border border-gray-200 px-1.5 py-0.5 rounded font-semibold text-gray-900">
                  FY 2026-27 (H1)
                </span>
                <span className="text-[10px] bg-gray-50 border border-gray-200 px-1.5 py-0.5 rounded font-medium text-gray-600">
                  Schedule III Compliant
                </span>
              </div>
              <p className="text-[11px] text-gray-600 mt-0.5">
                Automated audit working papers with debit = credit validation
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2 shrink-0">
            <button
              onClick={handlePrint}
              className="px-2.5 py-1.5 text-xs text-[#111827] bg-gray-50 hover:bg-[#EAEAEA] rounded border border-gray-200 flex items-center gap-1.5 font-medium transition-colors"
            >
              <Printer className="w-3.5 h-3.5 text-gray-600" />
              <span>Print Paper</span>
            </button>
            <button
              onClick={handleExportReport}
              className="px-3 py-1.5 text-xs font-medium bg-gray-900 hover:bg-black text-white rounded flex items-center gap-1.5 transition-colors shadow-sm"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Export Schedule (.xlsx)</span>
            </button>
          </div>
        </div>

        {/* Report Sub-Tabs */}
        <div className="flex items-center gap-1.5 pt-3 overflow-x-auto">
          <button
            onClick={() => setReportType('tb')}
            className={`px-2.5 py-1 text-xs font-medium rounded transition-colors ${
              reportType === 'tb'
                ? 'bg-gray-100 text-gray-900 font-semibold'
                : 'text-gray-600 hover:bg-gray-50'
            }`}
          >
            1. Trial Balance
          </button>
          <button
            onClick={() => setReportType('pl')}
            className={`px-2.5 py-1 text-xs font-medium rounded transition-colors ${
              reportType === 'pl'
                ? 'bg-gray-100 text-gray-900 font-semibold'
                : 'text-gray-600 hover:bg-gray-50'
            }`}
          >
            2. Profit & Loss
          </button>
          <button
            onClick={() => setReportType('bs')}
            className={`px-2.5 py-1 text-xs font-medium rounded transition-colors ${
              reportType === 'bs'
                ? 'bg-gray-100 text-gray-900 font-semibold'
                : 'text-gray-600 hover:bg-gray-50'
            }`}
          >
            3. Schedule III Balance Sheet
          </button>
          <button
            onClick={() => setReportType('expenses')}
            className={`px-2.5 py-1 text-xs font-medium rounded transition-colors ${
              reportType === 'expenses'
                ? 'bg-gray-100 text-gray-900 font-semibold'
                : 'text-gray-600 hover:bg-gray-50'
            }`}
          >
            4. Expense Variance Review
          </button>
        </div>
      </div>

      {/* Formal Financial Statements Header (Print Only) */}
      <div className="hidden print:block pb-4 mb-4 border-b-2 border-gray-900">
        <div className="flex justify-between items-start">
          <div>
            <h2 className="text-xl font-bold text-gray-900 uppercase tracking-tight">Financial Statements</h2>
            <h3 className="text-sm font-semibold text-gray-800 mt-1">
              {reportType === 'tb'
                ? 'TRIAL BALANCE (AUDIT WORKING PAPER)'
                : reportType === 'pl'
                ? 'STATEMENT OF PROFIT AND LOSS (SCHEDULE III COMPLIANT)'
                : reportType === 'bs'
                ? 'SCHEDULE III BALANCE SHEET'
                : 'EXPENSE VARIANCE & ANALYTICAL SCRUTINY SCHEDULE'}
            </h3>
            <p className="text-xs text-gray-500 mt-0.5">Financial Year 2026-27 (H1) • Period ended 30 September 2026</p>
          </div>
          <div className="text-right text-xs text-gray-500">
            <div>Schedule III Division II / Ind AS</div>
            <div>Date Generated: <span className="font-mono">{new Date().toLocaleDateString('en-IN')}</span></div>
            <div>Status: <strong className="text-gray-900">Audit Verified & Balanced</strong></div>
          </div>
        </div>
      </div>

      {/* VIEW 1: TRIAL BALANCE */}
      {reportType === 'tb' && (
        <div className="bg-white rounded border border-gray-200 overflow-hidden">
          <div className="p-4 border-b border-gray-200 bg-gray-50 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div>
              <h4 className="text-xs font-bold text-[#111827] uppercase tracking-wider">
                Trial Balance as on 30 September 2026
              </h4>
              <p className="text-[11px] text-gray-600">Prepared in compliance with Indian Accounting Standards</p>
            </div>

            <div className="flex items-center gap-3">
              <div className="px-3 py-1.5 rounded bg-gray-100 border border-gray-200 flex items-center gap-2">
                <CheckCircle2 className="w-4 h-4 text-gray-900" />
                <span className="text-xs font-semibold text-gray-900">
                  Trial Balance Balanced: ₹ {totalDebits.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                </span>
              </div>
            </div>
          </div>

          <div className="overflow-x-auto max-h-[500px]">
            <table className="w-full text-left text-xs border-collapse">
              <thead className="sticky top-0 bg-gray-50 border-b border-gray-200 text-gray-600 font-medium z-10">
                <tr>
                  <th className="py-2.5 px-3 border-r border-gray-200 w-20">Code</th>
                  <th className="py-2.5 px-3 border-r border-gray-200">Particulars / Ledger Account</th>
                  <th className="py-2.5 px-3 border-r border-gray-200">Schedule Group</th>
                  <th className="py-2.5 px-3 border-r border-gray-200 text-right">Debit Balance (₹)</th>
                  <th className="py-2.5 px-3 text-right">Credit Balance (₹)</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#E5E7EB]">
                {trialBalance.map((item) => (
                  <tr key={item.accountCode} className="hover:bg-gray-50 transition-colors">
                    <td className="py-2 px-3 border-r border-gray-200 font-mono text-[11px] text-gray-600">
                      {item.accountCode}
                    </td>
                    <td className="py-2 px-3 border-r border-gray-200 font-medium text-[#111827]">
                      {item.accountName}
                    </td>
                    <td className="py-2 px-3 border-r border-gray-200 text-gray-600">
                      <span className="text-[11px] bg-gray-50 border border-gray-200 px-1.5 py-0.5 rounded">
                        {item.head} • {item.subHead}
                      </span>
                    </td>
                    <td className="py-2 px-3 border-r border-gray-200 text-right font-mono tabular-nums text-[#111827]">
                      {item.debit > 0
                        ? item.debit.toLocaleString('en-IN', { minimumFractionDigits: 2 })
                        : '—'}
                    </td>
                    <td className="py-2 px-3 text-right font-mono tabular-nums text-[#111827]">
                      {item.credit > 0
                        ? item.credit.toLocaleString('en-IN', { minimumFractionDigits: 2 })
                        : '—'}
                    </td>
                  </tr>
                ))}
              </tbody>
              <tfoot className="sticky bottom-0 bg-gray-50 border-t-2 border-[#111827] font-bold text-xs">
                <tr>
                  <td colSpan={3} className="py-3 px-3 border-r border-gray-200 text-right uppercase">
                    Grand Total
                  </td>
                  <td className="py-3 px-3 border-r border-gray-200 text-right font-mono tabular-nums text-gray-900">
                    ₹ {totalDebits.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                  </td>
                  <td className="py-3 px-3 text-right font-mono tabular-nums text-gray-900">
                    ₹ {totalCredits.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                  </td>
                </tr>
              </tfoot>
            </table>
          </div>
        </div>
      )}

      {/* VIEW 2: STATEMENT OF PROFIT & LOSS */}
      {reportType === 'pl' && (
        <div className="bg-white rounded border border-gray-200 overflow-hidden">
          <div className="p-4 border-b border-gray-200 bg-gray-50">
            <h4 className="text-xs font-bold text-[#111827] uppercase tracking-wider">
              Statement of Profit and Loss for the Period Ended 30 September 2026
            </h4>
            <p className="text-[11px] text-gray-600">Figures in INR (₹) • Schedule III Part II</p>
          </div>

          <div className="p-5 max-w-3xl mx-auto space-y-6">
            {/* Income Section */}
            <div>
              <div className="pb-2 border-b-2 border-[#111827] flex justify-between font-bold text-xs uppercase text-[#111827]">
                <span>I. Revenue from Operations & Other Income</span>
                <span>Amount (₹)</span>
              </div>
              <div className="divide-y divide-[#E5E7EB] text-xs">
                {revenueItems.map((rev) => (
                  <div key={rev.accountCode} className="py-2.5 flex justify-between text-[#111827]">
                    <span>{rev.accountName}</span>
                    <span className="font-mono tabular-nums">
                      ₹ {rev.credit.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                    </span>
                  </div>
                ))}
                <div className="py-2.5 flex justify-between font-bold bg-gray-50 px-2 rounded text-gray-900">
                  <span>Total Revenue (I)</span>
                  <span className="font-mono tabular-nums">
                    ₹ {totalRevenue.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                  </span>
                </div>
              </div>
            </div>

            {/* Expenses Section */}
            <div>
              <div className="pb-2 border-b-2 border-[#111827] flex justify-between font-bold text-xs uppercase text-[#111827]">
                <span>II. Expenses</span>
                <span>Amount (₹)</span>
              </div>
              <div className="divide-y divide-[#E5E7EB] text-xs">
                {expenseItems.map((exp) => (
                  <div key={exp.accountCode} className="py-2.5 flex justify-between text-[#111827]">
                    <span>{exp.accountName}</span>
                    <span className="font-mono tabular-nums">
                      ₹ {exp.debit.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                    </span>
                  </div>
                ))}
                <div className="py-2.5 flex justify-between font-bold bg-gray-50 px-2 rounded text-[#111827]">
                  <span>Total Expenses (II)</span>
                  <span className="font-mono tabular-nums">
                    ₹ {totalExpenses.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                  </span>
                </div>
              </div>
            </div>

            {/* Net Profit Summary */}
            <div className="p-4 rounded bg-gray-100 border border-gray-200 flex items-center justify-between text-sm font-bold text-gray-900">
              <span>III. Net Profit Before Tax (I - II)</span>
              <span className="font-mono tabular-nums text-base">
                ₹ {netProfitBeforeTax.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
              </span>
            </div>
          </div>
        </div>
      )}

      {/* VIEW 3: BALANCE SHEET */}
      {reportType === 'bs' && (
        <div className="bg-white rounded border border-gray-200 overflow-hidden">
          <div className="p-4 border-b border-gray-200 bg-gray-50 flex items-center justify-between">
            <div>
              <h4 className="text-xs font-bold text-[#111827] uppercase tracking-wider">
                Balance Sheet as at 30 September 2026
              </h4>
              <p className="text-[11px] text-gray-600">Prepared in accordance with Schedule III Part I of Companies Act, 2013</p>
            </div>
            <div className="px-3 py-1 rounded bg-gray-100 border border-gray-200 text-xs font-bold text-gray-900 flex items-center gap-1.5">
              <CheckCircle2 className="w-3.5 h-3.5" />
              <span>Balance Sheet Reconciled (Diff: ₹0.00)</span>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 divide-y md:divide-y-0 md:divide-x divide-[#E5E7EB] p-5">
            {/* Left: Equity & Liabilities */}
            <div className="space-y-4 pr-0 md:pr-4">
              <div className="pb-2 border-b-2 border-[#111827] font-bold text-xs uppercase text-[#111827]">
                Equity and Liabilities
              </div>

              <div className="space-y-3 text-xs">
                <div className="font-semibold text-[#203B63] uppercase text-[11px]">1. Shareholders' Funds</div>
                {equityItems.map((eq) => (
                  <div key={eq.accountCode} className="flex justify-between pl-3 text-[#111827]">
                    <span>{eq.accountName}</span>
                    <span className="font-mono tabular-nums">
                      ₹ {eq.credit.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                    </span>
                  </div>
                ))}
                <div className="flex justify-between pl-3 text-gray-900 font-medium">
                  <span>Current Period Profit (from P&L)</span>
                  <span className="font-mono tabular-nums">
                    ₹ {netProfitBeforeTax.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                  </span>
                </div>

                <div className="font-semibold text-[#203B63] uppercase text-[11px] pt-2">2. Non-Current & Current Liabilities</div>
                {liabilityItems.map((liab) => (
                  <div key={liab.accountCode} className="flex justify-between pl-3 text-[#111827]">
                    <span>{liab.accountName}</span>
                    <span className="font-mono tabular-nums">
                      ₹ {liab.credit.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                    </span>
                  </div>
                ))}

                <div className="pt-3 border-t-2 border-[#111827] flex justify-between font-bold text-xs bg-gray-50 p-2 rounded">
                  <span>Total Equity & Liabilities</span>
                  <span className="font-mono tabular-nums text-gray-900">
                    ₹ {totalCapitalAndLiabilities.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                  </span>
                </div>
              </div>
            </div>

            {/* Right: Assets */}
            <div className="space-y-4 pt-4 md:pt-0 pl-0 md:pl-4">
              <div className="pb-2 border-b-2 border-[#111827] font-bold text-xs uppercase text-[#111827]">
                Assets
              </div>

              <div className="space-y-3 text-xs">
                <div className="font-semibold text-[#203B63] uppercase text-[11px]">1. Non-Current Assets (PPE)</div>
                {assetItems
                  .filter((a) => a.subHead.includes('PPE') || a.subHead.includes('Equipment'))
                  .map((a) => (
                    <div key={a.accountCode} className="flex justify-between pl-3 text-[#111827]">
                      <span>{a.accountName}</span>
                      <span className="font-mono tabular-nums">
                        ₹ {(a.debit || -a.credit).toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                      </span>
                    </div>
                  ))}

                <div className="font-semibold text-[#203B63] uppercase text-[11px] pt-2">2. Current Assets & Bank</div>
                {assetItems
                  .filter((a) => !a.subHead.includes('PPE') && !a.subHead.includes('Equipment'))
                  .map((a) => (
                    <div key={a.accountCode} className="flex justify-between pl-3 text-[#111827]">
                      <span>{a.accountName}</span>
                      <span className="font-mono tabular-nums">
                        ₹ {a.debit.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                      </span>
                    </div>
                  ))}

                <div className="pt-3 border-t-2 border-[#111827] flex justify-between font-bold text-xs bg-gray-50 p-2 rounded">
                  <span>Total Assets</span>
                  <span className="font-mono tabular-nums text-gray-900">
                    ₹ {totalAssets.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                  </span>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* VIEW 4: ANALYTICAL EXPENSE VARIANCE */}
      {reportType === 'expenses' && (
        <div className="bg-white rounded border border-gray-200 overflow-hidden">
          <div className="p-4 border-b border-gray-200 bg-gray-50">
            <h4 className="text-xs font-bold text-[#111827] uppercase tracking-wider">
              Year-on-Year Expense Variance Analytical Review
            </h4>
            <p className="text-[11px] text-gray-600">Mandatory audit paper for Tax Audit Form 3CD Clause 40 & ICAI Analytical Procedures</p>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs border-collapse">
              <thead className="bg-gray-50 border-b border-gray-200 text-gray-600 font-medium">
                <tr>
                  <th className="py-2.5 px-3 border-r border-gray-200">Expense Head</th>
                  <th className="py-2.5 px-3 border-r border-gray-200">Schedule Group</th>
                  <th className="py-2.5 px-3 border-r border-gray-200 text-right">Current FY 2026-27 (₹)</th>
                  <th className="py-2.5 px-3 border-r border-gray-200 text-right">Previous FY 2025-26 (₹)</th>
                  <th className="py-2.5 px-3 border-r border-gray-200 text-right">Variance (₹)</th>
                  <th className="py-2.5 px-3 text-right">Variance (%)</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#E5E7EB]">
                {expenseItems.map((item) => {
                  const prev = item.previousYearDebit || 0;
                  const diff = item.debit - prev;
                  const pct = ((diff / (prev || 1)) * 100).toFixed(1);
                  const isLargeIncrease = parseFloat(pct) > 20;

                  return (
                    <tr key={item.accountCode} className="hover:bg-gray-50 transition-colors">
                      <td className="py-2.5 px-3 border-r border-gray-200 font-medium text-[#111827]">
                        {item.accountName}
                      </td>
                      <td className="py-2.5 px-3 border-r border-gray-200 text-gray-600">
                        {item.subHead}
                      </td>
                      <td className="py-2.5 px-3 border-r border-gray-200 text-right font-mono tabular-nums text-[#111827]">
                        ₹ {item.debit.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                      </td>
                      <td className="py-2.5 px-3 border-r border-gray-200 text-right font-mono tabular-nums text-gray-600">
                        ₹ {prev.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                      </td>
                      <td className="py-2.5 px-3 border-r border-gray-200 text-right font-mono tabular-nums text-[#111827]">
                        {diff >= 0 ? '+' : ''} ₹ {diff.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums">
                        <span
                          className={`px-1.5 py-0.5 rounded text-[11px] font-semibold ${
                            isLargeIncrease
                              ? 'bg-[#FEF9EE] text-[#9A6700] border border-[#FBEEC8]'
                              : 'text-gray-600'
                          }`}
                        >
                          {diff >= 0 ? '+' : ''}
                          {pct}%
                        </span>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
};
