import { ToolHeader } from '../common/ToolHeader';
import React, { useState, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  UploadCloud,
  FileSpreadsheet,
  AlertTriangle,
  AlertCircle,
  CheckCircle2,
  Check,
  X,
  Search,
  Sliders,
  ShieldCheck,
  Eye,
  RefreshCw,
  Building,
  DollarSign,
  TrendingUp,
  TrendingDown,
  Printer,
  Calendar,
  Layers,
} from 'lucide-react';
import { InnerToolBadge } from '../common/InnerToolBadge';
import {
  type CashFlowLineItem,
  type CashFlowStatementResult,
  type CashFlowColumnMapping,
  detectCashFlowColumns,
  generateCashFlowStatement,
  exportCashFlowWorkbook,
} from '../../utils/cashFlowEngine';

interface CashFlowToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

export const CashFlowTool: React.FC<CashFlowToolProps> = ({
  onBack,
  onShowToast,
}) => {
  const [rawRows, setRawRows] = useState<Record<string, any>[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [mapping, setMapping] = useState<CashFlowColumnMapping>({
    accountNameCol: '',
    categoryCol: '',
    cyAmountCol: '',
    pyAmountCol: '',
  });
  const [fileName, setFileName] = useState('');
  const [clientName, setClientName] = useState('');
  const [cyPeriod, setCyPeriod] = useState('FY 2026-27');
  const [pyPeriod, setPyPeriod] = useState('FY 2025-26');
  const [activeTab, setActiveTab] = useState<'statement' | 'reconciliation' | 'unmapped' | 'printable_statement'>('statement');
  const [showMappingPanel, setShowMappingPanel] = useState(false);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Generate Cash Flow Statement
  const statementResult: CashFlowStatementResult = useMemo(() => {
    if (rawRows.length === 0) {
      return {
        clientName,
        cyPeriod,
        pyPeriod,
        operatingItems: [],
        netOperatingCash: 0,
        investingItems: [],
        netInvestingCash: 0,
        financingItems: [],
        netFinancingCash: 0,
        netChangeInCash: 0,
        openingCash: 0,
        computedClosingCash: 0,
        actualClosingCash: 0,
        reconciliationDifference: 0,
        isReconciled: true,
        unmappedAccounts: [],
      };
    }

    return generateCashFlowStatement(rawRows, mapping, clientName, cyPeriod, pyPeriod);
  }, [rawRows, mapping, clientName, cyPeriod, pyPeriod]);

  // Handle file upload
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json<Record<string, any>>(ws, { defval: '' });

        if (json.length === 0) {
          onShowToast('Empty File', 'The uploaded file has no data rows.', 'error');
          return;
        }

        const detectedHeaders = Object.keys(json[0]);
        setHeaders(detectedHeaders);
        setRawRows(json);

        const auto = detectCashFlowColumns(detectedHeaders);
        setMapping(auto);
      } catch (err: any) {
        onShowToast('Upload Error', err.message || 'Failed to read financial statement', 'error');
      }
    };

    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  // Export Excel workbook
  const handleExportExcel = () => {
    if (statementResult.operatingItems.length === 0) {
      onShowToast('No Data', 'Please load financial data before exporting.', 'error');
      return;
    }

    try {
      const wb = exportCashFlowWorkbook(statementResult);
      XLSX.writeFile(wb, `Cash_Flow_Statement_${clientName.replace(/\s+/g, '_')}_${cyPeriod}.xlsx`);
      onShowToast('Excel Workbook Exported', 'Exported AS-3 compliant Cash Flow Statement and Reconciliation Schedule.', 'success');
    } catch (err: any) {
      onShowToast('Export Error', err.message || 'Failed to generate Excel file.', 'error');
    }
  };

  // Print PDF
  const handlePrint = () => {
    if (statementResult.operatingItems.length === 0) {
      onShowToast('No Data', 'Please load financial data before printing.', 'error');
      return;
    }
    setActiveTab('printable_statement');
    setTimeout(() => {
      window.print();
    }, 300);
  };

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="cashflow-generator"
        title="Cash Flow Statement Generator"
        description="Automated preparation of Operating, Investing, and Financing cash flows with balance sheet cash tie-out validation."
        badge="AS-3 Compliant"
        onBack={onBack}
        hidePrintHeader={activeTab === 'printable_statement'}
        actionButtons={
          <div className="flex flex-wrap items-center gap-2">
            

            <button
              onClick={() => fileInputRef.current?.click()}
              className="px-3 py-1.5 text-xs font-semibold rounded-md bg-white text-gray-900 border border-gray-200 hover:bg-gray-50 transition-colors flex items-center gap-1.5 shadow-sm"
            >
              <UploadCloud className="w-3.5 h-3.5 text-gray-900" />
              {fileName ? 'Re-upload Statements' : 'Import Statements'}
            </button>
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleFileUpload}
              accept=".xlsx,.xls,.csv"
              className="hidden"
              id="cashflow-file-upload"
            />

            <button
              onClick={() => setShowMappingPanel(!showMappingPanel)}
              className={`px-3 py-1.5 text-xs font-semibold rounded-md border transition-colors flex items-center gap-1.5 ${
                showMappingPanel
                  ? 'bg-gray-900 text-white border-gray-300'
                  : 'bg-white text-gray-900 border-gray-200 hover:bg-gray-50'
              }`}
            >
              <Sliders className="w-3.5 h-3.5" />
              Map Accounts
            </button>

            <button
              onClick={handlePrint}
              disabled={statementResult.operatingItems.length === 0}
              className="px-3 py-1.5 text-xs font-semibold rounded-md bg-white text-gray-900 border border-gray-200 hover:bg-gray-50 disabled:opacity-50 transition-colors flex items-center gap-1.5"
            >
              <Printer className="w-3.5 h-3.5 text-gray-600" />
              PDF Statement
            </button>

            <button
              onClick={handleExportExcel}
              disabled={statementResult.operatingItems.length === 0}
              className="px-3.5 py-1.5 text-xs font-semibold rounded-md bg-gray-900 text-white hover:bg-black disabled:opacity-50 transition-colors flex items-center gap-1.5 shadow-sm cursor-pointer"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              Export Excel (.xlsx)
            </button>
          </div>
        }
      />

      {/* Client & Period Header */}
      <div className={`bg-white border border-gray-200 rounded-xl p-4 shadow-2xs flex flex-wrap items-center gap-4 text-xs ${activeTab === 'printable_statement' ? 'print:hidden' : ''}`}>
          <div className="flex items-center gap-2">
            <Building className="w-4 h-4 text-gray-600" />
            <span className="font-semibold text-gray-600">Entity:</span>
            <input
              type="text"
              value={clientName}
              onChange={(e) => setClientName(e.target.value)}
              className="px-2 py-1 border border-gray-200 rounded text-xs font-medium focus:outline-none focus:border-gray-300 w-56"
            />
          </div>

          <div className="flex items-center gap-2">
            <Calendar className="w-4 h-4 text-gray-600" />
            <span className="font-semibold text-gray-600">Current Period:</span>
            <input
              type="text"
              value={cyPeriod}
              onChange={(e) => setCyPeriod(e.target.value)}
              className="px-2 py-1 border border-gray-200 rounded text-xs font-medium focus:outline-none focus:border-gray-300 w-28"
            />
          </div>

          <div className="flex items-center gap-2">
            <span className="font-semibold text-gray-600">Prior Period:</span>
            <input
              type="text"
              value={pyPeriod}
              onChange={(e) => setPyPeriod(e.target.value)}
              className="px-2 py-1 border border-gray-200 rounded text-xs font-medium focus:outline-none focus:border-gray-300 w-28"
            />
          </div>

          {fileName && (
            <div className="ml-auto text-xs text-gray-600">
              Source: <span className="font-semibold text-gray-900">{fileName}</span> ({rawRows.length} accounts)
            </div>
          )}
      </div>

      {/* Upload Dropzone when no file selected */}
      {!fileName && (
        <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-8 sm:p-10 text-center transition-colors shadow-2xs">
          <label htmlFor="cashflow-file-upload" className="cursor-pointer flex flex-col items-center">
            <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mx-auto mb-3 shadow-2xs">
              <UploadCloud className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-semibold text-gray-900">Upload Financial Statements or Trial Balance</h3>
            <p className="text-xs text-gray-500 mt-1 max-w-md mx-auto leading-relaxed">
              Upload client Balance Sheet and P&L statements (.xlsx, .xls, .csv). Automatically classifies Operating, Investing, and Financing cash flows with balance sheet cash tie-out.
            </p>
            <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
              <UploadCloud className="w-4 h-4" />
              <span>Import Financial Statements</span>
            </div>
            <span className="text-[11px] text-gray-400 font-medium mt-3 block">
              Supports .xlsx, .xls, .csv · AS-3 and Ind AS-7 compliant cash flow statement generation
            </span>
          </label>
        </div>
      )}

      {/* Main Container */}
      <div className="space-y-6">
        {/* Mapping Drawer */}
        {showMappingPanel && (
          <div className="bg-white border border-gray-200 rounded-xl p-5 shadow-sm space-y-4">
            <div className="flex items-center justify-between pb-3 border-b border-gray-200">
              <div>
                <h3 className="text-sm font-bold text-gray-900">Column Mapping: Balance Sheet & P&L Statement</h3>
                <p className="text-xs text-gray-600">Map financial line items and period figures</p>
              </div>
              <button onClick={() => setShowMappingPanel(false)} className="text-xs text-gray-600 hover:text-gray-900">Close</button>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-4 gap-4 text-xs">
              <div>
                <label className="block text-gray-600 mb-1 font-semibold">Account / Particulars *</label>
                <select
                  value={mapping.accountNameCol}
                  onChange={(e) => setMapping({ ...mapping, accountNameCol: e.target.value })}
                  className="w-full border border-gray-200 rounded p-2 outline-none"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-gray-600 mb-1 font-semibold">Category / Schedule Head</label>
                <select
                  value={mapping.categoryCol}
                  onChange={(e) => setMapping({ ...mapping, categoryCol: e.target.value })}
                  className="w-full border border-gray-200 rounded p-2 outline-none"
                >
                  <option value="">-- Select (Optional) --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-gray-600 mb-1 font-semibold">Current Year Amount (CY) *</label>
                <select
                  value={mapping.cyAmountCol}
                  onChange={(e) => setMapping({ ...mapping, cyAmountCol: e.target.value })}
                  className="w-full border border-gray-200 rounded p-2 outline-none"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-gray-600 mb-1 font-semibold">Prior Year Amount (PY) *</label>
                <select
                  value={mapping.pyAmountCol}
                  onChange={(e) => setMapping({ ...mapping, pyAmountCol: e.target.value })}
                  className="w-full border border-gray-200 rounded p-2 outline-none"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
            </div>
          </div>
        )}

        {/* Empty State */}
        {rawRows.length === 0 ? (
          <div className="bg-white border border-dashed border-gray-200 rounded-2xl p-12 text-center shadow-sm">
            <div className="w-16 h-16 bg-gray-50 rounded-full flex items-center justify-center mx-auto mb-4 border border-gray-300/20">
              <Layers className="w-8 h-8 text-gray-900" />
            </div>
            <h3 className="text-base font-bold text-gray-900">No Balance Sheet or P&L Loaded</h3>
            <p className="text-xs text-gray-600 max-w-md mx-auto mt-1 mb-6">
              Upload your client's comparative Balance Sheet and Statement of Profit & Loss to automatically generate an AS-3 / Ind AS 7 compliant Cash Flow Statement.
            </p>
            <div className="flex items-center justify-center gap-3">
              
              <button
                onClick={() => fileInputRef.current?.click()}
                className="px-4 py-2 text-xs font-semibold rounded-lg bg-gray-900 text-white hover:bg-black transition-colors shadow-sm inline-flex items-center gap-2"
              >
                <UploadCloud className="w-4 h-4" />
                <span>Upload Financial Statements</span>
              </button>
            </div>
          </div>
        ) : (
          <div className="space-y-6">
            {/* KPI Summary Cards */}
            <div className={`grid grid-cols-2 md:grid-cols-5 gap-4 ${activeTab === 'printable_statement' ? 'print:hidden' : ''}`}>
              <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
                <span className="text-xs font-semibold text-gray-600 uppercase">A. Operating Cash</span>
                <div className={`text-lg font-bold mt-1 ${statementResult.netOperatingCash >= 0 ? 'text-gray-900' : 'text-[#DC2626]'}`}>
                  ₹{statementResult.netOperatingCash.toLocaleString('en-IN')}
                </div>
                <div className="text-xs text-gray-600 mt-0.5">Core business operations</div>
              </div>

              <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
                <span className="text-xs font-semibold text-gray-600 uppercase">B. Investing Cash</span>
                <div className={`text-lg font-bold mt-1 ${statementResult.netInvestingCash >= 0 ? 'text-gray-900' : 'text-[#DC2626]'}`}>
                  ₹{statementResult.netInvestingCash.toLocaleString('en-IN')}
                </div>
                <div className="text-xs text-gray-600 mt-0.5">Capex and fixed assets</div>
              </div>

              <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
                <span className="text-xs font-semibold text-gray-600 uppercase">C. Financing Cash</span>
                <div className={`text-lg font-bold mt-1 ${statementResult.netFinancingCash >= 0 ? 'text-gray-900' : 'text-[#DC2626]'}`}>
                  ₹{statementResult.netFinancingCash.toLocaleString('en-IN')}
                </div>
                <div className="text-xs text-gray-600 mt-0.5">Debt, equity & dividends</div>
              </div>

              <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
                <span className="text-xs font-semibold text-gray-600 uppercase">Net Change (A+B+C)</span>
                <div className={`text-lg font-bold mt-1 ${statementResult.netChangeInCash >= 0 ? 'text-gray-900' : 'text-[#DC2626]'}`}>
                  ₹{statementResult.netChangeInCash.toLocaleString('en-IN')}
                </div>
                <div className="text-xs text-gray-600 mt-0.5">Total cash movement</div>
              </div>

              <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
                <span className="text-xs font-semibold text-gray-600 uppercase">Balance Sheet Tie-Out</span>
                <div className="flex items-center gap-1.5 mt-1">
                  {statementResult.isReconciled ? (
                    <span className="px-2 py-0.5 text-xs font-bold rounded bg-[#F3F4F6] text-gray-900">
                      ✓ Balanced (₹0.00)
                    </span>
                  ) : (
                    <span className="px-2 py-0.5 text-xs font-bold rounded bg-[#FEE2E2] text-[#DC2626]">
                      Diff: ₹{statementResult.reconciliationDifference.toLocaleString('en-IN')}
                    </span>
                  )}
                </div>
                <div className="text-xs text-gray-600 mt-0.5">
                  Closing Cash: ₹{statementResult.actualClosingCash.toLocaleString('en-IN')}
                </div>
              </div>
            </div>

            {/* Tabs Bar */}
            <div className="flex items-center justify-between border-b border-gray-200 pb-1 print:hidden">
              <div className="flex items-center gap-2">
                <button
                  onClick={() => setActiveTab('statement')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'statement'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  AS-3 Cash Flow Statement
                </button>
                <button
                  onClick={() => setActiveTab('reconciliation')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'reconciliation'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Reconciliation Tie-Out Schedule
                </button>
                <button
                  onClick={() => setActiveTab('unmapped')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'unmapped'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Unmapped Line Items ({statementResult.unmappedAccounts.length})
                </button>
                <button
                  onClick={() => setActiveTab('printable_statement')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'printable_statement'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Formal Statement Preview
                </button>
              </div>
            </div>

            {/* TAB 1: AS-3 Statement Table */}
            {activeTab === 'statement' && (
              <div className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm">
                <div className="p-4 bg-gray-900 text-white flex justify-between items-center">
                  <div>
                    <h3 className="font-bold text-sm">Statement of Cash Flows (Indirect Method)</h3>
                    <p className="text-xs text-gray-100">Prepared in compliance with Accounting Standard 3 (AS-3)</p>
                  </div>
                  <div className="text-xs text-right font-semibold">
                    Figures in INR (₹)
                  </div>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-xs text-left">
                    <thead className="bg-gray-100 font-semibold text-gray-900">
                      <tr>
                        <th className="p-3 w-3/5">Particulars / Cash Flow Items</th>
                        <th className="p-3 w-1/5">Underlying Accounts / Schedule</th>
                        <th className="p-3 w-1/5 text-right">{cyPeriod} (₹)</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB]">
                      {/* Section A: Operating */}
                      <tr className="bg-gray-50 font-bold text-gray-900">
                        <td colSpan={3} className="p-3 uppercase">
                          A. Cash Flow from Operating Activities
                        </td>
                      </tr>
                      {statementResult.operatingItems.map((item) => (
                        <tr
                          key={item.id}
                          className={`${
                            item.isTotal
                              ? 'bg-gray-50 font-bold text-gray-900'
                              : item.isSubtotal
                              ? 'font-semibold text-gray-900'
                              : 'text-gray-600 hover:bg-gray-50'
                          }`}
                        >
                          <td className={`p-2.5 ${item.isTotal || item.isSubtotal ? '' : 'pl-6'}`}>
                            {item.description}
                          </td>
                          <td className="p-2.5 text-[11px] text-gray-600">
                            {item.notes || item.sourceAccounts}
                          </td>
                          <td className="p-2.5 text-right font-mono font-semibold">
                            {item.cyAmount < 0 ? (
                              <span className="text-[#DC2626]">({Math.abs(item.cyAmount).toLocaleString('en-IN')})</span>
                            ) : (
                              item.cyAmount.toLocaleString('en-IN')
                            )}
                          </td>
                        </tr>
                      ))}

                      {/* Section B: Investing */}
                      <tr className="bg-gray-50 font-bold text-gray-900">
                        <td colSpan={3} className="p-3 uppercase">
                          B. Cash Flow from Investing Activities
                        </td>
                      </tr>
                      {statementResult.investingItems.map((item) => (
                        <tr
                          key={item.id}
                          className={`${
                            item.isTotal
                              ? 'bg-gray-50 font-bold text-gray-900'
                              : 'text-gray-600 hover:bg-gray-50'
                          }`}
                        >
                          <td className={`p-2.5 ${item.isTotal ? '' : 'pl-6'}`}>
                            {item.description}
                          </td>
                          <td className="p-2.5 text-[11px] text-gray-600">
                            {item.notes || item.sourceAccounts}
                          </td>
                          <td className="p-2.5 text-right font-mono font-semibold">
                            {item.cyAmount < 0 ? (
                              <span className="text-[#DC2626]">({Math.abs(item.cyAmount).toLocaleString('en-IN')})</span>
                            ) : (
                              item.cyAmount.toLocaleString('en-IN')
                            )}
                          </td>
                        </tr>
                      ))}

                      {/* Section C: Financing */}
                      <tr className="bg-gray-50 font-bold text-gray-900">
                        <td colSpan={3} className="p-3 uppercase">
                          C. Cash Flow from Financing Activities
                        </td>
                      </tr>
                      {statementResult.financingItems.map((item) => (
                        <tr
                          key={item.id}
                          className={`${
                            item.isTotal
                              ? 'bg-gray-50 font-bold text-gray-900'
                              : 'text-gray-600 hover:bg-gray-50'
                          }`}
                        >
                          <td className={`p-2.5 ${item.isTotal ? '' : 'pl-6'}`}>
                            {item.description}
                          </td>
                          <td className="p-2.5 text-[11px] text-gray-600">
                            {item.notes || item.sourceAccounts}
                          </td>
                          <td className="p-2.5 text-right font-mono font-semibold">
                            {item.cyAmount < 0 ? (
                              <span className="text-[#DC2626]">({Math.abs(item.cyAmount).toLocaleString('en-IN')})</span>
                            ) : (
                              item.cyAmount.toLocaleString('en-IN')
                            )}
                          </td>
                        </tr>
                      ))}

                      {/* Net Movement & Reconciliation Rows */}
                      <tr className="bg-gray-100 font-bold text-gray-900 border-t-2 border-gray-300">
                        <td className="p-3">NET INCREASE / (DECREASE) IN CASH & EQUIVALENTS (A + B + C)</td>
                        <td className="p-3 text-[11px]">Net Periodic Movement</td>
                        <td className="p-3 text-right font-mono font-bold text-sm">
                          {statementResult.netChangeInCash < 0 ? (
                            <span className="text-[#DC2626]">({Math.abs(statementResult.netChangeInCash).toLocaleString('en-IN')})</span>
                          ) : (
                            statementResult.netChangeInCash.toLocaleString('en-IN')
                          )}
                        </td>
                      </tr>
                      <tr>
                        <td className="p-2.5 pl-6 text-gray-600">Cash and Cash Equivalents at Beginning of Year</td>
                        <td className="p-2.5 text-[11px] text-gray-600">Opening Balance Sheet</td>
                        <td className="p-2.5 text-right font-mono font-semibold">
                          ₹{statementResult.openingCash.toLocaleString('en-IN')}
                        </td>
                      </tr>
                      <tr className="bg-gray-50 font-bold text-gray-900">
                        <td className="p-3">Cash and Cash Equivalents at End of Year (Closing)</td>
                        <td className="p-3 text-[11px]">Closing Balance Sheet</td>
                        <td className="p-3 text-right font-mono font-bold text-sm">
                          ₹{statementResult.computedClosingCash.toLocaleString('en-IN')}
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* TAB 2: Reconciliation Tie-Out Schedule */}
            {activeTab === 'reconciliation' && (
              <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-sm space-y-6">
                <div>
                  <h3 className="text-sm font-bold text-gray-900">Balance Sheet Cash Tie-Out Verification</h3>
                  <p className="text-xs text-gray-600">
                    Confirms that computed closing cash from the cash flow statement equals the actual ending cash on the Balance Sheet
                  </p>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-6 text-xs">
                  <div className="border border-gray-200 rounded-lg p-4 space-y-3">
                    <h4 className="font-bold text-gray-900 uppercase border-b border-gray-200 pb-2">
                      Statement Derivations
                    </h4>
                    <div className="flex justify-between py-1 border-b border-gray-100">
                      <span className="text-gray-600">Opening Cash (PY Balance Sheet):</span>
                      <span className="font-mono font-semibold">₹{statementResult.openingCash.toLocaleString('en-IN')}</span>
                    </div>
                    <div className="flex justify-between py-1 border-b border-gray-100">
                      <span className="text-gray-600">Add: Net Cash from Operations (A):</span>
                      <span className="font-mono font-semibold">₹{statementResult.netOperatingCash.toLocaleString('en-IN')}</span>
                    </div>
                    <div className="flex justify-between py-1 border-b border-gray-100">
                      <span className="text-gray-600">Add: Net Cash from Investing (B):</span>
                      <span className="font-mono font-semibold">₹{statementResult.netInvestingCash.toLocaleString('en-IN')}</span>
                    </div>
                    <div className="flex justify-between py-1 border-b border-gray-100">
                      <span className="text-gray-600">Add: Net Cash from Financing (C):</span>
                      <span className="font-mono font-semibold">₹{statementResult.netFinancingCash.toLocaleString('en-IN')}</span>
                    </div>
                    <div className="flex justify-between py-1 font-bold text-sm text-gray-900">
                      <span>Computed Ending Cash Balance:</span>
                      <span className="font-mono">₹{statementResult.computedClosingCash.toLocaleString('en-IN')}</span>
                    </div>
                  </div>

                  <div className="border border-gray-200 rounded-lg p-4 space-y-3">
                    <h4 className="font-bold text-gray-900 uppercase border-b border-gray-200 pb-2">
                      Balance Sheet Verification
                    </h4>
                    <div className="flex justify-between py-1 border-b border-gray-100">
                      <span className="text-gray-600">Actual Cash on Balance Sheet:</span>
                      <span className="font-mono font-semibold">₹{statementResult.actualClosingCash.toLocaleString('en-IN')}</span>
                    </div>
                    <div className="flex justify-between py-1 border-b border-gray-100">
                      <span className="text-gray-600">Variance / Tie-Out Difference:</span>
                      <span className={`font-mono font-bold ${statementResult.reconciliationDifference === 0 ? 'text-gray-900' : 'text-[#DC2626]'}`}>
                        ₹{statementResult.reconciliationDifference.toLocaleString('en-IN')}
                      </span>
                    </div>
                    <div className="p-3 rounded bg-gray-50 border border-gray-300/30 text-gray-900 font-medium mt-4">
                      {statementResult.isReconciled ? (
                        <div>
                          <strong>✓ 100% Balanced:</strong> The statement reconciles exactly with the closing cash & bank balances of the Balance Sheet with zero discrepancy.
                        </div>
                      ) : (
                        <div className="text-[#DC2626]">
                          <strong>Variance Detected:</strong> An unmapped balance-sheet movement or non-cash item exists that requires CA adjustment.
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              </div>
            )}

            {/* TAB 3: Unmapped Accounts */}
            {activeTab === 'unmapped' && (
              <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-sm space-y-4">
                <div>
                  <h3 className="text-sm font-bold text-gray-900">Unmapped or Flagged Line Items</h3>
                  <p className="text-xs text-gray-600">
                    Accounts from the uploaded file that were not classified into operating, investing, or financing heads
                  </p>
                </div>

                {statementResult.unmappedAccounts.length === 0 ? (
                  <div className="p-6 text-center text-xs text-gray-900 bg-gray-50 rounded-lg border border-gray-300/30">
                    ✓ All accounts from the uploaded financial statement were mapped and classified successfully.
                  </div>
                ) : (
                  <ul className="divide-y divide-[#E5E7EB] text-xs">
                    {statementResult.unmappedAccounts.map((item, idx) => (
                      <li key={idx} className="py-2.5 flex justify-between items-center text-gray-600">
                        <span>• {item}</span>
                        <span className="text-[11px] bg-amber-50 text-amber-700 px-2 py-0.5 rounded">Unclassified</span>
                      </li>
                    ))}
                  </ul>
                )}
              </div>
            )}

            {/* TAB 4: Printable Formal Statement */}
            {activeTab === 'printable_statement' && (
              <div className="bg-white border border-gray-200 rounded-xl p-8 shadow-sm print:border-none print:shadow-none print:p-0">
                <div className="border-b-2 border-gray-300 pb-4 mb-6">
                  <div className="flex justify-between items-start">
                    <div>
                      <h2 className="text-xl font-bold text-gray-900">{clientName}</h2>
                      <h3 className="text-base font-semibold text-gray-900">STATEMENT OF CASH FLOWS (AS-3)</h3>
                      <p className="text-xs text-gray-600">For the period ended {cyPeriod}</p>
                    </div>
                    <div className="text-right text-xs text-gray-600">
                      <div>Amounts in Indian Rupees (₹)</div>
                      <div>Indirect Method Format</div>
                      <div>Date: {new Date().toLocaleDateString('en-IN')}</div>
                    </div>
                  </div>
                </div>

                <table className="w-full text-xs text-left mb-6 border-collapse">
                  <thead>
                    <tr className="bg-gray-900 text-white">
                      <th className="p-2 border border-gray-300 w-3/5">Particulars</th>
                      <th className="p-2 border border-gray-300 w-1/5">Schedules / Notes</th>
                      <th className="p-2 border border-gray-300 w-1/5 text-right">{cyPeriod} (₹)</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr className="font-bold bg-gray-100">
                      <td colSpan={3} className="p-2 border border-gray-300">A. CASH FLOW FROM OPERATING ACTIVITIES</td>
                    </tr>
                    {statementResult.operatingItems.map((item) => (
                      <tr key={item.id} className={item.isTotal ? 'font-bold bg-gray-50' : ''}>
                        <td className={`p-2 border border-gray-300 ${item.isTotal ? '' : 'pl-4'}`}>{item.description}</td>
                        <td className="p-2 border border-gray-300 text-[10px] text-gray-600">{item.notes || item.sourceAccounts}</td>
                        <td className="p-2 border border-gray-300 text-right font-mono">
                          {item.cyAmount < 0 ? `(${Math.abs(item.cyAmount).toLocaleString('en-IN')})` : item.cyAmount.toLocaleString('en-IN')}
                        </td>
                      </tr>
                    ))}

                    <tr className="font-bold bg-gray-100">
                      <td colSpan={3} className="p-2 border border-gray-300">B. CASH FLOW FROM INVESTING ACTIVITIES</td>
                    </tr>
                    {statementResult.investingItems.map((item) => (
                      <tr key={item.id} className={item.isTotal ? 'font-bold bg-gray-50' : ''}>
                        <td className={`p-2 border border-gray-300 ${item.isTotal ? '' : 'pl-4'}`}>{item.description}</td>
                        <td className="p-2 border border-gray-300 text-[10px] text-gray-600">{item.notes || item.sourceAccounts}</td>
                        <td className="p-2 border border-gray-300 text-right font-mono">
                          {item.cyAmount < 0 ? `(${Math.abs(item.cyAmount).toLocaleString('en-IN')})` : item.cyAmount.toLocaleString('en-IN')}
                        </td>
                      </tr>
                    ))}

                    <tr className="font-bold bg-gray-100">
                      <td colSpan={3} className="p-2 border border-gray-300">C. CASH FLOW FROM FINANCING ACTIVITIES</td>
                    </tr>
                    {statementResult.financingItems.map((item) => (
                      <tr key={item.id} className={item.isTotal ? 'font-bold bg-gray-50' : ''}>
                        <td className={`p-2 border border-gray-300 ${item.isTotal ? '' : 'pl-4'}`}>{item.description}</td>
                        <td className="p-2 border border-gray-300 text-[10px] text-gray-600">{item.notes || item.sourceAccounts}</td>
                        <td className="p-2 border border-gray-300 text-right font-mono">
                          {item.cyAmount < 0 ? `(${Math.abs(item.cyAmount).toLocaleString('en-IN')})` : item.cyAmount.toLocaleString('en-IN')}
                        </td>
                      </tr>
                    ))}

                    <tr className="font-bold bg-gray-200">
                      <td className="p-2 border border-gray-300">Net Increase / (Decrease) in Cash and Cash Equivalents</td>
                      <td className="p-2 border border-gray-300">A + B + C</td>
                      <td className="p-2 border border-gray-300 text-right font-mono font-bold">
                        {statementResult.netChangeInCash.toLocaleString('en-IN')}
                      </td>
                    </tr>
                    <tr>
                      <td className="p-2 border border-gray-300 pl-4">Cash and cash equivalents at the beginning of the year</td>
                      <td className="p-2 border border-gray-300">Opening Balance</td>
                      <td className="p-2 border border-gray-300 text-right font-mono">{statementResult.openingCash.toLocaleString('en-IN')}</td>
                    </tr>
                    <tr className="font-bold bg-gray-50">
                      <td className="p-2 border border-gray-300">Cash and cash equivalents at the end of the year</td>
                      <td className="p-2 border border-gray-300">Closing Balance</td>
                      <td className="p-2 border border-gray-300 text-right font-mono font-bold text-gray-900">
                        {statementResult.computedClosingCash.toLocaleString('en-IN')}
                      </td>
                    </tr>
                  </tbody>
                </table>

                <div className="mt-8 pt-6 border-t border-gray-200 text-xs text-gray-600 flex justify-between">
                  <div>Prepared by: Audit Assistant</div>
                  <div>For and on behalf of the Board of Directors</div>
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
};