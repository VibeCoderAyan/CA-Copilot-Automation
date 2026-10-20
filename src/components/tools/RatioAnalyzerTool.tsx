import { ToolHeader } from '../common/ToolHeader';
import React, { useState, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  UploadCloud,
  Download,
  Printer,
  FileSpreadsheet,
  AlertTriangle,
  AlertCircle,
  CheckCircle2,
  TrendingUp,
  TrendingDown,
  Minus,
  Sliders,
  ShieldCheck,
  Eye,
  RefreshCw,
  Building,
  Calculator,
  HelpCircle,
} from 'lucide-react';
import { InnerToolBadge } from '../common/InnerToolBadge';
import {
  type FinancialInputs,
  type RatioResult,
  type RatioAnalysisResult,
  type RatioColumnMapping,
  detectRatioColumns,
  parseFinancialStatements,
  computeFinancialRatios,
  exportRatioWorkbook,
} from '../../utils/ratioAnalyzerEngine';

interface RatioAnalyzerToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

export const RatioAnalyzerTool: React.FC<RatioAnalyzerToolProps> = ({
  onBack,
  onShowToast,
}) => {
  const [rawRows, setRawRows] = useState<Record<string, any>[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [mapping, setMapping] = useState<RatioColumnMapping>({
    accountNameCol: '',
    categoryCol: '',
    cyAmountCol: '',
    pyAmountCol: '',
  });
  const [fileName, setFileName] = useState('');
  const [clientName, setClientName] = useState('');
  const [cyPeriod, setCyPeriod] = useState('FY 2026-27');
  const [pyPeriod, setPyPeriod] = useState('FY 2025-26');
  const [activeTab, setActiveTab] = useState<'ratios' | 'inputs' | 'formulas' | 'printable_report'>('ratios');
  const [selectedCategory, setSelectedCategory] = useState<string>('all');
  const [showMappingPanel, setShowMappingPanel] = useState(false);
  const [showConfigPanel, setShowConfigPanel] = useState(false);

  // Benchmarks configurable by CA
  const [benchmarks, setBenchmarks] = useState({
    currentRatioMin: 1.33,
    currentRatioIdeal: 2.0,
    quickRatioMin: 1.0,
    debtEquityMax: 2.0,
    grossProfitMin: 20.0,
    netProfitMin: 5.0,
  });

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Parse financial inputs & compute ratios
  const analysisResult: RatioAnalysisResult = useMemo(() => {
    if (rawRows.length === 0) {
      // Default empty inputs
      const emptyInputs: FinancialInputs = {
        periodLabel: cyPeriod,
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
      return {
        clientName,
        cyPeriod,
        pyPeriod,
        cyInputs: emptyInputs,
        pyInputs: { ...emptyInputs, periodLabel: pyPeriod },
        ratios: [],
        alerts: [],
      };
    }

    const { cy, py, unmapped } = parseFinancialStatements(rawRows, mapping, cyPeriod, pyPeriod);
    const computed = computeFinancialRatios(cy, py, clientName);

    if (unmapped.length > 0) {
      computed.alerts.push(`${unmapped.length} line items were not mapped to standard categories. Check Statements tab.`);
    }

    return computed;
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
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json<Record<string, any>>(worksheet, { defval: '' });

        if (json.length === 0) {
          onShowToast('Empty File', 'The uploaded file has no readable data rows.', 'error');
          return;
        }

        const detectedHeaders = Object.keys(json[0]);
        setHeaders(detectedHeaders);
        setRawRows(json);

        const autoMap = detectRatioColumns(detectedHeaders);
        setMapping(autoMap);
      } catch (err: any) {
        onShowToast('Upload Error', err.message || 'Failed to parse file', 'error');
      }
    };

    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  // Export Excel workbook
  const handleExportExcel = () => {
    if (analysisResult.ratios.length === 0) {
      onShowToast('No Data', 'Please load financial data before exporting.', 'error');
      return;
    }

    try {
      const wb = exportRatioWorkbook(analysisResult);
      XLSX.writeFile(wb, `Financial_Ratio_Analysis_${clientName.replace(/\s+/g, '_')}_${cyPeriod}.xlsx`);
      onShowToast('Excel Export Ready', 'Exported 3-sheet Financial Ratio Workbook with formulas and inputs.', 'success');
    } catch (err: any) {
      onShowToast('Export Failed', err.message || 'Unable to generate Excel file.', 'error');
    }
  };

  // Print PDF
  const handlePrint = () => {
    if (analysisResult.ratios.length === 0) {
      onShowToast('No Data', 'Please load financial data before printing.', 'error');
      return;
    }
    setActiveTab('printable_report');
    setTimeout(() => {
      window.print();
    }, 300);
  };

  // Filter ratios by category
  const filteredRatios = useMemo(() => {
    if (selectedCategory === 'all') return analysisResult.ratios;
    return analysisResult.ratios.filter((r) => r.category.toLowerCase().includes(selectedCategory.toLowerCase()));
  }, [analysisResult.ratios, selectedCategory]);

  // Overall Health Summary counts
  const healthStats = useMemo(() => {
    const total = analysisResult.ratios.length;
    const healthy = analysisResult.ratios.filter((r) => r.interpretation === 'healthy').length;
    const neutral = analysisResult.ratios.filter((r) => r.interpretation === 'neutral').length;
    const concern = analysisResult.ratios.filter((r) => r.interpretation === 'concern').length;
    return { total, healthy, neutral, concern };
  }, [analysisResult.ratios]);

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="ratio-analyzer"
        title="Financial Ratio Analyzer"
        description="Analyze Balance Sheet & P&L ratios across comparative periods with formula transparency, ICAI benchmarks, and reports."
        onBack={onBack}
        actionButtons={
          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={() => fileInputRef.current?.click()}
              className="px-3 py-1.5 text-xs font-semibold rounded-md bg-white text-gray-900 border border-gray-200 hover:bg-gray-50 transition-colors flex items-center gap-1.5 shadow-sm cursor-pointer"
            >
              <UploadCloud className="w-3.5 h-3.5 text-gray-900" />
              {fileName ? 'Re-upload Statements' : 'Import Excel / CSV'}
            </button>
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleFileUpload}
              accept=".xlsx,.xls,.csv"
              className="hidden"
              id="ratio-analyzer-upload"
            />

            <button
              onClick={() => setShowMappingPanel(!showMappingPanel)}
              className={`px-3 py-1.5 text-xs font-semibold rounded-md border transition-colors flex items-center gap-1.5 cursor-pointer ${
                showMappingPanel
                  ? 'bg-gray-900 text-white border-gray-300'
                  : 'bg-white text-gray-900 border-gray-200 hover:bg-gray-50'
              }`}
            >
              <Sliders className="w-3.5 h-3.5" />
              Map Columns
            </button>

            <button
              onClick={handlePrint}
              disabled={analysisResult.ratios.length === 0}
              className="px-3 py-1.5 text-xs font-semibold rounded-md bg-white text-gray-900 border border-gray-200 hover:bg-gray-50 disabled:opacity-50 transition-colors flex items-center gap-1.5 cursor-pointer"
            >
              <Printer className="w-3.5 h-3.5 text-gray-600" />
              PDF Report
            </button>

            <button
              onClick={handleExportExcel}
              disabled={analysisResult.ratios.length === 0}
              className="px-3.5 py-1.5 text-xs font-semibold rounded-md bg-gray-900 text-white hover:bg-black disabled:opacity-50 transition-colors flex items-center gap-1.5 shadow-sm cursor-pointer"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              Export Excel (.xlsx)
            </button>
          </div>
        }
      />

      {/* Client & Period Header Bar */}
      <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-2xs flex flex-wrap items-center gap-4 text-xs">
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
              Source: <span className="font-semibold text-gray-900">{fileName}</span> ({rawRows.length} rows parsed)
            </div>
          )}
      </div>

      {/* Upload Dropzone when no file selected */}
      {!fileName && (
        <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-8 sm:p-10 text-center transition-colors shadow-2xs">
          <label htmlFor="ratio-analyzer-upload" className="cursor-pointer flex flex-col items-center">
            <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mx-auto mb-3 shadow-2xs">
              <UploadCloud className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-semibold text-gray-900">Upload Financial Statements or Trial Balance</h3>
            <p className="text-xs text-gray-500 mt-1 max-w-md mx-auto leading-relaxed">
              Upload Balance Sheet and Profit & Loss figures (.xlsx, .xls, .csv). Automatically evaluates all 11 Schedule III mandatory financial ratios with CARO 2020 variance audit thresholds.
            </p>
            <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
              <UploadCloud className="w-4 h-4" />
              <span>Import Financial Statements</span>
            </div>
            <span className="text-[11px] text-gray-400 font-medium mt-3 block">
              Supports .xlsx, .xls, .csv · Schedule III MCA compliance · Identifies ±25% variance triggers
            </span>
          </label>
        </div>
      )}

      {/* Main Body */}
      <div className="space-y-6">
        {/* Mapping Drawer */}
        {showMappingPanel && (
          <div className="bg-white border border-gray-200 rounded-xl p-5 shadow-sm">
            <div className="flex items-center justify-between pb-3 border-b border-gray-200">
              <div>
                <h3 className="text-sm font-bold text-gray-900">Column Mapping: Financial Statement Data</h3>
                <p className="text-xs text-gray-600">Map your uploaded workbook headers to the analyzer inputs</p>
              </div>
              <button
                onClick={() => setShowMappingPanel(false)}
                className="text-xs text-gray-600 hover:text-gray-900"
              >
                Close
              </button>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mt-4">
              <div>
                <label className="block text-xs font-semibold text-gray-600 mb-1">
                  Account Name / Particulars *
                </label>
                <select
                  value={mapping.accountNameCol}
                  onChange={(e) => setMapping({ ...mapping, accountNameCol: e.target.value })}
                  className="w-full text-xs border border-gray-200 rounded p-2 focus:border-gray-300 outline-none"
                >
                  <option value="">-- Select Column --</option>
                  {headers.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-xs font-semibold text-gray-600 mb-1">
                  Category / Schedule Head
                </label>
                <select
                  value={mapping.categoryCol}
                  onChange={(e) => setMapping({ ...mapping, categoryCol: e.target.value })}
                  className="w-full text-xs border border-gray-200 rounded p-2 focus:border-gray-300 outline-none"
                >
                  <option value="">-- Select Column (Optional) --</option>
                  {headers.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-xs font-semibold text-gray-600 mb-1">
                  Current Year (CY) Amount *
                </label>
                <select
                  value={mapping.cyAmountCol}
                  onChange={(e) => setMapping({ ...mapping, cyAmountCol: e.target.value })}
                  className="w-full text-xs border border-gray-200 rounded p-2 focus:border-gray-300 outline-none"
                >
                  <option value="">-- Select Column --</option>
                  {headers.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-xs font-semibold text-gray-600 mb-1">
                  Prior Year (PY) Amount
                </label>
                <select
                  value={mapping.pyAmountCol}
                  onChange={(e) => setMapping({ ...mapping, pyAmountCol: e.target.value })}
                  className="w-full text-xs border border-gray-200 rounded p-2 focus:border-gray-300 outline-none"
                >
                  <option value="">-- Select Column (Optional) --</option>
                  {headers.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>
            </div>
          </div>
        )}

        {/* Alerts / Warnings Banner */}
        {analysisResult.alerts.length > 0 && (
          <div className="bg-[#FEF2F2] border border-[#FCA5A5] rounded-xl p-4">
            <div className="flex items-start gap-2.5">
              <AlertTriangle className="w-5 h-5 text-[#DC2626] shrink-0 mt-0.5" />
              <div>
                <h4 className="text-xs font-bold text-[#DC2626]">Audit Flags & Missing Denominators Detected</h4>
                <ul className="mt-1 space-y-1 text-xs text-[#991B1B]">
                  {analysisResult.alerts.map((alert, i) => (
                    <li key={i}>• {alert}</li>
                  ))}
                </ul>
              </div>
            </div>
          </div>
        )}

        {/* KPI Summary Cards */}
        {analysisResult.ratios.length > 0 && (
          <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
              <span className="text-xs font-semibold text-gray-600 uppercase">Net Sales / Revenue</span>
              <div className="text-lg font-bold text-gray-900 mt-1">
                ₹{analysisResult.cyInputs.revenue.toLocaleString('en-IN')}
              </div>
              <div className="text-xs text-gray-600 mt-0.5">
                PY: ₹{analysisResult.pyInputs.revenue.toLocaleString('en-IN')}
              </div>
            </div>

            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
              <span className="text-xs font-semibold text-gray-600 uppercase">Net Profit After Tax</span>
              <div className="text-lg font-bold text-gray-900 mt-1">
                ₹{analysisResult.cyInputs.profitAfterTax.toLocaleString('en-IN')}
              </div>
              <div className="text-xs text-gray-600 mt-0.5">
                PAT Margin: {analysisResult.cyInputs.revenue > 0 ? ((analysisResult.cyInputs.profitAfterTax / analysisResult.cyInputs.revenue) * 100).toFixed(2) : 0}%
              </div>
            </div>

            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
              <span className="text-xs font-semibold text-gray-600 uppercase">Current Assets vs Liab.</span>
              <div className="text-lg font-bold text-gray-900 mt-1">
                ₹{analysisResult.cyInputs.currentAssets.toLocaleString('en-IN')}
              </div>
              <div className="text-xs text-gray-600 mt-0.5">
                Liabilities: ₹{analysisResult.cyInputs.currentLiabilities.toLocaleString('en-IN')}
              </div>
            </div>

            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
              <span className="text-xs font-semibold text-gray-600 uppercase">Auditor Health Index</span>
              <div className="flex items-center gap-2 mt-1">
                <span className="px-2 py-0.5 text-xs font-bold rounded bg-[#F3F4F6] text-gray-900">
                  {healthStats.healthy} Healthy
                </span>
                {healthStats.concern > 0 && (
                  <span className="px-2 py-0.5 text-xs font-bold rounded bg-[#FEE2E2] text-[#DC2626]">
                    {healthStats.concern} Concern
                  </span>
                )}
              </div>
              <div className="text-xs text-gray-600 mt-0.5">
                Total Ratios Evaluated: {healthStats.total}
              </div>
            </div>
          </div>
        )}

        {/* Empty State */}
        {rawRows.length === 0 ? (
          <div className="bg-white border border-dashed border-gray-200 rounded-2xl p-12 text-center shadow-sm">
            <div className="w-16 h-16 bg-gray-50 rounded-full flex items-center justify-center mx-auto mb-4 border border-gray-300/20">
              <Calculator className="w-8 h-8 text-gray-900" />
            </div>
            <h3 className="text-base font-bold text-gray-900">No Financial Statement Loaded</h3>
            <p className="text-xs text-gray-600 max-w-md mx-auto mt-1 mb-6">
              Upload your client's audited or provisional Balance Sheet and Profit & Loss Statement (Excel or CSV), or load our standard test dataset.
            </p>
            <div className="flex items-center justify-center gap-3">
              <button
                onClick={() => fileInputRef.current?.click()}
                className="px-4 py-2 text-xs font-semibold rounded-lg bg-gray-900 text-white hover:bg-black transition-colors shadow-sm inline-flex items-center gap-2 cursor-pointer"
              >
                <UploadCloud className="w-4 h-4" />
                <span>Upload Financial Statement</span>
              </button>
            </div>
          </div>
        ) : (
          <div className="space-y-6">
            {/* Tabs Bar */}
            <div className="flex items-center justify-between border-b border-gray-200 pb-1 print:hidden">
              <div className="flex items-center gap-2">
                <button
                  onClick={() => setActiveTab('ratios')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'ratios'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Financial Ratios Matrix ({analysisResult.ratios.length})
                </button>
                <button
                  onClick={() => setActiveTab('inputs')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'inputs'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Mapped Statement Inputs
                </button>
                <button
                  onClick={() => setActiveTab('formulas')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'formulas'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Formula Transparency & Notes
                </button>
                <button
                  onClick={() => setActiveTab('printable_report')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'printable_report'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Executive Audit Report Preview
                </button>
              </div>

              {activeTab === 'ratios' && (
                <div className="flex items-center gap-2 text-xs">
                  <span className="text-gray-600 font-medium">Filter Category:</span>
                  <select
                    value={selectedCategory}
                    onChange={(e) => setSelectedCategory(e.target.value)}
                    className="border border-gray-200 rounded px-2 py-1 text-xs outline-none focus:border-gray-300"
                  >
                    <option value="all">All Categories</option>
                    <option value="Liquidity">Liquidity</option>
                    <option value="Solvency">Solvency</option>
                    <option value="Profitability">Profitability</option>
                    <option value="Activity">Activity / Turnover</option>
                  </select>
                </div>
              )}
            </div>

            {/* TAB 1: Ratios Matrix */}
            {activeTab === 'ratios' && (
              <div className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm">
                <div className="overflow-x-auto">
                  <table className="w-full text-xs text-left">
                    <thead className="bg-gray-900 text-white font-semibold">
                      <tr>
                        <th className="p-3 w-12 text-center">#</th>
                        <th className="p-3">Category</th>
                        <th className="p-3">Financial Ratio</th>
                        <th className="p-3 text-right">{cyPeriod} (CY)</th>
                        <th className="p-3 text-right">{pyPeriod} (PY)</th>
                        <th className="p-3 text-center">YoY Trend</th>
                        <th className="p-3">Benchmark</th>
                        <th className="p-3 text-center">Audit Status</th>
                        <th className="p-3">Formula Substituted & Auditor Observation</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB]">
                      {filteredRatios.map((r, idx) => {
                        const trend =
                          r.cyValue !== null && r.pyValue !== null
                            ? r.cyValue > r.pyValue
                              ? 'up'
                              : r.cyValue < r.pyValue
                              ? 'down'
                              : 'flat'
                            : 'na';

                        return (
                          <tr key={r.id} className="hover:bg-gray-50 transition-colors">
                            <td className="p-3 text-center text-gray-600 font-medium">{idx + 1}</td>
                            <td className="p-3 font-semibold text-gray-600">{r.category}</td>
                            <td className="p-3 font-bold text-gray-900">
                              {r.name}
                              <div className="text-[11px] text-gray-600 font-normal">{r.formulaDescription}</div>
                            </td>
                            <td className="p-3 text-right font-bold text-gray-900 text-sm">
                              {r.cyValue !== null ? `${r.cyValue.toFixed(2)}${r.unit === 'percentage' ? '%' : r.unit === 'days' ? ' d' : r.unit === 'times' ? 'x' : ''}` : 'N/A'}
                            </td>
                            <td className="p-3 text-right text-gray-600">
                              {r.pyValue !== null ? `${r.pyValue.toFixed(2)}${r.unit === 'percentage' ? '%' : r.unit === 'days' ? ' d' : r.unit === 'times' ? 'x' : ''}` : 'N/A'}
                            </td>
                            <td className="p-3 text-center">
                              {trend === 'up' && (
                                <span className="inline-flex items-center text-gray-900 font-bold gap-0.5">
                                  <TrendingUp className="w-3.5 h-3.5" /> Up
                                </span>
                              )}
                              {trend === 'down' && (
                                <span className="inline-flex items-center text-[#DC2626] font-bold gap-0.5">
                                  <TrendingDown className="w-3.5 h-3.5" /> Down
                                </span>
                              )}
                              {trend === 'flat' && (
                                <span className="inline-flex items-center text-gray-600 gap-0.5">
                                  <Minus className="w-3.5 h-3.5" /> Stable
                                </span>
                              )}
                              {trend === 'na' && <span className="text-gray-600">-</span>}
                            </td>
                            <td className="p-3 font-medium text-gray-600">{r.benchmark}</td>
                            <td className="p-3 text-center">
                              {r.interpretation === 'healthy' && (
                                <span className="px-2 py-0.5 rounded font-bold bg-[#F3F4F6] text-gray-900 text-[11px]">
                                  HEALTHY
                                </span>
                              )}
                              {r.interpretation === 'neutral' && (
                                <span className="px-2 py-0.5 rounded font-bold bg-[#FEF3C7] text-[#D97706] text-[11px]">
                                  ACCEPTABLE
                                </span>
                              )}
                              {r.interpretation === 'concern' && (
                                <span className="px-2 py-0.5 rounded font-bold bg-[#FEE2E2] text-[#DC2626] text-[11px]">
                                  CONCERN
                                </span>
                              )}
                              {r.interpretation === 'not_applicable' && (
                                <span className="px-2 py-0.5 rounded font-bold bg-gray-100 text-gray-600 text-[11px]">
                                  N/A
                                </span>
                              )}
                            </td>
                            <td className="p-3">
                              <div className="font-mono text-[11px] text-gray-900 bg-gray-50 px-2 py-0.5 rounded inline-block">
                                CY: {r.cyFormulaExpanded}
                              </div>
                              <div className="text-[11px] text-gray-600 mt-1">{r.notes}</div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* TAB 2: Statement Inputs */}
            {activeTab === 'inputs' && (
              <div className="bg-white border border-gray-200 rounded-xl p-5 shadow-sm space-y-4">
                <div>
                  <h3 className="text-sm font-bold text-gray-900">Extracted Financial Statement Inputs</h3>
                  <p className="text-xs text-gray-600">
                    Aggregated balance sheet and profit & loss figures used across all financial ratio calculations
                  </p>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  {/* Balance Sheet Items */}
                  <div className="border border-gray-200 rounded-lg p-4">
                    <h4 className="text-xs font-bold text-gray-900 uppercase border-b border-gray-200 pb-2 mb-3">
                      Balance Sheet Position (₹)
                    </h4>
                    <div className="space-y-2 text-xs">
                      <div className="flex justify-between py-1 border-b border-gray-100">
                        <span className="text-gray-600">Current Assets</span>
                        <span className="font-bold">₹{analysisResult.cyInputs.currentAssets.toLocaleString('en-IN')}</span>
                      </div>
                      <div className="flex justify-between py-1 border-b border-gray-100 pl-3">
                        <span className="text-gray-600">Inventories / Stock</span>
                        <span>₹{analysisResult.cyInputs.inventories.toLocaleString('en-IN')}</span>
                      </div>
                      <div className="flex justify-between py-1 border-b border-gray-100 pl-3">
                        <span className="text-gray-600">Trade Receivables (Debtors)</span>
                        <span>₹{analysisResult.cyInputs.tradeReceivables.toLocaleString('en-IN')}</span>
                      </div>
                      <div className="flex justify-between py-1 border-b border-gray-100 pl-3">
                        <span className="text-gray-600">Cash & Bank Balances</span>
                        <span>₹{analysisResult.cyInputs.cashAndBank.toLocaleString('en-IN')}</span>
                      </div>
                      <div className="flex justify-between py-1 border-b border-gray-100">
                        <span className="text-gray-600">Current Liabilities</span>
                        <span className="font-bold">₹{analysisResult.cyInputs.currentLiabilities.toLocaleString('en-IN')}</span>
                      </div>
                      <div className="flex justify-between py-1 border-b border-gray-100 pl-3">
                        <span className="text-gray-600">Trade Payables (Creditors)</span>
                        <span>₹{analysisResult.cyInputs.tradePayables.toLocaleString('en-IN')}</span>
                      </div>
                      <div className="flex justify-between py-1 border-b border-gray-100">
                        <span className="text-gray-600">Total Debt (Long + Short Term)</span>
                        <span className="font-bold">₹{analysisResult.cyInputs.totalDebt.toLocaleString('en-IN')}</span>
                      </div>
                      <div className="flex justify-between py-1 border-b border-gray-100">
                        <span className="text-gray-600">Shareholders Equity (Net Worth)</span>
                        <span className="font-bold">₹{analysisResult.cyInputs.shareholdersEquity.toLocaleString('en-IN')}</span>
                      </div>
                    </div>
                  </div>

                  {/* P&L Items */}
                  <div className="border border-gray-200 rounded-lg p-4">
                    <h4 className="text-xs font-bold text-gray-900 uppercase border-b border-gray-200 pb-2 mb-3">
                      Statement of Profit & Loss (₹)
                    </h4>
                    <div className="space-y-2 text-xs">
                      <div className="flex justify-between py-1 border-b border-gray-100">
                        <span className="text-gray-600">Revenue from Operations (Gross)</span>
                        <span className="font-bold">₹{analysisResult.cyInputs.revenue.toLocaleString('en-IN')}</span>
                      </div>
                      <div className="flex justify-between py-1 border-b border-gray-100">
                        <span className="text-gray-600">Cost of Goods Sold (COGS)</span>
                        <span>₹{analysisResult.cyInputs.cogs.toLocaleString('en-IN')}</span>
                      </div>
                      <div className="flex justify-between py-1 border-b border-gray-100">
                        <span className="text-gray-600">Gross Profit</span>
                        <span className="font-bold">₹{analysisResult.cyInputs.grossProfit.toLocaleString('en-IN')}</span>
                      </div>
                      <div className="flex justify-between py-1 border-b border-gray-100">
                        <span className="text-gray-600">Operating Expenses</span>
                        <span>₹{analysisResult.cyInputs.operatingExpenses.toLocaleString('en-IN')}</span>
                      </div>
                      <div className="flex justify-between py-1 border-b border-gray-100">
                        <span className="text-gray-600">Operating Profit (EBIT)</span>
                        <span className="font-bold">₹{analysisResult.cyInputs.operatingProfit.toLocaleString('en-IN')}</span>
                      </div>
                      <div className="flex justify-between py-1 border-b border-gray-100">
                        <span className="text-gray-600">Finance Costs (Interest)</span>
                        <span>₹{analysisResult.cyInputs.financeCosts.toLocaleString('en-IN')}</span>
                      </div>
                      <div className="flex justify-between py-1 border-b border-gray-100">
                        <span className="text-gray-600">Depreciation & Amortisation</span>
                        <span>₹{analysisResult.cyInputs.depreciation.toLocaleString('en-IN')}</span>
                      </div>
                      <div className="flex justify-between py-1 border-b border-gray-100">
                        <span className="text-gray-600">Profit After Tax (PAT / Net Profit)</span>
                        <span className="font-bold text-gray-900">₹{analysisResult.cyInputs.profitAfterTax.toLocaleString('en-IN')}</span>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            )}

            {/* TAB 3: Formula Transparency */}
            {activeTab === 'formulas' && (
              <div className="bg-white border border-gray-200 rounded-xl p-5 shadow-sm space-y-4">
                <div>
                  <h3 className="text-sm font-bold text-gray-900">Mathematical Formulas & Statutory Benchmarks</h3>
                  <p className="text-xs text-gray-600">
                    Auditor reference guide for banking covenants, ICAI standards, and working capital credit appraisal
                  </p>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {analysisResult.ratios.map((r) => (
                    <div key={r.id} className="border border-gray-200 rounded-lg p-3 bg-gray-50/50">
                      <div className="flex items-center justify-between">
                        <span className="font-bold text-xs text-gray-900">{r.name}</span>
                        <span className="text-[11px] font-semibold text-gray-900 bg-[#F3F4F6] px-2 py-0.5 rounded">
                          Benchmark: {r.benchmark}
                        </span>
                      </div>
                      <div className="mt-1 font-mono text-[11px] text-gray-600">
                        Formula: {r.formulaDescription}
                      </div>
                      <div className="mt-2 text-xs text-gray-900">
                        <span className="font-semibold text-gray-900">CY Calculation:</span> {r.cyFormulaExpanded} ={' '}
                        <span className="font-bold">{r.cyValue !== null ? r.cyValue.toFixed(2) : 'N/A'}</span>
                      </div>
                      <p className="mt-1.5 text-[11px] text-gray-600 leading-relaxed">
                        {r.notes}
                      </p>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* TAB 4: Executive Report Preview */}
            {activeTab === 'printable_report' && (
              <div className="bg-white border border-gray-200 rounded-xl p-8 shadow-sm print:border-none print:shadow-none print:p-0">
                {/* Printable Header */}
                <div className="border-b-2 border-gray-300 pb-4 mb-6">
                  <div className="flex justify-between items-start">
                    <div>
                      <h2 className="text-xl font-bold text-gray-900">FINANCIAL RATIO ANALYSIS REPORT</h2>
                      <p className="text-xs text-gray-600 mt-0.5">
                        Statutory Audit & Working Capital Appraisal • Formulated under ICAI Guidelines
                      </p>
                    </div>
                    <div className="text-right text-xs">
                      <div className="font-bold text-gray-900">{clientName}</div>
                      <div className="text-gray-600">Comparative: {cyPeriod} vs {pyPeriod}</div>
                      <div className="text-gray-600">Generated: {new Date().toLocaleDateString('en-IN')}</div>
                    </div>
                  </div>
                </div>

                {/* Executive Summary Narrative */}
                <div className="mb-6 p-4 bg-[#F8F9FA] rounded-lg border border-gray-200 text-xs space-y-2">
                  <h4 className="font-bold text-gray-900">Executive Analytical Summary</h4>
                  <p className="text-gray-600 leading-relaxed">
                    Based on the audited/provisional financial statements for the period ended {cyPeriod}, the entity has evaluated {healthStats.total} core financial indicators across liquidity, solvency, operating profitability, and working capital turnaround. Overall, <span className="font-bold text-gray-900">{healthStats.healthy} ratios</span> satisfy standard prudential benchmarks, while <span className="font-bold text-[#DC2626]">{healthStats.concern} areas</span> require analytical scrutiny by the statutory auditor or management.
                  </p>
                </div>

                {/* Printable Ratios Table */}
                <table className="w-full text-xs text-left mb-6 border-collapse">
                  <thead>
                    <tr className="bg-gray-900 text-white">
                      <th className="p-2 border border-gray-300">Category</th>
                      <th className="p-2 border border-gray-300">Financial Ratio</th>
                      <th className="p-2 border border-gray-300 text-right">{cyPeriod} (CY)</th>
                      <th className="p-2 border border-gray-300 text-right">{pyPeriod} (PY)</th>
                      <th className="p-2 border border-gray-300">ICAI / Bank Benchmark</th>
                      <th className="p-2 border border-gray-300 text-center">Status</th>
                      <th className="p-2 border border-gray-300">Auditor Notes</th>
                    </tr>
                  </thead>
                  <tbody>
                    {analysisResult.ratios.map((r) => (
                      <tr key={r.id} className="border-b border-gray-200">
                        <td className="p-2 border border-gray-300 font-semibold text-gray-600">{r.category}</td>
                        <td className="p-2 border border-gray-300 font-bold text-gray-900">{r.name}</td>
                        <td className="p-2 border border-gray-300 text-right font-bold">
                          {r.cyValue !== null ? `${r.cyValue.toFixed(2)}${r.unit === 'percentage' ? '%' : ''}` : 'N/A'}
                        </td>
                        <td className="p-2 border border-gray-300 text-right text-gray-600">
                          {r.pyValue !== null ? `${r.pyValue.toFixed(2)}${r.unit === 'percentage' ? '%' : ''}` : 'N/A'}
                        </td>
                        <td className="p-2 border border-gray-300">{r.benchmark}</td>
                        <td className="p-2 border border-gray-300 text-center font-bold text-[10px]">
                          {r.interpretation.toUpperCase()}
                        </td>
                        <td className="p-2 border border-gray-300 text-[10px] text-gray-600">{r.notes}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>

                {/* Sign-off footer */}
                <div className="mt-12 pt-6 border-t border-gray-200 flex justify-between text-xs text-gray-600">
                  <div>
                    <div>Prepared by: Audit Assistant / Articleship Trainee</div>
                    <div>Verified by: Partner / Senior Chartered Accountant</div>
                  </div>
                  <div className="text-right">
                    <div className="font-semibold text-gray-900">{clientName}</div>
                    <div>Authorized Signatory & Stamp</div>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
};