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
  FileText,
  Percent,
  Calendar,
  Printer,
  HelpCircle,
} from 'lucide-react';
import { InnerToolBadge } from '../common/InnerToolBadge';
import {
  type TdsRecord,
  type TdsReconciliationSummary,
  type TdsColumnMapping,
  DEFAULT_TDS_SECTIONS,
  detectTdsColumns,
  reconcileTdsRecords,
  exportTdsWorkbook,
} from '../../utils/tdsReconciliationEngine';

interface TdsReconciliationToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

export const TdsReconciliationTool: React.FC<TdsReconciliationToolProps> = ({
  onBack,
  onShowToast,
}) => {
  const [rawRows, setRawRows] = useState<Record<string, any>[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [mapping, setMapping] = useState<TdsColumnMapping>({
    deducteeCol: '',
    panCol: '',
    sectionCol: '',
    bookingDateCol: '',
    taxableCol: '',
    tdsDeductedCol: '',
    tdsDepositedCol: '',
    challanNoCol: '',
    bsrCodeCol: '',
    challanDateCol: '',
  });
  const [fileName, setFileName] = useState('');
  const [clientName, setClientName] = useState('');
  const [financialYear, setFinancialYear] = useState('FY 2026-27');
  const [activeTab, setActiveTab] = useState<'deductees' | 'sections' | 'interest_schedule' | 'printable_report'>('deductees');
  const [searchQuery, setSearchQuery] = useState('');
  const [sectionFilter, setSectionFilter] = useState('all');
  const [statusFilter, setStatusFilter] = useState('all');
  const [showMappingPanel, setShowMappingPanel] = useState(false);
  const [showConfigPanel, setShowConfigPanel] = useState(false);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Reconciliation summary
  const summary: TdsReconciliationSummary = useMemo(() => {
    if (rawRows.length === 0) {
      return {
        totalRecords: 0,
        totalTaxableValue: 0,
        totalTdsDeducted: 0,
        totalTdsDeposited: 0,
        reconciledCount: 0,
        unreconciledCount: 0,
        missingChallanCount: 0,
        invalidPanCount: 0,
        lateDepositCount: 0,
        totalEstimatedInterest: 0,
        netTdsVariance: 0,
        isReconciled: true,
        records: [],
        sectionBreakdown: [],
      };
    }

    return reconcileTdsRecords(rawRows, mapping, financialYear);
  }, [rawRows, mapping, financialYear]);

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
          onShowToast('Empty File', 'The uploaded TDS file contains no data rows.', 'error');
          return;
        }

        const detectedHeaders = Object.keys(json[0]);
        setHeaders(detectedHeaders);
        setRawRows(json);

        const auto = detectTdsColumns(detectedHeaders);
        setMapping(auto);
      } catch (err: any) {
        onShowToast('Upload Error', err.message || 'Failed to read file', 'error');
      }
    };

    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  // Export Excel workbook
  const handleExportExcel = () => {
    if (summary.records.length === 0) {
      onShowToast('No Data', 'Please load TDS records before exporting.', 'error');
      return;
    }

    try {
      const wb = exportTdsWorkbook(summary, clientName, financialYear);
      XLSX.writeFile(wb, `TDS_Reconciliation_${clientName.replace(/\s+/g, '_')}_${financialYear}.xlsx`);
      onShowToast('Workbook Exported', 'Exported 3-sheet TDS Reconciliation report (Summary, Deductees, Sections).', 'success');
    } catch (err: any) {
      onShowToast('Export Error', err.message || 'Failed to generate Excel file.', 'error');
    }
  };

  // Print PDF
  const handlePrint = () => {
    if (summary.records.length === 0) {
      onShowToast('No Data', 'Please load TDS records before printing.', 'error');
      return;
    }
    setActiveTab('printable_report');
    setTimeout(() => {
      window.print();
    }, 300);
  };

  // Filter records
  const filteredRecords = useMemo(() => {
    return summary.records.filter((r) => {
      if (sectionFilter !== 'all' && !r.section.includes(sectionFilter)) return false;
      if (statusFilter !== 'all' && r.status !== statusFilter) return false;

      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase();
        const name = r.deducteeName.toLowerCase();
        const pan = r.pan.toLowerCase();
        const sec = r.section.toLowerCase();
        const challan = r.challanNo.toLowerCase();
        return name.includes(q) || pan.includes(q) || sec.includes(q) || challan.includes(q);
      }

      return true;
    });
  }, [summary.records, sectionFilter, statusFilter, searchQuery]);

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="tds-reconciliation"
        title="TDS Reconciliation Tool"
        description="Reconcile TDS deducted vs deposited, verify 206AA PAN compliance, and compute Section 201(1A) late interest."
        badge="Statutory"
        onBack={onBack}
        actionButtons={
          <div className="flex flex-wrap items-center gap-2">
            

            <button
              onClick={() => fileInputRef.current?.click()}
              className="px-3 py-1.5 text-xs font-semibold rounded-md bg-white text-gray-900 border border-gray-200 hover:bg-gray-50 transition-colors flex items-center gap-1.5 shadow-sm"
            >
              <UploadCloud className="w-3.5 h-3.5 text-gray-900" />
              {fileName ? 'Re-upload File' : 'Import TDS File'}
            </button>
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleFileUpload}
              accept=".xlsx,.xls,.csv"
              className="hidden"
              id="tds-recon-upload"
            />

            <button
              onClick={() => setShowConfigPanel(!showConfigPanel)}
              className="px-3 py-1.5 text-xs font-semibold rounded-md bg-white text-gray-900 border border-gray-200 hover:bg-gray-50 transition-colors flex items-center gap-1.5"
            >
              <Percent className="w-3.5 h-3.5 text-gray-900" />
              Statutory Rates
            </button>

            <button
              onClick={() => setShowMappingPanel(!showMappingPanel)}
              className={`px-3 py-1.5 text-xs font-semibold rounded-md border transition-colors flex items-center gap-1.5 ${
                showMappingPanel
                  ? 'bg-gray-900 text-white border-gray-300'
                  : 'bg-white text-gray-900 border-gray-200 hover:bg-gray-50'
              }`}
            >
              <Sliders className="w-3.5 h-3.5" />
              Map Fields
            </button>

            <button
              onClick={handlePrint}
              disabled={summary.records.length === 0}
              className="px-3 py-1.5 text-xs font-semibold rounded-md bg-white text-gray-900 border border-gray-200 hover:bg-gray-50 disabled:opacity-50 transition-colors flex items-center gap-1.5"
            >
              <Printer className="w-3.5 h-3.5 text-gray-600" />
              PDF Report
            </button>

            <button
              onClick={handleExportExcel}
              disabled={summary.records.length === 0}
              className="px-3.5 py-1.5 text-xs font-semibold rounded-md bg-gray-900 text-white hover:bg-black disabled:opacity-50 transition-colors flex items-center gap-1.5 shadow-sm cursor-pointer"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              Export Excel (.xlsx)
            </button>
          </div>
        }
      />

      {/* Client & Period Header */}
      <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-2xs flex flex-wrap items-center gap-4 text-xs">
          <div className="flex items-center gap-2">
            <Building className="w-4 h-4 text-gray-600" />
            <span className="font-semibold text-gray-600">Entity / TAN Holder:</span>
            <input
              type="text"
              value={clientName}
              onChange={(e) => setClientName(e.target.value)}
              className="px-2 py-1 border border-gray-200 rounded text-xs font-medium focus:outline-none focus:border-gray-300 w-56"
            />
          </div>

          <div className="flex items-center gap-2">
            <Calendar className="w-4 h-4 text-gray-600" />
            <span className="font-semibold text-gray-600">Financial Year:</span>
            <input
              type="text"
              value={financialYear}
              onChange={(e) => setFinancialYear(e.target.value)}
              className="px-2 py-1 border border-gray-200 rounded text-xs font-medium focus:outline-none focus:border-gray-300 w-28"
            />
          </div>

          {fileName && (
            <div className="ml-auto text-xs text-gray-600">
              Source: <span className="font-semibold text-gray-900">{fileName}</span> ({rawRows.length} deductees)
            </div>
          )}
      </div>

      {/* Upload Dropzone when no file selected */}
      {!fileName && (
        <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-8 sm:p-10 text-center transition-colors shadow-2xs">
          <label htmlFor="tds-recon-upload" className="cursor-pointer flex flex-col items-center">
            <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mx-auto mb-3 shadow-2xs">
              <UploadCloud className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-semibold text-gray-900">Upload TDS Deductions or 26AS Spreadsheet (.xlsx, .xls, .csv)</h3>
            <p className="text-xs text-gray-500 mt-1 max-w-md mx-auto leading-relaxed">
              Upload client TDS registers with party PAN, gross contract amount, section (194C, 194J, 194I, 194Q), and tax deducted. Automatically audits rate compliance and threshold breaches.
            </p>
            <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
              <UploadCloud className="w-4 h-4" />
              <span>Import TDS Spreadsheet</span>
            </div>
            <span className="text-[11px] text-gray-400 font-medium mt-3 block">
              Supports .xlsx, .xls, .csv · Finance Act 2024 / FY 2026-27 statutory rate validation
            </span>
          </label>
        </div>
      )}

      {/* Main Container */}
      <div className="space-y-6">
        {/* Statutory Disclaimer Alert */}
        <div className="bg-gray-50 border border-gray-300/30 rounded-xl p-3 text-xs text-gray-900 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <ShieldCheck className="w-4 h-4 shrink-0 text-gray-900" />
            <span>
              <strong>Statutory Audit Disclosure:</strong> This tool performs mathematical and data reconciliation. It does not certify that Form 24Q/26Q has been formally filed on TRACES or with the Income Tax Department.
            </span>
          </div>
        </div>

        {/* Configuration Drawer */}
        {showConfigPanel && (
          <div className="bg-white border border-gray-200 rounded-xl p-5 shadow-sm space-y-3">
            <div className="flex items-center justify-between pb-3 border-b border-gray-200">
              <div>
                <h3 className="text-sm font-bold text-gray-900">Statutory TDS Rate Matrix (Finance Act 2024 / FY 2026-27)</h3>
                <p className="text-xs text-gray-600">Active benchmark rates and annual thresholds applied during reconciliation</p>
              </div>
              <button onClick={() => setShowConfigPanel(false)} className="text-xs text-gray-600 hover:text-gray-900">Close</button>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-4 gap-3 text-xs">
              {Object.values(DEFAULT_TDS_SECTIONS).map((rule) => (
                <div key={rule.section} className="border border-gray-200 rounded-lg p-3 bg-gray-50/50">
                  <div className="flex justify-between items-center font-bold">
                    <span className="text-gray-900">Sec {rule.section}</span>
                    <span className="bg-[#F3F4F6] text-gray-900 px-2 py-0.5 rounded text-[11px]">{rule.standardRate}%</span>
                  </div>
                  <div className="text-[11px] font-semibold text-gray-900 mt-1">{rule.name}</div>
                  <div className="text-[10px] text-gray-600 mt-1">Threshold: ₹{rule.thresholdLimit.toLocaleString('en-IN')}</div>
                  <div className="text-[10px] text-gray-600 mt-0.5">{rule.description}</div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* Mapping Drawer */}
        {showMappingPanel && (
          <div className="bg-white border border-gray-200 rounded-xl p-5 shadow-sm space-y-4">
            <div className="flex items-center justify-between pb-3 border-b border-gray-200">
              <div>
                <h3 className="text-sm font-bold text-gray-900">Column Mapping: TDS Deductee & Challan Data</h3>
                <p className="text-xs text-gray-600">Map your raw workbook columns to standard TDS audit fields</p>
              </div>
              <button onClick={() => setShowMappingPanel(false)} className="text-xs text-gray-600 hover:text-gray-900">Close</button>
            </div>

            <div className="grid grid-cols-2 md:grid-cols-5 gap-4 text-xs">
              <div>
                <label className="block text-gray-600 mb-1 font-semibold">Deductee Name *</label>
                <select
                  value={mapping.deducteeCol}
                  onChange={(e) => setMapping({ ...mapping, deducteeCol: e.target.value })}
                  className="w-full border border-gray-200 rounded p-1.5 outline-none"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-gray-600 mb-1 font-semibold">PAN *</label>
                <select
                  value={mapping.panCol}
                  onChange={(e) => setMapping({ ...mapping, panCol: e.target.value })}
                  className="w-full border border-gray-200 rounded p-1.5 outline-none"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-gray-600 mb-1 font-semibold">TDS Section *</label>
                <select
                  value={mapping.sectionCol}
                  onChange={(e) => setMapping({ ...mapping, sectionCol: e.target.value })}
                  className="w-full border border-gray-200 rounded p-1.5 outline-none"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-gray-600 mb-1 font-semibold">Taxable Amount *</label>
                <select
                  value={mapping.taxableCol}
                  onChange={(e) => setMapping({ ...mapping, taxableCol: e.target.value })}
                  className="w-full border border-gray-200 rounded p-1.5 outline-none"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-gray-600 mb-1 font-semibold">TDS Deducted *</label>
                <select
                  value={mapping.tdsDeductedCol}
                  onChange={(e) => setMapping({ ...mapping, tdsDeductedCol: e.target.value })}
                  className="w-full border border-gray-200 rounded p-1.5 outline-none"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-gray-600 mb-1 font-semibold">TDS Deposited</label>
                <select
                  value={mapping.tdsDepositedCol}
                  onChange={(e) => setMapping({ ...mapping, tdsDepositedCol: e.target.value })}
                  className="w-full border border-gray-200 rounded p-1.5 outline-none"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-gray-600 mb-1 font-semibold">Challan No / CIN</label>
                <select
                  value={mapping.challanNoCol}
                  onChange={(e) => setMapping({ ...mapping, challanNoCol: e.target.value })}
                  className="w-full border border-gray-200 rounded p-1.5 outline-none"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-gray-600 mb-1 font-semibold">BSR Code</label>
                <select
                  value={mapping.bsrCodeCol}
                  onChange={(e) => setMapping({ ...mapping, bsrCodeCol: e.target.value })}
                  className="w-full border border-gray-200 rounded p-1.5 outline-none"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-gray-600 mb-1 font-semibold">Challan Date</label>
                <select
                  value={mapping.challanDateCol}
                  onChange={(e) => setMapping({ ...mapping, challanDateCol: e.target.value })}
                  className="w-full border border-gray-200 rounded p-1.5 outline-none"
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
              <FileText className="w-8 h-8 text-gray-900" />
            </div>
            <h3 className="text-base font-bold text-gray-900">No TDS Register Loaded</h3>
            <p className="text-xs text-gray-600 max-w-md mx-auto mt-1 mb-6">
              Upload your client's TDS deduction register, Tally ledger export, or Form 26AS/challan payment details to run automated compliance checks.
            </p>
            <div className="flex items-center justify-center gap-3">
              
              <button
                onClick={() => fileInputRef.current?.click()}
                className="px-4 py-2 text-xs font-semibold rounded-lg bg-gray-900 text-white hover:bg-black transition-colors shadow-sm inline-flex items-center gap-2"
              >
                <UploadCloud className="w-4 h-4" />
                <span>Upload TDS Deduction Register</span>
              </button>
            </div>
          </div>
        ) : (
          <div className="space-y-6">
            {/* KPI Summary Cards */}
            <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
              <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
                <span className="text-xs font-semibold text-gray-600 uppercase">Total Taxable Value</span>
                <div className="text-lg font-bold text-gray-900 mt-1">
                  ₹{summary.totalTaxableValue.toLocaleString('en-IN')}
                </div>
                <div className="text-xs text-gray-600 mt-0.5">
                  Across {summary.totalRecords} deductees
                </div>
              </div>

              <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
                <span className="text-xs font-semibold text-gray-600 uppercase">TDS Deducted vs Paid</span>
                <div className="text-lg font-bold text-gray-900 mt-1">
                  ₹{summary.totalTdsDeducted.toLocaleString('en-IN')}
                </div>
                <div className="text-xs text-gray-600 mt-0.5">
                  Paid: ₹{summary.totalTdsDeposited.toLocaleString('en-IN')}
                </div>
              </div>

              <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
                <span className="text-xs font-semibold text-gray-600 uppercase">Net Reconciliation Var.</span>
                <div className="text-lg font-bold text-[#DC2626] mt-1">
                  ₹{summary.netTdsVariance.toLocaleString('en-IN')}
                </div>
                <div className="text-xs text-gray-600 mt-0.5">
                  {summary.reconciledCount} reconciled, {summary.unreconciledCount} flagged
                </div>
              </div>

              <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
                <span className="text-xs font-semibold text-gray-600 uppercase">Sec 206AA Invalid PAN</span>
                <div className="text-lg font-bold text-[#DC2626] mt-1">
                  {summary.invalidPanCount}
                </div>
                <div className="text-xs text-gray-600 mt-0.5">
                  20% higher deduction required
                </div>
              </div>

              <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
                <span className="text-xs font-semibold text-gray-600 uppercase">Late Interest Exposure</span>
                <div className="text-lg font-bold text-[#D97706] mt-1">
                  ₹{summary.totalEstimatedInterest.toLocaleString('en-IN')}
                </div>
                <div className="text-xs text-gray-600 mt-0.5">
                  Sec 201(1A) 1.5% simple int.
                </div>
              </div>
            </div>

            {/* Tabs Bar */}
            <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 border-b border-gray-200 pb-1 print:hidden">
              <div className="flex items-center gap-2">
                <button
                  onClick={() => setActiveTab('deductees')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'deductees'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Deductee Master Register ({summary.records.length})
                </button>
                <button
                  onClick={() => setActiveTab('sections')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'sections'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Section-Wise Breakdown
                </button>
                <button
                  onClick={() => setActiveTab('interest_schedule')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'interest_schedule'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Sec 201(1A) Late Deposit Schedule ({summary.lateDepositCount})
                </button>
                <button
                  onClick={() => setActiveTab('printable_report')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'printable_report'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Audit Certificate Preview
                </button>
              </div>

              {activeTab === 'deductees' && (
                <div className="flex items-center gap-2">
                  <div className="relative">
                    <Search className="w-3.5 h-3.5 absolute left-2.5 top-1/2 -translate-y-1/2 text-gray-600" />
                    <input
                      type="text"
                      placeholder="Search deductee, PAN, challan..."
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      className="pl-8 pr-3 py-1.5 text-xs border border-gray-200 rounded-lg w-52 focus:outline-none focus:border-gray-300"
                    />
                  </div>

                  <select
                    value={statusFilter}
                    onChange={(e) => setStatusFilter(e.target.value)}
                    className="text-xs border border-gray-200 rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-gray-300"
                  >
                    <option value="all">All Statuses</option>
                    <option value="reconciled">Reconciled</option>
                    <option value="amount_mismatch">Amount Mismatch</option>
                    <option value="missing_challan">Missing Challan</option>
                    <option value="invalid_pan_206aa">Invalid PAN (206AA)</option>
                    <option value="late_deposit_interest">Late Deposit</option>
                  </select>
                </div>
              )}
            </div>

            {/* TAB 1: Deductees Register */}
            {activeTab === 'deductees' && (
              <div className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm">
                <div className="overflow-x-auto">
                  <table className="w-full text-xs text-left">
                    <thead className="bg-gray-900 text-white font-semibold">
                      <tr>
                        <th className="p-3 w-10 text-center">Row</th>
                        <th className="p-3">Deductee Name</th>
                        <th className="p-3">PAN</th>
                        <th className="p-3">Section</th>
                        <th className="p-3 text-right">Taxable (₹)</th>
                        <th className="p-3 text-right">Deducted (₹)</th>
                        <th className="p-3 text-right">Deposited (₹)</th>
                        <th className="p-3 text-right">Variance (₹)</th>
                        <th className="p-3">Challan / BSR / Date</th>
                        <th className="p-3 text-center">Status</th>
                        <th className="p-3">Statutory Audit Observation</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB]">
                      {filteredRecords.map((r) => (
                        <tr
                          key={r.id}
                          className={`hover:bg-gray-50 transition-colors ${
                            r.status === 'invalid_pan_206aa'
                              ? 'bg-red-50/40'
                              : r.status === 'missing_challan'
                              ? 'bg-amber-50/30'
                              : r.status === 'late_deposit_interest'
                              ? 'bg-orange-50/20'
                              : ''
                          }`}
                        >
                          <td className="p-3 text-center text-gray-600 font-medium">{r.rowIndex}</td>
                          <td className="p-3 font-bold text-gray-900">{r.deducteeName}</td>
                          <td className="p-3 font-mono">
                            <div>{r.pan || 'MISSING'}</div>
                            {r.isPanValid ? (
                              <span className="text-[10px] text-gray-900 font-semibold">✓ Valid PAN</span>
                            ) : (
                              <span className="text-[10px] text-[#DC2626] font-bold">✗ Invalid (Sec 206AA)</span>
                            )}
                          </td>
                          <td className="p-3 font-semibold text-gray-900">{r.section}</td>
                          <td className="p-3 text-right font-mono">₹{r.taxableAmount.toLocaleString('en-IN')}</td>
                          <td className="p-3 text-right font-mono">
                            <div>₹{r.tdsDeducted.toLocaleString('en-IN')}</div>
                            <span className="text-[10px] text-gray-600">{r.effectiveRatePercent.toFixed(1)}%</span>
                          </td>
                          <td className="p-3 text-right font-mono">₹{r.tdsDeposited.toLocaleString('en-IN')}</td>
                          <td className="p-3 text-right font-bold font-mono">
                            {r.varianceAmount > 0 ? (
                              <span className="text-[#DC2626]">₹{r.varianceAmount.toLocaleString('en-IN')}</span>
                            ) : (
                              <span className="text-gray-900">₹0.00</span>
                            )}
                          </td>
                          <td className="p-3 text-[11px] text-gray-600">
                            <div>No: {r.challanNo || 'None'}</div>
                            <div>Date: {r.challanDate || '-'}</div>
                            {r.bsrCode && <div>BSR: {r.bsrCode}</div>}
                          </td>
                          <td className="p-3 text-center">
                            {r.status === 'reconciled' && (
                              <span className="px-2 py-0.5 rounded font-bold bg-[#F3F4F6] text-gray-900 text-[10px]">
                                RECONCILED
                              </span>
                            )}
                            {r.status === 'amount_mismatch' && (
                              <span className="px-2 py-0.5 rounded font-bold bg-[#FEE2E2] text-[#DC2626] text-[10px]">
                                AMOUNT DIFF
                              </span>
                            )}
                            {r.status === 'missing_challan' && (
                              <span className="px-2 py-0.5 rounded font-bold bg-[#FEF3C7] text-[#B45309] text-[10px]">
                                NO CHALLAN
                              </span>
                            )}
                            {r.status === 'invalid_pan_206aa' && (
                              <span className="px-2 py-0.5 rounded font-bold bg-[#FEE2E2] text-[#B91C1C] text-[10px]">
                                SEC 206AA
                              </span>
                            )}
                            {r.status === 'late_deposit_interest' && (
                              <span className="px-2 py-0.5 rounded font-bold bg-[#FFEDD5] text-[#C2410C] text-[10px]">
                                LATE DEPOSIT
                              </span>
                            )}
                          </td>
                          <td className="p-3 text-[11px] text-gray-900 max-w-xs">
                            {r.auditReasons.join(' | ')}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* TAB 2: Sections Breakdown */}
            {activeTab === 'sections' && (
              <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-sm space-y-4">
                <div>
                  <h3 className="text-sm font-bold text-gray-900">Section-Wise Deductions & Statutory Deposits</h3>
                  <p className="text-xs text-gray-600">Summary cross-tabulation of total taxable base and tax remittances</p>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-xs text-left">
                    <thead className="bg-gray-900 text-white font-semibold">
                      <tr>
                        <th className="p-3">Section</th>
                        <th className="p-3 text-center">Records Count</th>
                        <th className="p-3 text-right">Total Taxable Value (₹)</th>
                        <th className="p-3 text-right">TDS Deducted (₹)</th>
                        <th className="p-3 text-right">TDS Deposited (₹)</th>
                        <th className="p-3 text-right">Net Variance (₹)</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB]">
                      {summary.sectionBreakdown.map((s) => (
                        <tr key={s.section} className="hover:bg-gray-50">
                          <td className="p-3 font-bold text-gray-900">{s.section}</td>
                          <td className="p-3 text-center font-semibold">{s.count}</td>
                          <td className="p-3 text-right font-mono">₹{s.taxable.toLocaleString('en-IN')}</td>
                          <td className="p-3 text-right font-mono font-semibold">₹{s.deducted.toLocaleString('en-IN')}</td>
                          <td className="p-3 text-right font-mono font-semibold">₹{s.deposited.toLocaleString('en-IN')}</td>
                          <td className="p-3 text-right font-mono font-bold">
                            {s.deducted - s.deposited !== 0 ? (
                              <span className="text-[#DC2626]">₹{(s.deducted - s.deposited).toLocaleString('en-IN')}</span>
                            ) : (
                              <span className="text-gray-900">₹0.00</span>
                            )}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* TAB 3: Late Deposit Schedule */}
            {activeTab === 'interest_schedule' && (
              <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-sm space-y-4">
                <div>
                  <h3 className="text-sm font-bold text-gray-900">Section 201(1A) Late Deposit Interest Schedule</h3>
                  <p className="text-xs text-gray-600">
                    Simple interest calculated at 1.5% per month or part of a month from deduction date to actual deposit date
                  </p>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-xs text-left">
                    <thead className="bg-gray-900 text-white font-semibold">
                      <tr>
                        <th className="p-3">Deductee</th>
                        <th className="p-3">Section</th>
                        <th className="p-3">Booking Date</th>
                        <th className="p-3">Challan Date</th>
                        <th className="p-3 text-center">Delay Months</th>
                        <th className="p-3 text-right">TDS Deposited (₹)</th>
                        <th className="p-3 text-right">Interest Rate</th>
                        <th className="p-3 text-right">Estimated Interest (₹)</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB]">
                      {summary.records
                        .filter((r) => r.isLateDeposit)
                        .map((r) => (
                          <tr key={r.id} className="hover:bg-gray-50">
                            <td className="p-3 font-bold text-gray-900">{r.deducteeName}</td>
                            <td className="p-3 text-gray-900 font-semibold">{r.section}</td>
                            <td className="p-3">{r.bookingDate}</td>
                            <td className="p-3 font-semibold text-[#DC2626]">{r.challanDate}</td>
                            <td className="p-3 text-center font-bold text-amber-700">{r.delayMonths} mo</td>
                            <td className="p-3 text-right font-mono">₹{r.tdsDeposited.toLocaleString('en-IN')}</td>
                            <td className="p-3 text-right font-mono">1.5% / mo</td>
                            <td className="p-3 text-right font-mono font-bold text-[#DC2626]">
                              ₹{r.estimatedInterest201.toLocaleString('en-IN')}
                            </td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                </div>

                <div className="p-3 bg-amber-50 rounded-lg border border-amber-200 text-xs text-amber-900">
                  Total computed statutory interest liability under Section 201(1A): <strong className="font-mono text-sm">₹{summary.totalEstimatedInterest.toLocaleString('en-IN')}</strong>. This must be remitted under Minor Head 200 before quarterly 24Q/26Q return filing.
                </div>
              </div>
            )}

            {/* TAB 4: Printable Audit Report */}
            {activeTab === 'printable_report' && (
              <div className="bg-white border border-gray-200 rounded-xl p-8 shadow-sm print:border-none print:shadow-none print:p-0">
                <div className="border-b-2 border-gray-300 pb-4 mb-6">
                  <div className="flex justify-between items-start">
                    <div>
                      <h2 className="text-xl font-bold text-gray-900">TDS RECONCILIATION & AUDIT EXCEPTION REPORT</h2>
                      <p className="text-xs text-gray-600 mt-0.5">
                        Statutory Verification of Deductions, Challan Tie-outs & PAN Compliance
                      </p>
                    </div>
                    <div className="text-right text-xs">
                      <div className="font-bold text-gray-900">{clientName}</div>
                      <div className="text-gray-600">{financialYear}</div>
                      <div className="text-gray-600">Audit Date: {new Date().toLocaleDateString('en-IN')}</div>
                    </div>
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-4 mb-6 text-xs">
                  <div className="border border-gray-200 rounded p-3">
                    <span className="text-gray-600">Total Deductions Booked:</span>
                    <span className="font-bold font-mono ml-2">₹{summary.totalTdsDeducted.toLocaleString('en-IN')}</span>
                  </div>
                  <div className="border border-gray-200 rounded p-3">
                    <span className="text-gray-600">Total Challan Payments Reconciled:</span>
                    <span className="font-bold font-mono ml-2 text-gray-900">₹{summary.totalTdsDeposited.toLocaleString('en-IN')}</span>
                  </div>
                </div>

                <table className="w-full text-xs text-left mb-6 border-collapse">
                  <thead>
                    <tr className="bg-gray-900 text-white">
                      <th className="p-2 border border-gray-300">Deductee</th>
                      <th className="p-2 border border-gray-300">PAN</th>
                      <th className="p-2 border border-gray-300">Sec</th>
                      <th className="p-2 border border-gray-300 text-right">Taxable (₹)</th>
                      <th className="p-2 border border-gray-300 text-right">Deducted (₹)</th>
                      <th className="p-2 border border-gray-300 text-right">Deposited (₹)</th>
                      <th className="p-2 border border-gray-300 text-center">Status</th>
                      <th className="p-2 border border-gray-300">Observation</th>
                    </tr>
                  </thead>
                  <tbody>
                    {summary.records.map((r) => (
                      <tr key={r.id}>
                        <td className="p-2 border border-gray-300 font-bold">{r.deducteeName}</td>
                        <td className="p-2 border border-gray-300 font-mono">{r.pan}</td>
                        <td className="p-2 border border-gray-300">{r.section}</td>
                        <td className="p-2 border border-gray-300 text-right">₹{r.taxableAmount.toLocaleString('en-IN')}</td>
                        <td className="p-2 border border-gray-300 text-right">₹{r.tdsDeducted.toLocaleString('en-IN')}</td>
                        <td className="p-2 border border-gray-300 text-right">₹{r.tdsDeposited.toLocaleString('en-IN')}</td>
                        <td className="p-2 border border-gray-300 text-center font-bold text-[10px]">
                          {r.status.replace(/_/g, ' ').toUpperCase()}
                        </td>
                        <td className="p-2 border border-gray-300 text-[10px] text-gray-600">
                          {r.auditReasons.join(' | ')}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>

                <div className="mt-8 pt-6 border-t border-gray-200 text-xs text-gray-600 flex justify-between">
                  <div>Audit Assistant: Verified against Books of Account</div>
                  <div>Authorized Signatory for {clientName}</div>
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
};