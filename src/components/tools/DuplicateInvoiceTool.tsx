import { ToolHeader } from '../common/ToolHeader';
import React, { useState, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { createStyledWorksheet } from '../../utils/excelStyler';
import {
  ArrowLeft,
  UploadCloud,
  Download,
  AlertTriangle,
  AlertCircle,
  CheckCircle2,
  FileSpreadsheet,
  Sparkles,
  Sliders,
  Search,
  Check,
  X,
  HelpCircle,
  Copy,
  ShieldCheck,
  RefreshCw,
  ExternalLink,
} from 'lucide-react';
import { InnerToolBadge } from '../common/InnerToolBadge';
import {
  type InvoiceRecord,
  type DuplicateFlag,
  type DuplicateDetectionConfig,
  type InvoiceColumnMapping,
  type ReviewStatus,
  detectInvoiceColumns,
  parseInvoiceRows,
  detectDuplicateInvoices,
  exportDuplicateAuditReport,
} from '../../utils/duplicateInvoiceEngine';

interface DuplicateInvoiceToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

export const DuplicateInvoiceTool: React.FC<DuplicateInvoiceToolProps> = ({
  onBack,
  onShowToast,
}) => {
  const [rawRows, setRawRows] = useState<Record<string, any>[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [mapping, setMapping] = useState<InvoiceColumnMapping>({
    vendorCol: '',
    gstinCol: '',
    invoiceNoCol: '',
    dateCol: '',
    taxableCol: '',
    taxCol: '',
    totalCol: '',
  });
  const [fileName, setFileName] = useState('');
  const [showConfigPanel, setShowConfigPanel] = useState(false);
  const [showMappingPanel, setShowMappingPanel] = useState(false);
  const [activeTab, setActiveTab] = useState<'all' | 'exact' | 'probable' | 'confirmed' | 'valid' | 'register'>('all');
  const [searchQuery, setSearchQuery] = useState('');
  const [userFlags, setUserFlags] = useState<DuplicateFlag[]>([]);

  // Configurable rules
  const [config, setConfig] = useState<DuplicateDetectionConfig>({
    checkExact: true,
    checkNormalizedNo: true,
    checkCloseDateSameAmount: true,
    checkSameNoDiffAmount: true,
    checkSameNoDiffVendor: true,
    dateToleranceDays: 7,
    amountTolerance: 1.0,
  });

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Parse structured records when rawRows or mapping changes
  const parsedRecords = useMemo(() => {
    if (rawRows.length === 0) return [];
    return parseInvoiceRows(rawRows, mapping);
  }, [rawRows, mapping]);

  // Run duplicate detection engine
  const detectionResult = useMemo(() => {
    if (parsedRecords.length === 0) {
      return {
        totalRecords: 0,
        duplicatePairsCount: 0,
        exactCount: 0,
        probableCount: 0,
        totalDuplicateExposure: 0,
        unresolvedCount: 0,
        confirmedDuplicateCount: 0,
        markedValidCount: 0,
        flags: [],
        records: [],
      };
    }
    return detectDuplicateInvoices(parsedRecords, config, userFlags);
  }, [parsedRecords, config, userFlags]);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });
        const firstSheet = wb.Sheets[wb.SheetNames[0]];
        const json: Record<string, any>[] = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });

        if (json.length === 0) {
          onShowToast('Empty File', 'The uploaded file has no data rows.', 'error');
          return;
        }

        const detectedHeaders = Object.keys(json[0]);
        setHeaders(detectedHeaders);
        const detectedMapping = detectInvoiceColumns(detectedHeaders);
        setMapping(detectedMapping);
        setRawRows(json);
        setUserFlags([]);
      } catch (err: any) {
        onShowToast('Import Error', err.message || 'Could not parse file.', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  
  const handleDownloadTemplate = () => {
    const templateRows = [
      {
        'Vendor Name': 'Reliance Retail Limited',
        'Vendor GSTIN': '24AAACR5055K1Z1',
        'Invoice Number': 'INV/2026/001',
        'Invoice Date': '15-09-2026',
        'Taxable Amount': 150000.0,
        'GST Amount': 27000.0,
        'Total Amount': 177000.0,
        'Remarks': 'Raw materials batch A',
      },
      {
        'Vendor Name': 'Tata Consultancy Services Ltd',
        'Vendor GSTIN': '27AAACT2727Q1ZW',
        'Invoice Number': 'TCS-99120',
        'Invoice Date': '18-09-2026',
        'Taxable Amount': 85000.0,
        'GST Amount': 15300.0,
        'Total Amount': 100300.0,
        'Remarks': 'IT consulting service',
      },
    ];
    const wb = XLSX.utils.book_new();
    const ws = createStyledWorksheet(templateRows, {
      currencyColumns: ['Taxable Amount', 'GST Amount', 'Total Amount'],
      columnWidths: {
        'Vendor Name': 32,
        'Vendor GSTIN': 18,
        'Invoice Number': 18,
        'Invoice Date': 16,
        'Taxable Amount': 18,
        'GST Amount': 16,
        'Total Amount': 18,
        'Remarks': 28,
      },
      freezeHeader: true,
    });
    XLSX.utils.book_append_sheet(wb, ws, 'Purchase_Register_Template');
    XLSX.writeFile(wb, 'Duplicate_Invoice_Checker_Template.xlsx');
    onShowToast('Template Downloaded', 'Template format saved to your downloads folder.', 'info');
  };

  const handleExportExcel = () => {
    if (detectionResult.records.length === 0) return;
    try {
      const wb = exportDuplicateAuditReport(detectionResult, 'Client Audit', 'FY 2026-27');
      XLSX.writeFile(wb, `Duplicate_Invoice_Audit_Report_${new Date().toISOString().split('T')[0]}.xlsx`);
      onShowToast('Audit Report Exported', 'Saved 3-sheet duplicate findings workbook with side-by-side comparison.', 'success');
    } catch (err: any) {
      onShowToast('Export Failed', err.message || 'Could not export Excel workbook.', 'error');
    }
  };

  const handleUpdateStatus = (flagId: string, newStatus: ReviewStatus) => {
    const updated = detectionResult.flags.map((f) => {
      if (f.id === flagId) {
        return { ...f, status: newStatus };
      }
      return f;
    });
    setUserFlags(updated);
    onShowToast('Review Status Updated', `Flag marked as ${newStatus.toUpperCase()}. Source records remain untouched.`, 'info');
  };

  const filteredFlags = detectionResult.flags.filter((f) => {
    if (activeTab === 'exact' && f.matchType !== 'exact') return false;
    if (activeTab === 'probable' && f.matchType === 'exact') return false;
    if (activeTab === 'confirmed' && f.status !== 'duplicate') return false;
    if (activeTab === 'valid' && f.status !== 'valid') return false;

    if (!searchQuery) return true;
    const q = searchQuery.toLowerCase();
    return (
      f.recordA.vendorName.toLowerCase().includes(q) ||
      f.recordB.vendorName.toLowerCase().includes(q) ||
      f.recordA.invoiceNumber.toLowerCase().includes(q) ||
      f.recordB.invoiceNumber.toLowerCase().includes(q) ||
      f.recordA.gstin.toLowerCase().includes(q) ||
      f.recordB.gstin.toLowerCase().includes(q)
    );
  });

  const formatInr = (val: number) => {
    return val.toLocaleString('en-IN', { style: 'currency', currency: 'INR', maximumFractionDigits: 2 });
  };

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="duplicate-invoice"
        title="Duplicate Invoice Detector"
        description="Scan purchase registers for exact and heuristic duplicate invoices using GSTIN, invoice #, date, and amount."
        badge="Popular"
        onBack={onBack}
        actionButtons={
          <div className="flex items-center gap-2 flex-wrap">
            <button
              onClick={() => setShowConfigPanel(!showConfigPanel)}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold border transition-all cursor-pointer ${
                showConfigPanel
                  ? 'bg-gray-900 text-white border-gray-900'
                  : 'bg-white text-gray-700 border-gray-200 hover:bg-gray-50'
              }`}
            >
              <Sliders className="w-3.5 h-3.5" />
              <span>Matching Rules</span>
            </button>

            {rawRows.length > 0 && (
              <button
                onClick={() => setShowMappingPanel(!showMappingPanel)}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-white text-gray-700 border border-gray-200 hover:bg-gray-50 transition-all cursor-pointer"
              >
                <span>Columns</span>
              </button>
            )}

            <button
              onClick={handleExportExcel}
              disabled={detectionResult.records.length === 0}
              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-semibold bg-gray-900 text-white hover:bg-gray-900 disabled:opacity-40 disabled:cursor-not-allowed shadow-xs transition-all cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Export Excel Report</span>
            </button>
          </div>
        }
      />

        {/* Matching Rules Drawer */}
        {showConfigPanel && (
          <div className="mt-4 pt-4 border-t border-gray-100 bg-[#F9FAFA] -mx-5 -mb-5 p-5 rounded-b-xl space-y-4">
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-gray-900">
                Configurable Duplicate Detection Rules
              </span>
              <span className="text-[11px] text-gray-500">
                Configure tolerances and fuzzy criteria
              </span>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
              <label className="flex items-center gap-2 p-2.5 bg-white border border-gray-200 rounded-lg text-xs text-gray-800 cursor-pointer">
                <input
                  type="checkbox"
                  checked={config.checkExact}
                  onChange={(e) => setConfig({ ...config, checkExact: e.target.checked })}
                  className="rounded text-gray-700 focus:ring-blue-500"
                />
                <div>
                  <span className="font-semibold block">Exact Duplicates</span>
                  <span className="text-[10px] text-gray-500">Same Vendor/GSTIN + Inv # + Amount</span>
                </div>
              </label>

              <label className="flex items-center gap-2 p-2.5 bg-white border border-gray-200 rounded-lg text-xs text-gray-800 cursor-pointer">
                <input
                  type="checkbox"
                  checked={config.checkNormalizedNo}
                  onChange={(e) => setConfig({ ...config, checkNormalizedNo: e.target.checked })}
                  className="rounded text-gray-700 focus:ring-blue-500"
                />
                <div>
                  <span className="font-semibold block">Normalized Invoice #</span>
                  <span className="text-[10px] text-gray-500">Ignore `/`, `-`, spaces, leading zeros</span>
                </div>
              </label>

              <label className="flex items-center gap-2 p-2.5 bg-white border border-gray-200 rounded-lg text-xs text-gray-800 cursor-pointer">
                <input
                  type="checkbox"
                  checked={config.checkCloseDateSameAmount}
                  onChange={(e) => setConfig({ ...config, checkCloseDateSameAmount: e.target.checked })}
                  className="rounded text-gray-700 focus:ring-blue-500"
                />
                <div>
                  <span className="font-semibold block">Close Date & Same Amount</span>
                  <span className="text-[10px] text-gray-500">Same vendor billed within window</span>
                </div>
              </label>

              <label className="flex items-center gap-2 p-2.5 bg-white border border-gray-200 rounded-lg text-xs text-gray-800 cursor-pointer">
                <input
                  type="checkbox"
                  checked={config.checkSameNoDiffAmount}
                  onChange={(e) => setConfig({ ...config, checkSameNoDiffAmount: e.target.checked })}
                  className="rounded text-gray-700 focus:ring-blue-500"
                />
                <div>
                  <span className="font-semibold block">Same Inv #, Diff Amount</span>
                  <span className="text-[10px] text-gray-500">Potential revised bill or partial booking</span>
                </div>
              </label>

              <label className="flex items-center gap-2 p-2.5 bg-white border border-gray-200 rounded-lg text-xs text-gray-800 cursor-pointer">
                <input
                  type="checkbox"
                  checked={config.checkSameNoDiffVendor}
                  onChange={(e) => setConfig({ ...config, checkSameNoDiffVendor: e.target.checked })}
                  className="rounded text-gray-700 focus:ring-blue-500"
                />
                <div>
                  <span className="font-semibold block">Same Inv # & Amount, Diff Vendor</span>
                  <span className="text-[10px] text-gray-500">Cross-vendor double entry or typo</span>
                </div>
              </label>

              <div className="p-2.5 bg-white border border-gray-200 rounded-lg flex items-center justify-between gap-3 text-xs">
                <div>
                  <span className="font-semibold text-gray-900 block">Date Tolerance Window</span>
                  <span className="text-[10px] text-gray-500">Maximum days apart</span>
                </div>
                <select
                  value={config.dateToleranceDays}
                  onChange={(e) => setConfig({ ...config, dateToleranceDays: Number(e.target.value) })}
                  className="px-2 py-1 bg-gray-50 border border-gray-200 rounded text-xs text-gray-800 font-semibold"
                >
                  <option value={0}>0 Days (Exact Date)</option>
                  <option value={3}>± 3 Days</option>
                  <option value={7}>± 7 Days (Default)</option>
                  <option value={15}>± 15 Days</option>
                  <option value={30}>± 30 Days</option>
                </select>
              </div>
            </div>
          </div>
        )}

        {/* Column Mapping Drawer */}
        {showMappingPanel && headers.length > 0 && (
          <div className="mt-4 pt-4 border-t border-gray-100 bg-white p-4 rounded-lg border border-gray-200 space-y-3">
            <span className="text-xs font-bold text-gray-900 block">
              Header Column Mapping
            </span>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3 text-xs">
              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Vendor Name</label>
                <select
                  value={mapping.vendorCol}
                  onChange={(e) => setMapping({ ...mapping, vendorCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs"
                >
                  {headers.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">GSTIN</label>
                <select
                  value={mapping.gstinCol}
                  onChange={(e) => setMapping({ ...mapping, gstinCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs"
                >
                  <option value="">-- Optional / None --</option>
                  {headers.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Invoice Number</label>
                <select
                  value={mapping.invoiceNoCol}
                  onChange={(e) => setMapping({ ...mapping, invoiceNoCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs"
                >
                  {headers.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Total Amount</label>
                <select
                  value={mapping.totalCol}
                  onChange={(e) => setMapping({ ...mapping, totalCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs"
                >
                  {headers.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>
            </div>
          </div>
        )}

      {/* 2. File Ingestion */}
      {rawRows.length === 0 ? (
        <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-8 sm:p-10 text-center space-y-4 shadow-2xs transition-colors">
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileUpload}
            accept=".xlsx,.xls,.csv"
            className="hidden"
            id="duplicate-invoice-upload"
          />
          <label htmlFor="duplicate-invoice-upload" className="cursor-pointer flex flex-col items-center">
            <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mx-auto mb-3 shadow-2xs">
              <UploadCloud className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-semibold text-gray-900">
              Import Purchase Register or Invoices File
            </h3>
            <p className="text-xs text-gray-500 mt-1 max-w-md mx-auto leading-relaxed">
              Upload Excel (.xlsx, .xls) or CSV purchase registers containing vendor names, invoice numbers, dates, and amounts.
            </p>
            <div className="mt-4 flex flex-wrap items-center justify-center gap-2.5">
              <div className="px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
                <UploadCloud className="w-4 h-4" />
                <span>Select Spreadsheet File</span>
              </div>
              <button
                type="button"
                onClick={(e) => {
                  e.preventDefault();
                  e.stopPropagation();
                  handleDownloadTemplate();
                }}
                className="px-4 py-2 bg-white hover:bg-gray-50 text-gray-700 border border-gray-200 text-xs font-semibold rounded-lg shadow-2xs transition-all cursor-pointer"
              >
                Download Template
              </button>
            </div>
            <span className="text-[11px] text-gray-400 font-medium mt-3 block">
              100% Offline Client-Side Processing · Original Files Never Mutated
            </span>
          </label>
        </div>
      ) : (
        /* 3. Executive KPI Metric Cards */
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
          <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs">
            <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider block">
              Invoices Scanned
            </span>
            <div className="flex items-baseline justify-between mt-1.5">
              <span className="text-xl font-bold text-gray-900">
                {detectionResult.totalRecords}
              </span>
              <span className="text-[11px] text-gray-500 truncate max-w-[100px]">
                {fileName}
              </span>
            </div>
          </div>

          <div className="bg-white p-4 rounded-xl border border-red-200 shadow-xs bg-red-50/20">
            <span className="text-[11px] font-semibold text-red-700 uppercase tracking-wider block">
              Flagged Duplicate Pairs
            </span>
            <div className="flex items-baseline justify-between mt-1.5">
              <span className="text-xl font-bold text-red-700">
                {detectionResult.duplicatePairsCount}
              </span>
              <span className="text-[11px] font-semibold text-red-600">
                {detectionResult.exactCount} Exact • {detectionResult.probableCount} Heuristic
              </span>
            </div>
          </div>

          <div className="bg-white p-4 rounded-xl border border-amber-200 shadow-xs bg-amber-50/20">
            <span className="text-[11px] font-semibold text-amber-700 uppercase tracking-wider block">
              Duplicate Exposure
            </span>
            <div className="flex items-baseline justify-between mt-1.5">
              <span className="text-xl font-bold text-amber-800">
                {formatInr(detectionResult.totalDuplicateExposure)}
              </span>
              <span className="text-[11px] text-amber-700">
                Potential Overpayment
              </span>
            </div>
          </div>

          <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs">
            <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider block">
              Auditor Review Progress
            </span>
            <div className="flex items-baseline justify-between mt-1.5">
              <span className="text-xl font-bold text-gray-900">
                {detectionResult.confirmedDuplicateCount + detectionResult.markedValidCount} / {detectionResult.duplicatePairsCount}
              </span>
              <span className="text-[11px] font-semibold text-gray-800">
                {detectionResult.unresolvedCount} Unresolved
              </span>
            </div>
          </div>
        </div>
      )}

      {/* 4. Active Findings Workbench */}
      {rawRows.length > 0 && (
        <div className="bg-white rounded-xl border border-gray-200 shadow-xs overflow-hidden">
          {/* Tabs and Search Header */}
          <div className="p-4 border-b border-gray-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div className="flex items-center gap-1.5 overflow-x-auto pb-1 sm:pb-0">
              {[
                { id: 'all' as const, label: 'All Flags', count: detectionResult.duplicatePairsCount },
                { id: 'exact' as const, label: 'Exact Matches', count: detectionResult.exactCount },
                { id: 'probable' as const, label: 'Heuristic Flags', count: detectionResult.probableCount },
                { id: 'confirmed' as const, label: 'Confirmed Duplicates', count: detectionResult.confirmedDuplicateCount },
                { id: 'valid' as const, label: 'Marked Valid', count: detectionResult.markedValidCount },
                { id: 'register' as const, label: 'All Invoices', count: detectionResult.totalRecords },
              ].map((tab) => (
                <button
                  key={tab.id}
                  onClick={() => setActiveTab(tab.id)}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all cursor-pointer shrink-0 ${
                    activeTab === tab.id
                      ? 'bg-gray-900 text-white shadow-2xs'
                      : 'text-gray-600 hover:bg-gray-100 hover:text-gray-900'
                  }`}
                >
                  {tab.label} ({tab.count})
                </button>
              ))}
            </div>

            <div className="relative w-full sm:w-64">
              <Search className="w-3.5 h-3.5 text-gray-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Search vendor, invoice #, GSTIN..."
                className="w-full pl-9 pr-7 py-1.5 text-xs bg-gray-50 border border-gray-200 rounded-lg focus:outline-none focus:border-gray-900"
              />
              {searchQuery && (
                <button
                  onClick={() => setSearchQuery('')}
                  className="absolute right-2.5 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600"
                >
                  <X className="w-3 h-3" />
                </button>
              )}
            </div>
          </div>

          {/* Tab View: Full Invoices Register */}
          {activeTab === 'register' ? (
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs text-gray-800">
                <thead className="bg-[#F8F9FA] border-b border-gray-200 text-gray-600 font-semibold uppercase text-[10px] tracking-wider">
                  <tr>
                    <th className="px-4 py-3">Row #</th>
                    <th className="px-4 py-3">Vendor Name</th>
                    <th className="px-4 py-3">GSTIN</th>
                    <th className="px-4 py-3">Invoice Number</th>
                    <th className="px-4 py-3">Invoice Date</th>
                    <th className="px-4 py-3 text-right">Taxable (₹)</th>
                    <th className="px-4 py-3 text-right">Tax (₹)</th>
                    <th className="px-4 py-3 text-right">Total (₹)</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {detectionResult.records.map((r) => (
                    <tr key={r.id} className="hover:bg-gray-50/80 transition-colors">
                      <td className="px-4 py-2.5 font-mono text-gray-500">{r.rowIndex}</td>
                      <td className="px-4 py-2.5 font-medium text-gray-900">{r.vendorName}</td>
                      <td className="px-4 py-2.5 font-mono text-gray-600">{r.gstin || 'N/A'}</td>
                      <td className="px-4 py-2.5 font-mono font-semibold text-gray-900">{r.invoiceNumber}</td>
                      <td className="px-4 py-2.5">{r.invoiceDate}</td>
                      <td className="px-4 py-2.5 text-right font-mono tabular-nums">{r.taxableAmount.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
                      <td className="px-4 py-2.5 text-right font-mono tabular-nums">{r.taxAmount.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
                      <td className="px-4 py-2.5 text-right font-mono font-bold text-gray-900 tabular-nums">{r.totalAmount.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            /* Side-by-Side Duplicate Pairs View */
            <div className="divide-y divide-gray-200">
              {filteredFlags.map((flag) => {
                const isExact = flag.matchType === 'exact';
                const isConfirmed = flag.status === 'duplicate';
                const isValid = flag.status === 'valid';

                return (
                  <div
                    key={flag.id}
                    className={`p-5 transition-colors ${
                      isConfirmed
                        ? 'bg-red-50/30'
                        : isValid
                        ? 'bg-gray-50/20'
                        : 'hover:bg-gray-50/50'
                    }`}
                  >
                    {/* Flag Header Info */}
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 pb-3 mb-3 border-b border-gray-100">
                      <div className="flex items-center gap-2">
                        <span
                          className={`px-2.5 py-0.5 rounded text-[10px] font-bold tracking-wider uppercase ${
                            isExact
                              ? 'bg-red-600 text-white'
                              : 'bg-amber-100 text-amber-800 border border-amber-200'
                          }`}
                        >
                          {flag.matchType.replace(/_/g, ' ')}
                        </span>
                        <span className="text-xs font-semibold text-gray-500">
                          Confidence: {flag.confidence}%
                        </span>
                        <span className="text-gray-300">•</span>
                        <div className="flex items-center gap-1 text-xs text-gray-600">
                          {flag.reasons.map((r, i) => (
                            <span key={i} className="inline-block bg-gray-100 px-2 py-0.5 rounded text-[11px]">
                              {r}
                            </span>
                          ))}
                        </div>
                      </div>

                      {/* Auditor Review Actions */}
                      <div className="flex items-center gap-2">
                        <span className="text-[11px] text-gray-400 font-medium">Mark:</span>
                        <button
                          onClick={() => handleUpdateStatus(flag.id, 'duplicate')}
                          className={`px-2.5 py-1 rounded text-xs font-semibold border transition-all cursor-pointer ${
                            isConfirmed
                              ? 'bg-red-700 text-white border-red-700 shadow-2xs'
                              : 'bg-white text-red-700 border-red-200 hover:bg-red-50'
                          }`}
                        >
                          Duplicate
                        </button>
                        <button
                          onClick={() => handleUpdateStatus(flag.id, 'valid')}
                          className={`px-2.5 py-1 rounded text-xs font-semibold border transition-all cursor-pointer ${
                            isValid
                              ? 'bg-gray-900 text-white border-gray-900 shadow-2xs'
                              : 'bg-white text-gray-800 border-gray-200 hover:bg-gray-50'
                          }`}
                        >
                          Valid Entry
                        </button>
                        {flag.status !== 'unresolved' && (
                          <button
                            onClick={() => handleUpdateStatus(flag.id, 'unresolved')}
                            className="p-1 text-gray-400 hover:text-gray-600"
                            title="Reset to Unresolved"
                          >
                            <RefreshCw className="w-3.5 h-3.5" />
                          </button>
                        )}
                      </div>
                    </div>

                    {/* Side-by-Side Comparison Grid */}
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      {/* Record A */}
                      <div className="p-3.5 rounded-lg border border-gray-200 bg-white space-y-2">
                        <div className="flex items-center justify-between text-xs pb-1.5 border-b border-gray-100">
                          <span className="font-semibold text-gray-700">
                            Primary Entry (Row #{flag.recordA.rowIndex})
                          </span>
                          <span className="font-mono text-gray-400 text-[10px]">
                            ID: {flag.recordA.id}
                          </span>
                        </div>
                        <div className="grid grid-cols-2 gap-2 text-xs">
                          <div>
                            <span className="text-[10px] text-gray-400 block uppercase">Vendor</span>
                            <span className="font-semibold text-gray-900 block truncate" title={flag.recordA.vendorName}>
                              {flag.recordA.vendorName}
                            </span>
                            <span className="text-[11px] font-mono text-gray-500">{flag.recordA.gstin || 'No GSTIN'}</span>
                          </div>
                          <div>
                            <span className="text-[10px] text-gray-400 block uppercase">Invoice #</span>
                            <span className="font-mono font-bold text-gray-900 block">
                              {flag.recordA.invoiceNumber}
                            </span>
                            <span className="text-[11px] text-gray-500">{flag.recordA.invoiceDate}</span>
                          </div>
                          <div>
                            <span className="text-[10px] text-gray-400 block uppercase">Taxable / Tax</span>
                            <span className="text-[11px] text-gray-700">
                              ₹{flag.recordA.taxableAmount.toLocaleString('en-IN')} + ₹{flag.recordA.taxAmount.toLocaleString('en-IN')}
                            </span>
                          </div>
                          <div>
                            <span className="text-[10px] text-gray-400 block uppercase">Total Amount</span>
                            <span className="text-sm font-bold text-gray-900 font-mono">
                              ₹{flag.recordA.totalAmount.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                            </span>
                          </div>
                        </div>
                      </div>

                      {/* Record B */}
                      <div className="p-3.5 rounded-lg border border-gray-200 bg-white space-y-2">
                        <div className="flex items-center justify-between text-xs pb-1.5 border-b border-gray-100">
                          <span className="font-semibold text-gray-700">
                            Matching Entry (Row #{flag.recordB.rowIndex})
                          </span>
                          <span className="font-mono text-gray-400 text-[10px]">
                            ID: {flag.recordB.id}
                          </span>
                        </div>
                        <div className="grid grid-cols-2 gap-2 text-xs">
                          <div>
                            <span className="text-[10px] text-gray-400 block uppercase">Vendor</span>
                            <span className="font-semibold text-gray-900 block truncate" title={flag.recordB.vendorName}>
                              {flag.recordB.vendorName}
                            </span>
                            <span className="text-[11px] font-mono text-gray-500">{flag.recordB.gstin || 'No GSTIN'}</span>
                          </div>
                          <div>
                            <span className="text-[10px] text-gray-400 block uppercase">Invoice #</span>
                            <span className="font-mono font-bold text-gray-900 block">
                              {flag.recordB.invoiceNumber}
                            </span>
                            <span className="text-[11px] text-gray-500">{flag.recordB.invoiceDate}</span>
                          </div>
                          <div>
                            <span className="text-[10px] text-gray-400 block uppercase">Taxable / Tax</span>
                            <span className="text-[11px] text-gray-700">
                              ₹{flag.recordB.taxableAmount.toLocaleString('en-IN')} + ₹{flag.recordB.taxAmount.toLocaleString('en-IN')}
                            </span>
                          </div>
                          <div>
                            <span className="text-[10px] text-gray-400 block uppercase">Total Amount</span>
                            <span className="text-sm font-bold text-gray-900 font-mono">
                              ₹{flag.recordB.totalAmount.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                            </span>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                );
              })}

              {filteredFlags.length === 0 && (
                <div className="text-center py-12 text-xs text-gray-500">
                  <CheckCircle2 className="w-8 h-8 text-gray-600 mx-auto mb-2" />
                  <p className="font-semibold text-gray-800">No duplicate invoice pairs found matching current criteria.</p>
                  <p className="mt-1 text-gray-400">All scanned records passed validation.</p>
                </div>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
};