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
  FileCheck,
  Filter,
  Printer,
  Calendar,
} from 'lucide-react';
import { InnerToolBadge } from '../common/InnerToolBadge';
import {
  type PurchaseMatchItem,
  type PurchaseMatchSummary,
  type PurchaseMatchingConfig,
  type PurchaseColumnMapping,
  detectPurchaseColumns,
  matchPurchaseRegisterWithInvoices,
  exportPurchaseMatchingWorkbook,
} from '../../utils/purchaseInvoiceMatchingEngine';

interface PurchaseInvoiceMatchingToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

export const PurchaseInvoiceMatchingTool: React.FC<PurchaseInvoiceMatchingToolProps> = ({
  onBack,
  onShowToast,
}) => {
  const [prRows, setPrRows] = useState<Record<string, any>[]>([]);
  const [invRows, setInvRows] = useState<Record<string, any>[]>([]);
  const [prHeaders, setPrHeaders] = useState<string[]>([]);
  const [invHeaders, setInvHeaders] = useState<string[]>([]);

  const [prFileName, setPrFileName] = useState('');
  const [invFileName, setInvFileName] = useState('');
  const [clientName, setClientName] = useState('');
  const [period, setPeriod] = useState('FY 2026-27 (Q2)');

  const [prMapping, setPrMapping] = useState<PurchaseColumnMapping>({
    supplierCol: '',
    gstinCol: '',
    invoiceNoCol: '',
    dateCol: '',
    taxableCol: '',
    cgstCol: '',
    sgstCol: '',
    igstCol: '',
    totalCol: '',
  });

  const [invMapping, setInvMapping] = useState<PurchaseColumnMapping>({
    supplierCol: '',
    gstinCol: '',
    invoiceNoCol: '',
    dateCol: '',
    taxableCol: '',
    cgstCol: '',
    sgstCol: '',
    igstCol: '',
    totalCol: '',
  });

  const [config, setConfig] = useState<PurchaseMatchingConfig>({
    amountTolerance: 1.0,
    taxTolerance: 0.5,
    dateToleranceDays: 3,
    requireGstinMatch: false,
    fuzzySupplierMatch: true,
  });

  const [activeTab, setActiveTab] = useState<'exceptions' | 'matched' | 'missing_in_pr' | 'all' | 'report'>('exceptions');
  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState('all');
  const [showConfigPanel, setShowConfigPanel] = useState(false);
  const [showMappingPanel, setShowMappingPanel] = useState(false);

  // Manual review status overrides
  const [reviewStatuses, setReviewStatuses] = useState<Record<string, any>>({});

  const prFileInputRef = useRef<HTMLInputElement>(null);
  const invFileInputRef = useRef<HTMLInputElement>(null);

  // Run matching engine
  const summary: PurchaseMatchSummary = useMemo(() => {
    if (prRows.length === 0 && invRows.length === 0) {
      return {
        totalPrRecords: 0,
        totalInvRecords: 0,
        matchedCount: 0,
        mismatchCount: 0,
        missingInInvCount: 0,
        missingInPrCount: 0,
        duplicateCount: 0,
        reviewNeededCount: 0,
        totalPrValue: 0,
        totalInvValue: 0,
        netVariance: 0,
        matchPercentage: 0,
        items: [],
      };
    }

    const res = matchPurchaseRegisterWithInvoices(
      prRows,
      invRows,
      prMapping,
      invMapping,
      config,
      prFileName || 'Purchase_Register.xlsx',
      invFileName || 'Supplier_Invoices.xlsx'
    );

    // Apply manual review status overrides
    res.items.forEach((item: PurchaseMatchItem) => {
      if (reviewStatuses[item.id]) {
        item.reviewStatus = reviewStatuses[item.id];
      }
    });

    return res;
  }, [prRows, invRows, prMapping, invMapping, config, prFileName, invFileName, reviewStatuses]);

  // Handle PR Upload
  const handlePrUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setPrFileName(file.name);
    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json<Record<string, any>>(ws, { defval: '' });

        if (json.length === 0) {
          onShowToast('Empty File', 'Purchase register contains no data rows.', 'error');
          return;
        }

        const headers = Object.keys(json[0]);
        setPrHeaders(headers);
        setPrRows(json);
        setPrMapping(detectPurchaseColumns(headers));
      } catch (err: any) {
        onShowToast('Upload Error', err.message || 'Failed to parse file', 'error');
      }
    };

    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  // Handle Invoice Upload
  const handleInvUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setInvFileName(file.name);
    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json<Record<string, any>>(ws, { defval: '' });

        if (json.length === 0) {
          onShowToast('Empty File', 'Supplier invoices file contains no data rows.', 'error');
          return;
        }

        const headers = Object.keys(json[0]);
        setInvHeaders(headers);
        setInvRows(json);
        setInvMapping(detectPurchaseColumns(headers));
      } catch (err: any) {
        onShowToast('Upload Error', err.message || 'Failed to parse file', 'error');
      }
    };

    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  // Export 6-sheet Excel report
  const handleExportExcel = () => {
    if (summary.items.length === 0) {
      onShowToast('No Data', 'Please load datasets before exporting.', 'error');
      return;
    }

    try {
      const wb = exportPurchaseMatchingWorkbook(summary, clientName, period);
      XLSX.writeFile(wb, `Purchase_Invoice_Matching_${clientName.replace(/\s+/g, '_')}_${period}.xlsx`);
      onShowToast('Workbook Exported', 'Exported 6-sheet reconciliation workbook (Summary, Matched, Exceptions, Duplicates).', 'success');
    } catch (err: any) {
      onShowToast('Export Error', err.message || 'Failed to generate Excel file.', 'error');
    }
  };

  // Filter items
  const filteredItems = useMemo(() => {
    return summary.items.filter((item) => {
      if (activeTab === 'exceptions' && item.status === 'matched') return false;
      if (activeTab === 'matched' && item.status !== 'matched') return false;
      if (activeTab === 'missing_in_pr' && item.status !== 'missing_in_pr') return false;

      if (statusFilter !== 'all' && item.status !== statusFilter) return false;

      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase();
        const inv = item.invoiceNo.toLowerCase();
        const supp = item.supplierName.toLowerCase();
        const gstin = item.gstin.toLowerCase();
        const desc = item.discrepancyDetails.toLowerCase();
        return inv.includes(q) || supp.includes(q) || gstin.includes(q) || desc.includes(q);
      }
      return true;
    });
  }, [summary.items, activeTab, statusFilter, searchQuery]);

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="purchase-invoice-matching"
        title="Purchase vs Invoice Matching"
        description="Reconcile Purchase Register with Supplier Invoices to detect amount, date, and tax variances, and unbooked invoices."
        badge="Audit Grade"
        onBack={onBack}
        actionButtons={
          <div className="flex flex-wrap items-center gap-2">
            

            <button
              onClick={() => setShowConfigPanel(!showConfigPanel)}
              className="px-3 py-1.5 text-xs font-semibold rounded-md bg-white text-gray-900 border border-gray-200 hover:bg-gray-50 transition-colors flex items-center gap-1.5"
            >
              <Sliders className="w-3.5 h-3.5 text-gray-900" />
              Matching Rules
            </button>

            <button
              onClick={() => setShowMappingPanel(!showMappingPanel)}
              className={`px-3 py-1.5 text-xs font-semibold rounded-md border transition-colors flex items-center gap-1.5 ${
                showMappingPanel
                  ? 'bg-gray-900 text-white border-gray-300'
                  : 'bg-white text-gray-900 border-gray-200 hover:bg-gray-50'
              }`}
            >
              Map Columns
            </button>

            <button
              onClick={handleExportExcel}
              disabled={summary.items.length === 0}
              className="px-3.5 py-1.5 text-xs font-semibold rounded-md bg-gray-900 text-white hover:bg-black disabled:opacity-50 transition-colors flex items-center gap-1.5 shadow-sm cursor-pointer"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              Export Excel Report (6 Sheets)
            </button>
          </div>
        }
      />

      {/* Entity & Period Row */}
      <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-2xs flex flex-wrap items-center gap-4 text-xs">
          <div className="flex items-center gap-2">
            <Building className="w-4 h-4 text-gray-600" />
            <span className="font-semibold text-gray-600">Entity:</span>
            <input
              type="text"
              value={clientName}
              onChange={(e) => setClientName(e.target.value)}
              className="px-2 py-1 border border-gray-200 rounded text-xs font-medium focus:border-gray-300 outline-none w-52"
            />
          </div>

          <div className="flex items-center gap-2">
            <Calendar className="w-4 h-4 text-gray-600" />
            <span className="font-semibold text-gray-600">Period:</span>
            <input
              type="text"
              value={period}
              onChange={(e) => setPeriod(e.target.value)}
              className="px-2 py-1 border border-gray-200 rounded text-xs font-medium focus:border-gray-300 outline-none w-36"
            />
          </div>

          <div className="flex items-center gap-2 text-gray-600">
            <span>PR: <strong className="text-gray-900">{prFileName || 'None'}</strong></span>
            <span>•</span>
            <span>Invoices: <strong className="text-gray-900">{invFileName || 'None'}</strong></span>
          </div>
      </div>

      {/* Main Container */}
      <div className="space-y-6">
        {/* Dual Upload Section */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
          {/* Box 1: Purchase Register */}
          <input
            type="file"
            ref={prFileInputRef}
            onChange={handlePrUpload}
            accept=".xlsx,.xls,.csv"
            className="hidden"
            id="pi-match-pr-upload"
          />
          {prRows.length === 0 ? (
            <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-6 sm:p-8 text-center transition-colors shadow-2xs">
              <label htmlFor="pi-match-pr-upload" className="cursor-pointer flex flex-col items-center">
                <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mx-auto mb-3 shadow-2xs">
                  <FileCheck className="w-6 h-6" />
                </div>
                <h3 className="text-sm font-semibold text-gray-900">1. Internal Purchase Register</h3>
                <p className="text-xs text-gray-500 mt-1 max-w-xs mx-auto">
                  Upload your accounting ERP purchase register (.xlsx, .xls, .csv)
                </p>
                <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
                  <UploadCloud className="w-4 h-4" />
                  <span>Choose Purchase Register</span>
                </div>
              </label>
            </div>
          ) : (
            <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-sm flex flex-col justify-between">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-lg bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 shrink-0">
                  <FileCheck className="w-5 h-5" />
                </div>
                <div className="min-w-0 flex-1">
                  <span className="text-xs font-bold text-gray-900 block truncate">{prFileName}</span>
                  <span className="text-[11px] text-gray-600 block mt-0.5">
                    Purchase Register • {prRows.length} vouchers
                  </span>
                </div>
              </div>
              <div className="mt-4 pt-3 border-t border-gray-100 flex items-center justify-between">
                <span className="text-[11px] text-gray-500 font-medium">Ready for matching</span>
                <button
                  onClick={() => prFileInputRef.current?.click()}
                  className="px-3 py-1.5 text-xs text-gray-900 bg-gray-50 hover:bg-[#EAEAEA] rounded-lg border border-gray-200 flex items-center gap-1.5 font-medium transition-colors cursor-pointer"
                >
                  <UploadCloud className="w-3.5 h-3.5 text-gray-600" />
                  <span>Change Register</span>
                </button>
              </div>
            </div>
          )}

          {/* Box 2: Supplier Invoices */}
          <input
            type="file"
            ref={invFileInputRef}
            onChange={handleInvUpload}
            accept=".xlsx,.xls,.csv"
            className="hidden"
            id="pi-match-inv-upload"
          />
          {invRows.length === 0 ? (
            <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-6 sm:p-8 text-center transition-colors shadow-2xs">
              <label htmlFor="pi-match-inv-upload" className="cursor-pointer flex flex-col items-center">
                <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mx-auto mb-3 shadow-2xs">
                  <FileSpreadsheet className="w-6 h-6" />
                </div>
                <h3 className="text-sm font-semibold text-gray-900">2. Supplier Invoices File</h3>
                <p className="text-xs text-gray-500 mt-1 max-w-xs mx-auto">
                  Upload supplier billed invoices list (.xlsx, .xls, .csv)
                </p>
                <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
                  <UploadCloud className="w-4 h-4" />
                  <span>Choose Supplier Invoices</span>
                </div>
              </label>
            </div>
          ) : (
            <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-sm flex flex-col justify-between">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-lg bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 shrink-0">
                  <FileSpreadsheet className="w-5 h-5" />
                </div>
                <div className="min-w-0 flex-1">
                  <span className="text-xs font-bold text-gray-900 block truncate">{invFileName}</span>
                  <span className="text-[11px] text-gray-600 block mt-0.5">
                    Supplier Invoices • {invRows.length} records
                  </span>
                </div>
              </div>
              <div className="mt-4 pt-3 border-t border-gray-100 flex items-center justify-between">
                <span className="text-[11px] text-gray-500 font-medium">Ready for matching</span>
                <button
                  onClick={() => invFileInputRef.current?.click()}
                  className="px-3 py-1.5 text-xs text-gray-900 bg-gray-50 hover:bg-[#EAEAEA] rounded-lg border border-gray-200 flex items-center gap-1.5 font-medium transition-colors cursor-pointer"
                >
                  <UploadCloud className="w-3.5 h-3.5 text-gray-600" />
                  <span>Change Invoices</span>
                </button>
              </div>
            </div>
          )}
        </div>

        {/* Configuration Drawer */}
        {showConfigPanel && (
          <div className="bg-white border border-gray-200 rounded-xl p-5 shadow-sm space-y-4">
            <div className="flex items-center justify-between pb-3 border-b border-gray-200">
              <div>
                <h3 className="text-sm font-bold text-gray-900">Matching Tolerances & Thresholds</h3>
                <p className="text-xs text-gray-600">Configure matching rules for invoice numbers, amounts, and dates</p>
              </div>
              <button onClick={() => setShowConfigPanel(false)} className="text-xs text-gray-600 hover:text-gray-900">Close</button>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs">
              <div>
                <label className="block font-semibold text-gray-600 mb-1">Total Amount Tolerance (₹)</label>
                <input
                  type="number"
                  step="0.1"
                  value={config.amountTolerance}
                  onChange={(e) => setConfig({ ...config, amountTolerance: parseFloat(e.target.value) || 0 })}
                  className="w-full border border-gray-200 rounded p-2 focus:border-gray-300 outline-none"
                />
                <span className="text-[11px] text-gray-600 mt-1 block">Tolerate fractional rounding differences (default ₹1.00)</span>
              </div>

              <div>
                <label className="block font-semibold text-gray-600 mb-1">Invoice Date Tolerance (Days)</label>
                <input
                  type="number"
                  value={config.dateToleranceDays}
                  onChange={(e) => setConfig({ ...config, dateToleranceDays: Math.max(0, parseInt(e.target.value) || 0) })}
                  className="w-full border border-gray-200 rounded p-2 focus:border-gray-300 outline-none"
                />
                <span className="text-[11px] text-gray-600 mt-1 block">Allow transit / booking lag (default 3 days)</span>
              </div>

              <div className="flex items-center gap-2 pt-6">
                <input
                  type="checkbox"
                  id="requireGstin"
                  checked={config.requireGstinMatch}
                  onChange={(e) => setConfig({ ...config, requireGstinMatch: e.target.checked })}
                  className="rounded text-gray-900 focus:ring-gray-900 h-4 w-4"
                />
                <label htmlFor="requireGstin" className="font-semibold text-gray-900 cursor-pointer">
                  Strict GSTIN Match Required
                </label>
              </div>
            </div>
          </div>
        )}

        {/* KPI Cards */}
        {summary.items.length > 0 && (
          <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
              <span className="text-xs font-semibold text-gray-600 uppercase">Matching Success Rate</span>
              <div className="text-xl font-bold text-gray-900 mt-1">
                {summary.matchPercentage.toFixed(1)}%
              </div>
              <div className="text-xs text-gray-600 mt-0.5">
                {summary.matchedCount} of {summary.totalPrRecords} matched
              </div>
            </div>

            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
              <span className="text-xs font-semibold text-gray-600 uppercase">Amount / Tax Variances</span>
              <div className="text-xl font-bold text-[#DC2626] mt-1">
                {summary.mismatchCount}
              </div>
              <div className="text-xs text-gray-600 mt-0.5">Discrepancy in values</div>
            </div>

            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
              <span className="text-xs font-semibold text-gray-600 uppercase">Missing in Invoices</span>
              <div className="text-xl font-bold text-[#D97706] mt-1">
                {summary.missingInInvCount}
              </div>
              <div className="text-xs text-gray-600 mt-0.5">PR bills not in supplier file</div>
            </div>

            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
              <span className="text-xs font-semibold text-gray-600 uppercase">Missing in PR (ITC Risk)</span>
              <div className="text-xl font-bold text-[#7C3AED] mt-1">
                {summary.missingInPrCount}
              </div>
              <div className="text-xs text-gray-600 mt-0.5">Billed by vendor, omitted in PR</div>
            </div>

            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
              <span className="text-xs font-semibold text-gray-600 uppercase">Duplicates Detected</span>
              <div className="text-xl font-bold text-[#DC2626] mt-1">
                {summary.duplicateCount}
              </div>
              <div className="text-xs text-gray-600 mt-0.5">Reused bill numbering</div>
            </div>
          </div>
        )}

        {/* Tab Navigation & Data Table */}
        {summary.items.length > 0 && (
          <div className="space-y-4">
            <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 border-b border-gray-200 pb-1">
              <div className="flex items-center gap-2">
                <button
                  onClick={() => setActiveTab('exceptions')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'exceptions'
                      ? 'border-[#DC2626] text-[#DC2626]'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Discrepancies & Exceptions ({summary.items.filter((i) => i.status !== 'matched').length})
                </button>
                <button
                  onClick={() => setActiveTab('matched')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'matched'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Matched Invoices ({summary.matchedCount})
                </button>
                <button
                  onClick={() => setActiveTab('missing_in_pr')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'missing_in_pr'
                      ? 'border-[#7C3AED] text-[#7C3AED]'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Missing in PR ({summary.missingInPrCount})
                </button>
                <button
                  onClick={() => setActiveTab('all')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'all'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  All Transactions ({summary.items.length})
                </button>
              </div>

              <div className="flex items-center gap-2">
                <div className="relative">
                  <Search className="w-3.5 h-3.5 absolute left-2.5 top-1/2 -translate-y-1/2 text-gray-600" />
                  <input
                    type="text"
                    placeholder="Search invoice # or supplier..."
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    className="pl-8 pr-3 py-1.5 text-xs border border-gray-200 rounded-lg w-52 focus:outline-none focus:border-gray-300"
                  />
                </div>
              </div>
            </div>

            {/* Table */}
            <div className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm">
              <div className="overflow-x-auto">
                <table className="w-full text-xs text-left">
                  <thead className="bg-gray-900 text-white font-semibold">
                    <tr>
                      <th className="p-3 w-10 text-center">#</th>
                      <th className="p-3">Invoice #</th>
                      <th className="p-3">Status</th>
                      <th className="p-3">Supplier Name & GSTIN</th>
                      <th className="p-3 text-right">PR Amount (₹)</th>
                      <th className="p-3 text-right">Invoice Amount (₹)</th>
                      <th className="p-3 text-right">Variance (₹)</th>
                      <th className="p-3">Dates (PR vs Inv)</th>
                      <th className="p-3">Audit Details & Rationale</th>
                      <th className="p-3 text-center">Resolution</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#E5E7EB]">
                    {filteredItems.map((item, idx) => (
                      <tr
                        key={item.id}
                        className={`hover:bg-gray-50 transition-colors ${
                          item.status === 'amount_mismatch'
                            ? 'bg-red-50/40'
                            : item.status === 'missing_in_pr'
                            ? 'bg-purple-50/30'
                            : item.status === 'missing_in_invoices'
                            ? 'bg-amber-50/30'
                            : ''
                        }`}
                      >
                        <td className="p-3 text-center text-gray-600 font-medium">{idx + 1}</td>
                        <td className="p-3 font-mono font-bold text-gray-900">{item.invoiceNo}</td>
                        <td className="p-3">
                          {item.status === 'matched' && (
                            <span className="px-2 py-0.5 rounded font-bold bg-[#F3F4F6] text-gray-900 text-[10px]">
                              MATCHED
                            </span>
                          )}
                          {item.status === 'amount_mismatch' && (
                            <span className="px-2 py-0.5 rounded font-bold bg-[#FEE2E2] text-[#DC2626] text-[10px]">
                              AMOUNT DIFF
                            </span>
                          )}
                          {item.status === 'tax_mismatch' && (
                            <span className="px-2 py-0.5 rounded font-bold bg-[#FEF3C7] text-[#D97706] text-[10px]">
                              TAX DIFF
                            </span>
                          )}
                          {item.status === 'date_mismatch' && (
                            <span className="px-2 py-0.5 rounded font-bold bg-[#FEF3C7] text-[#D97706] text-[10px]">
                              DATE DIFF
                            </span>
                          )}
                          {item.status === 'missing_in_invoices' && (
                            <span className="px-2 py-0.5 rounded font-bold bg-[#FEF3C7] text-[#B45309] text-[10px]">
                              NOT IN INVOICES
                            </span>
                          )}
                          {item.status === 'missing_in_pr' && (
                            <span className="px-2 py-0.5 rounded font-bold bg-[#EDE9FE] text-[#6D28D9] text-[10px]">
                              NOT IN PR
                            </span>
                          )}
                          {item.status === 'duplicate_invoice' && (
                            <span className="px-2 py-0.5 rounded font-bold bg-[#FEE2E2] text-[#B91C1C] text-[10px]">
                              DUPLICATE
                            </span>
                          )}
                        </td>
                        <td className="p-3 font-medium text-gray-900">
                          <div>{item.supplierName}</div>
                          <div className="font-mono text-[10px] text-gray-600">{item.gstin || 'GSTIN Not Available'}</div>
                        </td>
                        <td className="p-3 text-right font-mono">
                          {item.prTotal > 0 ? `₹${item.prTotal.toLocaleString('en-IN')}` : '-'}
                        </td>
                        <td className="p-3 text-right font-mono">
                          {item.invTotal > 0 ? `₹${item.invTotal.toLocaleString('en-IN')}` : '-'}
                        </td>
                        <td className="p-3 text-right font-bold font-mono">
                          {item.totalVariance > 0 ? (
                            <span className="text-[#DC2626]">₹{item.totalVariance.toLocaleString('en-IN')}</span>
                          ) : (
                            <span className="text-gray-900">₹0.00</span>
                          )}
                        </td>
                        <td className="p-3 text-[11px] text-gray-600">
                          <div>PR: {item.prDate || '-'}</div>
                          <div>Inv: {item.invDate || '-'}</div>
                        </td>
                        <td className="p-3 text-[11px] text-gray-900 max-w-xs">
                          {item.discrepancyDetails}
                        </td>
                        <td className="p-3 text-center">
                          {item.status !== 'matched' ? (
                            <select
                              value={item.reviewStatus}
                              onChange={(e) =>
                                setReviewStatuses((prev) => ({ ...prev, [item.id]: e.target.value }))
                              }
                              className="text-[11px] border border-gray-200 rounded px-1.5 py-0.5 outline-none font-medium"
                            >
                              <option value="unresolved">Unresolved</option>
                              <option value="approved_match">Approve Match</option>
                              <option value="rejected_discrepancy">Discrepancy</option>
                              <option value="query_sent">Query to Vendor</option>
                            </select>
                          ) : (
                            <span className="text-gray-900 font-bold text-xs">✓ Matched</span>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};