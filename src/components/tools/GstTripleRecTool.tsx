import { ToolHeader } from '../common/ToolHeader';
import React, { useState, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  FileSpreadsheet,
  Download,
  Search,
  CheckCircle2,
  AlertTriangle,
  AlertOctagon,
  FileText,
  Sliders,
  Check,
  X,
  Scale,
  RefreshCw,
  Info,
  Clock,
  ArrowRight,
  ShieldCheck,
  UploadCloud,
} from 'lucide-react';
import { extractSmartTableFromSheet } from '../../utils/dataSanitizer';
import {
  type GstTripleRecSummary,
  type Monthly3WayRecord,
  detectGstTripleRecColumns,
  reconcileGstTripleData,
  exportGstTripleRecWorkbook,
} from '../../utils/gstTripleRecEngine';

interface GstTripleRecToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

export const GstTripleRecTool: React.FC<GstTripleRecToolProps> = ({
  onBack,
  onShowToast,
}) => {
  const [rawRows, setRawRows] = useState<Record<string, any>[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [fileName, setFileName] = useState('');
  const [clientLegalName, setClientLegalName] = useState('');
  const [clientGstin, setClientGstin] = useState('27AAACP1234F1Z8');
  const [financialYear, setFinancialYear] = useState('FY 2025-26');
  const [drc01bThreshold, setDrc01bThreshold] = useState<number>(25000); // Standard statutory ₹25k
  const [activeTab, setActiveTab] = useState<'monthly_grid' | 'drc01b_notices' | 'tax_heads' | 'gstr9c_bridge'>('monthly_grid');
  const [showMappingPanel, setShowMappingPanel] = useState(false);

  const [mapping, setMapping] = useState({
    monthCol: '',
    booksTurnoverCol: '',
    booksIgstCol: '',
    booksCgstCol: '',
    booksSgstCol: '',
    booksCessCol: '',
    gstr1TurnoverCol: '',
    gstr1IgstCol: '',
    gstr1CgstCol: '',
    gstr1SgstCol: '',
    gstr1CessCol: '',
    gstr3bTurnoverCol: '',
    gstr3bIgstCol: '',
    gstr3bCgstCol: '',
    gstr3bSgstCol: '',
    gstr3bCessCol: '',
  });

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Run 3-Way Tripartite Reconciliation
  const summary: GstTripleRecSummary = useMemo(() => {
    if (rawRows.length === 0) {
      return {
        clientGstin,
        clientLegalName,
        financialYear,
        annualBooksTurnover: 0,
        annualGstr1Turnover: 0,
        annualGstr3bTurnover: 0,
        netTurnoverVariance1Vs3b: 0,
        netTurnoverVarianceBooksVs1: 0,
        annualBooksTotalTax: 0,
        annualGstr1TotalTax: 0,
        annualGstr3bTotalTax: 0,
        netAnnualTaxVariance1Vs3b: 0,
        annualIgstVariance: 0,
        annualCgstVariance: 0,
        annualSgstVariance: 0,
        annualCessVariance: 0,
        criticalDrc01bMonthsCount: 0,
        totalDrc01bTaxShortfall: 0,
        totalSection50Interest: 0,
        reconciliationMatchPercent: 100,
        monthlyBreakdown: [],
        drc01bNoticeItems: [],
      };
    }

    return reconcileGstTripleData(
      rawRows,
      mapping,
      clientGstin,
      clientLegalName,
      financialYear,
      drc01bThreshold
    );
  }, [rawRows, mapping, clientGstin, clientLegalName, financialYear, drc01bThreshold]);

  // Handle Excel upload
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
        const smartTable = extractSmartTableFromSheet(ws);

        if (!smartTable.rows || smartTable.rows.length === 0) {
          onShowToast('Empty File', 'Could not detect rows in uploaded sheet', 'error');
          return;
        }

        setHeaders(smartTable.headers);
        setRawRows(smartTable.rows);

        // Auto-detect columns
        const detected = detectGstTripleRecColumns(smartTable.headers);
        setMapping({
          monthCol: detected.monthCol || smartTable.headers[0] || '',
          booksTurnoverCol: detected.booksTurnoverCol || '',
          booksIgstCol: detected.booksIgstCol || '',
          booksCgstCol: detected.booksCgstCol || '',
          booksSgstCol: detected.booksSgstCol || '',
          booksCessCol: detected.booksCessCol || '',
          gstr1TurnoverCol: detected.gstr1TurnoverCol || '',
          gstr1IgstCol: detected.gstr1IgstCol || '',
          gstr1CgstCol: detected.gstr1CgstCol || '',
          gstr1SgstCol: detected.gstr1SgstCol || '',
          gstr1CessCol: detected.gstr1CessCol || '',
          gstr3bTurnoverCol: detected.gstr3bTurnoverCol || '',
          gstr3bIgstCol: detected.gstr3bIgstCol || '',
          gstr3bCgstCol: detected.gstr3bCgstCol || '',
          gstr3bSgstCol: detected.gstr3bSgstCol || '',
          gstr3bCessCol: detected.gstr3bCessCol || '',
        });
      } catch (err: any) {
        onShowToast('Parsing Error', err?.message || 'Failed to read spreadsheet', 'error');
      }
    };

    reader.readAsArrayBuffer(file);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };


  const handleResetData = () => {
    setRawRows([]);
    setHeaders([]);
    setFileName('');
    onShowToast('Data Cleared', 'Reset all 3-way reconciliation data.', 'info');
  };

  // Export full workbook
  const handleExport = () => {
    if (summary.monthlyBreakdown.length === 0) {
      onShowToast('No Data', 'Please load data before exporting', 'error');
      return;
    }
    exportGstTripleRecWorkbook(summary);
    onShowToast('Export Complete', 'Downloaded 4-sheet 3-Way Reconciliation Excel workbook', 'success');
  };

  return (
    <div className="space-y-6">
      {/* Hidden File Input */}
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        className="hidden"
        id="gst-triple-rec-upload"
        onChange={handleFileUpload}
      />

      <ToolHeader
        toolId="gst-triple-rec"
        title="GSTR-1 vs 3B vs Books"
        description="Tripartite reconciliation of outward supplies and tax liabilities (IGST, CGST, SGST, Cess) to preempt automated DRC-01B notices."
        badge="Notice Risk"
        onBack={onBack}
        actionButtons={
          <div className="flex items-center gap-2.5 flex-wrap">
            {rawRows.length > 0 && (
              <button
                onClick={handleResetData}
                className="px-3 py-1.5 text-xs font-medium text-gray-600 hover:text-gray-900 bg-white hover:bg-gray-50 border border-gray-200 rounded-lg transition-colors cursor-pointer"
              >
                Clear
              </button>
            )}

            {rawRows.length > 0 && (
              <button
                onClick={() => setShowMappingPanel(!showMappingPanel)}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg border border-gray-200 bg-white text-gray-700 hover:bg-gray-50 transition-colors cursor-pointer shadow-2xs"
              >
                <Sliders className="w-3.5 h-3.5 text-gray-500" />
                <span>Map Columns</span>
              </button>
            )}

            <button
              onClick={() => fileInputRef.current?.click()}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg bg-gray-900 text-white hover:bg-black transition-colors cursor-pointer shadow-xs"
            >
              <UploadCloud className="w-3.5 h-3.5" />
              <span>Upload 3-Way Sheet</span>
            </button>

            {rawRows.length > 0 && (
              <button
                onClick={handleExport}
                className="flex items-center gap-1.5 px-3.5 py-1.5 text-xs font-semibold rounded-lg bg-gray-900 text-white hover:bg-black transition-colors cursor-pointer shadow-xs"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Export 3-Way Excel</span>
              </button>
            )}
          </div>
        }
      />

      {/* Column Mapping Drawer */}
      {showMappingPanel && (
        <div className="bg-white border border-gray-200 rounded-xl p-5 shadow-xs space-y-4">
          <div className="flex items-center justify-between pb-3 border-b border-gray-100">
            <div>
              <h3 className="text-sm font-bold text-gray-900">3-Way Column Mapping (Books vs GSTR-1 vs GSTR-3B)</h3>
              <p className="text-xs text-gray-500">Map uploaded columns to respective source registers</p>
            </div>
            <button
              onClick={() => setShowMappingPanel(false)}
              className="text-xs text-gray-500 hover:text-gray-900 font-medium cursor-pointer"
            >
              Close
            </button>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-6 text-xs">
            {/* Books Columns */}
            <div className="space-y-3 bg-gray-50/70 p-3.5 rounded-lg border border-gray-200">
              <h4 className="font-bold text-gray-900 uppercase tracking-wider text-[11px] pb-1 border-b border-gray-200">
                1. Books of Account (Sales)
              </h4>
              <div>
                <label className="block text-gray-600 mb-1 font-semibold">Month / Period *</label>
                <select
                  value={mapping.monthCol}
                  onChange={(e) => setMapping({ ...mapping, monthCol: e.target.value })}
                  className="w-full border border-gray-200 rounded-md p-1.5 outline-none bg-white focus:border-gray-900"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-gray-600 mb-1 font-semibold">Books Turnover *</label>
                <select
                  value={mapping.booksTurnoverCol}
                  onChange={(e) => setMapping({ ...mapping, booksTurnoverCol: e.target.value })}
                  className="w-full border border-gray-200 rounded-md p-1.5 outline-none bg-white focus:border-gray-900"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-gray-600 mb-1 font-semibold">Books IGST</label>
                <select
                  value={mapping.booksIgstCol}
                  onChange={(e) => setMapping({ ...mapping, booksIgstCol: e.target.value })}
                  className="w-full border border-gray-200 rounded-md p-1.5 outline-none bg-white focus:border-gray-900"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-gray-600 mb-1 font-semibold">Books CGST</label>
                <select
                  value={mapping.booksCgstCol}
                  onChange={(e) => setMapping({ ...mapping, booksCgstCol: e.target.value })}
                  className="w-full border border-gray-200 rounded-md p-1.5 outline-none bg-white focus:border-gray-900"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-gray-600 mb-1 font-semibold">Books SGST</label>
                <select
                  value={mapping.booksSgstCol}
                  onChange={(e) => setMapping({ ...mapping, booksSgstCol: e.target.value })}
                  className="w-full border border-gray-200 rounded-md p-1.5 outline-none bg-white focus:border-gray-900"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
            </div>

            {/* GSTR-1 Columns */}
            <div className="space-y-3 bg-gray-50/70 p-3.5 rounded-lg border border-gray-200">
              <h4 className="font-bold text-gray-900 uppercase tracking-wider text-[11px] pb-1 border-b border-gray-200">
                2. GSTR-1 Outward Supplies
              </h4>
              <div>
                <label className="block text-gray-600 mb-1 font-semibold">GSTR-1 Taxable Value *</label>
                <select
                  value={mapping.gstr1TurnoverCol}
                  onChange={(e) => setMapping({ ...mapping, gstr1TurnoverCol: e.target.value })}
                  className="w-full border border-gray-200 rounded-md p-1.5 outline-none bg-white focus:border-gray-900"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-gray-600 mb-1 font-semibold">GSTR-1 IGST</label>
                <select
                  value={mapping.gstr1IgstCol}
                  onChange={(e) => setMapping({ ...mapping, gstr1IgstCol: e.target.value })}
                  className="w-full border border-gray-200 rounded-md p-1.5 outline-none bg-white focus:border-gray-900"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-gray-600 mb-1 font-semibold">GSTR-1 CGST</label>
                <select
                  value={mapping.gstr1CgstCol}
                  onChange={(e) => setMapping({ ...mapping, gstr1CgstCol: e.target.value })}
                  className="w-full border border-gray-200 rounded-md p-1.5 outline-none bg-white focus:border-gray-900"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-gray-600 mb-1 font-semibold">GSTR-1 SGST</label>
                <select
                  value={mapping.gstr1SgstCol}
                  onChange={(e) => setMapping({ ...mapping, gstr1SgstCol: e.target.value })}
                  className="w-full border border-gray-200 rounded-md p-1.5 outline-none bg-white focus:border-gray-900"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
            </div>

            {/* GSTR-3B Columns */}
            <div className="space-y-3 bg-gray-50/70 p-3.5 rounded-lg border border-gray-200">
              <h4 className="font-bold text-gray-900 uppercase tracking-wider text-[11px] pb-1 border-b border-gray-200">
                3. GSTR-3B Summary Return
              </h4>
              <div>
                <label className="block text-gray-600 mb-1 font-semibold">GSTR-3B Table 3.1 Taxable *</label>
                <select
                  value={mapping.gstr3bTurnoverCol}
                  onChange={(e) => setMapping({ ...mapping, gstr3bTurnoverCol: e.target.value })}
                  className="w-full border border-gray-200 rounded-md p-1.5 outline-none bg-white focus:border-gray-900"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-gray-600 mb-1 font-semibold">GSTR-3B IGST Paid</label>
                <select
                  value={mapping.gstr3bIgstCol}
                  onChange={(e) => setMapping({ ...mapping, gstr3bIgstCol: e.target.value })}
                  className="w-full border border-gray-200 rounded-md p-1.5 outline-none bg-white focus:border-gray-900"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-gray-600 mb-1 font-semibold">GSTR-3B CGST Paid</label>
                <select
                  value={mapping.gstr3bCgstCol}
                  onChange={(e) => setMapping({ ...mapping, gstr3bCgstCol: e.target.value })}
                  className="w-full border border-gray-200 rounded-md p-1.5 outline-none bg-white focus:border-gray-900"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-gray-600 mb-1 font-semibold">GSTR-3B SGST Paid</label>
                <select
                  value={mapping.gstr3bSgstCol}
                  onChange={(e) => setMapping({ ...mapping, gstr3bSgstCol: e.target.value })}
                  className="w-full border border-gray-200 rounded-md p-1.5 outline-none bg-white focus:border-gray-900"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Empty State / Dropzone */}
      {rawRows.length === 0 ? (
        <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-8 sm:p-10 text-center transition-colors shadow-2xs">
          <label htmlFor="gst-triple-rec-upload" className="cursor-pointer flex flex-col items-center">
            <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mx-auto mb-3 shadow-2xs">
              <Scale className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-semibold text-gray-900">Upload Tripartite GST Reconciliation Data</h3>
            <p className="text-xs text-gray-500 mt-1 max-w-md mx-auto leading-relaxed">
              Upload your monthly sales register alongside GSTR-1 outward filing and GSTR-3B tax payment returns (.xlsx, .xls, .csv) to identify turnover mismatches and automated DRC-01B triggers.
            </p>
            <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
              <UploadCloud className="w-4 h-4" />
              <span>Upload Consolidated Sheet</span>
            </div>
            <span className="text-[11px] text-gray-400 font-medium mt-3 block">
              Tripartite 3-way reconciliation · Identifies DRC-01B notice risks across CGST, SGST & IGST
            </span>
          </label>
        </div>
      ) : (
        <div className="space-y-6">
          {/* Executive KPI Summary Cards */}
          <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-xs">
              <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">Annual Books Turnover</span>
              <div className="text-xl font-bold text-gray-900 font-mono mt-1">
                ₹{summary.annualBooksTurnover.toLocaleString('en-IN')}
              </div>
              <div className="text-xs text-gray-500 mt-0.5">
                Audited P&L Revenue
              </div>
            </div>

            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-xs">
              <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">GSTR-1 vs 3B Turnover Diff</span>
              <div className={`text-xl font-bold font-mono mt-1 ${summary.netTurnoverVariance1Vs3b !== 0 ? 'text-amber-600' : 'text-gray-900'}`}>
                ₹{summary.netTurnoverVariance1Vs3b.toLocaleString('en-IN')}
              </div>
              <div className="text-xs text-gray-500 mt-0.5">
                GSTR-1: ₹{summary.annualGstr1Turnover.toLocaleString('en-IN')}
              </div>
            </div>

            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-xs">
              <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">Net Tax Shortfall (1 vs 3B)</span>
              <div className={`text-xl font-bold font-mono mt-1 ${summary.netAnnualTaxVariance1Vs3b > 0 ? 'text-rose-600' : 'text-gray-800'}`}>
                ₹{summary.netAnnualTaxVariance1Vs3b.toLocaleString('en-IN')}
              </div>
              <div className="text-xs text-gray-500 mt-0.5">
                GSTR-1 tax minus GSTR-3B paid
              </div>
            </div>

            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-xs">
              <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">DRC-01B Notice Risk</span>
              <div className="text-xl font-bold text-rose-600 font-mono mt-1">
                {summary.criticalDrc01bMonthsCount} Months
              </div>
              <div className="text-xs text-gray-500 mt-0.5">
                Shortfall &gt; ₹{drc01bThreshold.toLocaleString('en-IN')}
              </div>
            </div>

            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-xs">
              <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">Est. Sec 50 Interest</span>
              <div className="text-xl font-bold text-amber-600 font-mono mt-1">
                ₹{summary.totalSection50Interest.toLocaleString('en-IN')}
              </div>
              <div className="text-xs text-gray-500 mt-0.5">
                18% p.a. delayed tax liability
              </div>
            </div>
          </div>

          {/* Sub-Tabs Navigation */}
          <div className="bg-white border border-gray-200 rounded-xl p-1.5 flex items-center justify-between shadow-xs">
            <div className="flex items-center gap-1 overflow-x-auto">
              {[
                { id: 'monthly_grid' as const, label: `12-Month Comparison Grid (${summary.monthlyBreakdown.length})` },
                { id: 'drc01b_notices' as const, label: `DRC-01B Notice Scrutiny (${summary.criticalDrc01bMonthsCount} at risk)` },
                { id: 'tax_heads' as const, label: 'Tax Liabilities (IGST/CGST/SGST)' },
                { id: 'gstr9c_bridge' as const, label: 'GSTR-9C Annual Turnover Bridge' },
              ].map((tab) => (
                <button
                  key={tab.id}
                  onClick={() => setActiveTab(tab.id)}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all cursor-pointer whitespace-nowrap ${
                    activeTab === tab.id
                      ? 'bg-gray-900 text-white shadow-xs'
                      : 'text-gray-600 hover:text-gray-900 hover:bg-gray-100'
                  }`}
                >
                  {tab.label}
                </button>
              ))}
            </div>

            {/* GSTIN / Context Details */}
            <div className="hidden lg:flex items-center gap-3 text-xs text-gray-500 font-medium px-2">
              <span>GSTIN: <strong className="text-gray-900 font-mono">{clientGstin}</strong></span>
              <span>•</span>
              <span>Match Score: <strong className="text-gray-800 font-mono font-bold">{summary.reconciliationMatchPercent}%</strong></span>
            </div>
          </div>

          {/* TAB 1: 12-MONTH MONTHLY COMPARISON GRID */}
          {activeTab === 'monthly_grid' && (
            <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-xs space-y-4">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div>
                  <h3 className="text-base font-bold text-gray-900">
                    Monthly Tripartite Reconciliation Grid (Books vs GSTR-1 vs GSTR-3B)
                  </h3>
                  <p className="text-xs text-gray-500 mt-0.5">
                    Side-by-side comparison of Outward Supplies Turnover and total taxes across all 12 fiscal periods.
                  </p>
                </div>
                <div className="text-xs text-gray-500">
                  Showing all <strong className="text-gray-900">{summary.monthlyBreakdown.length}</strong> fiscal periods
                </div>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-xs text-left">
                  <thead className="bg-gray-50 border-y border-gray-200 text-gray-600 uppercase font-semibold">
                    <tr>
                      <th className="py-2.5 px-3">Tax Period</th>
                      <th className="py-2.5 px-3 text-right">Books Turnover</th>
                      <th className="py-2.5 px-3 text-right">GSTR-1 Turnover</th>
                      <th className="py-2.5 px-3 text-right">GSTR-3B Turnover</th>
                      <th className="py-2.5 px-3 text-right">Turnover Diff (1 vs 3B)</th>
                      <th className="py-2.5 px-3 text-right">GSTR-1 Total Tax</th>
                      <th className="py-2.5 px-3 text-right">GSTR-3B Total Tax</th>
                      <th className="py-2.5 px-3 text-right">Tax Shortfall</th>
                      <th className="py-2.5 px-3 text-center">DRC-01B Status</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100 text-gray-700">
                    {summary.monthlyBreakdown.map((m) => (
                      <tr
                        key={m.month}
                        className={
                          m.drc01bThresholdExceeded
                            ? 'bg-rose-50/20 hover:bg-rose-50/40'
                            : m.drc01bStatus === 'WARNING'
                            ? 'bg-amber-50/20 hover:bg-amber-50/30'
                            : 'hover:bg-gray-50'
                        }
                      >
                        <td className="py-3 px-3 font-semibold text-gray-900">{m.month}</td>
                        <td className="py-3 px-3 text-right font-mono font-medium text-gray-900">
                          ₹{m.booksTurnover.toLocaleString('en-IN')}
                        </td>
                        <td className="py-3 px-3 text-right font-mono font-medium text-gray-900">
                          ₹{m.gstr1Turnover.toLocaleString('en-IN')}
                        </td>
                        <td className="py-3 px-3 text-right font-mono font-medium text-gray-900">
                          ₹{m.gstr3bTurnover.toLocaleString('en-IN')}
                        </td>
                        <td className={`py-3 px-3 text-right font-mono font-bold ${m.turnoverVarianceGstr1Vs3b !== 0 ? 'text-amber-600' : 'text-gray-500'}`}>
                          ₹{m.turnoverVarianceGstr1Vs3b.toLocaleString('en-IN')}
                        </td>
                        <td className="py-3 px-3 text-right font-mono text-gray-900">
                          ₹{m.gstr1TotalTax.toLocaleString('en-IN')}
                        </td>
                        <td className="py-3 px-3 text-right font-mono text-gray-900">
                          ₹{m.gstr3bTotalTax.toLocaleString('en-IN')}
                        </td>
                        <td className={`py-3 px-3 text-right font-mono font-bold ${m.netTaxShortfall1Vs3b > 0 ? 'text-rose-600' : m.netTaxShortfall1Vs3b < 0 ? 'text-gray-800' : 'text-gray-500'}`}>
                          ₹{m.netTaxShortfall1Vs3b.toLocaleString('en-IN')}
                        </td>
                        <td className="py-3 px-3 text-center">
                          {m.drc01bStatus === 'CRITICAL_NOTICE_RISK' ? (
                            <span className="px-2 py-0.5 rounded text-[11px] font-semibold bg-rose-50 text-rose-700 border border-rose-200">
                              DRC-01B Risk
                            </span>
                          ) : m.drc01bStatus === 'WARNING' ? (
                            <span className="px-2 py-0.5 rounded text-[11px] font-semibold bg-amber-50 text-amber-700 border border-amber-200">
                              Minor Diff
                            </span>
                          ) : m.drc01bStatus === 'EXCESS_TAX_PAID' ? (
                            <span className="px-2 py-0.5 rounded text-[11px] font-semibold bg-gray-50 text-gray-800 border border-gray-200">
                              Excess Paid
                            </span>
                          ) : (
                            <span className="px-2 py-0.5 rounded text-[11px] font-semibold bg-gray-100 text-gray-700 border border-gray-200">
                              Matched
                            </span>
                          )}
                        </td>
                      </tr>
                    ))}
                    <tr className="bg-gray-50 font-bold">
                      <td className="py-3 px-3 text-gray-900 uppercase">Annual Totals:</td>
                      <td className="py-3 px-3 text-right font-mono text-gray-900">
                        ₹{summary.annualBooksTurnover.toLocaleString('en-IN')}
                      </td>
                      <td className="py-3 px-3 text-right font-mono text-gray-900">
                        ₹{summary.annualGstr1Turnover.toLocaleString('en-IN')}
                      </td>
                      <td className="py-3 px-3 text-right font-mono text-gray-900">
                        ₹{summary.annualGstr3bTurnover.toLocaleString('en-IN')}
                      </td>
                      <td className="py-3 px-3 text-right font-mono text-amber-600">
                        ₹{summary.netTurnoverVariance1Vs3b.toLocaleString('en-IN')}
                      </td>
                      <td className="py-3 px-3 text-right font-mono text-gray-900">
                        ₹{summary.annualGstr1TotalTax.toLocaleString('en-IN')}
                      </td>
                      <td className="py-3 px-3 text-right font-mono text-gray-900">
                        ₹{summary.annualGstr3bTotalTax.toLocaleString('en-IN')}
                      </td>
                      <td className="py-3 px-3 text-right font-mono text-rose-600">
                        ₹{summary.netAnnualTaxVariance1Vs3b.toLocaleString('en-IN')}
                      </td>
                      <td className="py-3 px-3 text-center text-xs text-gray-500">
                        {summary.criticalDrc01bMonthsCount} Notices
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* TAB 2: DRC-01B NOTICE SCRUTINY & AUTOMATED REPLIES */}
          {activeTab === 'drc01b_notices' && (
            <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-xs space-y-4">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div>
                  <h3 className="text-base font-bold text-gray-900">
                    Rule 88C / Form DRC-01B Notice Scrutiny & Draft Explanations
                  </h3>
                  <p className="text-xs text-gray-500 mt-0.5">
                    Whenever tax liability determined in GSTR-1 exceeds tax paid in GSTR-3B by ₹25,000 or 10%, the GST portal automatically issues DRC-01B. Below are items requiring immediate action.
                  </p>
                </div>
                <div className="text-xs font-semibold bg-rose-50 text-rose-800 border border-rose-200 px-3 py-1.5 rounded-lg">
                  Total Tax at Risk: <strong className="font-mono text-rose-700">₹{summary.totalDrc01bTaxShortfall.toLocaleString('en-IN')}</strong>
                </div>
              </div>

              {summary.drc01bNoticeItems.length === 0 ? (
                <div className="p-8 text-center text-gray-500 border border-dashed border-gray-200 rounded-xl">
                  <CheckCircle2 className="w-8 h-8 text-gray-600 mx-auto mb-2" />
                  <p className="text-sm font-semibold text-gray-800">Zero DRC-01B Notice Risk Detected</p>
                  <p className="text-xs text-gray-500 mt-0.5">No tax period exhibits GSTR-1 liability exceeding GSTR-3B beyond the ₹{drc01bThreshold.toLocaleString('en-IN')} statutory limit.</p>
                </div>
              ) : (
                <div className="space-y-4">
                  {summary.drc01bNoticeItems.map((n, idx) => (
                    <div key={n.month} className="border border-rose-200 rounded-xl p-4 bg-rose-50/10 space-y-3">
                      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 pb-2.5 border-b border-rose-100">
                        <div className="flex items-center gap-2">
                          <span className="w-6 h-6 rounded-full bg-rose-100 text-rose-700 flex items-center justify-center font-bold text-xs">
                            {idx + 1}
                          </span>
                          <span className="font-bold text-gray-900 text-sm">{n.month}</span>
                          <span className="px-2 py-0.5 rounded text-[10px] font-semibold bg-rose-100 text-rose-800 uppercase">
                            DRC-01B Triggered
                          </span>
                        </div>
                        <div className="flex items-center gap-4 text-xs font-mono">
                          <span>GSTR-1 Tax: <strong className="text-gray-900">₹{n.gstr1TotalTax.toLocaleString('en-IN')}</strong></span>
                          <span>GSTR-3B Paid: <strong className="text-gray-900">₹{n.gstr3bTotalTax.toLocaleString('en-IN')}</strong></span>
                          <span>Tax Shortfall: <strong className="text-rose-600 font-bold">₹{n.netTaxShortfall1Vs3b.toLocaleString('en-IN')}</strong></span>
                          <span>Est. Interest: <strong className="text-amber-600 font-bold">₹{n.estimatedInterestSec50.toLocaleString('en-IN')}</strong></span>
                        </div>
                      </div>

                      <div className="text-xs space-y-1">
                        <div className="text-gray-500 font-semibold uppercase tracking-wider text-[10px]">
                          Automated Suggested Reply to GST Officer / Portal:
                        </div>
                        <p className="text-gray-800 bg-white p-3 rounded-lg border border-gray-200 leading-relaxed font-sans">
                          {n.suggestedDepartmentReply}
                        </p>
                      </div>

                      <div className="flex items-center justify-between text-[11px] text-gray-500 pt-1">
                        <div>Section 50(1) interest applicable @ 18% p.a. on net cash liability.</div>
                        <div className="font-medium text-gray-700">Form DRC-03 Payment or Part B reply required within 7 days</div>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          {/* TAB 3: TAX HEAD-WISE RECONCILIATION */}
          {activeTab === 'tax_heads' && (
            <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-xs space-y-4">
              <div>
                <h3 className="text-base font-bold text-gray-900">
                  Tax Head-wise Reconciliation (IGST, CGST, SGST, Cess)
                </h3>
                <p className="text-xs text-gray-500 mt-0.5">
                  Verifies whether tax was mistakenly booked under IGST instead of CGST/SGST (Inter-state vs Intra-state mismatch).
                </p>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-xs text-left">
                  <thead className="bg-gray-50 border-y border-gray-200 text-gray-600 uppercase font-semibold">
                    <tr>
                      <th className="py-2.5 px-3">Month</th>
                      <th className="py-2.5 px-3 text-right">GSTR-1 IGST</th>
                      <th className="py-2.5 px-3 text-right">3B IGST Paid</th>
                      <th className="py-2.5 px-3 text-right">IGST Diff</th>
                      <th className="py-2.5 px-3 text-right">GSTR-1 CGST</th>
                      <th className="py-2.5 px-3 text-right">3B CGST Paid</th>
                      <th className="py-2.5 px-3 text-right">CGST Diff</th>
                      <th className="py-2.5 px-3 text-right">GSTR-1 SGST</th>
                      <th className="py-2.5 px-3 text-right">3B SGST Paid</th>
                      <th className="py-2.5 px-3 text-right">SGST Diff</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100 text-gray-700">
                    {summary.monthlyBreakdown.map((m) => (
                      <tr key={m.month} className="hover:bg-gray-50">
                        <td className="py-3 px-3 font-medium text-gray-900">{m.month}</td>
                        <td className="py-3 px-3 text-right font-mono text-gray-900">₹{m.gstr1Igst.toLocaleString('en-IN')}</td>
                        <td className="py-3 px-3 text-right font-mono text-gray-900">₹{m.gstr3bIgst.toLocaleString('en-IN')}</td>
                        <td className={`py-3 px-3 text-right font-mono font-bold ${m.igstDiff1Vs3b !== 0 ? 'text-amber-600' : 'text-gray-500'}`}>
                          ₹{m.igstDiff1Vs3b.toLocaleString('en-IN')}
                        </td>
                        <td className="py-3 px-3 text-right font-mono text-gray-900">₹{m.gstr1Cgst.toLocaleString('en-IN')}</td>
                        <td className="py-3 px-3 text-right font-mono text-gray-900">₹{m.gstr3bCgst.toLocaleString('en-IN')}</td>
                        <td className={`py-3 px-3 text-right font-mono font-bold ${m.cgstDiff1Vs3b !== 0 ? 'text-amber-600' : 'text-gray-500'}`}>
                          ₹{m.cgstDiff1Vs3b.toLocaleString('en-IN')}
                        </td>
                        <td className="py-3 px-3 text-right font-mono text-gray-900">₹{m.gstr1Sgst.toLocaleString('en-IN')}</td>
                        <td className="py-3 px-3 text-right font-mono text-gray-900">₹{m.gstr3bSgst.toLocaleString('en-IN')}</td>
                        <td className={`py-3 px-3 text-right font-mono font-bold ${m.sgstDiff1Vs3b !== 0 ? 'text-amber-600' : 'text-gray-500'}`}>
                          ₹{m.sgstDiff1Vs3b.toLocaleString('en-IN')}
                        </td>
                      </tr>
                    ))}
                    <tr className="bg-gray-50 font-bold">
                      <td className="py-3 px-3 text-gray-900 uppercase">Annual Variances:</td>
                      <td className="py-3 px-3"></td>
                      <td className="py-3 px-3"></td>
                      <td className="py-3 px-3 text-right font-mono text-amber-600">
                        ₹{summary.annualIgstVariance.toLocaleString('en-IN')}
                      </td>
                      <td className="py-3 px-3"></td>
                      <td className="py-3 px-3"></td>
                      <td className="py-3 px-3 text-right font-mono text-amber-600">
                        ₹{summary.annualCgstVariance.toLocaleString('en-IN')}
                      </td>
                      <td className="py-3 px-3"></td>
                      <td className="py-3 px-3"></td>
                      <td className="py-3 px-3 text-right font-mono text-amber-600">
                        ₹{summary.annualSgstVariance.toLocaleString('en-IN')}
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* TAB 4: GSTR-9C TABLE 5 & 7 BRIDGE */}
          {activeTab === 'gstr9c_bridge' && (
            <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-xs space-y-4">
              <div>
                <h3 className="text-base font-bold text-gray-900">
                  GSTR-9C Annual Audit Reconciliation Walk (Audited Financials vs GST Returns)
                </h3>
                <p className="text-xs text-gray-500 mt-0.5">
                  Statutory reconciliation walk mapping Table 5 (Gross Turnover) and Table 7 (Taxable Turnover) of Form GSTR-9C.
                </p>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-xs text-left">
                  <thead className="bg-gray-50 border-y border-gray-200 text-gray-600 uppercase font-semibold">
                    <tr>
                      <th className="py-2.5 px-4">GSTR-9C Table Reference</th>
                      <th className="py-2.5 px-4">Particulars / Turnover Component</th>
                      <th className="py-2.5 px-4 text-right">Amount (INR)</th>
                      <th className="py-2.5 px-4">Auditor Classification & Remarks</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100 text-gray-700">
                    <tr>
                      <td className="py-3 px-4 font-bold text-gray-900">Table 5A</td>
                      <td className="py-3 px-4 font-medium text-gray-900">
                        Turnover as per Audited Financial Statements / P&L
                      </td>
                      <td className="py-3 px-4 text-right font-mono font-bold text-gray-900">
                        ₹{summary.annualBooksTurnover.toLocaleString('en-IN')}
                      </td>
                      <td className="py-3 px-4 text-gray-500">
                        Base gross revenue from audited books
                      </td>
                    </tr>

                    <tr>
                      <td className="py-3 px-4 font-bold text-gray-900">Table 5Q</td>
                      <td className="py-3 px-4 font-medium text-gray-900">
                        Turnover as declared in Annual Return (GSTR-9 Table 5N)
                      </td>
                      <td className="py-3 px-4 text-right font-mono font-bold text-gray-900">
                        ₹{summary.annualGstr1Turnover.toLocaleString('en-IN')}
                      </td>
                      <td className="py-3 px-4 text-gray-500">
                        Cumulative outward turnover from GSTR-1 returns
                      </td>
                    </tr>

                    <tr className="bg-amber-50/30">
                      <td className="py-3 px-4 font-bold text-gray-900">Table 5R</td>
                      <td className="py-3 px-4 font-bold text-gray-900">
                        Unreconciled Turnover (Table 5A minus Table 5Q)
                      </td>
                      <td className={`py-3 px-4 text-right font-mono font-bold text-base ${summary.netTurnoverVarianceBooksVs1 !== 0 ? 'text-rose-600' : 'text-gray-800'}`}>
                        ₹{summary.netTurnoverVarianceBooksVs1.toLocaleString('en-IN')}
                      </td>
                      <td className="py-3 px-4 text-gray-700 font-medium">
                        {summary.netTurnoverVarianceBooksVs1 === 0
                          ? 'Zero turnover variance across Audited Accounts and GST Returns.'
                          : 'Requires explanation in GSTR-9C Table 6 (timing differences, exempt supplies, or unbooked revenue).'}
                      </td>
                    </tr>

                    <tr>
                      <td className="py-3 px-4 font-bold text-gray-900">Table 7A</td>
                      <td className="py-3 px-4 font-medium text-gray-900">
                        Taxable Turnover as per GSTR-3B Table 3.1
                      </td>
                      <td className="py-3 px-4 text-right font-mono font-bold text-gray-900">
                        ₹{summary.annualGstr3bTurnover.toLocaleString('en-IN')}
                      </td>
                      <td className="py-3 px-4 text-gray-500">
                        Turnover upon which tax was actually discharged via cash/credit ledger
                      </td>
                    </tr>

                    <tr className="bg-gray-50">
                      <td className="py-3 px-4 font-bold text-gray-900">Table 9</td>
                      <td className="py-3 px-4 font-medium text-gray-900">
                        Reconciliation of Rate-wise Tax Liability vs Tax Paid
                      </td>
                      <td className="py-3 px-4 text-right font-mono font-bold text-rose-600">
                        ₹{summary.netAnnualTaxVariance1Vs3b.toLocaleString('en-IN')}
                      </td>
                      <td className="py-3 px-4 text-gray-600">
                        Differential tax payable under DRC-03 with Section 50 interest if liability &gt; paid
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
};