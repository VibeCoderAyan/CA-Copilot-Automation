import { ToolHeader } from '../common/ToolHeader';
import React, { useState, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  UploadCloud,
  Download,
  Printer,
  Calendar,
  Clock,
  CheckCircle2,
  AlertTriangle,
  Sliders,
  Search,
  Check,
  X,
  FileSpreadsheet,
  Building,
  Info,
  ShieldCheck,
} from 'lucide-react';
import { InnerToolBadge } from '../common/InnerToolBadge';
import {
  type LedgerType,
  type AgingBucketConfig,
  type AgingColumnMapping,
  type AgingCalculationResult,
  DEFAULT_STANDARD_BUCKETS,
  SCHEDULE_III_BUCKETS,
  detectAgingColumns,
  calculateAgingAnalysis,
  exportAgingWorkbook,
} from '../../utils/receivablesAgingEngine';

interface ReceivablesAgingToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

export const ReceivablesAgingTool: React.FC<ReceivablesAgingToolProps> = ({
  onBack,
  onShowToast,
}) => {
  const [rawRows, setRawRows] = useState<Record<string, any>[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [mapping, setMapping] = useState<AgingColumnMapping>({
    partyCol: '',
    gstinCol: '',
    invoiceNoCol: '',
    dateCol: '',
    dueDateCol: '',
    grossAmountCol: '',
    paidAmountCol: '',
    balanceCol: '',
  });
  const [fileName, setFileName] = useState('');
  const [ledgerType, setLedgerType] = useState<LedgerType>('receivables');
  const [asOfDateStr, setAsOfDateStr] = useState<string>('2027-03-31');
  const [creditPeriodDays, setCreditPeriodDays] = useState<number>(30);
  const [bucketProfile, setBucketProfile] = useState<'standard' | 'schedule_iii'>('standard');
  const [showConfigPanel, setShowConfigPanel] = useState(false);
  const [showMappingPanel, setShowMappingPanel] = useState(false);
  const [activeTab, setActiveTab] = useState<'party_matrix' | 'open_invoices' | 'summary_view' | 'assumptions'>('party_matrix');
  const [searchQuery, setSearchQuery] = useState('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  const activeBuckets = useMemo(() => {
    return bucketProfile === 'schedule_iii' ? SCHEDULE_III_BUCKETS : DEFAULT_STANDARD_BUCKETS;
  }, [bucketProfile]);

  // Compute aging results
  const agingResult = useMemo<AgingCalculationResult | null>(() => {
    if (rawRows.length === 0) return null;
    return calculateAgingAnalysis(rawRows, mapping, {
      ledgerType,
      asOfDateStr,
      creditPeriodDays,
      buckets: activeBuckets,
      allocateCreditsFifo: true,
    });
  }, [rawRows, mapping, ledgerType, asOfDateStr, creditPeriodDays, activeBuckets]);

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
          onShowToast('Empty File', 'Uploaded spreadsheet contains no data rows.', 'error');
          return;
        }

        const detectedHeaders = Object.keys(json[0]);
        setHeaders(detectedHeaders);
        const detectedMapping = detectAgingColumns(detectedHeaders);
        setMapping(detectedMapping);
        setRawRows(json);
      } catch (err: any) {
        onShowToast('Parsing Error', err.message || 'Failed to read ledger spreadsheet.', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  

  const handleResetLedger = () => {
    setRawRows([]);
    setHeaders([]);
    setFileName('');
    setShowMappingPanel(false);
    onShowToast('Ledger Cleared', 'You can upload another ledger or re-run analysis.', 'info');
  };

  const handleExportExcel = () => {
    if (!agingResult) return;
    try {
      const wb = exportAgingWorkbook(agingResult, activeBuckets, 'Client Entity', 'FY 2026-27');
      XLSX.writeFile(wb, `${ledgerType === 'receivables' ? 'Debtors' : 'Creditors'}_Aging_Report_${asOfDateStr}.xlsx`);
      onShowToast('Report Exported', 'Downloaded multi-sheet aging matrix and statutory disclosure workbook.', 'success');
    } catch (err: any) {
      onShowToast('Export Failed', err.message || 'Could not export Excel workbook.', 'error');
    }
  };

  const handlePrintSummary = () => {
    window.print();
  };

  const filteredParties = useMemo(() => {
    if (!agingResult) return [];
    if (!searchQuery) return agingResult.parties;
    const q = searchQuery.toLowerCase();
    return agingResult.parties.filter(
      (p) =>
        p.partyName.toLowerCase().includes(q) ||
        p.gstin.toLowerCase().includes(q)
    );
  }, [agingResult, searchQuery]);

  const filteredTransactions = useMemo(() => {
    if (!agingResult) return [];
    if (!searchQuery) return agingResult.transactions;
    const q = searchQuery.toLowerCase();
    return agingResult.transactions.filter(
      (t) =>
        t.partyName.toLowerCase().includes(q) ||
        t.invoiceNo.toLowerCase().includes(q) ||
        t.gstin.toLowerCase().includes(q)
    );
  }, [agingResult, searchQuery]);

  const formatInr = (val: number) => {
    return val.toLocaleString('en-IN', { style: 'currency', currency: 'INR', maximumFractionDigits: 2 });
  };

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="receivables-aging"
        title="Receivables & Payables Aging"
        description="Classify outstanding customer & vendor balances into aging buckets with due dates, MSMED interest alerts, and reconciliation checks."
        badge="Schedule III"
        onBack={onBack}
        actionButtons={
          <div className="flex items-center gap-2 flex-wrap">
            {rawRows.length > 0 && (
              <button
                onClick={handleResetLedger}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-white text-rose-600 border border-rose-200 hover:bg-rose-50 transition-all cursor-pointer"
              >
                <X className="w-3.5 h-3.5 text-rose-500" />
                <span>Clear / Re-upload</span>
              </button>
            )}

            <button
              onClick={() => setShowConfigPanel(!showConfigPanel)}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold border transition-all cursor-pointer ${
                showConfigPanel
                  ? 'bg-gray-900 text-white border-gray-900'
                  : 'bg-white text-gray-700 border-gray-200 hover:bg-gray-50'
              }`}
            >
              <Sliders className="w-3.5 h-3.5" />
              <span>Aging Parameters</span>
            </button>

            {rawRows.length > 0 && (
              <button
                onClick={() => setShowMappingPanel(!showMappingPanel)}
                className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold border transition-all cursor-pointer ${
                  showMappingPanel
                    ? 'bg-blue-50 text-blue-800 border-blue-300'
                    : 'bg-white text-gray-700 border-gray-200 hover:bg-gray-50'
                }`}
              >
                <span>Column Mapping</span>
              </button>
            )}

            <button
              onClick={handlePrintSummary}
              disabled={!agingResult}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-white text-gray-700 border border-gray-200 hover:bg-gray-50 disabled:opacity-40 disabled:cursor-not-allowed shadow-xs transition-all cursor-pointer"
              title="Print formatted executive aging summary"
            >
              <Printer className="w-3.5 h-3.5 text-gray-600" />
              <span>Print PDF Summary</span>
            </button>

            <button
              onClick={handleExportExcel}
              disabled={!agingResult}
              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-semibold bg-gray-900 text-white hover:bg-black disabled:opacity-40 disabled:cursor-not-allowed shadow-xs transition-all cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Export Excel</span>
            </button>
          </div>
        }
      />

        {/* Configurable Aging Parameters Panel */}
        {showConfigPanel && (
          <div className="mt-4 pt-4 border-t border-gray-100 bg-[#F9FAFA] -mx-5 -mb-5 p-5 rounded-b-xl space-y-4">
            <span className="text-xs font-bold text-gray-900 block">
              Aging Classification & Cut-off Parameters
            </span>
            <div className="grid grid-cols-1 sm:grid-cols-4 gap-4 text-xs">
              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Ledger Type</label>
                <select
                  value={ledgerType}
                  onChange={(e) => setLedgerType(e.target.value as LedgerType)}
                  className="w-full px-2.5 py-1.5 bg-white border border-gray-200 rounded text-xs font-semibold"
                >
                  <option value="receivables">Trade Receivables (Sundry Debtors)</option>
                  <option value="payables">Trade Payables (Sundry Creditors)</option>
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Aging "As Of" Date</label>
                <input
                  type="date"
                  value={asOfDateStr}
                  onChange={(e) => setAsOfDateStr(e.target.value)}
                  className="w-full px-2.5 py-1.5 bg-white border border-gray-200 rounded text-xs"
                />
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Credit Terms (Days)</label>
                <select
                  value={creditPeriodDays}
                  onChange={(e) => setCreditPeriodDays(Number(e.target.value))}
                  className="w-full px-2.5 py-1.5 bg-white border border-gray-200 rounded text-xs font-semibold"
                >
                  <option value={0}>0 Days (From Invoice Date)</option>
                  <option value={15}>15 Days Credit</option>
                  <option value={30}>30 Days Credit (Standard)</option>
                  <option value={45}>45 Days (MSMED Statutory Limit)</option>
                  <option value={60}>60 Days Credit</option>
                  <option value={90}>90 Days Credit</option>
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Aging Buckets Format</label>
                <select
                  value={bucketProfile}
                  onChange={(e) => setBucketProfile(e.target.value as any)}
                  className="w-full px-2.5 py-1.5 bg-white border border-gray-200 rounded text-xs font-semibold"
                >
                  <option value="standard">Standard (0-30, 31-60, 61-90, 91-180, 180+)</option>
                  <option value="schedule_iii">Schedule III (&lt;6m, 6m-1y, 1-2y, 2-3y, &gt;3y)</option>
                </select>
              </div>
            </div>
          </div>
        )}

        {/* Column Mapping Panel */}
        {showMappingPanel && headers.length > 0 && (
          <div className="mt-4 pt-4 border-t border-gray-100 bg-white p-4 rounded-lg border border-gray-200 space-y-3">
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-gray-900 block">
                Ledger Column Mapping
              </span>
              <span className="text-[11px] text-gray-500">
                Verified auto-detection based on intelligent column aliases
              </span>
            </div>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3 text-xs">
              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Party / Debtor Name *</label>
                <select
                  value={mapping.partyCol}
                  onChange={(e) => setMapping({ ...mapping, partyCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs font-medium"
                >
                  <option value="">-- Select Party Column --</option>
                  {headers.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">GSTIN (Optional)</label>
                <select
                  value={mapping.gstinCol}
                  onChange={(e) => setMapping({ ...mapping, gstinCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs"
                >
                  <option value="">-- None / Not Mapped --</option>
                  {headers.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Invoice / Ref #</label>
                <select
                  value={mapping.invoiceNoCol}
                  onChange={(e) => setMapping({ ...mapping, invoiceNoCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs"
                >
                  <option value="">-- Select Invoice # --</option>
                  {headers.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Invoice Date *</label>
                <select
                  value={mapping.dateCol}
                  onChange={(e) => setMapping({ ...mapping, dateCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs font-medium"
                >
                  <option value="">-- Select Invoice Date --</option>
                  {headers.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Due Date (Optional)</label>
                <select
                  value={mapping.dueDateCol}
                  onChange={(e) => setMapping({ ...mapping, dueDateCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs"
                >
                  <option value="">-- None (Derive from terms) --</option>
                  {headers.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Gross / Billed Amount</label>
                <select
                  value={mapping.grossAmountCol}
                  onChange={(e) => setMapping({ ...mapping, grossAmountCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs"
                >
                  <option value="">-- Select Gross Amount --</option>
                  {headers.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Paid / Credit Amount</label>
                <select
                  value={mapping.paidAmountCol}
                  onChange={(e) => setMapping({ ...mapping, paidAmountCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs"
                >
                  <option value="">-- Select Paid Amount --</option>
                  {headers.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Outstanding Balance *</label>
                <select
                  value={mapping.balanceCol}
                  onChange={(e) => setMapping({ ...mapping, balanceCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs font-medium"
                >
                  <option value="">-- Select Balance Col --</option>
                  {headers.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>
            </div>
          </div>
        )}

      {/* 2. File Upload Dropzone */}
      {rawRows.length === 0 ? (
        <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-8 sm:p-10 text-center space-y-4 shadow-2xs transition-colors">
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileUpload}
            accept=".xlsx,.xls,.csv"
            className="hidden"
            id="receivables-aging-upload"
          />
          <label htmlFor="receivables-aging-upload" className="cursor-pointer flex flex-col items-center">
            <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mx-auto mb-3 shadow-2xs">
              <UploadCloud className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-semibold text-gray-900">
              Import Customer or Vendor Open Balances Ledger
            </h3>
            <p className="text-xs text-gray-500 mt-1 max-w-md mx-auto leading-relaxed">
              Upload customer receivables (sundry debtors) or vendor payables spreadsheets containing transaction dates, invoice references, and outstanding amounts.
            </p>
            <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
              <UploadCloud className="w-4 h-4" />
              <span>Select Ledger File</span>
            </div>
            <span className="text-[11px] text-gray-400 font-medium mt-3 block">
              Schedule III MCA Aging Buckets Supported · Zero variance balance reconciliation guarantee
            </span>
          </label>
        </div>
      ) : (
        /* 3. Executive Aging Dashboard & Bucket Breakdown Cards */
        agingResult && (
          <div className="space-y-4">
            {/* KPI Cards */}
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
              <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs">
                <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider block">
                  Total Parties Audited
                </span>
                <div className="flex items-baseline justify-between mt-1.5">
                  <span className="text-xl font-bold text-gray-900">
                    {agingResult.totalParties}
                  </span>
                  <span className="text-[11px] text-gray-500">
                    {agingResult.transactions.length} invoices
                  </span>
                </div>
              </div>

              <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs">
                <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider block">
                  Net Outstanding Balance
                </span>
                <div className="flex items-baseline justify-between mt-1.5">
                  <span className="text-xl font-bold text-gray-900">
                    {formatInr(agingResult.totalOutstanding)}
                  </span>
                  <span className="text-[11px] text-gray-800 font-semibold">
                    100% Reconciled
                  </span>
                </div>
              </div>

              <div className="bg-white p-4 rounded-xl border border-amber-200 shadow-xs bg-amber-50/20">
                <span className="text-[11px] font-semibold text-amber-800 uppercase tracking-wider block">
                  Overdue Beyond Terms (&gt; {creditPeriodDays}d)
                </span>
                <div className="flex items-baseline justify-between mt-1.5">
                  <span className="text-xl font-bold text-amber-800">
                    {formatInr(
                      Math.max(
                        0,
                        agingResult.totalOutstanding -
                          (agingResult.bucketTotals[activeBuckets[0].id] || 0)
                      )
                    )}
                  </span>
                  <span className="text-[11px] text-amber-700 font-semibold">
                    {agingResult.totalOutstanding > 0
                      ? `${(
                          ((agingResult.totalOutstanding -
                            (agingResult.bucketTotals[activeBuckets[0].id] || 0)) /
                            agingResult.totalOutstanding) *
                          100
                        ).toFixed(1)}%`
                      : '0%'}
                  </span>
                </div>
              </div>

              <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs">
                <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider block">
                  Cut-off As Of Date
                </span>
                <div className="flex items-baseline justify-between mt-1.5">
                  <span className="text-xl font-bold text-gray-900 font-mono">
                    {agingResult.asOfDate}
                  </span>
                  <span className="text-[11px] text-gray-500">
                    {creditPeriodDays}d Credit
                  </span>
                </div>
              </div>
            </div>

            {/* Visual Bucket Progress Distribution */}
            <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs space-y-2">
              <div className="flex items-center justify-between text-xs pb-1">
                <span className="font-bold text-gray-900">
                  Aging Bucket Proportions ({bucketProfile === 'schedule_iii' ? 'Schedule III MCA' : 'Standard 30-Day Intervals'})
                </span>
                <span className="text-[11px] text-gray-500">
                  Total Outstanding: {formatInr(agingResult.totalOutstanding)}
                </span>
              </div>

              {/* Multi-segment Progress Bar */}
              <div className="h-3 w-full bg-gray-100 rounded-full overflow-hidden flex shadow-inner">
                {activeBuckets.map((b, idx) => {
                  const pct = agingResult.bucketPercentages[b.id] || 0;
                  if (pct <= 0) return null;
                  const colors = [
                    'bg-gray-500',
                    'bg-blue-500',
                    'bg-amber-500',
                    'bg-orange-500',
                    'bg-red-600',
                  ];
                  return (
                    <div
                      key={b.id}
                      style={{ width: `${pct}%` }}
                      className={`${colors[idx % colors.length]} transition-all`}
                      title={`${b.label}: ${formatInr(agingResult.bucketTotals[b.id] || 0)} (${pct.toFixed(1)}%)`}
                    />
                  );
                })}
              </div>

              {/* Bucket Legend Pills */}
              <div className="flex items-center gap-4 pt-1 flex-wrap text-xs">
                {activeBuckets.map((b, idx) => {
                  const dotColors = [
                    'bg-gray-500',
                    'bg-blue-500',
                    'bg-amber-500',
                    'bg-orange-500',
                    'bg-red-600',
                  ];
                  const amt = agingResult.bucketTotals[b.id] || 0;
                  const pct = agingResult.bucketPercentages[b.id] || 0;
                  return (
                    <div key={b.id} className="flex items-center gap-1.5">
                      <span className={`w-2 h-2 rounded-full ${dotColors[idx % dotColors.length]}`} />
                      <span className="font-semibold text-gray-700">{b.label}:</span>
                      <span className="font-mono text-gray-900">{formatInr(amt)}</span>
                      <span className="text-[10px] text-gray-400">({pct.toFixed(1)}%)</span>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        )
      )}

      {/* 4. Active Workbench: Matrix vs Open Invoices */}
      {agingResult && (
        <div className="bg-white rounded-xl border border-gray-200 shadow-xs overflow-hidden">
          {/* Workbench Tabs & Search */}
          <div className="p-4 border-b border-gray-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3 no-print">
            <div className="flex items-center gap-1.5 overflow-x-auto pb-1 sm:pb-0">
              {[
                { id: 'party_matrix' as const, label: 'Party-Wise Aging Matrix', count: agingResult.parties.length },
                { id: 'open_invoices' as const, label: 'Detailed Open Invoices', count: agingResult.transactions.length },
                { id: 'assumptions' as const, label: 'Assumptions & Methodology', count: agingResult.assumptions.length },
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
                placeholder="Filter party, GSTIN, invoice #..."
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

          {/* Printable Report Header (Active only when printing) */}
          <div className="hidden print:block p-6 border-b border-gray-200 space-y-2">
            <div className="flex items-center justify-between">
              <div>
                <h2 className="text-xl font-bold text-gray-900">
                  {ledgerType === 'receivables' ? 'Trade Receivables Aging Schedule' : 'Trade Payables Aging Schedule'}
                </h2>
                <p className="text-xs text-gray-600">
                  Cut-off Date: {agingResult.asOfDate} • Credit Terms: {creditPeriodDays} Days
                </p>
              </div>
              <div className="text-right text-xs">
                <span className="font-bold text-gray-900 block">Client Entity</span>
                <span className="text-gray-500">FY 2026-27 Audit Working Paper</span>
              </div>
            </div>
          </div>

          {/* Tab 1: Party-Wise Matrix */}
          {activeTab === 'party_matrix' && (
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs text-gray-800">
                <thead className="bg-[#F8F9FA] border-b border-gray-200 text-gray-600 font-semibold uppercase text-[10px] tracking-wider">
                  <tr>
                    <th className="px-4 py-3">Party Name</th>
                    <th className="px-4 py-3">GSTIN</th>
                    <th className="px-4 py-3 text-right">Net Balance (₹)</th>
                    {activeBuckets.map((b) => (
                      <th key={b.id} className="px-4 py-3 text-right">
                        {b.label}
                      </th>
                    ))}
                    <th className="px-4 py-3 text-right">Overdue %</th>
                    <th className="px-4 py-3 text-center">Risk Level</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {filteredParties.map((p) => {
                    return (
                      <tr key={p.partyName} className="hover:bg-gray-50/80 transition-colors">
                        <td className="px-4 py-2.5 font-semibold text-gray-900 max-w-[200px] truncate" title={p.partyName}>
                          {p.partyName}
                        </td>
                        <td className="px-4 py-2.5 font-mono text-gray-500 text-[11px]">
                          {p.gstin || 'Unregistered'}
                        </td>
                        <td className={`px-4 py-2.5 text-right font-mono font-bold tabular-nums ${p.netOutstanding < 0 ? 'text-blue-700' : 'text-gray-900'}`}>
                          {p.netOutstanding.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                        </td>
                        {activeBuckets.map((b) => {
                          const val = p.bucketAmounts[b.id] || 0;
                          return (
                            <td key={b.id} className="px-4 py-2.5 text-right font-mono tabular-nums text-gray-700">
                              {val > 0 ? val.toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '-'}
                            </td>
                          );
                        })}
                        <td className="px-4 py-2.5 text-right font-mono text-gray-700">
                          {p.overduePercentage > 0 ? `${p.overduePercentage.toFixed(1)}%` : '0.0%'}
                        </td>
                        <td className="px-4 py-2.5 text-center">
                          {p.riskCategory === 'high_risk' ? (
                            <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-red-50 text-red-700 border border-red-200">
                              High Risk
                            </span>
                          ) : p.riskCategory === 'moderate' ? (
                            <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-amber-50 text-amber-700 border border-amber-200">
                              Moderate
                            </span>
                          ) : p.riskCategory === 'credit_balance' ? (
                            <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-blue-50 text-blue-700 border border-blue-200">
                              Advance / Cr
                            </span>
                          ) : (
                            <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-gray-50 text-gray-800 border border-gray-200">
                              Current
                            </span>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
                <tfoot className="bg-[#F8F9FA] border-t-2 border-gray-300 font-bold text-gray-900">
                  <tr>
                    <td className="px-4 py-3" colSpan={2}>TOTAL OUTSTANDING</td>
                    <td className="px-4 py-3 text-right font-mono tabular-nums">
                      {agingResult.totalOutstanding.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                    </td>
                    {activeBuckets.map((b) => (
                      <td key={b.id} className="px-4 py-3 text-right font-mono tabular-nums">
                        {(agingResult.bucketTotals[b.id] || 0).toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                      </td>
                    ))}
                    <td className="px-4 py-3" colSpan={2}></td>
                  </tr>
                </tfoot>
              </table>
            </div>
          )}

          {/* Tab 2: Detailed Open Invoices */}
          {activeTab === 'open_invoices' && (
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs text-gray-800">
                <thead className="bg-[#F8F9FA] border-b border-gray-200 text-gray-600 font-semibold uppercase text-[10px] tracking-wider">
                  <tr>
                    <th className="px-4 py-3">Party Name</th>
                    <th className="px-4 py-3">Invoice #</th>
                    <th className="px-4 py-3">Date</th>
                    <th className="px-4 py-3">Due Date</th>
                    <th className="px-4 py-3 text-center">Age (Days)</th>
                    <th className="px-4 py-3">Bucket</th>
                    <th className="px-4 py-3 text-right">Gross (₹)</th>
                    <th className="px-4 py-3 text-right">Paid (₹)</th>
                    <th className="px-4 py-3 text-right">Balance (₹)</th>
                    {ledgerType === 'payables' && <th className="px-4 py-3">MSMED Alert</th>}
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {filteredTransactions.map((t) => (
                    <tr key={t.id} className="hover:bg-gray-50/80 transition-colors">
                      <td className="px-4 py-2.5 font-medium text-gray-900 max-w-[180px] truncate">{t.partyName}</td>
                      <td className="px-4 py-2.5 font-mono font-semibold text-gray-900">{t.invoiceNo}</td>
                      <td className="px-4 py-2.5 text-gray-600">{t.invoiceDate}</td>
                      <td className="px-4 py-2.5 text-gray-600">{t.dueDate}</td>
                      <td className="px-4 py-2.5 text-center font-mono font-bold text-gray-800">{t.ageDays} d</td>
                      <td className="px-4 py-2.5">
                        <span className="px-2 py-0.5 rounded text-[10px] font-semibold bg-gray-100 text-gray-700">
                          {t.bucketLabel}
                        </span>
                      </td>
                      <td className="px-4 py-2.5 text-right font-mono tabular-nums">{t.grossAmount.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
                      <td className="px-4 py-2.5 text-right font-mono tabular-nums">{t.paidAmount.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
                      <td className="px-4 py-2.5 text-right font-mono font-bold tabular-nums text-gray-900">{t.outstandingBalance.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
                      {ledgerType === 'payables' && (
                        <td className="px-4 py-2.5">
                          {t.msmedExceeded ? (
                            <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-red-50 text-red-700 border border-red-200">
                              Exceeded 45d
                            </span>
                          ) : (
                            <span className="text-gray-400 text-[11px]">Compliant</span>
                          )}
                        </td>
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {/* Tab 3: Assumptions & Methodology */}
          {activeTab === 'assumptions' && (
            <div className="p-6 space-y-4">
              <h3 className="text-sm font-bold text-gray-900">
                Accounting Assumptions & Audit Disclosure Trail
              </h3>
              <p className="text-xs text-gray-500">
                The following statutory criteria and accounting standards (AS 4, AS 9, Ind AS 109, and Schedule III of Companies Act 2013) were applied:
              </p>
              <div className="space-y-2">
                {agingResult.assumptions.map((a, i) => (
                  <div key={i} className="flex items-start gap-2.5 p-3 rounded-lg bg-gray-50 border border-gray-200 text-xs text-gray-800">
                    <Info className="w-4 h-4 text-blue-600 shrink-0 mt-0.5" />
                    <span>{a}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
};