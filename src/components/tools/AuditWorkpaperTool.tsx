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
  Check,
  X,
  Search,
  Sliders,
  ShieldCheck,
  UserCheck,
  BookOpen,
  Info,
  Calendar,
} from 'lucide-react';
import { InnerToolBadge } from '../common/InnerToolBadge';
import { ToolHeader } from '../common/ToolHeader';
import {
  type AuditLeadSchedule,
  type AuditWorkpaperAccount,
  type AuditWorkpaperHeader,
  type AuditWorkpaperPack,
  type WorkpaperColumnMapping,
  detectWorkpaperColumns,
  generateAuditWorkingPaperPack,
  exportAuditWorkpaperWorkbook,
} from '../../utils/auditWorkpaperEngine';

interface AuditWorkpaperToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

export const AuditWorkpaperTool: React.FC<AuditWorkpaperToolProps> = ({
  onBack,
  onShowToast,
}) => {
  const [rawRows, setRawRows] = useState<Record<string, any>[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [mapping, setMapping] = useState<WorkpaperColumnMapping>({
    accountCodeCol: '',
    accountNameCol: '',
    groupCol: '',
    priorYearCol: '',
    openingCol: '',
    debitCol: '',
    creditCol: '',
    closingCol: '',
  });
  const [fileName, setFileName] = useState('');
  const [selectedSchedule, setSelectedSchedule] = useState<string>('ALL');
  const [activeTab, setActiveTab] = useState<'lead_schedules' | 'analytical_review' | 'audit_procedures' | 'confirmations'>('lead_schedules');
  const [searchQuery, setSearchQuery] = useState('');
  const [showMetadataPanel, setShowMetadataPanel] = useState(false);
  const [showMappingPanel, setShowMappingPanel] = useState(false);

  // Audit engagement metadata
  const [header, setHeader] = useState<AuditWorkpaperHeader>({
    clientName: '',
    engagementType: 'Statutory Audit under Companies Act 2013',
    financialYear: 'FY 2026-27',
    asOfDate: '31-03-2027',
    preparedBy: 'CA Articled Clerk / Intern',
    reviewedBy: 'Audit Senior / Partner',
    materialityThresholdAmount: 100000,
    materialityThresholdPercent: 10,
  });

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Generate working paper pack
  const pack = useMemo<AuditWorkpaperPack | null>(() => {
    if (rawRows.length === 0) return null;
    return generateAuditWorkingPaperPack(rawRows, mapping, header);
  }, [rawRows, mapping, header]);

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
        const detectedMapping = detectWorkpaperColumns(detectedHeaders);
        setMapping(detectedMapping);
        setRawRows(json);
      } catch (err: any) {
        onShowToast('Import Error', err.message || 'Could not parse spreadsheet file.', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  
  const handleExportExcel = () => {
    if (!pack) return;
    try {
      const wb = exportAuditWorkpaperWorkbook(pack);
      XLSX.writeFile(wb, `Audit_Working_Papers_${pack.header.clientName.replace(/\s+/g, '_')}_${pack.header.financialYear}.xlsx`);
      onShowToast('Working Papers Exported', 'Saved 5-sheet audit working paper package with lead schedules and analytical procedures.', 'success');
    } catch (err: any) {
      onShowToast('Export Failed', err.message || 'Could not export Excel workbook.', 'error');
    }
  };

  const handlePrint = () => {
    window.print();
  };

  const filteredLeadSchedules = useMemo(() => {
    if (!pack) return [];
    if (selectedSchedule === 'ALL') return pack.leadSchedules;
    return pack.leadSchedules.filter((s) => s.wpRef === selectedSchedule);
  }, [pack, selectedSchedule]);

  const filteredAccounts = useMemo(() => {
    if (!pack) return [];
    let list = pack.allAccounts;
    if (selectedSchedule !== 'ALL') {
      list = list.filter((a) => a.leadGroup === selectedSchedule);
    }
    if (searchQuery) {
      const q = searchQuery.toLowerCase();
      list = list.filter(
        (a) =>
          a.accountName.toLowerCase().includes(q) ||
          a.accountCode.toLowerCase().includes(q) ||
          a.leadGroup.toLowerCase().includes(q)
      );
    }
    return list;
  }, [pack, selectedSchedule, searchQuery]);

  const formatInr = (val: number) => {
    return val.toLocaleString('en-IN', { style: 'currency', currency: 'INR', maximumFractionDigits: 2 });
  };

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="audit-workpapers"
        title="Audit Working Paper Generator"
        description="Generate structured audit lead schedules (WP-A to WP-I), materiality analytical reviews, substantive testing plans, and confirmation trackers."
        badge="Standards on Auditing (SAs)"
        onBack={onBack}
        actionButtons={
          <div className="flex items-center gap-2 flex-wrap">
            <button
              onClick={() => setShowMetadataPanel(!showMetadataPanel)}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold border transition-all cursor-pointer ${
                showMetadataPanel
                  ? 'bg-gray-900 text-white border-gray-900'
                  : 'bg-white text-gray-700 border-gray-200 hover:bg-gray-50'
              }`}
            >
              <UserCheck className="w-3.5 h-3.5" />
              <span>Engagement Info</span>
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
              onClick={handlePrint}
              disabled={!pack}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-white text-gray-700 border border-gray-200 hover:bg-gray-50 disabled:opacity-40 disabled:cursor-not-allowed shadow-xs transition-all cursor-pointer"
              title="Print draft working papers"
            >
              <Printer className="w-3.5 h-3.5 text-gray-600" />
              <span>Print Draft WP</span>
            </button>

            <button
              onClick={handleExportExcel}
              disabled={!pack}
              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-semibold bg-gray-900 text-white hover:bg-black disabled:opacity-40 disabled:cursor-not-allowed shadow-xs transition-all cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Export Excel Package</span>
            </button>
          </div>
        }
      />

      {/* Engagement Metadata Panel */}
      {showMetadataPanel && (
        <div className="bg-white border border-gray-200 p-5 rounded-xl space-y-3 shadow-2xs">
          <span className="text-xs font-bold text-gray-900 block">
            Audit Engagement & Sign-off Details
          </span>
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
            <div>
              <label className="block text-[10px] font-semibold text-gray-600 mb-1">Client Entity Name</label>
              <input
                type="text"
                value={header.clientName}
                onChange={(e) => setHeader({ ...header, clientName: e.target.value })}
                className="w-full px-2.5 py-1.5 bg-white border border-gray-200 rounded text-xs font-semibold"
              />
            </div>

            <div>
              <label className="block text-[10px] font-semibold text-gray-600 mb-1">Financial Year</label>
              <input
                type="text"
                value={header.financialYear}
                onChange={(e) => setHeader({ ...header, financialYear: e.target.value })}
                className="w-full px-2.5 py-1.5 bg-white border border-gray-200 rounded text-xs"
              />
            </div>

            <div>
              <label className="block text-[10px] font-semibold text-gray-600 mb-1">Prepared By (Intern/Staff)</label>
              <input
                type="text"
                value={header.preparedBy}
                onChange={(e) => setHeader({ ...header, preparedBy: e.target.value })}
                className="w-full px-2.5 py-1.5 bg-white border border-gray-200 rounded text-xs"
              />
            </div>

            <div>
              <label className="block text-[10px] font-semibold text-gray-600 mb-1">Reviewed By (Partner)</label>
              <input
                type="text"
                value={header.reviewedBy}
                onChange={(e) => setHeader({ ...header, reviewedBy: e.target.value })}
                className="w-full px-2.5 py-1.5 bg-white border border-gray-200 rounded text-xs"
              />
            </div>

            <div>
              <label className="block text-[10px] font-semibold text-gray-600 mb-1">Materiality Threshold (₹)</label>
              <input
                type="number"
                value={header.materialityThresholdAmount}
                onChange={(e) => setHeader({ ...header, materialityThresholdAmount: Number(e.target.value) })}
                className="w-full px-2.5 py-1.5 bg-white border border-gray-200 rounded text-xs"
              />
            </div>

            <div>
              <label className="block text-[10px] font-semibold text-gray-600 mb-1">Materiality Fluctuation (%)</label>
              <input
                type="number"
                value={header.materialityThresholdPercent}
                onChange={(e) => setHeader({ ...header, materialityThresholdPercent: Number(e.target.value) })}
                className="w-full px-2.5 py-1.5 bg-white border border-gray-200 rounded text-xs"
              />
            </div>
          </div>
        </div>
      )}

        {/* Column Mapping Panel */}
        {showMappingPanel && headers.length > 0 && (
          <div className="mt-4 pt-4 border-t border-gray-100 bg-white p-4 rounded-lg border border-gray-200 space-y-3">
            <span className="text-xs font-bold text-gray-900 block">
              Trial Balance Column Mapping
            </span>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3 text-xs">
              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Account Code</label>
                <select
                  value={mapping.accountCodeCol}
                  onChange={(e) => setMapping({ ...mapping, accountCodeCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs"
                >
                  <option value="">-- Auto-generate --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Account Name</label>
                <select
                  value={mapping.accountNameCol}
                  onChange={(e) => setMapping({ ...mapping, accountNameCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs"
                >
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Prior Year Audited</label>
                <select
                  value={mapping.priorYearCol}
                  onChange={(e) => setMapping({ ...mapping, priorYearCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs"
                >
                  <option value="">-- Optional --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Closing Balance</label>
                <select
                  value={mapping.closingCol}
                  onChange={(e) => setMapping({ ...mapping, closingCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs"
                >
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
            </div>
          </div>
        )}

      {/* 2. File Import Dropzone */}
      {rawRows.length === 0 ? (
        <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-8 sm:p-10 text-center space-y-4 shadow-2xs transition-colors">
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileUpload}
            accept=".xlsx,.xls,.csv"
            className="hidden"
            id="audit-workpaper-upload"
          />
          <label htmlFor="audit-workpaper-upload" className="cursor-pointer flex flex-col items-center">
            <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mx-auto mb-3 shadow-2xs">
              <UploadCloud className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-semibold text-gray-900">
              Import Trial Balance or Ledger Spreadsheet
            </h3>
            <p className="text-xs text-gray-500 mt-1 max-w-md mx-auto leading-relaxed">
              Upload Trial Balance with opening balances, period debits/credits, closing balances, and prior year figures.
            </p>
            <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
              <UploadCloud className="w-4 h-4" />
              <span>Select Trial Balance File</span>
            </div>
            <span className="text-[11px] text-gray-400 font-medium mt-3 block">
              Automated Standard Lead Schedules (WP-A to WP-I) · Zero fabrication of audit evidence guarantee
            </span>
          </label>
        </div>
      ) : (
        /* 3. Executive Lead Schedule KPI Metric Cards */
        pack && (
          <div className="space-y-4">
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
              <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs">
                <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider block">
                  Lead Schedules Generated
                </span>
                <div className="flex items-baseline justify-between mt-1.5">
                  <span className="text-xl font-bold text-gray-900">
                    {pack.leadSchedules.length} Schedules
                  </span>
                  <span className="text-[11px] text-gray-500">
                    {pack.allAccounts.length} heads
                  </span>
                </div>
              </div>

              <div className="bg-white p-4 rounded-xl border border-amber-200 shadow-xs bg-amber-50/20">
                <span className="text-[11px] font-semibold text-amber-800 uppercase tracking-wider block">
                  Material Analytical Fluctuations
                </span>
                <div className="flex items-baseline justify-between mt-1.5">
                  <span className="text-xl font-bold text-amber-800">
                    {pack.analyticalExceptionsCount}
                  </span>
                  <span className="text-[11px] text-amber-700 font-semibold">
                    &gt; {header.materialityThresholdPercent}% & ₹{header.materialityThresholdAmount.toLocaleString('en-IN')}
                  </span>
                </div>
              </div>

              <div className="bg-white p-4 rounded-xl border border-red-200 shadow-xs bg-red-50/20">
                <span className="text-[11px] font-semibold text-red-700 uppercase tracking-wider block">
                  Red Flags / Missing Evidence
                </span>
                <div className="flex items-baseline justify-between mt-1.5">
                  <span className="text-xl font-bold text-red-700">
                    {pack.missingEvidenceCount}
                  </span>
                  <span className="text-[11px] text-red-600 font-semibold">
                    SA 500 Documentation Required
                  </span>
                </div>
              </div>

              <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs">
                <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider block">
                  Trial Balance Status
                </span>
                <div className="flex items-baseline justify-between mt-1.5">
                  <span className="text-xl font-bold text-gray-800">
                    {pack.isTbBalanced ? 'BALANCED' : 'IMBALANCED'}
                  </span>
                  <span className="text-[11px] text-gray-500">
                    {header.financialYear}
                  </span>
                </div>
              </div>
            </div>

            {/* Statutory Disclaimer Alert */}
            <div className="p-3.5 rounded-lg bg-gray-50 border border-gray-200 flex items-start gap-3 text-xs text-gray-600">
              <Info className="w-4 h-4 text-gray-800 shrink-0 mt-0.5" />
              <div>
                <span className="font-semibold text-gray-800 block">
                  Integrity Standard & Professional Ethics Notice
                </span>
                <span>{pack.disclaimer}</span>
              </div>
            </div>
          </div>
        )
      )}

      {/* 4. Active Workbench */}
      {pack && (
        <div className="bg-white rounded-xl border border-gray-200 shadow-xs overflow-hidden">
          {/* Working Paper Tabs & Filters */}
          <div className="p-4 border-b border-gray-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3 no-print">
            <div className="flex items-center gap-1.5 overflow-x-auto pb-1 sm:pb-0">
              {[
                { id: 'lead_schedules' as const, label: 'Lead Schedules (WP-A to I)', count: pack.leadSchedules.length },
                { id: 'analytical_review' as const, label: 'Detailed Accounts & Analytical', count: pack.allAccounts.length },
                { id: 'audit_procedures' as const, label: 'Substantive Testing Plan', count: pack.procedures.length },
                { id: 'confirmations' as const, label: 'External Confirmations', count: pack.confirmations.length },
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

            {/* Lead Schedule Selector Filter */}
            <div className="flex items-center gap-2">
              <select
                value={selectedSchedule}
                onChange={(e) => setSelectedSchedule(e.target.value)}
                className="px-2.5 py-1.5 text-xs bg-gray-50 border border-gray-200 rounded-lg text-gray-800 font-semibold"
              >
                <option value="ALL">All Schedules (WP-A to WP-I)</option>
                {pack.leadSchedules.map((s) => (
                  <option key={s.wpRef} value={s.wpRef}>
                    {s.wpRef}: {s.leadScheduleName}
                  </option>
                ))}
              </select>

              <div className="relative w-44">
                <Search className="w-3.5 h-3.5 text-gray-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
                <input
                  type="text"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  placeholder="Filter heads..."
                  className="w-full pl-9 pr-6 py-1.5 text-xs bg-gray-50 border border-gray-200 rounded-lg focus:outline-none focus:border-gray-900"
                />
              </div>
            </div>
          </div>

          {/* Printable Working Paper Letterhead (Print-only) */}
          <div className="hidden print:block p-6 border-b border-gray-300 space-y-2">
            <div className="flex items-center justify-between pb-3 border-b border-gray-200">
              <div>
                <h1 className="text-xl font-bold text-gray-900">{header.clientName}</h1>
                <p className="text-xs text-gray-600 font-medium">
                  {header.engagementType} • {header.financialYear} (As of {header.asOfDate})
                </p>
              </div>
              <div className="text-right text-xs">
                <span className="font-semibold text-gray-800 block">Prepared: {header.preparedBy}</span>
                <span className="font-semibold text-gray-800 block">Reviewed: {header.reviewedBy}</span>
                <span className="text-[10px] text-gray-500">Draft Subject to Partner Sign-off</span>
              </div>
            </div>
          </div>

          {/* Tab 1: Lead Schedules Master Summary */}
          {activeTab === 'lead_schedules' && (
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs text-gray-800">
                <thead className="bg-[#F8F9FA] border-b border-gray-200 text-gray-600 font-semibold uppercase text-[10px] tracking-wider">
                  <tr>
                    <th className="px-4 py-3">WP Ref</th>
                    <th className="px-4 py-3">Lead Schedule Description</th>
                    <th className="px-4 py-3">Class</th>
                    <th className="px-4 py-3 text-right">Prior Year (₹)</th>
                    <th className="px-4 py-3 text-right">Opening (₹)</th>
                    <th className="px-4 py-3 text-right">Debits (₹)</th>
                    <th className="px-4 py-3 text-right">Credits (₹)</th>
                    <th className="px-4 py-3 text-right">Closing (₹)</th>
                    <th className="px-4 py-3 text-right">Fluctuation %</th>
                    <th className="px-4 py-3 text-center">Audit Risk</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {filteredLeadSchedules.map((s) => (
                    <tr key={s.wpRef} className="hover:bg-gray-50/80 transition-colors">
                      <td className="px-4 py-3 font-mono font-bold text-gray-900">{s.wpRef}</td>
                      <td className="px-4 py-3 font-semibold text-gray-900">{s.leadScheduleName}</td>
                      <td className="px-4 py-3 text-gray-500">{s.groupCategory}</td>
                      <td className="px-4 py-3 text-right font-mono tabular-nums">{s.totalPriorYear.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
                      <td className="px-4 py-3 text-right font-mono tabular-nums">{s.totalOpening.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
                      <td className="px-4 py-3 text-right font-mono tabular-nums text-gray-600">{s.totalDebits > 0 ? s.totalDebits.toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '-'}</td>
                      <td className="px-4 py-3 text-right font-mono tabular-nums text-gray-600">{s.totalCredits > 0 ? s.totalCredits.toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '-'}</td>
                      <td className="px-4 py-3 text-right font-mono font-bold text-gray-900 tabular-nums">{s.totalClosing.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
                      <td className={`px-4 py-3 text-right font-mono font-bold tabular-nums ${s.isMaterialVariance ? 'text-amber-800' : 'text-gray-700'}`}>
                        {s.variancePercent.toFixed(1)}%
                      </td>
                      <td className="px-4 py-3 text-center">
                        {s.auditRisk === 'high' ? (
                          <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-amber-50 text-amber-800 border border-amber-200">
                            High Variance
                          </span>
                        ) : (
                          <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-gray-50 text-gray-800 border border-gray-200">
                            Normal
                          </span>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {/* Tab 2: Detailed Accounts & Analytical Review */}
          {activeTab === 'analytical_review' && (
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs text-gray-800">
                <thead className="bg-[#F8F9FA] border-b border-gray-200 text-gray-600 font-semibold uppercase text-[10px] tracking-wider">
                  <tr>
                    <th className="px-4 py-3">WP Ref</th>
                    <th className="px-4 py-3">A/c Code</th>
                    <th className="px-4 py-3">Account Description</th>
                    <th className="px-4 py-3 text-right">Prior Year (₹)</th>
                    <th className="px-4 py-3 text-right">Debit (₹)</th>
                    <th className="px-4 py-3 text-right">Credit (₹)</th>
                    <th className="px-4 py-3 text-right">Closing (₹)</th>
                    <th className="px-4 py-3 text-right">Variance %</th>
                    <th className="px-4 py-3">Tick Mark</th>
                    <th className="px-4 py-3">Audit Red Flags & Notes</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {filteredAccounts.map((a) => (
                    <tr key={a.id} className="hover:bg-gray-50/80 transition-colors">
                      <td className="px-4 py-2.5 font-mono font-bold text-gray-600">{a.leadGroup}</td>
                      <td className="px-4 py-2.5 font-mono text-gray-500">{a.accountCode}</td>
                      <td className="px-4 py-2.5 font-medium text-gray-900 max-w-[200px] truncate" title={a.accountName}>
                        {a.accountName}
                      </td>
                      <td className="px-4 py-2.5 text-right font-mono tabular-nums">{a.priorYearBalance.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
                      <td className="px-4 py-2.5 text-right font-mono tabular-nums">{a.periodDebit > 0 ? a.periodDebit.toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '-'}</td>
                      <td className="px-4 py-2.5 text-right font-mono tabular-nums">{a.periodCredit > 0 ? a.periodCredit.toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '-'}</td>
                      <td className="px-4 py-2.5 text-right font-mono font-bold text-gray-900 tabular-nums">{a.closingBalance.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
                      <td className={`px-4 py-2.5 text-right font-mono font-bold tabular-nums ${a.isMaterial ? 'text-amber-800' : 'text-gray-600'}`}>
                        {a.variancePercent.toFixed(1)}%
                      </td>
                      <td className="px-4 py-2.5 font-mono text-[11px] text-gray-500">
                        {a.auditTickMark}
                      </td>
                      <td className="px-4 py-2.5 text-[11px]">
                        {a.missingEvidenceFlags.length > 0 ? (
                          <span className="text-red-700 font-semibold bg-red-50 px-2 py-0.5 rounded">
                            {a.missingEvidenceFlags[0]}
                          </span>
                        ) : (
                          <span className="text-gray-400">Clean</span>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {/* Tab 3: Substantive Audit Testing Program */}
          {activeTab === 'audit_procedures' && (
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs text-gray-800">
                <thead className="bg-[#F8F9FA] border-b border-gray-200 text-gray-600 font-semibold uppercase text-[10px] tracking-wider">
                  <tr>
                    <th className="px-4 py-3">WP Ref</th>
                    <th className="px-4 py-3">Audit Area</th>
                    <th className="px-4 py-3">Objective & Substantive Procedure</th>
                    <th className="px-4 py-3">Evidence Required</th>
                    <th className="px-4 py-3">Status</th>
                    <th className="px-4 py-3">Auditor Findings & Conclusion</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {pack.procedures.map((p) => (
                    <tr key={p.wpRef} className="hover:bg-gray-50/80 transition-colors">
                      <td className="px-4 py-3 font-mono font-bold text-gray-900">{p.wpRef}</td>
                      <td className="px-4 py-3 font-bold text-gray-900">{p.procedureArea}</td>
                      <td className="px-4 py-3 max-w-sm">
                        <div className="font-semibold text-gray-900">{p.objective}</div>
                        <div className="text-[11px] text-gray-500 mt-0.5">{p.substantiveProcedure}</div>
                      </td>
                      <td className="px-4 py-3 text-[11px] text-gray-600 max-w-xs">{p.evidenceRequired}</td>
                      <td className="px-4 py-3">
                        <span className="px-2 py-0.5 rounded text-[10px] font-semibold bg-gray-100 text-gray-800">
                          {p.verificationStatus}
                        </span>
                      </td>
                      <td className="px-4 py-3 text-[11px] text-gray-600 max-w-xs">{p.auditFindingNote}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {/* Tab 4: External Balance Confirmations Tracker */}
          {activeTab === 'confirmations' && (
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs text-gray-800">
                <thead className="bg-[#F8F9FA] border-b border-gray-200 text-gray-600 font-semibold uppercase text-[10px] tracking-wider">
                  <tr>
                    <th className="px-4 py-3">Party Name</th>
                    <th className="px-4 py-3">Category</th>
                    <th className="px-4 py-3 text-right">Book Balance (₹)</th>
                    <th className="px-4 py-3 text-right">Confirmed (₹)</th>
                    <th className="px-4 py-3 text-right">Difference (₹)</th>
                    <th className="px-4 py-3">Status</th>
                    <th className="px-4 py-3">Auditor Reconciliation Note</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {pack.confirmations.map((c) => (
                    <tr key={c.id} className="hover:bg-gray-50/80 transition-colors">
                      <td className="px-4 py-2.5 font-bold text-gray-900">{c.partyName}</td>
                      <td className="px-4 py-2.5 text-gray-600">{c.category}</td>
                      <td className="px-4 py-2.5 text-right font-mono tabular-nums">{c.bookBalance.toLocaleString('en-IN')}</td>
                      <td className="px-4 py-2.5 text-right font-mono tabular-nums">{c.confirmedBalance > 0 ? c.confirmedBalance.toLocaleString('en-IN') : '-'}</td>
                      <td className={`px-4 py-2.5 text-right font-mono font-bold tabular-nums ${c.difference > 0 ? 'text-amber-800' : 'text-gray-800'}`}>
                        {c.difference.toLocaleString('en-IN')}
                      </td>
                      <td className="px-4 py-2.5">
                        <span className={`px-2 py-0.5 rounded text-[10px] font-semibold ${
                          c.status === 'Confirmed Matched'
                            ? 'bg-gray-50 text-gray-800 border border-gray-200'
                            : c.status === 'Variance Disputed'
                            ? 'bg-red-50 text-red-700 border border-red-200'
                            : 'bg-amber-50 text-amber-700 border border-amber-200'
                        }`}>
                          {c.status}
                        </span>
                      </td>
                      <td className="px-4 py-2.5 text-[11px] text-gray-600 max-w-xs">{c.remarks}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}
    </div>
  );
};