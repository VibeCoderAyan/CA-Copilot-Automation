import { ToolHeader } from '../common/ToolHeader';
import React, { useState, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  UploadCloud,
  Download,
  AlertTriangle,
  AlertCircle,
  CheckCircle2,
  FileSpreadsheet,
  Search,
  Check,
  X,
  Sliders,
  ShieldCheck,
  HelpCircle,
  Copy,
  Building,
  RefreshCw,
} from 'lucide-react';
import { InnerToolBadge } from '../common/InnerToolBadge';
import {
  type DeprMethod,
  type FarAssetRecord,
  type FarIssue,
  type FarAuditSummary,
  type FarColumnMapping,
  type FarCheckerConfig,
  SCHEDULE_II_GUIDE,
  detectFarColumns,
  auditFixedAssetRegister,
  exportFarAuditWorkbook,
} from '../../utils/farCheckerEngine';

interface FarCheckerToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

export const FarCheckerTool: React.FC<FarCheckerToolProps> = ({
  onBack,
  onShowToast,
}) => {
  const [rawRows, setRawRows] = useState<Record<string, any>[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [mapping, setMapping] = useState<FarColumnMapping>({
    assetIdCol: '',
    descCol: '',
    categoryCol: '',
    dateCol: '',
    costCol: '',
    accumDeprCol: '',
    openingWdvCol: '',
    bookDeprCol: '',
    closingWdvCol: '',
    usefulLifeCol: '',
    methodCol: '',
  });
  const [fileName, setFileName] = useState('');
  const [showConfigPanel, setShowConfigPanel] = useState(false);
  const [showMappingPanel, setShowMappingPanel] = useState(false);
  const [activeTab, setActiveTab] = useState<'all_assets' | 'issues_only' | 'recomputation' | 'schedule_guide'>('all_assets');
  const [searchQuery, setSearchQuery] = useState('');

  // Configurable parameters
  const [config, setConfig] = useState<FarCheckerConfig>({
    deprMethod: 'companies_act_slm',
    residualValuePercent: 5.0,
    fyStartYear: 2026,
    enforce180DayRule: true,
  });

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Run FAR Audit Engine
  const { summary, assets } = useMemo(() => {
    if (rawRows.length === 0) {
      return {
        summary: {
          totalAssets: 0,
          totalGrossCost: 0,
          totalOpeningAccumDepr: 0,
          totalBookDepreciation: 0,
          totalRecomputedDepreciation: 0,
          netDepreciationVariance: 0,
          totalClosingBookWdv: 0,
          criticalIssuesCount: 0,
          warningIssuesCount: 0,
          duplicateIdsCount: 0,
          depreciationDiscrepanciesCount: 0,
          issues: [],
        },
        assets: [],
      };
    }
    return auditFixedAssetRegister(rawRows, mapping, config);
  }, [rawRows, mapping, config]);

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
          onShowToast('Empty File', 'The uploaded asset register has no data rows.', 'error');
          return;
        }

        const detectedHeaders = Object.keys(json[0]);
        setHeaders(detectedHeaders);
        const detectedMapping = detectFarColumns(detectedHeaders);
        setMapping(detectedMapping);
        setRawRows(json);
      } catch (err: any) {
        onShowToast('Import Error', err.message || 'Could not parse fixed asset register.', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  
  const handleExportReviewReport = () => {
    if (assets.length === 0) return;
    try {
      const wb = exportFarAuditWorkbook(summary, assets, 'Client Entity', 'FY 2026-27');
      XLSX.writeFile(wb, `FAR_Audit_Review_Report_${new Date().toISOString().split('T')[0]}.xlsx`);
      onShowToast('Review Report Exported', 'Saved 4-sheet audit findings, recomputation schedule, and approved FAR copy.', 'success');
    } catch (err: any) {
      onShowToast('Export Failed', err.message || 'Could not export Excel workbook.', 'error');
    }
  };

  const filteredAssets = useMemo(() => {
    return assets.filter((a) => {
      if (activeTab === 'issues_only' && a.issues.length === 0) return false;

      if (!searchQuery) return true;
      const q = searchQuery.toLowerCase();
      return (
        a.assetId.toLowerCase().includes(q) ||
        a.assetDescription.toLowerCase().includes(q) ||
        a.category.toLowerCase().includes(q)
      );
    });
  }, [assets, activeTab, searchQuery]);

  const formatInr = (val: number) => {
    return val.toLocaleString('en-IN', { style: 'currency', currency: 'INR', maximumFractionDigits: 2 });
  };

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="far-checker"
        title="Fixed Asset Register (FAR) Auditor"
        description="Verify asset IDs, capitalization dates, Schedule II useful lives, 5% residual values, and recompute depreciation."
        badge="Statutory"
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
              <span>Depreciation Rules</span>
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
              onClick={handleExportReviewReport}
              disabled={assets.length === 0}
              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-semibold bg-gray-900 text-white hover:bg-black disabled:opacity-40 disabled:cursor-not-allowed shadow-xs transition-all cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Export Audit Report & Corrected Copy</span>
            </button>
          </div>
        }
      />

      {/* Configuration Panel */}
      {showConfigPanel && (
          <div className="mt-4 pt-4 border-t border-gray-100 bg-[#F9FAFA] -mx-5 -mb-5 p-5 rounded-b-xl space-y-4">
            <span className="text-xs font-bold text-gray-900 block">
              Depreciation & Useful Life Parameters
            </span>
            <div className="grid grid-cols-1 sm:grid-cols-4 gap-4 text-xs">
              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Depreciation Method</label>
                <select
                  value={config.deprMethod}
                  onChange={(e) => setConfig({ ...config, deprMethod: e.target.value as DeprMethod })}
                  className="w-full px-2.5 py-1.5 bg-white border border-gray-200 rounded text-xs font-semibold"
                >
                  <option value="companies_act_slm">Companies Act 2013 (Schedule II - SLM)</option>
                  <option value="companies_act_wdv">Companies Act 2013 (Schedule II - WDV)</option>
                  <option value="it_act_block">Income Tax Act 1961 (Sec 32 Block Rates)</option>
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Residual Value % (Schedule II)</label>
                <select
                  value={config.residualValuePercent}
                  onChange={(e) => setConfig({ ...config, residualValuePercent: Number(e.target.value) })}
                  className="w-full px-2.5 py-1.5 bg-white border border-gray-200 rounded text-xs font-semibold"
                >
                  <option value={5.0}>5.0% (Mandatory Schedule II Default)</option>
                  <option value={2.0}>2.0% (Specialized Machinery)</option>
                  <option value={0.0}>0.0% (Full Depr / Nil Scrap Value)</option>
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Financial Year</label>
                <select
                  value={config.fyStartYear}
                  onChange={(e) => setConfig({ ...config, fyStartYear: Number(e.target.value) })}
                  className="w-full px-2.5 py-1.5 bg-white border border-gray-200 rounded text-xs font-semibold"
                >
                  <option value={2030}>FY 2030-31</option>
                  <option value={2029}>FY 2029-30</option>
                  <option value={2028}>FY 2028-29</option>
                  <option value={2027}>FY 2027-28</option>
                  <option value={2026}>FY 2026-27 (Current)</option>
                  <option value={2025}>FY 2025-26</option>
                  <option value={2024}>FY 2024-25</option>
                  <option value={2023}>FY 2023-24</option>
                  <option value={2022}>FY 2022-23</option>
                </select>
              </div>

              <div className="flex items-center pt-5">
                <label className="flex items-center gap-2 cursor-pointer text-xs text-gray-800">
                  <input
                    type="checkbox"
                    checked={config.enforce180DayRule}
                    onChange={(e) => setConfig({ ...config, enforce180DayRule: e.target.checked })}
                    className="rounded text-purple-600"
                  />
                  <span>Enforce &lt; 180 Days Half-Rate Rule</span>
                </label>
              </div>
            </div>
          </div>
        )}

        {/* Column Mapping Panel */}
        {showMappingPanel && headers.length > 0 && (
          <div className="mt-4 pt-4 border-t border-gray-100 bg-white p-4 rounded-lg border border-gray-200 space-y-3">
            <span className="text-xs font-bold text-gray-900 block">
              FAR Column Mapping
            </span>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3 text-xs">
              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Asset ID / Tag</label>
                <select
                  value={mapping.assetIdCol}
                  onChange={(e) => setMapping({ ...mapping, assetIdCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs"
                >
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Description</label>
                <select
                  value={mapping.descCol}
                  onChange={(e) => setMapping({ ...mapping, descCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs"
                >
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Capitalization Date</label>
                <select
                  value={mapping.dateCol}
                  onChange={(e) => setMapping({ ...mapping, dateCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs"
                >
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Historical Cost</label>
                <select
                  value={mapping.costCol}
                  onChange={(e) => setMapping({ ...mapping, costCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs"
                >
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Book Depreciation</label>
                <select
                  value={mapping.bookDeprCol}
                  onChange={(e) => setMapping({ ...mapping, bookDeprCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs"
                >
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Useful Life (Years)</label>
                <select
                  value={mapping.usefulLifeCol}
                  onChange={(e) => setMapping({ ...mapping, usefulLifeCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs"
                >
                  <option value="">-- Auto-detect from Class --</option>
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
            id="far-file-upload"
          />
          <label htmlFor="far-file-upload" className="cursor-pointer flex flex-col items-center">
            <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mx-auto mb-3 shadow-2xs">
              <UploadCloud className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-semibold text-gray-900">
              Import Client Fixed Asset Register (FAR)
            </h3>
            <p className="text-xs text-gray-500 mt-1 max-w-md mx-auto leading-relaxed">
              Supports Excel (.xlsx, .xls) and CSV schedules containing asset descriptions, capitalization dates, gross costs, and depreciation claimed.
            </p>
            <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
              <UploadCloud className="w-4 h-4" />
              <span>Select Asset Register</span>
            </div>
            <span className="text-[11px] text-gray-400 font-medium mt-3 block">
              Schedule II Companies Act 2013 Verification · Zero-modification safety guarantee
            </span>
          </label>
        </div>
      ) : (
        /* 3. Executive Scrutiny KPI Metric Cards */
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
          <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs">
            <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider block">
              Total Assets Audited
            </span>
            <div className="flex items-baseline justify-between mt-1.5">
              <span className="text-xl font-bold text-gray-900">
                {summary.totalAssets}
              </span>
              <span className="text-[11px] text-gray-500">
                Gross: {formatInr(summary.totalGrossCost)}
              </span>
            </div>
          </div>

          <div className={`p-4 rounded-xl border shadow-xs ${
            summary.netDepreciationVariance <= 50
              ? 'bg-white border-gray-200'
              : 'bg-red-50/30 border-red-200'
          }`}>
            <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider block">
              Depreciation Variance
            </span>
            <div className="flex items-baseline justify-between mt-1.5">
              <span className={`text-xl font-bold ${summary.netDepreciationVariance <= 50 ? 'text-gray-800' : 'text-red-700'}`}>
                {summary.netDepreciationVariance <= 50 ? 'TIED OUT' : formatInr(summary.netDepreciationVariance)}
              </span>
              <span className="text-[11px] text-gray-500">
                {summary.depreciationDiscrepanciesCount} Discrepancies
              </span>
            </div>
          </div>

          <div className="bg-white p-4 rounded-xl border border-red-200 shadow-xs bg-red-50/20">
            <span className="text-[11px] font-semibold text-red-700 uppercase tracking-wider block">
              Critical Exceptions
            </span>
            <div className="flex items-baseline justify-between mt-1.5">
              <span className="text-xl font-bold text-red-700">
                {summary.criticalIssuesCount}
              </span>
              <span className="text-[11px] text-red-600 font-semibold">
                {summary.duplicateIdsCount} Duplicates • {summary.issues.filter(i => i.category === 'missing_date').length} Missing Dates
              </span>
            </div>
          </div>

          <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs">
            <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider block">
              Audited Net Book Value
            </span>
            <div className="flex items-baseline justify-between mt-1.5">
              <span className="text-xl font-bold text-gray-900">
                {formatInr(summary.totalClosingBookWdv)}
              </span>
              <span className="text-[11px] text-purple-700 font-semibold">
                5% Residual Retained
              </span>
            </div>
          </div>
        </div>
      )}

      {/* 4. Active Workbench */}
      {assets.length > 0 && (
        <div className="bg-white rounded-xl border border-gray-200 shadow-xs overflow-hidden">
          {/* Tabs and Search Header */}
          <div className="p-4 border-b border-gray-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div className="flex items-center gap-1.5 overflow-x-auto pb-1 sm:pb-0">
              {[
                { id: 'all_assets' as const, label: 'All Fixed Assets', count: assets.length },
                { id: 'issues_only' as const, label: 'Audit Exceptions', count: summary.issues.length },
                { id: 'recomputation' as const, label: 'Depreciation Recomputation', count: assets.length },
                { id: 'schedule_guide' as const, label: 'Schedule II Useful Life Guide', count: Object.keys(SCHEDULE_II_GUIDE).length },
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
                placeholder="Search tag #, description, class..."
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

          {/* Tab 1: All Fixed Assets Register View */}
          {activeTab === 'all_assets' && (
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs text-gray-800">
                <thead className="bg-[#F8F9FA] border-b border-gray-200 text-gray-600 font-semibold uppercase text-[10px] tracking-wider">
                  <tr>
                    <th className="px-4 py-3">Asset ID</th>
                    <th className="px-4 py-3">Description</th>
                    <th className="px-4 py-3">Category</th>
                    <th className="px-4 py-3">Put-to-Use</th>
                    <th className="px-4 py-3 text-right">Gross Cost (₹)</th>
                    <th className="px-4 py-3 text-right">Book Depr (₹)</th>
                    <th className="px-4 py-3 text-right">Audited Depr (₹)</th>
                    <th className="px-4 py-3 text-right">Closing WDV (₹)</th>
                    <th className="px-4 py-3">Audit Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {filteredAssets.map((a) => (
                    <tr key={a.id} className="hover:bg-gray-50/80 transition-colors">
                      <td className="px-4 py-2.5 font-mono font-bold text-gray-900">{a.assetId}</td>
                      <td className="px-4 py-2.5 font-medium text-gray-900 max-w-[200px] truncate" title={a.assetDescription}>
                        {a.assetDescription}
                      </td>
                      <td className="px-4 py-2.5 text-gray-600">{a.category}</td>
                      <td className="px-4 py-2.5 text-gray-600">{a.capitalizationDate || <span className="text-red-600 font-bold">Missing</span>}</td>
                      <td className="px-4 py-2.5 text-right font-mono tabular-nums">{a.grossCost.toLocaleString('en-IN')}</td>
                      <td className="px-4 py-2.5 text-right font-mono tabular-nums">{a.bookDepreciation.toLocaleString('en-IN')}</td>
                      <td className="px-4 py-2.5 text-right font-mono font-bold text-purple-800 tabular-nums">{a.recomputedDepreciation.toLocaleString('en-IN')}</td>
                      <td className="px-4 py-2.5 text-right font-mono tabular-nums font-semibold text-gray-900">{a.closingBookWdv.toLocaleString('en-IN')}</td>
                      <td className="px-4 py-2.5">
                        {a.issues.length === 0 ? (
                          <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-gray-800 bg-gray-50 px-2 py-0.5 rounded">
                            <Check className="w-3 h-3" /> Verified Clean
                          </span>
                        ) : (
                          <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-red-700 bg-red-50 px-2 py-0.5 rounded">
                            <AlertTriangle className="w-3 h-3" /> {a.issues.length} Flag{a.issues.length > 1 ? 's' : ''}
                          </span>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {/* Tab 2: Audit Exceptions Only */}
          {activeTab === 'issues_only' && (
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs text-gray-800">
                <thead className="bg-[#F8F9FA] border-b border-gray-200 text-gray-600 font-semibold uppercase text-[10px] tracking-wider">
                  <tr>
                    <th className="px-4 py-3">Severity</th>
                    <th className="px-4 py-3">Asset ID</th>
                    <th className="px-4 py-3">Category</th>
                    <th className="px-4 py-3">Audit Finding</th>
                    <th className="px-4 py-3">Detailed Observation</th>
                    <th className="px-4 py-3 text-right">Variance (₹)</th>
                    <th className="px-4 py-3">Suggested Action</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {summary.issues.map((iss) => (
                    <tr key={iss.id} className="hover:bg-gray-50/80 transition-colors">
                      <td className="px-4 py-3">
                        <span
                          className={`inline-block px-2 py-0.5 rounded text-[10px] font-bold tracking-wider uppercase ${
                            iss.severity === 'critical'
                              ? 'bg-red-600 text-white'
                              : 'bg-amber-100 text-amber-800 border border-amber-200'
                          }`}
                        >
                          {iss.severity}
                        </span>
                      </td>
                      <td className="px-4 py-3 font-mono font-bold text-gray-900">{iss.assetId}</td>
                      <td className="px-4 py-3 text-gray-600 font-semibold">{iss.category.replace(/_/g, ' ')}</td>
                      <td className="px-4 py-3 font-semibold text-gray-900">{iss.title}</td>
                      <td className="px-4 py-3 text-[11px] text-gray-600 max-w-sm">{iss.description}</td>
                      <td className="px-4 py-3 text-right font-mono font-bold text-red-700">
                        {iss.variance > 0 ? formatInr(iss.variance) : '-'}
                      </td>
                      <td className="px-4 py-3 text-[11px] text-gray-600 max-w-xs">{iss.suggestedAction}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {/* Tab 3: Depreciation Recomputation Schedule */}
          {activeTab === 'recomputation' && (
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs text-gray-800">
                <thead className="bg-[#F8F9FA] border-b border-gray-200 text-gray-600 font-semibold uppercase text-[10px] tracking-wider">
                  <tr>
                    <th className="px-4 py-3">Asset ID</th>
                    <th className="px-4 py-3">Asset Description</th>
                    <th className="px-4 py-3 text-center">Life (Yrs)</th>
                    <th className="px-4 py-3 text-center">Days in FY</th>
                    <th className="px-4 py-3 text-right">Gross Cost (₹)</th>
                    <th className="px-4 py-3 text-right">Book Depr (₹)</th>
                    <th className="px-4 py-3 text-right">Audited Depr (₹)</th>
                    <th className="px-4 py-3 text-right">Variance (₹)</th>
                    <th className="px-4 py-3 text-right">Audited WDV (₹)</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {filteredAssets.map((a) => (
                    <tr key={a.id} className="hover:bg-gray-50/80 transition-colors">
                      <td className="px-4 py-2.5 font-mono font-bold text-gray-900">{a.assetId}</td>
                      <td className="px-4 py-2.5 font-medium text-gray-900 max-w-[200px] truncate">{a.assetDescription}</td>
                      <td className="px-4 py-2.5 text-center font-mono">{a.usefulLifeYears} y</td>
                      <td className="px-4 py-2.5 text-center font-mono">{a.daysUsedInYear} d</td>
                      <td className="px-4 py-2.5 text-right font-mono tabular-nums">{a.grossCost.toLocaleString('en-IN')}</td>
                      <td className="px-4 py-2.5 text-right font-mono tabular-nums">{a.bookDepreciation.toLocaleString('en-IN')}</td>
                      <td className="px-4 py-2.5 text-right font-mono font-bold text-purple-800 tabular-nums">{a.recomputedDepreciation.toLocaleString('en-IN')}</td>
                      <td className={`px-4 py-2.5 text-right font-mono font-bold tabular-nums ${a.depreciationVariance > 50 ? 'text-red-700' : 'text-gray-400'}`}>
                        {a.depreciationVariance > 0 ? (a.bookDepreciation - a.recomputedDepreciation).toLocaleString('en-IN') : '-'}
                      </td>
                      <td className="px-4 py-2.5 text-right font-mono font-bold tabular-nums text-gray-900">{a.recomputedClosingWdv.toLocaleString('en-IN')}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {/* Tab 4: Schedule II Guide */}
          {activeTab === 'schedule_guide' && (
            <div className="p-6 space-y-4">
              <h3 className="text-sm font-bold text-gray-900">
                Companies Act 2013 - Schedule II Useful Life Reference Chart
              </h3>
              <p className="text-xs text-gray-500">
                Statutory useful lives and Income Tax block rates for standard asset categories:
              </p>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3 text-xs">
                {Object.values(SCHEDULE_II_GUIDE).map((guide) => (
                  <div key={guide.category} className="p-3.5 rounded-lg border border-gray-200 bg-gray-50 flex items-center justify-between">
                    <div>
                      <span className="font-semibold text-gray-900 block">{guide.category}</span>
                      <span className="text-[11px] text-gray-500">Companies Act Schedule II</span>
                    </div>
                    <div className="text-right">
                      <span className="font-bold text-purple-700 text-sm block font-mono">{guide.defaultUsefulLifeYears} Years</span>
                      <span className="text-[11px] text-gray-500">IT Act: {guide.itBlockRatePercent}% Block</span>
                    </div>
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