import { ToolHeader } from '../common/ToolHeader';
import React, { useState, useRef, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  UploadCloud,
  Download,
  CheckCheck,
  AlertTriangle,
  AlertCircle,
  CheckCircle2,
  FileSpreadsheet,
  Sparkles,
  Sliders,
  Filter,
  Search,
  Eye,
  Info
} from 'lucide-react';
import {
  type TbAuditSummary,
  type RawTbRow,
  type TbColumnMapping,
  type TbIssue,
  detectTbColumns,
  auditTrialBalance,
  exportTbAuditWorkbook,
} from '../../utils/tbCheckerEngine';
import { DataFormatGuide } from '../common/DataFormatGuide';

interface TrialBalanceCheckerToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

export const TrialBalanceCheckerTool: React.FC<TrialBalanceCheckerToolProps> = ({
  onBack,
  onShowToast,
}) => {
  const [rawRows, setRawRows] = useState<Record<string, any>[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [mapping, setMapping] = useState<TbColumnMapping>({
    accountNameCol: '',
    accountCodeCol: '',
    groupCol: '',
    debitCol: '',
    creditCol: '',
    balanceCol: '',
  });
  const [fileName, setFileName] = useState('');
  const [activeTab, setActiveTab] = useState<'all_issues' | 'critical' | 'duplicates' | 'round_sums' | 'tb_view'>('all_issues');
  const [searchTerm, setSearchTerm] = useState('');
  const [showMappingPanel, setShowMappingPanel] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Run audit engine whenever raw rows or mapping changes
  const { summary, parsedRows } = useMemo(() => {
    if (rawRows.length === 0) {
      return {
        summary: {
          totalAccounts: 0,
          totalDebit: 0,
          totalCredit: 0,
          netDifference: 0,
          isBalanced: true,
          criticalIssuesCount: 0,
          warningIssuesCount: 0,
          infoIssuesCount: 0,
          issues: [],
        },
        parsedRows: [],
      };
    }
    return auditTrialBalance(rawRows, mapping);
  }, [rawRows, mapping]);

  // Synchronize TB audit figures to active session for Dossier
  useEffect(() => {
    if (parsedRows.length > 0) {
      try {
        sessionStorage.setItem(
          'ca_copilot_session_tb',
          JSON.stringify({
            totalAccounts: summary.totalAccounts,
            totalDebit: summary.totalDebit,
            totalCredit: summary.totalCredit,
            netDifference: summary.netDifference,
            isBalanced: summary.isBalanced,
            criticalIssuesCount: summary.criticalIssuesCount,
            timestamp: new Date().toISOString(),
          })
        );
      } catch {
        // Storage unavailable
      }
    }
  }, [parsedRows.length, summary]);

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
          onShowToast('Empty File', 'The uploaded workbook has no data rows.', 'error');
          return;
        }

        const detectedHeaders = Object.keys(json[0]);
        setHeaders(detectedHeaders);
        const detectedMapping = detectTbColumns(detectedHeaders);
        setMapping(detectedMapping);
        setRawRows(json);
      } catch (err: any) {
        onShowToast('Parsing Error', err.message || 'Failed to read trial balance file.', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  
  const handleExportExcel = () => {
    if (parsedRows.length === 0) return;
    try {
      const wb = exportTbAuditWorkbook(summary, parsedRows, 'Client Entity', 'FY 2026-27');
      XLSX.writeFile(wb, `TB_Audit_Report_${new Date().toISOString().split('T')[0]}.xlsx`);
      onShowToast('Report Exported', 'Downloaded 3-sheet audit findings and standardized trial balance.', 'success');
    } catch (err: any) {
      onShowToast('Export Failed', err.message || 'Could not export Excel report.', 'error');
    }
  };

  const filteredIssues = summary.issues.filter((iss) => {
    if (activeTab === 'critical' && iss.severity !== 'critical') return false;
    if (activeTab === 'duplicates' && iss.category !== 'duplicate_account' && iss.category !== 'missing_code') return false;
    if (activeTab === 'round_sums' && iss.category !== 'round_sum' && iss.category !== 'suspense_account') return false;

    if (!searchTerm) return true;
    const term = searchTerm.toLowerCase();
    return (
      iss.accountName.toLowerCase().includes(term) ||
      iss.accountCode.toLowerCase().includes(term) ||
      iss.title.toLowerCase().includes(term) ||
      iss.description.toLowerCase().includes(term) ||
      iss.suggestedAction.toLowerCase().includes(term)
    );
  });

  const formatInr = (val: number) => {
    return val.toLocaleString('en-IN', { style: 'currency', currency: 'INR', maximumFractionDigits: 2 });
  };

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="tb-checker"
        title="Trial Balance & Ledger Checker"
        description="Automated audit scrutiny for debit-credit imbalances, negative cash, inverted revenue/capital balances, and duplicate codes."
        badge="Scrutiny"
        onBack={onBack}
        actionButtons={
          rawRows.length > 0 && (
            <div className="flex items-center gap-2 flex-wrap">
              <button
                onClick={() => setShowMappingPanel(!showMappingPanel)}
                className="flex items-center gap-1.5 px-3 py-2 bg-white border border-gray-200 hover:border-gray-300 text-gray-900 rounded-lg text-xs font-semibold shadow-sm transition-colors cursor-pointer"
              >
                <Sliders className="w-3.5 h-3.5 text-gray-900" />
                Column Mapping
              </button>
              <button
                onClick={handleExportExcel}
                className="flex items-center gap-1.5 px-4 py-2 bg-gray-900 hover:bg-black text-white rounded-lg text-xs font-semibold shadow-sm transition-colors cursor-pointer"
              >
                <Download className="w-3.5 h-3.5" />
                Export Audit Report (Excel)
              </button>
            </div>
          )
        }
      />

      <DataFormatGuide
        toolType="trial-balance"
        acceptedFormats="Excel (.xlsx, .xls), CSV (.csv)"
        requiredFields={[
          { name: 'Particulars / Account Head', description: 'Ledger account name', aliases: ['particulars', 'account', 'head'] },
          { name: 'Debit / Credit', description: 'Closing debit or credit balance', aliases: ['debit', 'credit', 'amount'] },
        ]}
      />

      {/* Upload Dropzone */}
      {rawRows.length === 0 ? (
        <div className="bg-white border-2 border-dashed border-gray-200 rounded-xl p-8 text-center hover:border-gray-300 transition-colors">
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileUpload}
            accept=".xlsx,.xls,.csv"
            className="hidden"
            id="tb-upload"
          />
          <label htmlFor="tb-upload" className="cursor-pointer flex flex-col items-center">
            <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mb-3">
              <UploadCloud className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-semibold text-gray-900">Upload Trial Balance or General Ledger</h3>
            <p className="text-xs text-gray-600 mt-1 max-w-md">
              Upload client Excel (.xlsx, .xls) or CSV files. All accounting headers and structures are supported.
            </p>
            <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-sm inline-flex items-center gap-2">
              <UploadCloud className="w-4 h-4" />
              <span>Upload Spreadsheet File</span>
            </div>
            <span className="text-[11px] text-gray-400 font-medium mt-3 block">
              Supports Excel (.xlsx, .xls) and CSV (.csv) with multi-column trial balances
            </span>
          </label>
        </div>
      ) : (
        /* Diagnostic Results View */
        <div className="space-y-5">
          {/* Column Mapping Collapsible Panel */}
          {showMappingPanel && (
            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm space-y-3">
              <div className="flex items-center justify-between pb-2 border-b border-gray-200">
                <h3 className="text-xs font-bold text-gray-900 uppercase tracking-wider">
                  Trial Balance Column Mapping
                </h3>
                <button
                  onClick={() => setShowMappingPanel(false)}
                  className="text-xs text-gray-900 font-semibold hover:underline"
                >
                  Done
                </button>
              </div>
              <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3 text-xs">
                <div>
                  <label className="text-[11px] font-semibold text-gray-600 block mb-1">Account Name *</label>
                  <select
                    value={mapping.accountNameCol}
                    onChange={(e) => setMapping({ ...mapping, accountNameCol: e.target.value })}
                    className="w-full px-2 py-1.5 bg-white border border-gray-200 rounded text-xs text-gray-900"
                  >
                    {headers.map((h) => (
                      <option key={h} value={h}>
                        {h}
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="text-[11px] font-semibold text-gray-600 block mb-1">Account Code</label>
                  <select
                    value={mapping.accountCodeCol}
                    onChange={(e) => setMapping({ ...mapping, accountCodeCol: e.target.value })}
                    className="w-full px-2 py-1.5 bg-white border border-gray-200 rounded text-xs text-gray-900"
                  >
                    <option value="">-- None --</option>
                    {headers.map((h) => (
                      <option key={h} value={h}>
                        {h}
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="text-[11px] font-semibold text-gray-600 block mb-1">Group / Type</label>
                  <select
                    value={mapping.groupCol}
                    onChange={(e) => setMapping({ ...mapping, groupCol: e.target.value })}
                    className="w-full px-2 py-1.5 bg-white border border-gray-200 rounded text-xs text-gray-900"
                  >
                    <option value="">-- None --</option>
                    {headers.map((h) => (
                      <option key={h} value={h}>
                        {h}
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="text-[11px] font-semibold text-gray-600 block mb-1">Debit Column</label>
                  <select
                    value={mapping.debitCol}
                    onChange={(e) => setMapping({ ...mapping, debitCol: e.target.value })}
                    className="w-full px-2 py-1.5 bg-white border border-gray-200 rounded text-xs text-gray-900"
                  >
                    {headers.map((h) => (
                      <option key={h} value={h}>
                        {h}
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="text-[11px] font-semibold text-gray-600 block mb-1">Credit Column</label>
                  <select
                    value={mapping.creditCol}
                    onChange={(e) => setMapping({ ...mapping, creditCol: e.target.value })}
                    className="w-full px-2 py-1.5 bg-white border border-gray-200 rounded text-xs text-gray-900"
                  >
                    {headers.map((h) => (
                      <option key={h} value={h}>
                        {h}
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="text-[11px] font-semibold text-gray-600 block mb-1">Net Balance</label>
                  <select
                    value={mapping.balanceCol}
                    onChange={(e) => setMapping({ ...mapping, balanceCol: e.target.value })}
                    className="w-full px-2 py-1.5 bg-white border border-gray-200 rounded text-xs text-gray-900"
                  >
                    <option value="">-- None --</option>
                    {headers.map((h) => (
                      <option key={h} value={h}>
                        {h}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
            </div>
          )}

          {/* Diagnostic Summary Cards */}
          <div className="bg-white border border-gray-200 rounded-xl p-5 shadow-sm space-y-4">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-gray-200">
              <div className="flex items-center gap-2">
                <FileSpreadsheet className="w-4 h-4 text-gray-900" />
                <span className="text-xs font-bold text-gray-900">{fileName}</span>
                <span className="text-[10px] text-gray-600">({parsedRows.length} Ledger Accounts)</span>
              </div>
              <div className="flex items-center gap-2">
                {summary.isBalanced ? (
                  <span className="inline-flex items-center gap-1.5 text-xs font-bold bg-gray-100 text-gray-900 px-3 py-1 rounded-full border border-gray-200">
                    <CheckCircle2 className="w-3.5 h-3.5" />
                    Trial Balance Balanced (Zero Difference)
                  </span>
                ) : (
                  <span className="inline-flex items-center gap-1.5 text-xs font-bold bg-red-50 text-red-700 px-3 py-1 rounded-full border border-red-200">
                    <AlertTriangle className="w-3.5 h-3.5" />
                    Debit-Credit Imbalance: {formatInr(summary.netDifference)}
                  </span>
                )}
                <button
                  onClick={() => {
                    setRawRows([]);
                    setFileName('');
                  }}
                  className="text-xs text-gray-600 hover:text-[#DC2626] ml-2"
                >
                  Change File
                </button>
              </div>
            </div>

            {/* 4 Cards */}
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
              <div className="p-3 bg-gray-50 border border-gray-200 rounded-lg">
                <span className="text-[11px] font-semibold text-gray-600 block">Total Debits</span>
                <span className="text-base font-bold text-gray-900 mt-0.5 block">{formatInr(summary.totalDebit)}</span>
                <span className="text-[10px] text-gray-600">{summary.totalAccounts} Ledgers</span>
              </div>
              <div className="p-3 bg-gray-50 border border-gray-200 rounded-lg">
                <span className="text-[11px] font-semibold text-gray-600 block">Total Credits</span>
                <span className="text-base font-bold text-gray-900 mt-0.5 block">{formatInr(summary.totalCredit)}</span>
                <span className="text-[10px] text-gray-600">{summary.totalAccounts} Ledgers</span>
              </div>
              <div className="p-3 bg-red-50 border border-red-200 rounded-lg">
                <span className="text-[11px] font-semibold text-red-800 flex items-center gap-1">
                  <AlertCircle className="w-3 h-3" /> Critical Red Flags
                </span>
                <span className="text-base font-bold text-red-900 mt-0.5 block">{summary.criticalIssuesCount}</span>
                <span className="text-[10px] text-red-700">Immediate CA action</span>
              </div>
              <div className="p-3 bg-amber-50 border border-amber-200 rounded-lg">
                <span className="text-[11px] font-semibold text-amber-800 flex items-center gap-1">
                  <AlertTriangle className="w-3 h-3" /> Warnings & COA Flags
                </span>
                <span className="text-base font-bold text-amber-900 mt-0.5 block">
                  {summary.warningIssuesCount + summary.infoIssuesCount}
                </span>
                <span className="text-[10px] text-amber-700">Audit classification</span>
              </div>
            </div>
          </div>

          {/* Navigation Tabs & Search */}
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-white border border-gray-200 rounded-xl p-3 shadow-sm">
            <div className="flex flex-wrap items-center gap-1.5">
              <button
                onClick={() => setActiveTab('all_issues')}
                className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors ${
                  activeTab === 'all_issues'
                    ? 'bg-gray-900 text-white shadow-sm'
                    : 'bg-transparent text-gray-600 hover:bg-gray-50'
                }`}
              >
                All Findings ({summary.issues.length})
              </button>
              <button
                onClick={() => setActiveTab('critical')}
                className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors ${
                  activeTab === 'critical'
                    ? 'bg-red-700 text-white shadow-sm'
                    : 'bg-transparent text-red-800 hover:bg-red-50'
                }`}
              >
                Critical ({summary.criticalIssuesCount})
              </button>
              <button
                onClick={() => setActiveTab('duplicates')}
                className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors ${
                  activeTab === 'duplicates'
                    ? 'bg-amber-600 text-white shadow-sm'
                    : 'bg-transparent text-amber-800 hover:bg-amber-50'
                }`}
              >
                Duplicate & Missing Codes
              </button>
              <button
                onClick={() => setActiveTab('round_sums')}
                className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors ${
                  activeTab === 'round_sums'
                    ? 'bg-blue-700 text-white shadow-sm'
                    : 'bg-transparent text-blue-800 hover:bg-blue-50'
                }`}
              >
                Round Sums & Suspense
              </button>
              <button
                onClick={() => setActiveTab('tb_view')}
                className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors ${
                  activeTab === 'tb_view'
                    ? 'bg-[#111827] text-white shadow-sm'
                    : 'bg-transparent text-gray-900 hover:bg-gray-50'
                }`}
              >
                Cleaned TB View ({parsedRows.length})
              </button>
            </div>

            <div className="relative w-full sm:w-64">
              <Search className="w-3.5 h-3.5 text-gray-600 absolute left-3 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                placeholder="Search issues, accounts..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="w-full pl-8 pr-3 py-1.5 bg-white border border-gray-200 rounded-lg text-xs text-gray-900 focus:outline-none focus:border-gray-300"
              />
            </div>
          </div>

          {/* View Mode: Issues or Cleaned TB */}
          {activeTab !== 'tb_view' ? (
            /* Issues Table */
            <div className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm">
              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="bg-gray-50 border-b border-gray-200 text-gray-600 uppercase font-bold text-[10px] tracking-wider">
                    <tr>
                      <th className="py-2.5 px-3 w-12 text-center">Row</th>
                      <th className="py-2.5 px-3 w-28">Severity</th>
                      <th className="py-2.5 px-4">Account / Ledger Head</th>
                      <th className="py-2.5 px-4">Audit Finding & Description</th>
                      <th className="py-2.5 px-4">Recommended CA Action</th>
                      <th className="py-2.5 px-3 text-right">Debit / Credit</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#E5E7EB]">
                    {filteredIssues.map((iss) => (
                      <tr key={iss.id} className="hover:bg-gray-50 transition-colors">
                        <td className="py-3 px-3 text-center font-mono text-[11px] text-gray-600">
                          {iss.rowIndex}
                        </td>
                        <td className="py-3 px-3">
                          {iss.severity === 'critical' ? (
                            <span className="text-[10px] font-bold uppercase tracking-wider bg-red-50 text-red-700 px-2 py-0.5 rounded border border-red-200">
                              Critical
                            </span>
                          ) : iss.severity === 'warning' ? (
                            <span className="text-[10px] font-semibold uppercase tracking-wider bg-amber-50 text-amber-800 px-2 py-0.5 rounded border border-amber-200">
                              Warning
                            </span>
                          ) : (
                            <span className="text-[10px] font-medium uppercase tracking-wider bg-blue-50 text-blue-700 px-2 py-0.5 rounded border border-blue-200">
                              Info
                            </span>
                          )}
                        </td>
                        <td className="py-3 px-4">
                          <div className="font-semibold text-xs text-gray-900">{iss.accountName}</div>
                          {iss.accountCode && iss.accountCode !== '-' && (
                            <div className="text-[10px] font-mono text-gray-600">Code: {iss.accountCode}</div>
                          )}
                        </td>
                        <td className="py-3 px-4">
                          <div className="font-bold text-xs text-gray-900">{iss.title}</div>
                          <div className="text-[11px] text-gray-600 mt-0.5">{iss.description}</div>
                        </td>
                        <td className="py-3 px-4">
                          <div className="text-[11px] text-gray-900 font-medium bg-gray-100 border border-gray-200 p-2 rounded">
                            {iss.suggestedAction}
                          </div>
                        </td>
                        <td className="py-3 px-3 text-right font-mono text-xs">
                          {iss.debit > 0 && <div className="text-gray-900">Dr: {formatInr(iss.debit)}</div>}
                          {iss.credit > 0 && <div className="text-gray-600">Cr: {formatInr(iss.credit)}</div>}
                        </td>
                      </tr>
                    ))}
                    {filteredIssues.length === 0 && (
                      <tr>
                        <td colSpan={6} className="py-8 text-center text-gray-600 text-xs">
                          No issues found in this category.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          ) : (
            /* Cleaned Trial Balance View */
            <div className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm">
              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="bg-gray-50 border-b border-gray-200 text-gray-600 uppercase font-bold text-[10px] tracking-wider">
                    <tr>
                      <th className="py-2.5 px-3 w-12 text-center">Row</th>
                      <th className="py-2.5 px-3 w-28">Code</th>
                      <th className="py-2.5 px-4">Account Particulars</th>
                      <th className="py-2.5 px-3">Classification</th>
                      <th className="py-2.5 px-4 text-right">Debit (₹)</th>
                      <th className="py-2.5 px-4 text-right">Credit (₹)</th>
                      <th className="py-2.5 px-4 text-right font-bold">Net Balance (₹)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#E5E7EB]">
                    {parsedRows.map((r) => (
                      <tr key={r.rowIndex} className="hover:bg-gray-50 transition-colors">
                        <td className="py-2.5 px-3 text-center font-mono text-gray-600">{r.rowIndex}</td>
                        <td className="py-2.5 px-3 font-mono text-gray-600">{r.accountCode || '-'}</td>
                        <td className="py-2.5 px-4 font-semibold text-xs text-gray-900">{r.accountName}</td>
                        <td className="py-2.5 px-3 text-gray-600">
                          <span className="text-[10px] bg-slate-100 text-slate-700 px-1.5 py-0.5 rounded">
                            {r.groupCategory || 'General'}
                          </span>
                        </td>
                        <td className="py-2.5 px-4 text-right font-mono text-xs text-gray-900">
                          {r.debit > 0 ? formatInr(r.debit) : '-'}
                        </td>
                        <td className="py-2.5 px-4 text-right font-mono text-xs text-gray-600">
                          {r.credit > 0 ? formatInr(r.credit) : '-'}
                        </td>
                        <td className="py-2.5 px-4 text-right font-mono text-xs font-bold text-gray-900">
                          {formatInr(r.balance)}
                        </td>
                      </tr>
                    ))}
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