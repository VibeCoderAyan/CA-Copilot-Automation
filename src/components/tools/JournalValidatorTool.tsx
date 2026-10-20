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
  Filter,
  Info,
  ShieldCheck,
  Scale,
  Sliders,
  RefreshCw,
} from 'lucide-react';
import { InnerToolBadge } from '../common/InnerToolBadge';
import {
  type JournalRow,
  type JournalIssue,
  type JournalValidationSummary,
  type JournalColumnMapping,
  type JournalIssueSeverity,
  detectJournalColumns,
  validateJournalEntries,
  exportJournalValidationWorkbook,
} from '../../utils/journalValidatorEngine';

interface JournalValidatorToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

export const JournalValidatorTool: React.FC<JournalValidatorToolProps> = ({
  onBack,
  onShowToast,
}) => {
  const [rawRows, setRawRows] = useState<Record<string, any>[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [mapping, setMapping] = useState<JournalColumnMapping>({
    voucherNoCol: '',
    dateCol: '',
    accountCol: '',
    accountCodeCol: '',
    debitCol: '',
    creditCol: '',
    narrationCol: '',
    voucherTypeCol: '',
  });
  const [fileName, setFileName] = useState('');
  const [activeTab, setActiveTab] = useState<'all_issues' | 'critical' | 'suspicious' | 'vouchers' | 'clean_view'>('all_issues');
  const [searchQuery, setSearchQuery] = useState('');
  const [showMappingPanel, setShowMappingPanel] = useState(false);
  const [userReviewStatuses, setUserReviewStatuses] = useState<Record<string, JournalIssue['reviewStatus']>>({});
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Validate journal entries
  const { summary, parsedRows } = useMemo(() => {
    if (rawRows.length === 0) {
      return {
        summary: {
          totalVouchers: 0,
          totalRows: 0,
          totalDebit: 0,
          totalCredit: 0,
          netFileDifference: 0,
          isFileBalanced: true,
          criticalCount: 0,
          suspiciousCount: 0,
          warningCount: 0,
          infoCount: 0,
          unbalancedVouchersCount: 0,
          cashViolationsCount: 0,
          issues: [],
          vouchers: [],
        },
        parsedRows: [],
      };
    }
    const res = validateJournalEntries(rawRows, mapping, fileName || 'Journal_Entries');
    // Apply user review overrides
    res.summary.issues = res.summary.issues.map((iss) => {
      if (userReviewStatuses[iss.id]) {
        return { ...iss, reviewStatus: userReviewStatuses[iss.id] };
      }
      return iss;
    });
    return res;
  }, [rawRows, mapping, fileName, userReviewStatuses]);

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
          onShowToast('Empty File', 'Uploaded workbook contains no data rows.', 'error');
          return;
        }

        const detectedHeaders = Object.keys(json[0]);
        setHeaders(detectedHeaders);
        const detectedMapping = detectJournalColumns(detectedHeaders);
        setMapping(detectedMapping);
        setRawRows(json);
        setUserReviewStatuses({});
      } catch (err: any) {
        onShowToast('Parsing Error', err.message || 'Failed to read journal entries spreadsheet.', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  
  const handleExportExcel = () => {
    if (parsedRows.length === 0) return;
    try {
      const wb = exportJournalValidationWorkbook(summary, parsedRows, 'Client Entity', 'FY 2026-27');
      XLSX.writeFile(wb, `Journal_Entry_Audit_Report_${new Date().toISOString().split('T')[0]}.xlsx`);
      onShowToast('Audit Report Exported', 'Downloaded 4-sheet journal findings and validated register workbook.', 'success');
    } catch (err: any) {
      onShowToast('Export Failed', err.message || 'Could not export Excel workbook.', 'error');
    }
  };

  const handleMarkReview = (issueId: string, status: JournalIssue['reviewStatus']) => {
    setUserReviewStatuses((prev) => ({ ...prev, [issueId]: status }));
    onShowToast('Review Status Updated', `Issue marked as ${status.replace(/_/g, ' ').toUpperCase()}.`, 'info');
  };

  const filteredIssues = summary.issues.filter((iss) => {
    if (activeTab === 'critical' && iss.severity !== 'critical') return false;
    if (activeTab === 'suspicious' && iss.severity !== 'suspicious') return false;

    if (!searchQuery) return true;
    const q = searchQuery.toLowerCase();
    return (
      iss.voucherNo.toLowerCase().includes(q) ||
      iss.accountName.toLowerCase().includes(q) ||
      iss.title.toLowerCase().includes(q) ||
      iss.description.toLowerCase().includes(q) ||
      iss.suggestedAction.toLowerCase().includes(q)
    );
  });

  const formatInr = (val: number) => {
    return val.toLocaleString('en-IN', { style: 'currency', currency: 'INR', maximumFractionDigits: 2 });
  };

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="journal-validator"
        title="Journal Entry Validator"
        description="Audit voucher debit-credit totals, missing dates, duplicate vouchers, Sec 40A(3) cash limits, and suspense postings."
        badge="Audit Scrutiny"
        onBack={onBack}
        actionButtons={
          <div className="flex items-center gap-2 flex-wrap">
            {rawRows.length > 0 && (
              <button
                onClick={() => setShowMappingPanel(!showMappingPanel)}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-white text-gray-700 border border-gray-200 hover:bg-gray-50 transition-all cursor-pointer"
              >
                <Sliders className="w-3.5 h-3.5" />
                <span>Map Columns</span>
              </button>
            )}

            <button
              onClick={handleExportExcel}
              disabled={parsedRows.length === 0}
              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-semibold bg-gray-900 text-white hover:bg-gray-900 disabled:opacity-40 disabled:cursor-not-allowed shadow-xs transition-all cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Export Issue Report</span>
            </button>
          </div>
        }
      />

        {/* Column Mapping Panel */}
        {showMappingPanel && headers.length > 0 && (
          <div className="mt-4 pt-4 border-t border-gray-100 bg-[#F9FAFA] -mx-5 -mb-5 p-5 rounded-b-xl space-y-3">
            <span className="text-xs font-bold text-gray-900 block">
              Journal Entry Column Mapping
            </span>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3 text-xs">
              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Voucher Number</label>
                <select
                  value={mapping.voucherNoCol}
                  onChange={(e) => setMapping({ ...mapping, voucherNoCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-white border border-gray-200 rounded text-xs"
                >
                  {headers.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Date</label>
                <select
                  value={mapping.dateCol}
                  onChange={(e) => setMapping({ ...mapping, dateCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-white border border-gray-200 rounded text-xs"
                >
                  {headers.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Account / Particulars</label>
                <select
                  value={mapping.accountCol}
                  onChange={(e) => setMapping({ ...mapping, accountCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-white border border-gray-200 rounded text-xs"
                >
                  {headers.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Debit Amount</label>
                <select
                  value={mapping.debitCol}
                  onChange={(e) => setMapping({ ...mapping, debitCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-white border border-gray-200 rounded text-xs"
                >
                  {headers.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Credit Amount</label>
                <select
                  value={mapping.creditCol}
                  onChange={(e) => setMapping({ ...mapping, creditCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-white border border-gray-200 rounded text-xs"
                >
                  {headers.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Narration / Remarks</label>
                <select
                  value={mapping.narrationCol}
                  onChange={(e) => setMapping({ ...mapping, narrationCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-white border border-gray-200 rounded text-xs"
                >
                  <option value="">-- Optional --</option>
                  {headers.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>
            </div>
          </div>
        )}

      {/* 2. File Ingestion & Drag Drop */}
      {rawRows.length === 0 ? (
        <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-8 sm:p-10 text-center space-y-4 shadow-2xs transition-colors">
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileUpload}
            accept=".xlsx,.xls,.csv"
            className="hidden"
            id="journal-validator-upload"
          />
          <label htmlFor="journal-validator-upload" className="cursor-pointer flex flex-col items-center">
            <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mx-auto mb-3 shadow-2xs">
              <UploadCloud className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-semibold text-gray-900">
              Import Journal Entry Ledger / Daybook
            </h3>
            <p className="text-xs text-gray-500 mt-1 max-w-md mx-auto leading-relaxed">
              Supports Excel (.xlsx, .xls) and CSV daybooks exported from Tally, Busy, SAP, Zoho Books, or QuickBooks.
            </p>
            <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
              <UploadCloud className="w-4 h-4" />
              <span>Select Journal File</span>
            </div>
            <span className="text-[11px] text-gray-400 font-medium mt-3 block">
              Source Journal Integrity Protected · Never modifies entries without explicit approval
            </span>
          </label>
        </div>
      ) : (
        /* 3. Executive Scrutiny KPI Metric Cards */
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
          <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs">
            <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider block">
              Vouchers Audited
            </span>
            <div className="flex items-baseline justify-between mt-1.5">
              <span className="text-xl font-bold text-gray-900">
                {summary.totalVouchers}
              </span>
              <span className="text-[11px] text-gray-500">
                {summary.totalRows} line items
              </span>
            </div>
          </div>

          <div className={`p-4 rounded-xl border shadow-xs ${
            summary.isFileBalanced
              ? 'bg-white border-gray-200'
              : 'bg-red-50/30 border-red-200'
          }`}>
            <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider block">
              Debit vs Credit Tie-Out
            </span>
            <div className="flex items-baseline justify-between mt-1.5">
              <span className={`text-xl font-bold ${summary.isFileBalanced ? 'text-gray-800' : 'text-red-700'}`}>
                {summary.isFileBalanced ? 'BALANCED' : `DIFF: ₹${summary.netFileDifference.toFixed(2)}`}
              </span>
              <span className="text-[11px] text-gray-500">
                Dr: {formatInr(summary.totalDebit)}
              </span>
            </div>
          </div>

          <div className="bg-white p-4 rounded-xl border border-red-200 shadow-xs bg-red-50/20">
            <span className="text-[11px] font-semibold text-red-700 uppercase tracking-wider block">
              Critical Exceptions
            </span>
            <div className="flex items-baseline justify-between mt-1.5">
              <span className="text-xl font-bold text-red-700">
                {summary.criticalCount}
              </span>
              <span className="text-[11px] text-red-600 font-semibold">
                {summary.unbalancedVouchersCount} Imbalances • {summary.cashViolationsCount} 40A(3)
              </span>
            </div>
          </div>

          <div className="bg-white p-4 rounded-xl border border-amber-200 shadow-xs bg-amber-50/20">
            <span className="text-[11px] font-semibold text-amber-700 uppercase tracking-wider block">
              Suspicious Inconsistencies
            </span>
            <div className="flex items-baseline justify-between mt-1.5">
              <span className="text-xl font-bold text-amber-800">
                {summary.suspiciousCount + summary.warningCount}
              </span>
              <span className="text-[11px] text-amber-700">
                Requires Audit Inquiry
              </span>
            </div>
          </div>
        </div>
      )}

      {/* 4. Findings & Audit Workspace */}
      {rawRows.length > 0 && (
        <div className="bg-white rounded-xl border border-gray-200 shadow-xs overflow-hidden">
          {/* Tabs and Search */}
          <div className="p-4 border-b border-gray-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div className="flex items-center gap-1.5 overflow-x-auto pb-1 sm:pb-0">
              {[
                { id: 'all_issues' as const, label: 'All Findings', count: summary.issues.length },
                { id: 'critical' as const, label: 'Critical Errors', count: summary.criticalCount },
                { id: 'suspicious' as const, label: 'Suspicious Items', count: summary.suspiciousCount },
                { id: 'vouchers' as const, label: 'Vouchers List', count: summary.vouchers.length },
                { id: 'clean_view' as const, label: 'Daybook View', count: summary.totalRows },
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
                placeholder="Filter voucher, account, finding..."
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

          {/* Tab View: Vouchers Overview */}
          {activeTab === 'vouchers' ? (
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs text-gray-800">
                <thead className="bg-[#F8F9FA] border-b border-gray-200 text-gray-600 font-semibold uppercase text-[10px] tracking-wider">
                  <tr>
                    <th className="px-4 py-3">Voucher #</th>
                    <th className="px-4 py-3">Date</th>
                    <th className="px-4 py-3 text-center">Lines</th>
                    <th className="px-4 py-3 text-right">Debit (₹)</th>
                    <th className="px-4 py-3 text-right">Credit (₹)</th>
                    <th className="px-4 py-3 text-right">Imbalance (₹)</th>
                    <th className="px-4 py-3">Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {summary.vouchers.map((v) => (
                    <tr key={v.voucherNo} className="hover:bg-gray-50/80 transition-colors">
                      <td className="px-4 py-2.5 font-mono font-bold text-gray-900">{v.voucherNo}</td>
                      <td className="px-4 py-2.5 text-gray-600">{v.date || 'Missing'}</td>
                      <td className="px-4 py-2.5 text-center font-mono">{v.rowCount}</td>
                      <td className="px-4 py-2.5 text-right font-mono tabular-nums">{v.totalDebit.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
                      <td className="px-4 py-2.5 text-right font-mono tabular-nums">{v.totalCredit.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
                      <td className={`px-4 py-2.5 text-right font-mono font-bold tabular-nums ${v.diff > 0.01 ? 'text-red-700' : 'text-gray-400'}`}>
                        {v.diff.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                      </td>
                      <td className="px-4 py-2.5">
                        {v.isBalanced ? (
                          <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-gray-800 bg-gray-50 px-2 py-0.5 rounded">
                            <Check className="w-3 h-3" /> Balanced
                          </span>
                        ) : (
                          <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-red-700 bg-red-50 px-2 py-0.5 rounded">
                            <AlertCircle className="w-3 h-3" /> Unbalanced
                          </span>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : activeTab === 'clean_view' ? (
            /* Tab View: Full Daybook Register */
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs text-gray-800">
                <thead className="bg-[#F8F9FA] border-b border-gray-200 text-gray-600 font-semibold uppercase text-[10px] tracking-wider">
                  <tr>
                    <th className="px-4 py-3">Row #</th>
                    <th className="px-4 py-3">Voucher #</th>
                    <th className="px-4 py-3">Date</th>
                    <th className="px-4 py-3">Account Name</th>
                    <th className="px-4 py-3">Code</th>
                    <th className="px-4 py-3 text-right">Debit (₹)</th>
                    <th className="px-4 py-3 text-right">Credit (₹)</th>
                    <th className="px-4 py-3">Narration</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {parsedRows.map((r) => (
                    <tr key={r.id} className="hover:bg-gray-50/80 transition-colors">
                      <td className="px-4 py-2.5 font-mono text-gray-500">{r.rowIndex}</td>
                      <td className="px-4 py-2.5 font-mono font-semibold text-gray-900">{r.voucherNo}</td>
                      <td className="px-4 py-2.5">{r.voucherDate || <span className="text-red-500">Missing</span>}</td>
                      <td className="px-4 py-2.5 font-medium text-gray-900">{r.accountName || <span className="text-red-500">(Blank)</span>}</td>
                      <td className="px-4 py-2.5 font-mono text-gray-500">{r.accountCode || '-'}</td>
                      <td className="px-4 py-2.5 text-right font-mono tabular-nums">{r.debit > 0 ? r.debit.toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '-'}</td>
                      <td className="px-4 py-2.5 text-right font-mono tabular-nums">{r.credit > 0 ? r.credit.toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '-'}</td>
                      <td className="px-4 py-2.5 text-gray-500 max-w-xs truncate">{r.narration || '-'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            /* Findings List Table */
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs text-gray-800">
                <thead className="bg-[#F8F9FA] border-b border-gray-200 text-gray-600 font-semibold uppercase text-[10px] tracking-wider">
                  <tr>
                    <th className="px-4 py-3">Severity</th>
                    <th className="px-4 py-3">Voucher #</th>
                    <th className="px-4 py-3">Row #</th>
                    <th className="px-4 py-3">Account Head</th>
                    <th className="px-4 py-3">Finding & Observation</th>
                    <th className="px-4 py-3">Suggested Action</th>
                    <th className="px-4 py-3 text-right">Dr / Cr (₹)</th>
                    <th className="px-4 py-3 text-center">Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {filteredIssues.map((iss) => {
                    const isCritical = iss.severity === 'critical';
                    const isSuspicious = iss.severity === 'suspicious';

                    return (
                      <tr key={iss.id} className="hover:bg-gray-50/80 transition-colors">
                        <td className="px-4 py-3">
                          <span
                            className={`inline-block px-2 py-0.5 rounded text-[10px] font-bold tracking-wider uppercase ${
                              isCritical
                                ? 'bg-red-600 text-white'
                                : isSuspicious
                                ? 'bg-amber-100 text-amber-800 border border-amber-200'
                                : 'bg-blue-50 text-blue-700'
                            }`}
                          >
                            {iss.severity}
                          </span>
                        </td>
                        <td className="px-4 py-3 font-mono font-semibold text-gray-900">
                          {iss.voucherNo}
                        </td>
                        <td className="px-4 py-3 font-mono text-gray-500">
                          Row {iss.rowIndex}
                        </td>
                        <td className="px-4 py-3 font-medium text-gray-900 max-w-[180px] truncate" title={iss.accountName}>
                          {iss.accountName}
                        </td>
                        <td className="px-4 py-3 max-w-sm">
                          <div className="font-semibold text-gray-900">{iss.title}</div>
                          <div className="text-[11px] text-gray-500 mt-0.5">{iss.description}</div>
                        </td>
                        <td className="px-4 py-3 text-gray-600 text-[11px] max-w-xs">
                          {iss.suggestedAction}
                        </td>
                        <td className="px-4 py-3 text-right font-mono text-gray-900">
                          {iss.debit > 0 ? `Dr: ${iss.debit.toLocaleString('en-IN')}` : iss.credit > 0 ? `Cr: ${iss.credit.toLocaleString('en-IN')}` : '-'}
                        </td>
                        <td className="px-4 py-3 text-center">
                          <div className="flex items-center justify-center gap-1">
                            <button
                              onClick={() => handleMarkReview(iss.id, 'reviewed')}
                              className={`p-1 rounded text-xs transition-colors ${
                                iss.reviewStatus === 'reviewed'
                                  ? 'bg-gray-900 text-white'
                                  : 'text-gray-400 hover:text-gray-800 hover:bg-gray-50'
                              }`}
                              title="Mark Reviewed"
                            >
                              <Check className="w-3.5 h-3.5" />
                            </button>
                            <button
                              onClick={() => handleMarkReview(iss.id, 'query_raised')}
                              className={`p-1 rounded text-xs transition-colors ${
                                iss.reviewStatus === 'query_raised'
                                  ? 'bg-amber-600 text-white'
                                  : 'text-gray-400 hover:text-amber-700 hover:bg-amber-50'
                              }`}
                              title="Flag for Query"
                            >
                              <AlertTriangle className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        </td>
                      </tr>
                    );
                  })}

                  {filteredIssues.length === 0 && (
                    <tr>
                      <td colSpan={8} className="py-12 text-center text-xs text-gray-500">
                        <CheckCircle2 className="w-8 h-8 text-gray-600 mx-auto mb-2" />
                        <p className="font-semibold text-gray-800">No issues found matching this filter.</p>
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}
    </div>
  );
};