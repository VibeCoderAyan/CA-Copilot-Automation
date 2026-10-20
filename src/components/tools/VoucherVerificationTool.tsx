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
  CheckSquare,
  HelpCircle,
  Printer,
} from 'lucide-react';
import { InnerToolBadge } from '../common/InnerToolBadge';
import {
  type VerificationItem,
  type VerificationSummary,
  type VerificationMatchingConfig,
  type VoucherColumnMapping,
  type LedgerColumnMapping,
  detectVoucherColumns,
  detectLedgerColumns,
  verifyVouchersAndLedger,
  exportVerificationWorkbook,
} from '../../utils/voucherVerificationEngine';

interface VoucherVerificationToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

export const VoucherVerificationTool: React.FC<VoucherVerificationToolProps> = ({
  onBack,
  onShowToast,
}) => {
  const [voucherRows, setVoucherRows] = useState<Record<string, any>[]>([]);
  const [ledgerRows, setLedgerRows] = useState<Record<string, any>[]>([]);
  const [voucherHeaders, setVoucherHeaders] = useState<string[]>([]);
  const [ledgerHeaders, setLedgerHeaders] = useState<string[]>([]);

  const [voucherFileName, setVoucherFileName] = useState('');
  const [ledgerFileName, setLedgerFileName] = useState('');

  const [clientName, setClientName] = useState('');
  const [period, setPeriod] = useState('FY 2026-27');

  const [voucherMapping, setVoucherMapping] = useState<VoucherColumnMapping>({
    voucherNoCol: '',
    dateCol: '',
    accountCol: '',
    debitCol: '',
    creditCol: '',
    narrationCol: '',
    refCol: '',
  });

  const [ledgerMapping, setLedgerMapping] = useState<LedgerColumnMapping>({
    voucherNoCol: '',
    dateCol: '',
    accountCol: '',
    debitCol: '',
    creditCol: '',
    narrationCol: '',
    refCol: '',
  });

  const [matchingConfig, setMatchingConfig] = useState<VerificationMatchingConfig>({
    dateToleranceDays: 3,
    amountTolerance: 0.5,
    fuzzyAccountMatch: true,
  });

  const [activeTab, setActiveTab] = useState<'exceptions' | 'matched' | 'all' | 'summary'>('exceptions');
  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState<string>('all');
  const [showMappingPanel, setShowMappingPanel] = useState(false);
  const [showConfigPanel, setShowConfigPanel] = useState(false);

  // In-memory approval states for exceptions
  const [reviewStatuses, setReviewStatuses] = useState<Record<string, 'unresolved' | 'approved_exception' | 'query_sent'>>({});

  const voucherFileInputRef = useRef<HTMLInputElement>(null);
  const ledgerFileInputRef = useRef<HTMLInputElement>(null);

  // Run verification engine
  const verificationSummary: VerificationSummary = useMemo(() => {
    if (voucherRows.length === 0 && ledgerRows.length === 0) {
      return {
        totalVouchers: 0,
        totalLedgerLines: 0,
        matchedCount: 0,
        mismatchCount: 0,
        missingInLedgerCount: 0,
        missingInVouchersCount: 0,
        duplicateCount: 0,
        totalVoucherAmount: 0,
        totalLedgerAmount: 0,
        netVariance: 0,
        reconciliationPercentage: 0,
        items: [],
      };
    }

    const summary = verifyVouchersAndLedger(
      voucherRows,
      ledgerRows,
      voucherMapping,
      ledgerMapping,
      matchingConfig,
      voucherFileName || 'Vouchers.xlsx',
      ledgerFileName || 'Ledger.xlsx'
    );

    // Apply manual review status overrides
    summary.items.forEach((item) => {
      if (reviewStatuses[item.id]) {
        item.reviewStatus = reviewStatuses[item.id];
      }
    });

    return summary;
  }, [voucherRows, ledgerRows, voucherMapping, ledgerMapping, matchingConfig, voucherFileName, ledgerFileName, reviewStatuses]);

  // Handle Voucher file upload
  const handleVoucherUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setVoucherFileName(file.name);
    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json<Record<string, any>>(ws, { defval: '' });

        if (json.length === 0) {
          onShowToast('Empty File', 'Voucher register has no data rows.', 'error');
          return;
        }

        const headers = Object.keys(json[0]);
        setVoucherHeaders(headers);
        setVoucherRows(json);
        const auto = detectVoucherColumns(headers);
        setVoucherMapping(auto);
      } catch (err: any) {
        onShowToast('Upload Error', err.message || 'Failed to read voucher register', 'error');
      }
    };

    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  // Handle Ledger file upload
  const handleLedgerUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setLedgerFileName(file.name);
    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json<Record<string, any>>(ws, { defval: '' });

        if (json.length === 0) {
          onShowToast('Empty File', 'General Ledger file has no data rows.', 'error');
          return;
        }

        const headers = Object.keys(json[0]);
        setLedgerHeaders(headers);
        setLedgerRows(json);
        const auto = detectLedgerColumns(headers);
        setLedgerMapping(auto);
      } catch (err: any) {
        onShowToast('Upload Error', err.message || 'Failed to read general ledger', 'error');
      }
    };

    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  // Export 3-sheet Excel workbook
  const handleExportExcel = () => {
    if (verificationSummary.items.length === 0) {
      onShowToast('No Data', 'Please load and cross-verify records before exporting.', 'error');
      return;
    }

    try {
      const wb = exportVerificationWorkbook(verificationSummary, clientName, period);
      XLSX.writeFile(wb, `Voucher_Ledger_Verification_${clientName.replace(/\s+/g, '_')}_${period}.xlsx`);
      onShowToast('Workbook Exported', 'Exported 3-sheet audit verification report (Summary, Matched, Exceptions).', 'success');
    } catch (err: any) {
      onShowToast('Export Error', err.message || 'Failed to generate Excel report.', 'error');
    }
  };

  // Filtered items
  const filteredItems = useMemo(() => {
    return verificationSummary.items.filter((item) => {
      // Tab filter
      if (activeTab === 'exceptions' && item.status === 'matched') return false;
      if (activeTab === 'matched' && item.status !== 'matched') return false;

      // Status dropdown filter
      if (statusFilter !== 'all' && item.status !== statusFilter) return false;

      // Search query
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase();
        const vNo = item.voucherNo.toLowerCase();
        const acc = item.accountName.toLowerCase();
        const det = item.discrepancyDetails.toLowerCase();
        return vNo.includes(q) || acc.includes(q) || det.includes(q);
      }

      return true;
    });
  }, [verificationSummary.items, activeTab, statusFilter, searchQuery]);

  // Update review status
  const handleReviewStatusChange = (id: string, newStatus: 'unresolved' | 'approved_exception' | 'query_sent') => {
    setReviewStatuses((prev) => ({
      ...prev,
      [id]: newStatus,
    }));
    onShowToast('Status Updated', `Exception marked as ${newStatus.replace(/_/g, ' ').toUpperCase()}`, 'info');
  };

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="voucher-verification"
        title="Voucher & Ledger Verification"
        description="Cross-verify Daybook vouchers against General Ledger postings to detect unposted items, variances, and duplicate numbers."
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
              Map Fields
            </button>

            <button
              onClick={handleExportExcel}
              disabled={verificationSummary.items.length === 0}
              className="px-3.5 py-1.5 text-xs font-semibold rounded-md bg-gray-900 text-white hover:bg-black disabled:opacity-50 transition-colors flex items-center gap-1.5 shadow-sm cursor-pointer"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              Export Excel Report
            </button>
          </div>
        }
      />

      {/* Client & Period Header */}
      <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-2xs flex flex-wrap items-center gap-4 text-xs">
          <div className="flex items-center gap-2">
            <Building className="w-4 h-4 text-gray-600" />
            <span className="font-semibold text-gray-600">Client:</span>
            <input
              type="text"
              value={clientName}
              onChange={(e) => setClientName(e.target.value)}
              className="px-2 py-1 border border-gray-200 rounded text-xs font-medium focus:outline-none focus:border-gray-300 w-52"
            />
          </div>

          <div className="flex items-center gap-2">
            <span className="font-semibold text-gray-600">Audit Period:</span>
            <input
              type="text"
              value={period}
              onChange={(e) => setPeriod(e.target.value)}
              className="px-2 py-1 border border-gray-200 rounded text-xs font-medium focus:outline-none focus:border-gray-300 w-36"
            />
          </div>

          <div className="flex items-center gap-2">
            <span className="font-semibold text-gray-600">Voucher File:</span>
            <span className="font-medium text-gray-900">{voucherFileName || 'None uploaded'}</span>
          </div>

          <div className="flex items-center gap-2">
            <span className="font-semibold text-gray-600">Ledger File:</span>
            <span className="font-medium text-gray-900">{ledgerFileName || 'None uploaded'}</span>
          </div>
      </div>

      {/* Main Container */}
      <div className="space-y-6">
        {/* Dual Upload Bars */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
          {/* Voucher Register Box */}
          <input
            type="file"
            ref={voucherFileInputRef}
            onChange={handleVoucherUpload}
            accept=".xlsx,.xls,.csv"
            className="hidden"
            id="vch-verif-vch-upload"
          />
          {voucherRows.length === 0 ? (
            <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-6 sm:p-8 text-center transition-colors shadow-2xs">
              <label htmlFor="vch-verif-vch-upload" className="cursor-pointer flex flex-col items-center">
                <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mx-auto mb-3 shadow-2xs">
                  <FileCheck className="w-6 h-6" />
                </div>
                <h3 className="text-sm font-semibold text-gray-900">1. Voucher Register (Daybook)</h3>
                <p className="text-xs text-gray-500 mt-1 max-w-xs mx-auto">
                  Upload payment, receipt, contra, and journal vouchers list (.xlsx or .csv)
                </p>
                <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
                  <UploadCloud className="w-4 h-4" />
                  <span>Choose Voucher File</span>
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
                  <span className="text-xs font-bold text-gray-900 block truncate">{voucherFileName}</span>
                  <span className="text-[11px] text-gray-600 block mt-0.5">
                    Vouchers • {voucherRows.length} entries
                  </span>
                </div>
              </div>
              <div className="mt-4 pt-3 border-t border-gray-100 flex items-center justify-between">
                <span className="text-[11px] text-gray-500 font-medium">Ready for verification</span>
                <button
                  onClick={() => voucherFileInputRef.current?.click()}
                  className="px-3 py-1.5 text-xs text-gray-900 bg-gray-50 hover:bg-[#EAEAEA] rounded-lg border border-gray-200 flex items-center gap-1.5 font-medium transition-colors cursor-pointer"
                >
                  <UploadCloud className="w-3.5 h-3.5 text-gray-600" />
                  <span>Change Vouchers</span>
                </button>
              </div>
            </div>
          )}

          {/* General Ledger Box */}
          <input
            type="file"
            ref={ledgerFileInputRef}
            onChange={handleLedgerUpload}
            accept=".xlsx,.xls,.csv"
            className="hidden"
            id="vch-verif-gl-upload"
          />
          {ledgerRows.length === 0 ? (
            <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-6 sm:p-8 text-center transition-colors shadow-2xs">
              <label htmlFor="vch-verif-gl-upload" className="cursor-pointer flex flex-col items-center">
                <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mx-auto mb-3 shadow-2xs">
                  <FileSpreadsheet className="w-6 h-6" />
                </div>
                <h3 className="text-sm font-semibold text-gray-900">2. General Ledger Export</h3>
                <p className="text-xs text-gray-500 mt-1 max-w-xs mx-auto">
                  Upload posted general ledger entries from Tally, SAP, or ERP (.xlsx or .csv)
                </p>
                <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
                  <UploadCloud className="w-4 h-4" />
                  <span>Choose Ledger File</span>
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
                  <span className="text-xs font-bold text-gray-900 block truncate">{ledgerFileName}</span>
                  <span className="text-[11px] text-gray-600 block mt-0.5">
                    General Ledger • {ledgerRows.length} records
                  </span>
                </div>
              </div>
              <div className="mt-4 pt-3 border-t border-gray-100 flex items-center justify-between">
                <span className="text-[11px] text-gray-500 font-medium">Ready for verification</span>
                <button
                  onClick={() => ledgerFileInputRef.current?.click()}
                  className="px-3 py-1.5 text-xs text-gray-900 bg-gray-50 hover:bg-[#EAEAEA] rounded-lg border border-gray-200 flex items-center gap-1.5 font-medium transition-colors cursor-pointer"
                >
                  <UploadCloud className="w-3.5 h-3.5 text-gray-600" />
                  <span>Change Ledger</span>
                </button>
              </div>
            </div>
          )}
        </div>

        {/* Configuration Drawer */}
        {showConfigPanel && (
          <div className="bg-white border border-gray-200 rounded-xl p-5 shadow-sm">
            <div className="flex items-center justify-between pb-3 border-b border-gray-200">
              <div>
                <h3 className="text-sm font-bold text-gray-900">Audit Matching Tolerance Rules</h3>
                <p className="text-xs text-gray-600">Configure matching thresholds for dates, amounts, and account titles</p>
              </div>
              <button onClick={() => setShowConfigPanel(false)} className="text-xs text-gray-600 hover:text-gray-900">Close</button>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mt-4 text-xs">
              <div>
                <label className="block font-semibold text-gray-600 mb-1">Date Tolerance (Days)</label>
                <input
                  type="number"
                  value={matchingConfig.dateToleranceDays}
                  onChange={(e) => setMatchingConfig({ ...matchingConfig, dateToleranceDays: Math.max(0, parseInt(e.target.value) || 0) })}
                  className="w-full border border-gray-200 rounded p-2 focus:border-gray-300 outline-none"
                />
                <span className="text-[11px] text-gray-600 mt-1 block">Allow posting lag (e.g., 3 days between voucher prep & ledger entry)</span>
              </div>
              <div>
                <label className="block font-semibold text-gray-600 mb-1">Amount Rounding Tolerance (₹)</label>
                <input
                  type="number"
                  step="0.1"
                  value={matchingConfig.amountTolerance}
                  onChange={(e) => setMatchingConfig({ ...matchingConfig, amountTolerance: parseFloat(e.target.value) || 0 })}
                  className="w-full border border-gray-200 rounded p-2 focus:border-gray-300 outline-none"
                />
                <span className="text-[11px] text-gray-600 mt-1 block">Tolerate paisa rounding differences (default ₹0.50)</span>
              </div>
              <div className="flex items-center gap-2 pt-6">
                <input
                  type="checkbox"
                  id="fuzzyAccount"
                  checked={matchingConfig.fuzzyAccountMatch}
                  onChange={(e) => setMatchingConfig({ ...matchingConfig, fuzzyAccountMatch: e.target.checked })}
                  className="rounded text-gray-900 focus:ring-gray-900 h-4 w-4"
                />
                <label htmlFor="fuzzyAccount" className="font-semibold text-gray-900 cursor-pointer">
                  Enable Fuzzy Account Matching
                </label>
              </div>
            </div>
          </div>
        )}

        {/* Mapping Drawer */}
        {showMappingPanel && (
          <div className="bg-white border border-gray-200 rounded-xl p-5 shadow-sm space-y-4">
            <div className="flex items-center justify-between pb-3 border-b border-gray-200">
              <div>
                <h3 className="text-sm font-bold text-gray-900">Field Mappings for Vouchers & Ledger</h3>
                <p className="text-xs text-gray-600">Ensure voucher numbers, dates, and amounts align accurately</p>
              </div>
              <button onClick={() => setShowMappingPanel(false)} className="text-xs text-gray-600 hover:text-gray-900">Close</button>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 text-xs">
              {/* Voucher Side */}
              <div className="border border-gray-200 rounded-lg p-4">
                <h4 className="font-bold text-gray-900 mb-3 uppercase">Voucher Register Fields</h4>
                <div className="space-y-3">
                  <div>
                    <label className="block text-gray-600 mb-1">Voucher Number *</label>
                    <select
                      value={voucherMapping.voucherNoCol}
                      onChange={(e) => setVoucherMapping({ ...voucherMapping, voucherNoCol: e.target.value })}
                      className="w-full border border-gray-200 rounded p-1.5 outline-none"
                    >
                      <option value="">-- Select --</option>
                      {voucherHeaders.map((h) => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="block text-gray-600 mb-1">Voucher Date *</label>
                    <select
                      value={voucherMapping.dateCol}
                      onChange={(e) => setVoucherMapping({ ...voucherMapping, dateCol: e.target.value })}
                      className="w-full border border-gray-200 rounded p-1.5 outline-none"
                    >
                      <option value="">-- Select --</option>
                      {voucherHeaders.map((h) => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="block text-gray-600 mb-1">Account Name *</label>
                    <select
                      value={voucherMapping.accountCol}
                      onChange={(e) => setVoucherMapping({ ...voucherMapping, accountCol: e.target.value })}
                      className="w-full border border-gray-200 rounded p-1.5 outline-none"
                    >
                      <option value="">-- Select --</option>
                      {voucherHeaders.map((h) => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="block text-gray-600 mb-1">Debit / Amount Column *</label>
                    <select
                      value={voucherMapping.debitCol}
                      onChange={(e) => setVoucherMapping({ ...voucherMapping, debitCol: e.target.value })}
                      className="w-full border border-gray-200 rounded p-1.5 outline-none"
                    >
                      <option value="">-- Select --</option>
                      {voucherHeaders.map((h) => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </div>
                </div>
              </div>

              {/* Ledger Side */}
              <div className="border border-gray-200 rounded-lg p-4">
                <h4 className="font-bold text-gray-900 mb-3 uppercase">General Ledger Fields</h4>
                <div className="space-y-3">
                  <div>
                    <label className="block text-gray-600 mb-1">Voucher / Ref Number *</label>
                    <select
                      value={ledgerMapping.voucherNoCol}
                      onChange={(e) => setLedgerMapping({ ...ledgerMapping, voucherNoCol: e.target.value })}
                      className="w-full border border-gray-200 rounded p-1.5 outline-none"
                    >
                      <option value="">-- Select --</option>
                      {ledgerHeaders.map((h) => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="block text-gray-600 mb-1">Posting Date *</label>
                    <select
                      value={ledgerMapping.dateCol}
                      onChange={(e) => setLedgerMapping({ ...ledgerMapping, dateCol: e.target.value })}
                      className="w-full border border-gray-200 rounded p-1.5 outline-none"
                    >
                      <option value="">-- Select --</option>
                      {ledgerHeaders.map((h) => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="block text-gray-600 mb-1">Ledger Account Title *</label>
                    <select
                      value={ledgerMapping.accountCol}
                      onChange={(e) => setLedgerMapping({ ...ledgerMapping, accountCol: e.target.value })}
                      className="w-full border border-gray-200 rounded p-1.5 outline-none"
                    >
                      <option value="">-- Select --</option>
                      {ledgerHeaders.map((h) => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="block text-gray-600 mb-1">Debit / Amount Column *</label>
                    <select
                      value={ledgerMapping.debitCol}
                      onChange={(e) => setLedgerMapping({ ...ledgerMapping, debitCol: e.target.value })}
                      className="w-full border border-gray-200 rounded p-1.5 outline-none"
                    >
                      <option value="">-- Select --</option>
                      {ledgerHeaders.map((h) => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Verification Summary KPIs */}
        {verificationSummary.items.length > 0 && (
          <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
              <span className="text-xs font-semibold text-gray-600 uppercase">Reconciliation Rate</span>
              <div className="text-xl font-bold text-gray-900 mt-1">
                {verificationSummary.reconciliationPercentage.toFixed(1)}%
              </div>
              <div className="text-xs text-gray-600 mt-0.5">
                {verificationSummary.matchedCount} of {verificationSummary.totalVouchers} matched
              </div>
            </div>

            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
              <span className="text-xs font-semibold text-gray-600 uppercase">Amount / Date Variances</span>
              <div className="text-xl font-bold text-[#DC2626] mt-1">
                {verificationSummary.mismatchCount}
              </div>
              <div className="text-xs text-gray-600 mt-0.5">
                Discrepancy in recorded values
              </div>
            </div>

            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
              <span className="text-xs font-semibold text-gray-600 uppercase">Unposted in Ledger</span>
              <div className="text-xl font-bold text-[#D97706] mt-1">
                {verificationSummary.missingInLedgerCount}
              </div>
              <div className="text-xs text-gray-600 mt-0.5">
                Vouchers without ledger postings
              </div>
            </div>

            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
              <span className="text-xs font-semibold text-gray-600 uppercase">Unsupported in Ledger</span>
              <div className="text-xl font-bold text-[#7C3AED] mt-1">
                {verificationSummary.missingInVouchersCount}
              </div>
              <div className="text-xs text-gray-600 mt-0.5">
                Ledger lines without vouchers
              </div>
            </div>

            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
              <span className="text-xs font-semibold text-gray-600 uppercase">Duplicate Numbers</span>
              <div className="text-xl font-bold text-[#DC2626] mt-1">
                {verificationSummary.duplicateCount}
              </div>
              <div className="text-xs text-gray-600 mt-0.5">
                Reused voucher numbering
              </div>
            </div>
          </div>
        )}

        {/* Data Tabs & Tables */}
        {verificationSummary.items.length > 0 && (
          <div className="space-y-4">
            {/* Filter / Tab Bar */}
            <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 border-b border-gray-200 pb-2">
              <div className="flex items-center gap-2">
                <button
                  onClick={() => setActiveTab('exceptions')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'exceptions'
                      ? 'border-[#DC2626] text-[#DC2626]'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Audit Exceptions ({verificationSummary.items.filter((i) => i.status !== 'matched').length})
                </button>
                <button
                  onClick={() => setActiveTab('matched')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'matched'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Matched Entries ({verificationSummary.matchedCount})
                </button>
                <button
                  onClick={() => setActiveTab('all')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'all'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  All Transactions ({verificationSummary.items.length})
                </button>
                <button
                  onClick={() => setActiveTab('summary')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'summary'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Executive Report
                </button>
              </div>

              {activeTab !== 'summary' && (
                <div className="flex items-center gap-2">
                  <div className="relative">
                    <Search className="w-3.5 h-3.5 absolute left-2.5 top-1/2 -translate-y-1/2 text-gray-600" />
                    <input
                      type="text"
                      placeholder="Search voucher # or account..."
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
                    <option value="amount_mismatch">Amount Mismatch</option>
                    <option value="date_mismatch">Date Mismatch</option>
                    <option value="missing_in_ledger">Missing in Ledger</option>
                    <option value="missing_in_vouchers">Missing in Vouchers</option>
                    <option value="duplicate_voucher">Duplicate Voucher</option>
                  </select>
                </div>
              )}
            </div>

            {/* Table for Exceptions, Matched, and All */}
            {activeTab !== 'summary' && (
              <div className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm">
                <div className="overflow-x-auto">
                  <table className="w-full text-xs text-left">
                    <thead className="bg-gray-900 text-white font-semibold">
                      <tr>
                        <th className="p-3 w-10 text-center">#</th>
                        <th className="p-3">Voucher #</th>
                        <th className="p-3">Status</th>
                        <th className="p-3">Account / Particulars</th>
                        <th className="p-3 text-right">Voucher Amt (₹)</th>
                        <th className="p-3 text-right">Ledger Amt (₹)</th>
                        <th className="p-3 text-right">Variance (₹)</th>
                        <th className="p-3">Dates (Voucher vs Ledger)</th>
                        <th className="p-3">Source Tracing</th>
                        <th className="p-3">Audit Details</th>
                        <th className="p-3 text-center">Auditor Review</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB]">
                      {filteredItems.map((item, idx) => {
                        const isMismatch = item.status !== 'matched';
                        return (
                          <tr
                            key={item.id}
                            className={`hover:bg-gray-50 transition-colors ${
                              item.status === 'amount_mismatch'
                                ? 'bg-red-50/40'
                                : item.status === 'missing_in_ledger' || item.status === 'missing_in_vouchers'
                                ? 'bg-amber-50/30'
                                : item.status === 'duplicate_voucher'
                                ? 'bg-purple-50/30'
                                : ''
                            }`}
                          >
                            <td className="p-3 text-center text-gray-600 font-medium">{idx + 1}</td>
                            <td className="p-3 font-bold font-mono text-gray-900">{item.voucherNo}</td>
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
                              {item.status === 'date_mismatch' && (
                                <span className="px-2 py-0.5 rounded font-bold bg-[#FEF3C7] text-[#D97706] text-[10px]">
                                  DATE DIFF
                                </span>
                              )}
                              {item.status === 'missing_in_ledger' && (
                                <span className="px-2 py-0.5 rounded font-bold bg-[#FEF3C7] text-[#B45309] text-[10px]">
                                  NOT IN LEDGER
                                </span>
                              )}
                              {item.status === 'missing_in_vouchers' && (
                                <span className="px-2 py-0.5 rounded font-bold bg-[#EDE9FE] text-[#6D28D9] text-[10px]">
                                  NOT IN VOUCHERS
                                </span>
                              )}
                              {item.status === 'duplicate_voucher' && (
                                <span className="px-2 py-0.5 rounded font-bold bg-[#FEE2E2] text-[#B91C1C] text-[10px]">
                                  DUPLICATE NO
                                </span>
                              )}
                            </td>
                            <td className="p-3 font-medium text-gray-900">{item.accountName}</td>
                            <td className="p-3 text-right font-mono">
                              {item.voucherAmount > 0 ? `₹${item.voucherAmount.toLocaleString('en-IN')}` : '-'}
                            </td>
                            <td className="p-3 text-right font-mono">
                              {item.ledgerAmount > 0 ? `₹${item.ledgerAmount.toLocaleString('en-IN')}` : '-'}
                            </td>
                            <td className="p-3 text-right font-bold font-mono">
                              {item.varianceAmount !== 0 ? (
                                <span className="text-[#DC2626]">₹{item.varianceAmount.toLocaleString('en-IN')}</span>
                              ) : (
                                <span className="text-gray-900">₹0.00</span>
                              )}
                            </td>
                            <td className="p-3 text-[11px] text-gray-600">
                              <div>Voucher: {item.voucherDate || '-'}</div>
                              <div>Ledger: {item.ledgerDate || '-'}</div>
                              {item.dateDifferenceDays > 0 && (
                                <span className="text-[#DC2626] font-semibold text-[10px]">
                                  Diff: {item.dateDifferenceDays} days
                                </span>
                              )}
                            </td>
                            <td className="p-3 text-[11px] text-gray-600">
                              <div>Voucher: {item.voucherRow ? `Row ${item.voucherRow.rowIndex}` : 'None'}</div>
                              <div>Ledger: {item.ledgerRow ? `Row ${item.ledgerRow.rowIndex}` : 'None'}</div>
                            </td>
                            <td className="p-3 text-[11px] text-gray-900 max-w-xs">
                              {item.discrepancyDetails}
                            </td>
                            <td className="p-3 text-center">
                              {isMismatch ? (
                                <select
                                  value={item.reviewStatus}
                                  onChange={(e) => handleReviewStatusChange(item.id, e.target.value as any)}
                                  className={`text-[11px] font-semibold border rounded px-2 py-1 outline-none ${
                                    item.reviewStatus === 'approved_exception'
                                      ? 'bg-gray-50 text-gray-900 border-gray-200'
                                      : item.reviewStatus === 'query_sent'
                                      ? 'bg-amber-50 text-amber-700 border-amber-300'
                                      : 'bg-red-50 text-red-700 border-red-300'
                                  }`}
                                >
                                  <option value="unresolved">Unresolved</option>
                                  <option value="approved_exception">Approved Exception</option>
                                  <option value="query_sent">Query to Client</option>
                                </select>
                              ) : (
                                <span className="text-gray-900 font-bold text-xs">✓ Verified</span>
                              )}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* TAB 4: Summary Tab */}
            {activeTab === 'summary' && (
              <div className="bg-white border border-gray-200 rounded-xl p-8 shadow-sm">
                <div className="border-b-2 border-gray-300 pb-4 mb-6">
                  <div className="flex justify-between items-start">
                    <div>
                      <h2 className="text-xl font-bold text-gray-900">VOUCHER & LEDGER SUBSTANTIVE AUDIT REPORT</h2>
                      <p className="text-xs text-gray-600 mt-0.5">
                        Independent Cross-Reconciliation of Books of Original Entry vs General Ledger
                      </p>
                    </div>
                    <div className="text-right text-xs">
                      <div className="font-bold text-gray-900">{clientName}</div>
                      <div className="text-gray-600">{period}</div>
                      <div className="text-gray-600">Audit Run Date: {new Date().toLocaleDateString('en-IN')}</div>
                    </div>
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8 text-xs">
                  <div className="border border-gray-200 rounded-lg p-4">
                    <h4 className="font-bold text-gray-900 mb-3 uppercase">Verification Scope & Totals</h4>
                    <div className="space-y-2">
                      <div className="flex justify-between py-1 border-b border-gray-100">
                        <span className="text-gray-600">Daybook Vouchers Scanned</span>
                        <span className="font-bold">{verificationSummary.totalVouchers}</span>
                      </div>
                      <div className="flex justify-between py-1 border-b border-gray-100">
                        <span className="text-gray-600">General Ledger Lines Scanned</span>
                        <span className="font-bold">{verificationSummary.totalLedgerLines}</span>
                      </div>
                      <div className="flex justify-between py-1 border-b border-gray-100">
                        <span className="text-gray-600">Total Value of Vouchers (₹)</span>
                        <span className="font-bold">₹{verificationSummary.totalVoucherAmount.toLocaleString('en-IN')}</span>
                      </div>
                      <div className="flex justify-between py-1 border-b border-gray-100">
                        <span className="text-gray-600">Total Value of Ledger Lines (₹)</span>
                        <span className="font-bold">₹{verificationSummary.totalLedgerAmount.toLocaleString('en-IN')}</span>
                      </div>
                      <div className="flex justify-between py-1 border-b border-gray-100">
                        <span className="text-gray-600">Net Reconciliation Variance (₹)</span>
                        <span className="font-bold text-[#DC2626]">₹{verificationSummary.netVariance.toLocaleString('en-IN')}</span>
                      </div>
                    </div>
                  </div>

                  <div className="border border-gray-200 rounded-lg p-4">
                    <h4 className="font-bold text-[#DC2626] mb-3 uppercase">Exception Breakdown</h4>
                    <div className="space-y-2">
                      <div className="flex justify-between py-1 border-b border-gray-100">
                        <span className="text-gray-600">100% Matched Entries</span>
                        <span className="font-bold text-gray-900">{verificationSummary.matchedCount}</span>
                      </div>
                      <div className="flex justify-between py-1 border-b border-gray-100">
                        <span className="text-gray-600">Amount / Date Discrepancies</span>
                        <span className="font-bold text-[#DC2626]">{verificationSummary.mismatchCount}</span>
                      </div>
                      <div className="flex justify-between py-1 border-b border-gray-100">
                        <span className="text-gray-600">Unposted Vouchers (Missing in Ledger)</span>
                        <span className="font-bold text-[#D97706]">{verificationSummary.missingInLedgerCount}</span>
                      </div>
                      <div className="flex justify-between py-1 border-b border-gray-100">
                        <span className="text-gray-600">Unsupported Lines (Missing in Vouchers)</span>
                        <span className="font-bold text-[#7C3AED]">{verificationSummary.missingInVouchersCount}</span>
                      </div>
                      <div className="flex justify-between py-1 border-b border-gray-100">
                        <span className="text-gray-600">Duplicate Voucher Numbers</span>
                        <span className="font-bold text-[#DC2626]">{verificationSummary.duplicateCount}</span>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="p-4 bg-gray-50 rounded-lg border border-gray-300/30 text-xs text-gray-900">
                  <h5 className="font-bold mb-1">Standard CA Audit Disclosure</h5>
                  <p className="leading-relaxed">
                    Source spreadsheets and raw accounting transactions remain strictly immutable. Unmatched items and variances must be resolved by examining original supporting physical invoices, bank debit advices, or through formal journal adjustments approved by authorized management.
                  </p>
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
};