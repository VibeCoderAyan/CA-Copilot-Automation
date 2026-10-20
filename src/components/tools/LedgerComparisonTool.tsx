import { ToolHeader } from '../common/ToolHeader';
import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  UploadCloud,
  Download,
  CheckCircle2,
  AlertTriangle,
  Scale,
  FileSpreadsheet,
  Sliders,
  Check,
  AlertOctagon,
  ArrowRightLeft
} from 'lucide-react';
import {
  parseRawLedgerRows,
  compareLedgers,
  createLedgerReconciliationWorkbook,
  type LedgerEntry,
  type MatchingCriteria,
  type LedgerComparisonResult,
  DEFAULT_MATCHING_CRITERIA,
} from '../../utils/ledgerComparisonEngine';
import { DataFormatGuide } from '../common/DataFormatGuide';
import { formatINR } from '../../utils/dataSanitizer';

interface LedgerComparisonToolProps {
  onBack: () => void;
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
}

export const LedgerComparisonTool: React.FC<LedgerComparisonToolProps> = ({ onBack, onShowToast }) => {
  const [fileA, setFileA] = useState<{ name: string; size: string } | null>(null);
  const [fileB, setFileB] = useState<{ name: string; size: string } | null>(null);

  const [entriesA, setEntriesA] = useState<LedgerEntry[]>([]);
  const [entriesB, setEntriesB] = useState<LedgerEntry[]>([]);

  const [criteria, setCriteria] = useState<MatchingCriteria>({ ...DEFAULT_MATCHING_CRITERIA });
  const [comparisonResult, setComparisonResult] = useState<LedgerComparisonResult | null>(null);
  const [isCompared, setIsCompared] = useState(false);
  const [activeTab, setActiveTab] = useState<'summary' | 'confirmed' | 'probable' | 'discrepancies' | 'unmatchedA' | 'unmatchedB'>('summary');
  const [validationError, setValidationError] = useState<string | null>(null);

  const fileInputARef = useRef<HTMLInputElement>(null);
  const fileInputBRef = useRef<HTMLInputElement>(null);

  const handleUploadA = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;
    setValidationError(null);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary', cellDates: true });
        const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]) as Record<string, any>[];
        if (!json || json.length === 0) throw new Error('Ledger A contains zero rows.');

        const parsed = parseRawLedgerRows(json, 'A');
        setFileA({ name: f.name, size: `${Math.round(f.size / 1024)} KB` });
        setEntriesA(parsed);
        setIsCompared(false);
        setComparisonResult(null);
      } catch (err: any) {
        setValidationError(`Failed to parse Ledger A: ${err.message}`);
        onShowToast('Ledger A Error', err.message, 'error');
      }
    };
    reader.readAsBinaryString(f);
  };

  const handleUploadB = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;
    setValidationError(null);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary', cellDates: true });
        const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]) as Record<string, any>[];
        if (!json || json.length === 0) throw new Error('Ledger B contains zero rows.');

        const parsed = parseRawLedgerRows(json, 'B');
        setFileB({ name: f.name, size: `${Math.round(f.size / 1024)} KB` });
        setEntriesB(parsed);
        setIsCompared(false);
        setComparisonResult(null);
      } catch (err: any) {
        setValidationError(`Failed to parse Ledger B: ${err.message}`);
        onShowToast('Ledger B Error', err.message, 'error');
      }
    };
    reader.readAsBinaryString(f);
  };


  const handleReset = () => {
    setFileA(null);
    setFileB(null);
    setEntriesA([]);
    setEntriesB([]);
    setComparisonResult(null);
    setIsCompared(false);
    setValidationError(null);
    if (fileInputARef.current) fileInputARef.current.value = '';
    if (fileInputBRef.current) fileInputBRef.current.value = '';
    onShowToast('Reset Complete', 'Cleared ledger entries.', 'info');
  };

  const handleCompare = () => {
    if (entriesA.length === 0 || entriesB.length === 0) {
      return;
    }

    const res = compareLedgers(entriesA, entriesB, fileA?.name || 'Ledger A', fileB?.name || 'Ledger B', criteria);
    setComparisonResult(res);
    setIsCompared(true);
    setActiveTab('summary');

    onShowToast(
      'Reconciliation Complete',
      `Match rate: ${res.matchRate}%. Confirmed: ${res.matchedPairs.length}, Probable: ${res.probablePairs.length}, Discrepancies: ${res.discrepancies.length}.`,
      'success'
    );
  };

  const handleDownload = () => {
    if (!comparisonResult) return;
    const wb = createLedgerReconciliationWorkbook(comparisonResult);
    const outName = `Ledger_Reconciliation_${fileA?.name.replace(/\.[^/.]+$/, '')}_vs_${fileB?.name.replace(/\.[^/.]+$/, '')}.xlsx`;
    XLSX.writeFile(wb, outName);
    onShowToast('Downloaded', `Exported 6-sheet reconciliation package as ${outName}.`, 'success');
  };

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="ledger-compare"
        title="Reciprocal Ledger & Inter-Co Rec"
        description="Reconcile two accounting ledgers (Vendor vs Client, Inter-Company, or Period vs Period) with fuzzy matching."
        badge="Audit Grade"
        onBack={onBack}
        actionButtons={
          (fileA || fileB || comparisonResult) ? (
            <div className="flex items-center gap-2">
              <button
                onClick={handleReset}
                className="px-3 py-1.5 text-xs font-medium text-gray-600 hover:text-gray-900 bg-white hover:bg-gray-50 border border-gray-200 rounded-lg transition-colors cursor-pointer"
              >
                Clear
              </button>
              {comparisonResult && (
                <button
                  onClick={handleDownload}
                  className="px-4 py-2 text-xs font-semibold text-white bg-gray-900 hover:bg-black rounded-lg shadow-sm flex items-center gap-2 transition-colors cursor-pointer"
                >
                  <Download className="w-4 h-4" />
                  <span>Export Reconciliation (.xlsx)</span>
                </button>
              )}
            </div>
          ) : undefined
        }
      />

      {/* 2. Format Requirements Guide */}
      <DataFormatGuide
        toolType="bank-rec"
        acceptedFormats=".xlsx, .xls, .csv"
        requiredFields={[
          { name: 'Posting Date', description: 'Transaction date', aliases: ['Date', 'Txn Date', 'Vch Date'] },
          { name: 'Voucher / Ref No', description: 'Voucher number for exact or fuzzy alignment', aliases: ['Voucher No', 'Vch No', 'Invoice No', 'Bill No'] },
          { name: 'Particulars / Account', description: 'Party or Account description', aliases: ['Particulars', 'Account Name', 'Party Name'] },
          { name: 'Debit & Credit Amounts', description: 'Financial values in debit and credit columns', aliases: ['Debit', 'Credit', 'Dr', 'Cr', 'Amount'] },
        ]}
        notes={[
          'Inter-Company Inverse Dr/Cr: Automatically maps Debit in our books to Credit in the counterparty ledger.',
          'Confidence Scoring: Differentiates 100% Confirmed matches from Probable matches with date drift.',
          'Amount Discrepancy Detection: Identifies entries with matching vouchers where amounts booked differ.',
        ]}
      />

      {/* 3. Validation Error Banner */}
      {validationError && (
        <div className="p-4 rounded-xl bg-[#FDF2F2] border border-[#F8D7DA] text-xs text-[#9B2C2C] flex items-center justify-between">
          <div className="flex items-center gap-2">
            <AlertTriangle className="w-4 h-4 shrink-0" />
            <span>{validationError}</span>
          </div>
          <button onClick={() => setValidationError(null)} className="text-xs font-semibold underline">
            Dismiss
          </button>
        </div>
      )}

      {/* 4. Dual Ledger Upload Grid */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
        {/* Ledger A Box */}
        <input type="file" ref={fileInputARef} onChange={handleUploadA} accept=".xlsx,.xls,.csv" className="hidden" id="ledger-a-upload" />
        {!fileA ? (
          <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-6 sm:p-8 text-center transition-colors shadow-2xs">
            <label htmlFor="ledger-a-upload" className="cursor-pointer flex flex-col items-center">
              <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mb-3 shadow-2xs">
                <FileSpreadsheet className="w-6 h-6" />
              </div>
              <h3 className="text-sm font-semibold text-gray-900">Ledger A (Our Books / Primary)</h3>
              <p className="text-xs text-gray-500 mt-1 max-w-xs mx-auto">
                Customer/Vendor ledger dump from your accounting ERP (.xlsx, .xls, .csv)
              </p>
              <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
                <UploadCloud className="w-4 h-4" />
                <span>Upload Primary Ledger A</span>
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
                <span className="text-xs font-bold text-gray-900 block truncate">{fileA.name}</span>
                <span className="text-[11px] text-gray-600 block mt-0.5">
                  Ledger A • {entriesA.length} entries • {fileA.size}
                </span>
              </div>
            </div>
            <div className="mt-4 pt-3 border-t border-gray-100 flex items-center justify-between">
              <span className="text-[11px] text-gray-500 font-medium">Ready for ledger tie-out</span>
              <button
                onClick={() => fileInputARef.current?.click()}
                className="px-3 py-1.5 text-xs text-gray-900 bg-gray-50 hover:bg-[#EAEAEA] rounded-lg border border-gray-200 flex items-center gap-1.5 font-medium transition-colors cursor-pointer"
              >
                <UploadCloud className="w-3.5 h-3.5 text-gray-600" />
                <span>Change Ledger A</span>
              </button>
            </div>
          </div>
        )}

        {/* Ledger B Box */}
        <input type="file" ref={fileInputBRef} onChange={handleUploadB} accept=".xlsx,.xls,.csv" className="hidden" id="ledger-b-upload" />
        {!fileB ? (
          <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-6 sm:p-8 text-center transition-colors shadow-2xs">
            <label htmlFor="ledger-b-upload" className="cursor-pointer flex flex-col items-center">
              <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mb-3 shadow-2xs">
                <ArrowRightLeft className="w-6 h-6" />
              </div>
              <h3 className="text-sm font-semibold text-gray-900">Ledger B (Counterparty / Secondary)</h3>
              <p className="text-xs text-gray-500 mt-1 max-w-xs mx-auto">
                Vendor confirmation statement, customer ledger, or 2nd ERP dump (.xlsx, .xls, .csv)
              </p>
              <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
                <UploadCloud className="w-4 h-4" />
                <span>Upload Counterparty Ledger B</span>
              </div>
            </label>
          </div>
        ) : (
          <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-sm flex flex-col justify-between">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-lg bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 shrink-0">
                <ArrowRightLeft className="w-5 h-5" />
              </div>
              <div className="min-w-0 flex-1">
                <span className="text-xs font-bold text-gray-900 block truncate">{fileB.name}</span>
                <span className="text-[11px] text-gray-600 block mt-0.5">
                  Ledger B • {entriesB.length} entries • {fileB.size}
                </span>
              </div>
            </div>
            <div className="mt-4 pt-3 border-t border-gray-100 flex items-center justify-between">
              <span className="text-[11px] text-gray-500 font-medium">Ready for ledger tie-out</span>
              <button
                onClick={() => fileInputBRef.current?.click()}
                className="px-3 py-1.5 text-xs text-gray-900 bg-gray-50 hover:bg-[#EAEAEA] rounded-lg border border-gray-200 flex items-center gap-1.5 font-medium transition-colors cursor-pointer"
              >
                <UploadCloud className="w-3.5 h-3.5 text-gray-600" />
                <span>Change Ledger B</span>
              </button>
            </div>
          </div>
        )}
      </div>

      {/* 5. Configurable Reconciliation Rules Bar */}
      <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm space-y-3">
        <div className="flex flex-wrap items-center justify-between gap-4 text-xs">
          <div className="flex items-center gap-4 flex-wrap">
            <div className="flex items-center gap-2">
              <Sliders className="w-4 h-4 text-gray-900" />
              <span className="font-bold text-gray-900">Matching Rules:</span>
            </div>

            <label className="flex items-center gap-1.5 cursor-pointer">
              <input
                type="checkbox"
                checked={criteria.matchInverseDrCr}
                onChange={(e) => setCriteria({ ...criteria, matchInverseDrCr: e.target.checked })}
                className="rounded text-gray-900 focus:ring-gray-900"
              />
              <span className="font-medium text-gray-900">Inter-Company Dr ↔ Cr Inversion</span>
            </label>

            <div className="flex items-center gap-1.5">
              <span className="text-gray-600">Amount Tolerance:</span>
              <select
                value={criteria.amountTolerance}
                onChange={(e) => setCriteria({ ...criteria, amountTolerance: parseFloat(e.target.value) })}
                className="px-2 py-1 bg-gray-50 border border-gray-200 rounded text-xs font-medium text-gray-900"
              >
                <option value={0.0}>Exact (±₹0.00)</option>
                <option value={1.0}>Rounding (±₹1.00)</option>
                <option value={5.0}>Tolerance (±₹5.00)</option>
                <option value={10.0}>Tolerance (±₹10.00)</option>
              </select>
            </div>

            <div className="flex items-center gap-1.5">
              <span className="text-gray-600">Date Tolerance:</span>
              <select
                value={criteria.dateToleranceDays}
                onChange={(e) => setCriteria({ ...criteria, dateToleranceDays: parseInt(e.target.value) })}
                className="px-2 py-1 bg-gray-50 border border-gray-200 rounded text-xs font-medium text-gray-900"
              >
                <option value={0}>Same Date (±0 days)</option>
                <option value={3}>Within ±3 Days</option>
                <option value={7}>Within ±7 Days</option>
              </select>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <button
              onClick={handleCompare}
              disabled={entriesA.length === 0 || entriesB.length === 0}
              className="px-5 py-1.5 text-xs font-semibold text-white bg-gray-900 hover:bg-black disabled:opacity-40 disabled:cursor-not-allowed rounded-lg shadow-sm flex items-center gap-2 transition-colors cursor-pointer"
            >
              <Scale className="w-4 h-4" />
              <span>Reconcile Ledgers</span>
            </button>
          </div>
        </div>
      </div>

      {/* 6. Reconciliation Results Workspace */}
      {isCompared && comparisonResult && (
        <div className="space-y-4">
          {/* Executive Reconciliation Metrics Cards */}
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <span className="text-[11px] font-semibold text-gray-600 block">Match Rate</span>
              <span className="text-xl font-bold text-gray-900 mt-1 block">
                {comparisonResult.matchRate}%
              </span>
              <span className="text-[10px] text-gray-900 block mt-0.5">
                {comparisonResult.matchedPairs.length} confirmed • {comparisonResult.probablePairs.length} probable
              </span>
            </div>

            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <span className="text-[11px] font-semibold text-gray-600 block">Net Variance Discrepancy</span>
              <span className={`text-xl font-bold mt-1 block ${comparisonResult.balanceDifference <= 1 ? 'text-gray-900' : 'text-[#C43D43]'}`}>
                {formatINR(comparisonResult.balanceDifference)}
              </span>
              <span className="text-[10px] text-gray-600 block mt-0.5">Closing balance delta</span>
            </div>

            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <span className="text-[11px] font-semibold text-gray-600 block">Amount Discrepancies</span>
              <span className="text-xl font-bold text-[#B45309] mt-1 block">
                {comparisonResult.discrepancies.length}
              </span>
              <span className="text-[10px] text-[#B45309] block mt-0.5">Same voucher, amount mismatch</span>
            </div>

            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <span className="text-[11px] font-semibold text-gray-600 block">Unreconciled Entries</span>
              <span className="text-xl font-bold text-gray-900 mt-1 block">
                {comparisonResult.unmatchedA.length + comparisonResult.unmatchedB.length}
              </span>
              <span className="text-[10px] text-[#C43D43] block mt-0.5">
                {comparisonResult.unmatchedA.length} in A • {comparisonResult.unmatchedB.length} in B
              </span>
            </div>
          </div>

          {/* Detailed Tabbed View */}
          <div className="bg-white rounded-xl border border-gray-200 overflow-hidden shadow-sm">
            {/* Tabs */}
            <div className="px-5 py-3 bg-gray-50 border-b border-gray-200 flex flex-wrap items-center justify-between gap-3 text-xs">
              <div className="flex items-center gap-1.5">
                <button
                  onClick={() => setActiveTab('summary')}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors ${
                    activeTab === 'summary'
                      ? 'bg-white text-gray-900 border border-gray-200 shadow-sm'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Summary
                </button>
                <button
                  onClick={() => setActiveTab('confirmed')}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors ${
                    activeTab === 'confirmed'
                      ? 'bg-white text-gray-900 border border-gray-200 shadow-sm'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Confirmed Matches ({comparisonResult.matchedPairs.length})
                </button>
                <button
                  onClick={() => setActiveTab('probable')}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors ${
                    activeTab === 'probable'
                      ? 'bg-white text-[#B45309] border border-gray-200 shadow-sm'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Probable Matches ({comparisonResult.probablePairs.length})
                </button>
                <button
                  onClick={() => setActiveTab('discrepancies')}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors ${
                    activeTab === 'discrepancies'
                      ? 'bg-white text-[#C43D43] border border-gray-200 shadow-sm'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Amount Mismatches ({comparisonResult.discrepancies.length})
                </button>
                <button
                  onClick={() => setActiveTab('unmatchedA')}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors ${
                    activeTab === 'unmatchedA'
                      ? 'bg-white text-[#C43D43] border border-gray-200 shadow-sm'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Missing in B ({comparisonResult.unmatchedA.length})
                </button>
                <button
                  onClick={() => setActiveTab('unmatchedB')}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors ${
                    activeTab === 'unmatchedB'
                      ? 'bg-white text-[#C43D43] border border-gray-200 shadow-sm'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Missing in A ({comparisonResult.unmatchedB.length})
                </button>
              </div>
            </div>

            {/* Tab 1: Summary */}
            {activeTab === 'summary' && (
              <div className="p-5 space-y-4 text-xs">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div className="border border-gray-200 rounded-lg p-4 space-y-2">
                    <span className="font-bold text-gray-900 block pb-1 border-b border-gray-200">
                      Ledger A: {comparisonResult.ledgerAName}
                    </span>
                    <div className="flex justify-between py-1 text-gray-600">
                      <span>Total Debit Booked:</span>
                      <strong className="text-gray-900">{formatINR(comparisonResult.totalDebitA)}</strong>
                    </div>
                    <div className="flex justify-between py-1 text-gray-600">
                      <span>Total Credit Booked:</span>
                      <strong className="text-gray-900">{formatINR(comparisonResult.totalCreditA)}</strong>
                    </div>
                    <div className="flex justify-between py-1 text-gray-600">
                      <span>Net Balance:</span>
                      <strong className="text-gray-900 font-bold">{formatINR(comparisonResult.netBalanceA)}</strong>
                    </div>
                  </div>

                  <div className="border border-gray-200 rounded-lg p-4 space-y-2">
                    <span className="font-bold text-gray-900 block pb-1 border-b border-gray-200">
                      Ledger B: {comparisonResult.ledgerBName}
                    </span>
                    <div className="flex justify-between py-1 text-gray-600">
                      <span>Total Debit Booked:</span>
                      <strong className="text-gray-900">{formatINR(comparisonResult.totalDebitB)}</strong>
                    </div>
                    <div className="flex justify-between py-1 text-gray-600">
                      <span>Total Credit Booked:</span>
                      <strong className="text-gray-900">{formatINR(comparisonResult.totalCreditB)}</strong>
                    </div>
                    <div className="flex justify-between py-1 text-gray-600">
                      <span>Net Balance:</span>
                      <strong className="text-gray-900 font-bold">{formatINR(comparisonResult.netBalanceB)}</strong>
                    </div>
                  </div>
                </div>
              </div>
            )}

            {/* Tab 2: Confirmed Matches */}
            {activeTab === 'confirmed' && (
              <div className="overflow-x-auto max-h-[460px]">
                {comparisonResult.matchedPairs.length === 0 ? (
                  <div className="p-8 text-center text-xs text-gray-600">Zero confirmed matches.</div>
                ) : (
                  <table className="w-full text-left text-xs border-collapse">
                    <thead className="sticky top-0 bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold">
                      <tr>
                        <th className="py-2.5 px-3 border-r border-gray-200 w-12 text-center">#</th>
                        <th className="py-2.5 px-3 border-r border-gray-200">Date A</th>
                        <th className="py-2.5 px-3 border-r border-gray-200">Voucher A</th>
                        <th className="py-2.5 px-3 border-r border-gray-200">Amount A (₹)</th>
                        <th className="py-2.5 px-3 border-r border-gray-200">Date B</th>
                        <th className="py-2.5 px-3 border-r border-gray-200">Voucher B</th>
                        <th className="py-2.5 px-3 border-r border-gray-200">Amount B (₹)</th>
                        <th className="py-2.5 px-3">Match Observation</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB] text-gray-900">
                      {comparisonResult.matchedPairs.map((p, idx) => (
                        <tr key={idx} className="hover:bg-gray-50">
                          <td className="py-2 px-3 border-r border-gray-200 text-center text-gray-600 text-[11px]">
                            {idx + 1}
                          </td>
                          <td className="py-2 px-3 border-r border-gray-200 whitespace-nowrap">{p.entryA.date}</td>
                          <td className="py-2 px-3 border-r border-gray-200 font-mono font-semibold">{p.entryA.voucherNo}</td>
                          <td className="py-2 px-3 border-r border-gray-200 text-right font-mono font-bold text-gray-900">
                            {(p.entryA.debit || p.entryA.credit).toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                          </td>
                          <td className="py-2 px-3 border-r border-gray-200 whitespace-nowrap">{p.entryB.date}</td>
                          <td className="py-2 px-3 border-r border-gray-200 font-mono font-semibold">{p.entryB.voucherNo}</td>
                          <td className="py-2 px-3 border-r border-gray-200 text-right font-mono font-bold text-gray-900">
                            {(p.entryB.debit || p.entryB.credit).toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                          </td>
                          <td className="py-2 px-3 text-gray-900 font-medium">{p.notes}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
            )}

            {/* Tab 3: Probable Matches */}
            {activeTab === 'probable' && (
              <div className="overflow-x-auto max-h-[460px]">
                {comparisonResult.probablePairs.length === 0 ? (
                  <div className="p-8 text-center text-xs text-gray-600">Zero probable matches.</div>
                ) : (
                  <table className="w-full text-left text-xs border-collapse">
                    <thead className="sticky top-0 bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold">
                      <tr>
                        <th className="py-2.5 px-3 border-r border-gray-200 w-12 text-center">Score</th>
                        <th className="py-2.5 px-3 border-r border-gray-200">Date A</th>
                        <th className="py-2.5 px-3 border-r border-gray-200">Voucher A</th>
                        <th className="py-2.5 px-3 border-r border-gray-200">Amount (₹)</th>
                        <th className="py-2.5 px-3 border-r border-gray-200">Date B</th>
                        <th className="py-2.5 px-3 border-r border-gray-200">Voucher B</th>
                        <th className="py-2.5 px-3">Match Observation</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB] text-gray-900">
                      {comparisonResult.probablePairs.map((p, idx) => (
                        <tr key={idx} className="hover:bg-gray-50">
                          <td className="py-2 px-3 border-r border-gray-200 text-center font-bold text-[#B45309]">
                            {p.confidenceScore}%
                          </td>
                          <td className="py-2 px-3 border-r border-gray-200 whitespace-nowrap">{p.entryA.date}</td>
                          <td className="py-2 px-3 border-r border-gray-200 font-mono">{p.entryA.voucherNo}</td>
                          <td className="py-2 px-3 border-r border-gray-200 text-right font-mono font-bold">
                            {(p.entryA.debit || p.entryA.credit).toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                          </td>
                          <td className="py-2 px-3 border-r border-gray-200 whitespace-nowrap">{p.entryB.date}</td>
                          <td className="py-2 px-3 border-r border-gray-200 font-mono">{p.entryB.voucherNo}</td>
                          <td className="py-2 px-3 text-gray-600">{p.notes}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
            )}

            {/* Tab 4: Amount Discrepancies */}
            {activeTab === 'discrepancies' && (
              <div className="overflow-x-auto max-h-[460px]">
                {comparisonResult.discrepancies.length === 0 ? (
                  <div className="p-8 text-center text-xs text-gray-900 font-semibold flex items-center justify-center gap-2">
                    <CheckCircle2 className="w-5 h-5 text-gray-900" />
                    <span>Zero voucher amount discrepancies detected.</span>
                  </div>
                ) : (
                  <table className="w-full text-left text-xs border-collapse">
                    <thead className="sticky top-0 bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold">
                      <tr>
                        <th className="py-2.5 px-3 border-r border-gray-200 w-12 text-center">#</th>
                        <th className="py-2.5 px-3 border-r border-gray-200">Voucher No</th>
                        <th className="py-2.5 px-3 border-r border-gray-200 text-right">Ledger A Amount (₹)</th>
                        <th className="py-2.5 px-3 border-r border-gray-200 text-right">Ledger B Amount (₹)</th>
                        <th className="py-2.5 px-3 border-r border-gray-200 text-right text-[#C43D43]">Variance (₹)</th>
                        <th className="py-2.5 px-3">Audit Observation</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB] text-gray-900">
                      {comparisonResult.discrepancies.map((d, idx) => (
                        <tr key={idx} className="bg-[#FDF2F2] hover:bg-[#FBE8E8]">
                          <td className="py-2 px-3 border-r border-gray-200 text-center font-bold text-[#C43D43]">
                            {idx + 1}
                          </td>
                          <td className="py-2 px-3 border-r border-gray-200 font-mono font-bold">{d.entryA.voucherNo}</td>
                          <td className="py-2 px-3 border-r border-gray-200 text-right font-mono font-semibold">
                            {(d.entryA.debit || d.entryA.credit).toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                          </td>
                          <td className="py-2 px-3 border-r border-gray-200 text-right font-mono font-semibold">
                            {(d.entryB.debit || d.entryB.credit).toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                          </td>
                          <td className="py-2 px-3 border-r border-gray-200 text-right font-mono font-bold text-[#C43D43]">
                            {d.amountDiff.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                          </td>
                          <td className="py-2 px-3 text-[#991B1B] font-medium">{d.reason}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
            )}

            {/* Tab 5: Missing in Ledger B */}
            {activeTab === 'unmatchedA' && (
              <div className="overflow-x-auto max-h-[460px]">
                {comparisonResult.unmatchedA.length === 0 ? (
                  <div className="p-8 text-center text-xs text-gray-900 font-semibold">All Ledger A entries reconciled.</div>
                ) : (
                  <table className="w-full text-left text-xs border-collapse">
                    <thead className="sticky top-0 bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold">
                      <tr>
                        <th className="py-2.5 px-3 border-r border-gray-200 w-12 text-center">#</th>
                        <th className="py-2.5 px-3 border-r border-gray-200">Date</th>
                        <th className="py-2.5 px-3 border-r border-gray-200">Voucher No</th>
                        <th className="py-2.5 px-3 border-r border-gray-200">Particulars</th>
                        <th className="py-2.5 px-3 border-r border-gray-200 text-right">Debit (₹)</th>
                        <th className="py-2.5 px-3 border-r border-gray-200 text-right">Credit (₹)</th>
                        <th className="py-2.5 px-3">Audit Finding</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB] text-gray-900">
                      {comparisonResult.unmatchedA.map((a, idx) => (
                        <tr key={idx} className="hover:bg-gray-50">
                          <td className="py-2 px-3 border-r border-gray-200 text-center text-gray-600">{idx + 1}</td>
                          <td className="py-2 px-3 border-r border-gray-200 whitespace-nowrap">{a.date}</td>
                          <td className="py-2 px-3 border-r border-gray-200 font-mono">{a.voucherNo}</td>
                          <td className="py-2 px-3 border-r border-gray-200">{a.particulars}</td>
                          <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">{a.debit > 0 ? a.debit.toLocaleString('en-IN') : '-'}</td>
                          <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">{a.credit > 0 ? a.credit.toLocaleString('en-IN') : '-'}</td>
                          <td className="py-2 px-3 text-[#C43D43] font-medium">Recorded in our books, missing in counterparty</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
            )}

            {/* Tab 6: Missing in Ledger A */}
            {activeTab === 'unmatchedB' && (
              <div className="overflow-x-auto max-h-[460px]">
                {comparisonResult.unmatchedB.length === 0 ? (
                  <div className="p-8 text-center text-xs text-gray-900 font-semibold">All Ledger B entries reconciled.</div>
                ) : (
                  <table className="w-full text-left text-xs border-collapse">
                    <thead className="sticky top-0 bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold">
                      <tr>
                        <th className="py-2.5 px-3 border-r border-gray-200 w-12 text-center">#</th>
                        <th className="py-2.5 px-3 border-r border-gray-200">Date</th>
                        <th className="py-2.5 px-3 border-r border-gray-200">Voucher No</th>
                        <th className="py-2.5 px-3 border-r border-gray-200">Particulars</th>
                        <th className="py-2.5 px-3 border-r border-gray-200 text-right">Debit (₹)</th>
                        <th className="py-2.5 px-3 border-r border-gray-200 text-right">Credit (₹)</th>
                        <th className="py-2.5 px-3">Audit Finding</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB] text-gray-900">
                      {comparisonResult.unmatchedB.map((b, idx) => (
                        <tr key={idx} className="hover:bg-gray-50">
                          <td className="py-2 px-3 border-r border-gray-200 text-center text-gray-600">{idx + 1}</td>
                          <td className="py-2 px-3 border-r border-gray-200 whitespace-nowrap">{b.date}</td>
                          <td className="py-2 px-3 border-r border-gray-200 font-mono">{b.voucherNo}</td>
                          <td className="py-2 px-3 border-r border-gray-200">{b.particulars}</td>
                          <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">{b.debit > 0 ? b.debit.toLocaleString('en-IN') : '-'}</td>
                          <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">{b.credit > 0 ? b.credit.toLocaleString('en-IN') : '-'}</td>
                          <td className="py-2 px-3 text-[#C43D43] font-medium">Recorded by counterparty, missing in our books</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
};