import { UploadIllustration } from '../common/UploadIllustration';
import { InnerToolBadge } from '../common/InnerToolBadge';
import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  UploadCloud,
  Download,
  AlertTriangle,
  FileCog,
  Sliders,
  TrendingUp,
  TrendingDown,
  Scale,
  CheckCircle2,
  FileText,
  Layers,
  BookOpen,
  Check,
  X,
  FileCheck,
  ListChecks,
} from 'lucide-react';
import {
  detectBankColumnMapping,
  convertBankStatementRows,
  createConvertedStatementWorkbook,
  exportBankStatementCsv,
  type ColumnMapping,
  type BankStatementConversionResult,
  type DraftJournalEntry,
} from '../../utils/bankStatementConverterEngine';
import { parsePdfFile } from '../../utils/pdfExtractor';
import { ToolIcon } from '../common/ToolIcon';
import { DataFormatGuide } from '../common/DataFormatGuide';
import { formatINR, extractSmartTableFromSheet } from '../../utils/dataSanitizer';

interface BankStatementConverterToolProps {
  onBack: () => void;
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
}

export const BankStatementConverterTool: React.FC<BankStatementConverterToolProps> = ({ onBack, onShowToast }) => {
  const [file, setFile] = useState<{ name: string; size: string } | null>(null);
  const [ledgerFile, setLedgerFile] = useState<{ name: string; size: string } | null>(null);
  const [rawRows, setRawRows] = useState<Record<string, any>[]>([]);
  const [rawLedgerRows, setRawLedgerRows] = useState<Record<string, any>[]>([]);
  const [availableHeaders, setAvailableHeaders] = useState<string[]>([]);
  const [mapping, setMapping] = useState<ColumnMapping>({
    dateCol: '',
    narrationCol: '',
    refCol: '',
    debitCol: '',
    creditCol: '',
    balanceCol: '',
  });

  const [bankName, setBankName] = useState<string>('HDFC Bank Statement');
  const [conversionResult, setConversionResult] = useState<BankStatementConversionResult | null>(null);
  const [draftJournals, setDraftJournals] = useState<DraftJournalEntry[]>([]);
  const [activeTab, setActiveTab] = useState<'statement' | 'journals' | 'brs' | 'continuity'>('statement');
  const [isConverted, setIsConverted] = useState(false);
  const [validationError, setValidationError] = useState<string | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const ledgerInputRef = useRef<HTMLInputElement>(null);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;
    setValidationError(null);

    const fileNameLower = f.name.toLowerCase();
    const isPdf = fileNameLower.endsWith('.pdf');
    const isSheet = fileNameLower.endsWith('.xlsx') || fileNameLower.endsWith('.xls') || fileNameLower.endsWith('.csv');

    if (!isPdf && !isSheet) {
      setValidationError('Unsupported format. Please upload a PDF bank statement or Excel/CSV spreadsheet.');
      onShowToast('Invalid Format', 'Supported formats: .pdf, .xlsx, .xls, .csv', 'error');
      return;
    }

    setIsProcessing(true);

    try {
      if (isPdf) {
        const pdfRes = await parsePdfFile(f);
        if (pdfRes.status === 'error' || pdfRes.rows.length === 0) {
          throw new Error(pdfRes.errorMessage || 'Could not parse tabular rows from PDF bank statement.');
        }

        const headers = ['date', 'particulars', 'refNo', 'debit', 'credit', 'balance'];
        const autoMap: ColumnMapping = {
          dateCol: 'date',
          narrationCol: 'particulars',
          refCol: 'refNo',
          debitCol: 'debit',
          creditCol: 'credit',
          balanceCol: 'balance',
        };

        setFile({ name: f.name, size: `${Math.round(f.size / 1024)} KB` });
        setRawRows(pdfRes.rows);
        setAvailableHeaders(headers);
        setMapping(autoMap);
        const derivedBank = f.name.replace(/\.[^/.]+$/, '');
        setBankName(derivedBank);

        const res = convertBankStatementRows(pdfRes.rows, autoMap, derivedBank, rawLedgerRows);
        setConversionResult(res);
        setDraftJournals(res.draftJournals);
        setIsConverted(true);
      } else {
        const reader = new FileReader();
        reader.onload = (evt) => {
          try {
            const buffer = evt.target?.result as ArrayBuffer;
            const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
            const smartTable = extractSmartTableFromSheet(wb.Sheets[wb.SheetNames[0]]);
            const json = smartTable.rows;

            if (!json || json.length === 0) throw new Error('Spreadsheet contains zero rows.');

            const headers = smartTable.headers.length > 0 ? smartTable.headers : Object.keys(json[0]);
            const detected = detectBankColumnMapping(headers);
            const derivedBank = f.name.replace(/\.[^/.]+$/, '');

            setFile({ name: f.name, size: `${Math.round(f.size / 1024)} KB` });
            setRawRows(json);
            setAvailableHeaders(headers);
            setMapping(detected);
            setBankName(derivedBank);

            const res = convertBankStatementRows(json, detected, derivedBank, rawLedgerRows);
            setConversionResult(res);
            setDraftJournals(res.draftJournals);
            setIsConverted(true);
          } catch (err: any) {
            setValidationError(`Failed to parse file: ${err.message}`);
            onShowToast('Parse Error', err.message, 'error');
          }
        };
        reader.readAsArrayBuffer(f);
      }
    } catch (err: any) {
      setValidationError(`Error reading bank statement: ${err.message}`);
      onShowToast('Statement Error', err.message, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  const handleLedgerUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;

    try {
      const buffer = await f.arrayBuffer();
      const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
      const smartTable = extractSmartTableFromSheet(wb.Sheets[wb.SheetNames[0]]);
      const json = smartTable.rows;

      if (!json || json.length === 0) throw new Error('Ledger file has no readable rows.');

      setLedgerFile({ name: f.name, size: `${Math.round(f.size / 1024)} KB` });
      setRawLedgerRows(json);

      // Re-run conversion with ledger
      if (rawRows.length > 0) {
        const res = convertBankStatementRows(rawRows, mapping, bankName, json);
        setConversionResult(res);
        setDraftJournals(res.draftJournals);
      }
    } catch (err: any) {
      onShowToast('Ledger Error', err.message, 'error');
    }
  };

  
  const handleRecalculateMapping = () => {
    if (rawRows.length === 0) return;
    const res = convertBankStatementRows(rawRows, mapping, bankName, rawLedgerRows);
    setConversionResult(res);
    setDraftJournals(res.draftJournals);
    onShowToast('Mappings Updated', 'Recalculated statement and draft journal entries.', 'success');
  };

  const handleToggleJournalApproval = (id: string) => {
    setDraftJournals((prev) =>
      prev.map((j) => (j.id === id ? { ...j, isApproved: !j.isApproved, status: !j.isApproved ? 'Approved' : 'Pending Review' } : j))
    );
  };

  const handleApproveAllJournals = () => {
    setDraftJournals((prev) =>
      prev.map((j) => ({ ...j, isApproved: true, status: 'Approved' }))
    );
    onShowToast('All Entries Approved', 'Approved all draft double entries for accounting posting.', 'success');
  };

  const handleDownloadExcel = () => {
    if (!conversionResult) return;
    // Embed latest approved draft journals
    const updatedResult = { ...conversionResult, draftJournals };
    const wb = createConvertedStatementWorkbook(updatedResult, rawRows);
    XLSX.writeFile(wb, `${bankName.replace(/\s+/g, '_')}_Reconciliation_FY26.xlsx`);
    onShowToast('5-Sheet Excel Exported', 'Downloaded complete BRS, draft journals, and continuity audit.', 'success');
  };

  const handleDownloadCsv = () => {
    if (!conversionResult) return;
    const csvData = exportBankStatementCsv(conversionResult);
    const blob = new Blob([csvData], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', `${bankName.replace(/\s+/g, '_')}_Statement.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    onShowToast('CSV Exported', 'Clean statement CSV downloaded.', 'success');
  };

  return (
    <div className="space-y-6">
      {/* 1. Header */}
      <div>
        <button
          onClick={onBack}
          className="inline-flex items-center gap-2 px-3 py-1.5 rounded-lg text-xs font-semibold text-gray-700 bg-white border border-gray-200 shadow-2xs hover:bg-gray-50 hover:text-gray-900 hover:border-gray-300 transition-all cursor-pointer mb-5 group"
        >
          <ArrowLeft className="w-3.5 h-3.5 text-gray-400 group-hover:text-gray-900 group-hover:-translate-x-0.5 transition-transform" />
          <span>Back to All Tools</span>
        </button>

        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-5 border-b border-gray-200">
          <div className="flex items-center gap-3.5">
            <InnerToolBadge toolId="bank-converter" />
            <div>
              <h1 className="text-xl font-bold text-gray-900 tracking-tight">Bank Statement-to-Ledger Automation & BRS</h1>
              <p className="text-xs text-gray-500 mt-0.5">
                Convert statement formats, audit running balance continuity, suggest double-entry accounts, and reconcile with accounting ledger.
              </p>
            </div>
          </div>

          {conversionResult && (
            <div className="flex items-center gap-2">
              <button
                onClick={handleDownloadCsv}
                className="px-3.5 py-2 text-xs font-semibold text-gray-700 bg-white border border-gray-300 hover:bg-gray-50 rounded-lg shadow-sm flex items-center gap-1.5 transition-colors"
                title="Export as CSV"
              >
                <FileText className="w-3.5 h-3.5 text-gray-500" />
                <span>Export CSV</span>
              </button>
              <button
                onClick={handleDownloadExcel}
                className="px-4 py-2 text-xs font-semibold text-white bg-gray-900 hover:bg-black rounded-lg shadow-sm flex items-center gap-2 transition-colors cursor-pointer"
                title="Export 5-Sheet Excel Workbook"
              >
                <Download className="w-4 h-4" />
                <span>Export 5-Sheet Excel (.xlsx)</span>
              </button>
            </div>
          )}
        </div>
      </div>

      {/* 2. Format Requirements Guide */}
      <DataFormatGuide
        toolType="bank-rec"
        acceptedFormats=".pdf, .xlsx, .xls, .csv"
        requiredFields={[
          { name: 'Transaction Date', description: 'Posting date in standard or Indian format', aliases: ['Date', 'Txn Date', 'Value Date'] },
          { name: 'Narration / Description', description: 'Transaction particulars, beneficiary, or remarks', aliases: ['Narration', 'Description', 'Particulars'] },
          { name: 'Withdrawal / Debit', description: 'Outgoing payments / debits in bank', aliases: ['Withdrawal', 'Debit', 'Dr', 'Payment'] },
          { name: 'Deposit / Credit', description: 'Incoming receipts / credits in bank', aliases: ['Deposit', 'Credit', 'Cr', 'Receipt'] },
          { name: 'Running Balance', description: 'Bank balance after each transaction', aliases: ['Balance', 'Closing Bal', 'Net Balance'] },
        ]}
        notes={[
          'Running Balance Continuity Check: Automatically flags any row where Balance[i-1] - Withdrawal[i] + Deposit[i] != Balance[i].',
          'Account Head Suggester: Categorizes transactions into GST, TDS, Salaries, Rent, Utilities, Bank Charges, and Trade Payables.',
          'Draft Journal Generator: Prepares double-entry journals with configurable review and user approval before posting.',
          'Side-by-Side Ledger Comparison: Upload existing accounting bank book to compute statutory Bank Reconciliation Statement (BRS).',
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

      {/* 4. Dual Upload Area: Bank Statement & Accounting Ledger */}
      <div className="space-y-4">
        <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
          {/* Box 1: Bank Statement */}
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileUpload}
            accept=".pdf,.xlsx,.xls,.csv"
            className="hidden"
            id="statement-converter-upload"
          />
          {!file ? (
            <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-6 sm:p-8 text-center transition-colors shadow-2xs">
              <label htmlFor="statement-converter-upload" className="cursor-pointer flex flex-col items-center">
                <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mb-3 shadow-2xs">
                  <UploadCloud className="w-6 h-6" />
                </div>
                <h3 className="text-sm font-semibold text-gray-900">1. Bank Statement</h3>
                <p className="text-xs text-gray-500 mt-1 max-w-xs mx-auto">
                  Upload bank statement in PDF, .xlsx, .xls, or .csv
                </p>
                <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
                  <UploadCloud className="w-4 h-4" />
                  <span>{isProcessing ? 'Processing...' : 'Upload Bank Statement'}</span>
                </div>
                <span className="text-[11px] text-gray-400 font-medium mt-3 block">
                  PDF / Excel / CSV · 100% Client-side
                </span>
              </label>
            </div>
          ) : (
            <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-sm flex flex-col justify-between">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-lg bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 shrink-0">
                  <UploadCloud className="w-5 h-5" />
                </div>
                <div className="min-w-0 flex-1">
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-bold text-gray-900 block truncate">{file.name}</span>
                    <span className="text-[10px] font-semibold text-gray-900 bg-gray-100 border border-gray-200 px-2 py-0.5 rounded">
                      Ready
                    </span>
                  </div>
                  <span className="text-[11px] text-gray-600 block mt-0.5">
                    Bank Statement • {rawRows.length} transactions loaded • {file.size}
                  </span>
                </div>
              </div>
              <div className="mt-4 pt-3 border-t border-gray-100 flex items-center justify-between">
                <span className="text-[11px] text-gray-500 font-medium">Ready for conversion</span>
                <button
                  type="button"
                  onClick={() => fileInputRef.current?.click()}
                  disabled={isProcessing}
                  className="cursor-pointer px-3 py-1.5 text-xs text-gray-900 bg-gray-50 hover:bg-[#EAEAEA] rounded-lg border border-gray-200 flex items-center gap-1.5 font-medium transition-colors"
                >
                  <UploadCloud className="w-3.5 h-3.5 text-gray-600" />
                  <span>Change Statement</span>
                </button>
              </div>
            </div>
          )}

          {/* Box 2: Accounting Ledger (Bank Book) */}
          <input
            type="file"
            ref={ledgerInputRef}
            onChange={handleLedgerUpload}
            accept=".xlsx,.xls,.csv"
            className="hidden"
            id="statement-converter-ledger-upload"
          />
          {!ledgerFile ? (
            <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-6 sm:p-8 text-center transition-colors shadow-2xs">
              <label htmlFor="statement-converter-ledger-upload" className="cursor-pointer flex flex-col items-center">
                <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mb-3 shadow-2xs">
                  <BookOpen className="w-6 h-6" />
                </div>
                <h3 className="text-sm font-semibold text-gray-900">2. Accounting Ledger (Bank Book)</h3>
                <p className="text-xs text-gray-500 mt-1 max-w-xs mx-auto">
                  Optional: Upload Bank Ledger from Tally/SAP for instant BRS
                </p>
                <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
                  <UploadCloud className="w-4 h-4" />
                  <span>Upload Accounting Ledger</span>
                </div>
                <span className="text-[11px] text-gray-400 font-medium mt-3 block">
                  Tally / SAP / ERP export (.xlsx, .csv)
                </span>
              </label>
            </div>
          ) : (
            <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-sm flex flex-col justify-between">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-lg bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 shrink-0">
                  <BookOpen className="w-5 h-5" />
                </div>
                <div className="min-w-0 flex-1">
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-bold text-gray-900 block truncate">{ledgerFile.name}</span>
                    <span className="text-[10px] font-semibold text-gray-900 bg-gray-100 border border-gray-200 px-2 py-0.5 rounded">
                      Ready
                    </span>
                  </div>
                  <span className="text-[11px] text-gray-600 block mt-0.5">
                    Accounting Ledger • {rawLedgerRows.length} vouchers • {ledgerFile.size}
                  </span>
                </div>
              </div>
              <div className="mt-4 pt-3 border-t border-gray-100 flex items-center justify-between">
                <span className="text-[11px] text-gray-500 font-medium">Ready for reconciliation</span>
                <button
                  type="button"
                  onClick={() => ledgerInputRef.current?.click()}
                  className="cursor-pointer px-3 py-1.5 text-xs text-gray-900 bg-gray-50 hover:bg-[#EAEAEA] rounded-lg border border-gray-200 flex items-center gap-1.5 font-medium transition-colors"
                >
                  <UploadCloud className="w-3.5 h-3.5 text-gray-600" />
                  <span>Change Ledger</span>
                </button>
              </div>
            </div>
          )}
        </div>

        {/* Column Mapping Grid */}
        {availableHeaders.length > 0 && (
          <div className="pt-4 border-t border-gray-200 space-y-3">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Sliders className="w-4 h-4 text-gray-900" />
                <span className="text-xs font-bold text-gray-900">Review & Confirm Column Mappings</span>
                <span className="text-[11px] text-gray-600">(Verify detected columns before final conversion)</span>
              </div>
              <button
                onClick={handleRecalculateMapping}
                className="px-3 py-1 text-xs font-semibold text-white bg-gray-900 hover:bg-black rounded-md transition-colors"
              >
                Apply Mapping Changes
              </button>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3 text-xs">
              <div>
                <label className="block text-[11px] font-semibold text-gray-600 mb-1">Date Column</label>
                <select
                  value={mapping.dateCol}
                  onChange={(e) => setMapping({ ...mapping, dateCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded-md text-xs font-medium text-gray-900 focus:outline-none focus:border-gray-300"
                >
                  <option value="">(None)</option>
                  {availableHeaders.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-[11px] font-semibold text-gray-600 mb-1">Narration Column</label>
                <select
                  value={mapping.narrationCol}
                  onChange={(e) => setMapping({ ...mapping, narrationCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded-md text-xs font-medium text-gray-900 focus:outline-none focus:border-gray-300"
                >
                  <option value="">(None)</option>
                  {availableHeaders.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-[11px] font-semibold text-gray-600 mb-1">Ref / Chq Column</label>
                <select
                  value={mapping.refCol}
                  onChange={(e) => setMapping({ ...mapping, refCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded-md text-xs font-medium text-gray-900 focus:outline-none focus:border-gray-300"
                >
                  <option value="">(None)</option>
                  {availableHeaders.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-[11px] font-semibold text-gray-600 mb-1">Debit / Wdl Column</label>
                <select
                  value={mapping.debitCol}
                  onChange={(e) => setMapping({ ...mapping, debitCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded-md text-xs font-medium text-gray-900 focus:outline-none focus:border-gray-300"
                >
                  <option value="">(None)</option>
                  {availableHeaders.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-[11px] font-semibold text-gray-600 mb-1">Credit / Dep Column</label>
                <select
                  value={mapping.creditCol}
                  onChange={(e) => setMapping({ ...mapping, creditCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded-md text-xs font-medium text-gray-900 focus:outline-none focus:border-gray-300"
                >
                  <option value="">(None)</option>
                  {availableHeaders.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-[11px] font-semibold text-gray-600 mb-1">Closing Bal Column</label>
                <select
                  value={mapping.balanceCol}
                  onChange={(e) => setMapping({ ...mapping, balanceCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded-md text-xs font-medium text-gray-900 focus:outline-none focus:border-gray-300"
                >
                  <option value="">(None)</option>
                  {availableHeaders.map((h) => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* 5. Converted Results Workspace */}
      {isConverted && conversionResult && (
        <div className="space-y-4">
          {/* Continuity Warning Banner if breaks found */}
          {conversionResult.continuityAudit.hasBreaks && (
            <div className="p-4 rounded-xl bg-amber-50 border border-amber-200 text-xs text-amber-900 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0" />
                <span>
                  <strong>Running Balance Break Detected:</strong> Found {conversionResult.continuityAudit.breakCount} continuity break(s) totalling {formatINR(conversionResult.continuityAudit.totalDiscrepancy)}. Review details in Continuity Audit tab.
                </span>
              </div>
              <button
                onClick={() => setActiveTab('continuity')}
                className="px-2.5 py-1 text-xs font-semibold bg-white border border-amber-300 rounded text-amber-800 hover:bg-amber-100 transition-colors"
              >
                Inspect Breaks
              </button>
            </div>
          )}

          {/* Executive Cash Flow Summary Cards */}
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <div className="flex items-center gap-1.5 text-gray-900 mb-1">
                <TrendingUp className="w-4 h-4" />
                <span className="text-[11px] font-semibold">Total Deposits (Cr)</span>
              </div>
              <span className="text-lg font-bold text-gray-900 block">
                {formatINR(conversionResult.totalDeposits)}
              </span>
              <span className="text-[10px] text-gray-600 block mt-0.5">Total inflows</span>
            </div>

            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <div className="flex items-center gap-1.5 text-[#C43D43] mb-1">
                <TrendingDown className="w-4 h-4" />
                <span className="text-[11px] font-semibold">Total Withdrawals (Dr)</span>
              </div>
              <span className="text-lg font-bold text-[#C43D43] block">
                {formatINR(conversionResult.totalWithdrawals)}
              </span>
              <span className="text-[10px] text-gray-600 block mt-0.5">Total outflows</span>
            </div>

            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <div className="flex items-center gap-1.5 text-gray-900 mb-1">
                <Scale className="w-4 h-4 text-gray-900" />
                <span className="text-[11px] font-semibold">Net Cash Flow</span>
              </div>
              <span className={`text-lg font-bold block ${conversionResult.netCashFlow >= 0 ? 'text-gray-900' : 'text-[#C43D43]'}`}>
                {formatINR(conversionResult.netCashFlow)}
              </span>
              <span className="text-[10px] text-gray-600 block mt-0.5">Period Net Position</span>
            </div>

            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <span className="text-[11px] font-semibold text-gray-600 block">Closing Statement Bal</span>
              <span className="text-lg font-bold text-gray-900 mt-1 block">
                {formatINR(conversionResult.closingBalance)}
              </span>
              <span className="text-[10px] text-gray-900 font-medium block mt-0.5">
                {conversionResult.totalTransactions} transactions standardized
              </span>
            </div>
          </div>

          {/* Interactive Navigation Tabs */}
          <div className="flex items-center gap-2 border-b border-gray-200 pb-1 text-xs">
            <button
              onClick={() => setActiveTab('statement')}
              className={`px-3 py-1.5 font-semibold rounded-t-lg transition-colors flex items-center gap-1.5 ${
                activeTab === 'statement'
                  ? 'bg-white text-gray-900 border-t-2 border-gray-300 shadow-xs'
                  : 'text-gray-600 hover:text-gray-900'
              }`}
            >
              <FileCheck className="w-3.5 h-3.5" />
              <span>Clean Statement ({conversionResult.transactions.length})</span>
            </button>

            <button
              onClick={() => setActiveTab('journals')}
              className={`px-3 py-1.5 font-semibold rounded-t-lg transition-colors flex items-center gap-1.5 ${
                activeTab === 'journals'
                  ? 'bg-white text-gray-900 border-t-2 border-gray-300 shadow-xs'
                  : 'text-gray-600 hover:text-gray-900'
              }`}
            >
              <ListChecks className="w-3.5 h-3.5" />
              <span>Draft Journal Entries ({draftJournals.length})</span>
            </button>

            <button
              onClick={() => setActiveTab('brs')}
              className={`px-3 py-1.5 font-semibold rounded-t-lg transition-colors flex items-center gap-1.5 ${
                activeTab === 'brs'
                  ? 'bg-white text-gray-900 border-t-2 border-gray-300 shadow-xs'
                  : 'text-gray-600 hover:text-gray-900'
              }`}
            >
              <Scale className="w-3.5 h-3.5" />
              <span>Reconciliation (BRS)</span>
            </button>

            <button
              onClick={() => setActiveTab('continuity')}
              className={`px-3 py-1.5 font-semibold rounded-t-lg transition-colors flex items-center gap-1.5 ${
                activeTab === 'continuity'
                  ? 'bg-white text-gray-900 border-t-2 border-gray-300 shadow-xs'
                  : 'text-gray-600 hover:text-gray-900'
              }`}
            >
              <AlertTriangle className={`w-3.5 h-3.5 ${conversionResult.continuityAudit.hasBreaks ? 'text-amber-600' : 'text-gray-400'}`} />
              <span>Continuity Audit ({conversionResult.continuityAudit.breakCount} issues)</span>
            </button>
          </div>

          {/* TAB 1: Clean Standardized Statement */}
          {activeTab === 'statement' && (
            <div className="bg-white rounded-xl border border-gray-200 overflow-hidden shadow-sm">
              <div className="px-5 py-3 bg-gray-50 border-b border-gray-200 flex items-center justify-between text-xs">
                <div className="flex items-center gap-2">
                  <span className="font-bold text-gray-900">Standardized Bank Transactions</span>
                  <span className="text-gray-600">({conversionResult.transactions.length} records ready for BRS)</span>
                </div>
                <span className="text-[11px] text-gray-900 font-semibold bg-slate-100 px-2 py-0.5 rounded border border-slate-200">
                  Dates standardized to DD-MM-YYYY
                </span>
              </div>

              <div className="overflow-x-auto max-h-[460px]">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="sticky top-0 bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold">
                    <tr>
                      <th className="py-2.5 px-3 border-r border-gray-200 w-12 text-center">#</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Date</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Particulars / Narration</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Chq / Ref No</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap text-right">Withdrawal (Dr ₹)</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap text-right">Deposit (Cr ₹)</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap text-right">Closing Balance (₹)</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Suggested Account Head</th>
                      <th className="py-2.5 px-3 whitespace-nowrap">Match Status</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#E5E7EB] text-gray-900">
                    {conversionResult.transactions.map((t, idx) => (
                      <tr key={t.id} className={`hover:bg-gray-50 transition-colors ${t.continuityBreak ? 'bg-amber-50/50' : ''}`}>
                        <td className="py-2 px-3 border-r border-gray-200 text-center text-gray-600 text-[11px]">
                          {idx + 1}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 whitespace-nowrap font-medium">
                          {t.date}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 font-mono text-[11px] max-w-[240px] truncate" title={t.narration}>
                          {t.narration}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 whitespace-nowrap text-gray-600">
                          {t.refNo}
                        </td>
                        <td className={`py-2 px-3 border-r border-gray-200 whitespace-nowrap text-right font-mono ${t.withdrawal > 0 ? 'text-[#C43D43] font-semibold' : 'text-gray-600'}`}>
                          {t.withdrawal > 0 ? t.withdrawal.toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '-'}
                        </td>
                        <td className={`py-2 px-3 border-r border-gray-200 whitespace-nowrap text-right font-mono ${t.deposit > 0 ? 'text-gray-900 font-semibold' : 'text-gray-600'}`}>
                          {t.deposit > 0 ? t.deposit.toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '-'}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 whitespace-nowrap text-right font-mono font-medium">
                          {t.balance > 0 ? t.balance.toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '-'}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 text-[11px] max-w-[200px] truncate" title={t.suggestedAccountHead}>
                          <span className="font-semibold text-gray-800">{t.suggestedAccountHead}</span>
                        </td>
                        <td className="py-2 px-3 whitespace-nowrap">
                          {t.matchStatus === 'Matched' ? (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-semibold bg-slate-100 text-gray-900 border border-slate-200">
                              <CheckCircle2 className="w-3 h-3" />
                              <span>Matched</span>
                            </span>
                          ) : (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-medium bg-gray-100 text-gray-600 border border-gray-200">
                              <span>Unmatched</span>
                            </span>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* TAB 2: Draft Journal Entries */}
          {activeTab === 'journals' && (
            <div className="bg-white rounded-xl border border-gray-200 overflow-hidden shadow-sm space-y-4">
              <div className="px-5 py-3 bg-gray-50 border-b border-gray-200 flex items-center justify-between text-xs">
                <div>
                  <span className="font-bold text-gray-900">Draft Double-Entry Journal Generator</span>
                  <span className="text-gray-600 ml-2">Review, edit account heads, and approve before accounting posting</span>
                </div>
                <button
                  onClick={handleApproveAllJournals}
                  className="px-3 py-1.5 text-xs font-semibold text-white bg-gray-900 hover:bg-black rounded-md transition-colors shadow-2xs flex items-center gap-1.5"
                >
                  <Check className="w-3.5 h-3.5" />
                  <span>Approve All Draft Entries</span>
                </button>
              </div>

              <div className="overflow-x-auto max-h-[460px]">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="sticky top-0 bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold">
                    <tr>
                      <th className="py-2.5 px-3 border-r border-gray-200 w-10 text-center">#</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Date</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Debit Account (Dr)</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Credit Account (Cr)</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap text-right">Amount (₹)</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Narration</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Classification Basis</th>
                      <th className="py-2.5 px-3 whitespace-nowrap text-center">Action</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#E5E7EB] text-gray-900">
                    {draftJournals.map((j, idx) => (
                      <tr key={j.id} className={`hover:bg-gray-50 transition-colors ${j.isApproved ? 'bg-gray-50/30' : ''}`}>
                        <td className="py-2 px-3 border-r border-gray-200 text-center text-gray-600 text-[11px]">
                          {idx + 1}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 whitespace-nowrap font-medium">
                          {j.date}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 font-semibold text-gray-900 max-w-[180px] truncate" title={j.debitAccount}>
                          {j.debitAccount}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 font-semibold text-gray-800 max-w-[180px] truncate" title={j.creditAccount}>
                          {j.creditAccount}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 text-right font-mono font-bold text-gray-900">
                          {j.amount.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 text-[11px] text-gray-600 max-w-[200px] truncate" title={j.narration}>
                          {j.narration}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 text-[11px] text-gray-600 max-w-[180px] truncate" title={j.suggestedRule}>
                          {j.suggestedRule}
                        </td>
                        <td className="py-2 px-3 text-center whitespace-nowrap">
                          <button
                            onClick={() => handleToggleJournalApproval(j.id)}
                            className={`px-2.5 py-1 rounded text-[10px] font-semibold transition-colors ${
                              j.isApproved
                                ? 'bg-slate-100 text-gray-900 border border-slate-200'
                                : 'bg-gray-100 text-gray-700 hover:bg-gray-200 border border-gray-300'
                            }`}
                          >
                            {j.isApproved ? 'Approved ✓' : 'Approve'}
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* TAB 3: BRS & Ledger Match */}
          {activeTab === 'brs' && (
            <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-sm space-y-5">
              <div>
                <h3 className="text-sm font-bold text-gray-900">Statutory Bank Reconciliation Statement (BRS)</h3>
                <p className="text-xs text-gray-600 mt-0.5">
                  Comparison between Bank Statement closing balance and Client Accounting Ledger.
                </p>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="space-y-3">
                  <div className="p-4 rounded-xl bg-gray-50 border border-gray-200 space-y-2 text-xs">
                    <span className="text-[11px] font-bold text-gray-700 block">Bank Reconciliation Summary</span>
                    <div className="flex justify-between py-1 border-b border-gray-200">
                      <span className="text-gray-600">Balance as per Bank Statement</span>
                      <span className="font-mono font-bold text-gray-900">{formatINR(conversionResult.closingBalance)}</span>
                    </div>
                    <div className="flex justify-between py-1 border-b border-gray-200 text-gray-900">
                      <span>Add: Cheques Issued in Books but Not Presented in Bank</span>
                      <span className="font-mono font-bold">{formatINR(conversionResult.ledgerReconciliation?.unpresentedCheques || 0)}</span>
                    </div>
                    <div className="flex justify-between py-1 border-b border-gray-200 text-[#C43D43]">
                      <span>Less: Cheques Deposited in Books but Not Credited by Bank</span>
                      <span className="font-mono font-bold">{formatINR(conversionResult.ledgerReconciliation?.uncreditedDeposits || 0)}</span>
                    </div>
                    <div className="flex justify-between py-1.5 font-bold text-sm text-gray-900 bg-white p-2 rounded border border-gray-200">
                      <span>Adjusted Reconciled Balance</span>
                      <span className="font-mono text-gray-900">
                        {formatINR(
                          conversionResult.closingBalance +
                            (conversionResult.ledgerReconciliation?.unpresentedCheques || 0) -
                            (conversionResult.ledgerReconciliation?.uncreditedDeposits || 0)
                        )}
                      </span>
                    </div>
                  </div>
                </div>

                <div className="space-y-3 text-xs">
                  <div className="p-4 rounded-xl bg-gray-50 border border-gray-200 space-y-2">
                    <span className="text-[11px] font-bold text-gray-700 block">Matching Statistics</span>
                    <div className="flex justify-between py-1 border-b border-gray-200">
                      <span className="text-gray-600">Matched Bank Transactions</span>
                      <span className="font-bold text-gray-900">{conversionResult.ledgerReconciliation?.matchedCount || 0} entries</span>
                    </div>
                    <div className="flex justify-between py-1 border-b border-gray-200">
                      <span className="text-gray-600">Unmatched Bank Transactions (Pending Vouchers)</span>
                      <span className="font-bold text-amber-700">{conversionResult.ledgerReconciliation?.unmatchedBankCount || conversionResult.totalTransactions} entries</span>
                    </div>
                    <div className="flex justify-between py-1 border-b border-gray-200">
                      <span className="text-gray-600">Unmatched Ledger Transactions (Timing Differences)</span>
                      <span className="font-bold text-gray-800">{conversionResult.ledgerReconciliation?.unmatchedLedgerCount || 0} entries</span>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* TAB 4: Continuity Audit */}
          {activeTab === 'continuity' && (
            <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-sm space-y-4">
              <div className="flex items-center justify-between pb-3 border-b border-gray-200">
                <div>
                  <h3 className="text-sm font-bold text-gray-900">Mathematical Running Balance Continuity Audit</h3>
                  <p className="text-xs text-gray-600 mt-0.5">
                    Validates: Balance[i-1] - Withdrawal[i] + Deposit[i] = Balance[i] across every sequential entry.
                  </p>
                </div>
                <span
                  className={`px-3 py-1 rounded text-xs font-semibold ${
                    conversionResult.continuityAudit.hasBreaks
                      ? 'bg-amber-100 text-amber-800 border border-amber-300'
                      : 'bg-slate-100 text-gray-900 border border-slate-200'
                  }`}
                >
                  {conversionResult.continuityAudit.hasBreaks
                    ? `${conversionResult.continuityAudit.breakCount} Breaks Detected`
                    : '100% Continuous & Mathematically Intact'}
                </span>
              </div>

              {conversionResult.continuityAudit.breaks.length > 0 ? (
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-xs border-collapse">
                    <thead className="bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold">
                      <tr>
                        <th className="py-2.5 px-3 border-r border-gray-200">Row</th>
                        <th className="py-2.5 px-3 border-r border-gray-200">Date</th>
                        <th className="py-2.5 px-3 border-r border-gray-200">Particulars</th>
                        <th className="py-2.5 px-3 border-r border-gray-200 text-right">Prev Balance (₹)</th>
                        <th className="py-2.5 px-3 border-r border-gray-200 text-right">Dr / Cr (₹)</th>
                        <th className="py-2.5 px-3 border-r border-gray-200 text-right">Expected Bal (₹)</th>
                        <th className="py-2.5 px-3 border-r border-gray-200 text-right">Actual Bal (₹)</th>
                        <th className="py-2.5 px-3 text-right font-bold text-amber-700">Variance (₹)</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB]">
                      {conversionResult.continuityAudit.breaks.map((b, idx) => (
                        <tr key={idx} className="bg-amber-50/50">
                          <td className="py-2 px-3 border-r border-gray-200 text-center font-mono">{b.row}</td>
                          <td className="py-2 px-3 border-r border-gray-200 whitespace-nowrap">{b.date}</td>
                          <td className="py-2 px-3 border-r border-gray-200 font-mono text-[11px] max-w-[200px] truncate">{b.narration}</td>
                          <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">{formatINR(b.prevBalance)}</td>
                          <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">{b.withdrawal > 0 ? `-${formatINR(b.withdrawal)}` : `+${formatINR(b.deposit)}`}</td>
                          <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">{formatINR(b.expectedBalance)}</td>
                          <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">{formatINR(b.actualBalance)}</td>
                          <td className="py-2 px-3 text-right font-mono font-bold text-amber-700">{formatINR(b.diff)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ) : (
                <div className="p-8 text-center bg-gray-50 rounded-xl border border-gray-200 space-y-2">
                  <CheckCircle2 className="w-8 h-8 text-gray-900 mx-auto" />
                  <h4 className="text-sm font-bold text-gray-900">Zero Balance Gaps Detected</h4>
                  <p className="text-xs text-gray-500 max-w-md mx-auto">
                    Every transaction's debit and credit strictly aligns with the reported running balance from opening to closing balance.
                  </p>
                </div>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
};
