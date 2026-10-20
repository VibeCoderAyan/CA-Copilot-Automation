import React, { useState, useRef } from 'react';
import {
  ArrowLeft,
  CheckSquare,
  Sliders,
  AlertOctagon,
  Download,
  UploadCloud,
  CheckCircle2,
  AlertTriangle,
  ArrowRight,
  RefreshCw,
  ShieldCheck,
  FileSpreadsheet,
  Info,
  XCircle,
  FileCode,
} from 'lucide-react';
import { INITIAL_AUDIT_CHECKLIST } from '../../data/mockData';
import type { AuditChecklistItem } from '../../types';
import {
  parseAuditTransactions,
  sampleTransactions,
  detectUnusualTransactions,
  createAuditWorkingPaper,
  type AuditTransaction,
  type FlaggedTransaction,
} from '../../utils/auditEngine';
import { parsePdfFile } from '../../utils/pdfExtractor';
import { exportStyledWorkbook } from '../../utils/excelStyler';
import { DataFormatGuide } from '../common/DataFormatGuide';
import { formatINR } from '../../utils/dataSanitizer';
import * as XLSX from 'xlsx';
import { ToolHeader } from '../common/ToolHeader';

interface AuditToolsToolProps {
  onBack: () => void;
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
}

export const AuditToolsTool: React.FC<AuditToolsToolProps> = ({ onBack, onShowToast }) => {
  const [selectedUtility, setSelectedUtility] = useState<'checklist' | 'sampling' | 'unusual' | null>(null);

  // Utility 1: Checklist State (Empty by default - loaded on demand)
  const [checklist, setChecklist] = useState<AuditChecklistItem[]>([]);
  const [checklistFilter, setChecklistFilter] = useState<'all' | 'verified' | 'progress'>('all');

  // File & Data State (Zero fake data by default)
  const [file, setFile] = useState<{ name: string; size: string; count: number } | null>(null);
  const [transactions, setTransactions] = useState<AuditTransaction[]>([]);
  const [sampledTxns, setSampledTxns] = useState<AuditTransaction[]>([]);
  const [unusualTxns, setUnusualTxns] = useState<FlaggedTransaction[]>([]);
  const [riskFilter, setRiskFilter] = useState<'all' | 'High' | 'Medium' | 'Low'>('all');
  const [validationError, setValidationError] = useState<string | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);

  // Sampling Configuration
  const [sampleMethod, setSampleMethod] = useState<'materiality' | 'mus' | 'random'>('materiality');
  const [cutoffAmount, setCutoffAmount] = useState(100000);
  const [samplingStats, setSamplingStats] = useState({ coverageAmount: 0, totalAmount: 0, coveragePercentage: 0 });

  const fileInputRef = useRef<HTMLInputElement>(null);

  const processLoadedTransactions = (parsed: AuditTransaction[], fileName: string, fileSize: string) => {
    if (parsed.length === 0) {
      setTransactions([]);
      setSampledTxns([]);
      setUnusualTxns([]);
      setValidationError('No valid transaction records found in uploaded file.');
      onShowToast('Empty Ledger', 'Zero transaction rows detected.', 'error');
      return;
    }

    setFile({ name: fileName, size: fileSize, count: parsed.length });
    setTransactions(parsed);
    setValidationError(null);

    // Run Sampling
    const sampledRes = sampleTransactions(parsed, sampleMethod, cutoffAmount);
    setSampledTxns(sampledRes.sampled);
    setSamplingStats({
      coverageAmount: sampledRes.coverageAmount,
      totalAmount: sampledRes.totalPopulationAmount,
      coveragePercentage: sampledRes.coveragePercentage,
    });

    // Run Exception Detectors
    const flagged = detectUnusualTransactions(parsed);
    setUnusualTxns(flagged);

    const highRiskCount = flagged.filter((f) => f.riskScore === 'High').length;
    onShowToast(
      'Ledger Analyzed',
      `Processed ${parsed.length} entries. Found ${flagged.length} audit exceptions (${highRiskCount} High Risk).`,
      'success'
    );
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;

    setValidationError(null);
    setIsProcessing(true);

    const lowerName = f.name.toLowerCase();
    const isPdf = lowerName.endsWith('.pdf');
    const isSheet = lowerName.endsWith('.xlsx') || lowerName.endsWith('.xls') || lowerName.endsWith('.csv');

    if (!isSheet && !isPdf) {
      setValidationError('Unsupported format. Please upload an Excel/CSV ledger dump (.xlsx, .csv) or a digital PDF statement.');
      onShowToast('Invalid Format', 'Supported formats: .xlsx, .csv, .pdf', 'error');
      setIsProcessing(false);
      return;
    }

    if (isPdf) {
      try {
        const pdfRes = await parsePdfFile(f);
        if (pdfRes.status === 'success' && pdfRes.rows.length > 0) {
          const rawRows = pdfRes.rows.map((r) => ({
            'Voucher No': r.refNo || 'TXN',
            Date: r.date,
            Particulars: r.particulars,
            Amount: r.debit > 0 ? r.debit : r.credit,
            'Payment Mode': r.debit > 0 ? 'Payment' : 'Receipt',
          }));
          const parsed = parseAuditTransactions(rawRows);
          processLoadedTransactions(parsed, f.name, `${Math.round(f.size / 1024)} KB`);
        } else {
          setTransactions([]);
          setSampledTxns([]);
          setUnusualTxns([]);
          setValidationError(pdfRes.errorMessage || 'Unsupported PDF: No ledger transaction records detected.');
          onShowToast('Extraction Stopped', 'Unsupported PDF. Halting to prevent corrupt data.', 'error');
        }
      } catch (err: any) {
        setValidationError(`Failed to parse PDF ledger: ${err.message || 'Corrupted file'}`);
        onShowToast('Read Error', 'Could not read PDF ledger.', 'error');
      } finally {
        setIsProcessing(false);
      }
      return;
    }

    // Spreadsheet parsing
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary', cellDates: true });
        const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]) as Record<string, any>[];

        if (!json || json.length === 0) {
          setValidationError(`Spreadsheet "${f.name}" contains zero transaction rows.`);
          onShowToast('Empty Ledger', 'No rows found in uploaded file.', 'error');
          setIsProcessing(false);
          return;
        }

        const parsed = parseAuditTransactions(json);
        processLoadedTransactions(parsed, f.name, `${Math.round(f.size / 1024)} KB`);
      } catch (err: any) {
        setValidationError(`Failed to parse file: ${err.message || 'Corrupted file format'}`);
        onShowToast('Parse Error', 'Could not parse ledger file.', 'error');
      } finally {
        setIsProcessing(false);
      }
    };
    reader.readAsBinaryString(f);
  };

  
  const handleRunSampling = () => {
    if (transactions.length === 0) {
      setValidationError('Please upload your ledger file before running sampling.');
      onShowToast('No File', 'Upload a ledger before sampling.', 'error');
      return;
    }

    const res = sampleTransactions(transactions, sampleMethod, cutoffAmount);
    setSampledTxns(res.sampled);
    setSamplingStats({
      coverageAmount: res.coverageAmount,
      totalAmount: res.totalPopulationAmount,
      coveragePercentage: res.coveragePercentage,
    });
    onShowToast(
      'Sampling Complete',
      `Selected ${res.sampled.length} vouchers covering ${res.coveragePercentage.toFixed(1)}% of total value.`,
      'success'
    );
  };

  const handleToggleChecklist = (id: string) => {
    setChecklist((prev) =>
      prev.map((item) => {
        if (item.id === id) {
          const nextStatus = item.status === 'Verified' ? 'In Progress' : 'Verified';
          return { ...item, status: nextStatus };
        }
        return item;
      })
    );
    onShowToast('Status Updated', 'Audit checklist clause updated.', 'info');
  };

  // Export styled checklist
  const handleDownloadChecklist = () => {
    const exportData = checklist.map((c) => ({
      'Clause / Section': c.clauseOrSection,
      'Clause Title': c.title,
      'Audit Description': c.description,
      'Verification Status': c.status,
      'Evidence Reference': c.evidenceRef || 'Verified with books of accounts',
      'Auditor Comments': c.comments || 'Compliant with reporting requirements',
      'Risk Level': c.riskLevel,
    }));

    exportStyledWorkbook(
      [
        {
          name: 'CARO_3CD_Checklist',
          data: exportData,
          columnWidths: {
            'Clause / Section': 18,
            'Clause Title': 28,
            'Audit Description': 45,
            'Verification Status': 18,
            'Evidence Reference': 32,
            'Auditor Comments': 35,
            'Risk Level': 14,
          },
        },
      ],
      'Audit_Checklist_Working_Paper.xlsx'
    );
    onShowToast('Downloaded', 'Exported styled audit checklist working paper.', 'success');
  };

  // Export styled audit working paper
  const handleDownloadAuditWP = () => {
    if (transactions.length === 0) return;
    const currentSampled = sampledTxns.length > 0 ? sampledTxns : sampleTransactions(transactions, sampleMethod, cutoffAmount).sampled;
    const wb = createAuditWorkingPaper(currentSampled, unusualTxns);
    const fileName = `Audit_Sampling_And_Exceptions_WP_${(file?.name || 'Ledger').replace(/\.[^/.]+$/, '')}.xlsx`;
    XLSX.writeFile(wb, fileName);
    onShowToast('Downloaded', `Saved ${fileName} in Excel format with custom styling and column widths.`, 'success');
  };

  const filteredUnusualTxns = unusualTxns.filter((u) => {
    if (riskFilter === 'all') return true;
    return u.riskScore === riskFilter;
  });

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="audit"
        title={
          selectedUtility === 'checklist'
            ? 'Statutory Audit Checklist (CARO 2020 & 3CD)'
            : selectedUtility === 'sampling'
            ? 'Transaction Sampling Engine (MUS & Materiality)'
            : selectedUtility === 'unusual'
            ? 'Unusual Transaction Flags & Tax Audit Detectors'
            : 'CARO & 3CD Audit Scrutiny Suite'
        }
        description="Comprehensive audit testing suite: CARO checklist, monetary unit sampling, Section 40A(3) cash flags, and round-sum detectors."
        onBack={() => (selectedUtility ? setSelectedUtility(null) : onBack())}
      />

      {/* 2. Global Upload & Ledger Status Banner */}
      <input
        type="file"
        ref={fileInputRef}
        onChange={handleFileUpload}
        accept=".xlsx,.xls,.csv,.pdf"
        className="hidden"
        id="audit-tools-upload"
      />

      {!file ? (
        <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-8 sm:p-10 text-center transition-colors shadow-2xs">
          <label htmlFor="audit-tools-upload" className="cursor-pointer flex flex-col items-center">
            <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mb-3 shadow-2xs">
              <FileSpreadsheet className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-semibold text-gray-900">Upload Client Ledger or Daybook</h3>
            <p className="text-xs text-gray-500 mt-1 max-w-md mx-auto leading-relaxed">
              Upload Daybook, General Ledger, Payment Register, or Cash Book in Excel (.xlsx, .xls), CSV, or digital PDF to run CARO, 3CD, MUS sampling, and audit testing.
            </p>
            <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
              <UploadCloud className="w-4 h-4" />
              <span>{isProcessing ? 'Analyzing...' : 'Upload Ledger / Daybook'}</span>
            </div>
            <span className="text-[11px] text-gray-400 font-medium mt-3 block">
              Supports .xlsx, .xls, .csv, digital .pdf · 100% private client-side processing
            </span>
          </label>
        </div>
      ) : (
        <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-sm space-y-4">
          <div className="flex flex-col sm:flex-row items-center justify-between gap-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-lg bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 shrink-0">
                <FileSpreadsheet className="w-5 h-5" />
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <span className="text-xs font-semibold text-gray-900 block">
                    {file.name}
                  </span>
                  <span className="text-[10px] font-semibold text-gray-900 bg-gray-100 border border-gray-200 px-2 py-0.5 rounded">
                    Ready
                  </span>
                </div>
                <span className="text-[11px] text-gray-600 block mt-0.5">
                  {file.size} • {transactions.length} vouchers loaded • {unusualTxns.length} exceptions flagged
                </span>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={() => fileInputRef.current?.click()}
                disabled={isProcessing}
                className="px-3.5 py-1.5 text-xs text-gray-900 bg-gray-50 hover:bg-[#EAEAEA] rounded-lg border border-gray-200 flex items-center gap-1.5 transition-colors font-medium disabled:opacity-50 cursor-pointer"
              >
                {isProcessing ? <RefreshCw className="w-3.5 h-3.5 animate-spin" /> : <UploadCloud className="w-3.5 h-3.5 text-gray-600" />}
                <span>{isProcessing ? 'Analyzing...' : 'Change Ledger'}</span>
              </button>
              {transactions.length > 0 && (
                <button
                  onClick={handleDownloadAuditWP}
                  className="inline-flex items-center gap-1.5 px-3.5 py-1.5 bg-gray-900 hover:bg-black text-white rounded-lg text-xs font-semibold shadow-sm transition-colors cursor-pointer"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>Export Working Paper (.xlsx)</span>
                </button>
              )}
            </div>
          </div>

          {/* Live Population Summary Strip */}
          {transactions.length > 0 && (
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 pt-4 border-t border-gray-200 text-xs">
              <div className="p-2.5 bg-gray-50 rounded-lg border border-gray-200">
                <span className="text-gray-600 block text-[11px]">Total Population Value</span>
                <span className="font-bold text-gray-900 font-mono mt-0.5 block">
                  {formatINR(samplingStats.totalAmount)}
                </span>
              </div>
              <div className="p-2.5 bg-gray-50 rounded-lg border border-gray-200">
                <span className="text-gray-600 block text-[11px]">Audit Sampled Value</span>
                <span className="font-bold text-gray-900 font-mono mt-0.5 block">
                  {formatINR(samplingStats.coverageAmount)} ({samplingStats.coveragePercentage.toFixed(1)}%)
                </span>
              </div>
              <div className="p-2.5 bg-gray-50 rounded-lg border border-gray-200">
                <span className="text-gray-600 block text-[11px]">Tax Audit Exceptions</span>
                <span className="font-bold text-[#B91C1C] font-mono mt-0.5 block">
                  {unusualTxns.length} items flagged
                </span>
              </div>
              <div className="p-2.5 bg-gray-50 rounded-lg border border-gray-200">
                <span className="text-gray-600 block text-[11px]">High Risk Disallowances</span>
                <span className="font-bold text-[#DC2626] font-mono mt-0.5 block">
                  {unusualTxns.filter((u) => u.riskScore === 'High').length} High Priority
                </span>
              </div>
            </div>
          )}
        </div>
      )}

      {/* Error banner */}
      {validationError && (
        <div className="p-4 rounded-xl bg-[#FEF2F2] border border-[#FCA5A5] text-xs text-[#991B1B] flex items-center justify-between">
          <div className="flex items-center gap-2">
            <AlertTriangle className="w-4 h-4 shrink-0" />
            <span>{validationError}</span>
          </div>
          <button onClick={() => setValidationError(null)} className="underline hover:no-underline font-semibold">
            Dismiss
          </button>
        </div>
      )}

      {/* 3. Utility Selector Cards (Shown when on main landing) */}
      {!selectedUtility && (
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          <div
            onClick={() => setSelectedUtility('checklist')}
            className="group bg-white border border-gray-200 hover:border-gray-300 rounded-xl p-6 h-[190px] flex flex-col justify-between cursor-pointer select-none transition-all duration-200 ease-out hover:-translate-y-1 hover:shadow-[0_8px_20px_rgba(22,118,83,0.08)] shadow-[0_1px_3px_rgba(0,0,0,0.02)]"
          >
            <div>
              <div className="w-9 h-9 rounded-lg bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 transition-colors group-hover:bg-gray-200 group-hover:border-gray-200 shrink-0">
                <CheckSquare className="w-4 h-4" strokeWidth={2.2} />
              </div>
            </div>

            <div className="my-auto pt-2">
              <h2 className="text-sm font-bold text-gray-900 tracking-tight group-hover:text-gray-900 transition-colors">
                Audit Checklist
              </h2>
              <p className="text-xs text-gray-600 leading-relaxed mt-1 line-clamp-2">
                CARO 2020, Tax Audit 3CD & ICAI SQC-1 compliance verification clauses.
              </p>
            </div>

            <div className="flex justify-end pt-1">
              <ArrowRight
                className="w-4 h-4 text-[#8D9B93] group-hover:text-gray-900 group-hover:translate-x-1 transition-all duration-200"
                strokeWidth={2}
              />
            </div>
          </div>

          <div
            onClick={() => {
              setSelectedUtility('sampling');
              if (transactions.length > 0) handleRunSampling();
            }}
            className="group bg-white border border-gray-200 hover:border-gray-300 rounded-xl p-6 h-[190px] flex flex-col justify-between cursor-pointer select-none transition-all duration-200 ease-out hover:-translate-y-1 hover:shadow-[0_8px_20px_rgba(22,118,83,0.08)] shadow-[0_1px_3px_rgba(0,0,0,0.02)]"
          >
            <div>
              <div className="w-9 h-9 rounded-lg bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 transition-colors group-hover:bg-gray-200 group-hover:border-gray-200 shrink-0">
                <Sliders className="w-4 h-4" strokeWidth={2.2} />
              </div>
            </div>

            <div className="my-auto pt-2">
              <h2 className="text-sm font-bold text-gray-900 tracking-tight group-hover:text-gray-900 transition-colors">
                Transaction Sampling Engine
              </h2>
              <p className="text-xs text-gray-600 leading-relaxed mt-1 line-clamp-2">
                Monetary Unit Sampling (MUS), materiality cutoffs, and stratified audit testing.
              </p>
            </div>

            <div className="flex justify-end pt-1">
              <ArrowRight
                className="w-4 h-4 text-[#8D9B93] group-hover:text-gray-900 group-hover:translate-x-1 transition-all duration-200"
                strokeWidth={2}
              />
            </div>
          </div>

          <div
            onClick={() => setSelectedUtility('unusual')}
            className="group bg-white border border-gray-200 hover:border-gray-300 rounded-xl p-6 h-[190px] flex flex-col justify-between cursor-pointer select-none transition-all duration-200 ease-out hover:-translate-y-1 hover:shadow-[0_8px_20px_rgba(22,118,83,0.08)] shadow-[0_1px_3px_rgba(0,0,0,0.02)]"
          >
            <div>
              <div className="w-9 h-9 rounded-lg bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 transition-colors group-hover:bg-gray-200 group-hover:border-gray-200 shrink-0">
                <AlertOctagon className="w-4 h-4" strokeWidth={2.2} />
              </div>
            </div>

            <div className="my-auto pt-2">
              <h2 className="text-sm font-bold text-gray-900 tracking-tight group-hover:text-gray-900 transition-colors">
                Unusual Transactions Scrutiny
              </h2>
              <p className="text-xs text-gray-600 leading-relaxed mt-1 line-clamp-2">
                Detect Section 40A(3) cash limits, split payments, round sums, and fraud indicators.
              </p>
            </div>

            <div className="flex justify-end pt-1">
              <ArrowRight
                className="w-4 h-4 text-[#8D9B93] group-hover:text-gray-900 group-hover:translate-x-1 transition-all duration-200"
                strokeWidth={2}
              />
            </div>
          </div>
        </div>
      )}

      {/* 4. Utility 1: Audit Checklist */}
      {selectedUtility === 'checklist' && (
        <div className="space-y-4">
          {checklist.length === 0 ? (
            <div className="bg-white border border-dashed border-gray-300 rounded-xl p-12 text-center max-w-xl mx-auto shadow-sm my-6">
              <div className="w-12 h-12 bg-gray-100 rounded-xl flex items-center justify-center mx-auto mb-4 text-gray-700">
                <CheckSquare className="w-6 h-6" />
              </div>
              <h3 className="text-sm font-bold text-gray-900 mb-1">No Audit Clauses Loaded</h3>
              <p className="text-xs text-gray-500 mb-6 max-w-md mx-auto">
                Load the standard CARO 2020 & Form 3CD compliance verification clauses to review fixed assets, stock verification, related party loans, statutory dues, and cash transactions.
              </p>
              <div className="flex items-center justify-center gap-3">
                <button
                  onClick={() => {
                    setChecklist(INITIAL_AUDIT_CHECKLIST);
                  }}
                  className="px-4 py-2 bg-gray-900 text-white rounded-lg text-xs font-semibold hover:bg-black transition-colors flex items-center gap-2 shadow-sm cursor-pointer"
                >
                  <CheckCircle2 className="w-4 h-4 text-emerald-400" />
                  Load Standard CARO & 3CD Checklist
                </button>
              </div>
            </div>
          ) : (
            <>
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <span className="text-xs text-gray-600">
                  {checklist.filter((c) => c.status === 'Verified' || c.status === 'Signed-off').length} of {checklist.length} Verified
                </span>
                <div className="flex items-center gap-2">
                  <button
                    onClick={() => setChecklistFilter('all')}
                    className={`px-2.5 py-1 rounded-md text-xs font-medium ${
                      checklistFilter === 'all' ? 'bg-gray-100 text-gray-900 font-semibold' : 'text-gray-600'
                    }`}
                  >
                    All Clauses
                  </button>
                  <button
                    onClick={() => setChecklistFilter('verified')}
                    className={`px-2.5 py-1 rounded-md text-xs font-medium ${
                      checklistFilter === 'verified' ? 'bg-gray-100 text-gray-900 font-semibold' : 'text-gray-600'
                    }`}
                  >
                    Verified Only
                  </button>
                  <button
                    onClick={handleDownloadChecklist}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-gray-900 hover:bg-black text-white rounded-lg text-xs font-semibold shadow-sm transition-colors"
                  >
                    <Download className="w-3.5 h-3.5" />
                    <span>Export Checklist (.xlsx)</span>
                  </button>
                </div>
              </div>

              <div className="bg-white rounded-xl border border-gray-200 divide-y divide-[#E5E7EB] shadow-sm">
                {checklist
                  .filter((item) => {
                    if (checklistFilter === 'verified') return item.status === 'Verified' || item.status === 'Signed-off';
                    return true;
                  })
                  .map((item) => (
                    <div key={item.id} className="p-4 flex items-start gap-3 hover:bg-gray-50">
                      <input
                        type="checkbox"
                        checked={item.status === 'Verified' || item.status === 'Signed-off'}
                        onChange={() => handleToggleChecklist(item.id)}
                        className="mt-1 rounded border-gray-200 text-gray-900 focus:ring-gray-900 cursor-pointer"
                      />
                      <div className="flex-1 space-y-1">
                        <div className="flex items-center gap-2">
                          <span className="text-[11px] font-semibold font-mono text-gray-900 bg-gray-100 px-2 py-0.5 rounded">
                            {item.clauseOrSection}
                          </span>
                          <span className="text-xs font-bold text-gray-900">{item.title}</span>
                          <span className="text-[10px] text-[#8D9B93] bg-gray-50 px-1.5 py-0.5 rounded">
                            {item.riskLevel} Risk
                          </span>
                        </div>
                        <p className="text-xs text-gray-600 leading-relaxed">{item.description}</p>
                        {item.comments && (
                          <p className="text-[11px] text-[#4A554E] italic">Auditor Note: {item.comments}</p>
                        )}
                      </div>
                    </div>
                  ))}
              </div>
            </>
          )}
        </div>
      )}

      {/* 5. Utility 2: Transaction Sampling Engine */}
      {selectedUtility === 'sampling' && (
        <div className="space-y-5">
          {/* Controls Card */}
          <div className="bg-white rounded-xl border border-gray-200 p-6 shadow-sm space-y-4">
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
              <div>
                <label className="text-xs font-semibold text-gray-900 block mb-1">Sampling Methodology</label>
                <select
                  value={sampleMethod}
                  onChange={(e) => setSampleMethod(e.target.value as any)}
                  className="w-full px-3 py-2 bg-gray-50 border border-gray-200 rounded-lg text-xs font-medium"
                >
                  <option value="materiality">Materiality Cutoff (100% High-Value)</option>
                  <option value="mus">Monetary Unit Sampling (MUS / PPS)</option>
                  <option value="random">Stratified Systematic (25% Sample)</option>
                </select>
              </div>

              <div>
                <label className="text-xs font-semibold text-gray-900 block mb-1">
                  {sampleMethod === 'materiality' ? 'Cutoff Threshold (₹)' : 'Sample Target Count'}
                </label>
                <input
                  type="number"
                  value={cutoffAmount}
                  onChange={(e) => setCutoffAmount(parseInt(e.target.value) || 0)}
                  className="w-full px-3 py-2 bg-gray-50 border border-gray-200 rounded-lg text-xs font-mono"
                />
              </div>

              <div className="flex items-end gap-2">
                <button
                  onClick={handleRunSampling}
                  disabled={transactions.length === 0}
                  className="w-full py-2 text-xs font-semibold text-white bg-gray-900 hover:bg-black rounded-lg transition-colors shadow-sm flex items-center justify-center gap-1.5 disabled:opacity-50"
                >
                  <RefreshCw className="w-3.5 h-3.5" />
                  <span>Execute Sample</span>
                </button>
              </div>
            </div>
          </div>

          {/* Sampled Vouchers Table */}
          {transactions.length > 0 ? (
            <div className="bg-white rounded-xl border border-gray-200 overflow-hidden shadow-sm">
              <div className="px-5 py-3 bg-gray-50 border-b border-gray-200 flex items-center justify-between text-xs">
                <span className="font-bold text-gray-900">
                  Sampled Audit Working Paper ({sampledTxns.length} vouchers selected)
                </span>
                <span className="text-gray-900 font-medium">
                  Covers {formatINR(samplingStats.coverageAmount)} ({samplingStats.coveragePercentage.toFixed(1)}%)
                </span>
              </div>

              <div className="overflow-x-auto max-h-[460px]">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="sticky top-0 bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold">
                    <tr>
                      <th className="py-2.5 px-3 border-r border-gray-200 w-10 text-center">#</th>
                      <th className="py-2.5 px-3 border-r border-gray-200">Voucher No</th>
                      <th className="py-2.5 px-3 border-r border-gray-200">Date</th>
                      <th className="py-2.5 px-3 border-r border-gray-200">Party / Ledger Particulars</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 text-right">Amount (₹)</th>
                      <th className="py-2.5 px-3">Audit Verification Status</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#E5E7EB] text-gray-900">
                    {sampledTxns.map((t, idx) => (
                      <tr key={t.id} className="hover:bg-gray-50">
                        <td className="py-2 px-3 border-r border-gray-200 text-center text-gray-600 text-[11px]">
                          {idx + 1}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 font-mono text-[11px]">
                          {t.voucherNo}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 whitespace-nowrap">
                          {t.date}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 font-medium">
                          {t.partyName}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 text-right font-mono font-semibold">
                          {formatINR(t.amount)}
                        </td>
                        <td className="py-2 px-3 text-gray-900 font-medium">
                          Substantive Sample (WP-AUDIT-{idx + 1})
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          ) : (
            <div className="bg-white rounded-xl border border-gray-200 p-8 text-center text-xs text-gray-600">
              <Sliders className="w-8 h-8 text-[#A0AAA2] mx-auto mb-2" />
              <p className="font-semibold text-gray-900">No Ledger Uploaded for Sampling</p>
              <p className="text-[11px] text-gray-600 mt-1 max-w-md mx-auto">
                Upload your client Daybook or General Ledger above to generate statistical MUS and materiality samples.
              </p>
            </div>
          )}
        </div>
      )}

      {/* 6. Utility 3: Unusual Transactions Scrutiny */}
      {selectedUtility === 'unusual' && (
        <div className="space-y-4">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-white p-3 rounded-xl border border-gray-200">
            <div className="flex items-center gap-2">
              <span className="text-xs font-semibold text-gray-900">Filter Risk:</span>
              <button
                onClick={() => setRiskFilter('all')}
                className={`px-2.5 py-1 text-xs rounded-md ${
                  riskFilter === 'all' ? 'bg-gray-900 text-white font-medium' : 'bg-gray-50 text-gray-600'
                }`}
              >
                All ({unusualTxns.length})
              </button>
              <button
                onClick={() => setRiskFilter('High')}
                className={`px-2.5 py-1 text-xs rounded-md ${
                  riskFilter === 'High' ? 'bg-[#DC2626] text-white font-medium' : 'bg-[#FDF2F2] text-[#991B1B]'
                }`}
              >
                High Risk ({unusualTxns.filter((u) => u.riskScore === 'High').length})
              </button>
              <button
                onClick={() => setRiskFilter('Medium')}
                className={`px-2.5 py-1 text-xs rounded-md ${
                  riskFilter === 'Medium' ? 'bg-[#D97706] text-white font-medium' : 'bg-[#FEF9EE] text-[#92400E]'
                }`}
              >
                Medium ({unusualTxns.filter((u) => u.riskScore === 'Medium').length})
              </button>
            </div>

            {transactions.length > 0 && (
              <span className="text-xs text-gray-600">
                Scrutinized {transactions.length} entries against Section 40A(3), Section 269ST, and SA 240.
              </span>
            )}
          </div>

          {filteredUnusualTxns.length > 0 ? (
            <div className="bg-white rounded-xl border border-gray-200 divide-y divide-[#E5E7EB] shadow-sm">
              {filteredUnusualTxns.map((u) => (
                <div key={u.id} className="p-4 flex items-start gap-3 hover:bg-gray-50">
                  <div
                    className={`p-2 rounded-lg mt-0.5 shrink-0 ${
                      u.riskScore === 'High' ? 'bg-[#FEE2E2] text-[#DC2626]' : 'bg-[#FEF9EE] text-[#D97706]'
                    }`}
                  >
                    <AlertOctagon className="w-4 h-4" />
                  </div>
                  <div className="flex-1 space-y-1.5">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-2">
                        <span className="text-xs font-bold text-gray-900">{u.partyName}</span>
                        <span className="text-[10px] font-mono text-gray-600 bg-gray-50 px-1.5 py-0.5 rounded">
                          {u.voucherNo}
                        </span>
                        <span
                          className={`text-[9px] px-1.5 py-0.5 rounded font-bold uppercase ${
                            u.riskScore === 'High' ? 'bg-[#FEE2E2] text-[#991B1B]' : 'bg-[#FEF9EE] text-[#92400E]'
                          }`}
                        >
                          {u.riskScore} Risk
                        </span>
                      </div>
                      <span className="text-xs font-bold font-mono text-gray-900">{formatINR(u.amount)}</span>
                    </div>

                    <div className="flex flex-wrap items-center gap-1.5">
                      {u.flagTypes.map((flag, idx) => (
                        <span
                          key={idx}
                          className="text-[10px] font-semibold bg-[#FEE2E2] text-[#991B1B] px-2 py-0.5 rounded"
                        >
                          {flag}
                        </span>
                      ))}
                      <span className="text-[10px] text-gray-600">• {u.statutoryClause}</span>
                      <span className="text-[10px] text-gray-600">• Posted on {u.date}</span>
                    </div>

                    <p className="text-xs text-[#4A554E] leading-relaxed">{u.explanation}</p>
                    <div className="p-2 bg-gray-50 rounded border border-gray-200 text-[11px] text-gray-900">
                      <strong>Auditor Action:</strong> {u.auditorAction}
                    </div>
                  </div>
                </div>
              ))}
            </div>
          ) : (
            <div className="bg-white rounded-xl border border-gray-200 p-8 text-center text-xs text-gray-600">
              <AlertOctagon className="w-8 h-8 text-[#A0AAA2] mx-auto mb-2" />
              <p className="font-semibold text-gray-900">No Unusual Transactions Flagged</p>
              <p className="text-[11px] text-gray-600 mt-1 max-w-md mx-auto">
                {transactions.length > 0
                  ? 'All entries in the loaded ledger passed Section 40A(3), Section 269ST, and cash threshold checks.'
                  : 'Upload your client ledger above to run the 8 statutory compliance detectors.'}
              </p>
            </div>
          )}
        </div>
      )}
    </div>
  );
};