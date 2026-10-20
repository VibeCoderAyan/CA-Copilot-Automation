import { UploadIllustration } from '../common/UploadIllustration';
import { InnerToolBadge } from '../common/InnerToolBadge';
import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  Download,
  CheckCircle2,
  AlertTriangle,
  AlertCircle,
  RefreshCw,
  FileText,
  Sliders,
  Check,
  X,
  HelpCircle,
  Copy,
  Info,
  FileSpreadsheet,
  FileCheck,
  UploadCloud,
} from 'lucide-react';
import {
  parsePurchaseRegisterRows,
  parseGstr2bRows,
  runGstReconciliation,
  createGstReconciliationWorkbook,
  exportGstReconciliationCsv,
  type GstComparisonRow,
  type GstReconciliationSummary,
  type GstVoucher,
  type GstReconciliationConfig,
  DEFAULT_GST_CONFIG,
} from '../../utils/gstReconciliationEngine';
import { ToolIcon } from '../common/ToolIcon';
import { DataFormatGuide } from '../common/DataFormatGuide';
import { formatINR, extractSmartTableFromSheet } from '../../utils/dataSanitizer';
import { getActiveClient } from '../../utils/clientStorage';

interface GstReconciliationToolProps {
  onBack: () => void;
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
}

export const GstReconciliationTool: React.FC<GstReconciliationToolProps> = ({ onBack, onShowToast }) => {
  const [booksFile, setBooksFile] = useState<{ name: string; size: string } | null>(null);
  const [gstr2bFile, setGstr2bFile] = useState<{ name: string; size: string } | null>(null);

  const [booksVouchers, setBooksVouchers] = useState<GstVoucher[]>([]);
  const [gstr2bVouchers, setGstr2bVouchers] = useState<GstVoucher[]>([]);

  const [comparisonRecords, setComparisonRecords] = useState<GstComparisonRow[]>([]);
  const [summary, setSummary] = useState<GstReconciliationSummary | null>(null);
  const [isCompared, setIsCompared] = useState(false);
  const [activeFilter, setActiveFilter] = useState<'all' | 'matched' | 'mismatch' | 'missing2b' | 'missingBooks' | 'duplicate'>('all');
  const [validationError, setValidationError] = useState<string | null>(null);

  // Matching configuration tolerances
  const [config, setConfig] = useState<GstReconciliationConfig>(DEFAULT_GST_CONFIG);
  const [showToleranceConfig, setShowToleranceConfig] = useState(false);

  const handleBooksUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;
    setValidationError(null);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const buffer = evt.target?.result as ArrayBuffer;
        const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
        const smartTable = extractSmartTableFromSheet(wb.Sheets[wb.SheetNames[0]]);
        const json = smartTable.rows;

        if (!json || json.length === 0) {
          setValidationError(`Purchase register "${f.name}" contains zero invoice records.`);
          onShowToast('Empty File', 'No data rows found in purchase register.', 'error');
          return;
        }

        const parsed = parsePurchaseRegisterRows(json);
        setBooksFile({ name: f.name, size: `${Math.round(f.size / 1024)} KB` });
        setBooksVouchers(parsed);
        setIsCompared(false);
      } catch (err: any) {
        setValidationError(`Failed to parse purchase register: ${err.message || 'Corrupted file'}`);
        onShowToast('Parse Error', 'Could not read purchase register.', 'error');
      }
    };
    reader.readAsArrayBuffer(f);
  };

  const handleGstr2bUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;
    setValidationError(null);

    const isJson = f.name.toLowerCase().endsWith('.json');
    const reader = new FileReader();

    reader.onload = (evt) => {
      try {
        let parsed: GstVoucher[] = [];
        if (isJson) {
          const jsonText = evt.target?.result as string;
          const parsedJson = JSON.parse(jsonText);
          parsed = parseGstr2bRows(parsedJson);
        } else {
          const buffer = evt.target?.result as ArrayBuffer;
          const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
          const b2bSheetName = wb.SheetNames.find((s) => /^b2b$/i.test(s)) ||
                               wb.SheetNames.find((s) => /b2b/i.test(s)) ||
                               wb.SheetNames[0];
          const smartTable = extractSmartTableFromSheet(wb.Sheets[b2bSheetName]);
          parsed = parseGstr2bRows(smartTable.rows);
        }

        if (parsed.length === 0) {
          setValidationError(`GSTR-2B file "${f.name}" contains zero valid B2B records.`);
          onShowToast('Empty GSTR-2B', 'No B2B invoices found in GSTR-2B file.', 'error');
          return;
        }

        setGstr2bFile({ name: f.name, size: `${Math.round(f.size / 1024)} KB` });
        setGstr2bVouchers(parsed);
        setIsCompared(false);
      } catch (err: any) {
        setValidationError(`Failed to parse GSTR-2B file: ${err.message || 'Corrupted file'}`);
        onShowToast('Parse Error', 'Could not read GSTR-2B file.', 'error');
      }
    };

    if (isJson) {
      reader.readAsText(f);
    } else {
      reader.readAsArrayBuffer(f);
    }
  };

  
  const handleCompare = () => {
    if (booksVouchers.length === 0 || gstr2bVouchers.length === 0) {
      return;
    }

    const result = runGstReconciliation(booksVouchers, gstr2bVouchers, config);
    setComparisonRecords(result.comparisonRecords);
    setSummary(result.summary);
    setIsCompared(true);

    // Synchronize GST figures to active session storage for 1-click Dossier pulling
    try {
      const activeClient = getActiveClient();
      sessionStorage.setItem(
        'ca_copilot_session_gst',
        JSON.stringify({
          clientName: activeClient?.name || 'Active Client',
          totalBooksTax: result.summary.totalBooksTax,
          total2bTax: result.summary.total2bTax,
          eligibleItc: result.summary.eligibleItc,
          atRiskItc: result.summary.atRiskItc,
          unclaimedItc: result.summary.unclaimedItc,
          matchedCount: result.summary.matchedCount,
          probableCount: result.summary.probableCount,
          mismatchCount: result.summary.mismatchCount,
          missingIn2bCount: result.summary.missingIn2bCount,
          missingInBooksCount: result.summary.missingInBooksCount,
          duplicateCount: result.summary.duplicateCount,
          timestamp: new Date().toISOString(),
        })
      );
    } catch {
      // Storage unavailable
    }

    onShowToast(
      'Reconciliation Completed',
      `Matched: ${result.summary.matchedCount + result.summary.probableCount}, Tax Mismatch: ${result.summary.mismatchCount}, Missing in 2B: ${result.summary.missingIn2bCount}, Duplicates: ${result.summary.duplicateCount}.`,
      'success'
    );
  };

  const handleSetResolution = (id: string, resolution: any) => {
    setComparisonRecords((prev) =>
      prev.map((r) => (r.id === id ? { ...r, resolutionStatus: resolution } : r))
    );
    onShowToast('Resolution Updated', `Marked record as ${resolution}.`, 'success');
  };

  const handleDownloadExcel = () => {
    if (!summary) return;
    const rawBooksRows = booksVouchers.map((b) => ({
      GSTIN: b.supplierGstin,
      'Supplier Name': b.supplierName,
      'Invoice No': b.invoiceNumber,
      'Invoice Date': b.invoiceDate,
      'Taxable Value (₹)': b.taxableAmount,
      'CGST (₹)': b.cgst,
      'SGST (₹)': b.sgst,
      'IGST (₹)': b.igst,
      'Total Value (₹)': b.taxableAmount + b.totalTax,
    }));
    const raw2bRows = gstr2bVouchers.map((g) => ({
      'GSTIN of Supplier': g.supplierGstin,
      'Trade/Legal Name': g.supplierName,
      'Invoice Number': g.invoiceNumber,
      'Invoice Date': g.invoiceDate,
      'Taxable Value (₹)': g.taxableAmount,
      'CGST (₹)': g.cgst,
      'SGST (₹)': g.sgst,
      'IGST (₹)': g.igst,
      'Total Tax (₹)': g.totalTax,
    }));
    const wb = createGstReconciliationWorkbook(comparisonRecords, summary, rawBooksRows, raw2bRows);
    XLSX.writeFile(wb, 'GST_Reconciliation_Report_FY26.xlsx');
    onShowToast('6-Sheet Excel Downloaded', 'Exported comprehensive GST reconciliation workbook.', 'success');
  };

  const handleDownloadCsv = () => {
    if (comparisonRecords.length === 0) return;
    const csvData = exportGstReconciliationCsv(comparisonRecords);
    const blob = new Blob([csvData], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', 'GST_Reconciliation_Report_FY26.csv');
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    onShowToast('CSV Downloaded', 'Exported reconciliation report as CSV.', 'success');
  };

  const displayRecords = comparisonRecords.filter((r) => {
    if (activeFilter === 'matched') return r.status === 'Exact Match' || r.status === 'Probable Match';
    if (activeFilter === 'mismatch') return r.status === 'Tax Mismatch';
    if (activeFilter === 'missing2b') return r.status === 'Missing in 2B';
    if (activeFilter === 'missingBooks') return r.status === 'Missing in Books';
    if (activeFilter === 'duplicate') return r.status === 'Duplicate Record';
    return true;
  });

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
            <InnerToolBadge toolId="gst-rec" />
            <div>
              <h1 className="text-xl font-bold text-gray-900 tracking-tight">GSTR-2B vs Books Reconciliation</h1>
              <p className="text-xs text-gray-500 mt-0.5">
                Automated invoice comparison, configurable tolerances, Section 16(2)(aa) compliance tagging, and vendor follow-up notices.
              </p>
            </div>
          </div>

          {summary && (
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
                title="Export 6-Sheet Excel Report"
              >
                <Download className="w-4 h-4" />
                <span>Export 6-Sheet Excel (.xlsx)</span>
              </button>
            </div>
          )}
        </div>
      </div>

      {/* 2. Format Requirements Guide */}
      <DataFormatGuide
        toolType="gst-rec"
        acceptedFormats=".xlsx, .xls, .csv, .json (GSTN Portal GSTR-2B JSON)"
        requiredFields={[
          { name: 'Supplier GSTIN', description: '15-digit GST identification number of seller', aliases: ['GSTIN', 'Supplier GSTIN', 'ctin', 'GST No'] },
          { name: 'Invoice Number', description: 'Tax invoice serial number', aliases: ['Invoice No', 'Bill No', 'inum', 'Vch No'] },
          { name: 'Invoice Date', description: 'Date of invoice issue', aliases: ['Date', 'Invoice Date', 'idt'] },
          { name: 'Taxable Value', description: 'Commercial net taxable amount', aliases: ['Taxable', 'Taxable Value (₹)', 'txval'] },
          { name: 'CGST, SGST, IGST, Cess', description: 'Statutory GST component breakup', aliases: ['CGST', 'SGST', 'IGST', 'Cess', 'Total Tax'] },
        ]}
        notes={[
          'Configurable Tolerances: Adjust tax difference threshold (₹1 - ₹10) and enable fuzzy matching for minor prefix shifts.',
          'Section 16(2)(aa) Safeguard: Reconciling GSTR-2B confirms appearance on the portal; ITC requires actual receipt of supply and valid tax invoice.',
          'Audit Action Templates: Generates pre-formatted vendor follow-up notices for missing invoices directly in the exported workbook.',
        ]}
      />

      {/* Validation Error Banner (Apple Callout Style) */}
      {validationError && (
        <div className="p-3 rounded-[14px] bg-white border border-gray-200/90 shadow-2xs text-xs text-gray-800 flex items-center justify-between animate-in fade-in duration-150">
          <div className="flex items-center gap-2 min-w-0">
            <span className="w-1.5 h-1.5 rounded-full bg-[#FF3B30] shrink-0" />
            <span className="font-medium text-gray-800 truncate">{validationError}</span>
          </div>
          <button
            onClick={() => setValidationError(null)}
            className="text-gray-400 hover:text-gray-700 text-[11px] font-semibold px-2 py-0.5 rounded-md hover:bg-gray-100 transition-colors shrink-0 cursor-pointer ml-3"
          >
            Dismiss
          </button>
        </div>
      )}

      {/* 3. Upload Area: Books + GSTR-2B */}
      <div className="bg-white rounded-xl border border-gray-200 p-6 shadow-sm space-y-4">
        <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
          {/* Box 1: Purchase Register */}
          <input
            type="file"
            accept=".xlsx,.xls,.csv"
            className="hidden"
            id="gstr2b-books-upload"
            onChange={handleBooksUpload}
          />
          {!booksFile ? (
            <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-6 sm:p-8 text-center transition-colors shadow-2xs">
              <label htmlFor="gstr2b-books-upload" className="cursor-pointer flex flex-col items-center">
                <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mb-3 shadow-2xs">
                  <FileSpreadsheet className="w-6 h-6" />
                </div>
                <h3 className="text-sm font-semibold text-gray-900">1. Purchase Register (Accounting Books)</h3>
                <p className="text-xs text-gray-500 mt-1 max-w-xs mx-auto">
                  Purchase ledger export from Tally, SAP, or ERP (.xlsx, .csv)
                </p>
                <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
                  <UploadCloud className="w-4 h-4" />
                  <span>Upload Purchase Register</span>
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
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-bold text-gray-900 block truncate">{booksFile.name}</span>
                    <span className="text-[10px] font-semibold text-gray-900 bg-gray-100 border border-gray-200 px-2 py-0.5 rounded">
                      Ready
                    </span>
                  </div>
                  <span className="text-[11px] text-gray-600 block mt-0.5">
                    Purchase Register • {booksVouchers.length} vouchers • {booksFile.size}
                  </span>
                </div>
              </div>
              <div className="mt-4 pt-3 border-t border-gray-100 flex items-center justify-between">
                <span className="text-[11px] text-gray-500 font-medium">Ready for reconciliation</span>
                <label htmlFor="gstr2b-books-upload" className="cursor-pointer px-3 py-1.5 text-xs text-gray-900 bg-gray-50 hover:bg-[#EAEAEA] rounded-lg border border-gray-200 flex items-center gap-1.5 font-medium transition-colors">
                  <UploadCloud className="w-3.5 h-3.5 text-gray-600" />
                  <span>Change Register</span>
                </label>
              </div>
            </div>
          )}

          {/* Box 2: GSTR-2B Portal Extract */}
          <input
            type="file"
            accept=".xlsx,.xls,.csv,.json"
            className="hidden"
            id="gstr2b-portal-upload"
            onChange={handleGstr2bUpload}
          />
          {!gstr2bFile ? (
            <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-6 sm:p-8 text-center transition-colors shadow-2xs">
              <label htmlFor="gstr2b-portal-upload" className="cursor-pointer flex flex-col items-center">
                <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mb-3 shadow-2xs">
                  <FileCheck className="w-6 h-6" />
                </div>
                <h3 className="text-sm font-semibold text-gray-900">2. GSTR-2B (GST Portal Extract)</h3>
                <p className="text-xs text-gray-500 mt-1 max-w-xs mx-auto">
                  Official GSTR-2B statement file from GST Portal (.xlsx, .json)
                </p>
                <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
                  <UploadCloud className="w-4 h-4" />
                  <span>Upload GSTR-2B File</span>
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
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-bold text-gray-900 block truncate">{gstr2bFile.name}</span>
                    <span className="text-[10px] font-semibold text-gray-900 bg-gray-100 border border-gray-200 px-2 py-0.5 rounded">
                      Ready
                    </span>
                  </div>
                  <span className="text-[11px] text-gray-600 block mt-0.5">
                    GSTR-2B • {gstr2bVouchers.length} invoices • {gstr2bFile.size}
                  </span>
                </div>
              </div>
              <div className="mt-4 pt-3 border-t border-gray-100 flex items-center justify-between">
                <span className="text-[11px] text-gray-500 font-medium">Ready for reconciliation</span>
                <label htmlFor="gstr2b-portal-upload" className="cursor-pointer px-3 py-1.5 text-xs text-gray-900 bg-gray-50 hover:bg-[#EAEAEA] rounded-lg border border-gray-200 flex items-center gap-1.5 font-medium transition-colors">
                  <UploadCloud className="w-3.5 h-3.5 text-gray-600" />
                  <span>Change GSTR-2B</span>
                </label>
              </div>
            </div>
          )}
        </div>

        {/* Action Controls & Tolerance Setting Bar */}
        <div className="flex flex-wrap items-center justify-between gap-3 pt-3 border-t border-gray-100">
          <div className="flex items-center gap-3">
            <button
              onClick={() => setShowToleranceConfig(!showToleranceConfig)}
              className="text-xs font-medium text-gray-700 hover:text-gray-900 flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-gray-200 hover:bg-gray-50 transition-colors"
            >
              <Sliders className="w-3.5 h-3.5 text-gray-500" />
              <span>Matching Tolerances (Tax: ₹{config.taxTolerance})</span>
            </button>
            
          </div>

          <button
            onClick={handleCompare}
            disabled={booksVouchers.length === 0 || gstr2bVouchers.length === 0}
            className={`px-4 py-2 rounded-lg text-xs font-bold text-white shadow-xs transition-colors flex items-center gap-1.5 ${
              booksVouchers.length > 0 && gstr2bVouchers.length > 0
                ? 'bg-gray-900 hover:bg-black cursor-pointer'
                : 'bg-gray-300 cursor-not-allowed'
            }`}
          >
            <RefreshCw className="w-3.5 h-3.5" />
            <span>Reconcile Invoices (PR vs 2B)</span>
          </button>
        </div>

        {/* Tolerance Configuration Drawer */}
        {showToleranceConfig && (
          <div className="p-4 rounded-xl bg-gray-50 border border-gray-200 space-y-3 text-xs">
            <span className="font-bold text-gray-800 block">Configure Matching Tolerances & Rules</span>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
              <div>
                <label className="block text-[11px] font-medium text-gray-600 mb-1">Tax Amount Tolerance (₹)</label>
                <input
                  type="number"
                  step="0.5"
                  value={config.taxTolerance}
                  onChange={(e) => setConfig({ ...config, taxTolerance: parseFloat(e.target.value) || 0 })}
                  className="w-full px-2 py-1 bg-white border border-gray-300 rounded text-xs font-mono"
                />
                <span className="text-[10px] text-gray-500 block mt-0.5">Allows for minor rounding differences</span>
              </div>
              <div>
                <label className="block text-[11px] font-medium text-gray-600 mb-1">Taxable Value Tolerance (₹)</label>
                <input
                  type="number"
                  step="1"
                  value={config.taxableTolerance}
                  onChange={(e) => setConfig({ ...config, taxableTolerance: parseFloat(e.target.value) || 0 })}
                  className="w-full px-2 py-1 bg-white border border-gray-300 rounded text-xs font-mono"
                />
                <span className="text-[10px] text-gray-500 block mt-0.5">Maximum commercial taxable variation</span>
              </div>
              <div className="flex items-center gap-2 pt-4">
                <input
                  type="checkbox"
                  id="fuzzyMatch"
                  checked={config.fuzzyMatching}
                  onChange={(e) => setConfig({ ...config, fuzzyMatching: e.target.checked })}
                  className="rounded text-gray-900 focus:ring-gray-900"
                />
                <label htmlFor="fuzzyMatch" className="text-xs font-medium text-gray-800 cursor-pointer">
                  Fuzzy Invoice Matching (Strips slashes & leading zeros)
                </label>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* 4. Statutory Legal Notice Banner (Section 16(2)(aa)) */}
      <div className="p-3.5 rounded-xl bg-gray-50 border border-gray-200 text-xs text-gray-900 flex items-start gap-2.5">
        <Info className="w-4 h-4 text-gray-800 shrink-0 mt-0.5" />
        <div className="space-y-0.5">
          <span className="font-bold block">Statutory Compliance Notice: Section 16(2)(aa) of CGST Act</span>
          <p className="text-[11px] text-gray-900 leading-relaxed">
            Data matching against GSTR-2B verifies invoice details furnished by suppliers. Under Indian GST law, legal entitlement to Input Tax Credit (ITC) additionally requires physical possession of a valid tax invoice, actual receipt of goods/services, and discharge of tax liability by the supplier in GSTR-3B.
          </p>
        </div>
      </div>

      {/* 5. Reconciliation Results Workspace */}
      {isCompared && summary && (
        <div className="space-y-5">
          {/* Executive ITC Metrics */}
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <span className="text-[11px] font-semibold text-gray-500 block">Total Books Tax (PR)</span>
              <span className="text-xl font-bold text-gray-900 mt-1 block">
                {formatINR(summary.totalBooksTax)}
              </span>
              <span className="text-[10px] text-gray-500 block mt-0.5">Claimed in accounting ledger</span>
            </div>

            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <span className="text-[11px] font-semibold text-gray-500 block">100% Eligible ITC (Matched)</span>
              <span className="text-xl font-bold text-gray-900 mt-1 block">
                {formatINR(summary.eligibleItc)}
              </span>
              <span className="text-[10px] text-gray-900 block mt-0.5 font-medium">Safe to claim in GSTR-3B</span>
            </div>

            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <span className="text-[11px] font-semibold text-gray-500 block">At-Risk ITC (Follow-up)</span>
              <span className="text-xl font-bold text-[#C43D43] mt-1 block">
                {formatINR(summary.atRiskItc)}
              </span>
              <span className="text-[10px] text-[#C43D43] block mt-0.5">Missing in 2B or Tax Mismatch</span>
            </div>

            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <span className="text-[11px] font-semibold text-gray-500 block">Unclaimed Available ITC</span>
              <span className="text-xl font-bold text-amber-600 mt-1 block">
                {formatINR(summary.unclaimedItc)}
              </span>
              <span className="text-[10px] text-gray-500 block mt-0.5">In 2B but missing in books</span>
            </div>
          </div>

          {/* Filter Pills */}
          <div className="bg-white rounded-xl border border-gray-200 overflow-hidden shadow-sm">
            <div className="px-5 py-3 bg-gray-50 border-b border-gray-200 flex flex-wrap items-center justify-between gap-3 text-xs">
              <div className="flex items-center gap-2">
                <span className="font-bold text-gray-900">Invoice Reconciliation Audit Table</span>
                <span className="text-gray-600">({displayRecords.length} records shown)</span>
              </div>

              <div className="flex items-center gap-1.5">
                <button
                  onClick={() => setActiveFilter('all')}
                  className={`px-2.5 py-1 rounded-md text-[11px] font-medium transition-colors ${
                    activeFilter === 'all'
                      ? 'bg-white text-gray-900 border border-gray-200 shadow-sm font-semibold'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  All ({comparisonRecords.length})
                </button>
                <button
                  onClick={() => setActiveFilter('matched')}
                  className={`px-2.5 py-1 rounded-md text-[11px] font-medium transition-colors ${
                    activeFilter === 'matched'
                      ? 'bg-white text-gray-900 border border-gray-200 shadow-sm font-semibold'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Matched ({summary.matchedCount + summary.probableCount})
                </button>
                <button
                  onClick={() => setActiveFilter('mismatch')}
                  className={`px-2.5 py-1 rounded-md text-[11px] font-medium transition-colors ${
                    activeFilter === 'mismatch'
                      ? 'bg-white text-[#C43D43] border border-gray-200 shadow-sm font-semibold'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Tax Mismatch ({summary.mismatchCount})
                </button>
                <button
                  onClick={() => setActiveFilter('missing2b')}
                  className={`px-2.5 py-1 rounded-md text-[11px] font-medium transition-colors ${
                    activeFilter === 'missing2b'
                      ? 'bg-white text-[#C43D43] border border-gray-200 shadow-sm font-semibold'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Missing in 2B ({summary.missingIn2bCount})
                </button>
                <button
                  onClick={() => setActiveFilter('missingBooks')}
                  className={`px-2.5 py-1 rounded-md text-[11px] font-medium transition-colors ${
                    activeFilter === 'missingBooks'
                      ? 'bg-white text-amber-700 border border-gray-200 shadow-sm font-semibold'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Missing in Books ({summary.missingInBooksCount})
                </button>
                {summary.duplicateCount > 0 && (
                  <button
                    onClick={() => setActiveFilter('duplicate')}
                    className={`px-2.5 py-1 rounded-md text-[11px] font-medium transition-colors ${
                      activeFilter === 'duplicate'
                        ? 'bg-white text-purple-700 border border-gray-200 shadow-sm font-semibold'
                        : 'text-gray-600 hover:text-gray-900'
                    }`}
                  >
                    Duplicates ({summary.duplicateCount})
                  </button>
                )}
              </div>
            </div>

            {/* Comparison Side-by-Side Table */}
            <div className="overflow-x-auto max-h-[500px]">
              <table className="w-full text-left text-xs border-collapse">
                <thead className="sticky top-0 bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold">
                  <tr>
                    <th className="py-2.5 px-3 border-r border-gray-200 w-10 text-center">#</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Supplier GSTIN</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Supplier Name</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Invoice No & Date</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap text-right">Books Taxable (₹)</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap text-right font-semibold">Books Tax (₹)</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap text-right">2B Taxable (₹)</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap text-right font-semibold">2B Tax (₹)</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap text-right font-bold">Tax Variance (₹)</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Reconciliation Status</th>
                    <th className="py-2.5 px-3 whitespace-nowrap text-center">Resolution Action</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#E5E7EB] text-gray-900">
                  {displayRecords.map((r, idx) => (
                    <tr
                      key={r.id}
                      className={`hover:bg-gray-50 transition-colors ${
                        r.status === 'Missing in 2B'
                          ? 'bg-rose-50/40'
                          : r.status === 'Tax Mismatch'
                          ? 'bg-amber-50/40'
                          : r.status === 'Duplicate Record'
                          ? 'bg-purple-50/40'
                          : ''
                      }`}
                    >
                      <td className="py-2 px-3 border-r border-gray-200 text-center text-gray-600 text-[11px]">
                        {idx + 1}
                      </td>
                      <td className="py-2 px-3 border-r border-gray-200 font-mono text-[11px] whitespace-nowrap font-medium">
                        {r.supplierGstin || '-'}
                      </td>
                      <td className="py-2 px-3 border-r border-gray-200 max-w-[180px] truncate" title={r.supplierName}>
                        {r.supplierName}
                      </td>
                      <td className="py-2 px-3 border-r border-gray-200 whitespace-nowrap font-mono text-[11px]">
                        <span className="font-semibold block">{r.invoiceNumber}</span>
                        <span className="text-[10px] text-gray-600 block">{r.invoiceDate}</span>
                      </td>
                      <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">
                        {r.booksTaxable > 0 ? r.booksTaxable.toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '-'}
                      </td>
                      <td className="py-2 px-3 border-r border-gray-200 text-right font-mono font-semibold text-gray-900">
                        {r.booksTotalTax > 0 ? r.booksTotalTax.toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '-'}
                      </td>
                      <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">
                        {r.gstr2bTaxable > 0 ? r.gstr2bTaxable.toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '-'}
                      </td>
                      <td className="py-2 px-3 border-r border-gray-200 text-right font-mono font-semibold text-gray-900">
                        {r.gstr2bTotalTax > 0 ? r.gstr2bTotalTax.toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '-'}
                      </td>
                      <td className={`py-2 px-3 border-r border-gray-200 text-right font-mono font-bold ${r.diffAmount > 0 ? 'text-[#C43D43]' : 'text-gray-900'}`}>
                        {r.diffAmount > 0 ? `₹${r.diffAmount.toFixed(2)}` : '₹0.00'}
                      </td>
                      <td className="py-2 px-3 border-r border-gray-200 whitespace-nowrap">
                        {r.status === 'Exact Match' ? (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-semibold bg-slate-100 text-gray-900 border border-slate-200">
                            <CheckCircle2 className="w-3 h-3" />
                            <span>Exact Match</span>
                          </span>
                        ) : r.status === 'Probable Match' ? (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-semibold bg-blue-50 text-blue-700 border border-blue-200">
                            <CheckCircle2 className="w-3 h-3" />
                            <span>Fuzzy Match</span>
                          </span>
                        ) : r.status === 'Tax Mismatch' ? (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-semibold bg-[#FDF2F2] text-[#C43D43] border border-[#F8D7DA]">
                            <AlertTriangle className="w-3 h-3" />
                            <span>Tax Mismatch</span>
                          </span>
                        ) : r.status === 'Duplicate Record' ? (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-semibold bg-purple-100 text-purple-800 border border-purple-200">
                            <Copy className="w-3 h-3" />
                            <span>Duplicate</span>
                          </span>
                        ) : (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-semibold bg-amber-50 text-amber-800 border border-amber-200">
                            <AlertCircle className="w-3 h-3" />
                            <span>{r.status}</span>
                          </span>
                        )}
                      </td>
                      <td className="py-2 px-3 text-center whitespace-nowrap">
                        <select
                          value={r.resolutionStatus || 'Pending Review'}
                          onChange={(e) => handleSetResolution(r.id, e.target.value)}
                          className="px-2 py-1 bg-white border border-gray-300 rounded text-[11px] font-medium text-gray-800 focus:outline-none focus:border-gray-300"
                        >
                          <option value="Approved for 3B">Approve for 3B</option>
                          <option value="Query Sent to Vendor">Query Vendor</option>
                          <option value="Debit Note Required">Issue Debit Note</option>
                          <option value="Entry Booked">Book in Accounts</option>
                          <option value="Pending Review">Pending Review</option>
                        </select>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
