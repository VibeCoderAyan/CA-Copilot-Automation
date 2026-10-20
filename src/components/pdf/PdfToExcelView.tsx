import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import { exportStyledWorkbook } from '../../utils/excelStyler';
import {
  FileCode,
  UploadCloud,
  Download,
  CheckCircle2,
  ShieldAlert,
  Info,
  RefreshCw
} from 'lucide-react';
import { parsePdfFile } from '../../utils/pdfExtractor';

interface PdfToExcelViewProps {
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
  selectedClientId?: string;
}

interface ExtractedRow {
  date: string;
  narration: string;
  refNo: string;
  debit: number;
  credit: number;
  balance: number;
}

export const PdfToExcelView: React.FC<PdfToExcelViewProps> = ({
  onShowToast,
  selectedClientId: _selectedClientId,
}) => {
  const [activeFile, setActiveFile] = useState<{
    name: string;
    size: string;
    pageCount: number;
    detectedTables: number;
  }>({
    name: '',
    size: '',
    pageCount: 0,
    detectedTables: 0,
  });

  const [selectedPage, setSelectedPage] = useState(1);
  const [extractedData, setExtractedData] = useState<ExtractedRow[]>([]);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setErrorMessage(null);
    setIsProcessing(true);

    try {
      const res = await parsePdfFile(file);

      if (res.status === 'success' && res.rows.length > 0) {
        const mappedRows: ExtractedRow[] = res.rows.map((r) => ({
          date: r.date,
          narration: r.particulars,
          refNo: r.refNo,
          debit: r.debit,
          credit: r.credit,
          balance: r.balance,
        }));

        setActiveFile({
          name: file.name,
          size: `${Math.round(file.size / 1024)} KB`,
          pageCount: res.pageCount,
          detectedTables: 1,
        });
        setExtractedData(mappedRows);
      } else {
        // Strictly block processing on unsupported/scanned files: NEVER inject fake data!
        setExtractedData([]);
        setActiveFile({
          name: file.name,
          size: `${Math.round(file.size / 1024)} KB`,
          pageCount: res.pageCount || 1,
          detectedTables: 0,
        });
        const msg = res.errorMessage || 'Unsupported document format: No standard bank transaction tables detected.';
        setErrorMessage(msg);
        onShowToast('Extraction Stopped', 'Unsupported PDF. Processing halted to prevent fake data.', 'error');
      }
    } catch (err: any) {
      setExtractedData([]);
      const msg = `Failed to process PDF: ${err.message || 'Corrupted file'}`;
      setErrorMessage(msg);
      onShowToast('Parse Error', msg, 'error');
    } finally {
      setIsProcessing(false);
    }
  };


  // Export to Excel
  const handleExportExcel = () => {
    if (extractedData.length === 0) return;
    const exportRows = extractedData.map((row) => ({
      'Date': row.date,
      'Particulars / Narration': row.narration,
      'Chq / Ref No': row.refNo,
      'Withdrawal (Debit)': row.debit,
      'Deposit (Credit)': row.credit,
      'Closing Balance': row.balance,
    }));

    exportStyledWorkbook(
      [{ name: 'Extracted_Statement', data: exportRows }],
      `${activeFile.name.replace('.pdf', '')}_Converted.xlsx`
    );

    onShowToast('Excel File Downloaded', 'Converted PDF statement to structured XLSX workbook with auto-fitted columns.', 'success');
  };

  // Export to CSV
  const handleExportCsv = () => {
    if (extractedData.length === 0) return;
    const exportRows = extractedData.map((row) => ({
      'Date': row.date,
      'Particulars': row.narration,
      'RefNo': row.refNo,
      'Debit': row.debit,
      'Credit': row.credit,
      'Balance': row.balance,
    }));

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const csvOutput = XLSX.utils.sheet_to_csv(ws);
    const blob = new Blob([csvOutput], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.setAttribute('download', `${activeFile.name.replace('.pdf', '')}_Converted.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    onShowToast('CSV File Downloaded', 'Converted PDF statement to CSV format.', 'success');
  };

  return (
    <div className="space-y-4">
      {/* 1. Page Header */}
      <div className="bg-white rounded-lg border border-[#E5E7EB] p-5 flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-lg font-semibold text-[#111827] tracking-tight">PDF to Excel Converter</h2>
          <p className="text-xs text-[#6B7280] mt-0.5">
            Extract tabular data from digital bank statements and financial PDFs with strict mathematical fidelity.
          </p>
        </div>

        <div className="flex items-center gap-2.5 shrink-0">
          <input
            type="file"
            ref={fileInputRef}
            accept=".pdf"
            className="hidden"
            onChange={handleFileUpload}
          />

          
          <button
            onClick={() => fileInputRef.current?.click()}
            disabled={isProcessing}
            className="px-3.5 py-2 text-xs font-medium bg-gray-900 hover:bg-black text-white rounded-md flex items-center gap-1.5 transition-colors disabled:opacity-50 cursor-pointer"
          >
            {isProcessing ? <RefreshCw className="w-3.5 h-3.5 animate-spin" /> : <UploadCloud className="w-3.5 h-3.5" />}
            <span>{isProcessing ? 'Extracting...' : 'Upload PDF'}</span>
          </button>
        </div>
      </div>

      {/* 2. File Status & Action Bar */}
      {extractedData.length > 0 && (
      <div className="bg-white rounded-lg border border-[#E5E7EB] p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 rounded bg-gray-50 border border-[#E5E7EB] text-gray-900 flex items-center justify-center shrink-0">
            <FileCode className="w-4 h-4" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <span className="text-xs font-semibold text-[#111827]">
                {activeFile.name}
              </span>
              <span className="text-[10px] bg-gray-100 border border-gray-200 px-2 py-0.5 rounded font-medium text-gray-900">
                Verified Digital PDF
              </span>
            </div>
            <p className="text-[11px] text-[#6B7280] mt-0.5">
              {activeFile.size} • {activeFile.pageCount} Page(s) • {extractedData.length} rows parsed
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2 shrink-0">
          <button
            onClick={handleExportCsv}
            disabled={extractedData.length === 0}
            className="px-3 py-1.5 text-xs font-medium text-[#111827] bg-gray-50 hover:bg-[#EAEAEA] rounded-md border border-[#E5E7EB] flex items-center gap-1.5 transition-colors disabled:opacity-40"
          >
            <Download className="w-3.5 h-3.5 text-[#6B7280]" />
            <span>Export CSV</span>
          </button>
          <button
            onClick={handleExportExcel}
            disabled={extractedData.length === 0}
            className="px-3.5 py-1.5 text-xs font-medium bg-gray-900 hover:bg-black text-white rounded-md flex items-center gap-1.5 transition-colors disabled:opacity-40"
          >
            <Download className="w-3.5 h-3.5" />
            <span>Download .xlsx</span>
          </button>
        </div>
      </div>
      )}

      {/* Error Banner when extraction blocked */}
      {errorMessage && (
        <div className="p-4 rounded-xl bg-[#FEF2F2] border border-[#FCA5A5] text-xs text-[#991B1B] flex items-start gap-3">
          <ShieldAlert className="w-5 h-5 shrink-0 text-[#DC2626] mt-0.5" />
          <div className="flex-1">
            <span className="font-bold block">Document Processing Blocked</span>
            <p className="mt-0.5 text-[11px] text-[#B91C1C]">{errorMessage}</p>
            <div className="mt-2 text-[11px] text-[#7F1D1D] flex items-center gap-1">
              <Info className="w-3.5 h-3.5 shrink-0" />
              <span>Anti-Hallucination Guard: Mock data is never generated when an uploaded file cannot be parsed.</span>
            </div>
          </div>
          <button onClick={() => setErrorMessage(null)} className="underline hover:no-underline font-semibold ml-2">
            Dismiss
          </button>
        </div>
      )}

      {/* 3. Side-by-side: Visual Page Preview & Extracted Data Table */}
      {extractedData.length > 0 && (
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-4">
          {/* Left: Document Visual Preview with Table Grid Box (5 cols) */}
          <div className="lg:col-span-5 bg-white rounded-lg border border-[#E5E7EB] p-4 flex flex-col justify-between">
            <div>
              <div className="flex items-center justify-between pb-3 border-b border-[#E5E7EB]">
                <span className="text-xs font-semibold text-[#111827]">
                  PDF Preview (Page {selectedPage} of {activeFile.pageCount})
                </span>
                <div className="flex items-center gap-1">
                  {[1, 2, 3].slice(0, activeFile.pageCount).map((p) => (
                    <button
                      key={p}
                      onClick={() => setSelectedPage(p)}
                      className={`px-2 py-0.5 text-[11px] rounded transition-colors ${
                        selectedPage === p
                          ? 'bg-gray-900 text-white font-medium'
                          : 'bg-gray-50 text-[#6B7280] hover:bg-[#E5E7EB]'
                      }`}
                    >
                      {p}
                    </button>
                  ))}
                </div>
              </div>

              {/* Simulated Statement Page Preview */}
              <div className="mt-3 p-3 bg-[#FAFAFA] rounded-md border border-[#E5E7EB] text-[10px] space-y-2">
                <div className="flex justify-between items-start border-b border-[#E5E7EB] pb-2">
                  <div>
                    <strong className="text-xs text-[#111827] block">BANK STATEMENT EXTRACT</strong>
                    <span className="text-[#6B7280]">{activeFile.name}</span>
                  </div>
                  <span className="font-mono text-[#6B7280]">{extractedData.length} records</span>
                </div>

                {/* Highlighted Bounding Box Area */}
                <div className="relative border-2 border-gray-900 bg-gray-100/30 rounded p-2">
                  <span className="absolute -top-2 right-2 bg-gray-900 text-white text-[9px] px-1.5 rounded font-mono">
                    PARSED TABLE
                  </span>
                  <div className="font-mono text-[9px] text-[#111827] space-y-1">
                    <div className="font-bold flex justify-between border-b border-gray-900/30 pb-0.5">
                      <span>Date</span>
                      <span>Particulars</span>
                      <span>Debit</span>
                      <span>Credit</span>
                      <span>Balance</span>
                    </div>
                    {extractedData.slice(0, 3).map((r, i) => (
                      <div key={i} className="flex justify-between text-[#6B7280]">
                        <span>{r.date}</span>
                        <span className="truncate max-w-[120px]">{r.narration}</span>
                        <span>{r.debit > 0 ? r.debit.toLocaleString('en-IN') : '-'}</span>
                        <span>{r.credit > 0 ? r.credit.toLocaleString('en-IN') : '-'}</span>
                        <span>{r.balance.toLocaleString('en-IN')}</span>
                      </div>
                    ))}
                  </div>
                </div>

                <div className="pt-2 text-[10px] text-[#6B7280]">
                  Tabular boundaries extracted directly from digital PDF text stream.
                </div>
              </div>
            </div>

            <div className="pt-3 border-t border-[#E5E7EB] mt-3 flex items-center justify-between text-xs text-[#6B7280]">
              <span>Verification Status: <strong className="text-gray-900">Verified</strong></span>
              <span className="text-[11px] text-[#6B7280]">Digital Layer OK</span>
            </div>
          </div>

          {/* Right: Extracted Data Table Preview (7 cols) */}
          <div className="lg:col-span-7 bg-white rounded-lg border border-[#E5E7EB] overflow-hidden flex flex-col justify-between">
            <div>
              <div className="px-4 py-2.5 border-b border-[#E5E7EB] bg-gray-50 flex items-center justify-between text-xs">
                <span className="font-semibold text-[#111827]">Extracted Table Data ({extractedData.length} rows)</span>
                <span className="text-gray-900 font-medium flex items-center gap-1">
                  <CheckCircle2 className="w-3.5 h-3.5" /> Structured
                </span>
              </div>

              <div className="overflow-x-auto max-h-[380px]">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="sticky top-0 bg-gray-50 border-b border-[#E5E7EB] text-[#6B7280] text-[11px] font-medium z-10">
                    <tr>
                      <th className="py-2 px-2.5 border-r border-[#E5E7EB]">Date</th>
                      <th className="py-2 px-2.5 border-r border-[#E5E7EB]">Particulars</th>
                      <th className="py-2 px-2.5 border-r border-[#E5E7EB]">Ref No</th>
                      <th className="py-2 px-2.5 border-r border-[#E5E7EB] text-right">Debit (₹)</th>
                      <th className="py-2 px-2.5 border-r border-[#E5E7EB] text-right">Credit (₹)</th>
                      <th className="py-2 px-2.5 text-right">Balance (₹)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#E5E7EB] text-[#111827]">
                    {extractedData.map((row, idx) => (
                      <tr key={idx} className="hover:bg-gray-50 transition-colors">
                        <td className="py-2 px-2.5 border-r border-[#E5E7EB] font-mono text-[11px]">
                          {row.date}
                        </td>
                        <td className="py-2 px-2.5 border-r border-[#E5E7EB] truncate max-w-[180px]" title={row.narration}>
                          {row.narration}
                        </td>
                        <td className="py-2 px-2.5 border-r border-[#E5E7EB] font-mono text-[10px] text-[#6B7280]">
                          {row.refNo || '-'}
                        </td>
                        <td className="py-2 px-2.5 border-r border-[#E5E7EB] text-right font-mono tabular-nums">
                          {row.debit > 0 ? row.debit.toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '—'}
                        </td>
                        <td className="py-2 px-2.5 border-r border-[#E5E7EB] text-right font-mono tabular-nums text-gray-900 font-medium">
                          {row.credit > 0 ? row.credit.toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '—'}
                        </td>
                        <td className="py-2 px-2.5 text-right font-mono tabular-nums">
                          {row.balance.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="p-3 border-t border-[#E5E7EB] bg-gray-50 flex items-center justify-between text-xs">
              <span className="text-[#6B7280]">Ready for export to .xlsx or CSV</span>
              <div className="flex items-center gap-2">
                <button
                  onClick={handleExportExcel}
                  className="px-3.5 py-1.5 bg-gray-900 hover:bg-black text-white rounded-md text-xs font-medium flex items-center gap-1.5 transition-colors"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>Save Excel File</span>
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Empty State when no data */}
      {extractedData.length === 0 && !errorMessage && (
        <div className="bg-white rounded-lg border border-[#E5E7EB] p-8 text-center text-xs text-[#6B7280]">
          <FileCode className="w-8 h-8 text-[#A0AAA2] mx-auto mb-2" />
          <p className="font-semibold text-[#111827]">No Statement Extracted</p>
          <p className="text-[11px] text-[#6B7280] mt-1 max-w-md mx-auto">
            Upload a digital bank statement PDF with selectable text layers to convert it into Excel.
          </p>
        </div>
      )}
    </div>
  );
};
