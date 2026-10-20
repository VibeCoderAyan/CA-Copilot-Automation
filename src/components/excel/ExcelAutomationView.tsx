import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import { exportStyledWorkbook } from '../../utils/excelStyler';
import { extractSmartTableFromSheet, parseDate, parseAmount, isValidGstin } from '../../utils/dataSanitizer';
import {
  UploadCloud,
  FileSpreadsheet,
  CheckCircle2,
  AlertTriangle,
  Download,
  RefreshCw,
  Check
} from 'lucide-react';

interface ExcelAutomationViewProps {
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
  selectedClientId?: string;
}

interface SpreadsheetRow {
  [key: string]: any;
}

export const ExcelAutomationView: React.FC<ExcelAutomationViewProps> = ({
  onShowToast,
  selectedClientId: _selectedClientId,
}) => {
  const [currentFile, setCurrentFile] = useState<{
    name: string;
    size: string;
    sheetNames: string[];
    selectedSheet: string;
    totalRows: number;
    columns: string[];
  } | null>(null);

  const [tableData, setTableData] = useState<SpreadsheetRow[]>([]);
  const [isProcessed, setIsProcessed] = useState(false);
  const [stats, setStats] = useState({
    datesNormalized: 0,
    amountsCleaned: 0,
    duplicatesDetected: 0,
    missingFields: 0,
  });

  const fileInputRef = useRef<HTMLInputElement>(null);

  // File upload handler using SheetJS with smart table extraction
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const buffer = evt.target?.result as ArrayBuffer;
        if (!buffer) throw new Error('Could not read file data');

        const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const table = extractSmartTableFromSheet(worksheet);

        if (table.rows.length > 0) {
          const rows: SpreadsheetRow[] = table.rows.map((r) => ({
            ...r,
            status: 'raw',
          }));

          setCurrentFile({
            name: file.name,
            size: `${Math.round(file.size / 1024)} KB`,
            sheetNames: workbook.SheetNames,
            selectedSheet: sheetName,
            totalRows: table.rows.length,
            columns: table.headers,
          });
          setTableData(rows);
          setIsProcessed(false);
        } else {
          onShowToast('Empty File', 'Could not detect any data rows in spreadsheet.', 'error');
        }
      } catch (err: any) {
        onShowToast('File Read Error', err.message || 'Unable to parse spreadsheet file format.', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Perform genuine clean & normalize automation
  const handleRunAutomation = () => {
    if (tableData.length === 0) {
      return;
    }

    let dNorm = 0;
    let aClean = 0;
    let dupCount = 0;
    let missingCount = 0;
    const seenSignatures = new Set<string>();

    const cleaned = tableData.map((row) => {
      const newRow = { ...row };

      // Identify date keys
      const dateKeys = Object.keys(newRow).filter((k) => /date|dt/i.test(k));
      for (const k of dateKeys) {
        const rawDate = newRow[k];
        if (rawDate !== undefined && rawDate !== null && String(rawDate).trim() !== '') {
          const normalized = parseDate(rawDate, 'DD-MM-YYYY');
          if (normalized && normalized !== String(rawDate).trim()) {
            newRow[k] = normalized;
            dNorm++;
          }
        }
      }

      // Identify amount keys
      const amountKeys = Object.keys(newRow).filter((k) =>
        /amt|amount|taxable|cgst|sgst|igst|tax|total|gross|net|debit|credit|dr|cr|balance/i.test(k)
      );
      for (const k of amountKeys) {
        const rawVal = newRow[k];
        if (rawVal !== undefined && rawVal !== null && typeof rawVal === 'string') {
          const strVal = rawVal.trim();
          if (/[₹$,]|^\s*\([0-9,.]+\)\s*$/.test(strVal)) {
            const parsed = parseAmount(strVal);
            newRow[k] = parsed;
            aClean++;
          }
        }
      }

      // Check duplicates (by Voucher/Bill/Doc No or composite row values)
      const vchKey = Object.keys(newRow).find((k) => /vch|voucher|inv|invoice|bill|doc/i.test(k));
      const sig = vchKey && newRow[vchKey] ? String(newRow[vchKey]).trim().toUpperCase() : JSON.stringify(newRow);
      if (seenSignatures.has(sig)) {
        newRow._duplicate = true;
        dupCount++;
      } else {
        seenSignatures.add(sig);
        newRow._duplicate = false;
      }

      // Check missing GSTIN
      const gstinKey = Object.keys(newRow).find((k) => /gstin|gst/i.test(k));
      if (gstinKey) {
        const gstinVal = String(newRow[gstinKey] || '').trim().toUpperCase();
        if (!gstinVal || !isValidGstin(gstinVal)) {
          newRow._missingGstin = true;
          missingCount++;
        } else {
          newRow._missingGstin = false;
        }
      }

      newRow.status = 'cleaned';
      return newRow;
    });

    setTableData(cleaned);
    setIsProcessed(true);
    setStats({
      datesNormalized: dNorm,
      amountsCleaned: aClean,
      duplicatesDetected: dupCount,
      missingFields: missingCount,
    });

    onShowToast(
      'Spreadsheet Cleaned',
      `Normalized ${dNorm} dates, cleaned ${aClean} amount fields, detected ${dupCount} duplicate(s).`,
      'success'
    );
  };

  // Export cleaned workbook as real XLSX
  const handleExportXLSX = () => {
    const exportRows = tableData.map(({ status: _s, _duplicate, _missingGstin, ...rest }) => rest);
    const fileName = currentFile?.name
      ? currentFile.name.replace(/\.[^/.]+$/, '') + '_Cleaned.xlsx'
      : 'LedgerFlow_Cleaned.xlsx';

    exportStyledWorkbook(
      [{ name: 'Cleaned_Register', data: exportRows }],
      fileName
    );

    onShowToast('File Downloaded', 'Exported standardized spreadsheet with auto-fitted columns and formatted headers.', 'success');
  };

  // Clear loaded spreadsheet
  const handleClear = () => {
    setCurrentFile(null);
    setTableData([]);
    setIsProcessed(false);
    setStats({ datesNormalized: 0, amountsCleaned: 0, duplicatesDetected: 0, missingFields: 0 });
    onShowToast('Spreadsheet Cleared', 'Ready for new client spreadsheet upload.', 'info');
  };

  return (
    <div className="space-y-4">
      {/* 1. Page Header (Section 5) */}
      <div className="bg-white rounded-lg border border-[#E5E7EB] p-5 flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-lg font-semibold text-[#111827] tracking-tight">Excel Cleaner</h2>
          <p className="text-xs text-[#6B7280] mt-0.5">
            Clean and standardize spreadsheets. Fix dates to DD-MM-YYYY, remove negative bracket syntax, and flag duplicates.
          </p>
        </div>

        <div className="flex items-center gap-2.5 shrink-0">
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileUpload}
            accept=".xlsx,.xls,.csv"
            className="hidden"
          />
          <button
            onClick={() => fileInputRef.current?.click()}
            className="px-3.5 py-2 text-xs font-medium bg-gray-900 hover:bg-black text-white rounded-md flex items-center gap-1.5 transition-colors cursor-pointer"
          >
            <UploadCloud className="w-3.5 h-3.5" />
            <span>Upload File</span>
          </button>
        </div>
      </div>

      {/* 2. File Status & Action Bar */}
      <div className="bg-white rounded-lg border border-[#E5E7EB] p-4 flex flex-col md:flex-row md:items-center justify-between gap-3">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 rounded bg-gray-50 text-gray-900 border border-[#E5E7EB] flex items-center justify-center shrink-0">
            <FileSpreadsheet className="w-4 h-4" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <span className="text-xs font-semibold text-[#111827]">
                {currentFile ? currentFile.name : 'No file loaded'}
              </span>
              <span className={`text-[10px] font-medium px-2 py-0.5 rounded border ${
                isProcessed
                  ? 'bg-gray-100 text-gray-900 border-gray-200'
                  : 'bg-gray-50 text-[#6B7280] border-[#E5E7EB]'
              }`}>
                {isProcessed ? 'Standardized' : currentFile ? 'Raw Data' : 'Empty'}
              </span>
            </div>
            <p className="text-[11px] text-[#6B7280] mt-0.5">
              {tableData.length > 0
                ? `${currentFile?.totalRows || tableData.length} records • ${currentFile?.size || 'Loaded'} • Formats: Tally, SAP, CSV`
                : 'No spreadsheet loaded • Upload an Excel (.xlsx, .csv) file to begin'}
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2 shrink-0">
          {tableData.length > 0 && (
            <button
              onClick={handleClear}
              className="px-2.5 py-1.5 text-xs text-[#6B7280] hover:text-[#111827] rounded-md border border-[#E5E7EB] hover:bg-gray-50 flex items-center gap-1.5 transition-colors cursor-pointer"
            >
              <RefreshCw className="w-3 h-3" />
              <span>Clear</span>
            </button>
          )}

          {!isProcessed ? (
            <button
              onClick={handleRunAutomation}
              disabled={tableData.length === 0}
              className="px-3.5 py-1.5 text-xs font-medium bg-gray-900 hover:bg-black text-white rounded-md flex items-center gap-1.5 transition-colors disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
            >
              <Check className="w-3.5 h-3.5" />
              <span>Clean Spreadsheet</span>
            </button>
          ) : (
            <button
              onClick={handleExportXLSX}
              className="px-3.5 py-1.5 text-xs font-medium bg-gray-900 hover:bg-black text-white rounded-md flex items-center gap-1.5 transition-colors cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Download Cleaned File</span>
            </button>
          )}
        </div>
      </div>

      {/* 3. Issues Summary (When available) */}
      {isProcessed && (
        <div className="bg-gray-100 border border-gray-200 rounded-lg p-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs text-gray-900">
          <div className="flex items-center gap-2.5">
            <CheckCircle2 className="w-4 h-4 text-gray-900 shrink-0" />
            <div>
              <span className="font-semibold">Processing Complete: </span>
              <span>
                Standardized {stats.datesNormalized} dates to DD-MM-YYYY, converted {stats.amountsCleaned} negative amounts to numbers, flagged{' '}
                <strong className="text-[#C43D43]">{stats.duplicatesDetected} duplicate voucher</strong> and{' '}
                <strong className="text-[#C43D43]">{stats.missingFields} missing GSTIN</strong>.
              </span>
            </div>
          </div>
          <button
            onClick={handleExportXLSX}
            className="inline-flex items-center gap-1 px-3 py-1 bg-white border border-gray-200 text-gray-900 font-medium rounded text-xs hover:bg-gray-50 transition-colors shrink-0"
          >
            <Download className="w-3 h-3" />
            <span>Download .xlsx</span>
          </button>
        </div>
      )}

      {/* 4. Data Preview Table */}
      <div className="bg-white border border-[#E5E7EB] rounded-lg overflow-hidden">
        <div className="px-4 py-2.5 border-b border-[#E5E7EB] bg-gray-50 flex items-center justify-between text-xs">
          <span className="font-semibold text-[#111827]">Data Preview</span>
          <div className="flex items-center gap-3 text-[11px] text-[#6B7280]">
            <span className="flex items-center gap-1">
              <span className="w-2 h-2 rounded-full bg-[#C43D43]" /> Duplicate / Missing
            </span>
            <span className="flex items-center gap-1">
              <span className="w-2 h-2 rounded-full bg-gray-900" /> Cleaned
            </span>
          </div>
        </div>

        {tableData.length === 0 ? (
          <div className="p-12 text-center">
            <div className="w-12 h-12 rounded-xl bg-gray-50 border border-gray-200 flex items-center justify-center mx-auto mb-3">
              <FileSpreadsheet className="w-6 h-6 text-gray-400" />
            </div>
            <p className="text-sm font-semibold text-gray-900">No Spreadsheet Loaded</p>
            <p className="text-xs text-gray-500 mt-1 max-w-sm mx-auto">
              Upload your client's Excel or CSV sales register, purchase ledger, or raw bank extract to clean dates, format currency, and detect duplicates.
            </p>
            <button
              onClick={() => fileInputRef.current?.click()}
              className="mt-4 px-4 py-2 text-xs font-semibold bg-gray-900 hover:bg-black text-white rounded-lg inline-flex items-center gap-1.5 shadow-sm transition-colors cursor-pointer"
            >
              <UploadCloud className="w-3.5 h-3.5" />
              <span>Upload Spreadsheet File</span>
            </button>
          </div>
        ) : (
          <div className="overflow-x-auto max-h-[460px]">
            <table className="w-full text-left text-xs border-collapse">
              <thead className="sticky top-0 bg-gray-50 border-b border-[#E5E7EB] text-[#6B7280] text-[11px] font-medium z-10">
                <tr>
                  <th className="py-2.5 px-3 border-r border-[#E5E7EB] w-10 text-center">#</th>
                  <th className="py-2.5 px-3 border-r border-[#E5E7EB]">Voucher No</th>
                  <th className="py-2.5 px-3 border-r border-[#E5E7EB]">Invoice Date</th>
                  <th className="py-2.5 px-3 border-r border-[#E5E7EB]">Party Name</th>
                  <th className="py-2.5 px-3 border-r border-[#E5E7EB]">GSTIN</th>
                  <th className="py-2.5 px-3 border-r border-[#E5E7EB] text-right">Taxable Amt</th>
                  <th className="py-2.5 px-3 border-r border-[#E5E7EB] text-right">CGST</th>
                  <th className="py-2.5 px-3 border-r border-[#E5E7EB] text-right">SGST</th>
                  <th className="py-2.5 px-3 border-r border-[#E5E7EB] text-right">IGST</th>
                  <th className="py-2.5 px-3 text-right">Total Invoice</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#E5E7EB] text-[#111827]">
                {tableData.map((row, idx) => {
                  const isDup = row._duplicate;
                  const isMissingGst = row._missingGstin;

                  return (
                    <tr
                      key={idx}
                      className={`transition-colors ${
                        isDup
                          ? 'bg-[#FDF2F2]'
                          : isMissingGst
                          ? 'bg-[#FEF9EE]'
                          : 'hover:bg-gray-50'
                      }`}
                    >
                      <td className="py-2 px-3 border-r border-[#E5E7EB] text-center text-[#6B7280] text-[11px]">
                        {idx + 1}
                      </td>
                      <td className="py-2 px-3 border-r border-[#E5E7EB] font-medium">
                        <div className="flex items-center gap-1.5">
                          <span>{row.Voucher_No}</span>
                          {isDup && (
                            <span className="text-[9px] px-1 py-0.2 rounded bg-[#FDF2F2] text-[#C43D43] border border-[#FDE8E8] font-semibold">
                              DUPLICATE
                            </span>
                          )}
                        </div>
                      </td>
                      <td className="py-2 px-3 border-r border-[#E5E7EB] font-mono text-[11px]">
                        <span className={isProcessed ? 'text-gray-900 font-medium' : 'text-[#6B7280]'}>
                          {row.Invoice_Date}
                        </span>
                      </td>
                      <td className="py-2 px-3 border-r border-[#E5E7EB]">
                        {row.Party_Name}
                      </td>
                      <td className="py-2 px-3 border-r border-[#E5E7EB] font-mono text-[11px]">
                        {row.GSTIN ? (
                          <span>{row.GSTIN}</span>
                        ) : (
                          <span className="text-[#C43D43] font-medium flex items-center gap-1">
                            <AlertTriangle className="w-3 h-3" /> Missing
                          </span>
                        )}
                      </td>
                      <td className="py-2 px-3 border-r border-[#E5E7EB] text-right font-mono tabular-nums">
                        {typeof row.Taxable_Amt === 'number'
                          ? row.Taxable_Amt.toLocaleString('en-IN', { minimumFractionDigits: 2 })
                          : row.Taxable_Amt}
                      </td>
                      <td className="py-2 px-3 border-r border-[#E5E7EB] text-right font-mono tabular-nums">
                        {typeof row.CGST === 'number'
                          ? row.CGST.toLocaleString('en-IN', { minimumFractionDigits: 2 })
                          : row.CGST}
                      </td>
                      <td className="py-2 px-3 border-r border-[#E5E7EB] text-right font-mono tabular-nums">
                        {typeof row.SGST === 'number'
                          ? row.SGST.toLocaleString('en-IN', { minimumFractionDigits: 2 })
                          : row.SGST}
                      </td>
                      <td className="py-2 px-3 border-r border-[#E5E7EB] text-right font-mono tabular-nums">
                        {typeof row.IGST === 'number'
                          ? row.IGST.toLocaleString('en-IN', { minimumFractionDigits: 2 })
                          : row.IGST}
                      </td>
                      <td className="py-2 px-3 text-right font-mono tabular-nums font-semibold">
                        {typeof row.Total_Amt === 'number'
                          ? row.Total_Amt.toLocaleString('en-IN', { minimumFractionDigits: 2 })
                          : row.Total_Amt}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
};
