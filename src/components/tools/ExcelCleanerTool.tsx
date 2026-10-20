import React, { useState, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  UploadCloud,
  Download,
  CheckCircle2,
  AlertTriangle,
  FileSpreadsheet,
  RefreshCw,
  Sliders,
  ShieldCheck,
  Eye,
  FileText,
  Layers,
  Sparkles,
  Info
} from 'lucide-react';
import {
  readWorkbookPreserved,
  cloneWorkbook,
  analyzeWorksheet,
  applyCleaningOperations,
  appendAuditLogSheet,
  type CleaningOptions,
  type SheetAnalysis,
  type ExecutionReport,
  DEFAULT_CLEANING_OPTIONS,
} from '../../utils/workbookPreserver';
import { DataFormatGuide } from '../common/DataFormatGuide';
import { ToolHeader } from '../common/ToolHeader';

interface ExcelCleanerToolProps {
  onBack: () => void;
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
}

export const ExcelCleanerTool: React.FC<ExcelCleanerToolProps> = ({ onBack, onShowToast }) => {
  const [file, setFile] = useState<{ name: string; size: string } | null>(null);
  const [originalWb, setOriginalWb] = useState<XLSX.WorkBook | null>(null);
  const [processedWb, setProcessedWb] = useState<XLSX.WorkBook | null>(null);

  // Selected sheet for inspection & cleaning
  const [selectedSheet, setSelectedSheet] = useState<string>('');
  const [previewSheet, setPreviewSheet] = useState<string>('');

  // User-controlled options
  const [options, setOptions] = useState<CleaningOptions>({ ...DEFAULT_CLEANING_OPTIONS });
  const [includeAuditSheet, setIncludeAuditSheet] = useState<boolean>(true);

  // Pre-processing analysis
  const [analysis, setAnalysis] = useState<SheetAnalysis | null>(null);

  // Post-processing execution report
  const [executionReport, setExecutionReport] = useState<ExecutionReport | null>(null);
  const [isProcessed, setIsProcessed] = useState(false);

  // Preview filtering
  const [activeFilter, setActiveFilter] = useState<'all' | 'modified' | 'duplicates'>('all');
  const [validationError, setValidationError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Handle spreadsheet file upload (.xlsx, .xls, .csv)
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;

    setValidationError(null);
    const validExtensions = ['.xlsx', '.xls', '.csv'];
    const fileNameLower = f.name.toLowerCase();
    const isValid = validExtensions.some((ext) => fileNameLower.endsWith(ext));

    if (!isValid) {
      setValidationError(`Unsupported file type: ${f.name}. Please upload a valid .xlsx, .xls, or .csv spreadsheet.`);
      onShowToast('Invalid Format', 'Supported formats: .xlsx, .xls, .csv', 'error');
      return;
    }

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const buffer = evt.target?.result;
        if (!buffer) throw new Error('Could not read file data');

        const wb = readWorkbookPreserved(buffer as ArrayBuffer);
        if (!wb.SheetNames || wb.SheetNames.length === 0) {
          throw new Error('Spreadsheet contains zero worksheets.');
        }

        const firstSheet = wb.SheetNames[0];
        setFile({ name: f.name, size: `${Math.round(f.size / 1024)} KB` });
        setOriginalWb(wb);
        setProcessedWb(null);
        setSelectedSheet(firstSheet);
        setPreviewSheet(firstSheet);
        setIsProcessed(false);
        setExecutionReport(null);

        // Run prospective analysis on first sheet
        const sheetAnalysis = analyzeWorksheet(wb.Sheets[firstSheet], firstSheet, options);
        setAnalysis(sheetAnalysis);
      } catch (err: any) {
        setValidationError(`Failed to parse file: ${err.message || 'Corrupted spreadsheet'}.`);
        onShowToast('File Parse Error', 'Could not read spreadsheet contents.', 'error');
      }
    };
    reader.readAsArrayBuffer(f);
  };

  // Switch target sheet for pre-analysis
  const handleSheetChange = (sheetName: string) => {
    setSelectedSheet(sheetName);
    setPreviewSheet(sheetName);
    if (originalWb && sheetName !== 'all') {
      const ws = originalWb.Sheets[sheetName];
      if (ws) {
        const a = analyzeWorksheet(ws, sheetName, options);
        setAnalysis(a);
      }
    }
  };

  // Toggle option handler
  const handleOptionToggle = (key: keyof CleaningOptions) => {
    const updated = { ...options, [key]: !options[key] };
    setOptions(updated);

    // Re-run prospective analysis with updated options
    if (originalWb && selectedSheet && selectedSheet !== 'all') {
      const ws = originalWb.Sheets[selectedSheet];
      if (ws) {
        setAnalysis(analyzeWorksheet(ws, selectedSheet, updated));
      }
    }
  };

  // Execute user-selected cleaning operations
  const handleClean = () => {
    if (!originalWb) {
      setValidationError('Please upload an Excel file (.xlsx, .xls, .csv) before processing.');
      return;
    }

    const { modifiedWorkbook, report } = applyCleaningOperations(originalWb, selectedSheet, options);

    if (includeAuditSheet) {
      appendAuditLogSheet(modifiedWorkbook, report, file?.name || 'Spreadsheet');
    }

    setProcessedWb(modifiedWorkbook);
    setExecutionReport(report);
    setIsProcessed(true);

    // If target sheet was 'all', preview first sheet
    const targetPreview = selectedSheet === 'all' ? modifiedWorkbook.SheetNames[0] : selectedSheet;
    setPreviewSheet(targetPreview);

    onShowToast(
      'Changes Applied Safely',
      `Processed ${report.sheetsProcessed.length} sheet(s): ${report.totalDuplicatesRemoved} duplicates removed, ${report.totalCellsTrimmed} spaces trimmed, ${report.formulasPreserved} formulas preserved.`,
      'success'
    );
  };

  // Download the exact modified workbook preserving original structure and requested changes
  const handleDownload = () => {
    if (!processedWb) return;
    const originalName = file?.name || 'Workbook.xlsx';
    const baseName = originalName.replace(/\.[^/.]+$/, '');
    const outName = `Cleaned_${baseName}.xlsx`;

    XLSX.writeFile(processedWb, outName);
    onShowToast('Workbook Downloaded', `Saved ${outName} preserving all original sheets, formulas & styles.`, 'success');
  };

  // Convert current preview sheet to rows array for table rendering
  const activeWb = processedWb || originalWb;
  const currentPreviewData = useMemo(() => {
    if (!activeWb || !previewSheet || !activeWb.Sheets[previewSheet]) return [];
    const ws = activeWb.Sheets[previewSheet];
    // Return array of rows
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as any[][];
  }, [activeWb, previewSheet]);

  const previewHeaders = currentPreviewData.length > 0 ? currentPreviewData[0] : [];
  const previewDataRows = currentPreviewData.slice(1);

  // Filtered rows for preview
  const displayRows = useMemo(() => {
    if (!isProcessed || !executionReport) return previewDataRows;
    if (activeFilter === 'modified') {
      const changedCells = new Set(executionReport.changes.map((c) => c.cell));
      return previewDataRows.filter((row, rIdx) => {
        return row.some((_, cIdx) => {
          const addr = XLSX.utils.encode_cell({ r: rIdx + 1, c: cIdx });
          return changedCells.has(addr);
        });
      });
    }
    return previewDataRows;
  }, [previewDataRows, isProcessed, executionReport, activeFilter]);

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="excel"
        title="Excel Sheet Cleaner"
        description="Preserves original formulas, multiple sheets, merges & layout. Modifies only user-selected operations."
        onBack={onBack}
        actionButtons={
          processedWb && (
            <button
              onClick={handleDownload}
              className="px-4 py-2 text-xs font-semibold text-white bg-gray-900 hover:bg-black rounded-lg shadow-sm flex items-center gap-2 transition-colors cursor-pointer"
            >
              <Download className="w-4 h-4" />
              <span>Download Cleaned Excel (.xlsx)</span>
            </button>
          )
        }
      />

      {/* 2. Format Requirements Guide */}
      <DataFormatGuide
        toolType="excel-cleaner"
        acceptedFormats=".xlsx, .xls, .csv"
        requiredFields={[
          { name: 'Voucher / Invoice No', description: 'Unique identifier for transaction', aliases: ['Voucher_No', 'Vch_No', 'Invoice_No', 'Bill_No', 'Doc_No'] },
          { name: 'Invoice / Entry Date', description: 'Transaction date in any format', aliases: ['Date', 'Invoice_Date', 'Bill_Date', 'Txn_Date', 'Vch_Date'] },
          { name: 'Party / Account Name', description: 'Supplier or Customer name', aliases: ['Party_Name', 'Vendor', 'Customer', 'Account_Name', 'Particulars'] },
          { name: 'GSTIN (Optional)', description: '15-character GST identification number', aliases: ['GSTIN', 'GST_No', 'Supplier_GSTIN', 'Party_GSTIN'] },
          { name: 'Amount / Taxable', description: 'Financial values in any currency format', aliases: ['Taxable_Amt', 'Amount', 'Total_Amt', 'Debit', 'Credit'] },
        ]}
        notes={[
          'Core Guarantee: Original multi-sheet structure, formulas (=SUM), merged headers, and unselected data are 100% preserved.',
          'Zero Silent Changes: Only checked operations are applied. Numeric values are never altered without user selection.',
          'Format Preservation: Output is exported in native .xlsx with auto-fitted column widths preventing "###" display.',
        ]}
      />

      {/* 3. Validation Error Banner */}
      {validationError && (
        <div className="p-4 rounded-xl bg-[#FDF2F2] border border-[#F8D7DA] text-xs text-[#9B2C2C] flex items-center justify-between">
          <div className="flex items-center gap-2">
            <AlertTriangle className="w-4 h-4 shrink-0" />
            <span>{validationError}</span>
          </div>
          <button
            onClick={() => setValidationError(null)}
            className="text-xs font-semibold underline hover:no-underline"
          >
            Dismiss
          </button>
        </div>
      )}

      {/* 4. Upload Dropzone / Target Selection Bar */}
      {!file ? (
        <div className="bg-white border-2 border-dashed border-gray-200 rounded-xl p-8 sm:p-10 text-center hover:border-gray-300 transition-colors shadow-2xs">
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileUpload}
            accept=".xlsx,.xls,.csv"
            className="hidden"
            id="excel-cleaner-upload"
          />
          <label htmlFor="excel-cleaner-upload" className="cursor-pointer flex flex-col items-center">
            <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mb-3 shadow-2xs">
              <UploadCloud className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-semibold text-gray-900">Upload Client Excel Spreadsheet</h3>
            <p className="text-xs text-gray-500 mt-1 max-w-md mx-auto leading-relaxed">
              Upload spreadsheets (.xlsx, .xls, .csv). Preserves original formulas, multiple sheets, merges, and layout.
            </p>
            <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
              <UploadCloud className="w-4 h-4" />
              <span>Upload Spreadsheet File</span>
            </div>
            <span className="text-[11px] text-gray-400 font-medium mt-3 block">
              Formulas (=SUM), cell alignments, and multi-sheet structures are 100% protected
            </span>
          </label>
        </div>
      ) : (
        <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-sm space-y-4">
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileUpload}
            accept=".xlsx,.xls,.csv"
            className="hidden"
          />

          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-lg bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 shrink-0">
                <FileSpreadsheet className="w-5 h-5" />
              </div>
              <div>
                <span className="text-xs font-semibold text-gray-900 block">
                  {file.name}
                </span>
                <span className="text-[11px] text-gray-600 block mt-0.5">
                  {file.size} • {originalWb ? `${originalWb.SheetNames.length} sheet(s) detected` : ''} • Formulas and formatting protected
                </span>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={() => fileInputRef.current?.click()}
                className="px-3.5 py-1.5 text-xs text-gray-700 bg-white hover:bg-gray-50 rounded-lg border border-gray-300 flex items-center gap-1.5 transition-colors font-medium shadow-2xs"
              >
                <UploadCloud className="w-3.5 h-3.5 text-gray-600" />
                <span>Change File</span>
              </button>
            </div>
          </div>

        {/* Sheet Selector (if workbook has multiple sheets) */}
        {originalWb && originalWb.SheetNames.length > 0 && (
          <div className="pt-3 border-t border-gray-200 flex flex-wrap items-center justify-between gap-3 text-xs">
            <div className="flex items-center gap-2">
              <Layers className="w-4 h-4 text-gray-900" />
              <span className="font-semibold text-gray-900">Target Sheet for Processing:</span>
              <select
                value={selectedSheet}
                onChange={(e) => handleSheetChange(e.target.value)}
                className="px-2.5 py-1.5 bg-gray-50 border border-gray-200 rounded-lg text-xs font-medium text-gray-900 focus:outline-none focus:border-gray-300"
              >
                <option value="all">⚡ All Worksheets ({originalWb.SheetNames.length} sheets)</option>
                {originalWb.SheetNames.map((s) => (
                  <option key={s} value={s}>
                    📄 Sheet: {s}
                  </option>
                ))}
              </select>
            </div>

            {analysis && (
              <div className="flex items-center gap-2 text-[11px] text-gray-600">
                <span className="bg-gray-50 px-2 py-0.5 rounded border border-gray-200">
                  Range: <strong>{analysis.range}</strong>
                </span>
                <span className="bg-gray-100 text-gray-900 px-2 py-0.5 rounded border border-gray-200 font-medium">
                  {analysis.formulaCells.length} Formulas Protected
                </span>
                <span className="bg-gray-50 px-2 py-0.5 rounded border border-gray-200">
                  {analysis.mergedRangesCount} Merged Ranges
                </span>
              </div>
            )}
          </div>
        )}
      </div>
      )}

      {/* 5. User-Controlled Operations Panel */}
      {originalWb && (
        <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-sm space-y-4">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 pb-3 border-b border-gray-200">
            <div className="flex items-center gap-2">
              <Sliders className="w-4 h-4 text-gray-900" />
              <h2 className="text-sm font-bold text-gray-900">User-Controlled Changes</h2>
              <span className="text-[11px] text-gray-600">
                (Only selected operations will be applied. Unchecked fields remain 100% untouched.)
              </span>
            </div>

            <div className="flex items-center gap-2">
              <label className="flex items-center gap-1.5 text-xs text-gray-900 cursor-pointer">
                <input
                  type="checkbox"
                  checked={includeAuditSheet}
                  onChange={(e) => setIncludeAuditSheet(e.target.checked)}
                  className="rounded text-gray-900 focus:ring-gray-900"
                />
                <span className="font-medium">Append Audit Log Sheet</span>
              </label>
            </div>
          </div>

          {/* Operation Checkboxes Grid */}
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
            {/* Op 1: Trim Whitespace */}
            <div
              onClick={() => handleOptionToggle('trimWhitespace')}
              className={`p-3 rounded-lg border cursor-pointer transition-colors ${
                options.trimWhitespace ? 'bg-slate-50 border-[#CBD5E1]' : 'bg-[#FAFAFA] border-gray-200 opacity-75'
              }`}
            >
              <div className="flex items-start gap-2.5">
                <input
                  type="checkbox"
                  checked={options.trimWhitespace}
                  onChange={() => {}}
                  className="mt-0.5 rounded text-gray-900 focus:ring-gray-900"
                />
                <div>
                  <span className="text-xs font-semibold text-gray-900 block">
                    Clean Extra Spaces & Non-Printable Chars
                  </span>
                  <span className="text-[11px] text-gray-600 block mt-0.5">
                    Trims leading/trailing spaces, replaces NBSP (\u00A0) with standard spaces.
                  </span>
                  {analysis && analysis.cellsWithWhitespace.length > 0 && (
                    <span className="inline-block mt-1 text-[10px] font-semibold text-gray-900 bg-gray-100 px-1.5 py-0.5 rounded">
                      {analysis.cellsWithWhitespace.length} cell(s) will be trimmed
                    </span>
                  )}
                </div>
              </div>
            </div>

            {/* Op 2: Remove Duplicate Records */}
            <div
              onClick={() => handleOptionToggle('removeDuplicates')}
              className={`p-3 rounded-lg border cursor-pointer transition-colors ${
                options.removeDuplicates ? 'bg-slate-50 border-[#CBD5E1]' : 'bg-[#FAFAFA] border-gray-200 opacity-75'
              }`}
            >
              <div className="flex items-start gap-2.5">
                <input
                  type="checkbox"
                  checked={options.removeDuplicates}
                  onChange={() => {}}
                  className="mt-0.5 rounded text-gray-900 focus:ring-gray-900"
                />
                <div>
                  <span className="text-xs font-semibold text-gray-900 block">
                    Remove Duplicate Records
                  </span>
                  <span className="text-[11px] text-gray-600 block mt-0.5">
                    Deletes redundant voucher rows, keeping first instance and safely shifting merges.
                  </span>
                  {analysis && analysis.duplicateRowIndices.length > 0 && (
                    <span className="inline-block mt-1 text-[10px] font-semibold text-[#C43D43] bg-[#FDF2F2] px-1.5 py-0.5 rounded">
                      {analysis.duplicateRowIndices.length} duplicate row(s) detected
                    </span>
                  )}
                </div>
              </div>
            </div>

            {/* Op 3: Standardize Dates */}
            <div
              className={`p-3 rounded-lg border transition-colors ${
                options.standardizeDates ? 'bg-slate-50 border-[#CBD5E1]' : 'bg-[#FAFAFA] border-gray-200 opacity-75'
              }`}
            >
              <div className="flex items-start gap-2.5">
                <input
                  type="checkbox"
                  checked={options.standardizeDates}
                  onChange={() => handleOptionToggle('standardizeDates')}
                  className="mt-0.5 rounded text-gray-900 focus:ring-gray-900"
                />
                <div className="flex-1">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-semibold text-gray-900">Standardize Dates</span>
                    {options.standardizeDates && (
                      <select
                        value={options.dateFormat}
                        onChange={(e) => setOptions({ ...options, dateFormat: e.target.value as any })}
                        onClick={(e) => e.stopPropagation()}
                        className="text-[10px] font-semibold bg-white border border-[#CBD5E1] rounded px-1.5 py-0.5 text-gray-900"
                      >
                        <option value="DD-MM-YYYY">DD-MM-YYYY (Indian)</option>
                        <option value="YYYY-MM-DD">YYYY-MM-DD (ISO)</option>
                        <option value="DD/MM/YYYY">DD/MM/YYYY</option>
                      </select>
                    )}
                  </div>
                  <span className="text-[11px] text-gray-600 block mt-0.5">
                    Normalizes inconsistent date formats into uniform standard dates.
                  </span>
                  {analysis && analysis.unstandardizedDates.length > 0 && (
                    <span className="inline-block mt-1 text-[10px] font-semibold text-gray-900 bg-gray-100 px-1.5 py-0.5 rounded">
                      {analysis.unstandardizedDates.length} date(s) will be formatted
                    </span>
                  )}
                </div>
              </div>
            </div>

            {/* Op 4: Standardize Numbers & Currency */}
            <div
              onClick={() => handleOptionToggle('standardizeNumbers')}
              className={`p-3 rounded-lg border cursor-pointer transition-colors ${
                options.standardizeNumbers ? 'bg-slate-50 border-[#CBD5E1]' : 'bg-[#FAFAFA] border-gray-200 opacity-75'
              }`}
            >
              <div className="flex items-start gap-2.5">
                <input
                  type="checkbox"
                  checked={options.standardizeNumbers}
                  onChange={() => {}}
                  className="mt-0.5 rounded text-gray-900 focus:ring-gray-900"
                />
                <div>
                  <span className="text-xs font-semibold text-gray-900 block">
                    Standardize Numbers & Indian Currency
                  </span>
                  <span className="text-[11px] text-gray-600 block mt-0.5">
                    Converts text currency strings (e.g. ₹ 1,50,000, (45,000)) to clean numbers with #,##0.00 format.
                  </span>
                  {analysis && analysis.unstandardizedNumbers.length > 0 && (
                    <span className="inline-block mt-1 text-[10px] font-semibold text-gray-900 bg-gray-100 px-1.5 py-0.5 rounded">
                      {analysis.unstandardizedNumbers.length} number(s) to convert
                    </span>
                  )}
                </div>
              </div>
            </div>

            {/* Op 5: Remove Blank Rows */}
            <div
              onClick={() => handleOptionToggle('removeBlankRows')}
              className={`p-3 rounded-lg border cursor-pointer transition-colors ${
                options.removeBlankRows ? 'bg-slate-50 border-[#CBD5E1]' : 'bg-[#FAFAFA] border-gray-200 opacity-75'
              }`}
            >
              <div className="flex items-start gap-2.5">
                <input
                  type="checkbox"
                  checked={options.removeBlankRows}
                  onChange={() => {}}
                  className="mt-0.5 rounded text-gray-900 focus:ring-gray-900"
                />
                <div>
                  <span className="text-xs font-semibold text-gray-900 block">
                    Remove Completely Blank Rows
                  </span>
                  <span className="text-[11px] text-gray-600 block mt-0.5">
                    Eliminates empty ghost rows while keeping valid row alignment intact.
                  </span>
                  {analysis && analysis.blankRowIndices.length > 0 && (
                    <span className="inline-block mt-1 text-[10px] font-semibold text-[#B45309] bg-[#FEF9EE] px-1.5 py-0.5 rounded">
                      {analysis.blankRowIndices.length} blank row(s) found
                    </span>
                  )}
                </div>
              </div>
            </div>

            {/* Op 6: Clean GSTINs */}
            <div
              onClick={() => handleOptionToggle('cleanGstin')}
              className={`p-3 rounded-lg border cursor-pointer transition-colors ${
                options.cleanGstin ? 'bg-slate-50 border-[#CBD5E1]' : 'bg-[#FAFAFA] border-gray-200 opacity-75'
              }`}
            >
              <div className="flex items-start gap-2.5">
                <input
                  type="checkbox"
                  checked={options.cleanGstin}
                  onChange={() => {}}
                  className="mt-0.5 rounded text-gray-900 focus:ring-gray-900"
                />
                <div>
                  <span className="text-xs font-semibold text-gray-900 block">
                    Validate & Standardize GSTINs
                  </span>
                  <span className="text-[11px] text-gray-600 block mt-0.5">
                    Uppercases and verifies 15-character GST identification numbers.
                  </span>
                </div>
              </div>
            </div>
          </div>

          {/* Action Execution Button */}
          <div className="pt-3 border-t border-gray-200 flex items-center justify-between">
            <div className="flex items-center gap-2 text-xs text-gray-600">
              <ShieldCheck className="w-4 h-4 text-gray-900" />
              <span>Formulas (=SUM) and unselected columns are permanently protected.</span>
            </div>

            <button
              onClick={handleClean}
              className="px-5 py-2 text-xs font-semibold text-white bg-gray-900 hover:bg-black rounded-lg shadow-sm flex items-center gap-2 transition-colors"
            >
              <Sparkles className="w-4 h-4" />
              <span>Apply Selected Changes</span>
            </button>
          </div>
        </div>
      )}

      {/* 6. Post-Processing Summary Banner */}
      {isProcessed && executionReport && (
        <div className="bg-gray-100 border border-gray-200 rounded-xl p-4 flex flex-col md:flex-row md:items-center justify-between gap-3 text-xs text-gray-900">
          <div className="flex items-center gap-3">
            <CheckCircle2 className="w-6 h-6 text-gray-900 shrink-0" />
            <div>
              <span className="font-bold text-sm text-gray-900 block">
                Cleaning Operations Applied Successfully
              </span>
              <span className="text-[11px] text-[#E5E7EB] block mt-0.5">
                Processed {executionReport.sheetsProcessed.length} sheet(s): Removed{' '}
                <strong>{executionReport.totalDuplicatesRemoved} duplicates</strong>, trimmed{' '}
                <strong>{executionReport.totalCellsTrimmed} spaces</strong>, formatted{' '}
                <strong>{executionReport.totalDatesStandardized} dates</strong>, standardized{' '}
                <strong>{executionReport.totalNumbersStandardized} currency amounts</strong>. Preserved{' '}
                <strong>{executionReport.formulasPreserved} formulas</strong> and{' '}
                <strong>{executionReport.mergedRangesPreserved} merged ranges</strong>.
              </span>
            </div>
          </div>

          <button
            onClick={handleDownload}
            className="inline-flex items-center gap-1.5 px-4 py-2 bg-gray-900 hover:bg-black text-white rounded-lg text-xs font-semibold transition-colors shrink-0 shadow-sm"
          >
            <Download className="w-4 h-4" />
            <span>Download Cleaned Excel</span>
          </button>
        </div>
      )}

      {/* 7. Sheet Preview Table */}
      {activeWb && activeWb.SheetNames.length > 0 && (
        <div className="bg-white rounded-xl border border-gray-200 overflow-hidden shadow-sm">
          {/* Preview Toolbar */}
          <div className="px-5 py-3 bg-gray-50 border-b border-gray-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs">
            <div className="flex items-center gap-3">
              <span className="font-bold text-gray-900">
                {isProcessed ? 'Cleaned Spreadsheet Preview' : 'Source Spreadsheet Preview'}
              </span>

              {/* Sheet tabs for preview */}
              <div className="flex items-center gap-1 bg-gray-100 p-0.5 rounded-lg border border-gray-200">
                {activeWb.SheetNames.map((s) => (
                  <button
                    key={s}
                    onClick={() => setPreviewSheet(s)}
                    className={`px-2.5 py-1 rounded-md text-[11px] font-medium transition-colors ${
                      previewSheet === s
                        ? 'bg-white text-gray-900 shadow-sm font-semibold'
                        : 'text-gray-600 hover:text-gray-900'
                    }`}
                  >
                    📄 {s}
                  </button>
                ))}
              </div>
            </div>

            {/* Filter buttons if processed */}
            {isProcessed && executionReport && (
              <div className="flex items-center gap-1.5">
                <button
                  onClick={() => setActiveFilter('all')}
                  className={`px-2.5 py-1 rounded-md text-[11px] font-medium transition-colors ${
                    activeFilter === 'all'
                      ? 'bg-white text-gray-900 border border-gray-200 shadow-sm font-semibold'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  All Rows ({previewDataRows.length})
                </button>
                <button
                  onClick={() => setActiveFilter('modified')}
                  className={`px-2.5 py-1 rounded-md text-[11px] font-medium transition-colors ${
                    activeFilter === 'modified'
                      ? 'bg-white text-gray-900 border border-gray-200 shadow-sm font-semibold'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Modified Only ({executionReport.changes.length})
                </button>
              </div>
            )}
          </div>

          {/* Table */}
          <div className="overflow-x-auto max-h-[460px]">
            {currentPreviewData.length === 0 ? (
              <div className="p-8 text-center text-xs text-gray-600">
                No data records found in worksheet "{previewSheet}".
              </div>
            ) : (
              <table className="w-full text-left text-xs border-collapse">
                <thead className="sticky top-0 bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold z-10">
                  <tr>
                    <th className="py-2.5 px-3 border-r border-gray-200 w-12 text-center">Row</th>
                    {previewHeaders.map((header, cIdx) => (
                      <th key={cIdx} className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">
                        {String(header || `Col ${cIdx + 1}`)}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#E5E7EB] text-gray-900">
                  {displayRows.map((row, rIdx) => (
                    <tr key={rIdx} className="hover:bg-gray-50 transition-colors">
                      <td className="py-2 px-3 border-r border-gray-200 text-center text-gray-600 text-[11px] font-mono">
                        {rIdx + 2}
                      </td>
                      {previewHeaders.map((_, cIdx) => {
                        const val = row[cIdx];
                        const isNum = typeof val === 'number';
                        return (
                          <td
                            key={cIdx}
                            className={`py-2 px-3 border-r border-gray-200 whitespace-nowrap ${
                              isNum ? 'text-right font-mono' : ''
                            }`}
                          >
                            {isNum
                              ? val.toLocaleString('en-IN', {
                                  minimumFractionDigits: 2,
                                  maximumFractionDigits: 2,
                                })
                              : String(val ?? '')}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </div>
      )}
    </div>
  );
};