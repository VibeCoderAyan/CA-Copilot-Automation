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
  CheckSquare,
  Square,
  Printer,
  Sparkles,
  HelpCircle,
  RotateCcw,
} from 'lucide-react';
import { InnerToolBadge } from '../common/InnerToolBadge';
import {
  type ProposedCellChange,
  type CleanupOptions,
  type SheetValidationSummary,
  DEFAULT_CLEANUP_OPTIONS,
  analyzeAndValidateSheetData,
  applyApprovedChanges,
  exportCleanedWorkbook,
} from '../../utils/excelDataValidationEngine';

interface ExcelDataValidationToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

export const ExcelDataValidationTool: React.FC<ExcelDataValidationToolProps> = ({
  onBack,
  onShowToast,
}) => {
  const [rawRows, setRawRows] = useState<Record<string, any>[]>([]);
  const [fileName, setFileName] = useState('');
  const [sheetName, setSheetName] = useState('Sheet1');
  const [availableSheets, setAvailableSheets] = useState<string[]>([]);
  const [workbookObj, setWorkbookObj] = useState<XLSX.WorkBook | null>(null);

  const [options, setOptions] = useState<CleanupOptions>(DEFAULT_CLEANUP_OPTIONS);
  const [activeTab, setActiveTab] = useState<'issues' | 'preview_cleaned' | 'raw_data' | 'audit_report'>('issues');
  const [searchQuery, setSearchQuery] = useState('');
  const [filterIssueType, setFilterIssueType] = useState('all');
  const [showConfigPanel, setShowConfigPanel] = useState(false);

  // In-memory approval state
  const [approvalMap, setApprovalMap] = useState<Record<string, boolean>>({});

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Analyze sheet data
  const summary: SheetValidationSummary = useMemo(() => {
    if (rawRows.length === 0) {
      return {
        sheetName,
        totalRows: 0,
        totalCols: 0,
        headers: [],
        totalIssuesDetected: 0,
        duplicateRowsCount: 0,
        emptyRowsCount: 0,
        whitespaceCount: 0,
        dateIssuesCount: 0,
        numberIssuesCount: 0,
        casingIssuesCount: 0,
        blankCellsCount: 0,
        changes: [],
        cleanedRows: [],
      };
    }

    const res = analyzeAndValidateSheetData(sheetName, rawRows, options);

    // Apply manual approval overrides
    res.changes.forEach((c) => {
      if (approvalMap[c.id] !== undefined) {
        c.approved = approvalMap[c.id];
      }
    });

    // Recompute cleaned rows with updated approvals
    res.cleanedRows = applyApprovedChanges(rawRows, res.changes, res.headers);
    return res;
  }, [rawRows, sheetName, options, approvalMap]);

  // Handle file upload
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });
        setWorkbookObj(wb);
        setAvailableSheets(wb.SheetNames);

        const firstSheet = wb.SheetNames[0];
        setSheetName(firstSheet);

        const ws = wb.Sheets[firstSheet];
        const json = XLSX.utils.sheet_to_json<Record<string, any>>(ws, { defval: '' });

        if (json.length === 0) {
          onShowToast('Empty Sheet', 'Selected sheet has no readable data.', 'error');
          return;
        }

        setRawRows(json);
        setApprovalMap({});
      } catch (err: any) {
        onShowToast('Upload Error', err.message || 'Failed to parse file', 'error');
      }
    };

    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  // Switch sheet
  const handleSheetChange = (newSheet: string) => {
    if (!workbookObj || !workbookObj.Sheets[newSheet]) return;
    setSheetName(newSheet);
    const ws = workbookObj.Sheets[newSheet];
    const json = XLSX.utils.sheet_to_json<Record<string, any>>(ws, { defval: '' });
    setRawRows(json);
    setApprovalMap({});
    onShowToast('Sheet Switched', `Active sheet changed to ${newSheet}`, 'info');
  };

  // Toggle individual change approval
  const handleToggleApproval = (id: string, current: boolean) => {
    setApprovalMap((prev) => ({
      ...prev,
      [id]: !current,
    }));
  };

  // Bulk actions
  const handleApproveAll = () => {
    const newMap: Record<string, boolean> = {};
    summary.changes.forEach((c) => {
      newMap[c.id] = true;
    });
    setApprovalMap(newMap);
    onShowToast('Approved All', `Approved all ${summary.changes.length} proposed cleanup changes`, 'success');
  };

  const handleRejectAll = () => {
    const newMap: Record<string, boolean> = {};
    summary.changes.forEach((c) => {
      newMap[c.id] = false;
    });
    setApprovalMap(newMap);
    onShowToast('Rejected All', 'Reset all proposed changes. Original data preserved.', 'info');
  };

  // Export Cleaned Excel Workbook
  const handleExportExcel = () => {
    if (summary.cleanedRows.length === 0) {
      onShowToast('No Data', 'Please load data before exporting.', 'error');
      return;
    }

    try {
      const wb = exportCleanedWorkbook(summary.cleanedRows, summary.changes, sheetName);
      XLSX.writeFile(wb, `Cleaned_${fileName.replace(/\.[^/.]+$/, '') || 'Data'}.xlsx`);
      onShowToast('Workbook Exported', 'Cleaned dataset and audit report exported successfully.', 'success');
    } catch (err: any) {
      onShowToast('Export Error', err.message || 'Failed to generate Excel file.', 'error');
    }
  };

  // Filtered changes
  const filteredChanges = useMemo(() => {
    return summary.changes.filter((c) => {
      if (filterIssueType !== 'all' && c.issueType !== filterIssueType) return false;
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase();
        const col = c.colHeader.toLowerCase();
        const desc = c.description.toLowerCase();
        const orig = String(c.originalValue).toLowerCase();
        const prop = String(c.proposedValue).toLowerCase();
        return col.includes(q) || desc.includes(q) || orig.includes(q) || prop.includes(q);
      }
      return true;
    });
  }, [summary.changes, filterIssueType, searchQuery]);

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="excel-data-cleanup"
        title="Excel Data Validation & Cleanup"
        description="Detect whitespace, irregular dates, duplicate rows, empty cells, and casing inconsistencies with itemized change approvals."
        badge="Popular"
        onBack={onBack}
        actionButtons={
          <div className="flex flex-wrap items-center gap-2">
            

            <button
              onClick={() => fileInputRef.current?.click()}
              className="px-3 py-1.5 text-xs font-semibold rounded-md bg-white text-gray-900 border border-gray-200 hover:bg-gray-50 transition-colors flex items-center gap-1.5 shadow-sm"
            >
              <UploadCloud className="w-3.5 h-3.5 text-gray-900" />
              {fileName ? 'Re-upload File' : 'Import Excel / CSV'}
            </button>
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleFileUpload}
              accept=".xlsx,.xls,.csv"
              className="hidden"
              id="excel-val-upload"
            />

            <button
              onClick={() => setShowConfigPanel(!showConfigPanel)}
              className={`px-3 py-1.5 text-xs font-semibold rounded-md border transition-colors flex items-center gap-1.5 ${
                showConfigPanel
                  ? 'bg-gray-900 text-white border-gray-300'
                  : 'bg-white text-gray-900 border-gray-200 hover:bg-gray-50'
              }`}
            >
              <Sliders className="w-3.5 h-3.5" />
              Cleanup Rules
            </button>

            <button
              onClick={handleExportExcel}
              disabled={summary.cleanedRows.length === 0}
              className="px-3.5 py-1.5 text-xs font-semibold rounded-md bg-gray-900 text-white hover:bg-black disabled:opacity-50 transition-colors flex items-center gap-1.5 shadow-sm cursor-pointer"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              Export Cleaned Excel (.xlsx)
            </button>
          </div>
        }
      />

      {/* Workbook Metadata Bar */}
      <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-2xs flex flex-wrap items-center gap-4 text-xs">
          <div className="flex items-center gap-2">
            <span className="font-semibold text-gray-600">File:</span>
            <span className="font-medium text-gray-900">{fileName || 'No file selected'}</span>
          </div>

          {availableSheets.length > 1 && (
            <div className="flex items-center gap-2">
              <span className="font-semibold text-gray-600">Select Sheet:</span>
              <select
                value={sheetName}
                onChange={(e) => handleSheetChange(e.target.value)}
                className="px-2 py-1 border border-gray-200 rounded text-xs font-medium focus:border-gray-300 outline-none"
              >
                {availableSheets.map((s) => (
                  <option key={s} value={s}>{s}</option>
                ))}
              </select>
            </div>
          )}

          {rawRows.length > 0 && (
            <div className="ml-auto flex items-center gap-4 text-gray-600">
              <span>Rows: <strong className="text-gray-900">{summary.totalRows}</strong></span>
              <span>Columns: <strong className="text-gray-900">{summary.totalCols}</strong></span>
              <span>Issues Detected: <strong className="text-[#DC2626]">{summary.totalIssuesDetected}</strong></span>
            </div>
          )}
      </div>

      {/* Upload Dropzone when no file selected */}
      {!fileName && (
        <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-8 sm:p-10 text-center transition-colors shadow-2xs">
          <label htmlFor="excel-val-upload" className="cursor-pointer flex flex-col items-center">
            <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mx-auto mb-3 shadow-2xs">
              <UploadCloud className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-semibold text-gray-900">Upload Excel Spreadsheet (.xlsx, .xls, .csv)</h3>
            <p className="text-xs text-gray-500 mt-1 max-w-md mx-auto leading-relaxed">
              Upload client spreadsheets to detect duplicate records, trailing whitespace, invalid GSTINs, broken dates, and text numbers without altering original formulas.
            </p>
            <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
              <UploadCloud className="w-4 h-4" />
              <span>Import Excel / CSV</span>
            </div>
            <span className="text-[11px] text-gray-400 font-medium mt-3 block">
              Supports .xlsx, .xls, .csv · Non-destructive audit rules · Zero silent mutations
            </span>
          </label>
        </div>
      )}

      {/* Main Container */}
      <div className="space-y-6">
        {/* Safe Non-Destructive Warning */}
        <div className="bg-gray-50 border border-gray-300/30 rounded-xl p-3 text-xs text-gray-900 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <ShieldCheck className="w-4 h-4 shrink-0 text-gray-900" />
            <span>
              <strong>Non-Destructive Guarantee:</strong> Your original uploaded file is never modified or overwritten. Proposed changes are reviewed individually and applied only upon your explicit approval.
            </span>
          </div>
        </div>

        {/* Cleanup Options Drawer */}
        {showConfigPanel && (
          <div className="bg-white border border-gray-200 rounded-xl p-5 shadow-sm space-y-4">
            <div className="flex items-center justify-between pb-3 border-b border-gray-200">
              <div>
                <h3 className="text-sm font-bold text-gray-900">Cleanup Rules & Automation Settings</h3>
                <p className="text-xs text-gray-600">Choose which data cleansing heuristics to run on imported spreadsheets</p>
              </div>
              <button onClick={() => setShowConfigPanel(false)} className="text-xs text-gray-600 hover:text-gray-900">Close</button>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs">
              <label className="flex items-start gap-2.5 p-2 rounded border border-gray-100 hover:bg-gray-50 cursor-pointer">
                <input
                  type="checkbox"
                  checked={options.trimWhitespace}
                  onChange={(e) => setOptions({ ...options, trimWhitespace: e.target.checked })}
                  className="rounded text-gray-900 focus:ring-gray-900 mt-0.5"
                />
                <div>
                  <strong className="text-gray-900">Trim Leading & Trailing Spaces</strong>
                  <p className="text-[11px] text-gray-600">Removes invisible whitespace at start or end of text</p>
                </div>
              </label>

              <label className="flex items-start gap-2.5 p-2 rounded border border-gray-100 hover:bg-gray-50 cursor-pointer">
                <input
                  type="checkbox"
                  checked={options.normalizeWhitespace}
                  onChange={(e) => setOptions({ ...options, normalizeWhitespace: e.target.checked })}
                  className="rounded text-gray-900 focus:ring-gray-900 mt-0.5"
                />
                <div>
                  <strong className="text-gray-900">Collapse Multiple Spaces</strong>
                  <p className="text-[11px] text-gray-600">Converts consecutive double/triple spaces to a single space</p>
                </div>
              </label>

              <label className="flex items-start gap-2.5 p-2 rounded border border-gray-100 hover:bg-gray-50 cursor-pointer">
                <input
                  type="checkbox"
                  checked={options.standardizeDates}
                  onChange={(e) => setOptions({ ...options, standardizeDates: e.target.checked })}
                  className="rounded text-gray-900 focus:ring-gray-900 mt-0.5"
                />
                <div>
                  <strong className="text-gray-900">Standardize Date Formats</strong>
                  <p className="text-[11px] text-gray-600">Format all dates consistently into DD-MM-YYYY</p>
                </div>
              </label>

              <label className="flex items-start gap-2.5 p-2 rounded border border-gray-100 hover:bg-gray-50 cursor-pointer">
                <input
                  type="checkbox"
                  checked={options.standardizeNumbers}
                  onChange={(e) => setOptions({ ...options, standardizeNumbers: e.target.checked })}
                  className="rounded text-gray-900 focus:ring-gray-900 mt-0.5"
                />
                <div>
                  <strong className="text-gray-900">Clean Currency & Numbers Stored as Text</strong>
                  <p className="text-[11px] text-gray-600">Removes ₹ and commas so numbers calculate properly</p>
                </div>
              </label>

              <label className="flex items-start gap-2.5 p-2 rounded border border-gray-100 hover:bg-gray-50 cursor-pointer">
                <input
                  type="checkbox"
                  checked={options.normalizeCasing}
                  onChange={(e) => setOptions({ ...options, normalizeCasing: e.target.checked })}
                  className="rounded text-gray-900 focus:ring-gray-900 mt-0.5"
                />
                <div>
                  <strong className="text-gray-900">Standardize PAN & GSTIN Uppercase</strong>
                  <p className="text-[11px] text-gray-600">Ensures tax codes and references are uniformly capitalized</p>
                </div>
              </label>

              <label className="flex items-start gap-2.5 p-2 rounded border border-gray-100 hover:bg-gray-50 cursor-pointer">
                <input
                  type="checkbox"
                  checked={options.removeDuplicates}
                  onChange={(e) => setOptions({ ...options, removeDuplicates: e.target.checked })}
                  className="rounded text-gray-900 focus:ring-gray-900 mt-0.5"
                />
                <div>
                  <strong className="text-gray-900">Detect & Omit Duplicate Rows</strong>
                  <p className="text-[11px] text-gray-600">Identifies identical rows across all columns</p>
                </div>
              </label>
            </div>
          </div>
        )}

        {/* Empty State */}
        {rawRows.length === 0 ? (
          <div className="bg-white border border-dashed border-gray-200 rounded-2xl p-12 text-center shadow-sm">
            <div className="w-16 h-16 bg-gray-50 rounded-full flex items-center justify-center mx-auto mb-4 border border-gray-300/20">
              <Sparkles className="w-8 h-8 text-gray-900" />
            </div>
            <h3 className="text-base font-bold text-gray-900">No Spreadsheet Data Loaded</h3>
            <p className="text-xs text-gray-600 max-w-md mx-auto mt-1 mb-6">
              Upload any client accounting file (Excel or CSV) to automatically diagnose whitespace, date errors, casing mismatches, and duplicate records.
            </p>
            <div className="flex items-center justify-center gap-3">
              
              <button
                onClick={() => fileInputRef.current?.click()}
                className="px-4 py-2 text-xs font-semibold rounded-lg bg-gray-900 text-white hover:bg-black transition-colors shadow-sm inline-flex items-center gap-2"
              >
                <UploadCloud className="w-4 h-4" />
                <span>Upload Excel / CSV File</span>
              </button>
            </div>
          </div>
        ) : (
          <div className="space-y-6">
            {/* KPI Summary Cards */}
            <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
              <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
                <span className="text-xs font-semibold text-gray-600 uppercase">Total Issues Flagged</span>
                <div className="text-xl font-bold text-[#DC2626] mt-1">
                  {summary.totalIssuesDetected}
                </div>
                <div className="text-xs text-gray-600 mt-0.5">Across {summary.totalRows} data rows</div>
              </div>

              <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
                <span className="text-xs font-semibold text-gray-600 uppercase">Whitespace Irregularities</span>
                <div className="text-xl font-bold text-[#D97706] mt-1">
                  {summary.whitespaceCount}
                </div>
                <div className="text-xs text-gray-600 mt-0.5">Leading, trailing & double spaces</div>
              </div>

              <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
                <span className="text-xs font-semibold text-gray-600 uppercase">Date & Number Fixes</span>
                <div className="text-xl font-bold text-gray-900 mt-1">
                  {summary.dateIssuesCount + summary.numberIssuesCount}
                </div>
                <div className="text-xs text-gray-600 mt-0.5">Formats & text-to-number</div>
              </div>

              <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
                <span className="text-xs font-semibold text-gray-600 uppercase">Casing Corrections</span>
                <div className="text-xl font-bold text-[#7C3AED] mt-1">
                  {summary.casingIssuesCount}
                </div>
                <div className="text-xs text-gray-600 mt-0.5">PAN / GSTIN / Title case</div>
              </div>

              <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
                <span className="text-xs font-semibold text-gray-600 uppercase">Duplicates & Blanks</span>
                <div className="text-xl font-bold text-[#DC2626] mt-1">
                  {summary.duplicateRowsCount + summary.emptyRowsCount}
                </div>
                <div className="text-xs text-gray-600 mt-0.5">Redundant rows detected</div>
              </div>
            </div>

            {/* Tabs Bar */}
            <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 border-b border-gray-200 pb-1 print:hidden">
              <div className="flex items-center gap-2">
                <button
                  onClick={() => setActiveTab('issues')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'issues'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Quality Issues & Change Approvals ({summary.changes.length})
                </button>
                <button
                  onClick={() => setActiveTab('preview_cleaned')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'preview_cleaned'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Cleaned Data Preview ({summary.cleanedRows.length} rows)
                </button>
                <button
                  onClick={() => setActiveTab('raw_data')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'raw_data'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Original Data Raw ({rawRows.length} rows)
                </button>
                <button
                  onClick={() => setActiveTab('audit_report')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'audit_report'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Data Quality Audit Log
                </button>
              </div>

              {activeTab === 'issues' && (
                <div className="flex items-center gap-2">
                  <button
                    onClick={handleApproveAll}
                    className="px-2.5 py-1 text-xs font-semibold rounded bg-gray-50 text-gray-900 border border-gray-200 hover:bg-gray-100 transition-colors"
                  >
                    ✓ Approve All
                  </button>
                  <button
                    onClick={handleRejectAll}
                    className="px-2.5 py-1 text-xs font-semibold rounded bg-gray-50 text-gray-600 border border-gray-300 hover:bg-gray-100 transition-colors"
                  >
                    ✗ Reset All
                  </button>

                  <select
                    value={filterIssueType}
                    onChange={(e) => setFilterIssueType(e.target.value)}
                    className="text-xs border border-gray-200 rounded px-2 py-1 outline-none"
                  >
                    <option value="all">All Issue Types</option>
                    <option value="leading_trailing_whitespace">Whitespace</option>
                    <option value="inconsistent_date_format">Dates</option>
                    <option value="invalid_numeric_format">Numbers</option>
                    <option value="inconsistent_casing">Casing</option>
                    <option value="duplicate_row">Duplicates</option>
                    <option value="empty_row">Empty Rows</option>
                    <option value="blank_cell">Blank Cells</option>
                  </select>
                </div>
              )}
            </div>

            {/* TAB 1: Issues & Itemized Approvals */}
            {activeTab === 'issues' && (
              <div className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm">
                <div className="overflow-x-auto">
                  <table className="w-full text-xs text-left">
                    <thead className="bg-gray-900 text-white font-semibold">
                      <tr>
                        <th className="p-3 w-10 text-center">Approve</th>
                        <th className="p-3 w-16 text-center">Row #</th>
                        <th className="p-3">Column Header</th>
                        <th className="p-3">Issue Category</th>
                        <th className="p-3">Original Source Value</th>
                        <th className="p-3">Proposed Cleaned Value</th>
                        <th className="p-3">Quality Observation</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB]">
                      {filteredChanges.map((c) => (
                        <tr
                          key={c.id}
                          className={`hover:bg-gray-50 transition-colors ${
                            !c.approved ? 'opacity-60 bg-gray-50/50' : ''
                          }`}
                        >
                          <td className="p-3 text-center">
                            <button
                              onClick={() => handleToggleApproval(c.id, c.approved)}
                              className="cursor-pointer"
                              title={c.approved ? 'Click to Reject' : 'Click to Approve'}
                            >
                              {c.approved ? (
                                <CheckSquare className="w-4 h-4 text-gray-900" />
                              ) : (
                                <Square className="w-4 h-4 text-gray-400" />
                              )}
                            </button>
                          </td>
                          <td className="p-3 text-center font-bold text-gray-600">{c.rowIndex}</td>
                          <td className="p-3 font-semibold text-gray-900">{c.colHeader}</td>
                          <td className="p-3">
                            <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-[#FEF3C7] text-[#D97706]">
                              {c.issueType.replace(/_/g, ' ').toUpperCase()}
                            </span>
                          </td>
                          <td className="p-3 font-mono text-[#DC2626] bg-red-50/30 max-w-xs truncate">
                            {String(c.originalValue)}
                          </td>
                          <td className="p-3 font-mono text-gray-900 bg-gray-50/30 max-w-xs truncate font-semibold">
                            {String(c.proposedValue)}
                          </td>
                          <td className="p-3 text-[11px] text-gray-600">
                            {c.description}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* TAB 2: Cleaned Data Preview */}
            {activeTab === 'preview_cleaned' && (
              <div className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm">
                <div className="p-3 bg-gray-50 border-b border-gray-200 flex justify-between items-center text-xs">
                  <span className="font-semibold text-gray-900">
                    Preview of dataset after applying approved cleanses ({summary.cleanedRows.length} rows)
                  </span>
                  <button
                    onClick={handleExportExcel}
                    className="px-3 py-1 bg-gray-900 text-white rounded text-xs font-semibold hover:bg-black"
                  >
                    Export Cleaned Workbook
                  </button>
                </div>
                <div className="overflow-x-auto max-h-[600px]">
                  <table className="w-full text-xs text-left">
                    <thead className="bg-gray-900 text-white sticky top-0 font-semibold">
                      <tr>
                        <th className="p-3 w-10 text-center">#</th>
                        {summary.headers.map((h) => (
                          <th key={h} className="p-3 whitespace-nowrap">{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB]">
                      {summary.cleanedRows.map((row, idx) => (
                        <tr key={idx} className="hover:bg-gray-50 font-mono">
                          <td className="p-3 text-center text-gray-600">{idx + 1}</td>
                          {summary.headers.map((h) => (
                            <td key={h} className="p-3 whitespace-nowrap text-gray-900">
                              {String(row[h] ?? '')}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* TAB 3: Raw Data */}
            {activeTab === 'raw_data' && (
              <div className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm">
                <div className="p-3 bg-gray-50 border-b border-gray-200 text-xs text-gray-600">
                  Original uploaded data without any modifications ({rawRows.length} rows)
                </div>
                <div className="overflow-x-auto max-h-[600px]">
                  <table className="w-full text-xs text-left">
                    <thead className="bg-gray-200 text-gray-900 sticky top-0 font-semibold">
                      <tr>
                        <th className="p-3 w-10 text-center">#</th>
                        {summary.headers.map((h) => (
                          <th key={h} className="p-3 whitespace-nowrap">{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB]">
                      {rawRows.map((row, idx) => (
                        <tr key={idx} className="hover:bg-gray-50 font-mono">
                          <td className="p-3 text-center text-gray-600">{idx + 1}</td>
                          {summary.headers.map((h) => (
                            <td key={h} className="p-3 whitespace-nowrap text-gray-900">
                              {String(row[h] ?? '')}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* TAB 4: Audit Report */}
            {activeTab === 'audit_report' && (
              <div className="bg-white border border-gray-200 rounded-xl p-8 shadow-sm space-y-6">
                <div className="border-b-2 border-gray-300 pb-4">
                  <h2 className="text-xl font-bold text-gray-900">DATA QUALITY AUDIT & RECTIFICATION CERTIFICATE</h2>
                  <p className="text-xs text-gray-600 mt-0.5">Automated Cleansing Log • Worksheet: {sheetName} • Source: {fileName}</p>
                </div>

                <div className="grid grid-cols-2 md:grid-cols-4 gap-4 text-xs">
                  <div className="border border-gray-200 rounded p-3">
                    <span className="text-gray-600">Initial Scanned Rows:</span>
                    <span className="font-bold ml-2">{summary.totalRows}</span>
                  </div>
                  <div className="border border-gray-200 rounded p-3">
                    <span className="text-gray-600">Cleaned Valid Rows:</span>
                    <span className="font-bold ml-2 text-gray-900">{summary.cleanedRows.length}</span>
                  </div>
                  <div className="border border-gray-200 rounded p-3">
                    <span className="text-gray-600">Total Issues Addressed:</span>
                    <span className="font-bold ml-2">{summary.changes.filter((c) => c.approved).length}</span>
                  </div>
                  <div className="border border-gray-200 rounded p-3">
                    <span className="text-gray-600">Redundant Rows Dropped:</span>
                    <span className="font-bold ml-2 text-[#DC2626]">{summary.totalRows - summary.cleanedRows.length}</span>
                  </div>
                </div>

                <table className="w-full text-xs text-left border-collapse">
                  <thead>
                    <tr className="bg-gray-900 text-white">
                      <th className="p-2 border border-gray-300">#</th>
                      <th className="p-2 border border-gray-300">Row</th>
                      <th className="p-2 border border-gray-300">Column</th>
                      <th className="p-2 border border-gray-300">Issue Type</th>
                      <th className="p-2 border border-gray-300">Original Value</th>
                      <th className="p-2 border border-gray-300">Cleaned Value</th>
                      <th className="p-2 border border-gray-300">Action Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {summary.changes.map((c, idx) => (
                      <tr key={c.id}>
                        <td className="p-2 border border-gray-300 text-center">{idx + 1}</td>
                        <td className="p-2 border border-gray-300 text-center font-bold">{c.rowIndex}</td>
                        <td className="p-2 border border-gray-300 font-semibold">{c.colHeader}</td>
                        <td className="p-2 border border-gray-300">{c.issueType.replace(/_/g, ' ')}</td>
                        <td className="p-2 border border-gray-300 font-mono text-red-700">{String(c.originalValue)}</td>
                        <td className="p-2 border border-gray-300 font-mono text-gray-800">{String(c.proposedValue)}</td>
                        <td className="p-2 border border-gray-300 font-semibold text-[10px]">
                          {c.approved ? 'APPROVED & APPLIED' : 'REJECTED'}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
};