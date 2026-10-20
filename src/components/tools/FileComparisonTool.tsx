import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  UploadCloud,
  Download,
  CheckCircle2,
  AlertTriangle,
  GitCompare,
  FileSpreadsheet,
  Layers,
  ArrowRight
} from 'lucide-react';
import {
  compareDatasets,
  createComparisonWorkbook,
  type FileComparisonResult,
} from '../../utils/fileComparisonEngine';
import { DataFormatGuide } from '../common/DataFormatGuide';
import { ToolHeader } from '../common/ToolHeader';

interface FileComparisonToolProps {
  onBack: () => void;
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
}

export const FileComparisonTool: React.FC<FileComparisonToolProps> = ({ onBack, onShowToast }) => {
  const [fileA, setFileA] = useState<{ name: string; size: string } | null>(null);
  const [fileB, setFileB] = useState<{ name: string; size: string } | null>(null);

  const [dataA, setDataA] = useState<Record<string, any>[]>([]);
  const [dataB, setDataB] = useState<Record<string, any>[]>([]);

  const [selectedKeyCol, setSelectedKeyCol] = useState<string>('');
  const [comparisonResult, setComparisonResult] = useState<FileComparisonResult | null>(null);
  const [isCompared, setIsCompared] = useState(false);
  const [activeTab, setActiveTab] = useState<'summary' | 'modified' | 'added' | 'removed' | 'matched'>('summary');
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
        if (!json || json.length === 0) throw new Error('File A contains no data rows.');

        setFileA({ name: f.name, size: `${Math.round(f.size / 1024)} KB` });
        setDataA(json);
        setIsCompared(false);
        setComparisonResult(null);

        // Auto-select key candidate
        const cols = Object.keys(json[0]);
        const cand = cols.find((c) => /\b(voucher|invoice|id|code|bill|sno|ref)\b/i.test(c)) || cols[0];
        if (!selectedKeyCol) setSelectedKeyCol(cand);
      } catch (err: any) {
        setValidationError(`Failed to parse File A: ${err.message}`);
        onShowToast('File A Error', err.message, 'error');
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
        if (!json || json.length === 0) throw new Error('File B contains no data rows.');

        setFileB({ name: f.name, size: `${Math.round(f.size / 1024)} KB` });
        setDataB(json);
        setIsCompared(false);
        setComparisonResult(null);
      } catch (err: any) {
        setValidationError(`Failed to parse File B: ${err.message}`);
        onShowToast('File B Error', err.message, 'error');
      }
    };
    reader.readAsBinaryString(f);
  };


  const handleReset = () => {
    setFileA(null);
    setFileB(null);
    setDataA([]);
    setDataB([]);
    setComparisonResult(null);
    setIsCompared(false);
    setSelectedKeyCol('');
    setValidationError(null);
    if (fileInputARef.current) fileInputARef.current.value = '';
    if (fileInputBRef.current) fileInputBRef.current.value = '';
    onShowToast('Comparison Cleared', 'Reset all loaded datasets.', 'info');
  };

  const handleCompare = () => {
    if (dataA.length === 0 || dataB.length === 0) {
      return;
    }

    const keyList = selectedKeyCol ? [selectedKeyCol] : [];
    const res = compareDatasets(dataA, dataB, fileA?.name || 'File A', fileB?.name || 'File B', keyList);
    setComparisonResult(res);
    setIsCompared(true);
    setActiveTab('summary');

    onShowToast(
      'Comparison Complete',
      `Identified ${res.modifiedRecords.length} modified, ${res.addedRecords.length} added, ${res.removedRecords.length} removed records.`,
      'success'
    );
  };

  const handleDownload = () => {
    if (!comparisonResult) return;
    const wb = createComparisonWorkbook(comparisonResult);
    const safeA = (fileA?.name || 'File_A').replace(/\.[^/.]+$/, '');
    const safeB = (fileB?.name || 'File_B').replace(/\.[^/.]+$/, '');
    const outName = `Comparison_Report_${safeA}_vs_${safeB}.xlsx`;
    XLSX.writeFile(wb, outName);
    onShowToast('Downloaded', `Exported 5-sheet comparison workbook as ${outName}.`, 'success');
  };

  const commonCols = dataA.length > 0 && dataB.length > 0
    ? Object.keys(dataA[0]).filter((k) => Object.keys(dataB[0]).includes(k))
    : [];

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="file-compare"
        title="File Comparison & Diff Auditor"
        description="Compare two Excel or CSV registers. Identify added, deleted, and modified records with cell-level diffs."
        badge="Popular"
        onBack={onBack}
        actionButtons={
          (fileA || fileB || comparisonResult) ? (
            <div className="flex items-center gap-2">
            {(fileA || fileB) && (
              <button
                onClick={handleReset}
                className="px-3 py-1.5 text-xs font-medium text-gray-600 hover:text-gray-900 bg-white hover:bg-gray-50 border border-gray-200 rounded-lg transition-colors cursor-pointer"
              >
                Clear
              </button>
            )}
            {comparisonResult && (
              <button
                onClick={handleDownload}
                className="px-4 py-2 text-xs font-semibold text-white bg-gray-900 hover:bg-black rounded-lg shadow-sm flex items-center gap-2 transition-colors cursor-pointer"
              >
                <Download className="w-4 h-4" />
                <span>Export Comparison (.xlsx)</span>
              </button>
            )}
          </div>
          ) : undefined
        }
      />

      {/* 2. Format Requirements Guide */}
      <DataFormatGuide
        toolType="excel-cleaner"
        acceptedFormats=".xlsx, .xls, .csv"
        requiredFields={[
          { name: 'Key Column (Recommended)', description: 'Unique identifier to match rows across files', aliases: ['Voucher_No', 'Invoice_No', 'Account_Code', 'ID'] },
          { name: 'Common Attributes', description: 'Columns to inspect for field-level modifications', aliases: ['Date', 'Party_Name', 'Taxable', 'Amount', 'Status'] },
        ]}
        notes={[
          'Cell-Level Diffs: Detects exact values changed between File A (baseline) and File B (revised).',
          'Structural Integrity: Preserves and reports columns added or dropped in revised files.',
          'Excel Export: Generates a 5-sheet audit working paper with separated summary, diffs, added, removed, and matched rows.',
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

      {/* 4. Dual File Upload Grid */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
        {/* File A Box */}
        <input type="file" ref={fileInputARef} onChange={handleUploadA} accept=".xlsx,.xls,.csv" className="hidden" id="file-a-upload" />
        {!fileA ? (
          <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-6 sm:p-8 text-center transition-colors shadow-2xs">
            <label htmlFor="file-a-upload" className="cursor-pointer flex flex-col items-center">
              <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mb-3 shadow-2xs">
                <FileSpreadsheet className="w-6 h-6" />
              </div>
              <h3 className="text-sm font-semibold text-gray-900">File A (Baseline / Reference)</h3>
              <p className="text-xs text-gray-500 mt-1 max-w-xs mx-auto">
                Original books, previous month audit dump, or master register (.xlsx, .xls, .csv)
              </p>
              <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
                <UploadCloud className="w-4 h-4" />
                <span>Upload Baseline File A</span>
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
                  File A (Baseline) • {dataA.length} rows • {fileA.size}
                </span>
              </div>
            </div>
            <div className="mt-4 pt-3 border-t border-gray-100 flex items-center justify-between">
              <span className="text-[11px] text-gray-500 font-medium">Ready for diff check</span>
              <button
                onClick={() => fileInputARef.current?.click()}
                className="px-3 py-1.5 text-xs text-gray-900 bg-gray-50 hover:bg-[#EAEAEA] rounded-lg border border-gray-200 flex items-center gap-1.5 font-medium transition-colors cursor-pointer"
              >
                <UploadCloud className="w-3.5 h-3.5 text-gray-600" />
                <span>Change File A</span>
              </button>
            </div>
          </div>
        )}

        {/* File B Box */}
        <input type="file" ref={fileInputBRef} onChange={handleUploadB} accept=".xlsx,.xls,.csv" className="hidden" id="file-b-upload" />
        {!fileB ? (
          <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-6 sm:p-8 text-center transition-colors shadow-2xs">
            <label htmlFor="file-b-upload" className="cursor-pointer flex flex-col items-center">
              <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mb-3 shadow-2xs">
                <GitCompare className="w-6 h-6" />
              </div>
              <h3 className="text-sm font-semibold text-gray-900">File B (Comparison / Revised)</h3>
              <p className="text-xs text-gray-500 mt-1 max-w-xs mx-auto">
                Revised register, current month dump, or amended spreadsheet (.xlsx, .xls, .csv)
              </p>
              <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
                <UploadCloud className="w-4 h-4" />
                <span>Upload Comparison File B</span>
              </div>
            </label>
          </div>
        ) : (
          <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-sm flex flex-col justify-between">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-lg bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 shrink-0">
                <GitCompare className="w-5 h-5" />
              </div>
              <div className="min-w-0 flex-1">
                <span className="text-xs font-bold text-gray-900 block truncate">{fileB.name}</span>
                <span className="text-[11px] text-gray-600 block mt-0.5">
                  File B (Comparison) • {dataB.length} rows • {fileB.size}
                </span>
              </div>
            </div>
            <div className="mt-4 pt-3 border-t border-gray-100 flex items-center justify-between">
              <span className="text-[11px] text-gray-500 font-medium">Ready for diff check</span>
              <button
                onClick={() => fileInputBRef.current?.click()}
                className="px-3 py-1.5 text-xs text-gray-900 bg-gray-50 hover:bg-[#EAEAEA] rounded-lg border border-gray-200 flex items-center gap-1.5 font-medium transition-colors cursor-pointer"
              >
                <UploadCloud className="w-3.5 h-3.5 text-gray-600" />
                <span>Change File B</span>
              </button>
            </div>
          </div>
        )}
      </div>

      {/* 5. Configuration & Run Bar */}
      <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div className="flex items-center gap-3">
          <Layers className="w-4 h-4 text-gray-900" />
          <span className="text-xs font-semibold text-gray-900">Key Column to Match Records:</span>
          {commonCols.length > 0 ? (
            <select
              value={selectedKeyCol}
              onChange={(e) => setSelectedKeyCol(e.target.value)}
              className="px-3 py-1.5 bg-gray-50 border border-gray-200 rounded-lg text-xs font-medium text-gray-900 focus:outline-none focus:border-gray-300"
            >
              {commonCols.map((c) => (
                <option key={c} value={c}>
                  🔑 {c}
                </option>
              ))}
            </select>
          ) : (
            <span className="text-[11px] text-gray-600">Upload files with common headers to select key</span>
          )}
        </div>

        <div className="flex items-center gap-2">
          
          <button
            onClick={handleCompare}
            disabled={dataA.length === 0 || dataB.length === 0}
            className="px-5 py-1.5 text-xs font-semibold text-white bg-gray-900 hover:bg-black disabled:opacity-40 disabled:cursor-not-allowed rounded-lg shadow-sm flex items-center gap-2 transition-colors cursor-pointer"
          >
            <GitCompare className="w-4 h-4" />
            <span>Compare Files</span>
          </button>
        </div>
      </div>

      {/* 6. Comparison Results Workspace */}
      {isCompared && comparisonResult && (
        <div className="space-y-4">
          {/* Metrics Overview Cards */}
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <span className="text-[11px] font-semibold text-gray-600 block">Matched Records</span>
              <span className="text-xl font-bold text-gray-900 mt-1 block">
                {comparisonResult.matchedRecords.length}
              </span>
              <span className="text-[10px] text-gray-600 block mt-0.5">100% Identical</span>
            </div>

            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <span className="text-[11px] font-semibold text-gray-600 block">Modified Records</span>
              <span className="text-xl font-bold text-[#B45309] mt-1 block">
                {comparisonResult.modifiedRecords.length}
              </span>
              <span className="text-[10px] text-[#B45309] block mt-0.5">Field-level differences</span>
            </div>

            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <span className="text-[11px] font-semibold text-gray-600 block">Added in File B</span>
              <span className="text-xl font-bold text-gray-900 mt-1 block">
                {comparisonResult.addedRecords.length}
              </span>
              <span className="text-[10px] text-gray-900 block mt-0.5">New entries</span>
            </div>

            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <span className="text-[11px] font-semibold text-gray-600 block">Removed in File B</span>
              <span className="text-xl font-bold text-[#C43D43] mt-1 block">
                {comparisonResult.removedRecords.length}
              </span>
              <span className="text-[10px] text-[#C43D43] block mt-0.5">Deleted from baseline</span>
            </div>
          </div>

          {/* Detailed Tabbed View */}
          <div className="bg-white rounded-xl border border-gray-200 overflow-hidden shadow-sm">
            {/* Tab Navigation */}
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
                  Summary Breakdown
                </button>
                <button
                  onClick={() => setActiveTab('modified')}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors ${
                    activeTab === 'modified'
                      ? 'bg-white text-[#B45309] border border-gray-200 shadow-sm'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Modified Diffs ({comparisonResult.modifiedRecords.length})
                </button>
                <button
                  onClick={() => setActiveTab('added')}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors ${
                    activeTab === 'added'
                      ? 'bg-white text-gray-900 border border-gray-200 shadow-sm'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Added Records ({comparisonResult.addedRecords.length})
                </button>
                <button
                  onClick={() => setActiveTab('removed')}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors ${
                    activeTab === 'removed'
                      ? 'bg-white text-[#C43D43] border border-gray-200 shadow-sm'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Removed Records ({comparisonResult.removedRecords.length})
                </button>
                <button
                  onClick={() => setActiveTab('matched')}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors ${
                    activeTab === 'matched'
                      ? 'bg-white text-gray-900 border border-gray-200 shadow-sm'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Matched Records ({comparisonResult.matchedRecords.length})
                </button>
              </div>
            </div>

            {/* Tab 1: Summary */}
            {activeTab === 'summary' && (
              <div className="p-5 space-y-4 text-xs">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div className="border border-gray-200 rounded-lg p-4 space-y-2">
                    <span className="font-bold text-gray-900 block pb-1 border-b border-gray-200">
                      Baseline: {comparisonResult.fileAName}
                    </span>
                    <div className="flex justify-between py-1 text-gray-600">
                      <span>Total Rows:</span>
                      <strong className="text-gray-900">{comparisonResult.totalRowsA}</strong>
                    </div>
                    <div className="flex justify-between py-1 text-gray-600">
                      <span>Columns:</span>
                      <strong className="text-gray-900">{comparisonResult.columnsA.length}</strong>
                    </div>
                    <div className="flex justify-between py-1 text-gray-600">
                      <span>Dropped in File B:</span>
                      <strong className="text-[#C43D43]">{comparisonResult.removedColumns.join(', ') || 'None'}</strong>
                    </div>
                  </div>

                  <div className="border border-gray-200 rounded-lg p-4 space-y-2">
                    <span className="font-bold text-gray-900 block pb-1 border-b border-gray-200">
                      Comparison: {comparisonResult.fileBName}
                    </span>
                    <div className="flex justify-between py-1 text-gray-600">
                      <span>Total Rows:</span>
                      <strong className="text-gray-900">{comparisonResult.totalRowsB}</strong>
                    </div>
                    <div className="flex justify-between py-1 text-gray-600">
                      <span>Columns:</span>
                      <strong className="text-gray-900">{comparisonResult.columnsB.length}</strong>
                    </div>
                    <div className="flex justify-between py-1 text-gray-600">
                      <span>New Columns Added:</span>
                      <strong className="text-gray-900">{comparisonResult.addedColumns.join(', ') || 'None'}</strong>
                    </div>
                  </div>
                </div>
              </div>
            )}

            {/* Tab 2: Modified Diffs */}
            {activeTab === 'modified' && (
              <div className="overflow-x-auto max-h-[460px]">
                {comparisonResult.modifiedRecords.length === 0 ? (
                  <div className="p-8 text-center text-xs text-gray-600">Zero modified records detected.</div>
                ) : (
                  <table className="w-full text-left text-xs border-collapse">
                    <thead className="sticky top-0 bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold">
                      <tr>
                        <th className="py-2.5 px-3 border-r border-gray-200 w-12 text-center">#</th>
                        <th className="py-2.5 px-3 border-r border-gray-200">Matching Key</th>
                        <th className="py-2.5 px-3 border-r border-gray-200">Column Changed</th>
                        <th className="py-2.5 px-3 border-r border-gray-200 bg-[#FDF2F2] text-[#C43D43]">File A (Original)</th>
                        <th className="py-2.5 px-3 bg-gray-100 text-gray-900">File B (Modified)</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB] text-gray-900">
                      {comparisonResult.modifiedRecords.flatMap((m, mIdx) =>
                        m.differences.map((d, dIdx) => (
                          <tr key={`${mIdx}-${dIdx}`} className="hover:bg-gray-50">
                            <td className="py-2 px-3 border-r border-gray-200 text-center text-gray-600 text-[11px]">
                              {mIdx + 1}.{dIdx + 1}
                            </td>
                            <td className="py-2 px-3 border-r border-gray-200 font-semibold font-mono text-gray-900">
                              {m.key}
                            </td>
                            <td className="py-2 px-3 border-r border-gray-200 font-medium text-[#B45309]">
                              {d.column}
                            </td>
                            <td className="py-2 px-3 border-r border-gray-200 bg-[#FDF9F9] font-mono text-[#C43D43]">
                              {String(d.oldValue ?? '(Empty)')}
                            </td>
                            <td className="py-2 px-3 bg-slate-50 font-mono text-gray-900 font-semibold">
                              {String(d.newValue ?? '(Empty)')}
                            </td>
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                )}
              </div>
            )}

            {/* Tab 3: Added Records */}
            {activeTab === 'added' && (
              <div className="overflow-x-auto max-h-[460px]">
                {comparisonResult.addedRecords.length === 0 ? (
                  <div className="p-8 text-center text-xs text-gray-600">Zero added records in File B.</div>
                ) : (
                  <table className="w-full text-left text-xs border-collapse">
                    <thead className="sticky top-0 bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold">
                      <tr>
                        <th className="py-2.5 px-3 border-r border-gray-200 w-12 text-center">#</th>
                        {Object.keys(comparisonResult.addedRecords[0]).map((h) => (
                          <th key={h} className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB] text-gray-900">
                      {comparisonResult.addedRecords.map((row, idx) => (
                        <tr key={idx} className="bg-slate-50 hover:bg-gray-100">
                          <td className="py-2 px-3 border-r border-gray-200 text-center text-gray-900 text-[11px] font-bold">
                            +{idx + 1}
                          </td>
                          {Object.keys(row).map((k) => (
                            <td key={k} className="py-2 px-3 border-r border-gray-200 whitespace-nowrap font-mono">
                              {String(row[k] ?? '')}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
            )}

            {/* Tab 4: Removed Records */}
            {activeTab === 'removed' && (
              <div className="overflow-x-auto max-h-[460px]">
                {comparisonResult.removedRecords.length === 0 ? (
                  <div className="p-8 text-center text-xs text-gray-600">Zero removed records from File A.</div>
                ) : (
                  <table className="w-full text-left text-xs border-collapse">
                    <thead className="sticky top-0 bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold">
                      <tr>
                        <th className="py-2.5 px-3 border-r border-gray-200 w-12 text-center">#</th>
                        {Object.keys(comparisonResult.removedRecords[0]).map((h) => (
                          <th key={h} className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB] text-gray-900">
                      {comparisonResult.removedRecords.map((row, idx) => (
                        <tr key={idx} className="bg-[#FDF2F2] hover:bg-[#FBE8E8]">
                          <td className="py-2 px-3 border-r border-gray-200 text-center text-[#C43D43] text-[11px] font-bold">
                            -{idx + 1}
                          </td>
                          {Object.keys(row).map((k) => (
                            <td key={k} className="py-2 px-3 border-r border-gray-200 whitespace-nowrap font-mono">
                              {String(row[k] ?? '')}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
            )}

            {/* Tab 5: Matched Records */}
            {activeTab === 'matched' && (
              <div className="overflow-x-auto max-h-[460px]">
                {comparisonResult.matchedRecords.length === 0 ? (
                  <div className="p-8 text-center text-xs text-gray-600">Zero matching records.</div>
                ) : (
                  <table className="w-full text-left text-xs border-collapse">
                    <thead className="sticky top-0 bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold">
                      <tr>
                        <th className="py-2.5 px-3 border-r border-gray-200 w-12 text-center">#</th>
                        {Object.keys(comparisonResult.matchedRecords[0]).map((h) => (
                          <th key={h} className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB] text-gray-900">
                      {comparisonResult.matchedRecords.map((row, idx) => (
                        <tr key={idx} className="hover:bg-gray-50">
                          <td className="py-2 px-3 border-r border-gray-200 text-center text-gray-600 text-[11px]">
                            {idx + 1}
                          </td>
                          {Object.keys(row).map((k) => (
                            <td key={k} className="py-2 px-3 border-r border-gray-200 whitespace-nowrap font-mono">
                              {String(row[k] ?? '')}
                            </td>
                          ))}
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