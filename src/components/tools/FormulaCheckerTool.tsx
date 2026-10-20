import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  UploadCloud,
  Download,
  CheckCircle2,
  AlertTriangle,
  Calculator,
  FileSpreadsheet,
  AlertOctagon,
  Sparkles,
  ShieldAlert,
  ShieldCheck,
  Check
} from 'lucide-react';
import {
  auditWorkbookFormulas,
  applyApprovedFormulaFixes,
  createFormulaAuditWorkbook,
  type FormulaAuditSummary,
  type FormulaIssue,
} from '../../utils/formulaCheckerEngine';
import { readWorkbookPreserved } from '../../utils/workbookPreserver';
import { DataFormatGuide } from '../common/DataFormatGuide';
import { ToolHeader } from '../common/ToolHeader';

interface FormulaCheckerToolProps {
  onBack: () => void;
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
}

export const FormulaCheckerTool: React.FC<FormulaCheckerToolProps> = ({ onBack, onShowToast }) => {
  const [file, setFile] = useState<{ name: string; size: string } | null>(null);
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [auditSummary, setAuditSummary] = useState<FormulaAuditSummary | null>(null);
  const [issues, setIssues] = useState<FormulaIssue[]>([]);
  const [activeSeverityFilter, setActiveSeverityFilter] = useState<'all' | 'Critical' | 'Warning' | 'Advisory'>('all');
  const [validationError, setValidationError] = useState<string | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;
    setValidationError(null);

    const fileNameLower = f.name.toLowerCase();
    if (!fileNameLower.endsWith('.xlsx') && !fileNameLower.endsWith('.xls')) {
      setValidationError('Please upload a Microsoft Excel workbook (.xlsx or .xls).');
      onShowToast('Invalid Format', 'Supported formats: .xlsx, .xls', 'error');
      return;
    }

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const buffer = evt.target?.result as ArrayBuffer;
        const wb = readWorkbookPreserved(buffer);
        if (!wb.SheetNames || wb.SheetNames.length === 0) throw new Error('Workbook contains no sheets.');

        setFile({ name: f.name, size: `${Math.round(f.size / 1024)} KB` });
        setWorkbook(wb);

        const summary = auditWorkbookFormulas(wb, f.name);
        setAuditSummary(summary);
        setIssues(summary.issues);
      } catch (err: any) {
        setValidationError(`Failed to parse workbook: ${err.message}`);
        onShowToast('Parse Error', err.message, 'error');
      }
    };
    reader.readAsArrayBuffer(f);
  };


  const handleReset = () => {
    setFile(null);
    setWorkbook(null);
    setAuditSummary(null);
    setIssues([]);
    setValidationError(null);
    if (fileInputRef.current) fileInputRef.current.value = '';
    onShowToast('Reset Complete', 'Workbook cleared.', 'info');
  };

  const handleToggleIssueApproval = (id: string) => {
    setIssues((prev) =>
      prev.map((i) => (i.id === id ? { ...i, approved: !i.approved } : i))
    );
  };

  const handleApproveAll = (onlyCritical = false) => {
    setIssues((prev) =>
      prev.map((i) => {
        if (onlyCritical && i.severity !== 'Critical') return i;
        return { ...i, approved: true };
      })
    );
    onShowToast('Approvals Updated', onlyCritical ? 'Approved all critical errors for fix.' : 'Approved all issues.', 'success');
  };

  const handleDownloadCorrected = () => {
    if (!workbook) return;
    const approvedIssues = issues.filter((i) => i.approved);
    if (approvedIssues.length === 0) {
      onShowToast('No Fixes Approved', 'Check at least one issue approval box before applying fixes.', 'info');
      return;
    }

    const { correctedWorkbook, appliedCount } = applyApprovedFormulaFixes(workbook, approvedIssues);
    const outName = `Corrected_${file?.name || 'Workbook.xlsx'}`;
    XLSX.writeFile(correctedWorkbook, outName);
    onShowToast('Corrected Workbook Downloaded', `Applied ${appliedCount} approved formula correction(s).`, 'success');
  };

  const handleDownloadAuditReport = () => {
    if (!auditSummary) return;
    const wb = createFormulaAuditWorkbook({ ...auditSummary, issues });
    const safeName = (file?.name || 'Workbook').replace(/\.[^/.]+$/, '');
    const outName = `Formula_Audit_Report_${safeName}.xlsx`;
    XLSX.writeFile(wb, outName);
    onShowToast('Audit Report Downloaded', `Saved 2-sheet audit report as ${outName}.`, 'success');
  };

  const filteredIssues = issues.filter((i) => {
    if (activeSeverityFilter === 'all') return true;
    return i.severity === activeSeverityFilter;
  });

  const approvedCount = issues.filter((i) => i.approved).length;

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="formula-checker"
        title="Spreadsheet Formula & Integrity Auditor"
        description="Inspect spreadsheets for #REF!, #DIV/0!, broken references, and hardcoded numbers in formula ranges."
        onBack={onBack}
        actionButtons={
          (file || auditSummary) ? (
            <div className="flex items-center gap-2">
              <button
                onClick={handleReset}
                className="px-3 py-1.5 text-xs font-medium text-gray-600 hover:text-gray-900 bg-white hover:bg-gray-50 border border-gray-200 rounded-lg transition-colors cursor-pointer"
              >
                Clear
              </button>
              {auditSummary && (
                <>
                  <button
                    onClick={handleDownloadAuditReport}
                    className="px-3 py-2 text-xs font-semibold text-gray-900 bg-white hover:bg-gray-50 rounded-lg border border-gray-200 shadow-sm flex items-center gap-1.5 transition-colors cursor-pointer"
                  >
                    <Download className="w-4 h-4 text-gray-600" />
                    <span>Export Audit Log (.xlsx)</span>
                  </button>
                  <button
                    onClick={handleDownloadCorrected}
                    className="px-4 py-2 text-xs font-semibold text-white bg-gray-900 hover:bg-black rounded-lg shadow-sm flex items-center gap-2 transition-colors cursor-pointer"
                  >
                    <Sparkles className="w-4 h-4" />
                    <span>Apply Fixes & Export ({approvedCount})</span>
                  </button>
                </>
              )}
            </div>
          ) : null
        }
      />

      {/* 2. Format Requirements Guide */}
      <DataFormatGuide
        toolType="excel-cleaner"
        acceptedFormats=".xlsx, .xls"
        requiredFields={[
          { name: 'Excel Formulas', description: 'Formulas prefixed with = in any worksheet', aliases: ['SUM', 'AVERAGE', 'VLOOKUP', 'IF', 'INDEX/MATCH'] },
          { name: 'Model Structure', description: 'Financial statements, projections, or working paper workbooks', aliases: ['Balance Sheet', 'P&L', 'Trial Balance', 'Schedules'] },
        ]}
        notes={[
          'Core Safety Rule: Never modifies formulas automatically. You review each suggestion and approve fixes individually.',
          'Formula Range Traps: Flags hardcoded numbers placed inside columns where neighboring rows calculate dynamically.',
          'Text Numbers: Detects numbers stored as text that are silently excluded by Excel SUM() functions.',
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

      {/* 4. Upload & Action Bar */}
      <input type="file" ref={fileInputRef} onChange={handleFileUpload} accept=".xlsx,.xls" className="hidden" id="formula-checker-upload" />

      {!file ? (
        <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-8 sm:p-10 text-center transition-colors shadow-2xs">
          <label htmlFor="formula-checker-upload" className="cursor-pointer flex flex-col items-center">
            <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mx-auto mb-3 shadow-2xs">
              <Calculator className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-semibold text-gray-900">Upload Excel Workbook (.xlsx, .xls) to Audit</h3>
            <p className="text-xs text-gray-500 mt-1 max-w-md mx-auto leading-relaxed">
              Inspect spreadsheets for #REF!, #DIV/0!, broken references, hardcoded numbers in formula ranges, and text formatted numbers across all sheets.
            </p>
            <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
              <UploadCloud className="w-4 h-4" />
              <span>Upload Excel Workbook</span>
            </div>
            <span className="text-[11px] text-gray-400 font-medium mt-3 block">
              Supports .xlsx, .xls spreadsheets · Formula definitions and cell styles are 100% preserved
            </span>
          </label>
        </div>
      ) : (
        <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-sm">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-lg bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 shrink-0">
                <Calculator className="w-5 h-5" />
              </div>
              <div>
                <span className="text-xs font-bold text-gray-900 block">
                  {file.name}
                </span>
                <span className="text-[11px] text-gray-600 block mt-0.5">
                  {auditSummary
                    ? `${file.size} • ${auditSummary.sheetsAudited} sheet(s) • ${auditSummary.totalFormulasScanned} formulas inspected`
                    : file.size}
                </span>
              </div>
            </div>
            <div className="flex items-center gap-2">
              <button
                onClick={() => fileInputRef.current?.click()}
                className="px-3.5 py-1.5 text-xs text-gray-900 bg-gray-50 hover:bg-[#EAEAEA] rounded-lg border border-gray-200 flex items-center gap-1.5 font-medium transition-colors cursor-pointer"
              >
                <UploadCloud className="w-3.5 h-3.5 text-gray-600" />
                <span>Change Workbook</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 5. Audit Results Workspace */}
      {auditSummary && (
        <div className="space-y-4">
          {/* Executive Metrics Overview */}
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <span className="text-[11px] font-semibold text-gray-600 block">Critical Errors</span>
              <span className="text-xl font-bold text-[#C43D43] mt-1 block">
                {auditSummary.criticalErrorsCount}
              </span>
              <span className="text-[10px] text-[#C43D43] block mt-0.5">#REF!, #DIV/0!, #VALUE!</span>
            </div>

            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <span className="text-[11px] font-semibold text-gray-600 block">Hardcoded in Formulas</span>
              <span className="text-xl font-bold text-[#B45309] mt-1 block">
                {auditSummary.hardcodedConstantsCount}
              </span>
              <span className="text-[10px] text-[#B45309] block mt-0.5">Formula pattern broken</span>
            </div>

            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <span className="text-[11px] font-semibold text-gray-600 block">Text Stored as Numbers</span>
              <span className="text-xl font-bold text-gray-900 mt-1 block">
                {auditSummary.textNumbersCount}
              </span>
              <span className="text-[10px] text-gray-900 block mt-0.5">Ignored by SUM()</span>
            </div>

            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <span className="text-[11px] font-semibold text-gray-600 block">Approved for Fix</span>
              <span className="text-xl font-bold text-gray-900 mt-1 block">
                {approvedCount} / {issues.length}
              </span>
              <span className="text-[10px] text-gray-900 font-medium block mt-0.5">Ready to export</span>
            </div>
          </div>

          {/* Issues Table */}
          <div className="bg-white rounded-xl border border-gray-200 overflow-hidden shadow-sm">
            <div className="px-5 py-3 bg-gray-50 border-b border-gray-200 flex flex-wrap items-center justify-between gap-3 text-xs">
              <div className="flex items-center gap-2">
                <span className="font-bold text-gray-900">Detected Formula Issues & Corrections</span>
                <span className="text-gray-600">({filteredIssues.length} issues shown)</span>
              </div>

              <div className="flex items-center gap-2">
                {/* Severity Filter */}
                <div className="flex items-center gap-1 bg-gray-100 p-0.5 rounded-lg border border-gray-200">
                  <button
                    onClick={() => setActiveSeverityFilter('all')}
                    className={`px-2 py-0.5 rounded text-[11px] font-medium transition-colors ${
                      activeSeverityFilter === 'all' ? 'bg-white text-gray-900 shadow-sm font-semibold' : 'text-gray-600'
                    }`}
                  >
                    All ({issues.length})
                  </button>
                  <button
                    onClick={() => setActiveSeverityFilter('Critical')}
                    className={`px-2 py-0.5 rounded text-[11px] font-medium transition-colors ${
                      activeSeverityFilter === 'Critical' ? 'bg-white text-[#C43D43] shadow-sm font-semibold' : 'text-gray-600'
                    }`}
                  >
                    Critical ({auditSummary.criticalErrorsCount})
                  </button>
                  <button
                    onClick={() => setActiveSeverityFilter('Warning')}
                    className={`px-2 py-0.5 rounded text-[11px] font-medium transition-colors ${
                      activeSeverityFilter === 'Warning' ? 'bg-white text-[#B45309] shadow-sm font-semibold' : 'text-gray-600'
                    }`}
                  >
                    Warning ({auditSummary.hardcodedConstantsCount})
                  </button>
                </div>

                <button
                  onClick={() => handleApproveAll(false)}
                  className="px-2.5 py-1 text-[11px] font-semibold text-gray-900 bg-gray-100 hover:bg-[#E5E7EB] rounded-md transition-colors"
                >
                  Approve All Fixes
                </button>
              </div>
            </div>

            <div className="overflow-x-auto max-h-[480px]">
              {filteredIssues.length === 0 ? (
                <div className="p-8 text-center text-xs text-gray-900 font-semibold flex items-center justify-center gap-2">
                  <CheckCircle2 className="w-5 h-5 text-gray-900" />
                  <span>Zero formula errors or inconsistencies detected in workbook.</span>
                </div>
              ) : (
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="sticky top-0 bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold">
                    <tr>
                      <th className="py-2.5 px-3 border-r border-gray-200 w-12 text-center">Fix?</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Sheet & Cell</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Severity</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Current Value / Formula</th>
                      <th className="py-2.5 px-3 border-r border-gray-200">Audit Finding</th>
                      <th className="py-2.5 px-3">Suggested Correction</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#E5E7EB] text-gray-900">
                    {filteredIssues.map((issue) => (
                      <tr
                        key={issue.id}
                        onClick={() => handleToggleIssueApproval(issue.id)}
                        className={`cursor-pointer transition-colors ${
                          issue.approved ? 'bg-slate-50' : 'hover:bg-gray-50'
                        }`}
                      >
                        <td className="py-2 px-3 border-r border-gray-200 text-center" onClick={(e) => e.stopPropagation()}>
                          <input
                            type="checkbox"
                            checked={issue.approved}
                            onChange={() => handleToggleIssueApproval(issue.id)}
                            className="rounded text-gray-900 focus:ring-gray-900"
                          />
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 whitespace-nowrap font-mono font-semibold text-gray-900">
                          {issue.sheetName}!{issue.cellAddress}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 whitespace-nowrap">
                          <span
                            className={`inline-block px-2 py-0.5 rounded text-[10px] font-semibold ${
                              issue.severity === 'Critical'
                                ? 'bg-[#FDF2F2] text-[#C43D43] border border-[#F8D7DA]'
                                : issue.severity === 'Warning'
                                ? 'bg-[#FEF9EE] text-[#B45309] border border-[#FDE68A]'
                                : 'bg-gray-100 text-gray-900 border border-gray-200'
                            }`}
                          >
                            {issue.severity}
                          </span>
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 whitespace-nowrap font-mono text-[11px] text-[#C43D43] font-semibold max-w-[200px] truncate">
                          {issue.currentFormula ? `=${issue.currentFormula}` : String(issue.currentValue ?? '')}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 text-gray-900">
                          {issue.description}
                        </td>
                        <td className="py-2 px-3 text-gray-900 font-medium">
                          {issue.suggestedFix}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
};