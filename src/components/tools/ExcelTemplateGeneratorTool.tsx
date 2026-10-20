import { UploadIllustration } from '../common/UploadIllustration';
import { InnerToolBadge } from '../common/InnerToolBadge';
import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  Download,
  FileSpreadsheet,
  Building2,
  Calendar,
  Layers,
  Info,
  Check,
  FileCheck,
  Award
} from 'lucide-react';
import {
  type TemplateId,
  type TemplateCategory,
  TEMPLATES,
  generateAccountingTemplateWorkbook,
} from '../../utils/templateGeneratorEngine';
import { DataFormatGuide } from '../common/DataFormatGuide';

interface ExcelTemplateGeneratorToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

export const ExcelTemplateGeneratorTool: React.FC<ExcelTemplateGeneratorToolProps> = ({
  onBack,
  onShowToast,
}) => {
  const [selectedTemplateId, setSelectedTemplateId] = useState<TemplateId>('bank_rec');
  const [companyName, setCompanyName] = useState('');
  const [financialYear, setFinancialYear] = useState('FY 2026-27');
  const [includeSampleData, setIncludeSampleData] = useState(true);
  const [includeInstructions, setIncludeInstructions] = useState(true);
  const [selectedCategory, setSelectedCategory] = useState<string>('All');

  const currentTemplate = TEMPLATES[selectedTemplateId];

  const handleDownloadTemplate = () => {
    try {
      const wb = generateAccountingTemplateWorkbook({
        templateId: selectedTemplateId,
        companyName,
        financialYear,
        includeSampleData,
        includeInstructionsSheet: includeInstructions,
      });

      const fileName = `${currentTemplate.sheetName}_${companyName.replace(/\s+/g, '_')}_${financialYear.replace(/\s+/g, '_')}.xlsx`;
      XLSX.writeFile(wb, fileName);
      onShowToast('Template Generated', `Downloaded ${currentTemplate.name} successfully.`, 'success');
    } catch (err: any) {
      onShowToast('Generation Error', err.message || 'Failed to generate template workbook.', 'error');
    }
  };

  const previewData = includeSampleData ? currentTemplate.sampleData : currentTemplate.blankData;

  const categories = [
    'All',
    'Reconciliation',
    'Tax Audit (Form 3CD)',
    'Statutory Audit',
    'Taxation & TDS',
    'Corporate Accounting',
  ];

  const filteredTemplates = Object.values(TEMPLATES).filter(
    (t) => selectedCategory === 'All' || t.category === selectedCategory
  );

  return (
    <div className="space-y-6">
      {/* Top Header with Breadcrumb Back Link */}
      <div>
        <button
          onClick={onBack}
          className="inline-flex items-center gap-2 px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-700 bg-white border border-slate-200 shadow-2xs hover:bg-slate-50 hover:text-slate-900 hover:border-slate-300 transition-all cursor-pointer mb-5 group"
        >
          <ArrowLeft className="w-3.5 h-3.5 text-slate-400 group-hover:text-slate-900 group-hover:-translate-x-0.5 transition-transform" />
          <span>Back to All Tools</span>
        </button>

        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-5 border-b border-slate-200">
          <div className="flex items-center gap-3.5">
            <InnerToolBadge toolId="excel-templates" />
            <div>
              <h1 className="text-xl font-bold text-slate-900 tracking-tight">Standardized CA Excel Templates</h1>
              <p className="text-xs text-slate-500 mt-0.5">
                ICAI, Companies Act 2013 & Income Tax compliant working papers engineered with strict accounting symmetry, professional typography, and audit trails.
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <button
              onClick={handleDownloadTemplate}
              className="flex items-center gap-1.5 px-4 py-2 bg-slate-900 hover:bg-slate-800 text-white rounded-lg text-xs font-semibold shadow-xs transition-colors cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              Download Excel Template (.xlsx)
            </button>
          </div>
        </div>
      </div>

      <DataFormatGuide
        toolType="excel-cleaner"
        acceptedFormats=".xlsx (Microsoft Excel & LibreOffice Calc compatible)"
        requiredFields={[
          { name: 'Template Standard', description: 'ICAI SA 230 & Statutory Regulation', aliases: ['SA 500', 'Sec 40A(3)', 'Sec 43B(h)', 'Rule 88C', 'Sch III'] },
          { name: 'Entity & Financial Year', description: 'Header metadata for audit working papers', aliases: ['Client', 'Assessee', 'FY', 'AY'] },
        ]}
        notes={[
          'Strict professional CA layout: Symmetrical column widths (wch), freeze panes on header row, and auto-filters applied.',
          'Classical accounting number formatting: Indian commas (Lakhs & Crores) with two decimals for monetary figures.',
          'Working paper audit sign-off block and statutory reference included on every schedule.',
          'Option to switch between Clean Blank Formats for live client input or Pre-filled Sample CA Practice Data with explanatory notes.',
        ]}
      />

      {/* Category Filter Pills */}
      <div className="flex flex-wrap items-center gap-1.5">
        {categories.map((cat) => {
          const count =
            cat === 'All'
              ? Object.keys(TEMPLATES).length
              : Object.values(TEMPLATES).filter((t) => t.category === cat).length;

          return (
            <button
              key={cat}
              onClick={() => setSelectedCategory(cat)}
              className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all cursor-pointer flex items-center gap-1.5 ${
                selectedCategory === cat
                  ? 'bg-slate-900 text-white shadow-xs'
                  : 'bg-white border border-slate-200 text-slate-600 hover:bg-slate-50 hover:text-slate-900'
              }`}
            >
              <span>{cat}</span>
              <span
                className={`text-[10px] px-1.5 py-0.2 rounded-full font-mono ${
                  selectedCategory === cat ? 'bg-slate-800 text-slate-300' : 'bg-slate-100 text-slate-500'
                }`}
              >
                {count}
              </span>
            </button>
          );
        })}
      </div>

      {/* Grid of Templates */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
        {filteredTemplates.map((tmpl) => {
          const isSelected = selectedTemplateId === tmpl.id;

          return (
            <div
              key={tmpl.id}
              onClick={() => setSelectedTemplateId(tmpl.id)}
              className={`p-4 rounded-xl border transition-all cursor-pointer text-left flex flex-col justify-between ${
                isSelected
                  ? 'bg-white border-gray-900 ring-1 ring-gray-900 shadow-2xs'
                  : 'bg-white border-gray-200 hover:border-gray-300 hover:shadow-2xs'
              }`}
            >
              <div>
                <div className="flex items-center justify-between gap-2 mb-1">
                  <h3 className="text-xs font-bold text-gray-900 tracking-tight">{tmpl.name}</h3>
                  {isSelected && (
                    <span className="w-4 h-4 rounded-full bg-gray-900 text-white flex items-center justify-center shrink-0">
                      <Check className="w-2.5 h-2.5" />
                    </span>
                  )}
                </div>
                <p className="text-[11px] text-gray-500 font-medium truncate">
                  {tmpl.statutoryStandard}
                </p>
                <p className="text-[11px] text-gray-600 mt-1.5 leading-relaxed line-clamp-2">
                  {tmpl.description}
                </p>
              </div>
            </div>
          );
        })}
      </div>

      {/* Customization Options Bar */}
      <div className="bg-white border border-slate-200 rounded-xl p-4 shadow-2xs space-y-3">
        <div className="flex items-center justify-between">
          <h3 className="text-xs font-bold text-slate-900 uppercase tracking-wider flex items-center gap-1.5">
            <Building2 className="w-3.5 h-3.5 text-slate-500" />
            Working Paper Engagement Properties
          </h3>
          <span className="text-[10px] text-slate-500 font-mono">ICAI SA 230 Compliant Header</span>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 text-xs">
          <div>
            <label className="text-[11px] font-semibold text-slate-600 block mb-1">Client Entity / Assessee Name</label>
            <input
              type="text"
              value={companyName}
              onChange={(e) => setCompanyName(e.target.value)}
              className="w-full px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-xs text-slate-900 focus:outline-none focus:border-slate-900"
            />
          </div>
          <div>
            <label className="text-[11px] font-semibold text-slate-600 block mb-1">Financial Year / Period</label>
            <select
              value={financialYear}
              onChange={(e) => setFinancialYear(e.target.value)}
              className="w-full px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-xs text-slate-900 focus:outline-none focus:border-slate-900"
            >
              <option value="FY 2030-31">FY 2030-31 (AY 2031-32)</option>
              <option value="FY 2029-30">FY 2029-30 (AY 2030-31)</option>
              <option value="FY 2028-29">FY 2028-29 (AY 2029-30)</option>
              <option value="FY 2027-28">FY 2027-28 (AY 2028-29)</option>
              <option value="FY 2026-27">FY 2026-27 (AY 2027-28)</option>
              <option value="FY 2025-26">FY 2025-26 (AY 2026-27)</option>
              <option value="FY 2024-25">FY 2024-25 (AY 2025-26)</option>
              <option value="FY 2023-24">FY 2023-24 (AY 2024-25)</option>
              <option value="FY 2022-23">FY 2022-23 (AY 2023-24)</option>
            </select>
          </div>
          <div>
            <label className="text-[11px] font-semibold text-slate-600 block mb-1">Working Paper Data Mode</label>
            <select
              value={includeSampleData ? 'sample' : 'blank'}
              onChange={(e) => setIncludeSampleData(e.target.value === 'sample')}
              className="w-full px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-xs text-slate-900 focus:outline-none focus:border-slate-900"
            >
              <option value="sample">Pre-filled Sample CA Practice Data</option>
              <option value="blank">Clean Blank Working Paper (Empty Rows)</option>
            </select>
          </div>
          <div className="flex items-end pb-1.5">
            <label className="flex items-center gap-2 cursor-pointer text-xs text-slate-900">
              <input
                type="checkbox"
                checked={includeInstructions}
                onChange={(e) => setIncludeInstructions(e.target.checked)}
                className="rounded border-slate-300 text-slate-900 focus:ring-slate-900 cursor-pointer"
              />
              <span className="font-medium text-[11px]">Include ICAI Audit Guidance Sheet</span>
            </label>
          </div>
        </div>
      </div>

      {/* Live Template Structure Preview */}
      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-2xs">
        <div className="p-4 border-b border-slate-200 bg-slate-50/70 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <div className="flex items-center gap-2">
              <FileSpreadsheet className="w-4 h-4 text-slate-900" />
              <h3 className="text-xs font-bold text-slate-900">{currentTemplate.name}</h3>
              <span className="text-[10px] text-slate-500 font-mono">Sheet: "{currentTemplate.sheetName}"</span>
            </div>
            <p className="text-[11px] text-slate-500 mt-0.5 flex items-center gap-1.5">
              <span>Standard: {currentTemplate.statutoryStandard}</span>
              <span>•</span>
              <span className="text-slate-700 font-medium">{includeSampleData ? 'Displaying Sample CA Audit Trail' : 'Displaying Blank Formatted Grid'}</span>
            </p>
          </div>
          <button
            onClick={handleDownloadTemplate}
            className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-900 hover:bg-slate-800 text-white text-xs font-semibold rounded-lg shadow-xs self-start sm:self-auto transition-colors cursor-pointer"
          >
            <Download className="w-3.5 h-3.5" />
            Download (.xlsx)
          </button>
        </div>

        {/* Standard CA Working Paper Header Banner Preview */}
        <div className="px-4 py-2.5 bg-slate-100/60 border-b border-slate-200 text-[11px] grid grid-cols-1 sm:grid-cols-3 gap-2 text-slate-700">
          <div>
            <span className="text-slate-500 font-medium">Assessee:</span> <strong className="text-slate-900">{companyName}</strong>
          </div>
          <div>
            <span className="text-slate-500 font-medium">Period / FY:</span> <strong className="text-slate-900">{financialYear}</strong>
          </div>
          <div className="sm:text-right">
            <span className="text-slate-500 font-medium">ICAI Audit Paper Ref:</span> <span className="font-mono font-semibold text-slate-900">WP-CA-{currentTemplate.id.slice(0, 4).toUpperCase()}-01</span>
          </div>
        </div>

        {/* Instructions Snippet */}
        <div className="p-3 bg-slate-50/40 border-b border-slate-200 text-xs text-slate-900">
          <div className="font-semibold text-[11px] mb-1 flex items-center gap-1 text-slate-800">
            <Info className="w-3.5 h-3.5 text-slate-500" />
            Standard Operating Procedures (SOP) & ICAI Working Paper Guidelines:
          </div>
          <ul className="list-disc list-inside space-y-0.5 text-[11px] text-slate-600 pl-1">
            {currentTemplate.instructions.map((inst, i) => (
              <li key={i}>{inst}</li>
            ))}
          </ul>
        </div>

        {/* Table Preview with Symmetrical CA Grid Styling */}
        <div className="overflow-x-auto max-h-[420px]">
          <table className="w-full text-left text-xs border-collapse">
            <thead className="bg-slate-100 border-b border-slate-300 text-slate-700 uppercase font-bold text-[10px] tracking-wider sticky top-0 z-10 shadow-2xs">
              <tr>
                <th className="py-2.5 px-3 w-10 text-center border-r border-slate-200">#</th>
                {currentTemplate.columns.map((col, idx) => {
                  const isCurrency = currentTemplate.currencyColumns?.includes(col);
                  const isDate = currentTemplate.dateColumns?.includes(col) || col.toLowerCase().includes('date');
                  const isCenter = isDate || col.toLowerCase().includes('s.no') || col.toLowerCase().includes('status');

                  return (
                    <th
                      key={col}
                      className={`py-2.5 px-3 whitespace-nowrap border-r border-slate-200 last:border-r-0 ${
                        isCurrency ? 'text-right' : isCenter ? 'text-center' : 'text-left'
                      }`}
                    >
                      {col}
                    </th>
                  );
                })}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {previewData.map((row, idx) => {
                const isTotalRow =
                  idx === previewData.length - 1 &&
                  (String(row[currentTemplate.columns[0]] || '').toUpperCase().includes('TOTAL') ||
                    String(row[currentTemplate.columns[1]] || '').toUpperCase().includes('TOTAL') ||
                    String(row[currentTemplate.columns[0]] || '').includes('Balance as per Official'));

                return (
                  <tr
                    key={idx}
                    className={`transition-colors ${
                      isTotalRow
                        ? 'bg-slate-100 font-bold border-t-2 border-b-2 border-slate-800 text-slate-900'
                        : idx % 2 === 0
                        ? 'bg-white hover:bg-slate-50/70'
                        : 'bg-slate-50/30 hover:bg-slate-50'
                    }`}
                  >
                    <td className={`py-2 px-3 text-center text-slate-500 font-mono text-[11px] border-r border-slate-200 ${isTotalRow ? 'font-bold text-slate-900' : ''}`}>
                      {isTotalRow ? 'Σ' : idx + 1}
                    </td>
                    {currentTemplate.columns.map((col) => {
                      const val = row[col];
                      const isCurrency = currentTemplate.currencyColumns?.includes(col);
                      const isDate = currentTemplate.dateColumns?.includes(col) || col.toLowerCase().includes('date');
                      const isCenter = isDate || col.toLowerCase().includes('status') || col.toLowerCase().includes('(y/n)');
                      const isCode =
                        col.toLowerCase().includes('pan') ||
                        col.toLowerCase().includes('gstin') ||
                        col.toLowerCase().includes('hsn') ||
                        col.toLowerCase().includes('voucher') ||
                        col.toLowerCase().includes('ref') ||
                        col.toLowerCase().includes('code');

                      const formatted =
                        typeof val === 'number' && isCurrency
                          ? val.toLocaleString('en-IN', { style: 'currency', currency: 'INR', maximumFractionDigits: 0 })
                          : typeof val === 'number'
                          ? val.toLocaleString('en-IN')
                          : val !== undefined && val !== null
                          ? String(val)
                          : '-';

                      const isStatusCol = col.toLowerCase().includes('status') || col.toLowerCase().includes('eligibility');

                      return (
                        <td
                          key={col}
                          className={`py-2 px-3 whitespace-nowrap text-xs border-r border-slate-200 last:border-r-0 ${
                            isCurrency
                              ? 'font-mono text-right text-slate-900 font-medium'
                              : isCenter
                              ? 'text-center'
                              : isCode
                              ? 'font-mono text-slate-800 font-medium'
                              : 'text-slate-800'
                          } ${isTotalRow ? 'font-bold' : ''}`}
                        >
                          {isStatusCol && !isTotalRow && typeof val === 'string' && val.length > 0 ? (
                            <span
                              className={`inline-block px-2 py-0.5 rounded text-[10px] font-semibold border ${
                                val.toLowerCase().includes('eligible') ||
                                val.toLowerCase().includes('verified') ||
                                val.toLowerCase().includes('compliant') ||
                                val.toLowerCase().includes('matched') ||
                                val.toLowerCase().includes('yes')
                                  ? 'bg-gray-50 text-gray-900 border-gray-200'
                                  : val.toLowerCase().includes('disallow') ||
                                    val.toLowerCase().includes('ineligible') ||
                                    val.toLowerCase().includes('blocked') ||
                                    val.toLowerCase().includes('overdue')
                                  ? 'bg-rose-50 text-rose-800 border-rose-200'
                                  : 'bg-slate-100 text-slate-700 border-slate-200'
                              }`}
                            >
                              {val}
                            </span>
                          ) : (
                            formatted || '-'
                          )}
                        </td>
                      );
                    })}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};
