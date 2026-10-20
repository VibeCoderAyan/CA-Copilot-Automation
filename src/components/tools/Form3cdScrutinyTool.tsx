import { ToolHeader } from '../common/ToolHeader';
import React, { useState, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  FileSpreadsheet,
  Download,
  AlertTriangle,
  Sliders,
  Scale,
  UploadCloud,
  CheckCircle2,
} from 'lucide-react';
import { extractSmartTableFromSheet } from '../../utils/dataSanitizer';
import {
  type Form3cdAuditSummary,
  detectForm3cdColumns,
  scrutinizeForm3cdRecords,
  exportForm3cdAnnexuresWorkbook,
} from '../../utils/form3cdEngine';

interface Form3cdScrutinyToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

export const Form3cdScrutinyTool: React.FC<Form3cdScrutinyToolProps> = ({
  onBack,
  onShowToast,
}) => {
  const [rawRows, setRawRows] = useState<Record<string, any>[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [fileName, setFileName] = useState('');
  const [clientName, setClientName] = useState('');
  const [financialYear, setFinancialYear] = useState('FY 2025-26');
  const [assessmentYear, setAssessmentYear] = useState('AY 2026-27');
  const [activeTab, setActiveTab] = useState<'overview' | 'clause40a3' | 'clause44' | 'clause34' | 'clause21a'>('overview');
  const [searchQuery, setSearchQuery] = useState('');
  const [showMappingPanel, setShowMappingPanel] = useState(false);
  const [rule6DdOverrides, setRule6DdOverrides] = useState<Record<string, boolean>>({});

  const [mapping, setMapping] = useState({
    voucherNoCol: '',
    dateCol: '',
    accountCol: '',
    payeeCol: '',
    amountCol: '',
    modeCol: '',
    narrationCol: '',
    panCol: '',
    gstinCol: '',
  });

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Run Form 3CD Audit Scrutiny
  const summary: Form3cdAuditSummary = useMemo(() => {
    if (rawRows.length === 0) {
      return {
        clientName,
        assessmentYear,
        financialYear,
        totalExpenditureScrutinized: 0,
        clause40a3ViolationsCount: 0,
        clause40a3TotalDisallowable: 0,
        clause40a3Rule6DdExemptCount: 0,
        totalClause44Expenditure: 0,
        totalRegisteredExpenditure: 0,
        totalUnregisteredExpenditure: 0,
        unregisteredExpenditureRatio: 0,
        clause34TotalCoveredExpenditure: 0,
        clause34TotalTdsRequired: 0,
        clause34TotalTdsDeducted: 0,
        clause34ShortDeduction: 0,
        clause34Disallowance40aIa: 0,
        clause21aPersonalAndPenalties: 0,
        clause22MsmeInterest: 0,
        totalTaxAdditions: 0,
        cashViolations: [],
        clause44Breakdown: [],
        clause34Summary: [],
        inadmissibleExpenses: [],
      };
    }

    const res = scrutinizeForm3cdRecords(rawRows, mapping, clientName, financialYear, assessmentYear);

    // Apply manual Rule 6DD toggles
    res.cashViolations = res.cashViolations.map((cv) => {
      if (rule6DdOverrides[cv.id] !== undefined) {
        const isExempt = rule6DdOverrides[cv.id];
        return {
          ...cv,
          isRule6DdExempt: isExempt,
          disallowableAmount: isExempt ? 0 : cv.dailyAggregateAmount,
        };
      }
      return cv;
    });

    // Recompute total disallowance
    const recalculated40a3 = res.cashViolations
      .filter((cv) => !cv.isRule6DdExempt)
      .reduce((acc, cv) => acc + cv.dailyAggregateAmount, 0);

    res.clause40a3TotalDisallowable = recalculated40a3;
    res.clause40a3ViolationsCount = res.cashViolations.filter((cv) => !cv.isRule6DdExempt).length;
    res.totalTaxAdditions =
      recalculated40a3 +
      res.clause34Disallowance40aIa +
      res.clause21aPersonalAndPenalties +
      res.clause22MsmeInterest;

    return res;
  }, [rawRows, mapping, clientName, financialYear, assessmentYear, rule6DdOverrides]);

  // Handle Excel upload
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const smartTable = extractSmartTableFromSheet(ws);

        if (!smartTable.rows || smartTable.rows.length === 0) {
          onShowToast('Empty File', 'Could not detect rows in the uploaded sheet', 'error');
          return;
        }

        setHeaders(smartTable.headers);
        setRawRows(smartTable.rows);

        // Auto-detect columns
        const detected = detectForm3cdColumns(smartTable.headers);
        setMapping({
          voucherNoCol: detected.voucherNoCol || smartTable.headers[0] || '',
          dateCol: detected.dateCol || smartTable.headers[1] || '',
          accountCol: detected.accountCol || smartTable.headers[2] || '',
          payeeCol: detected.payeeCol || smartTable.headers[3] || '',
          amountCol: detected.amountCol || smartTable.headers[4] || '',
          modeCol: detected.modeCol || '',
          narrationCol: detected.narrationCol || '',
          panCol: detected.panCol || '',
          gstinCol: detected.gstinCol || '',
        });
      } catch (err: any) {
        onShowToast('Parsing Error', err?.message || 'Failed to read spreadsheet', 'error');
      }
    };

    reader.readAsArrayBuffer(file);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };


  const handleResetData = () => {
    setRawRows([]);
    setHeaders([]);
    setFileName('');
    setRule6DdOverrides({});
    onShowToast('Data Reset', 'Cleared all vouchers and mappings.', 'info');
  };
  const handleExport = () => {
    if (summary.totalExpenditureScrutinized === 0) {
      onShowToast('No Data', 'Please load data before exporting', 'error');
      return;
    }
    exportForm3cdAnnexuresWorkbook(summary);
    onShowToast('Export Complete', 'Form 3CD Multi-Sheet Annexure workbook downloaded successfully', 'success');
  };

  const toggleRule6Dd = (violId: string) => {
    setRule6DdOverrides((prev) => ({
      ...prev,
      [violId]: !prev[violId],
    }));
  };

  return (
    <div className="space-y-6">
      {/* Hidden File Input */}
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        className="hidden"
        id="form3cd-file-upload"
        onChange={handleFileUpload}
      />

      <ToolHeader
        toolId="form3cd-scrutiny"
        title="Form 3CD Tax Audit Scrutiny"
        description="Scrutinize Clause 21(d) cash limits (> ₹10k/35k), Clause 44 GST spend break-up, Clause 34 Chapter XVII-B TDS, and Clause 21(a) disallowances."
        badge="Sec 44AB"
        onBack={onBack}
        actionButtons={
          <div className="flex items-center gap-2.5 flex-wrap">
            {rawRows.length > 0 && (
              <button
                onClick={handleResetData}
                className="px-3 py-1.5 text-xs font-medium text-gray-600 hover:text-gray-900 bg-white hover:bg-gray-50 border border-gray-200 rounded-lg transition-colors cursor-pointer"
              >
                Clear
              </button>
            )}

            {rawRows.length > 0 && (
              <button
                onClick={() => setShowMappingPanel(!showMappingPanel)}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg border border-gray-200 bg-white text-gray-700 hover:bg-gray-50 transition-colors cursor-pointer shadow-2xs"
              >
                <Sliders className="w-3.5 h-3.5 text-gray-500" />
                <span>Map Columns</span>
              </button>
            )}

            <button
              onClick={() => fileInputRef.current?.click()}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg bg-gray-900 text-white hover:bg-black transition-colors cursor-pointer shadow-xs"
            >
              <UploadCloud className="w-3.5 h-3.5" />
              <span>Upload Expense Ledger</span>
            </button>

            {rawRows.length > 0 && (
              <button
                onClick={handleExport}
                className="flex items-center gap-1.5 px-3.5 py-1.5 text-xs font-semibold rounded-lg bg-gray-900 text-white hover:bg-black transition-colors cursor-pointer shadow-xs"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Export Form 3CD Excel</span>
              </button>
            )}
          </div>
        }
      />

      {/* Column Mapping Drawer */}
      {showMappingPanel && (
        <div className="bg-white border border-gray-200 rounded-xl p-5 shadow-xs space-y-4">
          <div className="flex items-center justify-between pb-3 border-b border-gray-100">
            <div>
              <h3 className="text-sm font-bold text-gray-900">Form 3CD Column Mapping</h3>
              <p className="text-xs text-gray-500">Map uploaded columns to statutory tax audit audit fields</p>
            </div>
            <button
              onClick={() => setShowMappingPanel(false)}
              className="text-xs text-gray-500 hover:text-gray-900 font-medium cursor-pointer"
            >
              Close
            </button>
          </div>

          <div className="grid grid-cols-2 md:grid-cols-4 gap-4 text-xs">
            <div>
              <label className="block text-gray-600 mb-1 font-semibold">Voucher / Doc No</label>
              <select
                value={mapping.voucherNoCol}
                onChange={(e) => setMapping({ ...mapping, voucherNoCol: e.target.value })}
                className="w-full border border-gray-200 rounded-md p-1.5 outline-none bg-gray-50 focus:border-gray-900"
              >
                <option value="">-- Select --</option>
                {headers.map((h) => (
                  <option key={h} value={h}>{h}</option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-gray-600 mb-1 font-semibold">Voucher Date</label>
              <select
                value={mapping.dateCol}
                onChange={(e) => setMapping({ ...mapping, dateCol: e.target.value })}
                className="w-full border border-gray-200 rounded-md p-1.5 outline-none bg-gray-50 focus:border-gray-900"
              >
                <option value="">-- Select --</option>
                {headers.map((h) => (
                  <option key={h} value={h}>{h}</option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-gray-600 mb-1 font-semibold">Account / Expense Head</label>
              <select
                value={mapping.accountCol}
                onChange={(e) => setMapping({ ...mapping, accountCol: e.target.value })}
                className="w-full border border-gray-200 rounded-md p-1.5 outline-none bg-gray-50 focus:border-gray-900"
              >
                <option value="">-- Select --</option>
                {headers.map((h) => (
                  <option key={h} value={h}>{h}</option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-gray-600 mb-1 font-semibold">Payee / Vendor Name</label>
              <select
                value={mapping.payeeCol}
                onChange={(e) => setMapping({ ...mapping, payeeCol: e.target.value })}
                className="w-full border border-gray-200 rounded-md p-1.5 outline-none bg-gray-50 focus:border-gray-900"
              >
                <option value="">-- Select --</option>
                {headers.map((h) => (
                  <option key={h} value={h}>{h}</option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-gray-600 mb-1 font-semibold">Payment Amount</label>
              <select
                value={mapping.amountCol}
                onChange={(e) => setMapping({ ...mapping, amountCol: e.target.value })}
                className="w-full border border-gray-200 rounded-md p-1.5 outline-none bg-gray-50 focus:border-gray-900"
              >
                <option value="">-- Select --</option>
                {headers.map((h) => (
                  <option key={h} value={h}>{h}</option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-gray-600 mb-1 font-semibold">Payment Mode (Cash/Bank)</label>
              <select
                value={mapping.modeCol}
                onChange={(e) => setMapping({ ...mapping, modeCol: e.target.value })}
                className="w-full border border-gray-200 rounded-md p-1.5 outline-none bg-gray-50 focus:border-gray-900"
              >
                <option value="">-- Select --</option>
                {headers.map((h) => (
                  <option key={h} value={h}>{h}</option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-gray-600 mb-1 font-semibold">Narration / Description</label>
              <select
                value={mapping.narrationCol}
                onChange={(e) => setMapping({ ...mapping, narrationCol: e.target.value })}
                className="w-full border border-gray-200 rounded-md p-1.5 outline-none bg-gray-50 focus:border-gray-900"
              >
                <option value="">-- Select --</option>
                {headers.map((h) => (
                  <option key={h} value={h}>{h}</option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-gray-600 mb-1 font-semibold">Payee GSTIN (For Clause 44)</label>
              <select
                value={mapping.gstinCol}
                onChange={(e) => setMapping({ ...mapping, gstinCol: e.target.value })}
                className="w-full border border-gray-200 rounded-md p-1.5 outline-none bg-gray-50 focus:border-gray-900"
              >
                <option value="">-- Select --</option>
                {headers.map((h) => (
                  <option key={h} value={h}>{h}</option>
                ))}
              </select>
            </div>
          </div>
        </div>
      )}

      {/* Empty State / Dropzone */}
      {rawRows.length === 0 ? (
        <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-8 sm:p-10 text-center transition-colors shadow-2xs">
          <label htmlFor="form3cd-file-upload" className="cursor-pointer flex flex-col items-center">
            <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mx-auto mb-3 shadow-2xs">
              <Scale className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-semibold text-gray-900">Upload General Ledger or Expense Vouchers</h3>
            <p className="text-xs text-gray-500 mt-1 max-w-md mx-auto leading-relaxed">
              Upload client audited general ledger, payment daybook, or Tally voucher export (.xlsx, .xls, .csv). Automatically audits Clause 21(d) cash limits, Clause 44 GST spend, and Clause 34 TDS compliance.
            </p>
            <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
              <UploadCloud className="w-4 h-4" />
              <span>Upload General Ledger / Vouchers</span>
            </div>
            <span className="text-[11px] text-gray-400 font-medium mt-3 block">
              Section 44AB Tax Audit Rules · 100% offline private processing · Preserves all client records
            </span>
          </label>
        </div>
      ) : (
        <div className="space-y-6">
          {/* Executive KPI Summary Cards */}
          <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-xs">
              <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">Scrutinized Expenditure</span>
              <div className="text-xl font-bold text-gray-900 font-mono mt-1">
                ₹{summary.totalExpenditureScrutinized.toLocaleString('en-IN')}
              </div>
              <div className="text-xs text-gray-500 mt-0.5">
                Across {rawRows.length} vouchers
              </div>
            </div>

            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-xs">
              <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">Clause 21(d) / 40A(3) Cash</span>
              <div className="text-xl font-bold text-rose-600 font-mono mt-1">
                ₹{summary.clause40a3TotalDisallowable.toLocaleString('en-IN')}
              </div>
              <div className="text-xs text-gray-500 mt-0.5">
                {summary.clause40a3ViolationsCount} daily breaches
              </div>
            </div>

            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-xs">
              <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">Clause 44 Unregistered %</span>
              <div className={`text-xl font-bold font-mono mt-1 ${summary.unregisteredExpenditureRatio > 35 ? 'text-amber-600' : 'text-gray-900'}`}>
                {summary.unregisteredExpenditureRatio}%
              </div>
              <div className="text-xs text-gray-500 mt-0.5">
                ₹{summary.totalUnregisteredExpenditure.toLocaleString('en-IN')} unregistered
              </div>
            </div>

            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-xs">
              <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">Clause 34 TDS 40(a)(ia)</span>
              <div className="text-xl font-bold text-rose-600 font-mono mt-1">
                ₹{summary.clause34Disallowance40aIa.toLocaleString('en-IN')}
              </div>
              <div className="text-xs text-gray-500 mt-0.5">
                30% addition on non-deductions
              </div>
            </div>

            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-xs">
              <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">Total Form 3CD Additions</span>
              <div className="text-xl font-bold text-rose-700 font-mono mt-1">
                ₹{summary.totalTaxAdditions.toLocaleString('en-IN')}
              </div>
              <div className="text-xs text-gray-500 mt-0.5">
                Net addition to taxable profit
              </div>
            </div>
          </div>

          {/* Sub-Tabs Navigation */}
          <div className="bg-white border border-gray-200 rounded-xl p-1.5 flex items-center justify-between shadow-xs">
            <div className="flex items-center gap-1 overflow-x-auto">
              {[
                { id: 'overview' as const, label: 'Form 3CD Clause Readiness' },
                { id: 'clause40a3' as const, label: `Clause 21(d) Cash Limit (${summary.cashViolations.length})` },
                { id: 'clause44' as const, label: `Clause 44 GST Breakdown (${summary.clause44Breakdown.length})` },
                { id: 'clause34' as const, label: `Clause 34 Chapter XVII-B TDS (${summary.clause34Summary.length})` },
                { id: 'clause21a' as const, label: `Clause 21(a) & 22 Penalties (${summary.inadmissibleExpenses.length})` },
              ].map((tab) => (
                <button
                  key={tab.id}
                  onClick={() => setActiveTab(tab.id)}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all cursor-pointer whitespace-nowrap ${
                    activeTab === tab.id
                      ? 'bg-gray-900 text-white shadow-xs'
                      : 'text-gray-600 hover:text-gray-900 hover:bg-gray-100'
                  }`}
                >
                  {tab.label}
                </button>
              ))}
            </div>

            {/* Client Context Details */}
            <div className="hidden lg:flex items-center gap-3 text-xs text-gray-500 font-medium px-2">
              <span>Client: <strong className="text-gray-900">{clientName}</strong></span>
              <span>•</span>
              <span>{financialYear} ({assessmentYear})</span>
            </div>
          </div>

          {/* TAB 1: FORM 3CD READINESS & EXECUTIVE AUDIT REPORT */}
          {activeTab === 'overview' && (
            <div className="space-y-6">
              <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-xs">
                <h3 className="text-base font-bold text-gray-900 mb-1">
                  Tax Audit Readiness Schedule under Section 44AB
                </h3>
                <p className="text-xs text-gray-500 mb-6">
                  Executive summary of disallowances and disclosures to be incorporated into Form 3CD report for {financialYear}.
                </p>

                <div className="overflow-x-auto">
                  <table className="w-full text-xs text-left">
                    <thead className="bg-gray-50 border-y border-gray-200 text-gray-600 uppercase font-semibold">
                      <tr>
                        <th className="py-2.5 px-4">Form 3CD Clause</th>
                        <th className="py-2.5 px-4">Statutory Scope & Legal Provision</th>
                        <th className="py-2.5 px-4 text-right">Disallowable Addition</th>
                        <th className="py-2.5 px-4 text-center">Status</th>
                        <th className="py-2.5 px-4">Auditor Recommended Action</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-100 text-gray-700">
                      <tr>
                        <td className="py-3 px-4 font-bold text-gray-900">Clause 21(d)</td>
                        <td className="py-3 px-4">
                          <div>Section 40A(3) / 40A(3A)</div>
                          <div className="text-[11px] text-gray-500">Cash payments exceeding ₹10,000 / ₹35,000 daily aggregate</div>
                        </td>
                        <td className="py-3 px-4 text-right font-mono font-bold text-rose-600">
                          ₹{summary.clause40a3TotalDisallowable.toLocaleString('en-IN')}
                        </td>
                        <td className="py-3 px-4 text-center">
                          {summary.clause40a3ViolationsCount > 0 ? (
                            <span className="px-2 py-0.5 text-[11px] font-semibold rounded bg-rose-50 text-rose-700 border border-rose-200">
                              Disallowance Required
                            </span>
                          ) : (
                            <span className="px-2 py-0.5 text-[11px] font-semibold rounded bg-gray-50 text-gray-800 border border-gray-200">
                              Clean
                            </span>
                          )}
                        </td>
                        <td className="py-3 px-4 text-gray-600">
                          {summary.clause40a3ViolationsCount > 0
                            ? `Report ${summary.clause40a3ViolationsCount} daily breaches in Annexure 21(d)`
                            : 'Nil report under Clause 21(d)'}
                        </td>
                      </tr>

                      <tr>
                        <td className="py-3 px-4 font-bold text-gray-900">Clause 34(a)</td>
                        <td className="py-3 px-4">
                          <div>Chapter XVII-B & Section 40(a)(ia)</div>
                          <div className="text-[11px] text-gray-500">30% disallowance on non-deductions / short deductions of TDS</div>
                        </td>
                        <td className="py-3 px-4 text-right font-mono font-bold text-rose-600">
                          ₹{summary.clause34Disallowance40aIa.toLocaleString('en-IN')}
                        </td>
                        <td className="py-3 px-4 text-center">
                          {summary.clause34Disallowance40aIa > 0 ? (
                            <span className="px-2 py-0.5 text-[11px] font-semibold rounded bg-rose-50 text-rose-700 border border-rose-200">
                              Addition to Profit
                            </span>
                          ) : (
                            <span className="px-2 py-0.5 text-[11px] font-semibold rounded bg-gray-50 text-gray-800 border border-gray-200">
                              Compliant
                            </span>
                          )}
                        </td>
                        <td className="py-3 px-4 text-gray-600">
                          Annexure under Chapter XVII-B generated. Short deduction: ₹{summary.clause34ShortDeduction.toLocaleString('en-IN')}.
                        </td>
                      </tr>

                      <tr>
                        <td className="py-3 px-4 font-bold text-gray-900">Clause 44</td>
                        <td className="py-3 px-4">
                          <div>Break-up of Total Expenditure under GST</div>
                          <div className="text-[11px] text-gray-500">Exempt vs Composition vs Registered vs Unregistered</div>
                        </td>
                        <td className="py-3 px-4 text-right font-mono font-bold text-gray-700">
                          ₹{summary.totalUnregisteredExpenditure.toLocaleString('en-IN')}
                        </td>
                        <td className="py-3 px-4 text-center">
                          <span className={`px-2 py-0.5 text-[11px] font-semibold rounded ${summary.unregisteredExpenditureRatio > 40 ? 'bg-amber-50 text-amber-700 border border-amber-200' : 'bg-gray-50 text-gray-800 border border-gray-200'}`}>
                            {summary.unregisteredExpenditureRatio}% Unregistered
                          </span>
                        </td>
                        <td className="py-3 px-4 text-gray-600">
                          Mandatory CBDT 7-column table ready for Annexure 44 filing.
                        </td>
                      </tr>

                      <tr>
                        <td className="py-3 px-4 font-bold text-gray-900">Clause 21(a)</td>
                        <td className="py-3 px-4">
                          <div>Section 37(1) Explanation 1</div>
                          <div className="text-[11px] text-gray-500">Penalties, personal expenses, capital items debited to P&L</div>
                        </td>
                        <td className="py-3 px-4 text-right font-mono font-bold text-rose-600">
                          ₹{summary.clause21aPersonalAndPenalties.toLocaleString('en-IN')}
                        </td>
                        <td className="py-3 px-4 text-center">
                          {summary.clause21aPersonalAndPenalties > 0 ? (
                            <span className="px-2 py-0.5 text-[11px] font-semibold rounded bg-rose-50 text-rose-700 border border-rose-200">
                              Inadmissible
                            </span>
                          ) : (
                            <span className="px-2 py-0.5 text-[11px] font-semibold rounded bg-gray-50 text-gray-800 border border-gray-200">
                              Clean
                            </span>
                          )}
                        </td>
                        <td className="py-3 px-4 text-gray-600">
                          Disclose penalty & personal expenditures itemized in Clause 21(a).
                        </td>
                      </tr>

                      <tr>
                        <td className="py-3 px-4 font-bold text-gray-900">Clause 22</td>
                        <td className="py-3 px-4">
                          <div>Section 23 of MSMED Act, 2006</div>
                          <div className="text-[11px] text-gray-500">Interest paid/payable on delayed payments to Micro/Small enterprises</div>
                        </td>
                        <td className="py-3 px-4 text-right font-mono font-bold text-rose-600">
                          ₹{summary.clause22MsmeInterest.toLocaleString('en-IN')}
                        </td>
                        <td className="py-3 px-4 text-center">
                          {summary.clause22MsmeInterest > 0 ? (
                            <span className="px-2 py-0.5 text-[11px] font-semibold rounded bg-rose-50 text-rose-700 border border-rose-200">
                              Inadmissible
                            </span>
                          ) : (
                            <span className="px-2 py-0.5 text-[11px] font-semibold rounded bg-gray-50 text-gray-800 border border-gray-200">
                              Nil Interest
                            </span>
                          )}
                        </td>
                        <td className="py-3 px-4 text-gray-600">
                          Report disallowable MSME interest under Section 23 of MSMED Act.
                        </td>
                      </tr>

                      <tr className="bg-gray-50 font-bold">
                        <td colSpan={2} className="py-3 px-4 text-gray-900 text-right uppercase">
                          Total Recommended Additions to Taxable Profit (Form 3CD):
                        </td>
                        <td className="py-3 px-4 text-right font-mono text-base text-rose-700">
                          ₹{summary.totalTaxAdditions.toLocaleString('en-IN')}
                        </td>
                        <td colSpan={2} className="py-3 px-4 text-xs text-gray-500">
                          Net taxable income adjustment before MAT / Chapter VI-A deductions
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* TAB 2: CLAUSE 21(d) / SEC 40A(3) CASH VIOLATIONS */}
          {activeTab === 'clause40a3' && (
            <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-xs space-y-4">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div>
                  <h3 className="text-base font-bold text-gray-900">
                    Clause 21(d) / Section 40A(3) Cash Payment Scrutiny
                  </h3>
                  <p className="text-xs text-gray-500 mt-0.5">
                    Scrutinizes daily aggregate cash payments to a person exceeding ₹10,000 (or ₹35,000 for goods carriages). Toggle Rule 6DD exceptions if valid statutory grounds exist.
                  </p>
                </div>
                <div className="text-xs font-semibold bg-gray-100 text-gray-800 px-3 py-1.5 rounded-lg border border-gray-200">
                  Total Disallowance: <span className="font-mono text-rose-600 font-bold">₹{summary.clause40a3TotalDisallowable.toLocaleString('en-IN')}</span>
                </div>
              </div>

              {summary.cashViolations.length === 0 ? (
                <div className="p-8 text-center text-gray-500 border border-dashed border-gray-200 rounded-xl">
                  <CheckCircle2 className="w-8 h-8 text-gray-600 mx-auto mb-2" />
                  <p className="text-sm font-semibold text-gray-800">No Section 40A(3) Cash Violations Detected</p>
                  <p className="text-xs text-gray-500 mt-0.5">All daily cash payments to individual payees remain strictly within statutory thresholds.</p>
                </div>
              ) : (
                <div className="overflow-x-auto">
                  <table className="w-full text-xs text-left">
                    <thead className="bg-gray-50 border-y border-gray-200 text-gray-600 uppercase font-semibold">
                      <tr>
                        <th className="py-2.5 px-3">Date</th>
                        <th className="py-2.5 px-3">Payee / Beneficiary Name</th>
                        <th className="py-2.5 px-3">Vouchers Involved</th>
                        <th className="py-2.5 px-3 text-right">Daily Aggregate</th>
                        <th className="py-2.5 px-3 text-right">Limit</th>
                        <th className="py-2.5 px-3 text-center">Rule 6DD Exception</th>
                        <th className="py-2.5 px-3 text-right">Disallowable Addition</th>
                        <th className="py-2.5 px-3">Statutory Remarks</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-100 text-gray-700">
                      {summary.cashViolations.map((cv) => (
                        <tr key={cv.id} className={cv.isRule6DdExempt ? 'bg-gray-50/50' : 'hover:bg-rose-50/20'}>
                          <td className="py-3 px-3 font-medium text-gray-900">{cv.date}</td>
                          <td className="py-3 px-3 font-semibold text-gray-900">{cv.payeeName}</td>
                          <td className="py-3 px-3 font-mono text-[11px] text-gray-600">{cv.voucherNumbers.join(', ')}</td>
                          <td className="py-3 px-3 text-right font-mono font-bold text-gray-900">
                            ₹{cv.dailyAggregateAmount.toLocaleString('en-IN')}
                          </td>
                          <td className="py-3 px-3 text-right font-mono text-gray-500">
                            ₹{cv.thresholdLimit.toLocaleString('en-IN')}
                          </td>
                          <td className="py-3 px-3 text-center">
                            <button
                              onClick={() => toggleRule6Dd(cv.id)}
                              className={`px-2.5 py-1 rounded text-[11px] font-semibold border cursor-pointer transition-colors ${
                                cv.isRule6DdExempt
                                  ? 'bg-gray-50 text-gray-900 border-gray-200 hover:bg-gray-100'
                                  : 'bg-gray-100 text-gray-700 border-gray-300 hover:bg-gray-200'
                              }`}
                            >
                              {cv.isRule6DdExempt ? '✓ Rule 6DD Exempt' : 'Apply Rule 6DD'}
                            </button>
                          </td>
                          <td className="py-3 px-3 text-right font-mono font-bold text-rose-600">
                            ₹{cv.disallowableAmount.toLocaleString('en-IN')}
                          </td>
                          <td className="py-3 px-3 text-gray-600 text-[11px]">
                            {cv.exemptionReason}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          )}

          {/* TAB 3: CLAUSE 44 BREAKDOWN OF TOTAL EXPENDITURE */}
          {activeTab === 'clause44' && (
            <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-xs space-y-4">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div>
                  <h3 className="text-base font-bold text-gray-900">
                    Clause 44: Break-up of Total Expenditure (Mandatory CBDT Format)
                  </h3>
                  <p className="text-xs text-gray-500 mt-0.5">
                    Exact statutory 7-column table reporting procurement from GST Registered entities (Exempt, Composition, Regular) vs Unregistered suppliers.
                  </p>
                </div>
                <div className="text-xs text-gray-600 bg-gray-50 border border-gray-200 px-3 py-1.5 rounded-lg">
                  Registered: <strong className="font-mono text-gray-900">₹{summary.totalRegisteredExpenditure.toLocaleString('en-IN')}</strong> | Unregistered: <strong className="font-mono text-rose-600">₹{summary.totalUnregisteredExpenditure.toLocaleString('en-IN')}</strong>
                </div>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-xs text-left">
                  <thead className="bg-gray-50 border-y border-gray-200 text-gray-600 uppercase font-semibold">
                    <tr>
                      <th className="py-2.5 px-3">S.No</th>
                      <th className="py-2.5 px-3">Head of Expenditure</th>
                      <th className="py-2.5 px-3 text-right">Total Expenditure</th>
                      <th className="py-2.5 px-3 text-right">GST Exempt</th>
                      <th className="py-2.5 px-3 text-right">Composition (Sec 10)</th>
                      <th className="py-2.5 px-3 text-right">Other Registered</th>
                      <th className="py-2.5 px-3 text-right">Total Registered</th>
                      <th className="py-2.5 px-3 text-right">Unregistered Spend</th>
                      <th className="py-2.5 px-3 text-right">% Unreg</th>
                      <th className="py-2.5 px-3">Audit Scrutiny Remarks</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100 text-gray-700">
                    {summary.clause44Breakdown.map((r) => (
                      <tr key={r.sNo} className="hover:bg-gray-50">
                        <td className="py-3 px-3 text-gray-500">{r.sNo}</td>
                        <td className="py-3 px-3 font-semibold text-gray-900">{r.headOfExpenditure}</td>
                        <td className="py-3 px-3 text-right font-mono font-bold text-gray-900">
                          ₹{r.totalExpenditure.toLocaleString('en-IN')}
                        </td>
                        <td className="py-3 px-3 text-right font-mono text-gray-600">
                          ₹{r.exemptSupplies.toLocaleString('en-IN')}
                        </td>
                        <td className="py-3 px-3 text-right font-mono text-gray-600">
                          ₹{r.compositionEntities.toLocaleString('en-IN')}
                        </td>
                        <td className="py-3 px-3 text-right font-mono text-gray-600">
                          ₹{r.otherRegisteredEntities.toLocaleString('en-IN')}
                        </td>
                        <td className="py-3 px-3 text-right font-mono font-semibold text-gray-800">
                          ₹{r.totalRegisteredExpenditure.toLocaleString('en-IN')}
                        </td>
                        <td className="py-3 px-3 text-right font-mono font-semibold text-rose-600">
                          ₹{r.unregisteredEntities.toLocaleString('en-IN')}
                        </td>
                        <td className="py-3 px-3 text-right font-mono font-bold">
                          {r.unregisteredPercent}%
                        </td>
                        <td className="py-3 px-3 text-[11px] text-gray-500">
                          {r.remarks}
                        </td>
                      </tr>
                    ))}
                    <tr className="bg-gray-50 font-bold">
                      <td colSpan={2} className="py-3 px-3 text-gray-900 uppercase">Total Clause 44 Spend:</td>
                      <td className="py-3 px-3 text-right font-mono text-gray-900">
                        ₹{summary.totalClause44Expenditure.toLocaleString('en-IN')}
                      </td>
                      <td colSpan={3} className="py-3 px-3"></td>
                      <td className="py-3 px-3 text-right font-mono text-gray-800">
                        ₹{summary.totalRegisteredExpenditure.toLocaleString('en-IN')}
                      </td>
                      <td className="py-3 px-3 text-right font-mono text-rose-600">
                        ₹{summary.totalUnregisteredExpenditure.toLocaleString('en-IN')}
                      </td>
                      <td className="py-3 px-3 text-right font-mono">
                        {summary.unregisteredExpenditureRatio}%
                      </td>
                      <td className="py-3 px-3 text-xs text-gray-500">
                        CBDT Compliant
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* TAB 4: CLAUSE 34 CHAPTER XVII-B TDS */}
          {activeTab === 'clause34' && (
            <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-xs space-y-4">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div>
                  <h3 className="text-base font-bold text-gray-900">
                    Clause 34(a): Compliance with Chapter XVII-B (TDS/TCS)
                  </h3>
                  <p className="text-xs text-gray-500 mt-0.5">
                    Section-wise audit of expenditure liable to tax deduction at source, short deductions, and Section 40(a)(ia) disallowance computation.
                  </p>
                </div>
                <div className="text-xs font-semibold bg-gray-100 text-gray-800 px-3 py-1.5 rounded-lg border border-gray-200">
                  Sec 40(a)(ia) Addition: <span className="font-mono text-rose-600 font-bold">₹{summary.clause34Disallowance40aIa.toLocaleString('en-IN')}</span>
                </div>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-xs text-left">
                  <thead className="bg-gray-50 border-y border-gray-200 text-gray-600 uppercase font-semibold">
                    <tr>
                      <th className="py-2.5 px-3">Section</th>
                      <th className="py-2.5 px-3">Nature of Payment</th>
                      <th className="py-2.5 px-3 text-right">Covered Spend</th>
                      <th className="py-2.5 px-3 text-right">Liable to TDS</th>
                      <th className="py-2.5 px-3 text-right">Tax Required</th>
                      <th className="py-2.5 px-3 text-right">Tax Deducted</th>
                      <th className="py-2.5 px-3 text-right">Short Deduction</th>
                      <th className="py-2.5 px-3 text-right">Sec 40(a)(ia) Disallowance</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100 text-gray-700">
                    {summary.clause34Summary.map((s) => (
                      <tr key={s.section} className="hover:bg-gray-50">
                        <td className="py-3 px-3 font-bold text-gray-900">Sec {s.section}</td>
                        <td className="py-3 px-3 font-medium text-gray-800">{s.sectionName}</td>
                        <td className="py-3 px-3 text-right font-mono text-gray-900">
                          ₹{s.totalExpenditureCovered.toLocaleString('en-IN')}
                        </td>
                        <td className="py-3 px-3 text-right font-mono text-gray-900">
                          ₹{s.amountEligibleForTds.toLocaleString('en-IN')}
                        </td>
                        <td className="py-3 px-3 text-right font-mono text-gray-900 font-semibold">
                          ₹{s.tdsAmountRequired.toLocaleString('en-IN')}
                        </td>
                        <td className="py-3 px-3 text-right font-mono text-gray-800 font-semibold">
                          ₹{s.tdsAmountDeducted.toLocaleString('en-IN')}
                        </td>
                        <td className="py-3 px-3 text-right font-mono text-rose-600 font-bold">
                          ₹{s.shortDeductionAmount.toLocaleString('en-IN')}
                        </td>
                        <td className="py-3 px-3 text-right font-mono text-rose-700 font-bold">
                          ₹{s.disallowance40aIa.toLocaleString('en-IN')}
                        </td>
                      </tr>
                    ))}
                    <tr className="bg-gray-50 font-bold">
                      <td colSpan={2} className="py-3 px-3 text-gray-900 uppercase">Chapter XVII-B Totals:</td>
                      <td className="py-3 px-3 text-right font-mono text-gray-900">
                        ₹{summary.clause34TotalCoveredExpenditure.toLocaleString('en-IN')}
                      </td>
                      <td className="py-3 px-3"></td>
                      <td className="py-3 px-3 text-right font-mono text-gray-900">
                        ₹{summary.clause34TotalTdsRequired.toLocaleString('en-IN')}
                      </td>
                      <td className="py-3 px-3 text-right font-mono text-gray-800">
                        ₹{summary.clause34TotalTdsDeducted.toLocaleString('en-IN')}
                      </td>
                      <td className="py-3 px-3 text-right font-mono text-rose-600">
                        ₹{summary.clause34ShortDeduction.toLocaleString('en-IN')}
                      </td>
                      <td className="py-3 px-3 text-right font-mono text-rose-700">
                        ₹{summary.clause34Disallowance40aIa.toLocaleString('en-IN')}
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* TAB 5: CLAUSE 21(a) & 22 PENALTIES & MSME INTEREST */}
          {activeTab === 'clause21a' && (
            <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-xs space-y-4">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div>
                  <h3 className="text-base font-bold text-gray-900">
                    Clause 21(a) Penalties & Clause 22 MSME Interest Disallowances
                  </h3>
                  <p className="text-xs text-gray-500 mt-0.5">
                    Scans narrations and accounts for fines, compounding fees, director personal expenses, and MSMED Act Section 23 interest.
                  </p>
                </div>
                <div className="text-xs font-semibold bg-gray-100 text-gray-800 px-3 py-1.5 rounded-lg border border-gray-200">
                  Total Disallowable: <span className="font-mono text-rose-600 font-bold">₹{(summary.clause21aPersonalAndPenalties + summary.clause22MsmeInterest).toLocaleString('en-IN')}</span>
                </div>
              </div>

              {summary.inadmissibleExpenses.length === 0 ? (
                <div className="p-8 text-center text-gray-500 border border-dashed border-gray-200 rounded-xl">
                  <CheckCircle2 className="w-8 h-8 text-gray-600 mx-auto mb-2" />
                  <p className="text-sm font-semibold text-gray-800">No Inadmissible Penalties or Personal Expenses Flagged</p>
                  <p className="text-xs text-gray-500 mt-0.5">Vouchers appear clean of statutory penalty and fine keywords.</p>
                </div>
              ) : (
                <div className="overflow-x-auto">
                  <table className="w-full text-xs text-left">
                    <thead className="bg-gray-50 border-y border-gray-200 text-gray-600 uppercase font-semibold">
                      <tr>
                        <th className="py-2.5 px-3">Date</th>
                        <th className="py-2.5 px-3">Voucher No</th>
                        <th className="py-2.5 px-3">Clause Reference</th>
                        <th className="py-2.5 px-3">Disallowance Category</th>
                        <th className="py-2.5 px-3">Account Head</th>
                        <th className="py-2.5 px-3">Narration & Detected Triggers</th>
                        <th className="py-2.5 px-3 text-right">Inadmissible Amount</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-100 text-gray-700">
                      {summary.inadmissibleExpenses.map((inad) => (
                        <tr key={inad.id} className="hover:bg-rose-50/20">
                          <td className="py-3 px-3 font-medium text-gray-900">{inad.date}</td>
                          <td className="py-3 px-3 font-mono text-gray-600">{inad.voucherNo}</td>
                          <td className="py-3 px-3 font-bold text-gray-900">{inad.clauseReference}</td>
                          <td className="py-3 px-3">
                            <span className="px-2 py-0.5 rounded text-[11px] font-semibold bg-rose-50 text-rose-700 border border-rose-200">
                              {inad.disallowanceType}
                            </span>
                          </td>
                          <td className="py-3 px-3 font-semibold text-gray-900">{inad.accountName}</td>
                          <td className="py-3 px-3 text-gray-600 text-[11px]">
                            <div>{inad.narration}</div>
                            <div className="text-rose-600 font-semibold mt-0.5">Keywords: {inad.flaggedKeywords.join(', ')}</div>
                          </td>
                          <td className="py-3 px-3 text-right font-mono font-bold text-rose-600">
                            ₹{inad.amount.toLocaleString('en-IN')}
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
      )}
    </div>
  );
};