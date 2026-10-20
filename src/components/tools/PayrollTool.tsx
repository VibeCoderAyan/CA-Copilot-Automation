import { ToolHeader } from '../common/ToolHeader';
import React, { useState, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  UploadCloud,
  Download,
  Printer,
  FileSpreadsheet,
  AlertTriangle,
  AlertCircle,
  CheckCircle2,
  Check,
  X,
  Search,
  Sliders,
  ShieldCheck,
  UserCheck,
  Eye,
  RefreshCw,
  Building,
} from 'lucide-react';
import { InnerToolBadge } from '../common/InnerToolBadge';
import {
  type EmployeeSalaryRow,
  type PayrollConfig,
  type PayrollSummary,
  type PayrollColumnMapping,
  detectPayrollColumns,
  calculatePayroll,
  exportPayrollWorkbook,
} from '../../utils/payrollEngine';

interface PayrollToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

export const PayrollTool: React.FC<PayrollToolProps> = ({
  onBack,
  onShowToast,
}) => {
  const [rawRows, setRawRows] = useState<Record<string, any>[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [mapping, setMapping] = useState<PayrollColumnMapping>({
    empIdCol: '',
    empNameCol: '',
    departmentCol: '',
    panCol: '',
    bankAccountCol: '',
    basicCol: '',
    hraCol: '',
    daCol: '',
    conveyanceCol: '',
    specialCol: '',
    bonusCol: '',
    grossCol: '',
    pfCol: '',
    esiCol: '',
    ptCol: '',
    tdsCol: '',
    loanCol: '',
    totalDeductionsCol: '',
    netCol: '',
  });
  const [fileName, setFileName] = useState('');
  const [clientName, setClientName] = useState('');
  const [monthYear, setMonthYear] = useState('September 2026');
  const [activeTab, setActiveTab] = useState<'register' | 'payslip_view' | 'exceptions'>('register');
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedEmpId, setSelectedEmpId] = useState<string>('');
  const [showConfigPanel, setShowConfigPanel] = useState(false);
  const [showMappingPanel, setShowMappingPanel] = useState(false);

  // Configurable statutory rules
  const [config, setConfig] = useState<PayrollConfig>({
    pfRateEmployee: 12.0,
    pfRateEmployer: 12.0,
    pfWageCeiling: 15000,
    enforcePfCeiling: true,
    esiRateEmployee: 0.75,
    esiRateEmployer: 3.25,
    esiWageCeiling: 21000,
    ptDefaultAmount: 200,
    recalculateDeductions: false,
  });

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Run calculation engine
  const { summary, employees } = useMemo(() => {
    if (rawRows.length === 0) {
      return {
        summary: {
          totalEmployees: 0,
          totalGrossSalary: 0,
          totalDeductions: 0,
          totalNetSalary: 0,
          totalPfEmployee: 0,
          totalPfEmployer: 0,
          totalPfLiability: 0,
          totalEsiEmployee: 0,
          totalEsiEmployer: 0,
          totalEsiLiability: 0,
          totalTdsDeducted: 0,
          totalPt: 0,
          totalCtc: 0,
          criticalIssuesCount: 0,
          warningIssuesCount: 0,
          issues: [],
        },
        employees: [],
      };
    }
    return calculatePayroll(rawRows, mapping, config);
  }, [rawRows, mapping, config]);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });
        const firstSheet = wb.Sheets[wb.SheetNames[0]];
        const json: Record<string, any>[] = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });

        if (json.length === 0) {
          onShowToast('Empty File', 'The uploaded salary sheet has no data rows.', 'error');
          return;
        }

        const detectedHeaders = Object.keys(json[0]);
        setHeaders(detectedHeaders);
        const detectedMapping = detectPayrollColumns(detectedHeaders);
        setMapping(detectedMapping);
        setRawRows(json);
        if (json.length > 0) {
          const firstId = String(json[0][detectedMapping.empIdCol] || 'EMP-101');
          setSelectedEmpId(firstId);
        }
      } catch (err: any) {
        onShowToast('Parsing Error', err.message || 'Could not parse salary file.', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  
  const handleExportExcel = () => {
    if (employees.length === 0) return;
    try {
      const wb = exportPayrollWorkbook(summary, employees, clientName, monthYear);
      XLSX.writeFile(wb, `Salary_Register_${monthYear.replace(/\s+/g, '_')}.xlsx`);
      onShowToast('Salary Register Exported', 'Downloaded 3-sheet salary register and statutory compliance report.', 'success');
    } catch (err: any) {
      onShowToast('Export Failed', err.message || 'Could not export Excel workbook.', 'error');
    }
  };

  const handlePrintPayslip = () => {
    window.print();
  };

  const filteredEmployees = useMemo(() => {
    return employees.filter((e) => {
      if (!searchQuery) return true;
      const q = searchQuery.toLowerCase();
      return (
        e.empName.toLowerCase().includes(q) ||
        e.empId.toLowerCase().includes(q) ||
        e.department.toLowerCase().includes(q) ||
        e.pan.toLowerCase().includes(q)
      );
    });
  }, [employees, searchQuery]);

  const selectedEmployee = useMemo(() => {
    if (!selectedEmpId && employees.length > 0) return employees[0];
    return employees.find((e) => e.empId === selectedEmpId) || employees[0];
  }, [employees, selectedEmpId]);

  const formatInr = (val: number) => {
    return val.toLocaleString('en-IN', { style: 'currency', currency: 'INR', maximumFractionDigits: 2 });
  };

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="payroll-calculator"
        title="Payroll & Salary Calculator"
        description="Calculate gross-to-net salary, validate PF/ESI/PT/TDS statutory rules, identify negative pay exceptions, and generate payslips."
        onBack={onBack}
        actionButtons={
          <div className="flex items-center gap-2 flex-wrap">
            <button
              onClick={() => setShowConfigPanel(!showConfigPanel)}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold border transition-all cursor-pointer ${
                showConfigPanel
                  ? 'bg-gray-900 text-white border-gray-900'
                  : 'bg-white text-gray-700 border-gray-200 hover:bg-gray-50'
              }`}
            >
              <Sliders className="w-3.5 h-3.5" />
              <span>Statutory Rules</span>
            </button>

            {rawRows.length > 0 && (
              <button
                onClick={() => setShowMappingPanel(!showMappingPanel)}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-white text-gray-700 border border-gray-200 hover:bg-gray-50 transition-all cursor-pointer"
              >
                <span>Columns</span>
              </button>
            )}

            <button
              onClick={handleExportExcel}
              disabled={employees.length === 0}
              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-semibold bg-gray-900 text-white hover:bg-gray-900 disabled:opacity-40 disabled:cursor-not-allowed shadow-xs transition-all cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Export Salary Register</span>
            </button>
          </div>
        }
      />

        {/* Statutory Rules Config Panel */}
        {showConfigPanel && (
          <div className="mt-4 pt-4 border-t border-gray-100 bg-[#F9FAFA] -mx-5 -mb-5 p-5 rounded-b-xl space-y-4">
            <span className="text-xs font-bold text-gray-900 block">
              Configurable Statutory Parameters & Deduction Rules
            </span>
            <div className="grid grid-cols-1 sm:grid-cols-4 gap-4 text-xs">
              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">PF Employee Rate (%)</label>
                <input
                  type="number"
                  value={config.pfRateEmployee}
                  onChange={(e) => setConfig({ ...config, pfRateEmployee: Number(e.target.value) })}
                  className="w-full px-2.5 py-1.5 bg-white border border-gray-200 rounded text-xs font-semibold"
                />
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">PF Statutory Wage Ceiling (₹)</label>
                <input
                  type="number"
                  value={config.pfWageCeiling}
                  onChange={(e) => setConfig({ ...config, pfWageCeiling: Number(e.target.value) })}
                  className="w-full px-2.5 py-1.5 bg-white border border-gray-200 rounded text-xs font-semibold"
                />
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">ESI Employee Rate (%)</label>
                <input
                  type="number"
                  step="0.05"
                  value={config.esiRateEmployee}
                  onChange={(e) => setConfig({ ...config, esiRateEmployee: Number(e.target.value) })}
                  className="w-full px-2.5 py-1.5 bg-white border border-gray-200 rounded text-xs font-semibold"
                />
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">ESI Wage Ceiling (₹)</label>
                <input
                  type="number"
                  value={config.esiWageCeiling}
                  onChange={(e) => setConfig({ ...config, esiWageCeiling: Number(e.target.value) })}
                  className="w-full px-2.5 py-1.5 bg-white border border-gray-200 rounded text-xs font-semibold"
                />
              </div>

              <div className="col-span-2 flex items-center gap-4 pt-1">
                <label className="flex items-center gap-2 cursor-pointer text-xs text-gray-800">
                  <input
                    type="checkbox"
                    checked={config.enforcePfCeiling}
                    onChange={(e) => setConfig({ ...config, enforcePfCeiling: e.target.checked })}
                    className="rounded text-gray-700"
                  />
                  <span>Cap PF at ₹15,000 ceiling (₹1,800/mo)</span>
                </label>

                <label className="flex items-center gap-2 cursor-pointer text-xs text-gray-800">
                  <input
                    type="checkbox"
                    checked={config.recalculateDeductions}
                    onChange={(e) => setConfig({ ...config, recalculateDeductions: e.target.checked })}
                    className="rounded text-gray-700"
                  />
                  <span>Recompute statutory deductions automatically</span>
                </label>
              </div>
            </div>
          </div>
        )}

        {/* Column Mapping Panel */}
        {showMappingPanel && headers.length > 0 && (
          <div className="mt-4 pt-4 border-t border-gray-100 bg-white p-4 rounded-lg border border-gray-200 space-y-3">
            <span className="text-xs font-bold text-gray-900 block">
              Salary Sheet Column Mapping
            </span>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3 text-xs">
              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Employee ID</label>
                <select
                  value={mapping.empIdCol}
                  onChange={(e) => setMapping({ ...mapping, empIdCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs"
                >
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Employee Name</label>
                <select
                  value={mapping.empNameCol}
                  onChange={(e) => setMapping({ ...mapping, empNameCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs"
                >
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Basic Salary</label>
                <select
                  value={mapping.basicCol}
                  onChange={(e) => setMapping({ ...mapping, basicCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs"
                >
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-gray-600 mb-1">Gross Salary</label>
                <select
                  value={mapping.grossCol}
                  onChange={(e) => setMapping({ ...mapping, grossCol: e.target.value })}
                  className="w-full px-2 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs"
                >
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
            </div>
          </div>
        )}

      {/* 2. File Import Dropzone */}
      {rawRows.length === 0 ? (
        <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-8 sm:p-10 text-center space-y-4 shadow-2xs transition-colors">
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileUpload}
            accept=".xlsx,.xls,.csv"
            className="hidden"
            id="payroll-salary-upload"
          />
          <label htmlFor="payroll-salary-upload" className="cursor-pointer flex flex-col items-center">
            <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mx-auto mb-3 shadow-2xs">
              <UploadCloud className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-semibold text-gray-900">
              Import Employee Salary Register
            </h3>
            <p className="text-xs text-gray-500 mt-1 max-w-md mx-auto leading-relaxed">
              Supports Excel (.xlsx, .xls) and CSV payroll sheets with basic, allowances, deductions, and statutory PF/ESI components.
            </p>
            <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
              <UploadCloud className="w-4 h-4" />
              <span>Select Salary Sheet</span>
            </div>
            <span className="text-[11px] text-gray-400 font-medium mt-3 block">
              100% Offline Processing · Rules configurable · No sensitive payroll data uploaded
            </span>
          </label>
        </div>
      ) : (
        /* 3. Executive Scrutiny KPI Metric Cards */
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 no-print">
          <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs">
            <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider block">
              Total Headcount
            </span>
            <div className="flex items-baseline justify-between mt-1.5">
              <span className="text-xl font-bold text-gray-900">
                {summary.totalEmployees}
              </span>
              <span className="text-[11px] text-gray-500">
                Employees
              </span>
            </div>
          </div>

          <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs">
            <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider block">
              Total Gross Earnings
            </span>
            <div className="flex items-baseline justify-between mt-1.5">
              <span className="text-xl font-bold text-gray-900">
                {formatInr(summary.totalGrossSalary)}
              </span>
              <span className="text-[11px] text-gray-500">
                Deductions: {formatInr(summary.totalDeductions)}
              </span>
            </div>
          </div>

          <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs">
            <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider block">
              Net Payable Disbursal
            </span>
            <div className="flex items-baseline justify-between mt-1.5">
              <span className="text-xl font-bold text-gray-800">
                {formatInr(summary.totalNetSalary)}
              </span>
              <span className="text-[11px] text-gray-800 font-semibold">
                Take-Home Pay
              </span>
            </div>
          </div>

          <div className={`p-4 rounded-xl border shadow-xs ${
            summary.criticalIssuesCount > 0 ? 'bg-red-50/20 border-red-200' : 'bg-white border-gray-200'
          }`}>
            <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider block">
              Statutory Liabilities & Flags
            </span>
            <div className="flex items-baseline justify-between mt-1.5">
              <span className="text-xl font-bold text-gray-900">
                {formatInr(summary.totalPfLiability + summary.totalEsiLiability + summary.totalTdsDeducted)}
              </span>
              <span className={`text-[11px] font-semibold ${summary.criticalIssuesCount > 0 ? 'text-red-700' : 'text-gray-500'}`}>
                {summary.criticalIssuesCount} Exceptions
              </span>
            </div>
          </div>
        </div>
      )}

      {/* 4. Active Workbench */}
      {employees.length > 0 && (
        <div className="bg-white rounded-xl border border-gray-200 shadow-xs overflow-hidden">
          {/* Tabs and Search Header */}
          <div className="p-4 border-b border-gray-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3 no-print">
            <div className="flex items-center gap-1.5 overflow-x-auto pb-1 sm:pb-0">
              {[
                { id: 'register' as const, label: 'Salary Register', count: employees.length },
                { id: 'payslip_view' as const, label: 'Individual Payslips', count: employees.length },
                { id: 'exceptions' as const, label: 'Audit Exceptions', count: summary.issues.length },
              ].map((tab) => (
                <button
                  key={tab.id}
                  onClick={() => setActiveTab(tab.id)}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all cursor-pointer shrink-0 ${
                    activeTab === tab.id
                      ? 'bg-gray-900 text-white shadow-2xs'
                      : 'text-gray-600 hover:bg-gray-100 hover:text-gray-900'
                  }`}
                >
                  {tab.label} ({tab.count})
                </button>
              ))}
            </div>

            <div className="relative w-full sm:w-64">
              <Search className="w-3.5 h-3.5 text-gray-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Search employee, ID, PAN..."
                className="w-full pl-9 pr-7 py-1.5 text-xs bg-gray-50 border border-gray-200 rounded-lg focus:outline-none focus:border-gray-900"
              />
            </div>
          </div>

          {/* Tab 1: Columnar Salary Register Table */}
          {activeTab === 'register' && (
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs text-gray-800">
                <thead className="bg-[#F8F9FA] border-b border-gray-200 text-gray-600 font-semibold uppercase text-[10px] tracking-wider">
                  <tr>
                    <th className="px-4 py-3">Emp ID</th>
                    <th className="px-4 py-3">Employee Name</th>
                    <th className="px-4 py-3">Department</th>
                    <th className="px-4 py-3">PAN</th>
                    <th className="px-4 py-3 text-right">Basic (₹)</th>
                    <th className="px-4 py-3 text-right">Gross (₹)</th>
                    <th className="px-4 py-3 text-right">PF (₹)</th>
                    <th className="px-4 py-3 text-right">ESI (₹)</th>
                    <th className="px-4 py-3 text-right">TDS (₹)</th>
                    <th className="px-4 py-3 text-right">Net Payable (₹)</th>
                    <th className="px-4 py-3 text-center">Action</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {filteredEmployees.map((e) => (
                    <tr key={e.id} className="hover:bg-gray-50/80 transition-colors">
                      <td className="px-4 py-2.5 font-mono font-bold text-gray-900">{e.empId}</td>
                      <td className="px-4 py-2.5 font-medium text-gray-900">{e.empName}</td>
                      <td className="px-4 py-2.5 text-gray-600">{e.department}</td>
                      <td className="px-4 py-2.5 font-mono text-[11px] text-gray-500">{e.pan || 'N/A'}</td>
                      <td className="px-4 py-2.5 text-right font-mono tabular-nums">{e.basic.toLocaleString('en-IN')}</td>
                      <td className="px-4 py-2.5 text-right font-mono font-bold tabular-nums text-gray-900">{e.grossSalary.toLocaleString('en-IN')}</td>
                      <td className="px-4 py-2.5 text-right font-mono tabular-nums">{e.pfEmployee > 0 ? e.pfEmployee.toLocaleString('en-IN') : '-'}</td>
                      <td className="px-4 py-2.5 text-right font-mono tabular-nums">{e.esiEmployee > 0 ? e.esiEmployee.toLocaleString('en-IN') : '-'}</td>
                      <td className="px-4 py-2.5 text-right font-mono tabular-nums">{e.tds > 0 ? e.tds.toLocaleString('en-IN') : '-'}</td>
                      <td className={`px-4 py-2.5 text-right font-mono font-bold tabular-nums ${e.netSalary < 0 ? 'text-red-700' : 'text-gray-800'}`}>
                        {e.netSalary.toLocaleString('en-IN')}
                      </td>
                      <td className="px-4 py-2.5 text-center">
                        <button
                          onClick={() => {
                            setSelectedEmpId(e.empId);
                            setActiveTab('payslip_view');
                          }}
                          className="px-2 py-1 text-[11px] font-semibold bg-gray-100 hover:bg-gray-200 text-gray-700 rounded transition-colors"
                        >
                          Payslip
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {/* Tab 2: Individual Employee Payslip View */}
          {activeTab === 'payslip_view' && selectedEmployee && (
            <div className="p-6 space-y-6 max-w-3xl mx-auto">
              <div className="flex items-center justify-between no-print pb-3 border-b border-gray-200">
                <div className="flex items-center gap-3">
                  <label className="text-xs font-semibold text-gray-700">Select Employee:</label>
                  <select
                    value={selectedEmployee.empId}
                    onChange={(e) => setSelectedEmpId(e.target.value)}
                    className="px-3 py-1.5 text-xs bg-gray-50 border border-gray-200 rounded-lg font-semibold"
                  >
                    {employees.map((e) => (
                      <option key={e.id} value={e.empId}>
                        {e.empId} - {e.empName} ({e.department})
                      </option>
                    ))}
                  </select>
                </div>

                <button
                  onClick={handlePrintPayslip}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-gray-900 text-white rounded-lg text-xs font-semibold hover:bg-black transition-colors"
                >
                  <Printer className="w-3.5 h-3.5" />
                  <span>Print Payslip</span>
                </button>
              </div>

              {/* Formal Printable Payslip Card */}
              <div className="bg-white border-2 border-gray-300 rounded-xl p-6 shadow-sm space-y-4 print:border-none print:shadow-none print:p-0">
                <div className="text-center pb-4 border-b-2 border-gray-200">
                  <h2 className="text-base font-bold text-gray-900 uppercase tracking-tight">{clientName}</h2>
                  <p className="text-xs text-gray-500 mt-0.5">Salary Slip for the month of {monthYear}</p>
                </div>

                <div className="grid grid-cols-2 gap-4 text-xs py-2 border-b border-gray-200">
                  <div className="space-y-1">
                    <div><span className="text-gray-500">Employee ID:</span> <span className="font-bold text-gray-900">{selectedEmployee.empId}</span></div>
                    <div><span className="text-gray-500">Name:</span> <span className="font-bold text-gray-900">{selectedEmployee.empName}</span></div>
                    <div><span className="text-gray-500">Department:</span> <span className="font-semibold text-gray-800">{selectedEmployee.department}</span></div>
                  </div>
                  <div className="space-y-1 text-right">
                    <div><span className="text-gray-500">PAN:</span> <span className="font-mono font-bold text-gray-900">{selectedEmployee.pan || 'N/A'}</span></div>
                    <div><span className="text-gray-500">Bank A/c:</span> <span className="font-mono text-gray-800">{selectedEmployee.bankAccount || 'On File'}</span></div>
                    <div><span className="text-gray-500">Pay Period:</span> <span className="text-gray-800">{monthYear}</span></div>
                  </div>
                </div>

                {/* Earnings & Deductions Dual Table */}
                <div className="grid grid-cols-2 gap-4 text-xs">
                  {/* Earnings */}
                  <div className="border border-gray-200 rounded-lg overflow-hidden">
                    <div className="bg-[#F8F9FA] px-3 py-2 font-bold text-gray-900 border-b border-gray-200 flex justify-between">
                      <span>Earnings</span>
                      <span>Amount (₹)</span>
                    </div>
                    <div className="p-3 space-y-1.5">
                      <div className="flex justify-between"><span>Basic Salary</span><span className="font-mono">{selectedEmployee.basic.toLocaleString('en-IN')}</span></div>
                      <div className="flex justify-between"><span>House Rent Allowance (HRA)</span><span className="font-mono">{selectedEmployee.hra.toLocaleString('en-IN')}</span></div>
                      <div className="flex justify-between"><span>Dearness Allowance (DA)</span><span className="font-mono">{selectedEmployee.da.toLocaleString('en-IN')}</span></div>
                      <div className="flex justify-between"><span>Conveyance Allowance</span><span className="font-mono">{selectedEmployee.conveyance.toLocaleString('en-IN')}</span></div>
                      <div className="flex justify-between"><span>Special Allowance</span><span className="font-mono">{selectedEmployee.specialAllowance.toLocaleString('en-IN')}</span></div>
                      {selectedEmployee.bonus > 0 && <div className="flex justify-between"><span>Bonus / Incentive</span><span className="font-mono">{selectedEmployee.bonus.toLocaleString('en-IN')}</span></div>}
                    </div>
                    <div className="bg-gray-50 px-3 py-2 font-bold text-gray-900 border-t border-gray-200 flex justify-between">
                      <span>Gross Earnings</span>
                      <span className="font-mono">₹{selectedEmployee.grossSalary.toLocaleString('en-IN')}</span>
                    </div>
                  </div>

                  {/* Deductions */}
                  <div className="border border-gray-200 rounded-lg overflow-hidden">
                    <div className="bg-[#F8F9FA] px-3 py-2 font-bold text-gray-900 border-b border-gray-200 flex justify-between">
                      <span>Deductions</span>
                      <span>Amount (₹)</span>
                    </div>
                    <div className="p-3 space-y-1.5">
                      <div className="flex justify-between"><span>Provident Fund (EPF)</span><span className="font-mono">{selectedEmployee.pfEmployee.toLocaleString('en-IN')}</span></div>
                      <div className="flex justify-between"><span>Employee State Insurance (ESI)</span><span className="font-mono">{selectedEmployee.esiEmployee.toLocaleString('en-IN')}</span></div>
                      <div className="flex justify-between"><span>Professional Tax (PT)</span><span className="font-mono">{selectedEmployee.professionalTax.toLocaleString('en-IN')}</span></div>
                      <div className="flex justify-between"><span>Income Tax (TDS u/s 192)</span><span className="font-mono">{selectedEmployee.tds.toLocaleString('en-IN')}</span></div>
                      {selectedEmployee.loanDeduction > 0 && <div className="flex justify-between"><span>Salary Advance / Loan</span><span className="font-mono">{selectedEmployee.loanDeduction.toLocaleString('en-IN')}</span></div>}
                    </div>
                    <div className="bg-gray-50 px-3 py-2 font-bold text-gray-900 border-t border-gray-200 flex justify-between">
                      <span>Total Deductions</span>
                      <span className="font-mono">₹{selectedEmployee.totalDeductions.toLocaleString('en-IN')}</span>
                    </div>
                  </div>
                </div>

                {/* Net Pay Highlight */}
                <div className="p-3.5 bg-gray-50/60 border border-gray-200 rounded-lg flex items-center justify-between text-xs">
                  <div>
                    <span className="text-gray-900 font-bold block">NET PAYABLE AMOUNT</span>
                    <span className="text-[11px] text-gray-800">Credited to registered bank account</span>
                  </div>
                  <span className="text-base font-bold text-gray-900 font-mono">
                    ₹{selectedEmployee.netSalary.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                  </span>
                </div>
              </div>
            </div>
          )}

          {/* Tab 3: Audit Exceptions */}
          {activeTab === 'exceptions' && (
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs text-gray-800">
                <thead className="bg-[#F8F9FA] border-b border-gray-200 text-gray-600 font-semibold uppercase text-[10px] tracking-wider">
                  <tr>
                    <th className="px-4 py-3">Severity</th>
                    <th className="px-4 py-3">Emp ID</th>
                    <th className="px-4 py-3">Row #</th>
                    <th className="px-4 py-3">Finding</th>
                    <th className="px-4 py-3">Observation</th>
                    <th className="px-4 py-3">Suggested Action</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {summary.issues.map((iss) => (
                    <tr key={iss.id} className="hover:bg-gray-50/80 transition-colors">
                      <td className="px-4 py-3">
                        <span
                          className={`inline-block px-2 py-0.5 rounded text-[10px] font-bold tracking-wider uppercase ${
                            iss.severity === 'critical'
                              ? 'bg-red-600 text-white'
                              : 'bg-amber-100 text-amber-800 border border-amber-200'
                          }`}
                        >
                          {iss.severity}
                        </span>
                      </td>
                      <td className="px-4 py-3 font-mono font-bold text-gray-900">{iss.empId}</td>
                      <td className="px-4 py-3 font-mono text-gray-500">Row {iss.rowIndex}</td>
                      <td className="px-4 py-3 font-semibold text-gray-900">{iss.title}</td>
                      <td className="px-4 py-3 text-[11px] text-gray-600 max-w-sm">{iss.description}</td>
                      <td className="px-4 py-3 text-[11px] text-gray-600 max-w-xs">{iss.suggestedAction}</td>
                    </tr>
                  ))}
                  {summary.issues.length === 0 && (
                    <tr>
                      <td colSpan={6} className="py-12 text-center text-xs text-gray-500">
                        <CheckCircle2 className="w-8 h-8 text-gray-600 mx-auto mb-2" />
                        <p className="font-semibold text-gray-800">All payroll records verified clean.</p>
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}
    </div>
  );
};