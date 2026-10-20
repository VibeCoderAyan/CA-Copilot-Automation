import * as XLSX from 'xlsx';
import { parseAmount, isValidPan } from './dataSanitizer';
import { createStyledWorksheet } from './excelStyler';

export interface SalaryIssue {
  id: string;
  empId: string;
  rowIndex: number;
  category:
    | 'duplicate_id'
    | 'missing_details'
    | 'negative_salary'
    | 'pf_mismatch'
    | 'esi_mismatch'
    | 'invalid_pan'
    | 'excessive_deduction'
    | 'zero_basic';
  severity: 'critical' | 'warning' | 'info';
  title: string;
  description: string;
  suggestedAction: string;
}

export interface EmployeeSalaryRow {
  id: string;
  rowIndex: number;
  empId: string;
  empName: string;
  department: string;
  pan: string;
  bankAccount: string;

  // Earnings
  basic: number;
  hra: number;
  da: number;
  conveyance: number;
  specialAllowance: number;
  bonus: number;
  otherEarnings: number;
  grossSalary: number;

  // Deductions
  pfEmployee: number;
  esiEmployee: number;
  professionalTax: number;
  tds: number;
  loanDeduction: number;
  otherDeductions: number;
  totalDeductions: number;

  // Net Pay
  netSalary: number;

  // Statutory Calculations (Reference)
  calculatedPf: number;
  calculatedEsi: number;
  pfEmployer: number;
  esiEmployer: number;
  isPfMismatch: boolean;
  isEsiMismatch: boolean;

  issues: SalaryIssue[];
  raw: Record<string, any>;
}

export interface PayrollConfig {
  pfRateEmployee: number; // default 12%
  pfRateEmployer: number; // default 12%
  pfWageCeiling: number; // default 15000
  enforcePfCeiling: boolean; // default true
  esiRateEmployee: number; // default 0.75%
  esiRateEmployer: number; // default 3.25%
  esiWageCeiling: number; // default 21000
  ptDefaultAmount: number; // default 200
  recalculateDeductions: boolean; // if true, applies formula values
}

export interface PayrollSummary {
  totalEmployees: number;
  totalGrossSalary: number;
  totalDeductions: number;
  totalNetSalary: number;
  totalPfEmployee: number;
  totalPfEmployer: number;
  totalPfLiability: number;
  totalEsiEmployee: number;
  totalEsiEmployer: number;
  totalEsiLiability: number;
  totalTdsDeducted: number;
  totalPt: number;
  totalCtc: number; // Gross + Employer PF + Employer ESI
  criticalIssuesCount: number;
  warningIssuesCount: number;
  issues: SalaryIssue[];
}

export interface PayrollColumnMapping {
  empIdCol: string;
  empNameCol: string;
  departmentCol: string;
  panCol: string;
  bankAccountCol: string;
  basicCol: string;
  hraCol: string;
  daCol: string;
  conveyanceCol: string;
  specialCol: string;
  bonusCol: string;
  grossCol: string;
  pfCol: string;
  esiCol: string;
  ptCol: string;
  tdsCol: string;
  loanCol: string;
  totalDeductionsCol: string;
  netCol: string;
}

/**
 * Auto-detects payroll columns from spreadsheet headers
 */
import { detectColumnCategory } from './columnDetector';

export function detectPayrollColumns(headers: string[]): PayrollColumnMapping {
  const mapping: PayrollColumnMapping = {
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
  };

  for (const h of headers) {
    const clean = h.toLowerCase().trim();
    const cat = detectColumnCategory(h);

    if (!mapping.empIdCol && (cat === 'accountCode' || /emp.*id|employee.*id|emp.*code|staff.*id|badge/i.test(clean))) {
      mapping.empIdCol = h;
    } else if (!mapping.empNameCol && (cat === 'partyName' || /emp.*name|employee.*name|name|staff.*name/i.test(clean))) {
      mapping.empNameCol = h;
    } else if (!mapping.departmentCol && /dept|department|designation|role/i.test(clean)) {
      mapping.departmentCol = h;
    } else if (!mapping.panCol && (cat === 'gstin' || /pan|pan.*no|income.*tax.*id/i.test(clean))) {
      mapping.panCol = h;
    } else if (!mapping.bankAccountCol && /bank|a\/c|account.*no|bank.*acc/i.test(clean)) {
      mapping.bankAccountCol = h;
    } else if (!mapping.basicCol && /basic|basic.*salary|basic.*pay/i.test(clean)) {
      mapping.basicCol = h;
    } else if (!mapping.hraCol && /hra|house.*rent/i.test(clean)) {
      mapping.hraCol = h;
    } else if (!mapping.daCol && /\bda\b|dearness/i.test(clean)) {
      mapping.daCol = h;
    } else if (!mapping.conveyanceCol && /conveyance|transport|travel/i.test(clean)) {
      mapping.conveyanceCol = h;
    } else if (!mapping.specialCol && /special|special.*allowance|other.*allowance/i.test(clean)) {
      mapping.specialCol = h;
    } else if (!mapping.bonusCol && /bonus|incentive|overtime|ot/i.test(clean)) {
      mapping.bonusCol = h;
    } else if (!mapping.grossCol && (cat === 'totalAmount' || /gross|gross.*salary|total.*earnings|gross.*pay/i.test(clean))) {
      mapping.grossCol = h;
    } else if (!mapping.pfCol && /\bpf\b|provident|epf/i.test(clean) && !/employer/i.test(clean)) {
      mapping.pfCol = h;
    } else if (!mapping.esiCol && /\besi\b|esic/i.test(clean) && !/employer/i.test(clean)) {
      mapping.esiCol = h;
    } else if (!mapping.ptCol && /\bpt\b|prof.*tax|professional.*tax/i.test(clean)) {
      mapping.ptCol = h;
    } else if (!mapping.tdsCol && /tds|tax.*deducted|income.*tax/i.test(clean)) {
      mapping.tdsCol = h;
    } else if (!mapping.loanCol && /loan|advance|salary.*adv/i.test(clean)) {
      mapping.loanCol = h;
    } else if (!mapping.totalDeductionsCol && /total.*deduct|deductions/i.test(clean)) {
      mapping.totalDeductionsCol = h;
    } else if (!mapping.netCol && (cat === 'balance' || /net|net.*salary|net.*pay|take.*home/i.test(clean))) {
      mapping.netCol = h;
    }
  }

  return mapping;
}

/**
 * Main Payroll Calculation & Audit Engine
 */
export function calculatePayroll(
  rows: Record<string, any>[],
  mapping: PayrollColumnMapping,
  config: PayrollConfig
): { summary: PayrollSummary; employees: EmployeeSalaryRow[] } {
  const employees: EmployeeSalaryRow[] = [];
  const allIssues: SalaryIssue[] = [];

  // Track duplicate Employee IDs
  const empIdCounts = new Map<string, number>();
  rows.forEach((r) => {
    const empId = String(r[mapping.empIdCol] || '').trim();
    if (empId) {
      empIdCounts.set(empId, (empIdCounts.get(empId) || 0) + 1);
    }
  });

  let totalGross = 0;
  let totalDeductions = 0;
  let totalNet = 0;
  let totalPfEmployee = 0;
  let totalPfEmployer = 0;
  let totalEsiEmployee = 0;
  let totalEsiEmployer = 0;
  let totalTds = 0;
  let totalPt = 0;

  rows.forEach((r, idx) => {
    const rowIndex = idx + 2;
    const empId = String(r[mapping.empIdCol] || '').trim() || `EMP-${idx + 1}`;
    const empName = String(r[mapping.empNameCol] || '').trim();
    const department = String(r[mapping.departmentCol] || '').trim();
    const pan = String(r[mapping.panCol] || '').trim().toUpperCase();
    const bankAccount = String(r[mapping.bankAccountCol] || '').trim();

    const basic = parseAmount(r[mapping.basicCol]);
    const hra = parseAmount(r[mapping.hraCol]);
    const da = parseAmount(r[mapping.daCol]);
    const conveyance = parseAmount(r[mapping.conveyanceCol]);
    const specialAllowance = parseAmount(r[mapping.specialCol]);
    const bonus = parseAmount(r[mapping.bonusCol]);

    // Computed or given gross
    let bookGross = parseAmount(r[mapping.grossCol]);
    const computedEarnings = basic + hra + da + conveyance + specialAllowance + bonus;
    const grossSalary = bookGross > 0 ? bookGross : computedEarnings;

    // Deductions from file
    let bookPf = parseAmount(r[mapping.pfCol]);
    let bookEsi = parseAmount(r[mapping.esiCol]);
    let pt = parseAmount(r[mapping.ptCol]);
    if (pt === 0 && config.ptDefaultAmount > 0 && grossSalary > 15000) {
      pt = config.ptDefaultAmount;
    }
    const tds = parseAmount(r[mapping.tdsCol]);
    const loanDeduction = parseAmount(r[mapping.loanCol]);

    // Statutory Calculations
    // PF Base: Basic + DA
    const pfWageBase = basic + da;
    const pfEligibleBase = config.enforcePfCeiling
      ? Math.min(pfWageBase, config.pfWageCeiling)
      : pfWageBase;
    const calculatedPf = Math.round(pfEligibleBase * (config.pfRateEmployee / 100));
    const pfEmployer = Math.round(pfEligibleBase * (config.pfRateEmployer / 100));

    // ESI Base: Gross salary (excluding bonus/annual incentives)
    const esiWageBase = grossSalary - bonus;
    let calculatedEsi = 0;
    let esiEmployer = 0;
    if (esiWageBase <= config.esiWageCeiling && esiWageBase > 0) {
      calculatedEsi = Math.ceil(esiWageBase * (config.esiRateEmployee / 100));
      esiEmployer = Math.ceil(esiWageBase * (config.esiRateEmployer / 100));
    }

    const pfEmployee = config.recalculateDeductions ? calculatedPf : bookPf > 0 ? bookPf : calculatedPf;
    const esiEmployee = config.recalculateDeductions ? calculatedEsi : bookEsi > 0 ? bookEsi : calculatedEsi;

    const isPfMismatch = bookPf > 0 && Math.abs(bookPf - calculatedPf) > 10;
    const isEsiMismatch = bookEsi > 0 && Math.abs(bookEsi - calculatedEsi) > 10;

    let bookTotalDeductions = parseAmount(r[mapping.totalDeductionsCol]);
    const computedDeductions = pfEmployee + esiEmployee + pt + tds + loanDeduction;
    const totalDeductionsEmp = bookTotalDeductions > 0 && !config.recalculateDeductions
      ? bookTotalDeductions
      : computedDeductions;

    let bookNet = parseAmount(r[mapping.netCol]);
    const netSalary = bookNet > 0 && !config.recalculateDeductions
      ? bookNet
      : Math.max(0, grossSalary - totalDeductionsEmp);

    const empIssues: SalaryIssue[] = [];

    // Validation 1: Missing Employee Name
    if (!empName) {
      const iss: SalaryIssue = {
        id: `iss-pr-${allIssues.length + 1}`,
        empId,
        rowIndex,
        category: 'missing_details',
        severity: 'critical',
        title: 'Missing Employee Name',
        description: `Row ${rowIndex} with ID "${empId}" lacks an employee name.`,
        suggestedAction: 'Update employee record with legal name as per Aadhaar/PAN.',
      };
      empIssues.push(iss);
      allIssues.push(iss);
    }

    // Validation 2: Duplicate Employee ID
    if (empIdCounts.get(empId)! > 1) {
      const iss: SalaryIssue = {
        id: `iss-pr-${allIssues.length + 1}`,
        empId,
        rowIndex,
        category: 'duplicate_id',
        severity: 'critical',
        title: `Duplicate Employee ID: ${empId}`,
        description: `Employee ID "${empId}" appears across ${empIdCounts.get(empId)} rows.`,
        suggestedAction: 'Ensure unique employee IDs are assigned across departments.',
      };
      empIssues.push(iss);
      allIssues.push(iss);
    }

    // Validation 3: Zero or Negative Basic Pay
    if (basic <= 0) {
      const iss: SalaryIssue = {
        id: `iss-pr-${allIssues.length + 1}`,
        empId,
        rowIndex,
        category: 'zero_basic',
        severity: 'critical',
        title: 'Zero or Negative Basic Salary',
        description: `Basic pay of ₹${basic} recorded for ${empName || empId}.`,
        suggestedAction: 'Verify salary structure against employment agreement.',
      };
      empIssues.push(iss);
      allIssues.push(iss);
    }

    // Validation 4: Negative Net Pay
    if (grossSalary < totalDeductionsEmp) {
      const iss: SalaryIssue = {
        id: `iss-pr-${allIssues.length + 1}`,
        empId,
        rowIndex,
        category: 'negative_salary',
        severity: 'critical',
        title: 'Deductions Exceed Gross Salary',
        description: `Total deductions (₹${totalDeductionsEmp.toLocaleString('en-IN')}) exceed gross earnings (₹${grossSalary.toLocaleString('en-IN')}).`,
        suggestedAction: 'Cap loan or voluntary deductions so net salary does not turn negative.',
      };
      empIssues.push(iss);
      allIssues.push(iss);
    }

    // Validation 5: Invalid PAN Format
    if (pan && !isValidPan(pan)) {
      const iss: SalaryIssue = {
        id: `iss-pr-${allIssues.length + 1}`,
        empId,
        rowIndex,
        category: 'invalid_pan',
        severity: 'warning',
        title: `Invalid PAN Format (${pan})`,
        description: `PAN entered does not match standard 10-character alphanumeric structure.`,
        suggestedAction: 'Obtain copy of PAN card to avoid higher TDS withholding.',
      };
      empIssues.push(iss);
      allIssues.push(iss);
    }

    // Validation 6: Statutory PF Mismatch
    if (isPfMismatch) {
      const iss: SalaryIssue = {
        id: `iss-pr-${allIssues.length + 1}`,
        empId,
        rowIndex,
        category: 'pf_mismatch',
        severity: 'warning',
        title: `PF Deduction Variance (Book: ₹${bookPf} vs Statutory: ₹${calculatedPf})`,
        description: `Deducted PF differs from statutory 12% calculation on basic pay ₹${basic.toLocaleString('en-IN')}.`,
        suggestedAction: 'Confirm whether voluntary PF (VPF) or wage ceiling was configured differently.',
      };
      empIssues.push(iss);
      allIssues.push(iss);
    }

    // Validation 7: ESI Mismatch
    if (isEsiMismatch) {
      const iss: SalaryIssue = {
        id: `iss-pr-${allIssues.length + 1}`,
        empId,
        rowIndex,
        category: 'esi_mismatch',
        severity: 'warning',
        title: `ESI Deduction Variance (Book: ₹${bookEsi} vs Statutory: ₹${calculatedEsi})`,
        description: `Deducted ESI differs from 0.75% calculated rate on eligible gross.`,
        suggestedAction: 'Verify whether employee gross exceeds ESI wage ceiling ₹21,000.',
      };
      empIssues.push(iss);
      allIssues.push(iss);
    }

    totalGross += grossSalary;
    totalDeductions += totalDeductionsEmp;
    totalNet += netSalary;
    totalPfEmployee += pfEmployee;
    totalPfEmployer += pfEmployer;
    totalEsiEmployee += esiEmployee;
    totalEsiEmployer += esiEmployer;
    totalTds += tds;
    totalPt += pt;

    employees.push({
      id: `emp-sal-${idx + 1}`,
      rowIndex,
      empId,
      empName: empName || `Employee ${idx + 1}`,
      department: department || 'General',
      pan,
      bankAccount,
      basic,
      hra,
      da,
      conveyance,
      specialAllowance,
      bonus,
      otherEarnings: 0,
      grossSalary,
      pfEmployee,
      esiEmployee,
      professionalTax: pt,
      tds,
      loanDeduction,
      otherDeductions: 0,
      totalDeductions: totalDeductionsEmp,
      netSalary,
      calculatedPf,
      calculatedEsi,
      pfEmployer,
      esiEmployer,
      isPfMismatch,
      isEsiMismatch,
      issues: empIssues,
      raw: r,
    });
  });

  const criticalIssuesCount = allIssues.filter((i) => i.severity === 'critical').length;
  const warningIssuesCount = allIssues.filter((i) => i.severity === 'warning').length;

  return {
    summary: {
      totalEmployees: employees.length,
      totalGrossSalary: totalGross,
      totalDeductions,
      totalNetSalary: totalNet,
      totalPfEmployee,
      totalPfEmployer,
      totalPfLiability: totalPfEmployee + totalPfEmployer,
      totalEsiEmployee,
      totalEsiEmployer,
      totalEsiLiability: totalEsiEmployee + totalEsiEmployer,
      totalTdsDeducted: totalTds,
      totalPt,
      totalCtc: totalGross + totalPfEmployer + totalEsiEmployer,
      criticalIssuesCount,
      warningIssuesCount,
      issues: allIssues,
    },
    employees,
  };
}

/**
 * Exports complete multi-sheet Excel salary register
 */
export function exportPayrollWorkbook(
  summary: PayrollSummary,
  employees: EmployeeSalaryRow[],
  clientName = 'Corporate Client',
  monthYear = 'September 2026'
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Executive Payroll Summary
  const summaryRows = [
    { Parameter: 'Company / Employer Name', Value: clientName },
    { Parameter: 'Salary Period', Value: monthYear },
    { Parameter: 'Total Headcount (Employees)', Value: summary.totalEmployees },
    { Parameter: 'Total Gross Earnings (₹)', Value: summary.totalGrossSalary },
    { Parameter: 'Total Employee Deductions (₹)', Value: summary.totalDeductions },
    { Parameter: 'Total Net Salary Disbursed (₹)', Value: summary.totalNetSalary },
    { Parameter: 'Total Employee PF Withheld (₹)', Value: summary.totalPfEmployee },
    { Parameter: 'Total Employer PF Contribution (₹)', Value: summary.totalPfEmployer },
    { Parameter: 'Total PF Remittance Liability (₹)', Value: summary.totalPfLiability },
    { Parameter: 'Total Employee ESI Withheld (₹)', Value: summary.totalEsiEmployee },
    { Parameter: 'Total Employer ESI Contribution (₹)', Value: summary.totalEsiEmployer },
    { Parameter: 'Total ESI Remittance Liability (₹)', Value: summary.totalEsiLiability },
    { Parameter: 'Total TDS u/s 192 (₹)', Value: summary.totalTdsDeducted },
    { Parameter: 'Total Professional Tax (₹)', Value: summary.totalPt },
    { Parameter: 'Total Cost to Company (CTC) (₹)', Value: summary.totalCtc },
    { Parameter: 'Critical Audit Exceptions', Value: summary.criticalIssuesCount },
    { Parameter: 'Statutory Verification Note', Value: 'Rules configurable. Values subject to company HR policy verification.' },
  ];
  const wsSummary = createStyledWorksheet(summaryRows, {
    currencyColumns: ['Value'],
    columnWidths: { Parameter: 38, Value: 32 },
  });
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Payroll_Summary');

  // Sheet 2: Employee Salary Register
  const registerRows = employees.map((e) => ({
    'Emp ID': e.empId,
    'Employee Name': e.empName,
    'Department': e.department,
    'PAN': e.pan,
    'Bank A/c': e.bankAccount,
    'Basic Pay (₹)': e.basic,
    'HRA (₹)': e.hra,
    'DA (₹)': e.da,
    'Conveyance (₹)': e.conveyance,
    'Special Allowance (₹)': e.specialAllowance,
    'Bonus / OT (₹)': e.bonus,
    'Gross Earnings (₹)': e.grossSalary,
    'PF (Employee) (₹)': e.pfEmployee,
    'ESI (Employee) (₹)': e.esiEmployee,
    'Prof Tax (PT) (₹)': e.professionalTax,
    'TDS (₹)': e.tds,
    'Loan / Adv (₹)': e.loanDeduction,
    'Total Deductions (₹)': e.totalDeductions,
    'Net Salary Payable (₹)': e.netSalary,
    'PF (Employer) (₹)': e.pfEmployer,
    'ESI (Employer) (₹)': e.esiEmployer,
    'Audit Status': e.issues.length > 0 ? `${e.issues.length} ISSUE(S)` : 'VERIFIED CLEAN',
  }));
  const wsRegister = createStyledWorksheet(registerRows, {
    currencyColumns: [
      'Basic Pay (₹)',
      'HRA (₹)',
      'DA (₹)',
      'Conveyance (₹)',
      'Special Allowance (₹)',
      'Bonus / OT (₹)',
      'Gross Earnings (₹)',
      'PF (Employee) (₹)',
      'ESI (Employee) (₹)',
      'Prof Tax (PT) (₹)',
      'TDS (₹)',
      'Loan / Adv (₹)',
      'Total Deductions (₹)',
      'Net Salary Payable (₹)',
      'PF (Employer) (₹)',
      'ESI (Employer) (₹)',
    ],
  });
  XLSX.utils.book_append_sheet(wb, wsRegister, 'Salary_Register');

  // Sheet 3: Audit Flags & Discrepancies
  const issueRows = summary.issues.map((iss, idx) => ({
    'Issue #': idx + 1,
    'Severity': iss.severity.toUpperCase(),
    'Category': iss.category.replace(/_/g, ' ').toUpperCase(),
    'Emp ID': iss.empId,
    'Row #': iss.rowIndex,
    'Audit Finding': iss.title,
    'Detailed Observation': iss.description,
    'Suggested Corrective Action': iss.suggestedAction,
  }));
  const wsIssues = createStyledWorksheet(issueRows, {
    columnWidths: {
      'Issue #': 10,
      'Severity': 14,
      'Category': 22,
      'Emp ID': 14,
      'Row #': 10,
      'Audit Finding': 35,
      'Detailed Observation': 50,
      'Suggested Corrective Action': 45,
    },
  });
  XLSX.utils.book_append_sheet(wb, wsIssues, 'Statutory_Audit_Flags');

  return wb;
}

/**
 * Realistic Indian Payroll Sample Dataset
 */
export function getSamplePayrollRows(): Record<string, any>[] {
  return [
    {
      'Emp ID': 'EMP-101',
      'Employee Name': 'Aarav Sharma',
      'Department': 'Finance & Accounts',
      'PAN': 'ABCPS1234A',
      'Bank Account': '9100100412891',
      'Basic Salary': 45000,
      'HRA': 18000,
      'DA': 4500,
      'Conveyance': 1600,
      'Special Allowance': 10900,
      'Bonus': 0,
      'Gross Salary': 80000,
      'PF Employee': 1800, // Capped at 15,000 ceiling
      'ESI Employee': 0, // Exceeds 21,000 ceiling
      'Professional Tax': 200,
      'TDS u/s 192': 3500,
      'Loan Deduction': 0,
      'Total Deductions': 5500,
      'Net Salary': 74500,
    },
    {
      'Emp ID': 'EMP-102',
      'Employee Name': 'Priya Deshmukh',
      'Department': 'Internal Audit',
      'PAN': 'ABCPD5678B',
      'Bank Account': '5010023910245',
      'Basic Salary': 38000,
      'HRA': 15200,
      'DA': 3800,
      'Conveyance': 1600,
      'Special Allowance': 6400,
      'Bonus': 0,
      'Gross Salary': 65000,
      'PF Employee': 1800,
      'ESI Employee': 0,
      'Professional Tax': 200,
      'TDS u/s 192': 2100,
      'Loan Deduction': 0,
      'Total Deductions': 4100,
      'Net Salary': 60900,
    },
    {
      'Emp ID': 'EMP-103',
      'Employee Name': 'Rohan Kulkarni',
      'Department': 'Plant Operations',
      'PAN': 'ABCPK9012C',
      'Bank Account': '3029104819023',
      'Basic Salary': 12000,
      'HRA': 4000,
      'DA': 1200,
      'Conveyance': 1200,
      'Special Allowance': 0,
      'Bonus': 1500,
      'Gross Salary': 19900,
      'PF Employee': 1584, // 12% of (12000 + 1200)
      'ESI Employee': 138, // 0.75% of (19900 - 1500 = 18400)
      'Professional Tax': 200,
      'TDS u/s 192': 0,
      'Loan Deduction': 500,
      'Total Deductions': 2422,
      'Net Salary': 17478,
    },
    {
      'Emp ID': 'EMP-104',
      'Employee Name': 'Neha Nair',
      'Department': 'Human Resources',
      'PAN': 'INVALID_PAN_123', // INTENTIONAL PAN VALIDATION ERROR
      'Bank Account': '4029102839101',
      'Basic Salary': 28000,
      'HRA': 11200,
      'DA': 2800,
      'Conveyance': 1600,
      'Special Allowance': 4400,
      'Bonus': 0,
      'Gross Salary': 48000,
      'PF Employee': 1800,
      'ESI Employee': 0,
      'Professional Tax': 200,
      'TDS u/s 192': 1200,
      'Loan Deduction': 0,
      'Total Deductions': 3200,
      'Net Salary': 44800,
    },
    {
      'Emp ID': 'EMP-104', // INTENTIONAL DUPLICATE EMPLOYEE ID
      'Employee Name': 'Vikram Patel',
      'Department': 'Logistics',
      'PAN': 'ABCPP3456D',
      'Bank Account': '7029104819201',
      'Basic Salary': 16000,
      'HRA': 5000,
      'DA': 1600,
      'Conveyance': 1200,
      'Special Allowance': 1200,
      'Bonus': 0,
      'Gross Salary': 25000,
      'PF Employee': 1800,
      'ESI Employee': 0,
      'Professional Tax': 200,
      'TDS u/s 192': 0,
      'Loan Deduction': 1000,
      'Total Deductions': 3000,
      'Net Salary': 22000,
    },
    {
      'Emp ID': 'EMP-106',
      'Employee Name': 'Suresh Verma',
      'Department': 'Maintenance',
      'PAN': 'ABCPV7890E',
      'Bank Account': '8019284910293',
      'Basic Salary': 10000,
      'HRA': 3000,
      'DA': 1000,
      'Conveyance': 1000,
      'Special Allowance': 0,
      'Bonus': 0,
      'Gross Salary': 15000,
      'PF Employee': 800, // ERROR: Book PF ₹800 vs calculated ₹1,320 (12% of 11,000)!
      'ESI Employee': 113, // 0.75% of 15,000
      'Professional Tax': 175,
      'TDS u/s 192': 0,
      'Loan Deduction': 14000, // ERROR: High loan deduction causes net salary to drop near zero!
      'Total Deductions': 15088,
      'Net Salary': -88, // NEGATIVE NET SALARY EXCEPTION!
    },
    {
      'Emp ID': 'EMP-107',
      'Employee Name': 'Ananya Iyer',
      'Department': 'Marketing',
      'PAN': 'ABCPI1122F',
      'Bank Account': '6019284910291',
      'Basic Salary': 52000,
      'HRA': 20800,
      'DA': 5200,
      'Conveyance': 1600,
      'Special Allowance': 15400,
      'Bonus': 10000,
      'Gross Salary': 105000,
      'PF Employee': 1800,
      'ESI Employee': 0,
      'Professional Tax': 200,
      'TDS u/s 192': 8200,
      'Loan Deduction': 0,
      'Total Deductions': 10200,
      'Net Salary': 94800,
    },
    {
      'Emp ID': 'EMP-108',
      'Employee Name': 'Manish Gupta',
      'Department': 'Production',
      'PAN': 'ABCPG3344G',
      'Bank Account': '9019284910294',
      'Basic Salary': 14000,
      'HRA': 4500,
      'DA': 1400,
      'Conveyance': 1200,
      'Special Allowance': 900,
      'Bonus': 0,
      'Gross Salary': 22000,
      'PF Employee': 1800,
      'ESI Employee': 0, // Exceeds 21,000 ceiling
      'Professional Tax': 200,
      'TDS u/s 192': 0,
      'Loan Deduction': 0,
      'Total Deductions': 2000,
      'Net Salary': 20000,
    },
  ];
}
