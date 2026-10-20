import * as XLSX from 'xlsx';
import { createStyledWorksheet } from './excelStyler';

export type ClosingTaskStatus = 'not_started' | 'in_progress' | 'blocked' | 'completed';
export type ClosingCategory =
  | 'Cash & Bank'
  | 'Receivables & Payables'
  | 'General Ledger & Journals'
  | 'Fixed Assets'
  | 'Accruals & Cut-off'
  | 'Taxation & Statutory'
  | 'Financial Reporting';

export interface ClosingTask {
  id: string;
  category: ClosingCategory;
  title: string;
  description: string;
  assignee: string;
  dueDate: string;
  completedDate?: string;
  status: ClosingTaskStatus;
  priority: 'low' | 'medium' | 'high' | 'critical';
  notes: string;
  supportingDocRef?: string;
  isOverdue?: boolean;
  history: {
    timestamp: string;
    action: string;
    changedBy: string;
  }[];
}

export interface MonthEndClosingState {
  companyName: string;
  accountingPeriod: string; // e.g. "September 2026"
  financialYear: string; // e.g. "FY 2026-27"
  targetClosingDate: string;
  tasks: ClosingTask[];
}

export interface ClosingMetrics {
  totalTasks: number;
  completedCount: number;
  inProgressCount: number;
  blockedCount: number;
  notStartedCount: number;
  overdueCount: number;
  completionPercentage: number;
  outstandingTasks: ClosingTask[];
  blockedTasks: ClosingTask[];
}

export const STANDARD_CLOSING_TEMPLATE_TASKS: Omit<ClosingTask, 'id' | 'isOverdue' | 'history'>[] = [
  {
    category: 'Cash & Bank',
    title: 'Bank Reconciliation (BRS)',
    description: 'Reconcile all operational bank accounts with general ledger cashbook balances and trace unpresented cheques.',
    assignee: '',
    dueDate: '04-10-2026',
    status: 'not_started',
    priority: 'critical',
    notes: '',
  },
  {
    category: 'Cash & Bank',
    title: 'Petty Cash Physical Verification & Vouching',
    description: 'Verify physical cash in hand with petty cashbook ledger. Ensure no single cash payment exceeds ₹10,000 (Sec 40A(3)).',
    assignee: '',
    dueDate: '02-10-2026',
    status: 'not_started',
    priority: 'high',
    notes: '',
  },
  {
    category: 'Receivables & Payables',
    title: 'Trade Debtors & Aging Scrutiny',
    description: 'Review debtor aging buckets (> 180 days, > 3 years) and identify doubtful debts requiring ECL provisioning.',
    assignee: '',
    dueDate: '05-10-2026',
    status: 'not_started',
    priority: 'high',
    notes: '',
  },
  {
    category: 'Receivables & Payables',
    title: 'Trade Creditors & MSMED 45-Day Compliance',
    description: 'Cross-check supplier invoices against MSME registrations. Flag payments pending beyond 45 days (Sec 43B(h)).',
    assignee: '',
    dueDate: '06-10-2026',
    status: 'not_started',
    priority: 'critical',
    notes: '',
  },
  {
    category: 'Accruals & Cut-off',
    title: 'Expense Accruals & Prepaid Expenses',
    description: 'Book electricity, telephone, rent, and legal fee provisions. Amortize monthly prepaid insurance.',
    assignee: '',
    dueDate: '05-10-2026',
    status: 'not_started',
    priority: 'medium',
    notes: '',
  },
  {
    category: 'Accruals & Cut-off',
    title: 'Revenue & Purchase Cut-Off Testing',
    description: 'Examine last 5 dispatches before month-end and first 5 dispatches of next month to ensure correct revenue recognition.',
    assignee: '',
    dueDate: '06-10-2026',
    status: 'not_started',
    priority: 'high',
    notes: '',
  },
  {
    category: 'Fixed Assets',
    title: 'Depreciation Booking & Capex Capitalization',
    description: 'Recompute Companies Act Schedule II depreciation and capitalize completed CWIP asset additions.',
    assignee: '',
    dueDate: '07-10-2026',
    status: 'not_started',
    priority: 'high',
    notes: '',
  },
  {
    category: 'General Ledger & Journals',
    title: 'Intercompany & Branch Reconciliation',
    description: 'Reconcile balances between Head Office, manufacturing plants, and subsidiary entities. Verify contra balances.',
    assignee: '',
    dueDate: '07-10-2026',
    status: 'not_started',
    priority: 'medium',
    notes: '',
  },
  {
    category: 'Taxation & Statutory',
    title: 'GSTR-2B vs Books Reconciliation & ITC Claim',
    description: 'Download GSTR-2B from GST portal, compare with purchase register, and compute eligible vs blocked ITC (Sec 17(5)).',
    assignee: '',
    dueDate: '10-10-2026',
    status: 'not_started',
    priority: 'critical',
    notes: '',
  },
  {
    category: 'Taxation & Statutory',
    title: 'TDS Deduction & Challan Deposit Check',
    description: 'Verify TDS deductions under 194C, 194J, 194I and prepare challan payment before statutory 7th due date.',
    assignee: '',
    dueDate: '07-10-2026',
    status: 'not_started',
    priority: 'critical',
    notes: '',
  },
  {
    category: 'General Ledger & Journals',
    title: 'Trial Balance Scrutiny & Debit/Credit Tie-Out',
    description: 'Review complete trial balance for negative cash balances, suspense accounts, and misclassified capital items.',
    assignee: '',
    dueDate: '08-10-2026',
    status: 'not_started',
    priority: 'high',
    notes: '',
  },
  {
    category: 'Financial Reporting',
    title: 'Monthly Flash P&L & Balance Sheet Review',
    description: 'Generate monthly provisional management accounts with variance analysis against approved annual budgets.',
    assignee: '',
    dueDate: '10-10-2026',
    status: 'not_started',
    priority: 'high',
    notes: '',
  },
];

/**
 * Calculates metrics and flags overdue/blocked tasks based on current date
 */
export function evaluateClosingMetrics(tasks: ClosingTask[]): ClosingMetrics {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let completedCount = 0;
  let inProgressCount = 0;
  let blockedCount = 0;
  let notStartedCount = 0;
  let overdueCount = 0;

  const outstandingTasks: ClosingTask[] = [];
  const blockedTasks: ClosingTask[] = [];

  tasks.forEach((t) => {
    // Check overdue
    let isOverdue = false;
    if (t.status !== 'completed' && t.dueDate) {
      const [d, m, y] = t.dueDate.split('-').map(Number);
      const dueDateObj = new Date(y, m - 1, d);
      if (dueDateObj < today) {
        isOverdue = true;
        overdueCount++;
      }
    }
    t.isOverdue = isOverdue;

    if (t.status === 'completed') {
      completedCount++;
    } else {
      outstandingTasks.push(t);
      if (t.status === 'in_progress') inProgressCount++;
      else if (t.status === 'blocked') {
        blockedCount++;
        blockedTasks.push(t);
      } else notStartedCount++;
    }
  });

  const totalTasks = tasks.length;
  const completionPercentage = totalTasks > 0 ? (completedCount / totalTasks) * 100 : 0;

  return {
    totalTasks,
    completedCount,
    inProgressCount,
    blockedCount,
    notStartedCount,
    overdueCount,
    completionPercentage,
    outstandingTasks,
    blockedTasks,
  };
}

/**
 * Creates a duplicate of an existing checklist for a new period,
 * intelligently resetting completed statuses to 'not_started' without carrying forward past completions.
 */
export function duplicateChecklistForNewPeriod(
  currentTasks: ClosingTask[],
  newPeriod: string,
  newDueDateOffsetDays = 30
): ClosingTask[] {
  return currentTasks.map((t, idx) => {
    // Calculate new due date by adding days
    let newDueDate = t.dueDate;
    if (t.dueDate) {
      const [d, m, y] = t.dueDate.split('-').map(Number);
      const dObj = new Date(y, m - 1, d);
      dObj.setDate(dObj.getDate() + newDueDateOffsetDays);
      const nd = String(dObj.getDate()).padStart(2, '0');
      const nm = String(dObj.getMonth() + 1).padStart(2, '0');
      newDueDate = `${nd}-${nm}-${dObj.getFullYear()}`;
    }

    return {
      id: `task-${Date.now()}-${idx + 1}`,
      category: t.category,
      title: t.title,
      description: t.description,
      assignee: t.assignee,
      dueDate: newDueDate,
      status: 'not_started',
      priority: t.priority,
      notes: '',
      supportingDocRef: '',
      isOverdue: false,
      history: [
        {
          timestamp: new Date().toISOString(),
          action: `Duplicated from prior month for ${newPeriod}`,
          changedBy: 'System / Closing Manager',
        },
      ],
    };
  });
}

/**
 * Exports complete month-end closing workbook
 */
export function exportMonthEndClosingWorkbook(
  state: MonthEndClosingState,
  metrics: ClosingMetrics
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Executive Summary
  const summaryRows = [
    { Parameter: 'Company Entity', Value: state.companyName },
    { Parameter: 'Accounting Period', Value: state.accountingPeriod },
    { Parameter: 'Financial Year', Value: state.financialYear },
    { Parameter: 'Target Closing Date', Value: state.targetClosingDate },
    { Parameter: 'Total Closing Activities', Value: metrics.totalTasks },
    { Parameter: 'Completed Tasks', Value: metrics.completedCount },
    { Parameter: 'Tasks In Progress', Value: metrics.inProgressCount },
    { Parameter: 'Blocked / Exception Tasks', Value: metrics.blockedCount },
    { Parameter: 'Not Started', Value: metrics.notStartedCount },
    { Parameter: 'Overdue Activities', Value: metrics.overdueCount },
    { Parameter: 'Overall Closing Progress', Value: `${metrics.completionPercentage.toFixed(1)}%` },
    { Parameter: 'Statutory Disclaimer', Value: 'Tracks operational completion; does not certify statutory compliance.' },
  ];
  const wsSummary = createStyledWorksheet(summaryRows, {
    columnWidths: { Parameter: 36, Value: 34 },
  });
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Closing_Summary');

  // Sheet 2: Master Checklist
  const taskRows = state.tasks.map((t, idx) => ({
    '#': idx + 1,
    'Category': t.category,
    'Activity / Task Title': t.title,
    'Priority': t.priority.toUpperCase(),
    'Assignee': t.assignee,
    'Due Date': t.dueDate,
    'Status': t.status.replace(/_/g, ' ').toUpperCase(),
    'Completed Date': t.completedDate || '-',
    'Overdue Flag': t.isOverdue ? 'OVERDUE' : 'ON TRACK',
    'Auditor Notes / Findings': t.notes || '-',
    'Supporting File Ref': t.supportingDocRef || '-',
  }));
  const wsTasks = createStyledWorksheet(taskRows, {
    columnWidths: {
      '#': 6,
      'Category': 24,
      'Activity / Task Title': 36,
      'Priority': 14,
      'Assignee': 20,
      'Due Date': 14,
      'Status': 16,
      'Completed Date': 16,
      'Overdue Flag': 14,
      'Auditor Notes / Findings': 45,
      'Supporting File Ref': 26,
    },
  });
  XLSX.utils.book_append_sheet(wb, wsTasks, 'Closing_Checklist');

  // Sheet 3: Outstanding & Blocked Items
  const exceptionRows = metrics.outstandingTasks.map((t, idx) => ({
    '#': idx + 1,
    'Category': t.category,
    'Task Title': t.title,
    'Current Status': t.status.replace(/_/g, ' ').toUpperCase(),
    'Assignee': t.assignee,
    'Due Date': t.dueDate,
    'Reason / Blockers': t.notes || 'Awaiting action',
    'Required Action': t.status === 'blocked' ? 'ESCALATE TO CA MANAGER' : 'COMPLETE BEFORE SIGN-OFF',
  }));
  const wsExceptions = createStyledWorksheet(exceptionRows, {
    columnWidths: {
      '#': 6,
      'Category': 24,
      'Task Title': 36,
      'Current Status': 18,
      'Assignee': 20,
      'Due Date': 14,
      'Reason / Blockers': 40,
      'Required Action': 30,
    },
  });
  XLSX.utils.book_append_sheet(wb, wsExceptions, 'Outstanding_Items');

  return wb;
}

/**
 * Returns initial standard tasks with unique IDs
 */
export function getInitialClosingTasks(): ClosingTask[] {
  return STANDARD_CLOSING_TEMPLATE_TASKS.map((t, idx) => ({
    ...t,
    id: `closing-task-${idx + 1}`,
    isOverdue: false,
    history: [
      {
        timestamp: new Date().toISOString(),
        action: 'Created task from standard CA closing template',
        changedBy: 'Audit Senior',
      },
    ],
  }));
}
