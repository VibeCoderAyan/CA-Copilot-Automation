/**
 * Module 1: Month-End Closing Automation Service
 * Manages period-end closing checklists, prerequisite dependencies, exception checks, and formal closing locks
 */

import type { ClosingPeriod, ClosingTaskItem, FirmContext } from '../types';
import { BackendStore } from '../db/store';
import { ExportService } from './exportService';

export class MonthEndClosingService {
  private store: BackendStore;
  private exportService: ExportService;

  constructor(store?: BackendStore) {
    this.store = store || BackendStore.getInstance();
    this.exportService = new ExportService(this.store);
  }

  /**
   * Creates a new closing period with standard CA closing checklist
   */
  public createPeriod(
    context: FirmContext,
    periodName: string, // e.g. "September 2026"
    financialYear: string, // e.g. "FY 2026-27"
    targetClosingDate: string = '10-10-2026'
  ): { period: ClosingPeriod; tasks: ClosingTaskItem[] } {
    const periodId = `period-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`;

    const period: ClosingPeriod = {
      id: periodId,
      firmId: context.firmId,
      clientId: context.clientId,
      periodName,
      financialYear,
      targetClosingDate,
      status: 'in_progress',
      progressPercent: 0,
      unresolvedExceptionsCount: 0,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };

    const standardTasks: ClosingTaskItem[] = [
      {
        id: `task-${periodId}-1`,
        periodId,
        firmId: context.firmId,
        clientId: context.clientId,
        category: 'Cash & Bank',
        title: 'Bank Reconciliation (BRS) & Continuity Audit',
        description: 'Verify 100% running balance continuity and reconcile bank statement with cashbook',
        assignee: 'Articleship Assistant',
        dueDate: '04-10-2026',
        status: 'not_started',
        priority: 'critical',
        associatedWorkflowId: 'wf_bank_rec',
        notes: 'Mandatory before ledger closing',
      },
      {
        id: `task-${periodId}-2`,
        periodId,
        firmId: context.firmId,
        clientId: context.clientId,
        category: 'Receivables & Payables',
        title: 'AP / AR Sub-ledger vs Control Accounts Reconciliation',
        description: 'Check customer aging, MSME 45-day overdue compliance, and duplicate payment risks',
        assignee: 'Audit Senior',
        dueDate: '06-10-2026',
        status: 'not_started',
        priority: 'high',
        prerequisiteTaskIds: [`task-${periodId}-1`],
        notes: 'Flag credit debtors and debit creditors',
      },
      {
        id: `task-${periodId}-3`,
        periodId,
        firmId: context.firmId,
        clientId: context.clientId,
        category: 'Taxation & Statutory',
        title: 'GST GSTR-2B vs Purchase Register Reconciliation',
        description: 'Match inward invoices, check Section 16(2)(aa) compliance, and segregate ineligible ITC',
        assignee: 'Tax Article',
        dueDate: '08-10-2026',
        status: 'not_started',
        priority: 'critical',
        associatedWorkflowId: 'wf_gst_rec',
        notes: 'Confirm vendor return filing status',
      },
      {
        id: `task-${periodId}-4`,
        periodId,
        firmId: context.firmId,
        clientId: context.clientId,
        category: 'General Ledger & Journals',
        title: 'Trial Balance Balance & Schedule III Verification',
        description: 'Verify total debits == total credits, check abnormal cash polarity, and classify heads',
        assignee: 'CA Partner',
        dueDate: '09-10-2026',
        status: 'not_started',
        priority: 'critical',
        prerequisiteTaskIds: [`task-${periodId}-1`, `task-${periodId}-2`, `task-${periodId}-3`],
        associatedWorkflowId: 'wf_tb_audit',
        notes: 'Final TB must have 0 variance',
      },
      {
        id: `task-${periodId}-5`,
        periodId,
        firmId: context.firmId,
        clientId: context.clientId,
        category: 'Accruals & Cut-off',
        title: 'Year-End / Month-End Accruals & Adjustments',
        description: 'Post closing stock, depreciation, prepaid expenses, and accrued liabilities',
        assignee: 'CA Partner',
        dueDate: '10-10-2026',
        status: 'not_started',
        priority: 'high',
        prerequisiteTaskIds: [`task-${periodId}-4`],
        notes: 'Requires dual review',
      },
    ];

    this.store.createClosingPeriod(period);
    this.store.setClosingTasks(periodId, standardTasks);

    // Audit Event
    this.store.logAuditEvent({
      id: `aud-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`,
      firmId: context.firmId,
      clientId: context.clientId,
      userId: context.userId,
      userRole: context.userRole,
      action: 'CLOSING_PERIOD_CREATED',
      entityType: 'workflow',
      entityId: periodId,
      details: `Created closing period "${periodName}" (${financialYear}) with 5 checklist prerequisites`,
      timestamp: new Date().toISOString(),
    });

    return { period, tasks: standardTasks };
  }

  /**
   * Evaluates current closing status, incomplete prerequisites, and unresolved exceptions
   */
  public evaluateClosingStatus(context: FirmContext, periodId: string) {
    const period = this.store.getClosingPeriod(context.firmId, periodId);
    if (!period) throw new Error('Closing period not found');

    const tasks = this.store.getClosingTasks(periodId);
    const completedTasks = tasks.filter((t) => t.status === 'completed');
    const blockedTasks = tasks.filter((t) => t.status === 'blocked');
    const unresolvedExceptions = this.store.listExceptions(context.firmId, {
      clientId: context.clientId,
      status: 'pending',
    });

    // Check prerequisites
    const incompletePrerequisites: string[] = [];
    for (const task of tasks) {
      if (task.prerequisiteTaskIds && task.prerequisiteTaskIds.length > 0) {
        for (const preId of task.prerequisiteTaskIds) {
          const preTask = tasks.find((t) => t.id === preId);
          if (preTask && preTask.status !== 'completed') {
            incompletePrerequisites.push(
              `Task "${task.title}" is blocked by incomplete prerequisite: "${preTask.title}"`
            );
          }
        }
      }
    }

    const progressPercent = Math.round((completedTasks.length / tasks.length) * 100);
    const canClose =
      completedTasks.length === tasks.length &&
      unresolvedExceptions.length === 0 &&
      blockedTasks.length === 0;

    this.store.updateClosingPeriod(context.firmId, periodId, {
      progressPercent,
      unresolvedExceptionsCount: unresolvedExceptions.length,
      status: canClose ? 'review_required' : 'in_progress',
    });

    return {
      period,
      tasks,
      metrics: {
        totalTasks: tasks.length,
        completedTasks: completedTasks.length,
        blockedTasks: blockedTasks.length,
        progressPercent,
        unresolvedExceptionsCount: unresolvedExceptions.length,
        canClose,
      },
      incompletePrerequisites,
      unresolvedExceptions,
    };
  }

  /**
   * Approves final closing of the accounting period with strict validation lock
   */
  public approveClosing(
    context: FirmContext,
    periodId: string,
    approvalNotes?: string
  ): { success: boolean; period?: ClosingPeriod; error?: string } {
    const status = this.evaluateClosingStatus(context, periodId);

    // Strict validation: never close if prerequisites or exceptions remain open
    if (!status.metrics.canClose) {
      const reasons: string[] = [];
      if (status.metrics.completedTasks < status.metrics.totalTasks) {
        reasons.push(`${status.metrics.totalTasks - status.metrics.completedTasks} checklist tasks incomplete`);
      }
      if (status.metrics.unresolvedExceptionsCount > 0) {
        reasons.push(`${status.metrics.unresolvedExceptionsCount} unresolved exceptions pending auditor review`);
      }
      return {
        success: false,
        error: `Cannot close period. Unmet conditions: ${reasons.join(', ')}`,
      };
    }

    const updated = this.store.updateClosingPeriod(context.firmId, periodId, {
      status: 'closed',
      approvedBy: `${context.userId} (${context.userRole})`,
      approvedAt: new Date().toISOString(),
      approvalNotes: approvalNotes || 'Period successfully audited and closed',
    });

    this.store.logAuditEvent({
      id: `aud-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`,
      firmId: context.firmId,
      clientId: context.clientId,
      userId: context.userId,
      userRole: context.userRole,
      action: 'CLOSING_PERIOD_APPROVED',
      entityType: 'workflow',
      entityId: periodId,
      details: `Period "${status.period.periodName}" locked and closed by ${context.userId}`,
      timestamp: new Date().toISOString(),
    });

    return { success: true, period: updated };
  }

  /**
   * Exports Closing Summary Workbook
   */
  public exportClosingSummary(context: FirmContext, periodId: string) {
    const period = this.store.getClosingPeriod(context.firmId, periodId);
    if (!period) throw new Error('Closing period not found');

    const tasks = this.store.getClosingTasks(periodId);
    const taskRows = tasks.map((t) => ({
      Category: t.category,
      Task: t.title,
      Assignee: t.assignee,
      'Due Date': t.dueDate,
      Status: t.status,
      Priority: t.priority,
      Notes: t.notes,
    }));

    const summaryRow = [
      {
        'Client ID': context.clientId,
        'Period Name': period.periodName,
        'Financial Year': period.financialYear,
        Status: period.status,
        'Progress %': `${period.progressPercent}%`,
        'Approved By': period.approvedBy || 'Pending',
        'Approved At': period.approvedAt || 'Pending',
      },
    ];

    return this.exportService.generateExcelWorkbook(
      context,
      `closing-${periodId}`,
      `Month_End_Closing_${period.periodName.replace(/\s+/g, '_')}.xlsx`,
      [
        { sheetName: 'Closing_Summary', rows: summaryRow },
        { sheetName: 'Closing_Checklist', rows: taskRows },
      ]
    );
  }
}
