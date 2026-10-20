/**
 * Module 6: Recurring Compliance & Task Automation Service
 * Schedules recurring accounting & statutory tasks without duplicate generation
 */

import type { TaskInstance, FirmContext } from '../types';
import { BackendStore } from '../db/store';
import { ExportService } from './exportService';

export class RecurringComplianceService {
  private store: BackendStore;
  private exportService: ExportService;

  constructor(store?: BackendStore) {
    this.store = store || BackendStore.getInstance();
    this.exportService = new ExportService(this.store);
  }

  /**
   * Generates task instances for a specific client and period based on active recurring definitions
   */
  public generatePeriodInstances(
    context: FirmContext,
    period: string = 'September 2026',
    referenceYear: number = 2026,
    referenceMonth: number = 9 // September
  ): { createdCount: number; instances: TaskInstance[] } {
    const definitions = this.store.listRecurringTaskDefinitions(context.firmId);
    const existingInstances = this.store.listTaskInstances(context.clientId);
    const existingKeys = new Set(
      existingInstances.map((t) => `${t.definitionId}|${t.period}`)
    );

    const newInstances: TaskInstance[] = [];

    for (const def of definitions) {
      const dedupKey = `${def.id}|${period}`;
      if (existingKeys.has(dedupKey)) continue;

      // Calculate Statutory Due Date (following month for monthly tasks)
      let dueYear = referenceYear;
      let dueMonth = referenceMonth + 1;
      if (dueMonth > 12) {
        dueMonth = 1;
        dueYear++;
      }
      const dueDay = String(def.dueDayOfMonth).padStart(2, '0');
      const dueMonthStr = String(dueMonth).padStart(2, '0');
      const dueDate = `${dueYear}-${dueMonthStr}-${dueDay}`;

      const instance: TaskInstance = {
        id: `tinst-${Date.now()}-${newInstances.length + 1}`,
        firmId: context.firmId,
        clientId: context.clientId,
        definitionId: def.id,
        title: `${def.title} - ${period}`,
        period,
        dueDate,
        status: 'upcoming',
        assignedTo: 'Compliance Team',
      };

      newInstances.push(instance);
      existingKeys.add(dedupKey);
    }

    const merged = [...existingInstances, ...newInstances];
    this.store.saveTaskInstances(context.clientId, merged);

    return {
      createdCount: newInstances.length,
      instances: merged,
    };
  }

  /**
   * Refreshes status of tasks against current date (upcoming, due_today, overdue)
   */
  public evaluateTaskStatuses(
    context: FirmContext,
    currentDateStr: string = new Date().toISOString().split('T')[0]
  ): TaskInstance[] {
    const instances = this.store.listTaskInstances(context.clientId);
    const currTime = new Date(currentDateStr).getTime();

    for (const inst of instances) {
      if (inst.status === 'completed') continue;

      const dueTime = new Date(inst.dueDate).getTime();
      const diffDays = Math.round((dueTime - currTime) / (1000 * 60 * 60 * 24));

      if (diffDays < 0) {
        inst.status = 'overdue';
      } else if (diffDays === 0) {
        inst.status = 'due_today';
      } else {
        inst.status = 'upcoming';
      }
    }

    this.store.saveTaskInstances(context.clientId, instances);
    return instances;
  }

  /**
   * Marks a compliance task as completed with audit proof reference
   */
  public completeTask(
    context: FirmContext,
    instanceId: string,
    proofRef: string
  ): { success: boolean; instance?: TaskInstance; error?: string } {
    const instances = this.store.listTaskInstances(context.clientId);
    const target = instances.find((t) => t.id === instanceId);
    if (!target) return { success: false, error: 'Task instance not found' };

    const updated = this.store.updateTaskInstance(context.clientId, instanceId, {
      status: 'completed',
      completedAt: new Date().toISOString(),
      completionProofRef: proofRef,
      verifiedBy: `${context.userId} (${context.userRole})`,
    });

    this.store.logAuditEvent({
      id: `aud-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`,
      firmId: context.firmId,
      clientId: context.clientId,
      userId: context.userId,
      userRole: context.userRole,
      action: 'COMPLIANCE_TASK_COMPLETED',
      entityType: 'workflow',
      entityId: instanceId,
      details: `Completed compliance task "${target.title}". Proof / Ack Ref: ${proofRef}`,
      timestamp: new Date().toISOString(),
    });

    return { success: true, instance: updated };
  }

  /**
   * Exports Statutory Compliance Calendar Report
   */
  public exportComplianceCalendar(context: FirmContext) {
    const list = this.store.listTaskInstances(context.clientId);
    const rows = list.map((t) => ({
      Task: t.title,
      Period: t.period,
      'Due Date': t.dueDate,
      Status: t.status.toUpperCase(),
      Assignee: t.assignedTo,
      'Completed Date': t.completedAt || 'Pending',
      'Filing Ack / Proof': t.completionProofRef || 'N/A',
      'Verified By': t.verifiedBy || 'Pending',
    }));

    return this.exportService.generateExcelWorkbook(
      context,
      `compliance-${context.clientId}`,
      `Statutory_Compliance_Schedule_${context.clientId}.xlsx`,
      [{ sheetName: 'Compliance_Calendar', rows }]
    );
  }
}
