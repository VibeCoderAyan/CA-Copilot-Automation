/**
 * Module 7: Multi-Client Batch Automation Service
 * Dispatches and coordinates parallel, firm-isolated processing jobs across multiple clients
 */

import type { BatchRun, BatchItem, FirmContext } from '../types';
import { BackendStore } from '../db/store';
import { WorkflowOrchestrator } from './workflowOrchestrator';
import { ExportService } from './exportService';

export class MultiClientBatchService {
  private store: BackendStore;
  private orchestrator: WorkflowOrchestrator;
  private exportService: ExportService;

  constructor(store?: BackendStore) {
    this.store = store || BackendStore.getInstance();
    this.orchestrator = new WorkflowOrchestrator(this.store);
    this.exportService = new ExportService(this.store);
  }

  /**
   * Starts a multi-client batch execution
   */
  public async startBatchRun(
    context: FirmContext,
    workflowId: string,
    clientItems: {
      clientId: string;
      clientName: string;
      files: { fileName: string; buffer: Buffer }[];
    }[]
  ): Promise<BatchRun> {
    const definition = this.store.getWorkflowDefinition(workflowId);
    if (!definition) throw new Error(`Workflow definition "${workflowId}" not found`);

    const batchId = `batch-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`;
    const items: BatchItem[] = clientItems.map((ci) => ({
      clientId: ci.clientId,
      clientName: ci.clientName,
      status: 'queued',
      fileCount: ci.files.length,
      recordsProcessed: 0,
      exceptionsCount: 0,
    }));

    const batchRun: BatchRun = {
      id: batchId,
      firmId: context.firmId,
      workflowId,
      workflowName: definition.name,
      startedAt: new Date().toISOString(),
      status: 'running',
      totalClients: clientItems.length,
      completedClients: 0,
      failedClients: 0,
      items,
    };

    this.store.createBatchRun(batchRun);

    // Audit Event
    this.store.logAuditEvent({
      id: `aud-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`,
      firmId: context.firmId,
      clientId: 'multi_client_batch',
      userId: context.userId,
      userRole: context.userRole,
      action: 'MULTI_CLIENT_BATCH_STARTED',
      entityType: 'job',
      entityId: batchId,
      details: `Started multi-client batch run "${definition.name}" across ${clientItems.length} clients`,
      timestamp: new Date().toISOString(),
    });

    // Execute client workflows in isolated context
    this.executeBatchItems(context, batchId, workflowId, clientItems).catch((err) => {
      console.error(`[MultiClientBatchService] Batch ${batchId} encountered error:`, err);
    });

    return batchRun;
  }

  /**
   * Internal Batch Item Executor
   */
  private async executeBatchItems(
    context: FirmContext,
    batchId: string,
    workflowId: string,
    clientItems: {
      clientId: string;
      clientName: string;
      files: { fileName: string; buffer: Buffer }[];
    }[]
  ): Promise<void> {
    const batch = this.store.getBatchRun(context.firmId, batchId);
    if (!batch) return;

    let completed = 0;
    let failed = 0;

    for (const item of clientItems) {
      const clientContext: FirmContext = {
        ...context,
        clientId: item.clientId,
        clientName: item.clientName,
      };

      const batchItem = batch.items.find((bi) => bi.clientId === item.clientId);
      if (batchItem) batchItem.status = 'running';
      this.store.updateBatchRun(context.firmId, batchId, { items: batch.items });

      try {
        const { job, run } = await this.orchestrator.startWorkflow(
          clientContext,
          workflowId,
          item.files,
          { idempotencyKey: `idemp-batch-${batchId}-${item.clientId}` }
        );

        if (batchItem) {
          batchItem.jobId = job.id;
          batchItem.runId = run.id;
          batchItem.status = job.status === 'review_required' ? 'review_required' : 'completed';
          batchItem.recordsProcessed = run.statistics.totalRecords;
          batchItem.exceptionsCount = run.statistics.exceptionsCount;
        }
        completed++;
      } catch (err: any) {
        failed++;
        if (batchItem) {
          batchItem.status = 'failed';
          batchItem.errorMessage = err.message || 'Workflow execution error';
        }
      }

      this.store.updateBatchRun(context.firmId, batchId, {
        completedClients: completed,
        failedClients: failed,
        items: batch.items,
      });
    }

    const finalStatus = failed === 0 ? 'completed' : completed > 0 ? 'partial_failure' : 'failed';
    this.store.updateBatchRun(context.firmId, batchId, {
      status: finalStatus,
      completedAt: new Date().toISOString(),
    });
  }

  /**
   * Generates Consolidated Batch Execution Summary
   */
  public exportBatchSummary(context: FirmContext, batchId: string) {
    const batch = this.store.getBatchRun(context.firmId, batchId);
    if (!batch) throw new Error('Batch run not found');

    const summaryRows = [
      {
        'Batch ID': batch.id,
        'Workflow Name': batch.workflowName,
        'Total Clients': batch.totalClients,
        'Completed Clients': batch.completedClients,
        'Failed Clients': batch.failedClients,
        Status: batch.status.toUpperCase(),
        'Started At': batch.startedAt,
        'Completed At': batch.completedAt || 'In Progress',
      },
    ];

    const clientRows = batch.items.map((it) => ({
      'Client ID': it.clientId,
      'Client Name': it.clientName,
      Status: it.status.toUpperCase(),
      'Files Processed': it.fileCount,
      'Records Processed': it.recordsProcessed,
      Exceptions: it.exceptionsCount,
      'Associated Job ID': it.jobId || 'N/A',
      'Error Message': it.errorMessage || 'None',
    }));

    return this.exportService.generateExcelWorkbook(
      context,
      `batch-${batchId}`,
      `Batch_Execution_Summary_${batchId}.xlsx`,
      [
        { sheetName: 'Executive_Summary', rows: summaryRows },
        { sheetName: 'Client_Breakdown', rows: clientRows },
      ]
    );
  }
}
