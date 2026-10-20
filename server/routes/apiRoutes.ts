/**
 * Complete REST API Routes for CA Tools Backend Automation Engine
 * Supports Core Engine + All 7 Advanced Automation Modules
 */

import { Router } from 'express';
import type { Request, Response, NextFunction } from 'express';
import type { FirmContext } from '../types/index.ts';
import { BackendStore } from '../db/store.ts';
import { WorkflowOrchestrator } from '../services/workflowOrchestrator.ts';
import { UniversalMatchingEngine } from '../services/matchingEngine.ts';
import { AccountingRulesEngine } from '../services/accountingRulesEngine.ts';
import { ExceptionService } from '../services/exceptionService.ts';
import { MonthEndClosingService } from '../services/monthEndClosingService.ts';
import { ApArAutomationService } from '../services/apArService.ts';
import { AuditEvidenceService } from '../services/auditEvidenceService.ts';
import { AnomalyDetectionService } from '../services/anomalyDetectionService.ts';
import { DocumentCollectionService } from '../services/documentCollectionService.ts';
import { RecurringComplianceService } from '../services/recurringComplianceService.ts';
import { MultiClientBatchService } from '../services/multiClientBatchService.ts';
import { TallyXmlExportService } from '../services/tallyXmlExportService.ts';
import { GovJsonParserService } from '../services/govJsonParserService.ts';
import { AsyncWorkerQueue } from '../services/asyncWorkerQueue.ts';
import { ComplianceReportService } from '../services/complianceReportService.ts';
import { Form3cdService } from '../services/form3cdService.ts';
import { GstTripleRecService } from '../services/gstTripleRecService.ts';
import { Section206abService } from '../services/section206abService.ts';
import { StreamingProcessorService } from '../services/streamingProcessorService.ts';
import { TallyLiveSyncService } from '../services/tallyLiveSyncService.ts';
import { HotFolderWatcherService } from '../services/hotFolderWatcherService.ts';
import { OfficeLanService } from '../services/officeLanService.ts';
import { LocalDbEngine } from '../db/localDbEngine.ts';
import { ClientQuerySheetService } from '../services/clientQuerySheetService.ts';
import { NightlyAuditCronService } from '../services/nightlyAuditCronService.ts';

export const apiRouter = Router();
const store = BackendStore.getInstance();
const orchestrator = new WorkflowOrchestrator(store);
const exceptionService = new ExceptionService(store);
const closingService = new MonthEndClosingService(store);
const apArService = new ApArAutomationService(store);
const evidenceService = new AuditEvidenceService(store);
const anomalyService = new AnomalyDetectionService(store);
const docService = new DocumentCollectionService(store);
const complianceService = new RecurringComplianceService(store);
const batchService = new MultiClientBatchService(store);
const tallyService = new TallyXmlExportService(store);
const govJsonService = new GovJsonParserService(store);
const complianceReportService = new ComplianceReportService(store);
const form3cdService = new Form3cdService(store);
const gstTripleRecService = new GstTripleRecService(store);
const section206abService = new Section206abService(store);
const workerQueue = AsyncWorkerQueue.getInstance();

const tallyLiveSyncService = new TallyLiveSyncService();
const hotFolderService = new HotFolderWatcherService();
const officeLanService = new OfficeLanService();
const localDbEngine = LocalDbEngine.getInstance();
const clientQueryService = new ClientQuerySheetService();
const nightlyAuditService = new NightlyAuditCronService();

/**
 * Firm Context Extraction Middleware
 * Enforces firm-level multi-tenancy and data isolation
 */
export function extractFirmContext(req: Request, _res: Response, next: NextFunction): void {
  const firmId = (req.headers['x-firm-id'] as string) || (req.query.firmId as string) || 'firm_apex_ca';
  const firmName = (req.headers['x-firm-name'] as string) || 'Apex Chartered Accountants & Co.';
  const userId = (req.headers['x-user-id'] as string) || 'ca_partner_01';
  const userRole = ((req.headers['x-user-role'] as string) || 'Partner') as FirmContext['userRole'];
  const clientId = (req.headers['x-client-id'] as string) || (req.query.clientId as string) || 'client_apex_technologies';

  (req as any).firmContext = {
    firmId,
    firmName,
    userId,
    userRole,
    clientId,
  };

  next();
}

apiRouter.use(extractFirmContext);

// ==========================================
// 1. HEALTH & SYSTEM STATUS
// ==========================================
apiRouter.get('/health', (_req: Request, res: Response) => {
  res.json({
    status: 'healthy',
    timestamp: new Date().toISOString(),
    engine: 'CA Tools Intelligent Backend Automation Engine',
    version: '2.0.0',
    activeWorkflows: store.listWorkflowDefinitions().length,
  });
});

// ==========================================
// 2. WORKFLOW ORCHESTRATION PIPELINES
// ==========================================
apiRouter.get('/workflows', (_req: Request, res: Response) => {
  const definitions = store.listWorkflowDefinitions();
  res.json({ success: true, count: definitions.length, data: definitions });
});

apiRouter.get('/workflows/:id', (req: Request, res: Response) => {
  const paramId = String(req.params.id);
  const def = store.getWorkflowDefinition(paramId);
  if (!def) return res.status(404).json({ success: false, error: 'Workflow definition not found' });
  res.json({ success: true, data: def });
});

apiRouter.post('/workflows/:id/run', async (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const workflowId = String(req.params.id);
    const { files = [], idempotencyKey, configOverrides } = req.body;

    if (!files || files.length === 0) {
      return res.status(400).json({ success: false, error: 'At least one file is required to start workflow' });
    }

    const bufferFiles = files.map((f: any) => {
      let buffer: Buffer;
      if (f.base64) {
        buffer = Buffer.from(f.base64, 'base64');
      } else if (f.content) {
        buffer = Buffer.from(f.content, 'utf-8');
      } else {
        buffer = Buffer.from('');
      }
      return { fileName: f.fileName, buffer };
    });

    const result = await orchestrator.startWorkflow(context, workflowId, bufferFiles, {
      idempotencyKey,
      configOverrides,
    });

    res.status(202).json({
      success: true,
      message: 'Workflow started successfully in background queue',
      jobId: result.job.id,
      runId: result.run.id,
      status: result.job.status,
    });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

apiRouter.get('/workflows/runs/:runId', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const runId = String(req.params.runId);
  const run = store.getWorkflowRun(context.firmId, runId);
  if (!run) return res.status(404).json({ success: false, error: 'Workflow run not found' });
  const matches = store.getRecordMatches(run.id);
  const exceptions = store.listExceptions(context.firmId, { workflowRunId: run.id });

  res.json({
    success: true,
    data: {
      ...run,
      matches,
      exceptions,
    },
  });
});

// ==========================================
// 3. BACKGROUND PROCESSING JOBS
// ==========================================
apiRouter.get('/jobs', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const jobs = store.listJobs(context.firmId, req.query.clientId as string);
  res.json({ success: true, count: jobs.length, data: jobs });
});

apiRouter.get('/jobs/:id', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const jobId = String(req.params.id);
  const job = store.getJob(context.firmId, jobId);
  if (!job) return res.status(404).json({ success: false, error: 'Job not found' });
  const artifacts = store.listArtifacts(context.firmId, job.runId);
  res.json({ success: true, data: { ...job, artifacts } });
});

apiRouter.post('/jobs/:id/retry', async (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const jobId = String(req.params.id);
    const retriedJob = await orchestrator.retryJob(context, jobId);
    res.json({ success: true, message: 'Job retry initiated', data: retriedJob });
  } catch (err: any) {
    res.status(400).json({ success: false, error: err.message });
  }
});

apiRouter.post('/jobs/:id/cancel', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const jobId = String(req.params.id);
  const updated = store.updateJob(context.firmId, jobId, {
    status: 'cancelled',
    currentStep: 'Cancelled by User',
  });
  if (!updated) return res.status(404).json({ success: false, error: 'Job not found' });
  res.json({ success: true, message: 'Job cancelled successfully', data: updated });
});

// SSE Real-Time Job Progress Streaming
apiRouter.get('/jobs/:id/stream', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const jobId = String(req.params.id);
  const job = store.getJob(context.firmId, jobId);
  if (!job) {
    res.status(404).json({ success: false, error: 'Job not found' });
    return;
  }

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');

  // Send initial state
  res.write(`data: ${JSON.stringify({ type: 'initial', job })}\n\n`);

  const heartbeat = setInterval(() => {
    res.write(': keepalive\n\n');
  }, 15000);

  const onProgress = (updatedJob: any) => {
    res.write(`data: ${JSON.stringify({ type: 'progress', job: updatedJob })}\n\n`);
    if (updatedJob.status === 'completed' || updatedJob.status === 'failed') {
      res.write(`data: ${JSON.stringify({ type: 'done', status: updatedJob.status })}\n\n`);
      clearInterval(heartbeat);
      res.end();
    }
  };

  orchestrator.on(`job:${jobId}`, onProgress);

  req.on('close', () => {
    clearInterval(heartbeat);
    orchestrator.off(`job:${jobId}`, onProgress);
  });
});

// ==========================================
// 4. STANDALONE MATCHING & RULES
// ==========================================
apiRouter.post('/matching/execute', (req: Request, res: Response) => {
  try {
    const { primaryRecords = [], secondaryRecords = [], config = {} } = req.body;
    const result = UniversalMatchingEngine.executeMatching(
      primaryRecords,
      secondaryRecords,
      `standalone-${Date.now()}`,
      config
    );
    res.json({ success: true, data: result });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

apiRouter.post('/rules/validate', (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const { records = [], docType = 'general' } = req.body;
    const config = store.getFirmConfig(context.firmId);
    const issues = AccountingRulesEngine.validateDataset(records, config, docType);
    res.json({ success: true, count: issues.length, data: issues });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

apiRouter.get('/rules/config', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const config = store.getFirmConfig(context.firmId);
  res.json({ success: true, data: config });
});

apiRouter.put('/rules/config', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const updated = store.updateFirmConfig(context.firmId, req.body);
  res.json({ success: true, message: 'Firm accounting rules updated', data: updated });
});

// ==========================================
// 5. EXCEPTIONS & AUDIT LOGS
// ==========================================
apiRouter.get('/exceptions', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const exceptions = exceptionService.listExceptions(context, {
    clientId: req.query.clientId as string,
    workflowRunId: req.query.workflowRunId as string,
    status: req.query.status as string,
  });
  res.json({ success: true, count: exceptions.length, data: exceptions });
});

apiRouter.post('/exceptions/:id/resolve', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const exceptionId = String(req.params.id);
  const { resolution, resolutionNotes, correctedValues } = req.body;
  if (!resolution || !['approved', 'rejected', 'corrected', 'investigating', 'dismissed'].includes(resolution)) {
    return res.status(400).json({ success: false, error: 'Valid resolution status is required' });
  }

  const result = exceptionService.resolveException(
    context,
    exceptionId,
    resolution,
    resolutionNotes,
    correctedValues
  );

  if (!result.success) {
    return res.status(404).json({ success: false, error: result.error });
  }

  res.json({ success: true, message: `Exception marked as ${resolution}`, data: result });
});

apiRouter.get('/audit', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const events = store.queryAuditEvents(context.firmId, {
    clientId: req.query.clientId as string,
    entityId: req.query.entityId as string,
    entityType: req.query.entityType as string,
  });
  res.json({ success: true, count: events.length, data: events });
});

// ==========================================
// 6. MODULE 1: MONTH-END CLOSING
// ==========================================
apiRouter.post('/closing/periods', (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const { periodName, financialYear, targetClosingDate } = req.body;
    const result = closingService.createPeriod(context, periodName || 'September 2026', financialYear || 'FY 2026-27', targetClosingDate);
    res.json({ success: true, data: result });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

apiRouter.get('/closing/periods', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const periods = store.listClosingPeriods(context.firmId, req.query.clientId as string);
  res.json({ success: true, count: periods.length, data: periods });
});

apiRouter.get('/closing/periods/:id', (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const status = closingService.evaluateClosingStatus(context, String(req.params.id));
    res.json({ success: true, data: status });
  } catch (err: any) {
    res.status(404).json({ success: false, error: err.message });
  }
});

apiRouter.put('/closing/periods/:id/tasks/:taskId', (req: Request, res: Response) => {
  const periodId = String(req.params.id);
  const taskId = String(req.params.taskId);
  const updated = store.updateClosingTask(periodId, taskId, req.body);
  if (!updated) return res.status(404).json({ success: false, error: 'Task not found' });
  res.json({ success: true, data: updated });
});

apiRouter.post('/closing/periods/:id/approve', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const periodId = String(req.params.id);
  const { notes } = req.body;
  const result = closingService.approveClosing(context, periodId, notes);
  if (!result.success) return res.status(400).json({ success: false, error: result.error });
  res.json({ success: true, message: 'Period closed and locked successfully', data: result.period });
});

apiRouter.get('/closing/periods/:id/export', (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const artifact = closingService.exportClosingSummary(context, String(req.params.id));
    res.json({ success: true, data: artifact });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// ==========================================
// 7. MODULE 2: AP & AR AUTOMATION
// ==========================================
apiRouter.post('/apar/process', (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const { records = [], type = 'payable', asOfDate } = req.body;
    const result = apArService.processLedgerRecords(context, records, type, asOfDate);
    res.json({ success: true, data: result });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

apiRouter.get('/apar/records', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const records = store.getAgingRecords(context.clientId, context.firmId);
  res.json({ success: true, count: records.length, data: records });
});

apiRouter.get('/apar/alerts', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const alerts = store.getPaymentRiskAlerts(context.clientId, context.firmId);
  res.json({ success: true, count: alerts.length, data: alerts });
});

apiRouter.get('/apar/payment-review', (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const artifact = apArService.generatePaymentReviewList(context);
    res.json({ success: true, data: artifact });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// ==========================================
// 8. MODULE 3: AUDIT EVIDENCE AUTOMATION
// ==========================================
apiRouter.post('/evidence/init', (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const { period } = req.body;
    const slots = evidenceService.initializeEvidenceSlots(context, period);
    res.json({ success: true, count: slots.length, data: slots });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

apiRouter.get('/evidence', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const evidenceList = store.getAuditEvidence(context.clientId, context.firmId);
  res.json({ success: true, count: evidenceList.length, data: evidenceList });
});

apiRouter.post('/evidence/link', (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const { evidenceId, transaction } = req.body;
    const link = evidenceService.linkTransactionToEvidence(context, evidenceId, transaction);
    res.json({ success: true, data: link });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

apiRouter.post('/evidence/detect-missing', (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const { transactions = [], materialityThreshold = 100000 } = req.body;
    const result = evidenceService.detectMissingEvidence(context, transactions, materialityThreshold);
    res.json({ success: true, data: result });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

apiRouter.get('/evidence/export', (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const artifact = evidenceService.exportEvidenceRegister(context, (req.query.period as string) || 'September 2026');
    res.json({ success: true, data: artifact });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// ==========================================
// 9. MODULE 4: ANOMALY DETECTION ENGINE
// ==========================================
apiRouter.get('/anomalies/rules', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const rules = store.listAnomalyRules(context.firmId);
  res.json({ success: true, count: rules.length, data: rules });
});

apiRouter.put('/anomalies/rules/:id', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const updated = store.updateAnomalyRule(context.firmId, String(req.params.id), req.body);
  if (!updated) return res.status(404).json({ success: false, error: 'Rule not found' });
  res.json({ success: true, data: updated });
});

apiRouter.post('/anomalies/scan', (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const { records = [] } = req.body;
    const result = anomalyService.scanTransactions(context, records);
    res.json({ success: true, count: result.newAlertsCount, data: result });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

apiRouter.get('/anomalies/alerts', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const alerts = store.listAnomalyAlerts(context.clientId, context.firmId);
  res.json({ success: true, count: alerts.length, data: alerts });
});

apiRouter.post('/anomalies/alerts/:id/review', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const { action, notes } = req.body;
  const result = anomalyService.reviewAlert(context, String(req.params.id), action, notes);
  if (!result.success) return res.status(404).json({ success: false, error: result.error });
  res.json({ success: true, data: result.alert });
});

apiRouter.get('/anomalies/export', (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const artifact = anomalyService.exportAnomalyReport(context);
    res.json({ success: true, data: artifact });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// ==========================================
// 10. MODULE 5: CLIENT DOCUMENT COLLECTION
// ==========================================
apiRouter.post('/document-requests/bundle', (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const { financialPeriod, engagementType, dueDate } = req.body;
    const requests = docService.createEngagementBundle(context, financialPeriod, engagementType, dueDate);
    res.json({ success: true, count: requests.length, data: requests });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

apiRouter.get('/document-requests', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const requests = store.listDocumentRequests(context.clientId, context.firmId);
  res.json({ success: true, count: requests.length, data: requests });
});

apiRouter.post('/document-requests/:id/upload', (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const { fileName, fileSizeBytes, fileHash } = req.body;
    const updated = docService.registerClientFile(context, String(req.params.id), fileName, fileSizeBytes || 0, fileHash || 'hash');
    res.json({ success: true, data: updated });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

apiRouter.get('/document-requests/reminder', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const draft = docService.draftReminderNotice(context, req.query.clientName as string, req.query.email as string);
  res.json({ success: true, data: draft });
});

apiRouter.get('/document-requests/export', (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const artifact = docService.exportOutstandingDocumentReport(context);
    res.json({ success: true, data: artifact });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// ==========================================
// 11. MODULE 6: RECURRING COMPLIANCE
// ==========================================
apiRouter.get('/compliance/definitions', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const defs = store.listRecurringTaskDefinitions(context.firmId);
  res.json({ success: true, count: defs.length, data: defs });
});

apiRouter.post('/compliance/generate', (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const { period, referenceYear, referenceMonth } = req.body;
    const result = complianceService.generatePeriodInstances(context, period, referenceYear, referenceMonth);
    res.json({ success: true, count: result.createdCount, data: result.instances });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

apiRouter.get('/compliance/instances', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const list = complianceService.evaluateTaskStatuses(context, (req.query.currentDate as string) || new Date().toISOString().split('T')[0]);
  res.json({ success: true, count: list.length, data: list });
});

apiRouter.post('/compliance/instances/:id/complete', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const { proofRef } = req.body;
  const result = complianceService.completeTask(context, String(req.params.id), proofRef || 'Filed');
  if (!result.success) return res.status(404).json({ success: false, error: result.error });
  res.json({ success: true, data: result.instance });
});

apiRouter.get('/compliance/export', (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const artifact = complianceService.exportComplianceCalendar(context);
    res.json({ success: true, data: artifact });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// ==========================================
// 12. MODULE 7: MULTI-CLIENT BATCH RUNS
// ==========================================
apiRouter.post('/batch/run', async (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const { workflowId, clientItems = [] } = req.body;

    const parsedClientItems = clientItems.map((ci: any) => ({
      clientId: ci.clientId,
      clientName: ci.clientName,
      files: (ci.files || []).map((f: any) => ({
        fileName: f.fileName,
        buffer: f.content ? Buffer.from(f.content, 'utf-8') : f.base64 ? Buffer.from(f.base64, 'base64') : Buffer.from(''),
      })),
    }));

    const batchRun = await batchService.startBatchRun(context, workflowId || 'wf_bank_rec', parsedClientItems);
    res.json({ success: true, data: batchRun });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

apiRouter.get('/batch', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const list = store.listBatchRuns(context.firmId);
  res.json({ success: true, count: list.length, data: list });
});

apiRouter.get('/batch/:id', (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const batch = store.getBatchRun(context.firmId, String(req.params.id));
  if (!batch) return res.status(404).json({ success: false, error: 'Batch run not found' });
  res.json({ success: true, data: batch });
});

apiRouter.get('/batch/:id/export', (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const artifact = batchService.exportBatchSummary(context, String(req.params.id));
    res.json({ success: true, data: artifact });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// ==========================================
// 13. EXPORTS & FILE DOWNLOAD
// ==========================================
const handleArtifactDownload = (req: Request, res: Response) => {
  const context: FirmContext = (req as any).firmContext;
  const artifactId = String(req.params.id);
  let artifact = context?.firmId ? store.getArtifact(context.firmId, artifactId) : undefined;
  if (!artifact) {
    artifact = store.getArtifactById(artifactId);
  }
  if (!artifact) return res.status(404).json({ success: false, error: 'Export artifact not found' });

  const rawBase64 = artifact.contentBase64 || (artifact as any).base64Payload || (artifact as any).xmlPayload;
  if (!rawBase64) {
    return res.status(404).json({ success: false, error: 'Export artifact content unavailable' });
  }

  const buffer = Buffer.from(rawBase64, 'base64');
  const fileExt = artifact.fileType || (artifact.fileName?.endsWith('.csv') ? 'csv' : artifact.fileName?.endsWith('.xml') ? 'xml' : 'xlsx');
  const mimeType = fileExt === 'xlsx'
    ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    : fileExt === 'xml'
    ? 'application/xml'
    : 'text/csv';

  res.setHeader('Content-Type', mimeType);
  res.setHeader('Content-Disposition', `attachment; filename="${artifact.fileName}"`);
  res.setHeader('Content-Length', buffer.length.toString());
  res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
  res.send(buffer);
};

apiRouter.get('/exports/:id/download', handleArtifactDownload);
apiRouter.get('/artifacts/:id/download', handleArtifactDownload);

// ==========================================
// 14. TALLY PRIME XML EXPORTS
// ==========================================
apiRouter.post('/exports/tally', (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const { workflowRunId = `run-${Date.now()}`, records = [], type = 'bank', bankLedgerName } = req.body;

    const vouchers = type === 'invoice'
      ? tallyService.convertInvoicesToTally(records)
      : tallyService.convertBankRecordsToTally(records, bankLedgerName);

    const artifact = tallyService.saveTallyExport(
      context,
      workflowRunId,
      `Tally_${type === 'invoice' ? 'Purchase' : 'Bank'}_Vouchers_${Date.now()}.xml`,
      vouchers
    );

    res.json({ success: true, data: artifact, voucherCount: vouchers.length });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// ==========================================
// 15. GOVERNMENT PORTAL RAW JSON INGESTION
// ==========================================
apiRouter.post('/ingest/gov-json', (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const { fileName, jsonContent } = req.body;

    if (!fileName || !jsonContent) {
      return res.status(400).json({ success: false, error: 'fileName and jsonContent are required' });
    }

    const payloadStr = typeof jsonContent === 'string' ? jsonContent : JSON.stringify(jsonContent);
    const result = govJsonService.ingestGovJsonFile(context, fileName, payloadStr);

    res.json({
      success: true,
      message: `Successfully ingested government JSON: ${result.sourceFile.documentType}`,
      data: {
        sourceFile: result.sourceFile,
        extractedCount: result.records.length,
      },
    });
  } catch (err: any) {
    res.status(400).json({ success: false, error: err.message });
  }
});

// ==========================================
// 16. ICAI SQC 1 AUDIT TRAIL & COMPLIANCE
// ==========================================
apiRouter.get('/audit-trail/report', (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const cert = complianceReportService.generateCertificate(context);

    if (req.query.format === 'html') {
      const html = complianceReportService.generateHtmlReport(cert);
      res.setHeader('Content-Type', 'text/html; charset=utf-8');
      return res.send(html);
    }

    res.json({ success: true, data: cert });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

apiRouter.get('/audit-trail/verify', (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const chained = complianceReportService.generateChainedAuditLog(context.firmId, context.clientId);
    const isValid = complianceReportService.verifyAuditChain(chained);

    res.json({
      success: true,
      data: {
        verified: isValid,
        totalEventsChecked: chained.length,
        status: isValid ? 'CRYPTOGRAPHIC_CHAIN_INTACT' : 'TAMPER_DETECTED',
        headHash: chained.length > 0 ? chained[chained.length - 1].currentHash : 'GENESIS',
      },
    });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// ==========================================
// 17. ASYNC WORKER QUEUE & STORE BACKUP
// ==========================================
apiRouter.get('/worker-queue/stats', (_req: Request, res: Response) => {
  res.json({ success: true, data: workerQueue.getStats() });
});

apiRouter.post('/store/backup', (_req: Request, res: Response) => {
  try {
    const backupPath = store.createBackup();
    res.json({ success: true, message: 'ACID store backup snapshot created', backupPath });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// ==========================================
// 18. FORM 3CD TAX AUDIT (SECTION 44AB)
// ==========================================
apiRouter.post('/audit/form3cd/scrutinize', extractFirmContext, (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const { vouchers, financialYear, assessmentYear } = req.body;
    if (!vouchers || !Array.isArray(vouchers)) {
      return res.status(400).json({ success: false, error: 'vouchers array is required' });
    }
    const result = form3cdService.scrutinizeLedger(
      context.firmId,
      context.clientId,
      vouchers,
      financialYear,
      assessmentYear
    );
    res.json({ success: true, data: result });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

apiRouter.post('/audit/form3cd/export', extractFirmContext, (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const { vouchers, financialYear, assessmentYear, clientLegalName } = req.body;
    if (!vouchers || !Array.isArray(vouchers)) {
      return res.status(400).json({ success: false, error: 'vouchers array is required' });
    }
    const artifact = form3cdService.exportForm3cdWorkbook(
      context,
      vouchers,
      financialYear,
      assessmentYear,
      clientLegalName || 'Apex Technologies Pvt Ltd'
    );
    res.json({ success: true, data: artifact });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// ==========================================
// 19. GSTR-1 VS 3B VS BOOKS 3-WAY RECONCILIATION
// ==========================================
apiRouter.post('/gst/triple-rec/reconcile', extractFirmContext, (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const { data, financialYear, drc01bThreshold } = req.body;
    if (!data || !Array.isArray(data)) {
      return res.status(400).json({ success: false, error: 'data array is required' });
    }
    const result = gstTripleRecService.reconcile(
      context.firmId,
      context.clientId,
      data,
      financialYear,
      drc01bThreshold
    );
    res.json({ success: true, data: result });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

apiRouter.post('/gst/triple-rec/export', extractFirmContext, (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const { data, financialYear, drc01bThreshold, clientLegalName, clientGstin } = req.body;
    if (!data || !Array.isArray(data)) {
      return res.status(400).json({ success: false, error: 'data array is required' });
    }
    const artifact = gstTripleRecService.exportTripleRecWorkbook(
      context,
      data,
      financialYear,
      drc01bThreshold,
      clientLegalName || 'Apex Technologies Pvt Ltd',
      clientGstin || '27AABCU9603R1ZM'
    );
    res.json({ success: true, data: artifact });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// --- Section 206AB / 206CCA Non-Filer Verification Routes ---
apiRouter.post('/audit/section206ab/evaluate', extractFirmContext, (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const { records, clientName, financialYear } = req.body;
    if (!records || !Array.isArray(records)) {
      return res.status(400).json({ success: false, error: 'records array is required' });
    }
    const result = section206abService.evaluateRecords(
      records,
      clientName || context.firmName,
      financialYear || 'FY 2026-27'
    );
    res.json({ success: true, data: result });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

apiRouter.post('/audit/section206ab/export', extractFirmContext, (req: Request, res: Response) => {
  try {
    const context: FirmContext = (req as any).firmContext;
    const { records, clientName, financialYear } = req.body;
    if (!records || !Array.isArray(records)) {
      return res.status(400).json({ success: false, error: 'records array is required' });
    }
    const summary = section206abService.evaluateRecords(
      records,
      clientName || context.firmName,
      financialYear || 'FY 2026-27'
    );
    const artifact = section206abService.exportAuditArtifact(
      summary,
      context.firmId,
      context.clientId
    );
    res.json({ success: true, data: artifact });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// --- High-Throughput Streaming & Chunked Processing Routes ---
apiRouter.post('/pipeline/stream-process', extractFirmContext, async (req: Request, res: Response) => {
  try {
    const { base64, fileName, chunkSize } = req.body;
    if (!base64 || !fileName) {
      return res.status(400).json({ success: false, error: 'base64 buffer and fileName are required' });
    }
    const buffer = Buffer.from(base64, 'base64');
    const result = await StreamingProcessorService.processLargeBufferInChunks(
      buffer,
      fileName,
      chunkSize || 2500
    );
    res.json({ success: true, data: result });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// =========================================================================
// 1. TALLY PRIME LIVE SYNC (PORT 9000 XML OVER HTTP LOOPBACK)
// =========================================================================
apiRouter.post('/tally-live/health', async (_req: Request, res: Response) => {
  try {
    const status = await tallyLiveSyncService.checkHealth();
    res.json({ success: true, data: status });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

apiRouter.post('/tally-live/daybook', async (req: Request, res: Response) => {
  try {
    const { companyName, fromDate, toDate } = req.body;
    if (!companyName) {
      return res.status(400).json({ success: false, error: 'companyName is required' });
    }
    const result = await tallyLiveSyncService.fetchDaybook(companyName, fromDate, toDate);
    res.json({ success: result.success, data: result });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

apiRouter.post('/tally-live/trial-balance', async (req: Request, res: Response) => {
  try {
    const { companyName, fromDate, toDate } = req.body;
    if (!companyName) {
      return res.status(400).json({ success: false, error: 'companyName is required' });
    }
    const result = await tallyLiveSyncService.fetchTrialBalance(companyName, fromDate, toDate);
    res.json({ success: result.success, data: result });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

apiRouter.post('/tally-live/push-vouchers', async (req: Request, res: Response) => {
  try {
    const { companyName, vouchers } = req.body;
    if (!companyName || !vouchers) {
      return res.status(400).json({ success: false, error: 'companyName and vouchers array are required' });
    }
    const result = await tallyLiveSyncService.pushVouchers(companyName, vouchers);
    res.json({ success: result.success, data: result });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// =========================================================================
// 2. HOT-FOLDER ZERO-CLICK INGESTION WATCHER
// =========================================================================
apiRouter.get('/hot-folders/status', (_req: Request, res: Response) => {
  res.json({ success: true, data: hotFolderService.getStatus() });
});

apiRouter.post('/hot-folders/start', (_req: Request, res: Response) => {
  const result = hotFolderService.start();
  res.json({ success: result.success, data: result });
});

apiRouter.post('/hot-folders/stop', (_req: Request, res: Response) => {
  const result = hotFolderService.stop();
  res.json({ success: result.success, data: result });
});

apiRouter.post('/hot-folders/scan', (_req: Request, res: Response) => {
  hotFolderService.scanInbox();
  res.json({ success: true, data: hotFolderService.getStatus() });
});

// =========================================================================
// 3. MULTI-USER OFFICE LAN COLLABORATION & WORKPAPER LOCK
// =========================================================================
apiRouter.get('/lan/network-info', (req: Request, res: Response) => {
  const port = parseInt((req.query.port as string) || '5000') || 5000;
  const netInfo = officeLanService.getNetworkInfo(port);
  res.json({ success: true, data: netInfo });
});

apiRouter.post('/lan/lock-workpaper', extractFirmContext, (req: Request, res: Response) => {
  const { workpaperId, partnerName, partnerMrn, notes } = req.body;
  const context = (req as any).firmContext;
  if (!workpaperId || !partnerName || !partnerMrn) {
    return res.status(400).json({ success: false, error: 'workpaperId, partnerName, and partnerMrn are required' });
  }
  const result = officeLanService.lockWorkpaper(workpaperId, context.clientId, partnerName, partnerMrn, notes);
  res.json({ success: result.success, data: result });
});

apiRouter.post('/lan/unlock-workpaper', (req: Request, res: Response) => {
  const { workpaperId, partnerName } = req.body;
  if (!workpaperId || !partnerName) {
    return res.status(400).json({ success: false, error: 'workpaperId and partnerName are required' });
  }
  const result = officeLanService.unlockWorkpaper(workpaperId, partnerName);
  res.json({ success: result.success, data: result });
});

apiRouter.get('/lan/workpaper-status/:id', (req: Request, res: Response) => {
  const status = officeLanService.getWorkpaperStatus(req.params.id);
  res.json({ success: true, data: status });
});

// =========================================================================
// 4. INDEXED LOCAL DATABASE ENGINE (OFFLINE ACID STORAGE)
// =========================================================================
apiRouter.post('/local-db/insert-batch', extractFirmContext, (req: Request, res: Response) => {
  try {
    const { records } = req.body;
    const context = (req as any).firmContext;
    if (!Array.isArray(records)) {
      return res.status(400).json({ success: false, error: 'records must be an array' });
    }
    const recordsWithClient = records.map((r: any) => ({
      ...r,
      clientId: r.clientId || context.clientId,
    }));
    const result = localDbEngine.insertBatch(recordsWithClient);
    res.json({ success: true, data: result });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

apiRouter.post('/local-db/query', extractFirmContext, (req: Request, res: Response) => {
  try {
    const params = req.body || {};
    const context = (req as any).firmContext;
    if (!params.clientId) {
      params.clientId = context.clientId;
    }
    const result = localDbEngine.query(params);
    res.json({ success: true, data: result });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

apiRouter.get('/local-db/stats', (_req: Request, res: Response) => {
  res.json({ success: true, data: localDbEngine.getStats() });
});

// =========================================================================
// 5. CLIENT QUERY SHEET & AUDIT OBSERVATION GENERATOR
// =========================================================================
apiRouter.post('/client-queries/generate', extractFirmContext, (req: Request, res: Response) => {
  try {
    const { clientName, financialYear, assessmentYear, queryItems, clientEmail, firmName, firmFrn, partnerName, partnerMrn } = req.body;
    const context = (req as any).firmContext;

    if (!Array.isArray(queryItems) || queryItems.length === 0) {
      return res.status(400).json({ success: false, error: 'queryItems array is required' });
    }

    const pkg = {
      clientName: clientName || context.clientName || 'Client Entity',
      financialYear: financialYear || 'FY 2026-27',
      assessmentYear: assessmentYear || 'AY 2027-28',
      firmName: firmName || context.firmName || 'Chartered Accountants',
      firmFrn: firmFrn || '120488N',
      partnerName: partnerName || 'CA Rajesh Sharma, FCA',
      partnerMrn: partnerMrn || '054321',
      clientEmail,
      queryDate: new Date().toLocaleDateString('en-IN'),
      queryItems,
    };

    const artifacts = clientQueryService.generateQueryPackage(pkg);
    res.json({
      success: true,
      data: {
        totalQueries: artifacts.totalQueries,
        totalDisputedAmount: artifacts.totalDisputedAmount,
        htmlLetter: artifacts.htmlLetter,
        emlRawContent: artifacts.emlRawContent,
        excelBase64: artifacts.excelBuffer.toString('base64'),
      },
    });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// =========================================================================
// 6. NIGHTLY AUDIT CRON & MORNING EXCEPTION BRIEFING
// =========================================================================
apiRouter.post('/audit-cron/sweep', extractFirmContext, (req: Request, res: Response) => {
  try {
    const { clientName, financialYear, vouchers } = req.body;
    const context = (req as any).firmContext;

    if (!Array.isArray(vouchers) || vouchers.length === 0) {
      return res.status(400).json({ success: false, error: 'vouchers array is required' });
    }

    const report = nightlyAuditService.runAuditSweep(
      clientName || context.clientName || 'Client Entity',
      financialYear || 'FY 2026-27',
      vouchers
    );

    res.json({ success: true, data: report });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

apiRouter.get('/audit-cron/last-briefing', (_req: Request, res: Response) => {
  res.json({ success: true, data: nightlyAuditService.getLastBriefing() });
});


