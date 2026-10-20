/**
 * File-backed ACID Data Store with Firm-Level Isolation
 * Server Data Repository supporting all 7 Advanced Automation Modules
 */

import fs from 'fs';
import path from 'path';
import type {
  ProcessingJob,
  WorkflowDefinition,
  WorkflowRun,
  SourceFile,
  ExtractedRecord,
  RecordMatch,
  ExceptionRecord,
  ApprovalRecord,
  ExportArtifact,
  AuditEvent,
  AccountingRuleConfig,
  ClosingPeriod,
  ClosingTaskItem,
  AgingRecord,
  PaymentRiskAlert,
  AuditEvidence,
  EvidenceLink,
  AnomalyRule,
  AnomalyAlert,
  ClientDocumentRequest,
  RecurringTaskDefinition,
  TaskInstance,
  BatchRun,
} from '../types';

interface StoreSchema {
  jobs: Record<string, ProcessingJob>;
  workflowDefinitions: Record<string, WorkflowDefinition>;
  workflowRuns: Record<string, WorkflowRun>;
  sourceFiles: Record<string, SourceFile>;
  extractedRecords: Record<string, ExtractedRecord[]>; // keyed by sourceFileId
  recordMatches: Record<string, RecordMatch[]>; // keyed by workflowRunId
  exceptions: Record<string, ExceptionRecord>;
  approvals: Record<string, ApprovalRecord[]>; // keyed by workflowRunId
  artifacts: Record<string, ExportArtifact>;
  auditEvents: AuditEvent[];
  firmConfigs: Record<string, AccountingRuleConfig>;
  // Advanced Modules
  closingPeriods: Record<string, ClosingPeriod>;
  closingTasks: Record<string, ClosingTaskItem[]>; // keyed by periodId
  agingRecords: Record<string, AgingRecord[]>; // keyed by clientId
  paymentRiskAlerts: Record<string, PaymentRiskAlert[]>; // keyed by clientId
  auditEvidence: Record<string, AuditEvidence[]>; // keyed by clientId
  evidenceLinks: Record<string, EvidenceLink[]>; // keyed by evidenceId
  anomalyRules: Record<string, AnomalyRule[]>; // keyed by firmId
  anomalyAlerts: Record<string, AnomalyAlert[]>; // keyed by clientId
  clientDocumentRequests: Record<string, ClientDocumentRequest[]>; // keyed by clientId
  recurringTaskDefinitions: Record<string, RecurringTaskDefinition[]>; // keyed by firmId
  taskInstances: Record<string, TaskInstance[]>; // keyed by clientId
  batchRuns: Record<string, BatchRun>;
}

export class BackendStore {
  private static instance: BackendStore;
  private filePath: string;
  private data: StoreSchema;
  private writeLock: boolean = false;

  private constructor(customPath?: string) {
    this.filePath = customPath || path.resolve(process.cwd(), 'data', 'ca-backend-store.json');
    this.data = this.initializeData();
    this.load();
    this.seedWorkflowDefinitions();
    this.seedDefaults();
  }

  public static getInstance(customPath?: string): BackendStore {
    if (!BackendStore.instance) {
      BackendStore.instance = new BackendStore(customPath);
    }
    return BackendStore.instance;
  }

  private initializeData(): StoreSchema {
    return {
      jobs: {},
      workflowDefinitions: {},
      workflowRuns: {},
      sourceFiles: {},
      extractedRecords: {},
      recordMatches: {},
      exceptions: {},
      approvals: {},
      artifacts: {},
      auditEvents: [],
      firmConfigs: {},
      closingPeriods: {},
      closingTasks: {},
      agingRecords: {},
      paymentRiskAlerts: {},
      auditEvidence: {},
      evidenceLinks: {},
      anomalyRules: {},
      anomalyAlerts: {},
      clientDocumentRequests: {},
      recurringTaskDefinitions: {},
      taskInstances: {},
      batchRuns: {},
    };
  }

  private load(): void {
    try {
      if (fs.existsSync(this.filePath)) {
        const raw = fs.readFileSync(this.filePath, 'utf-8');
        const parsed = JSON.parse(raw);
        this.data = { ...this.initializeData(), ...parsed };
      } else {
        this.save();
      }
    } catch {
      this.data = this.initializeData();
    }
  }

  private transactionSnapshot: string | null = null;

  public beginTransaction(): void {
    this.transactionSnapshot = JSON.stringify(this.data);
  }

  public commitTransaction(): void {
    this.transactionSnapshot = null;
    this.save();
  }

  public rollbackTransaction(): void {
    if (this.transactionSnapshot) {
      this.data = JSON.parse(this.transactionSnapshot);
      this.transactionSnapshot = null;
      this.save();
    }
  }

  public createBackup(): string {
    const backupDir = path.resolve(path.dirname(this.filePath), 'backups');
    if (!fs.existsSync(backupDir)) {
      fs.mkdirSync(backupDir, { recursive: true });
    }
    const backupFile = path.resolve(backupDir, `ca-store-backup-${Date.now()}.json`);
    fs.writeFileSync(backupFile, JSON.stringify(this.data, null, 2), 'utf-8');

    // Keep only last 5 backups
    try {
      const files = fs.readdirSync(backupDir)
        .filter((f) => f.startsWith('ca-store-backup-') && f.endsWith('.json'))
        .sort()
        .reverse();
      if (files.length > 5) {
        for (let i = 5; i < files.length; i++) {
          fs.unlinkSync(path.resolve(backupDir, files[i]));
        }
      }
    } catch {
      // Ignore cleanup error
    }

    return backupFile;
  }

  private pendingSave: boolean = false;

  private save(): void {
    if (this.writeLock) {
      this.pendingSave = true;
      return;
    }
    this.writeLock = true;
    try {
      const dir = path.dirname(this.filePath);
      if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
      }
      fs.writeFileSync(this.filePath, JSON.stringify(this.data, null, 2), 'utf-8');
    } catch (err) {
      console.error('[BackendStore] Failed to persist data:', err);
    } finally {
      this.writeLock = false;
      if (this.pendingSave) {
        this.pendingSave = false;
        this.save();
      }
    }
  }

  // --- Seed Default Workflow Definitions ---
  private seedWorkflowDefinitions(): void {
    const definitions: WorkflowDefinition[] = [
      {
        id: 'wf_bank_rec',
        name: 'Bank Statement-to-Ledger Reconciliation & Journal Prep',
        version: '1.0.0',
        description: 'End-to-end bank statement parsing, continuity audit, ledger comparison, rule-based head suggestions, and draft double-entry journals.',
        category: 'bank_statement',
        requiresApprovalBeforeCompletion: true,
        retryPolicy: { maxAttempts: 3, backoffMs: 1000 },
        steps: [
          { stepId: 'step_validate_files', name: 'File Validation & Format Check', description: 'Validate bank statement and ledger structures', required: true, dependencies: [], action: 'validate_file' },
          { stepId: 'step_extract_txns', name: 'Transaction Extraction & Normalization', description: 'Extract deposits, withdrawals, dates, and calculate balance continuity', required: true, dependencies: ['step_validate_files'], action: 'extract_data' },
          { stepId: 'step_validate_continuity', name: 'Running Balance Continuity Audit', description: 'Verify mathematical integrity of bank running balances', required: true, dependencies: ['step_extract_txns'], action: 'validate_rules' },
          { stepId: 'step_reconcile_ledger', name: 'Ledger Matching Engine', description: 'Match statement against ledger entries with configurable tolerances', required: true, dependencies: ['step_validate_continuity'], action: 'execute_match' },
          { stepId: 'step_suggest_journals', name: 'Account Head & Draft Journal Generation', description: 'Suggest account heads and draft double-entry vouchers for unmatched items', required: true, dependencies: ['step_reconcile_ledger'], action: 'generate_drafts' },
          { stepId: 'step_export_reports', name: 'BRS Workbook & Audit Export', description: 'Export multi-sheet BRS reconciliation and journal registers', required: true, dependencies: ['step_suggest_journals'], action: 'export_reports' },
        ],
      },
      {
        id: 'wf_invoice_proc',
        name: 'Bulk Invoice-to-Excel & Duplicate Auditor',
        version: '1.0.0',
        description: 'Batch invoice extraction, statutory arithmetic verification, customer GSTIN/POS extraction, duplicate detection, and register generation.',
        category: 'invoice_processing',
        requiresApprovalBeforeCompletion: false,
        retryPolicy: { maxAttempts: 3, backoffMs: 1000 },
        steps: [
          { stepId: 'step_validate_invoices', name: 'Invoice Ingestion & Format Check', description: 'Verify file types and readability', required: true, dependencies: [], action: 'validate_file' },
          { stepId: 'step_extract_fields', name: 'Tax & Line-Item Extraction', description: 'Extract supplier, customer, GSTIN, taxable amount, taxes, cess', required: true, dependencies: ['step_validate_invoices'], action: 'extract_data' },
          { stepId: 'step_arithmetic_audit', name: 'Statutory Arithmetic Verification', description: 'Audit Taxable + CGST + SGST + IGST + Cess == Total', required: true, dependencies: ['step_extract_fields'], action: 'validate_rules' },
          { stepId: 'step_detect_duplicates', name: 'Batch Duplicate Detection', description: 'Identify duplicate invoices by GSTIN, number, and value', required: true, dependencies: ['step_arithmetic_audit'], action: 'detect_exceptions' },
          { stepId: 'step_export_register', name: 'Consolidated Register & Exceptions Export', description: 'Export 4-sheet formatted Excel invoice register', required: true, dependencies: ['step_detect_duplicates'], action: 'export_reports' },
        ],
      },
      {
        id: 'wf_gst_rec',
        name: 'GST Purchase Register vs GSTR-2B Reconciliation',
        version: '1.0.0',
        description: 'Multi-stream invoice comparison, alphanumeric normalization, Section 16(2)(aa) compliance notice, and difference reporting.',
        category: 'gst_reconciliation',
        requiresApprovalBeforeCompletion: true,
        retryPolicy: { maxAttempts: 3, backoffMs: 1000 },
        steps: [
          { stepId: 'step_validate_gst_inputs', name: 'Validate Purchase Register & 2B Data', description: 'Check portal export and books register headers', required: true, dependencies: [], action: 'validate_file' },
          { stepId: 'step_normalize_invoices', name: 'Normalize Invoice Numbers & GSTINs', description: 'Strip punctuation, spaces, and leading zeros for fuzzy matching', required: true, dependencies: ['step_validate_gst_inputs'], action: 'normalize' },
          { stepId: 'step_match_gst', name: 'Multi-Level Matching Engine', description: 'Identify exact matches, tax mismatches, missing in 2B, and missing in books', required: true, dependencies: ['step_normalize_invoices'], action: 'execute_match' },
          { stepId: 'step_flag_itc_risks', name: 'Section 16(2)(aa) Risk Classification', description: 'Segregate eligible vs at-risk ITC and formulate vendor follow-up actions', required: true, dependencies: ['step_match_gst'], action: 'detect_exceptions' },
          { stepId: 'step_export_gst_workbook', name: 'Reconciliation Workbook Export', description: 'Generate 6-sheet reconciliation register', required: true, dependencies: ['step_flag_itc_risks'], action: 'export_reports' },
        ],
      },
      {
        id: 'wf_tb_audit',
        name: 'Trial Balance-to-Schedule III & Audit Scrutiny',
        version: '1.0.0',
        description: 'Trial balance balance verification, Companies Act Schedule III statement preparation, abnormal balance audit, and adjustments recalculation.',
        category: 'trial_balance_audit',
        requiresApprovalBeforeCompletion: true,
        retryPolicy: { maxAttempts: 3, backoffMs: 1000 },
        steps: [
          { stepId: 'step_validate_tb', name: 'Trial Balance Ingestion & Balance Check', description: 'Verify equality of debits and credits', required: true, dependencies: [], action: 'validate_file' },
          { stepId: 'step_schedule_iii_mapping', name: 'Schedule III Group Classification', description: 'Map accounts to PPE, Inventory, Receivables, Payables, Revenue, Costs', required: true, dependencies: ['step_validate_tb'], action: 'normalize' },
          { stepId: 'step_audit_scrutiny', name: 'Abnormal Balances & Scrutiny Rules', description: 'Flag negative cash, credit debtors, debit creditors, Section 40A(3) violations', required: true, dependencies: ['step_schedule_iii_mapping'], action: 'validate_rules' },
          { stepId: 'step_apply_adjustments', name: 'Dynamic Audit Adjustment Processing', description: 'Apply approved closing stock, depreciation, and accrual adjustments', required: true, dependencies: ['step_audit_scrutiny'], action: 'generate_drafts' },
          { stepId: 'step_export_financials', name: 'Draft Financial Statements Export', description: 'Export 5-sheet financial statements and audit exceptions log', required: true, dependencies: ['step_apply_adjustments'], action: 'export_reports' },
        ],
      },
    ];

    for (const def of definitions) {
      this.data.workflowDefinitions[def.id] = def;
    }
    this.save();
  }

  // --- Seed Default Anomaly Rules & Compliance Schedules ---
  private seedDefaults(): void {
    const defaultFirm = 'firm_apex_ca';
    if (!this.data.anomalyRules[defaultFirm]) {
      this.data.anomalyRules[defaultFirm] = [
        { id: 'ar-1', firmId: defaultFirm, ruleCode: 'ANOMALY_DUPLICATE_JOURNAL', name: 'Duplicate Journal Entry', description: 'Flags identical debits, credits, and accounts posted within close proximity', category: 'journal_duplicate', severity: 'critical', enabled: true },
        { id: 'ar-2', firmId: defaultFirm, ruleCode: 'ANOMALY_ROUND_SUM_EXPENSE', name: 'Unusual Round Sum Expense', description: 'Flags round sum expenses (e.g. ₹50,000, ₹1,00,000) lacking breakdown', category: 'round_sum', severity: 'warning', thresholdValue: 50000, enabled: true },
        { id: 'ar-3', firmId: defaultFirm, ruleCode: 'ANOMALY_40A3_CASH', name: 'Section 40A(3) Cash Payment Violation', description: 'Flags single cash payments exceeding statutory limit of ₹10,000', category: 'cash_limit', severity: 'critical', thresholdValue: 10000, enabled: true },
        { id: 'ar-4', firmId: defaultFirm, ruleCode: 'ANOMALY_WEEKEND_POSTING', name: 'Weekend / Off-Hours Voucher', description: 'Flags significant journal vouchers posted on Sundays or holidays', category: 'weekend_posting', severity: 'info', thresholdValue: 100000, enabled: true },
        { id: 'ar-5', firmId: defaultFirm, ruleCode: 'ANOMALY_POLARITY_INVERSION', name: 'Abnormal Head Polarity', description: 'Flags credit cash, debit sundry creditors, or credit sundry debtors', category: 'abnormal_polarity', severity: 'critical', enabled: true },
        { id: 'ar-6', firmId: defaultFirm, ruleCode: 'ANOMALY_VELOCITY_SPIKE', name: 'Expense Velocity Surge', description: 'Flags sudden expense category increase > 200% above period average', category: 'velocity_spike', severity: 'warning', thresholdValue: 200, enabled: true },
      ];
    }

    if (!this.data.recurringTaskDefinitions[defaultFirm]) {
      this.data.recurringTaskDefinitions[defaultFirm] = [
        { id: 'rec-1', firmId: defaultFirm, title: 'TDS / TCS Deposit Challan ITNS-281', statutoryAct: 'Income Tax', frequency: 'monthly', dueDayOfMonth: 7, description: 'Deposit tax deducted under Section 200(3) of Income Tax Act for previous month', associatedWorkflowId: 'wf_bank_rec', isStatutory: true },
        { id: 'rec-2', firmId: defaultFirm, title: 'GSTR-1 Monthly Outward Supplies', statutoryAct: 'GST', frequency: 'monthly', dueDayOfMonth: 11, description: 'File details of outward supplies for registered regular taxpayers', associatedWorkflowId: 'wf_invoice_proc', isStatutory: true },
        { id: 'rec-3', firmId: defaultFirm, title: 'GSTR-3B Summary Return & Tax Settlement', statutoryAct: 'GST', frequency: 'monthly', dueDayOfMonth: 20, description: 'Reconcile ITC and pay monthly net GST liability', associatedWorkflowId: 'wf_gst_rec', isStatutory: true },
        { id: 'rec-4', firmId: defaultFirm, title: 'Monthly Bank Reconciliation & Continuity Review', statutoryAct: 'Internal Audit', frequency: 'monthly', dueDayOfMonth: 5, description: 'Verify zero balance discontinuity across all client operational bank accounts', associatedWorkflowId: 'wf_bank_rec', isStatutory: false },
        { id: 'rec-5', firmId: defaultFirm, title: 'Advance Tax Instalment under Section 211', statutoryAct: 'Income Tax', frequency: 'quarterly', dueDayOfMonth: 15, description: 'Calculate and deposit advance tax (15% Jun, 45% Sep, 75% Dec, 100% Mar)', isStatutory: true },
      ];
    }
    this.save();
  }

  // --- Job Operations ---
  public createJob(job: ProcessingJob): ProcessingJob {
    this.data.jobs[job.id] = job;
    this.save();
    return job;
  }

  public getJob(firmId: string, jobId: string): ProcessingJob | undefined {
    const job = this.data.jobs[jobId];
    if (job && job.firmId === firmId) return job;
    return undefined;
  }

  public updateJob(firmId: string, jobId: string, patch: Partial<ProcessingJob>): ProcessingJob | undefined {
    const job = this.getJob(firmId, jobId);
    if (!job) return undefined;
    const updated = { ...job, ...patch, updatedAt: new Date().toISOString() };
    this.data.jobs[jobId] = updated;
    this.save();
    return updated;
  }

  public listJobs(firmId: string, clientId?: string): ProcessingJob[] {
    return Object.values(this.data.jobs).filter((j) => {
      if (j.firmId !== firmId) return false;
      if (clientId && j.clientId !== clientId) return false;
      return true;
    }).sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());
  }

  public findJobByIdempotencyKey(firmId: string, idempotencyKey: string): ProcessingJob | undefined {
    return Object.values(this.data.jobs).find((j) => j.firmId === firmId && j.idempotencyKey === idempotencyKey);
  }

  // --- Workflow Definitions ---
  public getWorkflowDefinition(id: string): WorkflowDefinition | undefined {
    return this.data.workflowDefinitions[id];
  }

  public listWorkflowDefinitions(): WorkflowDefinition[] {
    return Object.values(this.data.workflowDefinitions);
  }

  // --- Workflow Runs ---
  public createWorkflowRun(run: WorkflowRun): WorkflowRun {
    this.data.workflowRuns[run.id] = run;
    this.save();
    return run;
  }

  public getWorkflowRun(firmId: string, runId: string): WorkflowRun | undefined {
    const run = this.data.workflowRuns[runId];
    if (run && run.firmId === firmId) return run;
    return undefined;
  }

  public updateWorkflowRun(firmId: string, runId: string, patch: Partial<WorkflowRun>): WorkflowRun | undefined {
    const run = this.getWorkflowRun(firmId, runId);
    if (!run) return undefined;
    const updated = { ...run, ...patch, updatedAt: new Date().toISOString() };
    this.data.workflowRuns[runId] = updated;
    this.save();
    return updated;
  }

  // --- Source Files ---
  public saveSourceFile(file: SourceFile): SourceFile {
    this.data.sourceFiles[file.id] = file;
    this.save();
    return file;
  }

  public getSourceFile(firmId: string, fileId: string): SourceFile | undefined {
    const file = this.data.sourceFiles[fileId];
    if (file && file.firmId === firmId) return file;
    return undefined;
  }

  public listSourceFiles(firmId: string, clientId?: string): SourceFile[] {
    return Object.values(this.data.sourceFiles).filter((f) => {
      if (f.firmId !== firmId) return false;
      if (clientId && f.clientId !== clientId) return false;
      return true;
    });
  }

  // --- Extracted Records ---
  public saveExtractedRecords(sourceFileId: string, records: ExtractedRecord[]): void {
    this.data.extractedRecords[sourceFileId] = records;
    this.save();
  }

  public getExtractedRecords(sourceFileId: string): ExtractedRecord[] {
    return this.data.extractedRecords[sourceFileId] || [];
  }

  // --- Record Matches ---
  public saveRecordMatches(workflowRunId: string, matches: RecordMatch[]): void {
    this.data.recordMatches[workflowRunId] = matches;
    this.save();
  }

  public getRecordMatches(workflowRunId: string): RecordMatch[] {
    return this.data.recordMatches[workflowRunId] || [];
  }

  public updateRecordMatch(workflowRunId: string, matchId: string, patch: Partial<RecordMatch>): RecordMatch | undefined {
    const matches = this.data.recordMatches[workflowRunId];
    if (!matches) return undefined;
    const idx = matches.findIndex((m) => m.id === matchId);
    if (idx === -1) return undefined;
    matches[idx] = { ...matches[idx], ...patch };
    this.save();
    return matches[idx];
  }

  // --- Exceptions ---
  public createException(exception: ExceptionRecord): ExceptionRecord {
    this.data.exceptions[exception.id] = exception;
    this.save();
    return exception;
  }

  public getException(firmId: string, exceptionId: string): ExceptionRecord | undefined {
    const exc = this.data.exceptions[exceptionId];
    if (exc && exc.firmId === firmId) return exc;
    return undefined;
  }

  public updateException(firmId: string, exceptionId: string, patch: Partial<ExceptionRecord>): ExceptionRecord | undefined {
    const exc = this.getException(firmId, exceptionId);
    if (!exc) return undefined;
    const updated = { ...exc, ...patch, updatedAt: new Date().toISOString() };
    this.data.exceptions[exceptionId] = updated;
    this.save();
    return updated;
  }

  public listExceptions(firmId: string, filters?: { clientId?: string; workflowRunId?: string; status?: string }): ExceptionRecord[] {
    return Object.values(this.data.exceptions).filter((e) => {
      if (e.firmId !== firmId) return false;
      if (filters?.clientId && e.clientId !== filters.clientId) return false;
      if (filters?.workflowRunId && e.workflowRunId !== filters.workflowRunId) return false;
      if (filters?.status && e.resolutionStatus !== filters.status) return false;
      return true;
    }).sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());
  }

  // --- Approvals ---
  public recordApproval(approval: ApprovalRecord): ApprovalRecord {
    if (!this.data.approvals[approval.workflowRunId]) {
      this.data.approvals[approval.workflowRunId] = [];
    }
    this.data.approvals[approval.workflowRunId].push(approval);
    this.save();
    return approval;
  }

  public listApprovals(firmId: string, workflowRunId?: string): ApprovalRecord[] {
    if (workflowRunId) {
      return (this.data.approvals[workflowRunId] || []).filter((a) => a.firmId === firmId);
    }
    const all: ApprovalRecord[] = [];
    for (const list of Object.values(this.data.approvals)) {
      for (const item of list) {
        if (item.firmId === firmId) all.push(item);
      }
    }
    return all.sort((a, b) => new Date(b.performedAt).getTime() - new Date(a.performedAt).getTime());
  }

  // --- Export Artifacts ---
  public saveArtifact(artifact: ExportArtifact): ExportArtifact {
    this.data.artifacts[artifact.id] = artifact;
    this.save();
    return artifact;
  }

  public getArtifact(firmId: string, artifactId: string): ExportArtifact | undefined {
    const art = this.data.artifacts[artifactId];
    if (art && art.firmId === firmId) return art;
    return undefined;
  }

  public getArtifactById(artifactId: string): ExportArtifact | undefined {
    return this.data.artifacts[artifactId];
  }

  public listArtifacts(firmId: string, workflowRunId?: string): ExportArtifact[] {
    return Object.values(this.data.artifacts).filter((a) => {
      if (a.firmId !== firmId) return false;
      if (workflowRunId && a.workflowRunId !== workflowRunId) return false;
      return true;
    }).sort((a, b) => new Date(b.generatedAt).getTime() - new Date(a.generatedAt).getTime());
  }

  // --- Audit Events ---
  public logAuditEvent(event: AuditEvent): void {
    this.data.auditEvents.push(event);
    this.save();
  }

  public queryAuditEvents(firmId: string, filters?: { clientId?: string; entityId?: string; entityType?: string }): AuditEvent[] {
    return this.data.auditEvents.filter((ev) => {
      if (ev.firmId !== firmId) return false;
      if (filters?.clientId && ev.clientId !== filters.clientId) return false;
      if (filters?.entityId && ev.entityId !== filters.entityId) return false;
      if (filters?.entityType && ev.entityType !== filters.entityType) return false;
      return true;
    }).sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
  }

  public getAuditEvents(firmId: string, clientId?: string): AuditEvent[] {
    return this.queryAuditEvents(firmId, clientId ? { clientId } : undefined);
  }

  // --- Firm Configs ---
  public getFirmConfig(firmId: string): AccountingRuleConfig {
    if (!this.data.firmConfigs[firmId]) {
      this.data.firmConfigs[firmId] = {
        firmId,
        maxCashVoucherAmount: 10000,
        gstTolerance: 2.0,
        allowNegativeCash: false,
        requireBalancedJournals: true,
        enforceBalanceContinuity: true,
      };
      this.save();
    }
    return this.data.firmConfigs[firmId];
  }

  public updateFirmConfig(firmId: string, patch: Partial<AccountingRuleConfig>): AccountingRuleConfig {
    const current = this.getFirmConfig(firmId);
    const updated = { ...current, ...patch };
    this.data.firmConfigs[firmId] = updated;
    this.save();
    return updated;
  }

  // ==========================================
  // MODULE 1: MONTH-END CLOSING
  // ==========================================
  public createClosingPeriod(period: ClosingPeriod): ClosingPeriod {
    this.data.closingPeriods[period.id] = period;
    this.save();
    return period;
  }

  public getClosingPeriod(firmId: string, periodId: string): ClosingPeriod | undefined {
    const p = this.data.closingPeriods[periodId];
    if (p && p.firmId === firmId) return p;
    return undefined;
  }

  public listClosingPeriods(firmId: string, clientId?: string): ClosingPeriod[] {
    return Object.values(this.data.closingPeriods).filter((p) => {
      if (p.firmId !== firmId) return false;
      if (clientId && p.clientId !== clientId) return false;
      return true;
    }).sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());
  }

  public updateClosingPeriod(firmId: string, periodId: string, patch: Partial<ClosingPeriod>): ClosingPeriod | undefined {
    const p = this.getClosingPeriod(firmId, periodId);
    if (!p) return undefined;
    const updated = { ...p, ...patch, updatedAt: new Date().toISOString() };
    this.data.closingPeriods[periodId] = updated;
    this.save();
    return updated;
  }

  public setClosingTasks(periodId: string, tasks: ClosingTaskItem[]): void {
    this.data.closingTasks[periodId] = tasks;
    this.save();
  }

  public getClosingTasks(periodId: string): ClosingTaskItem[] {
    return this.data.closingTasks[periodId] || [];
  }

  public updateClosingTask(periodId: string, taskId: string, patch: Partial<ClosingTaskItem>): ClosingTaskItem | undefined {
    const list = this.data.closingTasks[periodId];
    if (!list) return undefined;
    const idx = list.findIndex((t) => t.id === taskId);
    if (idx === -1) return undefined;
    list[idx] = { ...list[idx], ...patch };
    this.save();
    return list[idx];
  }

  // ==========================================
  // MODULE 2: AP & AR AUTOMATION
  // ==========================================
  public saveAgingRecords(clientId: string, records: AgingRecord[], firmId?: string): void {
    const effectiveFirmId = firmId || records[0]?.firmId;
    if (effectiveFirmId) {
      const existing = (this.data.agingRecords[clientId] || []).filter((r) => r.firmId !== effectiveFirmId);
      this.data.agingRecords[clientId] = [...existing, ...records];
    } else {
      this.data.agingRecords[clientId] = records;
    }
    this.save();
  }

  public getAgingRecords(clientId: string, firmId?: string): AgingRecord[] {
    const list = this.data.agingRecords[clientId] || [];
    return firmId ? list.filter((r) => r.firmId === firmId) : list;
  }

  public savePaymentRiskAlerts(clientId: string, alerts: PaymentRiskAlert[], firmId?: string): void {
    const effectiveFirmId = firmId || alerts[0]?.firmId;
    if (effectiveFirmId) {
      const existing = (this.data.paymentRiskAlerts[clientId] || []).filter((a) => a.firmId !== effectiveFirmId);
      this.data.paymentRiskAlerts[clientId] = [...existing, ...alerts];
    } else {
      this.data.paymentRiskAlerts[clientId] = alerts;
    }
    this.save();
  }

  public getPaymentRiskAlerts(clientId: string, firmId?: string): PaymentRiskAlert[] {
    const list = this.data.paymentRiskAlerts[clientId] || [];
    return firmId ? list.filter((a) => a.firmId === firmId) : list;
  }

  // ==========================================
  // MODULE 3: AUDIT EVIDENCE AUTOMATION
  // ==========================================
  public saveAuditEvidence(clientId: string, evidenceList: AuditEvidence[], firmId?: string): void {
    const effectiveFirmId = firmId || evidenceList[0]?.firmId;
    if (effectiveFirmId) {
      const existing = (this.data.auditEvidence[clientId] || []).filter((e) => e.firmId !== effectiveFirmId);
      this.data.auditEvidence[clientId] = [...existing, ...evidenceList];
    } else {
      this.data.auditEvidence[clientId] = evidenceList;
    }
    this.save();
  }

  public getAuditEvidence(clientId: string, firmId?: string): AuditEvidence[] {
    const list = this.data.auditEvidence[clientId] || [];
    return firmId ? list.filter((e) => e.firmId === firmId) : list;
  }

  public saveEvidenceLink(link: EvidenceLink): void {
    if (!this.data.evidenceLinks[link.evidenceId]) {
      this.data.evidenceLinks[link.evidenceId] = [];
    }
    this.data.evidenceLinks[link.evidenceId].push(link);
    this.save();
  }

  public getEvidenceLinks(evidenceId: string): EvidenceLink[] {
    return this.data.evidenceLinks[evidenceId] || [];
  }

  // ==========================================
  // MODULE 4: ANOMALY DETECTION ENGINE
  // ==========================================
  public listAnomalyRules(firmId: string): AnomalyRule[] {
    return this.data.anomalyRules[firmId] || [];
  }

  public updateAnomalyRule(firmId: string, ruleId: string, patch: Partial<AnomalyRule>): AnomalyRule | undefined {
    const rules = this.data.anomalyRules[firmId] || [];
    const idx = rules.findIndex((r) => r.id === ruleId);
    if (idx === -1) return undefined;
    rules[idx] = { ...rules[idx], ...patch };
    this.save();
    return rules[idx];
  }

  public saveAnomalyAlerts(clientId: string, alerts: AnomalyAlert[], firmId?: string): void {
    const effectiveFirmId = firmId || alerts[0]?.firmId;
    if (effectiveFirmId) {
      const existing = (this.data.anomalyAlerts[clientId] || []).filter((a) => a.firmId !== effectiveFirmId);
      this.data.anomalyAlerts[clientId] = [...existing, ...alerts];
    } else {
      this.data.anomalyAlerts[clientId] = alerts;
    }
    this.save();
  }

  public listAnomalyAlerts(clientId: string, firmId?: string): AnomalyAlert[] {
    const list = this.data.anomalyAlerts[clientId] || [];
    return firmId ? list.filter((a) => a.firmId === firmId) : list;
  }

  public updateAnomalyAlert(clientId: string, alertId: string, patch: Partial<AnomalyAlert>): AnomalyAlert | undefined {
    const list = this.data.anomalyAlerts[clientId];
    if (!list) return undefined;
    const idx = list.findIndex((a) => a.id === alertId);
    if (idx === -1) return undefined;
    list[idx] = { ...list[idx], ...patch };
    this.save();
    return list[idx];
  }

  // ==========================================
  // MODULE 5: CLIENT DOCUMENT COLLECTION
  // ==========================================
  public createDocumentRequest(req: ClientDocumentRequest): ClientDocumentRequest {
    if (!this.data.clientDocumentRequests[req.clientId]) {
      this.data.clientDocumentRequests[req.clientId] = [];
    }
    this.data.clientDocumentRequests[req.clientId].push(req);
    this.save();
    return req;
  }

  public listDocumentRequests(clientId: string, firmId?: string): ClientDocumentRequest[] {
    const list = this.data.clientDocumentRequests[clientId] || [];
    return firmId ? list.filter((r) => r.firmId === firmId) : list;
  }

  public updateDocumentRequest(clientId: string, requestId: string, patch: Partial<ClientDocumentRequest>): ClientDocumentRequest | undefined {
    const list = this.data.clientDocumentRequests[clientId];
    if (!list) return undefined;
    const idx = list.findIndex((r) => r.id === requestId);
    if (idx === -1) return undefined;
    list[idx] = { ...list[idx], ...patch, updatedAt: new Date().toISOString() };
    this.save();
    return list[idx];
  }

  // ==========================================
  // MODULE 6: RECURRING COMPLIANCE & TASKS
  // ==========================================
  public listRecurringTaskDefinitions(firmId: string): RecurringTaskDefinition[] {
    return this.data.recurringTaskDefinitions[firmId] || [];
  }

  public saveTaskInstances(clientId: string, instances: TaskInstance[], firmId?: string): void {
    const effectiveFirmId = firmId || instances[0]?.firmId;
    if (effectiveFirmId) {
      const existing = (this.data.taskInstances[clientId] || []).filter((t) => t.firmId !== effectiveFirmId);
      this.data.taskInstances[clientId] = [...existing, ...instances];
    } else {
      this.data.taskInstances[clientId] = instances;
    }
    this.save();
  }

  public listTaskInstances(clientId: string, firmId?: string): TaskInstance[] {
    const list = this.data.taskInstances[clientId] || [];
    return firmId ? list.filter((t) => t.firmId === firmId) : list;
  }

  public updateTaskInstance(clientId: string, instanceId: string, patch: Partial<TaskInstance>): TaskInstance | undefined {
    const list = this.data.taskInstances[clientId];
    if (!list) return undefined;
    const idx = list.findIndex((t) => t.id === instanceId);
    if (idx === -1) return undefined;
    list[idx] = { ...list[idx], ...patch };
    this.save();
    return list[idx];
  }

  // ==========================================
  // MODULE 7: MULTI-CLIENT BATCH RUNS
  // ==========================================
  public createBatchRun(run: BatchRun): BatchRun {
    this.data.batchRuns[run.id] = run;
    this.save();
    return run;
  }

  public getBatchRun(firmId: string, batchId: string): BatchRun | undefined {
    const b = this.data.batchRuns[batchId];
    if (b && b.firmId === firmId) return b;
    return undefined;
  }

  public listBatchRuns(firmId: string): BatchRun[] {
    return Object.values(this.data.batchRuns)
      .filter((b) => b.firmId === firmId)
      .sort((a, b) => new Date(b.startedAt).getTime() - new Date(a.startedAt).getTime());
  }

  public updateBatchRun(firmId: string, batchId: string, patch: Partial<BatchRun>): BatchRun | undefined {
    const b = this.getBatchRun(firmId, batchId);
    if (!b) return undefined;
    const updated = { ...b, ...patch };
    this.data.batchRuns[batchId] = updated;
    this.save();
    return updated;
  }

  // For testing cleanup
  public clearAll(): void {
    this.data = this.initializeData();
    this.seedWorkflowDefinitions();
    this.seedDefaults();
    this.save();
  }
}
