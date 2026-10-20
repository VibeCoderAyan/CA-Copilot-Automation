/**
 * Intelligent Backend Automation Engine - Core Data Models & Schemas
 * CA Tools Web Application
 */

export type JobStatus = 'queued' | 'running' | 'review_required' | 'completed' | 'failed' | 'cancelled';
export type StepStatus = 'pending' | 'running' | 'completed' | 'failed' | 'skipped';
export type ExceptionSeverity = 'info' | 'warning' | 'error' | 'critical';
export type ExceptionResolution = 'pending' | 'approved' | 'rejected' | 'corrected' | 'investigating';
export type MatchType = 'exact' | 'probable' | 'partial' | 'unmatched' | 'duplicate_candidate';

export interface FirmContext {
  firmId: string;
  firmName: string;
  userId: string;
  userRole: 'Partner' | 'Senior' | 'Article' | 'Auditor' | 'System';
  clientId: string;
  clientName?: string;
}

export interface SourceFile {
  id: string;
  firmId: string;
  clientId: string;
  fileName: string;
  fileSizeBytes: number;
  mimeType: string;
  fileHash: string; // SHA-256 for dedup and integrity
  detectedFormat: 'csv' | 'xlsx' | 'xls' | 'pdf' | 'json' | 'unknown';
  documentType: 'bank_statement' | 'sales_invoice' | 'purchase_invoice' | 'purchase_register' | 'gstr2b' | 'trial_balance' | 'ledger' | 'unknown';
  uploadedAt: string;
  storagePath?: string;
  contentBase64?: string; // Optional embedded content for lightweight storage
  metadata?: Record<string, any>;
}

export interface ExtractedRecord {
  id: string;
  sourceFileId: string;
  rowIndex: number;
  pageNumber?: number;
  rawFields: Record<string, any>;
  normalizedFields: {
    date?: string;
    amount?: number;
    debit?: number;
    credit?: number;
    narration?: string;
    refNo?: string;
    chqNo?: string;
    balance?: number;
    gstin?: string;
    partyName?: string;
    invoiceNo?: string;
    taxableAmount?: number;
    cgst?: number;
    sgst?: number;
    igst?: number;
    cess?: number;
    totalAmount?: number;
    withdrawal?: number;
    deposit?: number;
    accountCode?: string;
    accountName?: string;
    [key: string]: any;
  };
  confidence?: number; // 0 to 100
  isSuspect?: boolean;
  warnings?: string[];
  status?: string;
}

export interface ValidationIssue {
  id: string;
  ruleCode: string;
  ruleName: string;
  severity: ExceptionSeverity;
  message: string;
  field?: string;
  actualValue?: any;
  expectedValue?: any;
  recordId?: string;
  sourceFileId?: string;
  suggestedCorrection?: any;
}

export interface RecordMatch {
  id: string;
  workflowRunId: string;
  matchType: MatchType;
  matchScore: number; // 0 to 100
  explanation: string;
  primaryRecord: ExtractedRecord;
  matchedRecords: ExtractedRecord[];
  isConfirmed: boolean;
  varianceAmount?: number;
  varianceDays?: number;
  resolutionStatus?: 'auto_matched' | 'user_confirmed' | 'user_rejected' | 'force_matched';
}

export interface ExceptionRecord {
  id: string;
  firmId: string;
  clientId: string;
  workflowRunId: string;
  jobId: string;
  sourceFileId?: string;
  recordId?: string;
  category: 'unmatched_transaction' | 'tax_difference' | 'missing_in_portal' | 'missing_in_books' | 'duplicate_invoice' | 'abnormal_balance' | 'discontinuity' | 'statutory_violation' | 'format_error';
  severity: ExceptionSeverity;
  title: string;
  description: string;
  originalValues: Record<string, any>;
  correctedValues?: Record<string, any>;
  suggestedResolution?: string;
  resolutionStatus: ExceptionResolution;
  assignedReviewer?: string;
  reviewedBy?: string;
  reviewedAt?: string;
  resolutionNotes?: string;
  createdAt: string;
  updatedAt: string;
}

export interface ApprovalRecord {
  id: string;
  firmId: string;
  clientId: string;
  workflowRunId: string;
  itemType: 'journal_entry' | 'matched_group' | 'adjusted_trial_balance' | 'exception_resolution';
  itemId: string;
  action: 'approved' | 'rejected' | 'modified';
  performedBy: string;
  performedAt: string;
  previousPayload: any;
  approvedPayload: any;
  comments?: string;
}

export interface ExportArtifact {
  id: string;
  firmId: string;
  clientId: string;
  workflowRunId: string;
  fileName: string;
  fileType: 'xlsx' | 'csv' | 'pdf' | 'zip' | 'json';
  format?: string;
  fileSizeBytes: number;
  downloadUrl: string;
  contentBase64?: string;
  generatedAt: string;
  sheetNames?: string[];
  totalRecords: number;
  integrityVerified: boolean;
  fileHash?: string;
  metadata?: Record<string, any>;
}

export interface AuditEvent {
  id: string;
  firmId: string;
  clientId: string;
  userId: string;
  userRole: string;
  action: string;
  entityType: 'job' | 'workflow' | 'file' | 'exception' | 'approval' | 'export';
  entityId: string;
  details: string;
  timestamp: string;
  ipAddress?: string;
  metadata?: Record<string, any>;
}

export interface WorkflowStepDefinition {
  stepId: string;
  name: string;
  description: string;
  required: boolean;
  dependencies: string[]; // stepIds that must complete first
  action: 'validate_file' | 'extract_data' | 'normalize' | 'validate_rules' | 'execute_match' | 'detect_exceptions' | 'generate_drafts' | 'export_reports';
  config?: Record<string, any>;
}

export interface WorkflowDefinition {
  id: string;
  name: string;
  version: string;
  description: string;
  category: 'bank_statement' | 'invoice_processing' | 'gst_reconciliation' | 'trial_balance_audit' | 'general';
  steps: WorkflowStepDefinition[];
  requiresApprovalBeforeCompletion: boolean;
  retryPolicy: {
    maxAttempts: number;
    backoffMs: number;
  };
}

export interface WorkflowStepRun {
  stepId: string;
  name: string;
  status: StepStatus;
  startedAt?: string;
  completedAt?: string;
  errorMessage?: string;
  processedItems: number;
  producedItems: number;
  issuesCount: number;
}

export interface WorkflowRun {
  id: string;
  firmId: string;
  clientId: string;
  workflowId: string;
  jobId: string;
  status: JobStatus;
  progressPercent: number;
  stepRuns: Record<string, WorkflowStepRun>;
  sourceFiles: SourceFile[];
  statistics: {
    totalRecords: number;
    matchedRecords: number;
    unmatchedRecords: number;
    exceptionsCount: number;
    pendingApprovalsCount: number;
  };
  createdAt: string;
  updatedAt: string;
  completedAt?: string;
}

export interface ProcessingJob {
  id: string;
  firmId: string;
  clientId: string;
  idempotencyKey?: string;
  workflowType: string;
  status: JobStatus;
  progress: number;
  currentStep: string;
  totalSteps: number;
  completedSteps: number;
  error?: string;
  retryCount: number;
  maxRetries: number;
  createdAt: string;
  updatedAt: string;
  completedAt?: string;
  runId?: string;
  artifacts?: ExportArtifact[];
}

export interface MatchingConfig {
  dateToleranceDays: number;
  amountTolerance: number;
  checkDirection: boolean;
  fuzzyNarrationThreshold: number; // 0 to 1
  allowOneToMany: boolean;
  preventReusingMatches: boolean;
}

export interface AccountingRuleConfig {
  firmId: string;
  maxCashVoucherAmount: number; // Section 40A(3) - e.g. 10000
  gstTolerance: number; // e.g. 2.0
  allowNegativeCash: boolean; // default false
  requireBalancedJournals: boolean; // default true
  enforceBalanceContinuity: boolean; // default true
}

// ==========================================
// MODULE 1: MONTH-END CLOSING AUTOMATION
// ==========================================
export type ClosingPeriodStatus = 'draft' | 'in_progress' | 'review_required' | 'closed' | 'reopened';

export interface ClosingPeriod {
  id: string;
  firmId: string;
  clientId: string;
  periodName: string; // e.g. "September 2026"
  financialYear: string; // e.g. "FY 2026-27"
  targetClosingDate: string;
  status: ClosingPeriodStatus;
  progressPercent: number;
  unresolvedExceptionsCount: number;
  approvedBy?: string;
  approvedAt?: string;
  approvalNotes?: string;
  createdAt: string;
  updatedAt: string;
}

export interface ClosingTaskItem {
  id: string;
  periodId: string;
  firmId: string;
  clientId: string;
  category: 'Cash & Bank' | 'Receivables & Payables' | 'General Ledger & Journals' | 'Fixed Assets' | 'Accruals & Cut-off' | 'Taxation & Statutory' | 'Financial Reporting';
  title: string;
  description: string;
  assignee: string;
  dueDate: string;
  completedDate?: string;
  completedAt?: string;
  completedBy?: string;
  status: 'not_started' | 'in_progress' | 'blocked' | 'completed';
  priority: 'low' | 'medium' | 'high' | 'critical';
  prerequisiteTaskIds?: string[];
  associatedWorkflowId?: string; // e.g. 'wf_bank_rec', 'wf_gst_rec', 'wf_tb_audit'
  notes: string;
  completionNotes?: string;
  supportingDocRef?: string;
  isOverdue?: boolean;
}

// ==========================================
// MODULE 2: AP & AR AUTOMATION
// ==========================================
export interface AgingRecord {
  id: string;
  firmId: string;
  clientId: string;
  type: 'payable' | 'receivable';
  partyName: string;
  gstin: string;
  invoiceNo: string;
  invoiceDate: string;
  dueDate: string;
  ageDays: number;
  grossAmount: number;
  paidAmount: number;
  outstandingBalance: number;
  bucketLabel: string; // '0-30 Days', '31-60 Days', etc.
  isCreditBalance: boolean;
  isMsmeOverdue: boolean; // Over 45 days under MSMED Act Section 15
  duplicateRisk: boolean;
  notes?: string;
}

export interface PaymentRiskAlert {
  id: string;
  firmId: string;
  clientId: string;
  vendorName: string;
  invoiceNo: string;
  amount: number;
  riskType: 'duplicate_invoice_number' | 'identical_amount_window' | 'overpayment_risk' | 'msme_45_day_violation';
  description: string;
  severity: 'warning' | 'critical';
  suggestedAction: string;
}

// ==========================================
// MODULE 3: AUDIT EVIDENCE AUTOMATION
// ==========================================
export type EvidenceStatus = 'Missing' | 'Requested' | 'Received' | 'Linked' | 'Reviewed' | 'Rejected';

export interface AuditEvidence {
  id: string;
  firmId: string;
  clientId: string;
  period: string;
  auditArea: 'Revenue' | 'Purchases' | 'Cash & Bank' | 'Fixed Assets' | 'Taxes' | 'Statutory' | 'General';
  wpIndex: string; // e.g. "WP-BNK-001", "WP-PUR-001"
  documentTitle: string;
  fileName?: string;
  fileHash?: string;
  status: EvidenceStatus;
  reviewerNotes?: string;
  reviewedBy?: string;
  reviewedAt?: string;
  uploadedAt?: string;
  sourceReference?: string;
}

export interface EvidenceLink {
  id: string;
  evidenceId: string;
  transactionRef: string;
  voucherNo: string;
  amount: number;
  date: string;
  accountHead: string;
  verificationStatus: 'Satisfactory' | 'Exceptions Found' | 'Pending Inquiry';
}

// ==========================================
// MODULE 4: ANOMALY DETECTION ENGINE
// ==========================================
export type AnomalySeverity = 'info' | 'warning' | 'critical';

export interface AnomalyRule {
  id: string;
  firmId: string;
  ruleCode: string;
  name: string;
  description: string;
  category: 'journal_duplicate' | 'round_sum' | 'cash_limit' | 'weekend_posting' | 'abnormal_polarity' | 'velocity_spike' | 'missing_tax_info';
  severity: AnomalySeverity;
  thresholdValue?: number;
  enabled: boolean;
}

export interface AnomalyAlert {
  id: string;
  firmId: string;
  clientId: string;
  ruleCode: string;
  ruleName: string;
  severity: AnomalySeverity;
  voucherNo?: string;
  accountName?: string;
  amount?: number;
  date?: string;
  explanation: string;
  supportingEvidence?: Record<string, any>;
  reviewStatus: 'open' | 'investigating' | 'dismissed_false_positive' | 'marked_for_adjustment';
  resolutionNotes?: string;
  reviewedBy?: string;
  reviewedAt?: string;
  createdAt: string;
}

// ==========================================
// MODULE 5: CLIENT DOCUMENT COLLECTION
// ==========================================
export type DocRequestStatus = 'requested' | 'partially_received' | 'received' | 'accepted' | 'rejected' | 'overdue';

export interface ClientDocumentRequest {
  id: string;
  firmId: string;
  clientId: string;
  financialPeriod: string;
  engagementType: 'Statutory Audit' | 'Tax Audit' | 'GST Annual Return' | 'Due Diligence' | 'Monthly Accounting';
  title: string;
  description: string;
  category: string;
  requiredDueDate: string;
  status: DocRequestStatus;
  reminderCount: number;
  lastReminderSentAt?: string;
  receivedFiles: {
    fileName: string;
    fileSizeBytes: number;
    uploadedAt: string;
    fileHash: string;
    accepted: boolean;
  }[];
  createdAt: string;
  updatedAt: string;
}

export interface DocumentReminderDraft {
  clientId: string;
  clientName: string;
  contactEmail: string;
  subject: string;
  body: string;
  pendingItemsCount: number;
  earliestDueDate: string;
}

// ==========================================
// MODULE 6: RECURRING COMPLIANCE & TASKS
// ==========================================
export type ComplianceFrequency = 'monthly' | 'quarterly' | 'half_yearly' | 'annual';

export interface RecurringTaskDefinition {
  id: string;
  firmId: string;
  title: string;
  statutoryAct: 'GST' | 'Income Tax' | 'Companies Act' | 'ESI/PF' | 'Internal Audit';
  frequency: ComplianceFrequency;
  dueDayOfMonth: number; // e.g. 11 for GSTR-1, 20 for GSTR-3B, 7 for TDS
  description: string;
  associatedWorkflowId?: string;
  isStatutory: boolean;
}

export interface TaskInstance {
  id: string;
  firmId: string;
  clientId: string;
  definitionId: string;
  title: string;
  period: string; // e.g. "September 2026"
  dueDate: string;
  status: 'upcoming' | 'due_today' | 'overdue' | 'completed';
  assignedTo: string;
  completedAt?: string;
  completionProofRef?: string;
  verifiedBy?: string;
}

// ==========================================
// MODULE 7: MULTI-CLIENT BATCH AUTOMATION
// ==========================================
export interface BatchItem {
  clientId: string;
  clientName: string;
  status: 'queued' | 'running' | 'review_required' | 'completed' | 'failed';
  jobId?: string;
  runId?: string;
  fileCount: number;
  recordsProcessed: number;
  exceptionsCount: number;
  errorMessage?: string;
}

export interface BatchRun {
  id: string;
  firmId: string;
  workflowId: string;
  workflowName: string;
  startedAt: string;
  completedAt?: string;
  status: 'queued' | 'running' | 'completed' | 'partial_failure' | 'failed';
  totalClients: number;
  completedClients: number;
  failedClients: number;
  items: BatchItem[];
}

