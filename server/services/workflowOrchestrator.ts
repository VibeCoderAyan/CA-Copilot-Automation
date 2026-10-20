/**
 * Intelligent Workflow Orchestration Layer
 * Manages states: queued -> running -> review_required -> completed / failed
 */

import { EventEmitter } from 'events';
import type {
  ProcessingJob,
  WorkflowRun,
  WorkflowDefinition,
  FirmContext,
  ExtractedRecord,
  RecordMatch,
  ExportArtifact,
} from '../types';
import { BackendStore } from '../db/store';
import { FileProcessorService } from './fileProcessor';
import { UniversalMatchingEngine } from './matchingEngine';
import { AccountingRulesEngine } from './accountingRulesEngine';
import { ExceptionService } from './exceptionService';
import { ExportService } from './exportService';
import { suggestAccountHead } from '../../src/utils/bankStatementConverterEngine';
import { mapToScheduleIII, inferAccountHead } from '../../src/utils/financialReportsEngine';

export class WorkflowOrchestrator extends EventEmitter {
  private store: BackendStore;
  private exceptionService: ExceptionService;
  private exportService: ExportService;

  constructor(store?: BackendStore) {
    super();
    this.store = store || BackendStore.getInstance();
    this.exceptionService = new ExceptionService(this.store);
    this.exportService = new ExportService(this.store);
  }

  /**
   * Starts a new workflow execution with idempotency protection
   */
  public async startWorkflow(
    context: FirmContext,
    workflowId: string,
    files: { fileName: string; buffer: Buffer }[],
    options: { idempotencyKey?: string; configOverrides?: Record<string, any> } = {}
  ): Promise<{ job: ProcessingJob; run: WorkflowRun }> {
    // 1. Check Idempotency Key
    if (options.idempotencyKey) {
      const existingJob = this.store.findJobByIdempotencyKey(context.firmId, options.idempotencyKey);
      if (existingJob) {
        const existingRun = this.store.getWorkflowRun(context.firmId, existingJob.runId || '');
        if (existingRun) {
          return { job: existingJob, run: existingRun };
        }
      }
    }

    const definition = this.store.getWorkflowDefinition(workflowId);
    if (!definition) {
      throw new Error(`Workflow definition "${workflowId}" not found`);
    }

    const jobId = `job-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`;
    const runId = `run-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`;

    // Process files immediately to register them
    const sourceFiles = [];
    const extractedFileRecords: Record<string, ExtractedRecord[]> = {};

    for (const f of files) {
      const result = FileProcessorService.processFileBuffer(f.buffer, f.fileName, context.firmId, context.clientId);
      this.store.saveSourceFile(result.sourceFile);
      this.store.saveExtractedRecords(result.sourceFile.id, result.records);
      sourceFiles.push(result.sourceFile);
      extractedFileRecords[result.sourceFile.id] = result.records;
    }

    const initialStepRuns: Record<string, any> = {};
    for (const step of definition.steps) {
      initialStepRuns[step.stepId] = {
        stepId: step.stepId,
        name: step.name,
        status: 'pending',
        processedItems: 0,
        producedItems: 0,
        issuesCount: 0,
      };
    }

    const run: WorkflowRun = {
      id: runId,
      firmId: context.firmId,
      clientId: context.clientId,
      workflowId,
      jobId,
      status: 'queued',
      progressPercent: 0,
      stepRuns: initialStepRuns,
      sourceFiles,
      statistics: {
        totalRecords: Object.values(extractedFileRecords).reduce((acc, r) => acc + r.length, 0),
        matchedRecords: 0,
        unmatchedRecords: 0,
        exceptionsCount: 0,
        pendingApprovalsCount: 0,
      },
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };

    const job: ProcessingJob = {
      id: jobId,
      firmId: context.firmId,
      clientId: context.clientId,
      idempotencyKey: options.idempotencyKey,
      workflowType: definition.name,
      status: 'queued',
      progress: 0,
      currentStep: 'Queued',
      totalSteps: definition.steps.length,
      completedSteps: 0,
      retryCount: 0,
      maxRetries: definition.retryPolicy.maxAttempts,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      runId,
      artifacts: [],
    };

    this.store.createWorkflowRun(run);
    this.store.createJob(job);

    // Audit Logging
    this.store.logAuditEvent({
      id: `aud-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`,
      firmId: context.firmId,
      clientId: context.clientId,
      userId: context.userId,
      userRole: context.userRole,
      action: 'WORKFLOW_STARTED',
      entityType: 'job',
      entityId: jobId,
      details: `Started workflow "${definition.name}" with ${files.length} source files`,
      timestamp: new Date().toISOString(),
    });

    // Execute asynchronously (or directly for local server response)
    this.executeWorkflowSteps(context, definition, job, run, extractedFileRecords, options.configOverrides).catch((err) => {
      console.error(`[WorkflowOrchestrator] Execution failed for Job ${jobId}:`, err);
      this.store.updateJob(context.firmId, jobId, {
        status: 'failed',
        error: err.message || 'Workflow execution error',
      });
      this.store.updateWorkflowRun(context.firmId, runId, {
        status: 'failed',
      });
    });

    return { job, run };
  }

  private notifyJobUpdate(firmId: string, jobId: string, patch: Partial<ProcessingJob>): ProcessingJob | undefined {
    const updated = this.store.updateJob(firmId, jobId, patch);
    if (updated) {
      this.emit('progress', updated);
      this.emit(`job:${jobId}`, updated);
    }
    return updated;
  }

  /**
   * Internal Step Executor
   */
  private async executeWorkflowSteps(
    context: FirmContext,
    definition: WorkflowDefinition,
    job: ProcessingJob,
    run: WorkflowRun,
    fileRecords: Record<string, ExtractedRecord[]>,
    configOverrides: Record<string, any> = {}
  ): Promise<void> {
    this.notifyJobUpdate(context.firmId, job.id, { status: 'running', currentStep: 'Initializing Pipeline' });
    this.store.updateWorkflowRun(context.firmId, run.id, { status: 'running' });

    let completedSteps = 0;
    const allRecords: ExtractedRecord[] = Object.values(fileRecords).flat();
    let generatedMatches: RecordMatch[] = [];
    const generatedArtifacts: ExportArtifact[] = [];
    let exceptionsCount = 0;

    for (const step of definition.steps) {
      this.notifyJobUpdate(context.firmId, job.id, {
        currentStep: step.name,
        progress: Math.round((completedSteps / definition.steps.length) * 100),
      });

      const stepStart = new Date().toISOString();

      try {
        if (step.action === 'validate_file') {
          // File validation step
          for (const file of run.sourceFiles) {
            if (file.fileSizeBytes === 0) {
              throw new Error(`File ${file.fileName} is empty`);
            }
          }
        } else if (step.action === 'extract_data') {
          // Already extracted into fileRecords, verify counts
          if (allRecords.length === 0) {
            this.exceptionService.logException(
              context,
              run.id,
              job.id,
              'format_error',
              'error',
              'No Extracted Records',
              'The uploaded files did not yield any structured transaction rows.',
              { fileCount: run.sourceFiles.length }
            );
            exceptionsCount++;
          }
        } else if (step.action === 'validate_rules') {
          // Accounting Rules Engine Verification
          const firmConfig = this.store.getFirmConfig(context.firmId);
          const firstDocType = run.sourceFiles[0]?.documentType || 'unknown';
          const issues = AccountingRulesEngine.validateDataset(allRecords, firmConfig, firstDocType);

          for (const issue of issues) {
            this.exceptionService.logException(
              context,
              run.id,
              job.id,
              issue.severity === 'critical' ? 'discontinuity' : 'statutory_violation',
              issue.severity,
              issue.ruleName,
              issue.message,
              { actual: issue.actualValue, expected: issue.expectedValue },
              issue.suggestedCorrection ? JSON.stringify(issue.suggestedCorrection) : undefined,
              issue.sourceFileId,
              issue.recordId
            );
            exceptionsCount++;
          }
        } else if (step.action === 'execute_match') {
          // Universal Matching Engine Execution - Smart primary/secondary detection
          const fileIds = Object.keys(fileRecords);
          let primaryFileId = fileIds[0];
          let secondaryFileId = fileIds[1];

          const bankFile = run.sourceFiles.find((f) => f.documentType === 'bank_statement');
          const ledgerFile = run.sourceFiles.find((f) => f.documentType === 'ledger');
          const prFile = run.sourceFiles.find((f) => f.documentType === 'purchase_register');
          const gstr2bFile = run.sourceFiles.find((f) => f.documentType === 'gstr2b');

          if (bankFile && ledgerFile) {
            primaryFileId = bankFile.id;
            secondaryFileId = ledgerFile.id;
          } else if (prFile && gstr2bFile) {
            primaryFileId = prFile.id;
            secondaryFileId = gstr2bFile.id;
          }

          const primaryRecords = fileRecords[primaryFileId] || [];
          const secondaryRecords = fileRecords[secondaryFileId] || [];

          if (secondaryRecords.length > 0) {
            const matchingResult = UniversalMatchingEngine.executeMatching(
              primaryRecords,
              secondaryRecords,
              run.id,
              configOverrides.matching
            );
            generatedMatches = matchingResult.matches;
            this.store.saveRecordMatches(run.id, generatedMatches);

            // Log exceptions for unmatched items
            for (const unmatched of matchingResult.unmatchedPrimary) {
              this.exceptionService.logException(
                context,
                run.id,
                job.id,
                'unmatched_transaction',
                'warning',
                'Unmatched Transaction',
                `No matching counterpart found for ${unmatched.normalizedFields.date || 'Undated'} ₹${(unmatched.normalizedFields.amount || 0).toFixed(2)} (${unmatched.normalizedFields.narration || unmatched.normalizedFields.invoiceNo || 'No narration'})`,
                unmatched.normalizedFields,
                'Review vendor ledger or draft adjusting journal entry',
                unmatched.sourceFileId,
                unmatched.id
              );
              exceptionsCount++;
            }
          }
        } else if (step.action === 'generate_drafts') {
          // Generate draft journal entries for unmatched bank transactions
          const draftJournals: any[] = [];
          for (const rec of allRecords) {
            const wdr = rec.normalizedFields.debit || 0;
            const dep = rec.normalizedFields.credit || 0;
            if (wdr > 0 || dep > 0) {
              const suggestion = suggestAccountHead(rec.normalizedFields.narration || '', wdr > 0);
              draftJournals.push({
                date: rec.normalizedFields.date || '2026-09-30',
                debitAccount: wdr > 0 ? suggestion.accountHead : 'HDFC Bank Account',
                creditAccount: wdr > 0 ? 'HDFC Bank Account' : suggestion.accountHead,
                amount: Math.max(wdr, dep),
                narration: rec.normalizedFields.narration || 'Bank Transaction',
                confidence: suggestion.confidence,
                category: suggestion.rule,
              });
            }
          }
          run.statistics.pendingApprovalsCount = draftJournals.length;
        } else if (step.action === 'detect_exceptions') {
          // Batch duplicate invoice detection
          const seen = new Map<string, ExtractedRecord>();
          for (const rec of allRecords) {
            const key = `${rec.normalizedFields.gstin}|${rec.normalizedFields.invoiceNo}|${rec.normalizedFields.totalAmount}`;
            if (rec.normalizedFields.invoiceNo && seen.has(key)) {
              this.exceptionService.logException(
                context,
                run.id,
                job.id,
                'duplicate_invoice',
                'critical',
                'Duplicate Invoice Detected',
                `Invoice ${rec.normalizedFields.invoiceNo} from GSTIN ${rec.normalizedFields.gstin} for ₹${(rec.normalizedFields.totalAmount || 0).toFixed(2)} appears multiple times in batch.`,
                rec.normalizedFields,
                'Discard duplicate voucher to avoid double taxation',
                rec.sourceFileId,
                rec.id
              );
              exceptionsCount++;
            } else if (rec.normalizedFields.invoiceNo) {
              seen.set(key, rec);
            }
          }
        } else if (step.action === 'export_reports') {
          // Multi-sheet Excel export based on workflow type
          if (definition.category === 'bank_statement') {
            const bankFile = run.sourceFiles.find((f) => f.documentType === 'bank_statement');
            const ledgerFile = run.sourceFiles.find((f) => f.documentType === 'ledger');
            const bankRecords = bankFile ? (fileRecords[bankFile.id] || []) : allRecords;
            const ledgerRecords = ledgerFile ? (fileRecords[ledgerFile.id] || []) : [];

            // 1. Clean Bank Statement Records
            const statementRows = bankRecords.map((r, idx) => ({
              'S.No': idx + 1,
              'Txn Date': r.normalizedFields.date || '',
              'Particulars / Narration': r.normalizedFields.narration || '',
              'Chq / Ref No': r.normalizedFields.refNo || '',
              'Withdrawal (Debit)': r.normalizedFields.debit || 0,
              'Deposit (Credit)': r.normalizedFields.credit || 0,
              'Running Balance': r.normalizedFields.balance || 0,
            }));

            // 2. Clean Books Ledger Records (if ingested)
            const ledgerRows = ledgerRecords.map((r, idx) => ({
              'S.No': idx + 1,
              'Voucher Date': r.normalizedFields.date || '',
              'Voucher No': r.normalizedFields.invoiceNo || r.normalizedFields.refNo || `VCH-${idx + 1}`,
              'Particulars': r.normalizedFields.narration || '',
              'Debit (Receipt in Books)': r.normalizedFields.debit || 0,
              'Credit (Payment in Books)': r.normalizedFields.credit || 0,
            }));

            // 3. Matched Transactions
            const matchedRows = generatedMatches.map((m, idx) => ({
              'Match #': idx + 1,
              'Match Type': m.matchType,
              'Confidence Score': `${m.matchScore}%`,
              'Bank Date': m.primaryRecord.normalizedFields.date || '',
              'Bank Narration': m.primaryRecord.normalizedFields.narration || '',
              'Bank Amount': m.primaryRecord.normalizedFields.amount || 0,
              'Ledger Voucher': m.matchedRecords[0]?.normalizedFields.invoiceNo || m.matchedRecords[0]?.normalizedFields.refNo || 'N/A',
              'Ledger Particulars': m.matchedRecords[0]?.normalizedFields.narration || '',
              'Ledger Amount': m.matchedRecords[0]?.normalizedFields.amount || 0,
              'Match Status': m.matchStatus,
            }));

            // 4. Clean Exceptions
            const excRecords = this.store.listExceptions(context.firmId, { workflowRunId: run.id });
            const exceptionRows = excRecords.map((e: any, idx: number) => ({
              'Issue #': idx + 1,
              'Severity': e.severity ? e.severity.toUpperCase() : 'INFO',
              'Category': e.category || 'General',
              'Title': e.title || '',
              'Description': e.description || '',
              'Status': e.resolutionStatus || 'pending',
              'Action Required': e.suggestedAction || 'Review entry against voucher',
            }));

            // 5. BRS Summary
            const closingBankBal = bankRecords.length > 0 ? (bankRecords[bankRecords.length - 1].normalizedFields.balance || 0) : 0;
            const unpresented = excRecords.filter((e) => e.category === 'unmatched_transaction').length;
            const brsSummary = [
              { Parameter: 'Bank Account Statement Closing Balance', 'Amount (₹)': closingBankBal },
              { Parameter: 'Total Bank Transactions Processed', 'Amount (₹)': bankRecords.length },
              { Parameter: 'Total Books Ledger Vouchers Processed', 'Amount (₹)': ledgerRecords.length },
              { Parameter: 'Total Successfully Matched Transactions', 'Amount (₹)': generatedMatches.length },
              { Parameter: 'Audit Exceptions / Unmatched Items Count', 'Amount (₹)': excRecords.length },
            ];

            const workbookSheets: any[] = [
              { sheetName: 'BRS_Executive_Summary', rows: brsSummary },
              { sheetName: 'Clean_Bank_Statement', rows: statementRows },
            ];

            if (ledgerRows.length > 0) {
              workbookSheets.push({ sheetName: 'Clean_Books_Ledger', rows: ledgerRows });
            }
            if (matchedRows.length > 0) {
              workbookSheets.push({ sheetName: 'Matched_Transactions', rows: matchedRows });
            }
            workbookSheets.push({
              sheetName: 'Exceptions_Audit',
              rows: exceptionRows.length > 0 ? exceptionRows : [{ Notice: 'Zero audit exceptions detected' }],
            });

            const artifact = this.exportService.generateExcelWorkbook(
              context,
              run.id,
              `Bank_Reconciliation_${context.clientId}_${Date.now()}.xlsx`,
              workbookSheets
            );
            generatedArtifacts.push(artifact);
          } else if (definition.category === 'invoice_processing') {
            const invoiceRows = allRecords.map((r, idx) => ({
              'S.No': idx + 1,
              'Invoice No': r.normalizedFields.invoiceNo || '',
              'Invoice Date': r.normalizedFields.date || '',
              'Supplier GSTIN': r.normalizedFields.gstin || '',
              'Supplier Name': r.normalizedFields.partyName || '',
              'Taxable Value': r.normalizedFields.taxableAmount || 0,
              'CGST': r.normalizedFields.cgst || 0,
              'SGST': r.normalizedFields.sgst || 0,
              'IGST': r.normalizedFields.igst || 0,
              'Cess': r.normalizedFields.cess || 0,
              'Total Amount': r.normalizedFields.totalAmount || 0,
            }));

            const excRecords = this.store.listExceptions(context.firmId, { workflowRunId: run.id });
            const exceptionRows = excRecords.map((e: any, idx: number) => ({
              'Issue #': idx + 1,
              'Severity': e.severity ? e.severity.toUpperCase() : 'INFO',
              'Category': e.category || 'General',
              'Title': e.title || '',
              'Description': e.description || '',
              'Status': e.resolutionStatus || 'pending',
              'Suggested Action': e.suggestedAction || 'Review invoice duplicate or GSTIN',
            }));

            const artifact = this.exportService.generateExcelWorkbook(
              context,
              run.id,
              `Invoice_Register_${context.clientId}_${Date.now()}.xlsx`,
              [
                { sheetName: 'Invoice_Register', rows: invoiceRows },
                {
                  sheetName: 'Audit_Exceptions',
                  rows: exceptionRows.length > 0 ? exceptionRows : [{ Notice: 'No duplicate or validation issues detected' }],
                },
              ]
            );
            generatedArtifacts.push(artifact);
          } else if (definition.category === 'gst_reconciliation') {
            const matchedRows = generatedMatches.map((m, idx) => ({
              'S.No': idx + 1,
              'Match Type': m.matchType,
              'Confidence Score': `${m.matchScore}%`,
              'Books Invoice No': m.primaryRecord?.normalizedFields?.invoiceNo || '',
              'Books Taxable': m.primaryRecord?.normalizedFields?.taxableAmount || 0,
              '2B Invoice No': m.matchedRecords[0]?.normalizedFields?.invoiceNo || '',
              '2B Taxable': m.matchedRecords[0]?.normalizedFields?.taxableAmount || 0,
              'Variance Amount': m.varianceAmount || 0,
              'Explanation': m.explanation || '',
            }));

            const excRecords = this.store.listExceptions(context.firmId, { workflowRunId: run.id });
            const exceptionRows = excRecords.map((e: any, idx: number) => ({
              'Issue #': idx + 1,
              'Severity': e.severity ? e.severity.toUpperCase() : 'INFO',
              'Category': e.category || 'GST Discrepancy',
              'Title': e.title || '',
              'Description': e.description || '',
              'Status': e.resolutionStatus || 'pending',
            }));

            const artifact = this.exportService.generateExcelWorkbook(
              context,
              run.id,
              `GST_Reconciliation_${context.clientId}_${Date.now()}.xlsx`,
              [
                { sheetName: 'Matched_Summary', rows: matchedRows.length > 0 ? matchedRows : [{ Notice: 'No matched records' }] },
                {
                  sheetName: 'Audit_Exceptions',
                  rows: exceptionRows.length > 0 ? exceptionRows : [{ Notice: 'Zero GST discrepancies' }],
                },
                {
                  sheetName: 'Statutory_Notice',
                  rows: [
                    { Statutory_Provision: 'Section 16(2)(aa) CGST Act', Notice: 'ITC can only be availed if invoice details are furnished by supplier in GSTR-1 and communicated to recipient in GSTR-2B.' },
                  ],
                },
              ]
            );
            generatedArtifacts.push(artifact);
          } else if (definition.category === 'trial_balance_audit') {
            const tbRows = allRecords.map((r, idx) => ({
              'S.No': idx + 1,
              'Account Code': r.normalizedFields.accountCode || '',
              'Account Name': r.normalizedFields.accountName || '',
              'Schedule III Group': mapToScheduleIII(r.normalizedFields.accountName || '', inferAccountHead(r.normalizedFields.accountName || '', r.normalizedFields.debit || 0, r.normalizedFields.credit || 0).head, (r.normalizedFields.debit || 0) - (r.normalizedFields.credit || 0)),
              'Debit (₹)': r.normalizedFields.debit || 0,
              'Credit (₹)': r.normalizedFields.credit || 0,
            }));

            const excRecords = this.store.listExceptions(context.firmId, { workflowRunId: run.id });
            const exceptionRows = excRecords.map((e: any, idx: number) => ({
              'Issue #': idx + 1,
              'Severity': e.severity ? e.severity.toUpperCase() : 'INFO',
              'Category': e.category || 'TB Imbalance',
              'Title': e.title || '',
              'Description': e.description || '',
              'Status': e.resolutionStatus || 'pending',
            }));

            const artifact = this.exportService.generateExcelWorkbook(
              context,
              run.id,
              `Schedule_III_Statements_${context.clientId}_${Date.now()}.xlsx`,
              [
                { sheetName: 'Schedule_III_TB', rows: tbRows },
                {
                  sheetName: 'Audit_Exceptions',
                  rows: exceptionRows.length > 0 ? exceptionRows : [{ Notice: 'Trial balance is balanced with zero audit exceptions' }],
                },
              ]
            );
            generatedArtifacts.push(artifact);
          }
        }

        completedSteps++;
        run.stepRuns[step.stepId] = {
          stepId: step.stepId,
          name: step.name,
          status: 'completed',
          startedAt: stepStart,
          completedAt: new Date().toISOString(),
          processedItems: allRecords.length,
          producedItems: allRecords.length,
          issuesCount: exceptionsCount,
        };
      } catch (err: any) {
        run.stepRuns[step.stepId] = {
          stepId: step.stepId,
          name: step.name,
          status: 'failed',
          startedAt: stepStart,
          completedAt: new Date().toISOString(),
          errorMessage: err.message,
          processedItems: 0,
          producedItems: 0,
          issuesCount: exceptionsCount + 1,
        };
        throw err;
      }
    }

    // Determine final status
    const pendingExceptions = this.store.listExceptions(context.firmId, { workflowRunId: run.id, status: 'pending' });
    const finalStatus = definition.requiresApprovalBeforeCompletion && pendingExceptions.length > 0 ? 'review_required' : 'completed';

    this.notifyJobUpdate(context.firmId, job.id, {
      status: finalStatus,
      progress: 100,
      currentStep: finalStatus === 'review_required' ? 'Awaiting Human-in-the-Loop Review' : 'Completed',
      completedSteps,
      artifacts: generatedArtifacts,
      completedAt: finalStatus === 'completed' ? new Date().toISOString() : undefined,
    });

    this.store.updateWorkflowRun(context.firmId, run.id, {
      status: finalStatus,
      progressPercent: 100,
      statistics: {
        totalRecords: allRecords.length,
        matchedRecords: generatedMatches.filter((m) => m.matchType !== 'unmatched').length,
        unmatchedRecords: generatedMatches.filter((m) => m.matchType === 'unmatched').length,
        exceptionsCount: pendingExceptions.length,
        pendingApprovalsCount: run.statistics.pendingApprovalsCount,
      },
      completedAt: finalStatus === 'completed' ? new Date().toISOString() : undefined,
    });
  }

  /**
   * Retries a failed or stalled job idempotently without duplicating artifacts
   */
  public async retryJob(context: FirmContext, jobId: string): Promise<ProcessingJob> {
    const job = this.store.getJob(context.firmId, jobId);
    if (!job) throw new Error('Job not found');
    if (job.status === 'completed') return job;

    if (job.retryCount >= job.maxRetries) {
      throw new Error(`Job exceeded maximum retry count (${job.maxRetries})`);
    }

    const run = this.store.getWorkflowRun(context.firmId, job.runId || '');
    if (!run) throw new Error('Workflow run not found');

    const definition = this.store.getWorkflowDefinition(run.workflowId);
    if (!definition) throw new Error('Workflow definition not found');

    const fileRecords: Record<string, ExtractedRecord[]> = {};
    for (const f of run.sourceFiles) {
      fileRecords[f.id] = this.store.getExtractedRecords(f.id);
    }

    this.store.updateJob(context.firmId, jobId, {
      retryCount: job.retryCount + 1,
      status: 'queued',
      error: undefined,
    });

    // Re-execute pipeline
    this.executeWorkflowSteps(context, definition, job, run, fileRecords).catch((err) => {
      this.store.updateJob(context.firmId, jobId, {
        status: 'failed',
        error: err.message,
      });
    });

    return job;
  }
}
