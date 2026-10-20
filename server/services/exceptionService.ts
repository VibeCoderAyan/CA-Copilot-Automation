/**
 * Exception Review, Approval & Audit Trail Service
 * Manages human-in-the-loop exception resolution
 */

import type { ExceptionRecord, ApprovalRecord, FirmContext, ExceptionResolution } from '../types';
import { BackendStore } from '../db/store';

export class ExceptionService {
  private store: BackendStore;

  constructor(store?: BackendStore) {
    this.store = store || BackendStore.getInstance();
  }

  /**
   * Creates an exception from a validation issue or match mismatch
   */
  public logException(
    context: FirmContext,
    workflowRunId: string,
    jobId: string,
    category: ExceptionRecord['category'],
    severity: ExceptionRecord['severity'],
    title: string,
    description: string,
    originalValues: Record<string, any>,
    suggestedResolution?: string,
    sourceFileId?: string,
    recordId?: string
  ): ExceptionRecord {
    const exception: ExceptionRecord = {
      id: `exc-${Date.now()}-${Math.random().toString(36).substring(2, 7)}`,
      firmId: context.firmId,
      clientId: context.clientId,
      workflowRunId,
      jobId,
      sourceFileId,
      recordId,
      category,
      severity,
      title,
      description,
      originalValues,
      suggestedResolution,
      resolutionStatus: 'pending',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };

    this.store.createException(exception);

    // Audit Event
    this.store.logAuditEvent({
      id: `aud-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`,
      firmId: context.firmId,
      clientId: context.clientId,
      userId: context.userId,
      userRole: context.userRole,
      action: 'EXCEPTION_DETECTED',
      entityType: 'exception',
      entityId: exception.id,
      details: `Exception flagged: [${category}] ${title}`,
      timestamp: new Date().toISOString(),
    });

    return exception;
  }

  /**
   * Resolves an exception with explicit approval, correction, or rejection
   */
  public resolveException(
    context: FirmContext,
    exceptionId: string,
    resolution: ExceptionResolution,
    resolutionNotes?: string,
    correctedValues?: Record<string, any>
  ): { success: boolean; exception?: ExceptionRecord; approval?: ApprovalRecord; error?: string } {
    const exc = this.store.getException(context.firmId, exceptionId);
    if (!exc) {
      return { success: false, error: 'Exception record not found' };
    }

    const previousStatus = exc.resolutionStatus;
    const updatedExc = this.store.updateException(context.firmId, exceptionId, {
      resolutionStatus: resolution,
      resolutionNotes,
      correctedValues: correctedValues || exc.correctedValues,
      reviewedBy: context.userId,
      reviewedAt: new Date().toISOString(),
    });

    // Record formal approval entry
    const approval: ApprovalRecord = {
      id: `appr-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`,
      firmId: context.firmId,
      clientId: context.clientId,
      workflowRunId: exc.workflowRunId,
      itemType: 'exception_resolution',
      itemId: exceptionId,
      action: resolution === 'approved' ? 'approved' : resolution === 'rejected' ? 'rejected' : 'modified',
      performedBy: `${context.userId} (${context.userRole})`,
      performedAt: new Date().toISOString(),
      previousPayload: { status: previousStatus, originalValues: exc.originalValues },
      approvedPayload: { status: resolution, correctedValues: correctedValues || exc.originalValues, notes: resolutionNotes },
      comments: resolutionNotes,
    };
    this.store.recordApproval(approval);

    // Immutable Audit Trail
    this.store.logAuditEvent({
      id: `aud-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`,
      firmId: context.firmId,
      clientId: context.clientId,
      userId: context.userId,
      userRole: context.userRole,
      action: `EXCEPTION_${resolution.toUpperCase()}`,
      entityType: 'approval',
      entityId: approval.id,
      details: `Exception "${exc.title}" marked as ${resolution} by ${context.userId}. Notes: ${resolutionNotes || 'None'}`,
      timestamp: new Date().toISOString(),
      metadata: { exceptionId, previousStatus, resolution, correctedValues },
    });

    return { success: true, exception: updatedExc, approval };
  }

  /**
   * Lists exceptions with filter
   */
  public listExceptions(context: FirmContext, filters?: { clientId?: string; workflowRunId?: string; status?: string }) {
    return this.store.listExceptions(context.firmId, {
      clientId: filters?.clientId || context.clientId,
      workflowRunId: filters?.workflowRunId,
      status: filters?.status,
    });
  }
}
