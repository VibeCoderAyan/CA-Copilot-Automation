/**
 * Frontend Typed API Client for CA Tools Backend Automation Engine
 */

export interface StartWorkflowPayload {
  files: { fileName: string; content?: string; base64?: string }[];
  idempotencyKey?: string;
  configOverrides?: Record<string, any>;
}

export class CAAutomationApiClient {
  private static baseUrl = '/api';

  private static getHeaders(firmId = 'firm_ca_office', clientId?: string): HeadersInit {
    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
      'x-firm-id': firmId,
      'x-firm-name': 'Chartered Accountants & Co.',
      'x-user-id': 'ca_partner_01',
      'x-user-role': 'Partner',
    };
    if (clientId) {
      headers['x-client-id'] = clientId;
    }
    return headers;
  }

  public static async getHealth() {
    const res = await fetch(`${this.baseUrl}/health`);
    return res.json();
  }

  public static async getWorkflows(firmId?: string) {
    const res = await fetch(`${this.baseUrl}/workflows`, {
      headers: this.getHeaders(firmId),
    });
    return res.json();
  }

  public static async startWorkflow(
    workflowId: string,
    payload: StartWorkflowPayload,
    firmId?: string,
    clientId?: string
  ) {
    const res = await fetch(`${this.baseUrl}/workflows/${workflowId}/run`, {
      method: 'POST',
      headers: this.getHeaders(firmId, clientId),
      body: JSON.stringify(payload),
    });
    return res.json();
  }

  public static async getWorkflowRun(runId: string, firmId?: string) {
    const res = await fetch(`${this.baseUrl}/workflows/runs/${runId}`, {
      headers: this.getHeaders(firmId),
    });
    return res.json();
  }

  public static async getJobs(firmId?: string, clientId?: string) {
    const url = clientId
      ? `${this.baseUrl}/jobs?clientId=${encodeURIComponent(clientId)}`
      : `${this.baseUrl}/jobs`;
    const res = await fetch(url, {
      headers: this.getHeaders(firmId, clientId),
    });
    return res.json();
  }

  public static async getJob(jobId: string, firmId?: string) {
    const res = await fetch(`${this.baseUrl}/jobs/${jobId}`, {
      headers: this.getHeaders(firmId),
    });
    return res.json();
  }

  public static async retryJob(jobId: string, firmId?: string) {
    const res = await fetch(`${this.baseUrl}/jobs/${jobId}/retry`, {
      method: 'POST',
      headers: this.getHeaders(firmId),
    });
    return res.json();
  }

  public static async cancelJob(jobId: string, firmId?: string) {
    const res = await fetch(`${this.baseUrl}/jobs/${jobId}/cancel`, {
      method: 'POST',
      headers: this.getHeaders(firmId),
    });
    return res.json();
  }

  public static async getExceptions(filters?: { clientId?: string; workflowRunId?: string; status?: string }, firmId?: string) {
    const params = new URLSearchParams();
    if (filters?.clientId) params.append('clientId', filters.clientId);
    if (filters?.workflowRunId) params.append('workflowRunId', filters.workflowRunId);
    if (filters?.status) params.append('status', filters.status);

    const res = await fetch(`${this.baseUrl}/exceptions?${params.toString()}`, {
      headers: this.getHeaders(firmId, filters?.clientId),
    });
    return res.json();
  }

  public static async resolveException(
    exceptionId: string,
    resolution: 'approved' | 'rejected' | 'corrected' | 'investigating',
    resolutionNotes?: string,
    correctedValues?: Record<string, any>,
    firmId?: string
  ) {
    const res = await fetch(`${this.baseUrl}/exceptions/${exceptionId}/resolve`, {
      method: 'POST',
      headers: this.getHeaders(firmId),
      body: JSON.stringify({ resolution, resolutionNotes, correctedValues }),
    });
    return res.json();
  }

  public static async getAuditEvents(filters?: { clientId?: string; entityId?: string }, firmId?: string) {
    const params = new URLSearchParams();
    if (filters?.clientId) params.append('clientId', filters.clientId);
    if (filters?.entityId) params.append('entityId', filters.entityId);

    const res = await fetch(`${this.baseUrl}/audit?${params.toString()}`, {
      headers: this.getHeaders(firmId, filters?.clientId),
    });
    return res.json();
  }

  // ========================================================
  // MODULE 1: MONTH-END CLOSING AUTOMATION
  // ========================================================
  public static async getClosingPeriods(firmId?: string, clientId?: string) {
    const url = clientId ? `${this.baseUrl}/closing/periods?clientId=${encodeURIComponent(clientId)}` : `${this.baseUrl}/closing/periods`;
    const res = await fetch(url, { headers: this.getHeaders(firmId, clientId) });
    return res.json();
  }

  public static async createClosingPeriod(
    payload: { periodName: string; financialYear: string; targetClosingDate?: string },
    firmId?: string,
    clientId?: string
  ) {
    const res = await fetch(`${this.baseUrl}/closing/periods`, {
      method: 'POST',
      headers: this.getHeaders(firmId, clientId),
      body: JSON.stringify(payload),
    });
    return res.json();
  }

  public static async evaluateClosingPeriod(periodId: string, firmId?: string, clientId?: string) {
    const res = await fetch(`${this.baseUrl}/closing/periods/${periodId}`, {
      headers: this.getHeaders(firmId, clientId),
    });
    return res.json();
  }

  public static async updateClosingTask(
    periodId: string,
    taskId: string,
    payload: { status?: string; completionNotes?: string; assignedTo?: string },
    firmId?: string,
    clientId?: string
  ) {
    const res = await fetch(`${this.baseUrl}/closing/periods/${periodId}/tasks/${taskId}`, {
      method: 'PUT',
      headers: this.getHeaders(firmId, clientId),
      body: JSON.stringify(payload),
    });
    return res.json();
  }

  public static async approveClosing(periodId: string, notes?: string, firmId?: string, clientId?: string) {
    const res = await fetch(`${this.baseUrl}/closing/periods/${periodId}/approve`, {
      method: 'POST',
      headers: this.getHeaders(firmId, clientId),
      body: JSON.stringify({ notes }),
    });
    return res.json();
  }

  public static async exportClosingSummary(periodId: string, firmId?: string, clientId?: string) {
    const res = await fetch(`${this.baseUrl}/closing/periods/${periodId}/export`, {
      headers: this.getHeaders(firmId, clientId),
    });
    return res.json();
  }

  // ========================================================
  // MODULE 2: AP & AR AUTOMATION
  // ========================================================
  public static async processApar(
    payload: { records: any[]; type: 'payable' | 'receivable'; asOfDate?: string },
    firmId?: string,
    clientId?: string
  ) {
    const res = await fetch(`${this.baseUrl}/apar/process`, {
      method: 'POST',
      headers: this.getHeaders(firmId, clientId),
      body: JSON.stringify(payload),
    });
    return res.json();
  }

  public static async getAparRecords(firmId?: string, clientId?: string) {
    const res = await fetch(`${this.baseUrl}/apar/records`, {
      headers: this.getHeaders(firmId, clientId),
    });
    return res.json();
  }

  public static async getPaymentAlerts(firmId?: string, clientId?: string) {
    const res = await fetch(`${this.baseUrl}/apar/alerts`, {
      headers: this.getHeaders(firmId, clientId),
    });
    return res.json();
  }

  public static async exportPaymentReviewList(firmId?: string, clientId?: string) {
    const res = await fetch(`${this.baseUrl}/apar/payment-review`, {
      headers: this.getHeaders(firmId, clientId),
    });
    return res.json();
  }

  // ========================================================
  // MODULE 3: AUDIT EVIDENCE AUTOMATION
  // ========================================================
  public static async initEvidenceSlots(period?: string, firmId?: string, clientId?: string) {
    const res = await fetch(`${this.baseUrl}/evidence/init`, {
      method: 'POST',
      headers: this.getHeaders(firmId, clientId),
      body: JSON.stringify({ period }),
    });
    return res.json();
  }

  public static async getEvidenceList(firmId?: string, clientId?: string) {
    const res = await fetch(`${this.baseUrl}/evidence`, {
      headers: this.getHeaders(firmId, clientId),
    });
    return res.json();
  }

  public static async linkEvidenceTransaction(
    evidenceId: string,
    transaction: { voucherNo: string; date: string; amount: number; partyName: string; accountHead: string; notes?: string },
    firmId?: string,
    clientId?: string
  ) {
    const res = await fetch(`${this.baseUrl}/evidence/link`, {
      method: 'POST',
      headers: this.getHeaders(firmId, clientId),
      body: JSON.stringify({ evidenceId, transaction }),
    });
    return res.json();
  }

  public static async detectMissingEvidence(
    payload: { transactions: any[]; materialityThreshold?: number },
    firmId?: string,
    clientId?: string
  ) {
    const res = await fetch(`${this.baseUrl}/evidence/detect-missing`, {
      method: 'POST',
      headers: this.getHeaders(firmId, clientId),
      body: JSON.stringify(payload),
    });
    return res.json();
  }

  public static async exportEvidenceRegister(period?: string, firmId?: string, clientId?: string) {
    const url = period ? `${this.baseUrl}/evidence/export?period=${encodeURIComponent(period)}` : `${this.baseUrl}/evidence/export`;
    const res = await fetch(url, {
      headers: this.getHeaders(firmId, clientId),
    });
    return res.json();
  }

  // ========================================================
  // MODULE 4: ANOMALY DETECTION ENGINE
  // ========================================================
  public static async getAnomalyRules(firmId?: string) {
    const res = await fetch(`${this.baseUrl}/anomalies/rules`, {
      headers: this.getHeaders(firmId),
    });
    return res.json();
  }

  public static async updateAnomalyRule(ruleId: string, payload: any, firmId?: string) {
    const res = await fetch(`${this.baseUrl}/anomalies/rules/${ruleId}`, {
      method: 'PUT',
      headers: this.getHeaders(firmId),
      body: JSON.stringify(payload),
    });
    return res.json();
  }

  public static async scanAnomalies(records: any[], firmId?: string, clientId?: string) {
    const res = await fetch(`${this.baseUrl}/anomalies/scan`, {
      method: 'POST',
      headers: this.getHeaders(firmId, clientId),
      body: JSON.stringify({ records }),
    });
    return res.json();
  }

  public static async getAnomalyAlerts(firmId?: string, clientId?: string) {
    const res = await fetch(`${this.baseUrl}/anomalies/alerts`, {
      headers: this.getHeaders(firmId, clientId),
    });
    return res.json();
  }

  public static async reviewAnomalyAlert(
    alertId: string,
    action: 'investigating' | 'dismissed_false_positive' | 'marked_for_adjustment',
    notes?: string,
    firmId?: string,
    clientId?: string
  ) {
    const res = await fetch(`${this.baseUrl}/anomalies/alerts/${alertId}/review`, {
      method: 'POST',
      headers: this.getHeaders(firmId, clientId),
      body: JSON.stringify({ action, notes }),
    });
    return res.json();
  }

  public static async exportAnomalyReport(firmId?: string, clientId?: string) {
    const res = await fetch(`${this.baseUrl}/anomalies/export`, {
      headers: this.getHeaders(firmId, clientId),
    });
    return res.json();
  }

  // ========================================================
  // MODULE 5: CLIENT DOCUMENT COLLECTION
  // ========================================================
  public static async createDocumentBundle(
    payload: { financialPeriod: string; engagementType: string; dueDate?: string },
    firmId?: string,
    clientId?: string
  ) {
    const res = await fetch(`${this.baseUrl}/document-requests/bundle`, {
      method: 'POST',
      headers: this.getHeaders(firmId, clientId),
      body: JSON.stringify(payload),
    });
    return res.json();
  }

  public static async getDocumentRequests(firmId?: string, clientId?: string) {
    const res = await fetch(`${this.baseUrl}/document-requests`, {
      headers: this.getHeaders(firmId, clientId),
    });
    return res.json();
  }

  public static async uploadClientDocument(
    requestId: string,
    payload: { fileName: string; fileSizeBytes?: number; fileHash?: string },
    firmId?: string,
    clientId?: string
  ) {
    const res = await fetch(`${this.baseUrl}/document-requests/${requestId}/upload`, {
      method: 'POST',
      headers: this.getHeaders(firmId, clientId),
      body: JSON.stringify(payload),
    });
    return res.json();
  }

  public static async draftReminderNotice(clientName?: string, email?: string, firmId?: string, clientId?: string) {
    const params = new URLSearchParams();
    if (clientName) params.append('clientName', clientName);
    if (email) params.append('email', email);

    const res = await fetch(`${this.baseUrl}/document-requests/reminder?${params.toString()}`, {
      headers: this.getHeaders(firmId, clientId),
    });
    return res.json();
  }

  public static async exportOutstandingDocuments(firmId?: string, clientId?: string) {
    const res = await fetch(`${this.baseUrl}/document-requests/export`, {
      headers: this.getHeaders(firmId, clientId),
    });
    return res.json();
  }

  // ========================================================
  // MODULE 6: RECURRING COMPLIANCE
  // ========================================================
  public static async getComplianceDefinitions(firmId?: string) {
    const res = await fetch(`${this.baseUrl}/compliance/definitions`, {
      headers: this.getHeaders(firmId),
    });
    return res.json();
  }

  public static async generateComplianceInstances(
    payload: { period: string; referenceYear: number; referenceMonth: number },
    firmId?: string,
    clientId?: string
  ) {
    const res = await fetch(`${this.baseUrl}/compliance/generate`, {
      method: 'POST',
      headers: this.getHeaders(firmId, clientId),
      body: JSON.stringify(payload),
    });
    return res.json();
  }

  public static async getComplianceInstances(currentDate?: string, firmId?: string, clientId?: string) {
    const url = currentDate ? `${this.baseUrl}/compliance/instances?currentDate=${encodeURIComponent(currentDate)}` : `${this.baseUrl}/compliance/instances`;
    const res = await fetch(url, {
      headers: this.getHeaders(firmId, clientId),
    });
    return res.json();
  }

  public static async completeComplianceTask(instanceId: string, proofRef?: string, firmId?: string, clientId?: string) {
    const res = await fetch(`${this.baseUrl}/compliance/instances/${instanceId}/complete`, {
      method: 'POST',
      headers: this.getHeaders(firmId, clientId),
      body: JSON.stringify({ proofRef }),
    });
    return res.json();
  }

  public static async exportComplianceCalendar(firmId?: string, clientId?: string) {
    const res = await fetch(`${this.baseUrl}/compliance/export`, {
      headers: this.getHeaders(firmId, clientId),
    });
    return res.json();
  }

  // ========================================================
  // MODULE 7: MULTI-CLIENT BATCH RUNS
  // ==========================================
  public static async startBatchRun(
    workflowId: string,
    clientItems: { clientId: string; clientName: string; files?: any[] }[],
    firmId?: string
  ) {
    const res = await fetch(`${this.baseUrl}/batch/run`, {
      method: 'POST',
      headers: this.getHeaders(firmId),
      body: JSON.stringify({ workflowId, clientItems }),
    });
    return res.json();
  }

  public static async getBatchRuns(firmId?: string) {
    const res = await fetch(`${this.baseUrl}/batch`, {
      headers: this.getHeaders(firmId),
    });
    return res.json();
  }

  public static async getBatchRun(batchId: string, firmId?: string) {
    const res = await fetch(`${this.baseUrl}/batch/${batchId}`, {
      headers: this.getHeaders(firmId),
    });
    return res.json();
  }

  public static async exportBatchSummary(batchId: string, firmId?: string) {
    const res = await fetch(`${this.baseUrl}/batch/${batchId}/export`, {
      headers: this.getHeaders(firmId),
    });
    return res.json();
  }

  public static getExportDownloadUrl(artifactId: string) {
    return `${this.baseUrl}/exports/${artifactId}/download`;
  }

  // ==========================================
  // MODULE 8: FORM 3CD TAX AUDIT (SEC 44AB)
  // ==========================================
  public static async scrutinizeForm3cd(
    vouchers: any[],
    financialYear: string = '2025-26',
    assessmentYear: string = '2026-27',
    firmId?: string,
    clientId?: string
  ) {
    const res = await fetch(`${this.baseUrl}/audit/form3cd/scrutinize`, {
      method: 'POST',
      headers: this.getHeaders(firmId, clientId),
      body: JSON.stringify({ vouchers, financialYear, assessmentYear }),
    });
    return res.json();
  }

  // ==========================================
  // MODULE 9: GSTR-1 VS 3B VS BOOKS 3-WAY REC
  // ==========================================
  public static async reconcileGstTriple(
    data: any[],
    financialYear: string = '2025-26',
    drc01bThreshold: number = 25000,
    firmId?: string,
    clientId?: string
  ) {
    const res = await fetch(`${this.baseUrl}/gst/triple-rec/reconcile`, {
      method: 'POST',
      headers: this.getHeaders(firmId, clientId),
      body: JSON.stringify({ data, financialYear, drc01bThreshold }),
    });
    return res.json();
  }

  public static async exportForm3cd(
    scrutinyResult: any,
    clientName: string = '',
    financialYear: string = '2025-26',
    assessmentYear: string = '2026-27',
    firmId?: string,
    clientId?: string
  ) {
    const res = await fetch(`${this.baseUrl}/audit/form3cd/export`, {
      method: 'POST',
      headers: this.getHeaders(firmId, clientId),
      body: JSON.stringify({ scrutinyResult, clientName, financialYear, assessmentYear }),
    });
    return res.json();
  }

  public static async exportGstTriple(
    tripleRecResult: any,
    clientName: string = '',
    financialYear: string = '2025-26',
    firmId?: string,
    clientId?: string
  ) {
    const res = await fetch(`${this.baseUrl}/gst/triple-rec/export`, {
      method: 'POST',
      headers: this.getHeaders(firmId, clientId),
      body: JSON.stringify({ tripleRecResult, clientName, financialYear }),
    });
    return res.json();
  }

  // ==========================================
  // MODULE 10: SECTION 206AB / 206CCA AUDIT
  // ==========================================
  public static async evaluateSection206ab(
    vendors: any[],
    clientName: string = '',
    financialYear: string = 'FY 2026-27',
    firmId?: string,
    clientId?: string
  ) {
    const res = await fetch(`${this.baseUrl}/audit/section206ab/evaluate`, {
      method: 'POST',
      headers: this.getHeaders(firmId, clientId),
      body: JSON.stringify({ vendors, clientName, financialYear }),
    });
    return res.json();
  }

  public static async exportSection206ab(
    summary: any,
    firmId?: string,
    clientId?: string
  ) {
    const res = await fetch(`${this.baseUrl}/audit/section206ab/export`, {
      method: 'POST',
      headers: this.getHeaders(firmId, clientId),
      body: JSON.stringify({ summary }),
    });
    return res.json();
  }

  // ==========================================
  // MODULE 11: MEMORY-BOUNDED STREAMING
  // ==========================================
  public static async streamProcessBuffer(
    base64: string,
    fileName: string,
    chunkSize: number = 2500,
    firmId?: string,
    clientId?: string
  ) {
    const res = await fetch(`${this.baseUrl}/pipeline/stream-process`, {
      method: 'POST',
      headers: this.getHeaders(firmId, clientId),
      body: JSON.stringify({ base64, fileName, chunkSize }),
    });
    return res.json();
  }
}


