import React, { useState, useEffect } from 'react';
import {
  ClipboardCheck,
  CreditCard,
  FileCheck2,
  AlertTriangle,
  Mail,
  CalendarCheck,
  Layers,
  Lock,
  Unlock,
  CheckCircle2,
  Clock,
  Download,
  Plus,
  RefreshCw,
  Search,
  ExternalLink,
  ShieldAlert,
  ChevronRight,
  FileText,
  Copy,
  AlertOctagon,
  Building,
  UserCheck,
  Send,
  Sliders,
  X,
  Sparkles,
  ArrowLeft,
} from 'lucide-react';
import { CAAutomationApiClient } from '../../services/apiClient';

interface AdvancedModulesHubProps {
  initialModule?: ActiveModule;
  onBack?: () => void;
  selectedClientId?: string;
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
}

type ActiveModule =
  | 'closing'
  | 'apar'
  | 'evidence'
  | 'anomalies'
  | 'document-requests'
  | 'compliance'
  | 'batch';

export const AdvancedModulesHub: React.FC<AdvancedModulesHubProps> = ({
  initialModule = 'closing',
  onBack,
  selectedClientId = '',
  onShowToast,
}) => {
  const [activeModule, setActiveModule] = useState<ActiveModule>(initialModule);
  const [isLoading, setIsLoading] = useState(false);


  // Module 1: Month-End Closing State
  const [closingPeriods, setClosingPeriods] = useState<any[]>([]);
  const [selectedPeriod, setSelectedPeriod] = useState<any | null>(null);
  const [isNewPeriodModalOpen, setIsNewPeriodModalOpen] = useState(false);
  const [newPeriodName, setNewPeriodName] = useState('September 2026');
  const [newPeriodFy, setNewPeriodFy] = useState('FY 2026-27');
  const [closingApprovalNotes, setClosingApprovalNotes] = useState('');

  // Module 2: AP & AR State
  const [aparType, setAparType] = useState<'payable' | 'receivable'>('payable');
  const [agingRecords, setAgingRecords] = useState<any[]>([]);
  const [paymentAlerts, setPaymentAlerts] = useState<any[]>([]);

  // Module 3: Audit Evidence State
  const [evidenceList, setEvidenceList] = useState<any[]>([]);
  const [isLinkModalOpen, setIsLinkModalOpen] = useState(false);
  const [targetEvidenceId, setTargetEvidenceId] = useState('');
  const [linkVoucherNo, setLinkVoucherNo] = useState('');
  const [linkVoucherAmount, setLinkVoucherAmount] = useState('');
  const [linkPartyName, setLinkPartyName] = useState('');
  const [missingEvidenceThreshold, setMissingEvidenceThreshold] = useState(100000);
  const [missingEvidenceResult, setMissingEvidenceResult] = useState<any | null>(null);

  // Module 4: Anomaly Detection State
  const [anomalyRules, setAnomalyRules] = useState<any[]>([]);
  const [anomalyAlerts, setAnomalyAlerts] = useState<any[]>([]);
  const [selectedAlertForReview, setSelectedAlertForReview] = useState<any | null>(null);
  const [alertReviewNotes, setAlertReviewNotes] = useState('');

  // Module 5: Document Collection State
  const [docRequests, setDocRequests] = useState<any[]>([]);
  const [reminderDraft, setReminderDraft] = useState<any | null>(null);
  const [uploadTargetRequestId, setUploadTargetRequestId] = useState<string | null>(null);
  const [simulatedFileName, setSimulatedFileName] = useState('');

  // Module 6: Recurring Compliance State
  const [complianceDefs, setComplianceDefs] = useState<any[]>([]);
  const [complianceInstances, setComplianceInstances] = useState<any[]>([]);
  const [selectedTaskToComplete, setSelectedTaskToComplete] = useState<any | null>(null);
  const [completionProofRef, setCompletionProofRef] = useState('');

  // Module 7: Multi-Client Batch State
  const [batchRuns, setBatchRuns] = useState<any[]>([]);
  const [selectedBatchWorkflow, setSelectedBatchWorkflow] = useState('wf_bank_rec');
  const [selectedBatchClients, setSelectedBatchClients] = useState<string[]>([
    '',
    'client_bharat_logistics',
    'client_coastal_trading',
  ]);

  // Load initial data on mount or tab change
  useEffect(() => {
    loadModuleData(activeModule);
  }, [activeModule, selectedClientId]);

  const loadModuleData = async (mod: ActiveModule) => {
    setIsLoading(true);
    try {
      if (mod === 'closing') {
        const res = await CAAutomationApiClient.getClosingPeriods(undefined, selectedClientId);
        if (res.success && res.data) {
          setClosingPeriods(res.data);
          if (res.data.length > 0) {
            loadPeriodDetails(res.data[0].id);
          }
        }
      } else if (mod === 'apar') {
        const [recRes, alertRes] = await Promise.all([
          CAAutomationApiClient.getAparRecords(undefined, selectedClientId),
          CAAutomationApiClient.getPaymentAlerts(undefined, selectedClientId),
        ]);
        if (recRes.success) setAgingRecords(recRes.data || []);
        if (alertRes.success) setPaymentAlerts(alertRes.data || []);
      } else if (mod === 'evidence') {
        const res = await CAAutomationApiClient.getEvidenceList(undefined, selectedClientId);
        if (res.success && res.data) {
          if (res.data.length === 0) {
            // auto-init standard slots
            const initRes = await CAAutomationApiClient.initEvidenceSlots('September 2026', undefined, selectedClientId);
            if (initRes.success) setEvidenceList(initRes.data || []);
          } else {
            setEvidenceList(res.data);
          }
        }
      } else if (mod === 'anomalies') {
        const [rulesRes, alertsRes] = await Promise.all([
          CAAutomationApiClient.getAnomalyRules(),
          CAAutomationApiClient.getAnomalyAlerts(undefined, selectedClientId),
        ]);
        if (rulesRes.success) setAnomalyRules(rulesRes.data || []);
        if (alertsRes.success) setAnomalyAlerts(alertsRes.data || []);
      } else if (mod === 'document-requests') {
        const res = await CAAutomationApiClient.getDocumentRequests(undefined, selectedClientId);
        if (res.success && res.data) {
          if (res.data.length === 0) {
            // auto-generate statutory bundle
            const bundleRes = await CAAutomationApiClient.createDocumentBundle(
              { financialPeriod: 'September 2026', engagementType: 'Statutory Audit Checklist' },
              undefined,
              selectedClientId
            );
            if (bundleRes.success) setDocRequests(bundleRes.data || []);
          } else {
            setDocRequests(res.data);
          }
        }
      } else if (mod === 'compliance') {
        const [defsRes, instRes] = await Promise.all([
          CAAutomationApiClient.getComplianceDefinitions(),
          CAAutomationApiClient.getComplianceInstances(undefined, undefined, selectedClientId),
        ]);
        if (defsRes.success) setComplianceDefs(defsRes.data || []);
        if (instRes.success && instRes.data) {
          if (instRes.data.length === 0) {
            const genRes = await CAAutomationApiClient.generateComplianceInstances(
              { period: 'September 2026', referenceYear: 2026, referenceMonth: 9 },
              undefined,
              selectedClientId
            );
            if (genRes.success) setComplianceInstances(genRes.data || []);
          } else {
            setComplianceInstances(instRes.data);
          }
        }
      } else if (mod === 'batch') {
        const res = await CAAutomationApiClient.getBatchRuns();
        if (res.success) setBatchRuns(res.data || []);
      }
    } catch (err: any) {
      console.error('Error loading module data:', err);
    } finally {
      setIsLoading(false);
    }
  };

  const loadPeriodDetails = async (periodId: string) => {
    try {
      const res = await CAAutomationApiClient.evaluateClosingPeriod(periodId, undefined, selectedClientId);
      if (res.success && res.data) {
        setSelectedPeriod(res.data);
      }
    } catch (err) {
      console.error(err);
    }
  };

  // ==========================================
  // MODULE 1 HANDLERS
  // ==========================================
  const handleCreatePeriod = async () => {
    try {
      const res = await CAAutomationApiClient.createClosingPeriod(
        { periodName: newPeriodName, financialYear: newPeriodFy },
        undefined,
        selectedClientId
      );
      if (res.success) {
        onShowToast('Closing Period Created', `${newPeriodName} initialized with standard checklist`, 'success');
        setIsNewPeriodModalOpen(false);
        loadModuleData('closing');
      }
    } catch (err: any) {
      onShowToast('Failed to Create Period', err.message, 'error');
    }
  };

  const handleUpdateTaskStatus = async (taskId: string, newStatus: string) => {
    if (!selectedPeriod) return;
    try {
      const res = await CAAutomationApiClient.updateClosingTask(
        selectedPeriod.period.id,
        taskId,
        { status: newStatus, completionNotes: `Updated via Executive Closing Hub at ${new Date().toLocaleTimeString()}` },
        undefined,
        selectedClientId
      );
      if (res.success) {
        onShowToast('Checklist Updated', 'Task status recorded', 'info');
        loadPeriodDetails(selectedPeriod.period.id);
      }
    } catch (err: any) {
      onShowToast('Update Failed', err.message, 'error');
    }
  };

  const handleApproveClosing = async () => {
    if (!selectedPeriod) return;
    try {
      const res = await CAAutomationApiClient.approveClosing(
        selectedPeriod.period.id,
        closingApprovalNotes || 'All prerequisites verified and sign-off approved by Partner.',
        undefined,
        selectedClientId
      );
      if (res.success) {
        onShowToast('Closing Approved & Locked', 'Financial period status is now officially closed', 'success');
        loadPeriodDetails(selectedPeriod.period.id);
      } else {
        onShowToast('Closing Approval Blocked', res.error, 'error');
      }
    } catch (err: any) {
      onShowToast('Approval Error', err.message, 'error');
    }
  };

  const handleExportClosing = async () => {
    if (!selectedPeriod) return;
    try {
      const res = await CAAutomationApiClient.exportClosingSummary(selectedPeriod.period.id, undefined, selectedClientId);
      if (res.success && res.data) {
        window.open(CAAutomationApiClient.getExportDownloadUrl(res.data.id), '_blank');
        onShowToast('Report Downloaded', res.data.fileName, 'success');
      }
    } catch (err: any) {
      onShowToast('Export Failed', err.message, 'error');
    }
  };

  // ==========================================
  // MODULE 2 HANDLERS
  // ==========================================

  const handleExportPaymentReview = async () => {
    try {
      const res = await CAAutomationApiClient.exportPaymentReviewList(undefined, selectedClientId);
      if (res.success && res.data) {
        window.open(CAAutomationApiClient.getExportDownloadUrl(res.data.id), '_blank');
        onShowToast('Payment Review Downloaded', res.data.fileName, 'success');
      }
    } catch (err: any) {
      onShowToast('Export Failed', err.message, 'error');
    }
  };

  // ==========================================
  // MODULE 3 HANDLERS
  // ==========================================
  const handleLinkTransaction = async () => {
    if (!targetEvidenceId || !linkVoucherNo) return;
    try {
      const res = await CAAutomationApiClient.linkEvidenceTransaction(
        targetEvidenceId,
        {
          voucherNo: linkVoucherNo,
          date: '2026-09-15',
          amount: parseFloat(linkVoucherAmount) || 0,
          partyName: linkPartyName || 'M/s Verified Vendor',
          accountHead: 'Audit Evidence Verification',
        },
        undefined,
        selectedClientId
      );

      if (res.success) {
        onShowToast('Voucher Linked to Evidence', `Linked ${linkVoucherNo} to working paper slot`, 'success');
        setIsLinkModalOpen(false);
        setLinkVoucherNo('');
        setLinkVoucherAmount('');
        loadModuleData('evidence');
      }
    } catch (err: any) {
      onShowToast('Link Failed', err.message, 'error');
    }
  };

  const handleScanMissingEvidence = async () => {
    try {
      const sampleTxns = [
        { voucherNo: 'VR-2026-901', date: '2026-09-10', amount: 450000, partyName: 'Reliance Power Infra', accountHead: 'Capital Equipment' },
        { voucherNo: 'VR-2026-902', date: '2026-09-12', amount: 280000, partyName: 'Adani Logistics Hub', accountHead: 'Freight Charges' },
        { voucherNo: 'VR-2026-903', date: '2026-09-14', amount: 65000, partyName: 'Office Depot Stationery', accountHead: 'Printing & Stationery' },
      ];

      const res = await CAAutomationApiClient.detectMissingEvidence(
        { transactions: sampleTxns, materialityThreshold: missingEvidenceThreshold },
        undefined,
        selectedClientId
      );

      if (res.success) {
        setMissingEvidenceResult(res.data);
        onShowToast('Materiality Scan Complete', `Identified ${res.data.missingEvidenceCount} items requiring supporting documentation`, 'info');
      }
    } catch (err: any) {
      onShowToast('Scan Failed', err.message, 'error');
    }
  };

  const handleExportEvidence = async () => {
    try {
      const res = await CAAutomationApiClient.exportEvidenceRegister('September 2026', undefined, selectedClientId);
      if (res.success && res.data) {
        window.open(CAAutomationApiClient.getExportDownloadUrl(res.data.id), '_blank');
        onShowToast('Evidence Register Downloaded', res.data.fileName, 'success');
      }
    } catch (err: any) {
      onShowToast('Export Failed', err.message, 'error');
    }
  };

  // ==========================================
  // MODULE 4 HANDLERS
  // ==========================================
  const handleRunAnomalyScan = async () => {
    try {
      const sampleJournals = [
        { voucherNo: 'JV-2026-501', date: '2026-09-15', accountHead: 'Audit & Legal Fees', debit: 50000, credit: 0, description: 'Legal retainer fees' },
        { voucherNo: 'JV-2026-502', date: '2026-09-15', accountHead: 'Audit & Legal Fees', debit: 50000, credit: 0, description: 'Legal retainer fees duplicate entry' },
        { voucherNo: 'VR-CASH-101', date: '2026-09-18', accountHead: 'Factory Site Repairs', debit: 25000, credit: 0, description: 'Cash paid to site contractor exceeding 40A(3)' },
        { voucherNo: 'BNK-SUN-001', date: '2026-09-20', accountHead: 'Director Remuneration', debit: 1500000, credit: 0, description: 'Sunday wire payment release' },
        { voucherNo: 'CB-NEG-001', date: '2026-09-25', accountHead: 'Cash in Hand Account', debit: 0, credit: 15000, balance: -15000, description: 'Negative cash closing balance' },
      ];

      const res = await CAAutomationApiClient.scanAnomalies(sampleJournals, undefined, selectedClientId);
      if (res.success) {
        onShowToast('Anomaly Scan Complete', `Identified ${res.count} suspicious transactions based on deterministic rules`, 'success');
        loadModuleData('anomalies');
      }
    } catch (err: any) {
      onShowToast('Scan Failed', err.message, 'error');
    }
  };

  const handleReviewAlert = async (action: 'investigating' | 'dismissed_false_positive' | 'marked_for_adjustment') => {
    if (!selectedAlertForReview) return;
    try {
      const res = await CAAutomationApiClient.reviewAnomalyAlert(
        selectedAlertForReview.id,
        action,
        alertReviewNotes || `Reviewed by Auditor with status: ${action}`,
        undefined,
        selectedClientId
      );
      if (res.success) {
        onShowToast('Alert Reviewed', `Anomaly marked as ${action.replace(/_/g, ' ')}`, 'info');
        setSelectedAlertForReview(null);
        setAlertReviewNotes('');
        loadModuleData('anomalies');
      }
    } catch (err: any) {
      onShowToast('Review Failed', err.message, 'error');
    }
  };

  const handleExportAnomalies = async () => {
    try {
      const res = await CAAutomationApiClient.exportAnomalyReport(undefined, selectedClientId);
      if (res.success && res.data) {
        window.open(CAAutomationApiClient.getExportDownloadUrl(res.data.id), '_blank');
        onShowToast('Anomaly Report Downloaded', res.data.fileName, 'success');
      }
    } catch (err: any) {
      onShowToast('Export Failed', err.message, 'error');
    }
  };

  // ==========================================
  // MODULE 5 HANDLERS
  // ==========================================
  const handleUploadSimulatedDoc = async () => {
    if (!uploadTargetRequestId || !simulatedFileName) return;
    try {
      const res = await CAAutomationApiClient.uploadClientDocument(
        uploadTargetRequestId,
        {
          fileName: simulatedFileName,
          fileSizeBytes: 245000,
          fileHash: 'sha256_' + Math.random().toString(36).substring(2, 12),
        },
        undefined,
        selectedClientId
      );

      if (res.success) {
        onShowToast('Document Registered', `${simulatedFileName} registered with verification hash`, 'success');
        setUploadTargetRequestId(null);
        setSimulatedFileName('');
        loadModuleData('document-requests');
      }
    } catch (err: any) {
      onShowToast('Upload Failed', err.message, 'error');
    }
  };

  const handleDraftReminder = async () => {
    try {
      const res = await CAAutomationApiClient.draftReminderNotice('Client Entity', 'finance@client.in', undefined, selectedClientId);
      if (res.success && res.data) {
        setReminderDraft(res.data);
        onShowToast('Reminder Draft Prepared', 'Ready for Partner review before dispatch', 'info');
      }
    } catch (err: any) {
      onShowToast('Draft Failed', err.message, 'error');
    }
  };

  const handleExportDocuments = async () => {
    try {
      const res = await CAAutomationApiClient.exportOutstandingDocuments(undefined, selectedClientId);
      if (res.success && res.data) {
        window.open(CAAutomationApiClient.getExportDownloadUrl(res.data.id), '_blank');
        onShowToast('Document Register Downloaded', res.data.fileName, 'success');
      }
    } catch (err: any) {
      onShowToast('Export Failed', err.message, 'error');
    }
  };

  // ==========================================
  // MODULE 6 HANDLERS
  // ==========================================
  const handleCompleteCompliance = async () => {
    if (!selectedTaskToComplete) return;
    try {
      const res = await CAAutomationApiClient.completeComplianceTask(
        selectedTaskToComplete.id,
        completionProofRef || 'ARN_AA270926012485C',
        undefined,
        selectedClientId
      );

      if (res.success) {
        onShowToast('Compliance Task Completed', `Filing proof recorded: ${res.data.statutoryFilingProof}`, 'success');
        setSelectedTaskToComplete(null);
        setCompletionProofRef('');
        loadModuleData('compliance');
      }
    } catch (err: any) {
      onShowToast('Completion Failed', err.message, 'error');
    }
  };

  const handleExportCompliance = async () => {
    try {
      const res = await CAAutomationApiClient.exportComplianceCalendar(undefined, selectedClientId);
      if (res.success && res.data) {
        window.open(CAAutomationApiClient.getExportDownloadUrl(res.data.id), '_blank');
        onShowToast('Compliance Calendar Downloaded', res.data.fileName, 'success');
      }
    } catch (err: any) {
      onShowToast('Export Failed', err.message, 'error');
    }
  };

  // ==========================================
  // MODULE 7 HANDLERS
  // ==========================================
  const handleStartBatchRun = async () => {
    try {
      const clientItems = selectedBatchClients.map((cId) => ({
        clientId: cId,
        clientName: cId.replace('client_', '').replace(/_/g, ' ').toUpperCase(),
        files: [
          { fileName: 'bank_statement.csv', content: 'Date,Particulars,Debit,Credit,Balance\n2026-09-01,Opening,0,0,100000' },
          { fileName: 'general_ledger.csv', content: 'Date,Particulars,Debit,Credit,Balance\n2026-09-01,Opening,0,0,100000' },
        ],
      }));

      const res = await CAAutomationApiClient.startBatchRun(selectedBatchWorkflow, clientItems);
      if (res.success && res.data) {
        onShowToast('Multi-Client Processing Started', `Running ${clientItems.length} clients concurrently with isolated data containers`, 'success');
        loadModuleData('batch');
      }
    } catch (err: any) {
      onShowToast('Processing Failed to Start', err.message, 'error');
    }
  };

  const handleExportBatchSummary = async (batchId: string) => {
    try {
      const res = await CAAutomationApiClient.exportBatchSummary(batchId);
      if (res.success && res.data) {
        window.open(CAAutomationApiClient.getExportDownloadUrl(res.data.id), '_blank');
        onShowToast('Consolidated Summary Downloaded', res.data.fileName, 'success');
      }
    } catch (err: any) {
      onShowToast('Export Failed', err.message, 'error');
    }
  };

  // Navigation tabs configuration
  const modules = [
    { id: 'closing', label: '1. Month-End Closing', icon: ClipboardCheck, badge: 'Workflow' },
    { id: 'apar', label: '2. AP & AR Aging', icon: CreditCard, badge: 'MSMED 45d' },
    { id: 'evidence', label: '3. Audit Evidence', icon: FileCheck2, badge: 'SA 500' },
    { id: 'anomalies', label: '4. Anomaly Engine', icon: AlertTriangle, badge: 'Deterministic' },
    { id: 'document-requests', label: '5. Document Tracker', icon: Mail, badge: 'Requisitions' },
    { id: 'compliance', label: '6. Recurring Compliance', icon: CalendarCheck, badge: 'Statutory' },
    { id: 'batch', label: '7. Multi-Client Processing', icon: Layers, badge: 'Multi-Client' },
  ];

  return (
    <div className="space-y-5">
      {/* Executive Header Banner */}
      <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-xs">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="flex items-start gap-3">
            {onBack && (
              <button
                onClick={onBack}
                className="mt-1 p-2 text-gray-500 hover:text-gray-900 bg-white border border-gray-200 rounded-lg hover:bg-gray-50 transition-colors cursor-pointer shrink-0 shadow-2xs"
                title="Back to All Tools"
              >
                <ArrowLeft className="w-4 h-4" />
              </button>
            )}
            <div>
              <div className="flex items-center gap-2">
                <span className="px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-gray-100 text-gray-900 border border-gray-200">
                  Enterprise Firm Edition
                </span>
                <span className="text-xs text-gray-500 font-medium">Firm Multi-Tenancy Active</span>
              </div>
              <h1 className="text-xl font-bold text-gray-900 mt-1.5 tracking-tight flex items-center gap-2">
                Advanced CA Firm Automation Suite
              </h1>
              <p className="text-xs text-gray-600 mt-0.5 max-w-3xl">
                End-to-end accounting orchestration: Closing workflows, AP/AR aging, audit evidence linking, deterministic anomaly detection, client requisitions, compliance calendars, and isolated multi-client automated runs.
              </p>
            </div>
          </div>


          <div className="flex items-center gap-2">
            <button
              onClick={() => loadModuleData(activeModule)}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-gray-700 bg-gray-50 hover:bg-gray-100 border border-gray-200 rounded-lg transition-colors cursor-pointer"
              title="Refresh Module Data"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${isLoading ? 'animate-spin text-gray-900' : 'text-gray-500'}`} />
              <span>Refresh</span>
            </button>
          </div>
        </div>

        {/* Module Segmented Tab Bar */}
        <div className="mt-5 border-t border-gray-100 pt-3 flex items-center gap-1.5 overflow-x-auto pb-1 scrollbar-none">
          {modules.map((m) => {
            const Icon = m.icon;
            const isActive = activeModule === m.id;
            return (
              <button
                key={m.id}
                onClick={() => setActiveModule(m.id as ActiveModule)}
                className={`flex items-center gap-2 px-3.5 py-2 rounded-lg text-xs font-semibold whitespace-nowrap transition-all cursor-pointer ${
                  isActive
                    ? 'bg-gray-900 text-white shadow-xs'
                    : 'text-gray-600 hover:text-gray-900 hover:bg-gray-50'
                }`}
              >
                <Icon className={`w-3.5 h-3.5 ${isActive ? 'text-white' : 'text-gray-400'}`} />
                <span>{m.label}</span>
                <span
                  className={`text-[9px] px-1.5 py-0.2 rounded-full font-medium ${
                    isActive ? 'bg-gray-900 text-white' : 'bg-gray-100 text-gray-600'
                  }`}
                >
                  {m.badge}
                </span>
              </button>
            );
          })}
        </div>
      </div>

      {/* ========================================================================= */}
      {/* MODULE 1: MONTH-END CLOSING WORKFLOW */}
      {/* ========================================================================= */}
      {activeModule === 'closing' && (
        <div className="space-y-4">
          {/* Top Period Selector & Actions */}
          <div className="bg-white rounded-xl border border-gray-200 p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div className="flex items-center gap-3">
              <div>
                <label className="text-[11px] font-semibold text-gray-500 uppercase">Active Financial Period</label>
                <select
                  value={selectedPeriod?.period?.id || ''}
                  onChange={(e) => loadPeriodDetails(e.target.value)}
                  className="mt-0.5 block w-64 rounded-md border border-gray-300 bg-white px-3 py-1.5 text-xs font-medium text-gray-900 shadow-2xs focus:border-gray-900 focus:outline-none"
                >
                  {closingPeriods.map((p) => (
                    <option key={p.id} value={p.id}>
                      {p.periodName} ({p.financialYear}) — {p.status.toUpperCase()}
                    </option>
                  ))}
                </select>
              </div>

              {selectedPeriod && (
                <div className="pt-4 flex items-center gap-2">
                  <span
                    className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-semibold ${
                      selectedPeriod.period.status === 'closed'
                        ? 'bg-gray-100 text-gray-900 border border-gray-300'
                        : selectedPeriod.period.status === 'review_pending'
                        ? 'bg-blue-50 text-blue-800 border border-blue-200'
                        : 'bg-amber-100 text-amber-800 border border-amber-200'
                    }`}
                  >
                    {selectedPeriod.period.status === 'closed' ? (
                      <Lock className="w-3 h-3 text-gray-900" />
                    ) : (
                      <Clock className="w-3 h-3 text-amber-700" />
                    )}
                    {selectedPeriod.period.status.toUpperCase()}
                  </span>
                </div>
              )}
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={() => setIsNewPeriodModalOpen(true)}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-gray-900 bg-gray-100 hover:bg-gray-100 border border-gray-200 rounded-lg transition-colors cursor-pointer"
              >
                <Plus className="w-3.5 h-3.5" />
                <span>New Period</span>
              </button>

              <button
                onClick={handleExportClosing}
                disabled={!selectedPeriod}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-gray-700 bg-white hover:bg-gray-50 border border-gray-300 rounded-lg transition-colors cursor-pointer disabled:opacity-50"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Export Summary</span>
              </button>
            </div>
          </div>

          {/* Checklist Progress Overview */}
          {selectedPeriod ? (
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-5">
              {/* Left 2 Cols: Sequential Checklist Graph */}
              <div className="lg:col-span-2 bg-white rounded-xl border border-gray-200 p-5 space-y-4">
                <div className="flex items-center justify-between border-b border-gray-100 pb-3">
                  <div>
                    <h3 className="text-sm font-bold text-gray-900">Closing Prerequisite Graph</h3>
                    <p className="text-xs text-gray-500">Prerequisites are strictly enforced before dependent checks can be completed</p>
                  </div>
                  <span className="text-xs font-semibold text-gray-700 bg-gray-100 px-2 py-0.5 rounded">
                    Score: {selectedPeriod.scorePercent}% Complete
                  </span>
                </div>

                <div className="space-y-3">
                  {selectedPeriod.period.checklist.map((task: any, idx: number) => {
                    const isClosed = selectedPeriod.period.status === 'closed';
                    return (
                      <div
                        key={task.id}
                        className={`p-3.5 rounded-lg border transition-all ${
                          task.status === 'completed'
                            ? 'bg-gray-100 border-gray-200'
                            : task.status === 'in_progress'
                            ? 'bg-blue-50/40 border-blue-200'
                            : 'bg-gray-50/70 border-gray-200'
                        }`}
                      >
                        <div className="flex items-start justify-between gap-3">
                          <div className="flex items-start gap-3">
                            <span className="w-5 h-5 rounded-full bg-white border border-gray-300 flex items-center justify-center text-[10px] font-bold text-gray-700 shrink-0 mt-0.5">
                              {idx + 1}
                            </span>
                            <div>
                              <h4 className="text-xs font-bold text-gray-900 flex items-center gap-2">
                                {task.taskName}
                                {task.required && (
                                  <span className="text-[10px] text-red-600 bg-red-50 border border-red-200 px-1 rounded font-normal">
                                    Mandatory
                                  </span>
                                )}
                              </h4>
                              <p className="text-[11px] text-gray-600 mt-0.5">{task.description}</p>
                              {task.prerequisiteTaskId && (
                                <p className="text-[10px] text-amber-700 mt-1 font-medium flex items-center gap-1">
                                  <span>Prerequisite:</span>
                                  <span className="underline">{task.prerequisiteTaskId}</span>
                                </p>
                              )}
                              {task.completionNotes && (
                                <p className="text-[11px] text-gray-500 italic mt-1 bg-white p-1.5 rounded border border-gray-200">
                                  Note: {task.completionNotes}
                                </p>
                              )}
                            </div>
                          </div>

                          <div className="shrink-0 flex items-center gap-1.5">
                            {!isClosed && (
                              <select
                                value={task.status}
                                onChange={(e) => handleUpdateTaskStatus(task.id, e.target.value)}
                                className="text-xs border border-gray-300 rounded px-2 py-1 bg-white font-medium focus:outline-none"
                              >
                                <option value="pending">Pending</option>
                                <option value="in_progress">In Progress</option>
                                <option value="completed">Completed</option>
                                <option value="waived_with_approval">Waived (Partner Approved)</option>
                              </select>
                            )}
                            {isClosed && (
                              <span className="text-xs font-semibold text-gray-900 bg-gray-100 px-2 py-1 rounded">
                                {task.status.toUpperCase()}
                              </span>
                            )}
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Right Col: Sign-off & Audit Lock Panel */}
              <div className="space-y-4">
                <div className="bg-white rounded-xl border border-gray-200 p-5 space-y-4">
                  <h3 className="text-sm font-bold text-gray-900 flex items-center gap-2">
                    <ShieldAlert className="w-4 h-4 text-gray-900" />
                    <span>Audit Sign-Off & Lock</span>
                  </h3>

                  {selectedPeriod.period.status === 'closed' ? (
                    <div className="bg-gray-100 border border-gray-200 rounded-lg p-3.5 space-y-2">
                      <div className="flex items-center gap-2 text-gray-900 font-bold text-xs">
                        <Lock className="w-4 h-4 text-gray-900" />
                        <span>PERIOD OFFICIALLY CLOSED & LOCKED</span>
                      </div>
                      <p className="text-[11px] text-gray-900">
                        Approved By: {selectedPeriod.period.approvedBy}
                      </p>
                      <p className="text-[11px] text-gray-900">
                        Approval Date: {selectedPeriod.period.approvedAt}
                      </p>
                      <p className="text-[11px] text-gray-900 italic border-t border-gray-200 pt-1">
                        "{selectedPeriod.period.approvalNotes}"
                      </p>
                    </div>
                  ) : (
                    <div className="space-y-3">
                      <div>
                        <label className="text-xs font-semibold text-gray-700 block mb-1">
                          Partner Sign-Off Notes
                        </label>
                        <textarea
                          rows={3}
                          value={closingApprovalNotes}
                          onChange={(e) => setClosingApprovalNotes(e.target.value)}
                          placeholder="Document verification checks, exception justifications, and closing approval sign-off..."
                          className="w-full text-xs rounded-lg border border-gray-300 p-2.5 focus:border-gray-900 focus:outline-none"
                        />
                      </div>

                      <div className="p-3 bg-amber-50 rounded-lg border border-amber-200 text-[11px] text-amber-900 space-y-1">
                        <p className="font-bold">Prerequisite Requirements:</p>
                        <ul className="list-disc pl-4 space-y-0.5">
                          <li>All mandatory checklist items must be marked completed.</li>
                          <li>Critical exceptions must be resolved.</li>
                          <li>Cannot modify historical transactions after closing.</li>
                        </ul>
                      </div>

                      <button
                        onClick={handleApproveClosing}
                        className="w-full flex items-center justify-center gap-2 px-4 py-2.5 text-xs font-bold text-white bg-gray-900 hover:bg-black rounded-lg shadow-xs transition-colors cursor-pointer"
                      >
                        <Lock className="w-4 h-4" />
                        <span>Approve & Lock Financial Period</span>
                      </button>
                    </div>
                  )}
                </div>

                <div className="bg-white rounded-xl border border-gray-200 p-4 space-y-2 text-xs">
                  <h4 className="font-bold text-gray-900">Period Information</h4>
                  <div className="flex justify-between py-1 border-b border-gray-100">
                    <span className="text-gray-500">Target Closing:</span>
                    <span className="font-medium text-gray-800">{selectedPeriod.period.targetClosingDate}</span>
                  </div>
                  <div className="flex justify-between py-1 border-b border-gray-100">
                    <span className="text-gray-500">Financial Year:</span>
                    <span className="font-medium text-gray-800">{selectedPeriod.period.financialYear}</span>
                  </div>
                  <div className="flex justify-between py-1">
                    <span className="text-gray-500">Client ID:</span>
                    <span className="font-mono text-gray-800">{selectedPeriod.period.clientId}</span>
                  </div>
                </div>
              </div>
            </div>
          ) : (
            <div className="bg-white rounded-xl border border-gray-200 p-10 text-center space-y-3">
              <ClipboardCheck className="w-10 h-10 text-gray-300 mx-auto" />
              <h3 className="text-sm font-semibold text-gray-800">No Closing Period Initialized</h3>
              <p className="text-xs text-gray-500 max-w-sm mx-auto">
                Create a financial closing period to initialize the prerequisite checklist and audit sign-off workflow.
              </p>
              <button
                onClick={() => setIsNewPeriodModalOpen(true)}
                className="px-4 py-2 text-xs font-semibold text-white bg-gray-900 hover:bg-black rounded-lg transition-colors cursor-pointer inline-flex items-center gap-1.5"
              >
                <Plus className="w-4 h-4" />
                <span>Initialize September 2026 Period</span>
              </button>
            </div>
          )}
        </div>
      )}

      {/* ========================================================================= */}
      {/* MODULE 2: ACCOUNTS PAYABLE & RECEIVABLE AUTOMATION */}
      {/* ========================================================================= */}
      {activeModule === 'apar' && (
        <div className="space-y-4">
          <div className="bg-white rounded-xl border border-gray-200 p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div className="flex items-center gap-3">
              <div className="inline-flex p-1 bg-gray-100 rounded-lg">
                <button
                  onClick={() => setAparType('payable')}
                  className={`px-3 py-1.5 text-xs font-bold rounded-md transition-all cursor-pointer ${
                    aparType === 'payable' ? 'bg-white text-gray-900 shadow-2xs' : 'text-gray-600'
                  }`}
                >
                  Accounts Payable (Vendors)
                </button>
                <button
                  onClick={() => setAparType('receivable')}
                  className={`px-3 py-1.5 text-xs font-bold rounded-md transition-all cursor-pointer ${
                    aparType === 'receivable' ? 'bg-white text-gray-900 shadow-2xs' : 'text-gray-600'
                  }`}
                >
                  Accounts Receivable (Customers)
                </button>
              </div>

              
            </div>

            <button
              onClick={handleExportPaymentReview}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-gray-700 bg-white hover:bg-gray-50 border border-gray-300 rounded-lg transition-colors cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Export Payment Review List</span>
            </button>
          </div>

          {/* Statutory Alerts (MSMED 45-Day & Duplicate Risks) */}
          {paymentAlerts.length > 0 && (
            <div className="space-y-2">
              {paymentAlerts.map((alt) => (
                <div
                  key={alt.id}
                  className={`p-3.5 rounded-lg border flex items-start gap-3 ${
                    alt.severity === 'high' ? 'bg-red-50 border-red-200' : 'bg-amber-50 border-amber-200'
                  }`}
                >
                  <AlertOctagon
                    className={`w-4 h-4 shrink-0 mt-0.5 ${
                      alt.severity === 'high' ? 'text-red-700' : 'text-amber-700'
                    }`}
                  />
                  <div className="text-xs">
                    <span className="font-bold text-gray-900 mr-2">
                      {alt.alertType.toUpperCase().replace(/_/g, ' ')}:
                    </span>
                    <span className="text-gray-800">{alt.message}</span>
                    <span className="ml-2 font-mono font-semibold text-gray-900">
                      ₹{alt.amount.toLocaleString('en-IN')}
                    </span>
                  </div>
                </div>
              ))}
            </div>
          )}

          {/* Aging Table */}
          <div className="bg-white rounded-xl border border-gray-200 overflow-hidden shadow-2xs">
            <div className="px-5 py-3.5 border-b border-gray-100 flex items-center justify-between">
              <div>
                <h3 className="text-sm font-bold text-gray-900">
                  {aparType === 'payable' ? 'Supplier Liabilities Aging' : 'Customer Receivables Aging'}
                </h3>
                <p className="text-xs text-gray-500">Calculated from verified invoice dates, terms, and payments</p>
              </div>
              <span className="text-xs font-semibold text-gray-700 bg-gray-100 px-2.5 py-1 rounded">
                Total Open Records: {agingRecords.length}
              </span>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-xs text-left">
                <thead className="bg-gray-50 text-gray-600 font-semibold border-b border-gray-200">
                  <tr>
                    <th className="px-4 py-3">Party Name</th>
                    <th className="px-4 py-3">Invoice #</th>
                    <th className="px-4 py-3">Invoice Date</th>
                    <th className="px-4 py-3">Due Date</th>
                    <th className="px-4 py-3 text-right">Invoice Amount</th>
                    <th className="px-4 py-3 text-right">Balance Due</th>
                    <th className="px-4 py-3 text-center">Overdue Days</th>
                    <th className="px-4 py-3 text-center">Aging Bucket</th>
                    <th className="px-4 py-3 text-center">MSME Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {agingRecords.length > 0 ? (
                    agingRecords.map((r) => (
                      <tr key={r.id} className="hover:bg-gray-50/60">
                        <td className="px-4 py-2.5 font-semibold text-gray-900">{r.partyName}</td>
                        <td className="px-4 py-2.5 font-mono text-gray-600">{r.invoiceNo}</td>
                        <td className="px-4 py-2.5 text-gray-600">{r.invoiceDate}</td>
                        <td className="px-4 py-2.5 text-gray-600">{r.dueDate}</td>
                        <td className="px-4 py-2.5 text-right font-mono">₹{r.amount.toLocaleString('en-IN')}</td>
                        <td className="px-4 py-2.5 text-right font-mono font-bold text-gray-900">
                          ₹{r.balance.toLocaleString('en-IN')}
                        </td>
                        <td className="px-4 py-2.5 text-center">
                          <span
                            className={`px-2 py-0.5 rounded font-bold ${
                              r.overdueDays > 45 ? 'bg-red-100 text-red-800' : 'bg-gray-100 text-gray-700'
                            }`}
                          >
                            {r.overdueDays}d
                          </span>
                        </td>
                        <td className="px-4 py-2.5 text-center">
                          <span className="bg-gray-100 text-gray-900 border border-gray-200 px-2 py-0.5 rounded font-medium">
                            {r.agingBucket}
                          </span>
                        </td>
                        <td className="px-4 py-2.5 text-center">
                          {r.isMsmeSupplier ? (
                            <span className="bg-blue-50 text-blue-800 border border-blue-200 px-2 py-0.5 rounded font-semibold text-[10px]">
                              MSME Verified
                            </span>
                          ) : (
                            <span className="text-gray-400">—</span>
                          )}
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={9} className="px-4 py-8 text-center text-gray-500">
                        No aging records loaded. Click "Analyze Ledger Records" to evaluate open invoices.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* MODULE 3: AUDIT EVIDENCE AUTOMATION */}
      {/* ========================================================================= */}
      {activeModule === 'evidence' && (
        <div className="space-y-4">
          <div className="bg-white rounded-xl border border-gray-200 p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div>
              <h3 className="text-sm font-bold text-gray-900">Audit Traceability & Evidence Register</h3>
              <p className="text-xs text-gray-500">Cross-reference ledger vouchers with substantive audit working papers</p>
            </div>

            <div className="flex items-center gap-2">
              <div className="flex items-center gap-1.5 border border-gray-300 rounded-lg px-2.5 py-1 bg-white">
                <span className="text-[11px] text-gray-500">Materiality:</span>
                <input
                  type="number"
                  value={missingEvidenceThreshold}
                  onChange={(e) => setMissingEvidenceThreshold(parseFloat(e.target.value) || 0)}
                  className="w-20 text-xs font-semibold focus:outline-none"
                />
              </div>

              <button
                onClick={handleScanMissingEvidence}
                className="px-3 py-1.5 text-xs font-semibold text-gray-900 bg-gray-100 hover:bg-gray-100 border border-gray-200 rounded-lg transition-colors cursor-pointer"
              >
                Scan Missing Evidence
              </button>

              <button
                onClick={handleExportEvidence}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-gray-700 bg-white hover:bg-gray-50 border border-gray-300 rounded-lg transition-colors cursor-pointer"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Export Working Papers</span>
              </button>
            </div>
          </div>

          {/* Missing Evidence Warning Alert */}
          {missingEvidenceResult && missingEvidenceResult.missingEvidenceCount > 0 && (
            <div className="p-4 bg-amber-50 border border-amber-200 rounded-xl space-y-2">
              <div className="flex items-center gap-2 text-amber-900 font-bold text-xs">
                <AlertTriangle className="w-4 h-4 text-amber-700" />
                <span>
                  MATERIAL TRANSACTIONS LACKING SUPPORTING EVIDENCE ({missingEvidenceResult.missingEvidenceCount} Items)
                </span>
              </div>
              <p className="text-xs text-amber-800">
                Found vouchers exceeding materiality threshold (₹{missingEvidenceThreshold.toLocaleString('en-IN')}) with status 'Missing' or 'Requested'. Audit verification cannot be marked complete until documentary evidence is reviewed.
              </p>
            </div>
          )}

          {/* Evidence Slots Cards */}
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
            {evidenceList.map((slot) => (
              <div key={slot.id} className="bg-white rounded-xl border border-gray-200 p-4 space-y-3 shadow-2xs">
                <div className="flex items-start justify-between">
                  <div>
                    <span className="font-mono text-[10px] font-bold bg-gray-100 text-gray-700 px-2 py-0.5 rounded">
                      {slot.workingPaperIndex}
                    </span>
                    <h4 className="text-xs font-bold text-gray-900 mt-1">{slot.documentTitle}</h4>
                    <p className="text-[11px] text-gray-500">{slot.auditArea}</p>
                  </div>
                  <span
                    className={`text-[10px] font-bold px-2 py-0.5 rounded-full ${
                      slot.status === 'Reviewed'
                        ? 'bg-gray-100 text-gray-900'
                        : slot.status === 'Linked'
                        ? 'bg-blue-100 text-blue-800'
                        : slot.status === 'Received'
                        ? 'bg-slate-100 text-slate-800 border border-slate-200'
                        : 'bg-amber-100 text-amber-800'
                    }`}
                  >
                    {slot.status}
                  </span>
                </div>

                <div className="p-2.5 bg-gray-50 rounded-lg text-[11px] space-y-1">
                  <div className="flex justify-between text-gray-600">
                    <span>Account Head:</span>
                    <span className="font-medium text-gray-800">{slot.accountHead}</span>
                  </div>
                  <div className="flex justify-between text-gray-600">
                    <span>Linked Vouchers:</span>
                    <span className="font-mono font-bold text-gray-900">{slot.linkedTransactions?.length || 0}</span>
                  </div>
                </div>

                <div className="pt-2 border-t border-gray-100 flex items-center justify-between">
                  <button
                    onClick={() => {
                      setTargetEvidenceId(slot.id);
                      setIsLinkModalOpen(true);
                    }}
                    className="text-xs font-semibold text-gray-900 hover:text-gray-900 flex items-center gap-1 cursor-pointer"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    <span>Link Voucher</span>
                  </button>

                  {slot.reviewerNotes && (
                    <span className="text-[10px] text-gray-500 italic max-w-[140px] truncate" title={slot.reviewerNotes}>
                      {slot.reviewerNotes}
                    </span>
                  )}
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* MODULE 4: ACCOUNTING ANOMALY DETECTION ENGINE */}
      {/* ========================================================================= */}
      {activeModule === 'anomalies' && (
        <div className="space-y-4">
          <div className="bg-white rounded-xl border border-gray-200 p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div>
              <h3 className="text-sm font-bold text-gray-900">Deterministic Anomaly Detection Engine</h3>
              <p className="text-xs text-gray-500">Explainable rule-based scrutiny for journal duplicates, cash limits, and velocity spikes</p>
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={handleRunAnomalyScan}
                className="flex items-center gap-1.5 px-3.5 py-1.5 text-xs font-bold text-white bg-gray-900 hover:bg-black rounded-lg transition-colors cursor-pointer"
              >
                <AlertTriangle className="w-3.5 h-3.5" />
                <span>Run Anomaly Scan</span>
              </button>

              <button
                onClick={handleExportAnomalies}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-gray-700 bg-white hover:bg-gray-50 border border-gray-300 rounded-lg transition-colors cursor-pointer"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Export Scrutiny Report</span>
              </button>
            </div>
          </div>

          {/* Active Rules Grid */}
          <div className="bg-white rounded-xl border border-gray-200 p-4 space-y-3">
            <h4 className="text-xs font-bold text-gray-900 uppercase tracking-wider">Configured Scrutiny Rules</h4>
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
              {anomalyRules.map((rule) => (
                <div key={rule.id} className="p-3 rounded-lg border border-gray-200 bg-gray-50/60 space-y-1.5">
                  <div className="flex items-center justify-between">
                    <span className="font-mono text-[10px] font-bold text-gray-900 bg-gray-100 px-1.5 py-0.5 rounded">
                      {rule.ruleCode}
                    </span>
                    <span
                      className={`text-[10px] font-bold px-1.5 rounded ${
                        rule.severity === 'high' ? 'bg-red-100 text-red-800' : 'bg-amber-100 text-amber-800'
                      }`}
                    >
                      {rule.severity.toUpperCase()}
                    </span>
                  </div>
                  <h5 className="text-xs font-bold text-gray-900">{rule.name}</h5>
                  <p className="text-[11px] text-gray-600">{rule.description}</p>
                </div>
              ))}
            </div>
          </div>

          {/* Detected Anomaly Alerts */}
          <div className="bg-white rounded-xl border border-gray-200 overflow-hidden shadow-2xs">
            <div className="px-5 py-3.5 border-b border-gray-100 flex items-center justify-between">
              <div>
                <h4 className="text-sm font-bold text-gray-900">Detected Scrutiny Alerts</h4>
                <p className="text-xs text-gray-500">Auditor review required before finalizing financial statements</p>
              </div>
              <span className="text-xs font-semibold bg-gray-100 text-gray-700 px-2 py-0.5 rounded">
                {anomalyAlerts.length} Alerts Open
              </span>
            </div>

            <div className="divide-y divide-gray-100">
              {anomalyAlerts.length > 0 ? (
                anomalyAlerts.map((alt) => (
                  <div key={alt.id} className="p-4 hover:bg-gray-50/50 flex flex-col md:flex-row md:items-center justify-between gap-4">
                    <div className="space-y-1">
                      <div className="flex items-center gap-2">
                        <span
                          className={`text-[10px] font-bold px-2 py-0.5 rounded-full ${
                            alt.severity === 'high'
                              ? 'bg-red-100 text-red-800'
                              : alt.severity === 'medium'
                              ? 'bg-amber-100 text-amber-800'
                              : 'bg-blue-100 text-blue-800'
                          }`}
                        >
                          {alt.severity.toUpperCase()}
                        </span>
                        <span className="font-mono text-xs font-bold text-gray-900">{alt.ruleCode}</span>
                        <span className="text-xs text-gray-400">•</span>
                        <span className="font-mono text-xs font-bold text-gray-900">
                          ₹{alt.affectedAmount?.toLocaleString('en-IN')}
                        </span>
                      </div>
                      <p className="text-xs font-medium text-gray-900">{alt.explanation}</p>
                      <p className="text-[11px] text-gray-500">
                        Affected Head: {alt.affectedAccountHead} | Voucher: {alt.sourceRecordIds?.join(', ')}
                      </p>
                      {alt.reviewNotes && (
                        <p className="text-[11px] text-gray-600 italic bg-gray-100 p-1.5 rounded">
                          Reviewer Note: {alt.reviewNotes}
                        </p>
                      )}
                    </div>

                    <div className="flex items-center gap-2 shrink-0">
                      <button
                        onClick={() => setSelectedAlertForReview(alt)}
                        className="px-3 py-1.5 text-xs font-semibold text-gray-900 bg-gray-100 hover:bg-gray-100 border border-gray-200 rounded-lg transition-colors cursor-pointer"
                      >
                        Review Alert
                      </button>
                    </div>
                  </div>
                ))
              ) : (
                <div className="p-8 text-center text-gray-500 text-xs">
                  No anomaly alerts detected. Click "Run Anomaly Scan" to evaluate journal entries against deterministic rules.
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* MODULE 5: CLIENT DOCUMENT COLLECTION & FOLLOW-UP */}
      {/* ========================================================================= */}
      {activeModule === 'document-requests' && (
        <div className="space-y-4">
          <div className="bg-white rounded-xl border border-gray-200 p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div>
              <h3 className="text-sm font-bold text-gray-900">Client Document Collection & Follow-up</h3>
              <p className="text-xs text-gray-500">Track client requisitions, verify file integrity hashes, and draft reminder notices</p>
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={handleDraftReminder}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-gray-900 bg-gray-100 hover:bg-gray-100 border border-gray-200 rounded-lg transition-colors cursor-pointer"
              >
                <Mail className="w-3.5 h-3.5" />
                <span>Draft Reminder Notice</span>
              </button>

              <button
                onClick={handleExportDocuments}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-gray-700 bg-white hover:bg-gray-50 border border-gray-300 rounded-lg transition-colors cursor-pointer"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Export Outstanding List</span>
              </button>
            </div>
          </div>

          {/* Reminder Draft Banner */}
          {reminderDraft && (
            <div className="bg-gray-100 border border-gray-200 rounded-xl p-4 space-y-3">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2 text-gray-900 font-bold text-xs">
                  <Mail className="w-4 h-4 text-gray-900" />
                  <span>DRAFTED REMINDER NOTICE FOR REVIEW</span>
                </div>
                <button
                  onClick={() => {
                    navigator.clipboard.writeText(reminderDraft.emailBody);
                    onShowToast('Copied to Clipboard', 'Reminder body copied', 'success');
                  }}
                  className="flex items-center gap-1 text-xs font-semibold text-gray-900 hover:underline cursor-pointer"
                >
                  <Copy className="w-3 h-3" />
                  <span>Copy Body</span>
                </button>
              </div>

              <div className="bg-white rounded-lg p-3 border border-gray-200 text-xs font-mono whitespace-pre-wrap text-gray-800">
                {reminderDraft.emailBody}
              </div>
            </div>
          )}

          {/* Document Requests Table */}
          <div className="bg-white rounded-xl border border-gray-200 overflow-hidden shadow-2xs">
            <div className="px-5 py-3.5 border-b border-gray-100 flex items-center justify-between">
              <h4 className="text-sm font-bold text-gray-900">Requisition Items</h4>
              <span className="text-xs font-semibold text-gray-700 bg-gray-100 px-2 py-0.5 rounded">
                {docRequests.length} Required Documents
              </span>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-xs text-left">
                <thead className="bg-gray-50 text-gray-600 font-semibold border-b border-gray-200">
                  <tr>
                    <th className="px-4 py-3">Document Title</th>
                    <th className="px-4 py-3">Category</th>
                    <th className="px-4 py-3">Due Date</th>
                    <th className="px-4 py-3">Mandatory</th>
                    <th className="px-4 py-3">Status</th>
                    <th className="px-4 py-3">Verified File</th>
                    <th className="px-4 py-3 text-right">Action</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {docRequests.map((req) => (
                    <tr key={req.id} className="hover:bg-gray-50/60">
                      <td className="px-4 py-3 font-semibold text-gray-900">{req.title}</td>
                      <td className="px-4 py-3 text-gray-600">{req.category}</td>
                      <td className="px-4 py-3 text-gray-600">{req.dueDate}</td>
                      <td className="px-4 py-3">
                        {req.mandatory ? (
                          <span className="text-red-700 font-bold text-[10px] bg-red-50 border border-red-200 px-1.5 py-0.5 rounded">
                            Required
                          </span>
                        ) : (
                          <span className="text-gray-400">Optional</span>
                        )}
                      </td>
                      <td className="px-4 py-3">
                        <span
                          className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                            req.status === 'accepted'
                              ? 'bg-gray-100 text-gray-900'
                              : req.status === 'received'
                              ? 'bg-blue-100 text-blue-800'
                              : 'bg-amber-100 text-amber-800'
                          }`}
                        >
                          {req.status.toUpperCase()}
                        </span>
                      </td>
                      <td className="px-4 py-3 font-mono text-[11px] text-gray-600">
                        {req.receivedFileName ? (
                          <div title={`Hash: ${req.receivedFileHash}`}>
                            <span className="font-semibold text-gray-900">{req.receivedFileName}</span>
                          </div>
                        ) : (
                          <span className="text-gray-400">—</span>
                        )}
                      </td>
                      <td className="px-4 py-3 text-right">
                        {req.status === 'pending' && (
                          <button
                            onClick={() => {
                              setUploadTargetRequestId(req.id);
                              setSimulatedFileName(`${req.title.toLowerCase().replace(/\s+/g, '_')}.pdf`);
                            }}
                            className="text-xs font-semibold text-gray-900 hover:text-gray-900 underline cursor-pointer"
                          >
                            Register Upload
                          </button>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* MODULE 6: RECURRING COMPLIANCE & TASK AUTOMATION */}
      {/* ========================================================================= */}
      {activeModule === 'compliance' && (
        <div className="space-y-4">
          <div className="bg-white rounded-xl border border-gray-200 p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div>
              <h3 className="text-sm font-bold text-gray-900">Recurring Compliance & Statutory Calendar</h3>
              <p className="text-xs text-gray-500">Track GSTR-1, GSTR-3B, TDS ITNS-281, and Advance Tax with verified filing proofs</p>
            </div>

            <button
              onClick={handleExportCompliance}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-gray-700 bg-white hover:bg-gray-50 border border-gray-300 rounded-lg transition-colors cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Export Compliance Calendar</span>
            </button>
          </div>

          {/* Compliance Tasks Table */}
          <div className="bg-white rounded-xl border border-gray-200 overflow-hidden shadow-2xs">
            <div className="px-5 py-3.5 border-b border-gray-100 flex items-center justify-between">
              <h4 className="text-sm font-bold text-gray-900">Statutory Filing Schedule</h4>
              <span className="text-xs font-semibold text-gray-700 bg-gray-100 px-2 py-0.5 rounded">
                Active Calendar: September 2026
              </span>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-xs text-left">
                <thead className="bg-gray-50 text-gray-600 font-semibold border-b border-gray-200">
                  <tr>
                    <th className="px-4 py-3">Compliance Task</th>
                    <th className="px-4 py-3">Jurisdiction / Section</th>
                    <th className="px-4 py-3">Statutory Due Date</th>
                    <th className="px-4 py-3">Status</th>
                    <th className="px-4 py-3">Filing Proof Reference</th>
                    <th className="px-4 py-3 text-right">Action</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {complianceInstances.map((t) => (
                    <tr key={t.id} className="hover:bg-gray-50/60">
                      <td className="px-4 py-3 font-semibold text-gray-900">{t.title}</td>
                      <td className="px-4 py-3 text-gray-600">{t.statutoryActRef}</td>
                      <td className="px-4 py-3 font-semibold text-gray-800">{t.dueDate}</td>
                      <td className="px-4 py-3">
                        <span
                          className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                            t.status === 'completed'
                              ? 'bg-gray-100 text-gray-900'
                              : t.status === 'overdue'
                              ? 'bg-red-100 text-red-800'
                              : t.status === 'due_today'
                              ? 'bg-amber-100 text-amber-800'
                              : 'bg-blue-100 text-blue-800'
                          }`}
                        >
                          {t.status.toUpperCase()}
                        </span>
                      </td>
                      <td className="px-4 py-3 font-mono text-gray-600">
                        {t.statutoryFilingProof ? (
                          <span className="text-gray-900 font-semibold">{t.statutoryFilingProof}</span>
                        ) : (
                          <span className="text-gray-400">Pending Filing</span>
                        )}
                      </td>
                      <td className="px-4 py-3 text-right">
                        {t.status !== 'completed' && (
                          <button
                            onClick={() => {
                              setSelectedTaskToComplete(t);
                              setCompletionProofRef('ARN_AA270926012485C');
                            }}
                            className="text-xs font-semibold text-gray-900 hover:text-gray-900 underline cursor-pointer"
                          >
                            Mark Completed
                          </button>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* MODULE 7: MULTI-CLIENT BATCH AUTOMATION */}
      {/* ========================================================================= */}
      {activeModule === 'batch' && (
        <div className="space-y-4">
          <div className="bg-white rounded-xl border border-gray-200 p-5 space-y-4 shadow-2xs">
            <div>
              <h3 className="text-sm font-bold text-gray-900">Multi-Client Audit Automation</h3>
              <p className="text-xs text-gray-500">
                Execute automated workflows across multiple clients concurrently with complete multi-tenant container isolation.
              </p>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 pt-2 border-t border-gray-100">
              <div>
                <label className="text-xs font-bold text-gray-700 block mb-1">Target Automation Workflow</label>
                <select
                  value={selectedBatchWorkflow}
                  onChange={(e) => setSelectedBatchWorkflow(e.target.value)}
                  className="w-full text-xs rounded-lg border border-gray-300 p-2 focus:border-gray-900 focus:outline-none bg-white font-medium"
                >
                  <option value="wf_bank_rec">Bank Reconciliation Workflow (Bank vs Ledger)</option>
                  <option value="wf_gst_itc">GST ITC 2B Rec (Purchase Register vs GSTR-2B)</option>
                  <option value="wf_invoice_ocr">Invoice Processing OCR & Math Validation</option>
                </select>
              </div>

              <div>
                <label className="text-xs font-bold text-gray-700 block mb-1">Client Selection (Isolated Containers)</label>
                <div className="space-y-1.5">
                  {[
                    { id: 'client_01', name: 'Enterprise Client A' },
                    { id: 'client_02', name: 'Enterprise Client B' },
                    { id: 'client_03', name: 'Enterprise Client C' },
                  ].map((cl) => (
                    <label key={cl.id} className="flex items-center gap-2 text-xs text-gray-800">
                      <input
                        type="checkbox"
                        checked={selectedBatchClients.includes(cl.id)}
                        onChange={(e) => {
                          if (e.target.checked) {
                            setSelectedBatchClients([...selectedBatchClients, cl.id]);
                          } else {
                            setSelectedBatchClients(selectedBatchClients.filter((x) => x !== cl.id));
                          }
                        }}
                        className="rounded text-gray-900 focus:ring-gray-900"
                      />
                      <span>{cl.name}</span>
                    </label>
                  ))}
                </div>
              </div>
            </div>

            <div className="pt-2 flex justify-end">
              <button
                onClick={handleStartBatchRun}
                disabled={selectedBatchClients.length === 0}
                className="flex items-center gap-2 px-4 py-2 text-xs font-bold text-white bg-gray-900 hover:bg-black rounded-lg shadow-xs transition-colors cursor-pointer disabled:opacity-50"
              >
                <Layers className="w-4 h-4" />
                <span>Process Client Accounts</span>
              </button>
            </div>
          </div>

          {/* Past Client Processing Runs */}
          <div className="bg-white rounded-xl border border-gray-200 overflow-hidden shadow-2xs">
            <div className="px-5 py-3.5 border-b border-gray-100 flex items-center justify-between">
              <h4 className="text-sm font-bold text-gray-900">Client Automation Monitor</h4>
              <span className="text-xs font-semibold text-gray-700 bg-gray-100 px-2 py-0.5 rounded">
                {batchRuns.length} Processing Records
              </span>
            </div>

            <div className="divide-y divide-gray-100">
              {batchRuns.length > 0 ? (
                batchRuns.map((b) => (
                  <div key={b.id} className="p-4 space-y-3">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                      <div>
                        <div className="flex items-center gap-2">
                          <span className="font-mono text-xs font-bold text-gray-900">{b.id}</span>
                          <span className="text-xs font-semibold text-gray-900 bg-gray-100 px-2 py-0.5 rounded">
                            {b.workflowId}
                          </span>
                          <span
                            className={`text-[10px] font-bold px-2 py-0.5 rounded-full ${
                              b.status === 'completed'
                                ? 'bg-gray-100 text-gray-900'
                                : b.status === 'failed'
                                ? 'bg-red-100 text-red-800'
                                : 'bg-blue-100 text-blue-800'
                            }`}
                          >
                            {b.status.toUpperCase()}
                          </span>
                        </div>
                        <p className="text-[11px] text-gray-500 mt-0.5">Started At: {b.createdAt}</p>
                      </div>

                      <div className="flex items-center gap-3">
                        <div className="text-right text-xs">
                          <span className="text-gray-500">Completed: </span>
                          <span className="font-bold text-gray-900">{b.completedClients}</span>
                          <span className="text-gray-400"> / </span>
                          <span className="font-bold text-gray-800">{b.totalClients}</span>
                        </div>

                        <button
                          onClick={() => handleExportBatchSummary(b.id)}
                          className="flex items-center gap-1 text-xs font-semibold text-gray-700 bg-gray-50 hover:bg-gray-100 border border-gray-300 px-2.5 py-1.5 rounded cursor-pointer"
                        >
                          <Download className="w-3.5 h-3.5" />
                          <span>Summary</span>
                        </button>
                      </div>
                    </div>

                    {/* Progress Bar */}
                    <div className="w-full bg-gray-100 rounded-full h-2 overflow-hidden">
                      <div
                        className="bg-gray-900 h-2 rounded-full transition-all duration-500"
                        style={{
                          width: `${b.totalClients > 0 ? (b.completedClients / b.totalClients) * 100 : 0}%`,
                        }}
                      />
                    </div>

                    {/* Isolated Client Items */}
                    <div className="grid grid-cols-1 md:grid-cols-3 gap-2 pt-1">
                      {b.clientItems.map((ci: any) => (
                        <div key={ci.clientId} className="p-2 bg-gray-50 rounded border border-gray-200 text-xs flex justify-between">
                          <span className="font-semibold text-gray-800 truncate">{ci.clientName}</span>
                          <span
                            className={`text-[10px] font-bold ${
                              ci.status === 'completed' ? 'text-gray-900' : 'text-amber-800'
                            }`}
                          >
                            {ci.status.toUpperCase()}
                          </span>
                        </div>
                      ))}
                    </div>
                  </div>
                ))
              ) : (
                <div className="p-8 text-center text-gray-500 text-xs">
                  No client reconciliation runs recorded yet. Select clients above and click "Process Client Accounts".
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* MODALS */}
      {/* ========================================================================= */}

      {/* Modal 1: New Closing Period */}
      {isNewPeriodModalOpen && (
        <div className="fixed inset-0 bg-black/40 backdrop-blur-2xs flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-xl border border-gray-200 shadow-xl max-w-md w-full p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-gray-100 pb-3">
              <h3 className="text-sm font-bold text-gray-900">Create New Closing Period</h3>
              <button
                onClick={() => setIsNewPeriodModalOpen(false)}
                className="text-gray-400 hover:text-gray-600 cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="space-y-3">
              <div>
                <label className="text-xs font-semibold text-gray-700 block mb-1">Period Name</label>
                <input
                  type="text"
                  value={newPeriodName}
                  onChange={(e) => setNewPeriodName(e.target.value)}
                  className="w-full text-xs rounded-lg border border-gray-300 p-2.5 focus:border-gray-900 focus:outline-none"
                  placeholder="e.g. September 2026"
                />
              </div>

              <div>
                <label className="text-xs font-semibold text-gray-700 block mb-1">Financial Year</label>
                <input
                  type="text"
                  value={newPeriodFy}
                  onChange={(e) => setNewPeriodFy(e.target.value)}
                  className="w-full text-xs rounded-lg border border-gray-300 p-2.5 focus:border-gray-900 focus:outline-none"
                  placeholder="FY 2026-27"
                />
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-gray-100">
              <button
                onClick={() => setIsNewPeriodModalOpen(false)}
                className="px-3 py-1.5 text-xs font-semibold text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-lg cursor-pointer"
              >
                Cancel
              </button>
              <button
                onClick={handleCreatePeriod}
                className="px-4 py-1.5 text-xs font-bold text-white bg-gray-900 hover:bg-black rounded-lg cursor-pointer"
              >
                Initialize Checklist
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal 2: Link Voucher to Evidence */}
      {isLinkModalOpen && (
        <div className="fixed inset-0 bg-black/40 backdrop-blur-2xs flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-xl border border-gray-200 shadow-xl max-w-md w-full p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-gray-100 pb-3">
              <h3 className="text-sm font-bold text-gray-900">Link Voucher to Working Paper</h3>
              <button onClick={() => setIsLinkModalOpen(false)} className="text-gray-400 hover:text-gray-600 cursor-pointer">
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="space-y-3">
              <div>
                <label className="text-xs font-semibold text-gray-700 block mb-1">Voucher #</label>
                <input
                  type="text"
                  value={linkVoucherNo}
                  onChange={(e) => setLinkVoucherNo(e.target.value)}
                  className="w-full text-xs rounded-lg border border-gray-300 p-2 focus:border-gray-900 focus:outline-none"
                  placeholder="e.g. VR-2026-081"
                />
              </div>

              <div>
                <label className="text-xs font-semibold text-gray-700 block mb-1">Transaction Amount (₹)</label>
                <input
                  type="number"
                  value={linkVoucherAmount}
                  onChange={(e) => setLinkVoucherAmount(e.target.value)}
                  className="w-full text-xs rounded-lg border border-gray-300 p-2 focus:border-gray-900 focus:outline-none"
                  placeholder="e.g. 450000"
                />
              </div>

              <div>
                <label className="text-xs font-semibold text-gray-700 block mb-1">Party / Vendor Name</label>
                <input
                  type="text"
                  value={linkPartyName}
                  onChange={(e) => setLinkPartyName(e.target.value)}
                  className="w-full text-xs rounded-lg border border-gray-300 p-2 focus:border-gray-900 focus:outline-none"
                  placeholder="e.g. Tata Steel Tubes Ltd"
                />
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-gray-100">
              <button
                onClick={() => setIsLinkModalOpen(false)}
                className="px-3 py-1.5 text-xs font-semibold text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-lg cursor-pointer"
              >
                Cancel
              </button>
              <button
                onClick={handleLinkTransaction}
                className="px-4 py-1.5 text-xs font-bold text-white bg-gray-900 hover:bg-black rounded-lg cursor-pointer"
              >
                Save Voucher Link
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal 3: Review Anomaly Alert */}
      {selectedAlertForReview && (
        <div className="fixed inset-0 bg-black/40 backdrop-blur-2xs flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-xl border border-gray-200 shadow-xl max-w-lg w-full p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-gray-100 pb-3">
              <h3 className="text-sm font-bold text-gray-900">Auditor Scrutiny Review</h3>
              <button onClick={() => setSelectedAlertForReview(null)} className="text-gray-400 hover:text-gray-600 cursor-pointer">
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="p-3 bg-gray-50 rounded-lg border border-gray-200 space-y-1 text-xs">
              <p className="font-bold text-gray-900">{selectedAlertForReview.ruleCode}: {selectedAlertForReview.explanation}</p>
              <p className="text-gray-600">Affected Head: {selectedAlertForReview.affectedAccountHead}</p>
              <p className="text-gray-600 font-mono">Amount: ₹{selectedAlertForReview.affectedAmount?.toLocaleString('en-IN')}</p>
            </div>

            <div>
              <label className="text-xs font-semibold text-gray-700 block mb-1">
                Auditor Disposition Notes & Justification
              </label>
              <textarea
                rows={3}
                value={alertReviewNotes}
                onChange={(e) => setAlertReviewNotes(e.target.value)}
                placeholder="Explain justification for false-positive dismissal or required journal adjustment..."
                className="w-full text-xs rounded-lg border border-gray-300 p-2 focus:border-gray-900 focus:outline-none"
              />
            </div>

            <div className="flex justify-between items-center pt-2 border-t border-gray-100">
              <button
                onClick={() => handleReviewAlert('dismissed_false_positive')}
                className="px-3 py-1.5 text-xs font-semibold text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-lg cursor-pointer"
              >
                Dismiss False Positive
              </button>

              <div className="flex items-center gap-2">
                <button
                  onClick={() => handleReviewAlert('investigating')}
                  className="px-3 py-1.5 text-xs font-semibold text-blue-800 bg-blue-50 hover:bg-blue-100 rounded-lg cursor-pointer"
                >
                  Mark Investigating
                </button>
                <button
                  onClick={() => handleReviewAlert('marked_for_adjustment')}
                  className="px-3 py-1.5 text-xs font-bold text-white bg-gray-900 hover:bg-black rounded-lg cursor-pointer"
                >
                  Flag for Adjustment
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Modal 4: Register Uploaded Document */}
      {uploadTargetRequestId && (
        <div className="fixed inset-0 bg-black/40 backdrop-blur-2xs flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-xl border border-gray-200 shadow-xl max-w-md w-full p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-gray-100 pb-3">
              <h3 className="text-sm font-bold text-gray-900">Register Client Received File</h3>
              <button onClick={() => setUploadTargetRequestId(null)} className="text-gray-400 hover:text-gray-600 cursor-pointer">
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="space-y-3">
              <div>
                <label className="text-xs font-semibold text-gray-700 block mb-1">File Name</label>
                <input
                  type="text"
                  value={simulatedFileName}
                  onChange={(e) => setSimulatedFileName(e.target.value)}
                  className="w-full text-xs rounded-lg border border-gray-300 p-2 focus:border-gray-900 focus:outline-none font-mono"
                  placeholder="bank_statement_sep2026.pdf"
                />
              </div>

              <div className="p-3 bg-gray-50 rounded-lg border border-gray-200 text-xs space-y-1">
                <p className="font-semibold text-gray-700">Integrity Verification:</p>
                <p className="text-gray-500 text-[11px]">
                  Files are validated for allowable mime types (PDF/XLSX/CSV) and stamped with SHA-256 hash for audit proof.
                </p>
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-gray-100">
              <button
                onClick={() => setUploadTargetRequestId(null)}
                className="px-3 py-1.5 text-xs font-semibold text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-lg cursor-pointer"
              >
                Cancel
              </button>
              <button
                onClick={handleUploadSimulatedDoc}
                className="px-4 py-1.5 text-xs font-bold text-white bg-gray-900 hover:bg-black rounded-lg cursor-pointer"
              >
                Confirm File Receipt
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal 5: Complete Compliance Task */}
      {selectedTaskToComplete && (
        <div className="fixed inset-0 bg-black/40 backdrop-blur-2xs flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-xl border border-gray-200 shadow-xl max-w-md w-full p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-gray-100 pb-3">
              <h3 className="text-sm font-bold text-gray-900">Record Statutory Filing Proof</h3>
              <button onClick={() => setSelectedTaskToComplete(null)} className="text-gray-400 hover:text-gray-600 cursor-pointer">
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="space-y-3">
              <div className="p-3 bg-gray-50 rounded-lg border border-gray-200 text-xs">
                <p className="font-bold text-gray-900">{selectedTaskToComplete.title}</p>
                <p className="text-gray-600">Act: {selectedTaskToComplete.statutoryActRef}</p>
                <p className="text-gray-600">Due Date: {selectedTaskToComplete.dueDate}</p>
              </div>

              <div>
                <label className="text-xs font-semibold text-gray-700 block mb-1">
                  Challan BSR / ARN / Acknowledgement Reference #
                </label>
                <input
                  type="text"
                  value={completionProofRef}
                  onChange={(e) => setCompletionProofRef(e.target.value)}
                  className="w-full text-xs rounded-lg border border-gray-300 p-2 focus:border-gray-900 focus:outline-none font-mono"
                  placeholder="e.g. ARN_AA270926012485C or BSR-0210045"
                />
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-gray-100">
              <button
                onClick={() => setSelectedTaskToComplete(null)}
                className="px-3 py-1.5 text-xs font-semibold text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-lg cursor-pointer"
              >
                Cancel
              </button>
              <button
                onClick={handleCompleteCompliance}
                className="px-4 py-1.5 text-xs font-bold text-white bg-gray-900 hover:bg-black rounded-lg cursor-pointer"
              >
                Confirm Compliance Completion
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
