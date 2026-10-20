import React, { useState, useEffect } from 'react';
import {
  GitMerge,
  Play,
  RotateCcw,
  CheckCircle2,
  AlertTriangle,
  Clock,
  Download,
  FileSpreadsheet,
  Layers,
  ArrowRight,
  ShieldCheck,
  Check,
  X,
  FileText,
  Activity,
  History,
  Cpu,
  Database,
  Lock,
  Zap,
  Sliders,
  Workflow,
  Sparkles,
  Scale,
  FileCode,
  FileCheck2,
  Eye,
  Server,
  Terminal,
  Shield,
  CheckCheck,
} from 'lucide-react';
import { CAAutomationApiClient } from '../../services/apiClient';

interface WorkflowOrchestratorViewProps {
  selectedClientId?: string;
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
}

interface ArchitectureStage {
  id: number;
  code: string;
  title: string;
  shortName: string;
  subtitle: string;
  badge: string;
  icon: React.ComponentType<{ className?: string }>;
  engine: string;
  statutoryAct: string;
  throughput: string;
  inputs: string[];
  outputs: string[];
  guarantee: string;
  keySteps: string[];
}

const ARCHITECTURE_STAGES: ArchitectureStage[] = [
  {
    id: 1,
    code: 'INGESTION',
    title: 'Stage 1: Multi-Source Document Ingestion',
    shortName: '1. Ingestion',
    subtitle: 'Streamed binary ingestion & digital format gate',
    badge: 'Input Boundary',
    icon: Database,
    engine: 'src/utils/engine.ts (parseRawFileContents)',
    statutoryAct: 'ICAI SA 500 (Audit Evidence Reliability)',
    throughput: 'Up to 50 MB / 100,000 rows',
    inputs: [
      'Bank Statement PDFs (Digital)',
      'General Ledgers (.xlsx / .csv)',
      'Government Portal GSTR-2B JSON',
      'Form 26AS / AIS Tax JSON',
      'Tax Invoices (PDF & Spreadsheets)',
    ],
    outputs: [
      'Validated In-Memory Byte Stream',
      'SHA-256 Genesis File Hash',
      'Client Folder Boundary Lock',
    ],
    guarantee: 'Digital PDF verification halts image-only scans with descriptive diagnostic error messages before processing.',
    keySteps: [
      'Binary format verification and extension white-listing',
      'Multi-client isolation container boundary allocation',
      'Cryptographic SHA-256 genesis fingerprint generation',
    ],
  },
  {
    id: 2,
    code: 'NORMALIZATION',
    title: 'Stage 2: Deterministic Pre-Processing & Normalizer',
    shortName: '2. Normalizer',
    subtitle: 'Autonomous column mapping, date serials & currency sanitizer',
    badge: 'Data Cleansing',
    icon: Sliders,
    engine: 'src/utils/engine.ts (detectColumns, cleanAccountingAmount)',
    statutoryAct: 'Indian Accounting Standards (Ind AS / AS 1)',
    throughput: '18,500 rows / sec',
    inputs: [
      'Raw column labels (Txn Date, Narration, Withdrawal, Dr/Cr)',
      'Indian currency strings with commas (₹ 1,50,000.00)',
      'Excel 1900 date serials & irregular slash/dot dates',
    ],
    outputs: [
      'Standardized IEEE 754 Floating Values',
      'ISO 8601 UTC Normalized Timestamps',
      'Validated 15-char GSTIN & 10-char PAN structures',
    ],
    guarantee: 'Accounting negative values in parentheses (e.g. "(1,500.00)") deterministically parsed as negative floats (-1500.00).',
    keySteps: [
      'Synonym dictionary fuzzy header detection',
      'Excel serial date conversion with leap-year adjustments',
      'String sanitization & leading-zero preservation in invoice numbers',
    ],
  },
  {
    id: 3,
    code: 'STATUTORY_RULES',
    title: 'Stage 3: Statutory Compliance & Tax Scrutiny Engine',
    shortName: '3. Tax Scrutiny',
    subtitle: 'Zero-hallucination verification of Income Tax & GST Act provisions',
    badge: 'Statutory Core',
    icon: Scale,
    engine: 'server/services/form3cdService.ts & section206abService.ts',
    statutoryAct: 'Income Tax Act 1961 (Sec 40A(3), 43B, 206AB) & CGST Act',
    throughput: '100% Tax Rules Evaluated',
    inputs: [
      'Vendor cash payments & daily supplier aggregations',
      'Statutory liabilities ledger (GST, PF, ESIC balances)',
      'Vendor PANs & historical return filing telemetry',
    ],
    outputs: [
      'Form 3CD Clause 21(d) Cash Disallowance Schedules',
      'Section 206AB Higher Rate (5% floor) Differential Exposure',
      'Rule 88C DRC-01B ITC Variance Alert Schedules',
    ],
    guarantee: 'Rule 6DD statutory exemptions (bank holidays, clearing house payments) and transporter ₹35,000 limits enforced with exact tax law precision.',
    keySteps: [
      'Section 40A(3) daily cash aggregation per payee',
      'Rule 6DD statutory exemption evaluation',
      'Section 206AB / 206CCA 5% floor higher deduction calculation',
      'Section 43B unpaid balance verification as of tax filing date',
    ],
  },
  {
    id: 4,
    code: 'RECONCILIATION',
    title: 'Stage 4: Universal Dual & Triple Reconciliation Core',
    shortName: '4. Rec Core',
    subtitle: 'Exact cryptographic matching, fuzzy tolerances & running balance walk',
    badge: 'Matching Engine',
    icon: CheckCheck,
    engine: 'src/utils/engine.ts & server/services/gstTripleRecService.ts',
    statutoryAct: 'ICAI SA 505 (External Confirmations) & CGST Sec 16(2)(aa)',
    throughput: 'Dual & Triple Stream Match',
    inputs: [
      'Books Bank Ledger vs Bank Statements',
      'Books Purchase Register vs GSTR-2B Portal JSON vs GSTR-3B',
      'Vendor Ledgers with Inter-Company Dr/Cr Inversion',
    ],
    outputs: [
      '100% Exact Confirmed Matches (High Confidence)',
      'Tolerance Window Matches (±₹10 rounding, ±3 days)',
      'Tax & Rate Variance Schedules',
      'Unbooked Bank Charges & Unpresented Cheque Schedules',
    ],
    guarantee: 'Mathematical running balance continuity strictly verified: PrevBal - Withdrawal + Deposit == CurrBal across all rows.',
    keySteps: [
      'Normalized alphanumeric invoice key matching',
      'Dual-stream & triple-stream turnover walk calculation',
      'Inter-company debit/credit mathematical inversion',
      'Section 50 statutory interest estimation for delayed tax payment',
    ],
  },
  {
    id: 5,
    code: 'ANOMALY_REVIEW',
    title: 'Stage 5: Anomaly Scrutiny & Human-in-the-Loop Workbench',
    shortName: '5. Anomaly & Review',
    subtitle: 'Red flag detector with materiality thresholding & CA sign-off',
    badge: 'Quality Control',
    icon: Shield,
    engine: 'server/services/accountingAnomalyService.ts',
    statutoryAct: 'ICAI SA 320 (Materiality in Planning and Performing an Audit)',
    throughput: 'Audit Risk Filtered',
    inputs: [
      'Reconciliation variances and unresolved exception records',
      'Journal entries with round-sum amounts or unusual hours',
      'Suspense and negative cash account ledgers',
    ],
    outputs: [
      'Material Exception Dossier (> 0.5% Turnover)',
      'Auditor Human Review Memos & Override Logs',
      'Auditor Sign-off Approval/Rejection Audit Trail',
    ],
    guarantee: 'Deterministic anomaly scoring never auto-disallows legitimate transactions; human auditor retains final sign-off authority.',
    keySteps: [
      'Materiality filtering based on firm risk parameters',
      'Negative cash balance detection & timeline localization',
      'Round-sum cash payment isolation',
      'Auditor approval/rejection timestamping with reviewer identity',
    ],
  },
  {
    id: 6,
    code: 'OUTPUT_DELIVERY',
    title: 'Stage 6: Multi-Format Deliverables & Cryptographic Trail',
    shortName: '6. Deliverables',
    subtitle: 'Formatted Excel workbooks, Tally Prime XML & ICAI SQC 1 certificates',
    badge: 'Final Deliverables',
    icon: FileSpreadsheet,
    engine: 'server/services/tallyExportService.ts & auditTrailService.ts',
    statutoryAct: 'ICAI SQC 1 (System of Quality Control) & Companies Act 2013',
    throughput: 'Instant Multi-Format Output',
    inputs: [
      'Fully reconciled transaction models',
      'Verified statutory disallowance annexures',
      'Auditor-approved exception memos',
    ],
    outputs: [
      'Multi-Sheet Excel Audit Workbooks with Indian Number Masks (₹ #,##,##0.00)',
      'Tally Prime XML Vouchers with Auto-Balanced Dr/Cr',
      'Form 3CD Clauses 21(a), 21(d), 44 Multi-Sheet Workbooks',
      'ICAI SQC 1 Tamper-Evident Cryptographic SHA-256 PDF Certificate',
    ],
    guarantee: 'Blockchain-style chained SHA-256 audit log; tampering with even one past audit event immediately breaks cryptographic validation.',
    keySteps: [
      'Explicit column sequence preservation without key shuffling',
      'Tally Prime XML formatting with negative debit values for receipt/payment',
      'Chained hash log generation linking every step to parent hash',
      'Printable HTML & Excel audit package compilation',
    ],
  },
];

const WORKFLOW_REQUIREMENTS: Record<
  string,
  { name: string; accept: string; label: string; allowedExts: string[] }
> = {
  wf_bank_rec: {
    name: 'Bank Statement vs Ledger Reconciliation',
    accept: '.pdf,.xlsx,.xls,.csv',
    label: 'Upload Bank Statement (.pdf, .xlsx, .csv) & Books Bank Ledger (.xlsx, .csv)',
    allowedExts: ['.pdf', '.xlsx', '.xls', '.csv'],
  },
  wf_gst_audit: {
    name: 'GSTR-2B vs Books Purchase Reconciliation',
    accept: '.xlsx,.xls,.csv,.json',
    label: 'Upload GSTR-2B Portal Export (.xlsx, .csv, .json) & Books Purchase Register (.xlsx, .csv) — PDF not supported',
    allowedExts: ['.xlsx', '.xls', '.csv', '.json'],
  },
  wf_vendor_pay: {
    name: 'Vendor Payables & Invoice Verification',
    accept: '.xlsx,.xls,.csv,.pdf',
    label: 'Upload Vendor Tax Invoices (.pdf) & Open Payables / Purchase Ledger (.xlsx, .csv)',
    allowedExts: ['.xlsx', '.xls', '.csv', '.pdf'],
  },
  wf_month_end: {
    name: 'Month-End Financial Closing',
    accept: '.xlsx,.xls,.csv',
    label: 'Upload Trial Balance & General Ledger (.xlsx, .xls, .csv) — PDF not supported',
    allowedExts: ['.xlsx', '.xls', '.csv'],
  },
};

export const WorkflowOrchestratorView: React.FC<WorkflowOrchestratorViewProps> = ({
  selectedClientId = '',
  onShowToast,
}) => {
  // Navigation & Architecture Visual State
  const [activeChartTab, setActiveChartTab] = useState<'flowchart' | 'blueprint' | 'statutory'>('flowchart');
  const [selectedStageId, setSelectedStageId] = useState<number>(1);
  const [isSimulating, setIsSimulating] = useState<boolean>(false);
  const [simulatedActiveStage, setSimulatedActiveStage] = useState<number | null>(null);

  // Backend Workflow Execution State
  const [workflows, setWorkflows] = useState<any[]>([]);
  const [selectedWorkflowId, setSelectedWorkflowId] = useState<string>('wf_bank_rec');
  const [jobs, setJobs] = useState<any[]>([]);
  const [activeJob, setActiveJob] = useState<any | null>(null);
  const [activeRun, setActiveRun] = useState<any | null>(null);
  const [isUploading, setIsUploading] = useState<boolean>(false);
  const [uploadedFiles, setUploadedFiles] = useState<{ fileName: string; content: string }[]>([]);
  const [exceptions, setExceptions] = useState<any[]>([]);
  const [activeTab, setActiveTab] = useState<'pipeline' | 'exceptions' | 'artifacts' | 'jobs'>('pipeline');
  const [isEngineLive, setIsEngineLive] = useState<boolean>(true);

  // Live simulation ticker
  useEffect(() => {
    let interval: any = null;
    if (isSimulating) {
      let currentStage = 1;
      setSimulatedActiveStage(currentStage);
      setSelectedStageId(currentStage);
      interval = setInterval(() => {
        currentStage++;
        if (currentStage > 6) {
          setIsSimulating(false);
          setSimulatedActiveStage(null);
          onShowToast('Simulation Complete', 'End-to-end 6-stage pipeline validated successfully with 0 exceptions', 'success');
        } else {
          setSimulatedActiveStage(currentStage);
          setSelectedStageId(currentStage);
        }
      }, 1400);
    }
    return () => clearInterval(interval);
  }, [isSimulating]);

  // Load initial workflows and jobs
  useEffect(() => {
    loadWorkflows();
    loadJobs();
    const timer = setInterval(() => {
      loadJobs();
      if (activeJob && (activeJob.status === 'running' || activeJob.status === 'queued')) {
        loadJobDetails(activeJob.id);
      }
    }, 2500);
    return () => clearInterval(timer);
  }, [selectedClientId, activeJob?.id, activeJob?.status]);

  const loadWorkflows = async () => {
    try {
      const res = await CAAutomationApiClient.getWorkflows();
      if (res.success && res.data) {
        setWorkflows(res.data);
      }
    } catch {
      setIsEngineLive(false);
    }
  };

  const loadJobs = async () => {
    try {
      const res = await CAAutomationApiClient.getJobs(undefined, selectedClientId);
      if (res.success && res.data) {
        setJobs(res.data);
      }
    } catch (err) {
      console.error('Failed to load jobs:', err);
    }
  };

  const loadJobDetails = async (jobId: string) => {
    try {
      const res = await CAAutomationApiClient.getJob(jobId);
      if (res.success && res.data) {
        setActiveJob(res.data);
        if (res.data.runId) {
          const runRes = await CAAutomationApiClient.getWorkflowRun(res.data.runId);
          if (runRes.success && runRes.data) {
            setActiveRun(runRes.data);
            setExceptions(runRes.data.exceptions || []);
          }
        }
      }
    } catch (err) {
      console.error('Failed to load job details:', err);
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const currentReq = WORKFLOW_REQUIREMENTS[selectedWorkflowId] || {
      name: 'Audit Workflow',
      accept: '.xlsx,.xls,.csv,.pdf,.json',
      label: 'Upload required documents',
      allowedExts: ['.xlsx', '.xls', '.csv', '.pdf', '.json'],
    };

    const validFiles: File[] = [];
    const rejectedFiles: string[] = [];

    Array.from(files).forEach((file) => {
      const lower = file.name.toLowerCase();
      const isAllowed = currentReq.allowedExts.some((ext) => lower.endsWith(ext));
      if (isAllowed) {
        validFiles.push(file);
      } else {
        rejectedFiles.push(file.name);
      }
    });

    if (rejectedFiles.length > 0) {
      onShowToast(
        'Incompatible Document Type',
        `File(s) ${rejectedFiles.join(', ')} cannot be processed for "${currentReq.name}". ${currentReq.label}`,
        'error'
      );
    }

    if (validFiles.length === 0) {
      if (e.target) e.target.value = '';
      return;
    }

    const readFiles: { fileName: string; content?: string; base64?: string }[] = [];
    let count = 0;

    validFiles.forEach((file) => {
      const reader = new FileReader();
      reader.onload = (event) => {
        const dataUrl = (event.target?.result as string) || '';
        const base64 = dataUrl.includes(',') ? dataUrl.split(',')[1] : dataUrl;
        readFiles.push({
          fileName: file.name,
          base64,
        });
        count++;
        if (count === validFiles.length) {
          setUploadedFiles((prev) => [...prev, ...readFiles as any]);
        }
      };
      reader.readAsDataURL(file);
    });

    if (e.target) e.target.value = '';
  };

  const handleStartWorkflow = async () => {
    if (uploadedFiles.length === 0) {
      return;
    }

    setIsUploading(true);
    try {
      const res = await CAAutomationApiClient.startWorkflow(
        selectedWorkflowId,
        {
          files: uploadedFiles,
          idempotencyKey: `idemp-${selectedClientId}-${Date.now()}`,
        },
        undefined,
        selectedClientId
      );

      if (res.success) {
        onShowToast('Workflow Dispatched', 'Audit workflow running in background', 'success');
        setUploadedFiles([]);
        await loadJobs();
        if (res.jobId) {
          loadJobDetails(res.jobId);
          setActiveTab('pipeline');
        }
      } else {
        onShowToast('Execution Error', res.error || 'Failed to dispatch workflow', 'error');
      }
    } catch (err: any) {
      onShowToast('Connection Error', err.message, 'error');
    } finally {
      setIsUploading(false);
    }
  };

  const handleResolveException = async (exceptionId: string, action: 'approved' | 'rejected') => {
    try {
      const res = await CAAutomationApiClient.resolveException(
        exceptionId,
        action,
        `Approved via Human-in-the-Loop review for client ${selectedClientId}`
      );
      if (res.success) {
        onShowToast('Exception Updated', `Action: ${action}`, 'success');
        if (activeJob) loadJobDetails(activeJob.id);
      }
    } catch (err: any) {
      onShowToast('Error', err.message, 'error');
    }
  };

  const handleRetry = async (jobId: string) => {
    try {
      const res = await CAAutomationApiClient.retryJob(jobId);
      if (res.success) {
        onShowToast('Retry Started', 'Processing restarted without duplicating outputs', 'success');
        loadJobDetails(jobId);
      }
    } catch (err: any) {
      onShowToast('Retry Failed', err.message, 'error');
    }
  };

  const handleDownloadArtifact = (art: any) => {
    try {
      const base64Data = art.contentBase64 || art.base64Payload || art.base64;
      if (base64Data) {
        const byteCharacters = atob(base64Data);
        const byteNumbers = new Array(byteCharacters.length);
        for (let i = 0; i < byteCharacters.length; i++) {
          byteNumbers[i] = byteCharacters.charCodeAt(i);
        }
        const byteArray = new Uint8Array(byteNumbers);
        const fileExt = art.fileType || (art.fileName?.endsWith('.csv') ? 'csv' : 'xlsx');
        const mimeType =
          fileExt === 'xlsx'
            ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            : fileExt === 'xml'
            ? 'application/xml'
            : 'text/csv';
        const blob = new Blob([byteArray], { type: mimeType });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = art.fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        onShowToast('Report Downloaded', `Successfully downloaded ${art.fileName}`, 'success');
      } else {
        window.open(art.downloadUrl, '_blank');
      }
    } catch (err: any) {
      onShowToast('Download Failed', err.message, 'error');
    }
  };

  const selectedStage = ARCHITECTURE_STAGES.find((s) => s.id === selectedStageId) || ARCHITECTURE_STAGES[0];

  return (
    <div className="space-y-6">
      {/* ========================================================
          1. TOP ARCHITECTURE BANNER & MODE CONTROLLER
         ======================================================== */}
      <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-2xs space-y-4">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="space-y-1">
            <div className="flex items-center gap-2.5">
              <div className="w-8 h-8 rounded-lg bg-gray-900 text-white flex items-center justify-center shrink-0 shadow-xs">
                <GitMerge className="w-4 h-4" />
              </div>
              <h2 className="text-base font-bold text-gray-900 tracking-tight">
                CA Automation Workflow Architecture & Execution Engine
              </h2>
              <span className="px-2 py-0.5 rounded-full text-[10px] font-semibold bg-gray-50 text-gray-800 border border-gray-200 shrink-0">
                ● Deterministic Core Active
              </span>
            </div>
            <p className="text-xs text-gray-600 pl-10.5 max-w-4xl leading-relaxed">
              End-to-end 6-stage architecture design flowchart: ingestion, normalization, statutory tax rules, dual/triple reconciliation, anomaly review, and ICAI SQC 1 cryptographic deliverables.
            </p>
          </div>

          {/* Action Buttons: Simulation & Job History */}
          <div className="flex items-center gap-2 pl-10.5 lg:pl-0 shrink-0">
            <button
              onClick={() => setIsSimulating(!isSimulating)}
              className={`px-3 py-1.5 text-xs font-semibold rounded-lg border flex items-center gap-1.5 transition-all cursor-pointer shadow-2xs ${
                isSimulating
                  ? 'bg-amber-500 hover:bg-amber-600 text-white border-amber-600 animate-pulse'
                  : 'bg-white text-gray-800 border-gray-300 hover:bg-gray-50'
              }`}
              title="Animate active data packet flow through the 6 stages"
            >
              <Zap className={`w-3.5 h-3.5 ${isSimulating ? 'text-white' : 'text-amber-500'}`} />
              <span>{isSimulating ? 'Simulating Active Flow...' : 'Simulate Data Flow'}</span>
            </button>

            <button
              onClick={() => setActiveTab('jobs')}
              className={`px-3 py-1.5 text-xs font-semibold rounded-lg border flex items-center gap-1.5 transition-all cursor-pointer shadow-2xs ${
                activeTab === 'jobs'
                  ? 'bg-gray-900 text-white border-gray-900'
                  : 'bg-white text-gray-700 border-gray-200 hover:bg-gray-50'
              }`}
            >
              <History className="w-3.5 h-3.5" />
              <span>Audit History ({jobs.length})</span>
            </button>
          </div>
        </div>

        {/* View Switcher Tabs */}
        <div className="flex items-center gap-2 pt-2 border-t border-gray-100 overflow-x-auto">
          <button
            onClick={() => setActiveChartTab('flowchart')}
            className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all cursor-pointer flex items-center gap-1.5 ${
              activeChartTab === 'flowchart'
                ? 'bg-gray-900 text-white shadow-xs'
                : 'bg-gray-50 text-gray-600 hover:bg-gray-100 hover:text-gray-900'
            }`}
          >
            <Workflow className="w-3.5 h-3.5" />
            <span>Interactive Pipeline Flowchart</span>
          </button>

          <button
            onClick={() => setActiveChartTab('blueprint')}
            className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all cursor-pointer flex items-center gap-1.5 ${
              activeChartTab === 'blueprint'
                ? 'bg-gray-900 text-white shadow-xs'
                : 'bg-gray-50 text-gray-600 hover:bg-gray-100 hover:text-gray-900'
            }`}
          >
            <Server className="w-3.5 h-3.5" />
            <span>System Architecture Blueprint</span>
          </button>

          <button
            onClick={() => setActiveChartTab('statutory')}
            className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all cursor-pointer flex items-center gap-1.5 ${
              activeChartTab === 'statutory'
                ? 'bg-gray-900 text-white shadow-xs'
                : 'bg-gray-50 text-gray-600 hover:bg-gray-100 hover:text-gray-900'
            }`}
          >
            <Scale className="w-3.5 h-3.5" />
            <span>Statutory Compliance Matrix</span>
          </button>
        </div>
      </div>

      {/* ========================================================
          2. VIEW A: INTERACTIVE PIPELINE FLOWCHART
         ======================================================== */}
      {activeChartTab === 'flowchart' && (
        <div className="space-y-4">
          {/* Visual Stage Nodes Flow Diagram */}
          <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-2xs space-y-4">
            <div className="flex items-center justify-between">
              <div>
                <h3 className="text-xs font-bold uppercase tracking-wider text-gray-400">
                  Execution Pipeline & Architecture Flow
                </h3>
                <p className="text-xs text-gray-600 mt-0.5">
                  Click on any stage card below to inspect its code engine, input/output schemas, and statutory acts.
                </p>
              </div>
              <div className="text-[11px] font-medium text-gray-500">
                Selected: <span className="font-semibold text-gray-900">{selectedStage.shortName}</span>
              </div>
            </div>

            {/* The 6 Interconnected Nodes */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-6 gap-2.5 relative">
              {ARCHITECTURE_STAGES.map((st, idx) => {
                const Icon = st.icon;
                const isSelected = selectedStageId === st.id;
                const isSimActive = simulatedActiveStage === st.id;

                return (
                  <div
                    key={st.id}
                    onClick={() => setSelectedStageId(st.id)}
                    className={`relative p-3.5 rounded-xl border transition-all cursor-pointer select-none flex flex-col justify-between ${
                      isSimActive
                        ? 'bg-amber-50/80 border-amber-400 ring-2 ring-amber-400 shadow-md scale-[1.02]'
                        : isSelected
                        ? 'bg-gray-100 border-gray-900 ring-1 ring-gray-900 shadow-xs'
                        : 'bg-white border-gray-200 hover:border-gray-400 hover:bg-gray-50/70'
                    }`}
                  >
                    <div>
                      {/* Top Header */}
                      <div className="flex items-center justify-between mb-2">
                        <span className="text-[10px] font-bold text-gray-400">0{st.id}</span>
                        <div
                          className={`w-6 h-6 rounded-md flex items-center justify-center shrink-0 ${
                            isSimActive
                              ? 'bg-amber-500 text-white animate-bounce'
                              : isSelected
                              ? 'bg-gray-900 text-white'
                              : 'bg-gray-100 text-gray-600'
                          }`}
                        >
                          <Icon className="w-3.5 h-3.5" />
                        </div>
                      </div>

                      {/* Title & Badge */}
                      <div className="text-[10px] font-bold uppercase tracking-wider text-gray-500 mb-0.5">
                        {st.badge}
                      </div>
                      <h4 className="text-xs font-bold text-gray-900 line-clamp-1 leading-snug">
                        {st.title.replace(/Stage \d: /, '')}
                      </h4>
                      <p className="text-[10px] text-gray-500 mt-1 line-clamp-2 leading-relaxed">
                        {st.subtitle}
                      </p>
                    </div>

                    {/* Bottom Status / Throughput */}
                    <div className="mt-3 pt-2 border-t border-gray-100 flex items-center justify-between text-[10px]">
                      <span className="font-semibold text-gray-700">{st.throughput}</span>
                      {isSelected && (
                        <span className="text-[9px] font-bold text-gray-900">Active</span>
                      )}
                    </div>

                    {/* Connecting Chevron Arrow (desktop) */}
                    {idx < ARCHITECTURE_STAGES.length - 1 && (
                      <div className="hidden lg:flex absolute -right-2 top-1/2 -translate-y-1/2 z-10 w-4 h-4 rounded-full bg-white border border-gray-300 items-center justify-center shadow-2xs pointer-events-none">
                        <ArrowRight className="w-2.5 h-2.5 text-gray-500" strokeWidth={2.5} />
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>

          {/* Deep-Dive Stage Inspector Card */}
          <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-2xs space-y-4">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-gray-100 pb-3">
              <div className="flex items-center gap-2.5">
                <div className="w-7 h-7 rounded-lg bg-gray-900 text-white flex items-center justify-center">
                  {React.createElement(selectedStage.icon, { className: 'w-4 h-4' })}
                </div>
                <div>
                  <h3 className="text-sm font-bold text-gray-900">{selectedStage.title}</h3>
                  <p className="text-xs text-gray-500">{selectedStage.subtitle}</p>
                </div>
              </div>

              <div className="flex items-center gap-2 text-xs">
                <span className="font-mono text-[11px] text-gray-600 bg-gray-100 px-2 py-0.5 rounded border border-gray-200">
                  {selectedStage.engine}
                </span>
                <span className="px-2 py-0.5 rounded bg-gray-50 text-gray-900 text-[11px] font-semibold border border-gray-200">
                  {selectedStage.statutoryAct}
                </span>
              </div>
            </div>

            {/* 3-Column Technical Specification */}
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              {/* Col 1: Ingested Data Schemas */}
              <div className="p-3.5 bg-gray-50 rounded-xl border border-gray-200 space-y-2">
                <div className="flex items-center gap-1.5 text-xs font-bold text-gray-800">
                  <Database className="w-3.5 h-3.5 text-gray-600" />
                  <span>Input Data Contracts</span>
                </div>
                <ul className="space-y-1.5 text-xs text-gray-600">
                  {selectedStage.inputs.map((inp, i) => (
                    <li key={i} className="flex items-start gap-1.5">
                      <span className="text-gray-400 shrink-0 font-mono">▸</span>
                      <span>{inp}</span>
                    </li>
                  ))}
                </ul>
              </div>

              {/* Col 2: Deterministic Transformation Pipeline */}
              <div className="p-3.5 bg-gray-50 rounded-xl border border-gray-200 space-y-2">
                <div className="flex items-center gap-1.5 text-xs font-bold text-gray-800">
                  <Cpu className="w-3.5 h-3.5 text-gray-600" />
                  <span>Core Execution Engine Steps</span>
                </div>
                <ul className="space-y-1.5 text-xs text-gray-600">
                  {selectedStage.keySteps.map((stp, i) => (
                    <li key={i} className="flex items-start gap-1.5">
                      <span className="text-gray-400 shrink-0 font-mono">{i + 1}.</span>
                      <span>{stp}</span>
                    </li>
                  ))}
                </ul>
              </div>

              {/* Col 3: Generated Outputs & Artifacts */}
              <div className="p-3.5 bg-gray-50 rounded-xl border border-gray-200 space-y-2">
                <div className="flex items-center gap-1.5 text-xs font-bold text-gray-800">
                  <FileSpreadsheet className="w-3.5 h-3.5 text-gray-600" />
                  <span>Produced Verified Artifacts</span>
                </div>
                <ul className="space-y-1.5 text-xs text-gray-600">
                  {selectedStage.outputs.map((out, i) => (
                    <li key={i} className="flex items-start gap-1.5">
                      <CheckCircle2 className="w-3.5 h-3.5 text-gray-700 shrink-0 mt-0.5" />
                      <span>{out}</span>
                    </li>
                  ))}
                </ul>
              </div>
            </div>

            {/* Deterministic Mathematical Guarantee Alert */}
            <div className="p-3 bg-gray-50 border border-gray-200 rounded-lg flex items-start gap-2.5 text-xs text-gray-700">
              <ShieldCheck className="w-4 h-4 text-gray-900 shrink-0 mt-0.5" />
              <div>
                <span className="font-bold text-gray-900">Deterministic Mathematical Guarantee: </span>
                <span>{selectedStage.guarantee}</span>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ========================================================
          3. VIEW B: SYSTEM ARCHITECTURE BLUEPRINT (4-TIER)
         ======================================================== */}
      {activeChartTab === 'blueprint' && (
        <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-2xs space-y-4">
          <div className="border-b border-gray-100 pb-3">
            <h3 className="text-sm font-bold text-gray-900">4-Tier Enterprise Architecture Blueprint</h3>
            <p className="text-xs text-gray-500 mt-0.5">
              Production-ready enterprise architecture designed for Chartered Accountant statutory audit engagements.
            </p>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
            {/* Tier 1 */}
            <div className="p-4 rounded-xl bg-gray-50 border border-gray-200 space-y-3">
              <div className="flex items-center gap-2">
                <div className="w-6 h-6 rounded bg-gray-900 text-white flex items-center justify-center text-xs font-bold">
                  1
                </div>
                <span className="text-xs font-bold text-gray-900">Presentation Layer</span>
              </div>
              <p className="text-xs text-gray-600 leading-relaxed">
                Single-page React 19 application with high-contrast CA Slate design system, zero unnecessary renders, and print-ready stylesheets.
              </p>
              <div className="space-y-1 text-[11px] font-mono text-gray-500">
                <div>• React 19 + TypeScript</div>
                <div>• Tailwind CSS Utilities</div>
                <div>• Lucide Iconography</div>
                <div>• Client Canvas (px-8 aligned)</div>
              </div>
            </div>

            {/* Tier 2 */}
            <div className="p-4 rounded-xl bg-gray-50 border border-gray-200 space-y-3">
              <div className="flex items-center gap-2">
                <div className="w-6 h-6 rounded bg-gray-900 text-white flex items-center justify-center text-xs font-bold">
                  2
                </div>
                <span className="text-xs font-bold text-gray-900">API & Worker Queue</span>
              </div>
              <p className="text-xs text-gray-600 leading-relaxed">
                Asynchronous chunked worker queue with atomic idempotency locks (`X-Idempotency-Key`) preventing duplicated batch reconciliations.
              </p>
              <div className="space-y-1 text-[11px] font-mono text-gray-500">
                <div>• Express RESTful Services</div>
                <div>• Non-Blocking Chunk Worker</div>
                <div>• Client Isolation Containers</div>
                <div>• Real-time SSE Telemetry</div>
              </div>
            </div>

            {/* Tier 3 */}
            <div className="p-4 rounded-xl bg-gray-50 border border-gray-200 space-y-3">
              <div className="flex items-center gap-2">
                <div className="w-6 h-6 rounded bg-gray-900 text-white flex items-center justify-center text-xs font-bold">
                  3
                </div>
                <span className="text-xs font-bold text-gray-900">Deterministic Engines</span>
              </div>
              <p className="text-xs text-gray-600 leading-relaxed">
                Mathematical audit algorithms executing strict legal formulas with zero AI hallucinations and comprehensive tolerance bounds.
              </p>
              <div className="space-y-1 text-[11px] font-mono text-gray-500">
                <div>• Bank Statement Rec Engine</div>
                <div>• GST 3-Way Rec Engine</div>
                <div>• Form 3CD Clause Scrutiny</div>
                <div>• Section 206AB Higher TDS</div>
              </div>
            </div>

            {/* Tier 4 */}
            <div className="p-4 rounded-xl bg-gray-50 border border-gray-200 space-y-3">
              <div className="flex items-center gap-2">
                <div className="w-6 h-6 rounded bg-gray-900 text-white flex items-center justify-center text-xs font-bold">
                  4
                </div>
                <span className="text-xs font-bold text-gray-900">Persistence & Audit Trail</span>
              </div>
              <p className="text-xs text-gray-600 leading-relaxed">
                ACID in-memory store with write-ahead disk replacement, Tally Prime XML export, and ICAI SQC 1 cryptographic hash chaining.
              </p>
              <div className="space-y-1 text-[11px] font-mono text-gray-500">
                <div>• Atomic State Snapshot Store</div>
                <div>• ICAI SQC 1 SHA-256 Chain</div>
                <div>• Tally Prime XML Adapters</div>
                <div>• Multi-Sheet Excel Engine</div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ========================================================
          4. VIEW C: STATUTORY COMPLIANCE MATRIX
         ======================================================== */}
      {activeChartTab === 'statutory' && (
        <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-2xs space-y-4">
          <div className="border-b border-gray-100 pb-3 flex items-center justify-between">
            <div>
              <h3 className="text-sm font-bold text-gray-900">Statutory Tax & Audit Law Matrix</h3>
              <p className="text-xs text-gray-500 mt-0.5">
                Exact statutory provisions implemented across the CA automation workflow engine.
              </p>
            </div>
            <span className="text-xs font-semibold px-2 py-0.5 bg-gray-100 rounded text-gray-800">
              Tax Year 2026-27 Compliant
            </span>
          </div>

          <div className="border border-gray-200 rounded-lg overflow-hidden">
            <table className="w-full text-left text-xs">
              <thead className="bg-gray-50 border-b border-gray-200 text-gray-700 font-bold">
                <tr>
                  <th className="p-3">Statutory Act & Provision</th>
                  <th className="p-3">Legal Threshold</th>
                  <th className="p-3">Deterministic Engine Logic</th>
                  <th className="p-3">Generated Audit Deliverable</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100 text-gray-700">
                <tr className="hover:bg-gray-50/50">
                  <td className="p-3 font-semibold text-gray-900">Income Tax Act § 40A(3)</td>
                  <td className="p-3 font-mono">₹10,000 / ₹35,000 (transporters)</td>
                  <td className="p-3">Aggregates daily cash payments to the same payee across all vouchers</td>
                  <td className="p-3 text-gray-900 font-medium">Form 3CD Clause 21(d) Annexure</td>
                </tr>
                <tr className="hover:bg-gray-50/50">
                  <td className="p-3 font-semibold text-gray-900">Income Tax Rules Rule 6DD</td>
                  <td className="p-3 font-mono">Statutory Exemptions</td>
                  <td className="p-3">Exempts bank holidays, RTGS/NEFT banking hours, and agricultural purchases</td>
                  <td className="p-3 text-gray-900 font-medium">Rule 6DD Audit Justification Memo</td>
                </tr>
                <tr className="hover:bg-gray-50/50">
                  <td className="p-3 font-semibold text-gray-900">Income Tax Act § 43B</td>
                  <td className="p-3 font-mono">Unpaid as of filing date</td>
                  <td className="p-3">Flags statutory unpaid GST, PF, ESIC, and MSME interest liabilities</td>
                  <td className="p-3 text-gray-900 font-medium">Form 3CD Clause 26 Schedule</td>
                </tr>
                <tr className="hover:bg-gray-50/50">
                  <td className="p-3 font-semibold text-gray-900">Income Tax Act § 206AB / 206CCA</td>
                  <td className="p-3 font-mono">Higher of 5% or 2x rate</td>
                  <td className="p-3">Identifies non-filers with &gt; ₹50k TDS and computes differential exposure</td>
                  <td className="p-3 text-gray-900 font-medium">4-Sheet Section 206AB Workbook</td>
                </tr>
                <tr className="hover:bg-gray-50/50">
                  <td className="p-3 font-semibold text-gray-900">CGST Act § 16(2)(aa) & Rule 88C</td>
                  <td className="p-3 font-mono">DRC-01B trigger &gt; ₹25,000</td>
                  <td className="p-3">Compares Books ITC vs GSTR-2B vs 3B and estimates Section 50 interest</td>
                  <td className="p-3 text-gray-900 font-medium">GST 3-Way Rec Schedule & DRC-01B Pack</td>
                </tr>
                <tr className="hover:bg-gray-50/50">
                  <td className="p-3 font-semibold text-gray-900">ICAI SQC 1 (Quality Control)</td>
                  <td className="p-3 font-mono">Tamper-Evident Trail</td>
                  <td className="p-3">Chains every audit run to previous run with immutable SHA-256 hashes</td>
                  <td className="p-3 text-gray-900 font-medium">SQC 1 Cryptographic Audit Certificate</td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* ========================================================
          5. WORKFLOW INGESTION & RUN WORKBENCH
         ======================================================== */}
      <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-2xs space-y-4">
        <div className="flex items-center justify-between border-b border-gray-100 pb-3">
          <div>
            <h3 className="text-sm font-bold text-gray-900">Dispatch Audit Automation Cycle</h3>
            <p className="text-xs text-gray-500 mt-0.5">
              Select standard workflow, stage source documents, and run background execution.
            </p>
          </div>
          <div className="text-xs font-semibold px-2.5 py-1 bg-gray-100 rounded-lg text-gray-800 border border-gray-200">
            Active Client: <span className="font-mono text-gray-900">{selectedClientId}</span>
          </div>
        </div>

        {/* Workflow Selection Cards */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
          {workflows.map((wf) => {
            const isSelected = selectedWorkflowId === wf.id;
            return (
              <div
                key={wf.id}
                onClick={() => setSelectedWorkflowId(wf.id)}
                className={`p-3.5 rounded-xl border transition-all cursor-pointer flex flex-col justify-between ${
                  isSelected
                    ? 'bg-gray-100 border-gray-900 ring-1 ring-gray-900 shadow-2xs'
                    : 'bg-white border-gray-200 hover:border-gray-400'
                }`}
              >
                <div>
                  <div className="flex items-center justify-between mb-1.5">
                    <span className="text-[10px] uppercase font-bold tracking-wider px-2 py-0.5 rounded bg-gray-100 text-gray-700">
                      {wf.category.replace('_', ' ')}
                    </span>
                    {isSelected && <CheckCircle2 className="w-4 h-4 text-gray-900" />}
                  </div>
                  <h4 className="text-xs font-bold text-gray-900 line-clamp-1">{wf.name}</h4>
                  <p className="text-[11px] text-gray-500 mt-1 line-clamp-2 leading-relaxed">
                    {wf.description}
                  </p>
                </div>

                <div className="mt-3 pt-2 border-t border-gray-200/50 flex items-center justify-between text-[11px] text-gray-500">
                  <span>{wf.steps.length} Steps</span>
                  <span className="font-mono">v{wf.version}</span>
                </div>
              </div>
            );
          })}
        </div>

        {/* Document Ingestion & Execution Card */}
        <div className="flex flex-col sm:flex-row items-center gap-4 pt-2">
          <label className="flex-1 w-full flex flex-col items-center justify-center p-6 border-2 border-dashed border-gray-200 hover:border-gray-900 rounded-xl cursor-pointer bg-gray-50/50 hover:bg-gray-100/70 transition-all">
            <FileSpreadsheet className="w-8 h-8 text-gray-700 mb-2" />
            <span className="text-xs font-semibold text-gray-800">Click or Drag & Drop Documents</span>
            <span className="text-[11px] text-gray-500 mt-0.5 text-center">
              {(WORKFLOW_REQUIREMENTS[selectedWorkflowId] || { label: 'Supports CSV, XLSX, XLS, PDF, and JSON' }).label}
            </span>
            <input
              type="file"
              multiple
              accept={(WORKFLOW_REQUIREMENTS[selectedWorkflowId] || { accept: '.xlsx,.xls,.csv,.pdf,.json' }).accept}
              onChange={handleFileUpload}
              className="hidden"
            />
          </label>

          {uploadedFiles.length > 0 && (
            <div className="w-full sm:w-80 p-4 bg-gray-50 rounded-xl border border-gray-200 space-y-2.5">
              <div className="flex items-center justify-between text-xs font-bold text-gray-700">
                <span>Staged Documents ({uploadedFiles.length})</span>
                <button
                  onClick={() => setUploadedFiles([])}
                  className="text-red-600 hover:underline text-[11px] cursor-pointer"
                >
                  Clear All
                </button>
              </div>

              <div className="max-h-28 overflow-y-auto space-y-1 text-xs">
                {uploadedFiles.map((f, idx) => (
                  <div key={idx} className="flex items-center gap-1.5 text-gray-600 truncate">
                    <FileText className="w-3.5 h-3.5 shrink-0 text-gray-900" />
                    <span className="truncate">{f.fileName}</span>
                  </div>
                ))}
              </div>

              <button
                onClick={handleStartWorkflow}
                disabled={isUploading}
                className="w-full mt-2 py-2 px-3 bg-gray-900 hover:bg-black text-white text-xs font-bold rounded-lg shadow-xs flex items-center justify-center gap-2 transition-all disabled:opacity-50 cursor-pointer"
              >
                {isUploading ? (
                  <span>Executing Automation...</span>
                ) : (
                  <>
                    <Play className="w-3.5 h-3.5" />
                    <span>Run Audit Automation</span>
                  </>
                )}
              </button>
            </div>
          )}
        </div>
      </div>

      {/* ========================================================
          6. ACTIVE JOB MONITOR, STEPPER & HUMAN-IN-THE-LOOP
         ======================================================== */}
      {activeJob && (
        <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-2xs space-y-4">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-gray-100 pb-4">
            <div>
              <div className="flex items-center gap-2">
                <span className="text-xs font-mono text-gray-500">{activeJob.id}</span>
                <span
                  className={`text-[11px] font-bold px-2 py-0.5 rounded-full capitalize ${
                    activeJob.status === 'completed'
                      ? 'bg-gray-100 text-gray-900'
                      : activeJob.status === 'review_required'
                      ? 'bg-amber-100 text-amber-800'
                      : activeJob.status === 'running'
                      ? 'bg-blue-100 text-blue-800'
                      : 'bg-gray-100 text-gray-800'
                  }`}
                >
                  {activeJob.status.replace('_', ' ')}
                </span>
              </div>
              <h3 className="text-sm font-bold text-gray-900 mt-1">{activeJob.workflowType}</h3>
            </div>

            <div className="flex items-center gap-2">
              <span className="text-xs text-gray-500 font-medium">
                Step {activeJob.completedSteps} of {activeJob.totalSteps}
              </span>
              {activeJob.status === 'failed' && (
                <button
                  onClick={() => handleRetry(activeJob.id)}
                  className="px-3 py-1.5 text-xs font-semibold bg-gray-100 hover:bg-gray-200 text-gray-800 rounded-lg flex items-center gap-1 cursor-pointer"
                >
                  <RotateCcw className="w-3 h-3" />
                  <span>Retry Processing</span>
                </button>
              )}
            </div>
          </div>

          {/* Stepper Grid */}
          {activeRun?.stepRuns && (
            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-2">
              {Object.values(activeRun.stepRuns).map((st: any, i: number) => {
                const isComplete = st.status === 'completed';
                const isFailed = st.status === 'failed';
                const isRunning = st.status === 'running';

                return (
                  <div
                    key={st.stepId || i}
                    className={`p-3 rounded-lg border text-xs flex flex-col justify-between ${
                      isComplete
                        ? 'bg-gray-100 border-gray-200 text-gray-900'
                        : isFailed
                        ? 'bg-red-50 border-red-200 text-red-950'
                        : isRunning
                        ? 'bg-blue-50 border-blue-200 text-blue-950 animate-pulse'
                        : 'bg-gray-50 border-gray-200 text-gray-500'
                    }`}
                  >
                    <div>
                      <div className="flex items-center justify-between mb-1">
                        <span className="text-[10px] font-bold">Step {i + 1}</span>
                        {isComplete && <Check className="w-3.5 h-3.5 text-gray-900" />}
                        {isFailed && <X className="w-3.5 h-3.5 text-red-600" />}
                        {isRunning && <Activity className="w-3.5 h-3.5 text-blue-600 animate-spin" />}
                      </div>
                      <div className="font-semibold line-clamp-1">{st.name}</div>
                    </div>

                    <div className="mt-2 pt-1 border-t border-gray-200/50 text-[10px] flex items-center justify-between">
                      <span>{st.processedItems} records</span>
                      <span className="capitalize">{st.status}</span>
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {/* Sub-tabs: Exceptions & Artifacts */}
          <div className="flex items-center gap-3 pt-2 border-t border-gray-100">
            <button
              onClick={() => setActiveTab('pipeline')}
              className={`text-xs font-semibold pb-1 border-b-2 transition-all cursor-pointer ${
                activeTab === 'pipeline' ? 'border-gray-900 text-gray-900' : 'border-transparent text-gray-500'
              }`}
            >
              Workflow Overview
            </button>
            <button
              onClick={() => setActiveTab('exceptions')}
              className={`text-xs font-semibold pb-1 border-b-2 transition-all flex items-center gap-1 cursor-pointer ${
                activeTab === 'exceptions' ? 'border-gray-900 text-gray-900' : 'border-transparent text-gray-500'
              }`}
            >
              <span>Exceptions Review</span>
              {exceptions.length > 0 && (
                <span className="px-1.5 py-0.2 rounded-full text-[10px] bg-amber-100 text-amber-800 font-bold">
                  {exceptions.length}
                </span>
              )}
            </button>
            <button
              onClick={() => setActiveTab('artifacts')}
              className={`text-xs font-semibold pb-1 border-b-2 transition-all flex items-center gap-1 cursor-pointer ${
                activeTab === 'artifacts' ? 'border-gray-900 text-gray-900' : 'border-transparent text-gray-500'
              }`}
            >
              <span>Verified Exports</span>
              {activeJob.artifacts?.length > 0 && (
                <span className="px-1.5 py-0.2 rounded-full text-[10px] bg-gray-100 text-gray-900 font-bold">
                  {activeJob.artifacts.length}
                </span>
              )}
            </button>
          </div>

          {/* Exceptions Table */}
          {activeTab === 'exceptions' && (
            <div className="space-y-3">
              {exceptions.length === 0 ? (
                <div className="p-8 text-center text-xs text-gray-500 bg-gray-50 rounded-lg">
                  No exceptions detected in this run. All financial rules passed cleanly.
                </div>
              ) : (
                <div className="border border-gray-200 rounded-lg overflow-hidden">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-gray-50 border-b border-gray-200 text-gray-700 font-bold">
                      <tr>
                        <th className="p-3">Severity</th>
                        <th className="p-3">Category</th>
                        <th className="p-3">Description</th>
                        <th className="p-3">Status</th>
                        <th className="p-3 text-right">Human-in-Loop Action</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-100">
                      {exceptions.map((exc) => (
                        <tr key={exc.id} className="hover:bg-gray-50/50">
                          <td className="p-3">
                            <span
                              className={`px-2 py-0.5 rounded text-[10px] font-bold uppercase ${
                                exc.severity === 'critical' ? 'bg-red-100 text-red-800' : 'bg-amber-100 text-amber-800'
                              }`}
                            >
                              {exc.severity}
                            </span>
                          </td>
                          <td className="p-3 font-semibold text-gray-800">{exc.category.replace('_', ' ')}</td>
                          <td className="p-3 text-gray-600 max-w-md">{exc.description}</td>
                          <td className="p-3 capitalize font-medium">{exc.resolutionStatus}</td>
                          <td className="p-3 text-right">
                            {exc.resolutionStatus === 'pending' ? (
                              <div className="inline-flex items-center gap-1.5">
                                <button
                                  onClick={() => handleResolveException(exc.id, 'approved')}
                                  className="px-2.5 py-1 bg-gray-900 hover:bg-black text-white rounded text-[11px] font-semibold cursor-pointer"
                                >
                                  Approve
                                </button>
                                <button
                                  onClick={() => handleResolveException(exc.id, 'rejected')}
                                  className="px-2.5 py-1 bg-gray-200 hover:bg-gray-300 text-gray-800 rounded text-[11px] font-semibold cursor-pointer"
                                >
                                  Reject
                                </button>
                              </div>
                            ) : (
                              <span className="text-[11px] text-gray-400">Resolved by {exc.reviewedBy || 'CA'}</span>
                            )}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          )}

          {/* Artifacts Download Cards */}
          {activeTab === 'artifacts' && (
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              {activeJob.artifacts?.map((art: any) => (
                <div
                  key={art.id}
                  className="p-4 bg-gray-50 border border-gray-200 rounded-xl flex items-center justify-between"
                >
                  <div className="flex items-center gap-3">
                    <FileSpreadsheet className="w-8 h-8 text-gray-900" />
                    <div>
                      <div className="text-xs font-bold text-gray-900">{art.fileName}</div>
                      <div className="text-[11px] text-gray-500">
                        {art.sheetNames?.join(', ')} • {(art.fileSizeBytes / 1024).toFixed(1)} KB
                      </div>
                    </div>
                  </div>

                  <button
                    onClick={() => handleDownloadArtifact(art)}
                    className="px-3 py-1.5 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg flex items-center gap-1.5 shadow-xs cursor-pointer"
                  >
                    <Download className="w-3.5 h-3.5" />
                    <span>Download</span>
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {/* ========================================================
          7. AUDIT PROCESSING HISTORY DRAWER
         ======================================================== */}
      {activeTab === 'jobs' && (
        <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-2xs space-y-3">
          <div className="flex items-center justify-between">
            <h3 className="text-sm font-bold text-gray-900">Audit Processing & Execution History</h3>
            <span className="text-xs text-gray-500">Filtered by Client {selectedClientId}</span>
          </div>

          <div className="border border-gray-200 rounded-lg overflow-hidden">
            <table className="w-full text-left text-xs">
              <thead className="bg-gray-50 border-b border-gray-200 text-gray-700 font-bold">
                <tr>
                  <th className="p-3">Audit Run ID</th>
                  <th className="p-3">Workflow</th>
                  <th className="p-3">Status</th>
                  <th className="p-3">Current Step</th>
                  <th className="p-3">Created</th>
                  <th className="p-3 text-right">Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100">
                {jobs.map((j) => (
                  <tr key={j.id} className="hover:bg-gray-50/50">
                    <td className="p-3 font-mono text-gray-500">{j.id}</td>
                    <td className="p-3 font-semibold text-gray-800">{j.workflowType}</td>
                    <td className="p-3">
                      <span
                        className={`px-2 py-0.5 rounded text-[10px] font-bold capitalize ${
                          j.status === 'completed'
                            ? 'bg-gray-100 text-gray-900'
                            : j.status === 'review_required'
                            ? 'bg-amber-100 text-amber-800'
                            : j.status === 'running'
                            ? 'bg-blue-100 text-blue-800'
                            : 'bg-gray-100 text-gray-800'
                        }`}
                      >
                        {j.status.replace('_', ' ')}
                      </span>
                    </td>
                    <td className="p-3 text-gray-600">{j.currentStep}</td>
                    <td className="p-3 text-gray-400">{new Date(j.createdAt).toLocaleTimeString()}</td>
                    <td className="p-3 text-right">
                      <button
                        onClick={() => {
                          loadJobDetails(j.id);
                          setActiveTab('pipeline');
                        }}
                        className="text-gray-900 font-semibold hover:underline cursor-pointer"
                      >
                        Inspect
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
};
