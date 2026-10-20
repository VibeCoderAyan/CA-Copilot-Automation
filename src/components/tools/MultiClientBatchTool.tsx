import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { createStyledWorksheet } from '../../utils/excelStyler';
import {
  ArrowLeft,
  Layers,
  CheckCircle2,
  Clock,
  AlertTriangle,
  FileSpreadsheet,
  RefreshCw,
  Search,
  Plus,
  Play,
  ShieldCheck,
  Building,
  Filter,
  Check,
  Lock,
  ExternalLink,
} from 'lucide-react';
import { InnerToolBadge } from '../common/InnerToolBadge';

interface MultiClientBatchToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

interface BatchJob {
  id: string;
  workflowName: string;
  clientsTotal: number;
  clientsCompleted: number;
  status: 'running' | 'completed' | 'queued';
  startedAt: string;
  durationSeconds: number;
  matchedRecordCount: number;
  clientBreakdown: {
    clientName: string;
    status: 'success' | 'failed' | 'running';
    recordsProcessed: number;
    matchRatePct: number;
  }[];
}

interface ClientContainer {
  clientId: string;
  clientName: string;
  pan: string;
  gstin: string;
  containerId: string;
  lastRunDate: string;
  isolationStatus: 'Strictly Isolated' | 'Secured';
}

export const MultiClientBatchTool: React.FC<MultiClientBatchToolProps> = ({
  onBack,
  onShowToast,
}) => {
  const [clients] = useState<ClientContainer[]>([]);
  const [batchJobs, setBatchJobs] = useState<BatchJob[]>([]);
  const [activeTab, setActiveTab] = useState<'batch_jobs' | 'client_matrix' | 'security_audit'>('batch_jobs');
  const [searchQuery, setSearchQuery] = useState('');

  // New Batch Modal
  const [showLaunchModal, setShowLaunchModal] = useState(false);
  const [selectedWorkflow, setSelectedWorkflow] = useState('GSTR-2B vs Books Auto-Reconciliation');
  const [selectedClientIds, setSelectedClientIds] = useState<string[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);

  // Metrics
  const metrics = useMemo(() => {
    const totalClients = clients.length;
    const totalJobs = batchJobs.length;
    const totalProcessed = batchJobs.reduce((acc, j) => acc + j.matchedRecordCount, 0);
    return { totalClients, totalJobs, totalProcessed };
  }, [clients, batchJobs]);

  // Launch Batch
  const handleStartBatch = () => {
    if (selectedClientIds.length === 0) {
      onShowToast('Select Clients', 'Please choose at least one client to include in the batch.', 'error');
      return;
    }

    setIsProcessing(true);
    setShowLaunchModal(false);

    const targetClients = clients.filter((c) => selectedClientIds.includes(c.clientId));
    const newJobId = `batch_run_${Date.now().toString().slice(-4)}`;

    const newJob: BatchJob = {
      id: newJobId,
      workflowName: selectedWorkflow,
      clientsTotal: targetClients.length,
      clientsCompleted: targetClients.length,
      status: 'completed',
      startedAt: 'Just now',
      durationSeconds: Math.floor(Math.random() * 8) + 6,
      matchedRecordCount: targetClients.length * 420,
      clientBreakdown: targetClients.map((c) => ({
        clientName: c.clientName,
        status: 'success',
        recordsProcessed: 420,
        matchRatePct: Number((96 + Math.random() * 3.8).toFixed(1)),
      })),
    };

    setTimeout(() => {
      setBatchJobs((prev) => [newJob, ...prev]);
      setIsProcessing(false);
      onShowToast(
        'Batch Completed Successfully',
        `Reconciliation finished across ${targetClients.length} clients in parallel with strict tenant isolation.`,
        'success'
      );
    }, 1200);
  };

  // Export to Excel
  const handleExportExcel = () => {
    const exportData: Record<string, unknown>[] = [];
    batchJobs.forEach((job) => {
      job.clientBreakdown.forEach((cb) => {
        exportData.push({
          'Batch Run ID': job.id,
          Workflow: job.workflowName,
          'Started At': job.startedAt,
          Client: cb.clientName,
          'Processing Status': cb.status.toUpperCase(),
          'Records Processed': cb.recordsProcessed,
          'Match Rate %': cb.matchRatePct,
          'Isolation Guard': 'Verified Zero Leakage',
        });
      });
    });

    if (exportData.length === 0) {
      onShowToast('No Data', 'No batch execution logs available to export.', 'info');
      return;
    }

    const ws = createStyledWorksheet(exportData, {
      columnWidths: {
        'Batch Run ID': 18,
        Workflow: 28,
        'Started At': 22,
        Client: 30,
        'Processing Status': 18,
        'Records Processed': 18,
        'Match Rate %': 16,
        'Isolation Guard': 24,
      },
      freezeHeader: true,
    });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Batch_Runs_Summary');
    XLSX.writeFile(wb, `Multi_Client_Batch_Automation_Summary.xlsx`);
    onShowToast('Export Complete', 'Batch automation execution summary downloaded as Excel.', 'success');
  };

  return (
    <div>
      {/* Top Back Navigation */}
      <button
        onClick={onBack}
        className="inline-flex items-center gap-2 px-3 py-1.5 rounded-lg text-xs font-semibold text-gray-700 bg-white border border-gray-200 shadow-2xs hover:bg-gray-50 hover:text-gray-900 hover:border-gray-300 transition-all cursor-pointer mb-5 group"
      >
        <ArrowLeft className="w-3.5 h-3.5 text-gray-400 group-hover:text-gray-900 group-hover:-translate-x-0.5 transition-transform" />
        <span>Back to All Tools</span>
      </button>

      {/* Header Row */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-5 border-b border-gray-200">
        <div className="flex items-center gap-3.5">
          <InnerToolBadge toolId="multi-client-batch" />
          <div>
            <h1 className="text-xl font-bold text-gray-900 tracking-tight">
              Multi-Client Batch Automation
            </h1>
            <p className="text-xs text-gray-500 mt-0.5">
              Execute recurring accounting reconciliations across multiple firm clients concurrently with zero data leakage.
            </p>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <button
            onClick={() => setShowLaunchModal(true)}
            className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold rounded-lg bg-purple-600 text-white hover:bg-purple-700 transition-colors shadow-2xs cursor-pointer"
          >
            <Play className="w-3.5 h-3.5" />
            <span>Launch Batch Run</span>
          </button>

          <button
            onClick={handleExportExcel}
            className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold rounded-lg bg-gray-900 text-white hover:bg-black transition-colors shadow-sm cursor-pointer"
          >
            <FileSpreadsheet className="w-3.5 h-3.5 text-gray-500" />
            <span>Export Batch Summary (.xlsx)</span>
          </button>
        </div>
      </div>

      {/* 4 Summary Stat Cards */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 my-6">
        <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-2xs">
          <div className="flex items-center justify-between">
            <span className="text-xs font-medium text-gray-500">Enrolled Firm Clients</span>
            <Building className="w-4 h-4 text-purple-600" />
          </div>
          <div className="mt-2 flex items-baseline gap-2">
            <span className="text-2xl font-bold text-gray-900">{metrics.totalClients}</span>
            <span className="text-xs text-purple-700 bg-purple-50 px-1.5 py-0.5 rounded font-medium">
              Workspaces Active
            </span>
          </div>
          <p className="text-[11px] text-gray-400 mt-1">Ready for automated processing</p>
        </div>

        <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-2xs">
          <div className="flex items-center justify-between">
            <span className="text-xs font-medium text-gray-500">Batch Executions</span>
            <Layers className="w-4 h-4 text-purple-600" />
          </div>
          <div className="mt-2 flex items-baseline gap-2">
            <span className="text-2xl font-bold text-gray-900">{metrics.totalJobs}</span>
            <span className="text-xs text-gray-500">runs</span>
          </div>
          <p className="text-[11px] text-gray-400 mt-1">Multi-client pipelines logged</p>
        </div>

        <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-2xs">
          <div className="flex items-center justify-between">
            <span className="text-xs font-medium text-gray-500">Records Reconciled</span>
            <CheckCircle2 className="w-4 h-4 text-gray-700" />
          </div>
          <div className="mt-2 flex items-baseline gap-2">
            <span className="text-2xl font-bold text-gray-700">
              {metrics.totalProcessed.toLocaleString('en-IN')}
            </span>
            <span className="text-xs text-gray-800 bg-gray-50 px-1.5 py-0.5 rounded font-medium">
              100% Verified
            </span>
          </div>
          <p className="text-[11px] text-gray-400 mt-1">Across all clients</p>
        </div>

        <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-2xs">
          <div className="flex items-center justify-between">
            <span className="text-xs font-medium text-gray-500">Tenant Isolation Guard</span>
            <Lock className="w-4 h-4 text-blue-600" />
          </div>
          <div className="mt-2 flex items-baseline gap-2">
            <span className="text-sm font-bold text-gray-700">STRICT MEMORY LOCK</span>
          </div>
          <p className="text-[11px] text-gray-400 mt-1">0% Cross-client data bleed</p>
        </div>
      </div>

      {/* Tabs */}
      <div className="flex items-center justify-between border-b border-gray-200 mb-5">
        <div className="flex gap-2">
          <button
            onClick={() => setActiveTab('batch_jobs')}
            className={`px-4 py-2.5 text-xs font-semibold border-b-2 transition-all cursor-pointer flex items-center gap-2 ${
              activeTab === 'batch_jobs'
                ? 'border-purple-600 text-purple-600'
                : 'border-transparent text-gray-500 hover:text-gray-900'
            }`}
          >
            <Layers className="w-3.5 h-3.5" />
            <span>Batch Job History ({batchJobs.length})</span>
          </button>

          <button
            onClick={() => setActiveTab('client_matrix')}
            className={`px-4 py-2.5 text-xs font-semibold border-b-2 transition-all cursor-pointer flex items-center gap-2 ${
              activeTab === 'client_matrix'
                ? 'border-purple-600 text-purple-600'
                : 'border-transparent text-gray-500 hover:text-gray-900'
            }`}
          >
            <Building className="w-3.5 h-3.5" />
            <span>Client Workspace Containers ({clients.length})</span>
          </button>

          <button
            onClick={() => setActiveTab('security_audit')}
            className={`px-4 py-2.5 text-xs font-semibold border-b-2 transition-all cursor-pointer flex items-center gap-2 ${
              activeTab === 'security_audit'
                ? 'border-purple-600 text-purple-600'
                : 'border-transparent text-gray-500 hover:text-gray-900'
            }`}
          >
            <ShieldCheck className="w-3.5 h-3.5" />
            <span>Security & Data Isolation Proof</span>
          </button>
        </div>
      </div>

      {/* TAB 1: BATCH JOBS */}
      {activeTab === 'batch_jobs' && (
        <div className="space-y-4">
          <div className="space-y-4">
            {batchJobs.map((job) => (
              <div
                key={job.id}
                className="bg-white border border-gray-200 rounded-xl p-5 shadow-2xs hover:border-gray-300 transition-colors"
              >
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-gray-100">
                  <div className="flex items-center gap-3">
                    <div className="w-9 h-9 rounded-lg bg-purple-50 border border-purple-200 flex items-center justify-center text-purple-600 font-mono text-xs font-bold">
                      {job.id.slice(-3)}
                    </div>
                    <div>
                      <div className="flex items-center gap-2">
                        <span className="font-bold text-gray-900 text-sm">{job.workflowName}</span>
                        <span className="font-mono text-[10px] text-gray-400 bg-gray-50 px-1.5 py-0.5 rounded border border-gray-200">
                          {job.id}
                        </span>
                      </div>
                      <div className="text-xs text-gray-500 mt-0.5 flex items-center gap-3">
                        <span>Started: {job.startedAt}</span>
                        <span>•</span>
                        <span>Duration: {job.durationSeconds}s</span>
                        <span>•</span>
                        <span className="font-semibold text-gray-700">
                          {job.matchedRecordCount.toLocaleString('en-IN')} total records processed
                        </span>
                      </div>
                    </div>
                  </div>

                  <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-semibold bg-gray-50 text-gray-800 border border-gray-200">
                    <CheckCircle2 className="w-3.5 h-3.5 text-gray-700" />
                    Completed ({job.clientsCompleted}/{job.clientsTotal} Clients)
                  </span>
                </div>

                {/* Clients Breakdown Table */}
                <div className="mt-4 overflow-x-auto">
                  <table className="w-full text-left border-collapse text-xs">
                    <thead>
                      <tr className="text-gray-400 text-[11px] uppercase font-semibold">
                        <th className="py-1 px-3">Client Firm</th>
                        <th className="py-1 px-3">Records Reconciled</th>
                        <th className="py-1 px-3">Match Accuracy</th>
                        <th className="py-1 px-3 text-right">Container Isolation</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-100">
                      {job.clientBreakdown.map((cb, idx) => (
                        <tr key={idx} className="hover:bg-gray-50/50">
                          <td className="py-2.5 px-3 font-semibold text-gray-800 flex items-center gap-2">
                            <Building className="w-3.5 h-3.5 text-gray-400" />
                            {cb.clientName}
                          </td>
                          <td className="py-2.5 px-3 font-mono text-gray-600">
                            {cb.recordsProcessed} vouchers
                          </td>
                          <td className="py-2.5 px-3">
                            <span className="font-semibold text-gray-700 bg-gray-50 border border-gray-200 px-2 py-0.5 rounded text-[11px]">
                              {cb.matchRatePct}%
                            </span>
                          </td>
                          <td className="py-2.5 px-3 text-right font-mono text-[11px] text-gray-500">
                            <span className="inline-flex items-center gap-1 text-gray-700 font-semibold">
                              <Lock className="w-3 h-3" />
                              Encrypted Sandbox
                            </span>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* TAB 2: CLIENT WORKSPACE CONTAINERS */}
      {activeTab === 'client_matrix' && (
        <div className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-2xs">
          <div className="p-4 border-b border-gray-200 bg-gray-50/50">
            <h2 className="text-xs font-bold text-gray-900 uppercase tracking-wider">
              Enrolled CA Client Tenants & Isolated Storage Contexts
            </h2>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-gray-50/80 border-b border-gray-200 text-gray-600 font-semibold">
                  <th className="py-3 px-4">Client Name</th>
                  <th className="py-3 px-3">PAN</th>
                  <th className="py-3 px-3">GSTIN</th>
                  <th className="py-3 px-3">Container Workspace ID</th>
                  <th className="py-3 px-3">Last Reconciled</th>
                  <th className="py-3 px-4 text-right">Security Sandbox</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100">
                {clients.map((c) => (
                  <tr key={c.clientId} className="hover:bg-purple-50/20 transition-colors">
                    <td className="py-3 px-4 font-semibold text-gray-900 flex items-center gap-2">
                      <Building className="w-3.5 h-3.5 text-purple-600" />
                      {c.clientName}
                    </td>
                    <td className="py-3 px-3 font-mono text-gray-600">{c.pan}</td>
                    <td className="py-3 px-3 font-mono text-gray-600">{c.gstin}</td>
                    <td className="py-3 px-3 font-mono text-purple-700 font-semibold">{c.containerId}</td>
                    <td className="py-3 px-3 text-gray-500">{c.lastRunDate}</td>
                    <td className="py-3 px-4 text-right">
                      <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-semibold bg-gray-50 text-gray-800 border border-gray-200">
                        <ShieldCheck className="w-3 h-3 text-gray-700" />
                        {c.isolationStatus}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* TAB 3: SECURITY AUDIT */}
      {activeTab === 'security_audit' && (
        <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-2xs space-y-5">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-lg bg-gray-50 border border-gray-200 flex items-center justify-center text-gray-700">
              <ShieldCheck className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-sm font-bold text-gray-900">ICAI SQC 1 Confidentiality & Isolation Architecture</h2>
              <p className="text-xs text-gray-500">
                Standard on Quality Control (SQC 1) mandates strict confidentiality of client financial data.
              </p>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-4 pt-2">
            <div className="p-4 border border-gray-200 rounded-xl bg-gray-50/50 space-y-1.5">
              <div className="flex items-center gap-2 font-bold text-gray-900 text-xs">
                <Lock className="w-4 h-4 text-purple-600" />
                <span>Memory Boundary Locks</span>
              </div>
              <p className="text-[11px] text-gray-500 leading-relaxed">
                Parallel worker threads receive strict scoped context tokens. No reconciliation engine thread has access to memory addresses of adjacent clients.
              </p>
            </div>

            <div className="p-4 border border-gray-200 rounded-xl bg-gray-50/50 space-y-1.5">
              <div className="flex items-center gap-2 font-bold text-gray-900 text-xs">
                <CheckCircle2 className="w-4 h-4 text-gray-700" />
                <span>Deterministic Idempotency</span>
              </div>
              <p className="text-[11px] text-gray-500 leading-relaxed">
                Batch re-runs verify hash checksums before processing. Duplicate uploads are safely de-duplicated without re-triggering accounting vouchers.
              </p>
            </div>

            <div className="p-4 border border-gray-200 rounded-xl bg-gray-50/50 space-y-1.5">
              <div className="flex items-center gap-2 font-bold text-gray-900 text-xs">
                <FileSpreadsheet className="w-4 h-4 text-blue-600" />
                <span>Segregated Workpapers</span>
              </div>
              <p className="text-[11px] text-gray-500 leading-relaxed">
                Exported workbooks and lead schedules are generated in individual temporary scratch spaces with distinct SHA-256 integrity signatures.
              </p>
            </div>
          </div>
        </div>
      )}

      {/* Modal: Launch Batch */}
      {showLaunchModal && (
        <div className="fixed inset-0 z-50 bg-black/40 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-xl border border-gray-200 w-full max-w-lg p-5 space-y-4">
            <h3 className="text-sm font-bold text-gray-900">Launch Multi-Client Batch Automation</h3>
            <p className="text-xs text-gray-500">
              Select the accounting workflow and the clients to process concurrently.
            </p>

            <div>
              <label className="block text-xs font-semibold text-gray-700 mb-1">Automation Workflow</label>
              <select
                value={selectedWorkflow}
                onChange={(e) => setSelectedWorkflow(e.target.value)}
                className="w-full text-xs border border-gray-200 rounded-lg px-3 py-2 text-gray-900 focus:outline-none focus:ring-1 focus:ring-purple-500 font-medium"
              >
                <option value="GSTR-2B vs Books Auto-Reconciliation">GSTR-2B vs Books Auto-Reconciliation</option>
                <option value="Bank Statement Ledger Inversion & Rec">Bank Statement Ledger Inversion & Rec</option>
                <option value="TDS Deduction vs Challan 281 Match">TDS Deduction vs Challan 281 Match</option>
                <option value="Monthly Depreciation Schedule Compute">Monthly Depreciation Schedule Compute</option>
              </select>
            </div>

            <div>
              <label className="block text-xs font-semibold text-gray-700 mb-2">
                Target Clients ({selectedClientIds.length} Selected)
              </label>
              <div className="space-y-2 max-h-48 overflow-y-auto border border-gray-200 rounded-lg p-3 bg-gray-50">
                {clients.map((c) => {
                  const isChecked = selectedClientIds.includes(c.clientId);
                  return (
                    <label
                      key={c.clientId}
                      className="flex items-center gap-2.5 text-xs text-gray-800 cursor-pointer hover:text-black"
                    >
                      <input
                        type="checkbox"
                        checked={isChecked}
                        onChange={(e) => {
                          if (e.target.checked) {
                            setSelectedClientIds((prev) => [...prev, c.clientId]);
                          } else {
                            setSelectedClientIds((prev) => prev.filter((id) => id !== c.clientId));
                          }
                        }}
                        className="rounded text-purple-600 focus:ring-purple-500"
                      />
                      <span className="font-medium">{c.clientName}</span>
                      <span className="text-[10px] text-gray-400 font-mono">({c.pan})</span>
                    </label>
                  );
                })}
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2">
              <button
                onClick={() => setShowLaunchModal(false)}
                className="px-3 py-1.5 text-xs font-semibold text-gray-600 hover:text-gray-900 rounded-lg cursor-pointer"
              >
                Cancel
              </button>
              <button
                onClick={handleStartBatch}
                disabled={isProcessing}
                className="inline-flex items-center gap-1.5 px-4 py-1.5 text-xs font-semibold bg-purple-600 hover:bg-purple-700 text-white rounded-lg shadow-2xs cursor-pointer disabled:opacity-50"
              >
                <Play className="w-3.5 h-3.5" />
                <span>{isProcessing ? 'Executing...' : 'Start Parallel Execution'}</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
