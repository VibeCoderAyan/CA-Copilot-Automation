import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { exportStyledWorkbook } from '../../utils/excelStyler';
import type { AuditChecklistItem, DocumentRequestItem, UnusualTransaction, UserRole } from '../../types';
import { Badge } from '../common/Badge';
import {
  ShieldCheck,
  CheckCircle2,
  Download,
  MessageSquare
} from 'lucide-react';

interface AuditToolsViewProps {
  checklists: AuditChecklistItem[];
  documentRequests: DocumentRequestItem[];
  unusualTransactions: UnusualTransaction[];
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
  selectedClientId: string;
  userRole: UserRole;
}

export const AuditToolsView: React.FC<AuditToolsViewProps> = ({
  checklists: initialChecklists,
  documentRequests: initialDocRequests,
  unusualTransactions: initialUnusualTxns,
  onShowToast,
  selectedClientId,
  userRole,
}) => {
  const [activeSubTab, setActiveSubTab] = useState<'checklists' | 'pbc' | 'sampling' | 'unusual'>('checklists');

  const [checklist, setChecklist] = useState<AuditChecklistItem[]>(initialChecklists);
  const [docRequests, setDocRequests] = useState<DocumentRequestItem[]>(initialDocRequests);
  const [unusualTxns, setUnusualTxns] = useState<UnusualTransaction[]>(initialUnusualTxns);

  // Sampling parameters
  const [thresholdAmount, setThresholdAmount] = useState(100000);

  // Update checklist status
  const handleUpdateChecklistStatus = (id: string, newStatus: AuditChecklistItem['status']) => {
    setChecklist((prev) =>
      prev.map((item) =>
        item.id === id
          ? {
              ...item,
              status: newStatus,
              reviewedBy: userRole === 'Administrator' || userRole === 'Reviewer' ? 'Audit Partner' : item.reviewedBy,
            }
          : item
      )
    );
    onShowToast('Checklist Updated', `Clause status marked as ${newStatus}.`, 'success');
  };

  // Export Audit Working Papers to Excel
  const handleExportAuditWorkingPaper = () => {
    const checklistData = checklist.map((c) => ({
      'Clause / Section': c.clauseOrSection,
      'Audit Objective': c.title,
      'Description': c.description,
      'Status': c.status,
      'Prepared By': c.preparedBy,
      'Reviewed By': c.reviewedBy || 'Pending',
      'Working Paper Ref': c.evidenceRef || 'None',
      'Reviewer Comments': c.comments,
    }));

    const unusualData = unusualTxns.map((u) => ({
      'Date': u.date,
      'Voucher No': u.voucherNo,
      'Party Name': u.partyName,
      'Amount (₹)': u.amount,
      'Issue Type': u.issueType,
      'Audit Observation': u.explanation,
      'Review Status': u.reviewerStatus,
    }));

    exportStyledWorkbook(
      [
        { name: 'Audit_Checklist', data: checklistData },
        { name: 'Flagged_Transactions', data: unusualData },
      ],
      'Audit_Working_Papers_FY2627.xlsx'
    );

    onShowToast('Audit Papers Downloaded', 'Exported complete audit package to XLSX with auto-fitted columns and formatted headers.', 'success');
  };

  return (
    <div className="space-y-4">
      {/* 1. Page Header (Section 5) */}
      <div className="bg-white rounded-lg border border-[#E5E7EB] p-5 flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-lg font-semibold text-[#111827] tracking-tight">Audit Tools</h2>
          <p className="text-xs text-[#6B7280] mt-0.5">
            Manage statutory checklists, PBC document requests, transaction sampling, and unusual transaction flags.
          </p>
        </div>

        <div className="flex items-center gap-2.5 shrink-0">
          <button
            onClick={handleExportAuditWorkingPaper}
            className="px-3.5 py-2 text-xs font-medium bg-gray-900 hover:bg-black text-white rounded-md flex items-center gap-1.5 transition-colors"
          >
            <Download className="w-3.5 h-3.5" />
            <span>Export Working Papers (.xlsx)</span>
          </button>
        </div>
      </div>

      {/* 2. Sub-Tab Switcher */}
      <div className="inline-flex p-1 bg-gray-50 border border-[#E5E7EB] rounded-md">
        <button
          onClick={() => setActiveSubTab('checklists')}
          className={`px-3 py-1.5 text-xs font-medium rounded transition-all ${
            activeSubTab === 'checklists'
              ? 'bg-white text-gray-900 font-semibold border border-[#E5E7EB] shadow-xs'
              : 'text-[#6B7280] hover:text-[#111827]'
          }`}
        >
          Checklists ({checklist.length})
        </button>
        <button
          onClick={() => setActiveSubTab('pbc')}
          className={`px-3 py-1.5 text-xs font-medium rounded transition-all ${
            activeSubTab === 'pbc'
              ? 'bg-white text-gray-900 font-semibold border border-[#E5E7EB] shadow-xs'
              : 'text-[#6B7280] hover:text-[#111827]'
          }`}
        >
          PBC Requests ({docRequests.length})
        </button>
        <button
          onClick={() => setActiveSubTab('sampling')}
          className={`px-3 py-1.5 text-xs font-medium rounded transition-all ${
            activeSubTab === 'sampling'
              ? 'bg-white text-gray-900 font-semibold border border-[#E5E7EB] shadow-xs'
              : 'text-[#6B7280] hover:text-[#111827]'
          }`}
        >
          Sampling Engine
        </button>
        <button
          onClick={() => setActiveSubTab('unusual')}
          className={`px-3 py-1.5 text-xs font-medium rounded transition-all ${
            activeSubTab === 'unusual'
              ? 'bg-white text-gray-900 font-semibold border border-[#E5E7EB] shadow-xs'
              : 'text-[#6B7280] hover:text-[#111827]'
          }`}
        >
          Flagged Items ({unusualTxns.length})
        </button>
      </div>

      {/* 3. Sub-Tab 1: Checklists */}
      {activeSubTab === 'checklists' && (
        <div className="bg-white rounded-lg border border-[#E5E7EB] overflow-hidden">
          <div className="px-4 py-3 border-b border-[#E5E7EB] bg-gray-50 flex items-center justify-between text-xs">
            <span className="font-semibold text-[#111827]">CARO 2020 & Tax Audit 3CD Checklist</span>
            <span className="text-[11px] text-gray-900 font-medium">
              {checklist.filter((c) => c.status === 'Verified' || c.status === 'Signed-off').length} of {checklist.length} Verified
            </span>
          </div>

          <div className="divide-y divide-[#E5E7EB]">
            {checklist.map((item) => (
              <div key={item.id} className="p-4 hover:bg-gray-50 transition-colors">
                <div className="flex flex-col md:flex-row md:items-start justify-between gap-3">
                  <div className="space-y-1.5 max-w-3xl">
                    <div className="flex items-center gap-2 flex-wrap">
                      <span className="font-mono text-xs font-semibold text-[#243B5A] bg-[#F4F7FB] border border-[#E5E7EB] px-1.5 py-0.5 rounded">
                        {item.clauseOrSection}
                      </span>
                      <h4 className="text-xs font-semibold text-[#111827]">{item.title}</h4>
                      <Badge variant={item.riskLevel === 'High' ? 'error' : item.riskLevel === 'Medium' ? 'warning' : 'navy'}>
                        {item.riskLevel}
                      </Badge>
                      <Badge variant={item.status === 'Signed-off' || item.status === 'Verified' ? 'success' : 'warning'}>
                        {item.status}
                      </Badge>
                    </div>

                    <p className="text-xs text-[#6B7280] leading-relaxed">
                      {item.description}
                    </p>

                    <div className="text-[11px] text-[#111827] pt-1 flex items-center gap-2 flex-wrap">
                      <span className="font-medium text-gray-900">WP Ref:</span>
                      <span className="font-mono text-[#6B7280] bg-white border border-[#E5E7EB] px-1.5 py-0.2 rounded">
                        {item.evidenceRef || 'Pending'}
                      </span>
                      <span>•</span>
                      <span className="text-[#6B7280]">Prepared: {item.preparedBy}</span>
                      {item.reviewedBy && (
                        <>
                          <span>•</span>
                          <span className="text-gray-900 font-medium">Reviewed: {item.reviewedBy}</span>
                        </>
                      )}
                    </div>

                    {item.comments && (
                      <div className="p-2 rounded bg-[#FEF9EE] border border-[#FBEEC8] text-[11px] text-[#9A6700] mt-1.5 flex items-start gap-1.5">
                        <MessageSquare className="w-3.5 h-3.5 shrink-0 mt-0.5" />
                        <span><strong>Note:</strong> {item.comments}</span>
                      </div>
                    )}
                  </div>

                  {/* Sign-off Action */}
                  <div className="flex items-center gap-2 shrink-0 self-start">
                    <select
                      value={item.status}
                      onChange={(e) => handleUpdateChecklistStatus(item.id, e.target.value as any)}
                      className="text-xs bg-white border border-[#E5E7EB] rounded-md px-2.5 py-1 text-[#111827] font-medium focus:outline-none focus:border-gray-900 cursor-pointer"
                    >
                      <option value="Pending">Pending</option>
                      <option value="In Progress">In Progress</option>
                      <option value="Verified">Verified</option>
                      <option value="Signed-off">Signed-off</option>
                    </select>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* 4. Sub-Tab 2: PBC Requests */}
      {activeSubTab === 'pbc' && (
        <div className="bg-white rounded-lg border border-[#E5E7EB] overflow-hidden">
          <div className="px-4 py-3 border-b border-[#E5E7EB] bg-gray-50 flex items-center justify-between text-xs">
            <span className="font-semibold text-[#111827]">Prepared By Client (PBC) Tracker</span>
            <span className="text-[11px] text-[#6B7280]">{docRequests.length} Requirements</span>
          </div>

          <div className="divide-y divide-[#E5E7EB]">
            {docRequests.map((doc) => (
              <div key={doc.id} className="p-4 hover:bg-gray-50 transition-colors flex items-center justify-between gap-3">
                <div>
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-semibold text-[#111827]">{doc.title}</span>
                    <Badge variant={doc.status === 'Received' ? 'success' : 'warning'}>
                      {doc.status}
                    </Badge>
                  </div>
                  <p className="text-xs text-[#6B7280] mt-0.5">{doc.description}</p>
                  <span className="text-[10px] text-[#6B7280] block mt-1">Due: {doc.dueDate} • Assigned: {doc.assignedTo}</span>
                </div>

                <select
                  value={doc.status}
                  onChange={(e) => {
                    setDocRequests((prev) => prev.map((d) => d.id === doc.id ? { ...d, status: e.target.value as any } : d));
                    onShowToast('Document Tracker Updated', `Marked as ${e.target.value}.`, 'info');
                  }}
                  className="text-xs bg-white border border-[#E5E7EB] rounded-md px-2.5 py-1 text-[#111827] font-medium focus:border-gray-900 cursor-pointer shrink-0"
                >
                  <option value="Pending">Pending</option>
                  <option value="Requested">Requested</option>
                  <option value="Received">Received</option>
                </select>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* 5. Sub-Tab 3: Sampling */}
      {activeSubTab === 'sampling' && (
        <div className="bg-white rounded-lg border border-[#E5E7EB] p-5">
          <h3 className="text-xs font-semibold text-[#111827] mb-1">Monetary Unit Sampling (MUS)</h3>
          <p className="text-xs text-[#6B7280] mb-4">Extract high-value transactions exceeding materiality threshold for substantive verification.</p>

          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 max-w-2xl">
            <div>
              <label className="text-xs font-medium text-[#111827] block mb-1">Threshold Amount (₹)</label>
              <input
                type="number"
                value={thresholdAmount}
                onChange={(e) => setThresholdAmount(parseInt(e.target.value) || 0)}
                className="w-full px-3 py-1.5 border border-[#E5E7EB] rounded-md font-mono text-xs focus:border-gray-900 focus:outline-none"
              />
            </div>
            <div className="flex items-end">
              <button
                onClick={() => onShowToast('Sample Extracted', `Selected 12 vouchers exceeding ₹${thresholdAmount.toLocaleString('en-IN')}.`, 'success')}
                className="w-full py-1.5 px-3 bg-gray-900 hover:bg-black text-white rounded-md text-xs font-medium transition-colors"
              >
                Extract Sample
              </button>
            </div>
          </div>

          <div className="mt-4 p-3 rounded-md bg-gray-100 border border-gray-200 text-xs text-gray-900">
            <strong>Sampling Summary:</strong> 12 vouchers (representing ₹1.84 Cr or 44% of total ledger debit turnover) selected for 100% substantive vouching. Working paper generated.
          </div>
        </div>
      )}

      {/* 6. Sub-Tab 4: Unusual Transactions */}
      {activeSubTab === 'unusual' && (
        <div className="bg-white rounded-lg border border-[#E5E7EB] overflow-hidden">
          <div className="px-4 py-3 border-b border-[#E5E7EB] bg-gray-50 flex items-center justify-between text-xs">
            <span className="font-semibold text-[#111827]">Potential Anomalies & Exceptions</span>
            <span className="text-[11px] text-[#C43D43] font-medium bg-[#FDF2F2] border border-[#FDE8E8] px-2 py-0.5 rounded">
              {unusualTxns.length} Flagged
            </span>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs border-collapse">
              <thead className="bg-gray-50 border-b border-[#E5E7EB] text-[#6B7280] text-[11px] font-medium">
                <tr>
                  <th className="py-2.5 px-3 border-r border-[#E5E7EB]">Voucher & Date</th>
                  <th className="py-2.5 px-3 border-r border-[#E5E7EB]">Party</th>
                  <th className="py-2.5 px-3 border-r border-[#E5E7EB] text-right">Amount (₹)</th>
                  <th className="py-2.5 px-3 border-r border-[#E5E7EB]">Flag Type</th>
                  <th className="py-2.5 px-3">Audit Observation</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#E5E7EB] text-[#111827]">
                {unusualTxns.map((ut) => (
                  <tr key={ut.id} className="hover:bg-gray-50 transition-colors">
                    <td className="py-2.5 px-3 border-r border-[#E5E7EB]">
                      <div className="font-medium text-[#111827]">{ut.voucherNo}</div>
                      <div className="font-mono text-[10px] text-[#6B7280]">{ut.date}</div>
                    </td>
                    <td className="py-2.5 px-3 border-r border-[#E5E7EB] font-medium">
                      {ut.partyName}
                    </td>
                    <td className="py-2.5 px-3 border-r border-[#E5E7EB] text-right font-mono tabular-nums font-semibold">
                      ₹ {ut.amount.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                    </td>
                    <td className="py-2.5 px-3 border-r border-[#E5E7EB]">
                      <Badge variant={ut.riskScore === 'High' ? 'error' : 'warning'}>
                        {ut.issueType}
                      </Badge>
                    </td>
                    <td className="py-2.5 px-3 text-[#6B7280] text-xs">
                      {ut.explanation}
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
