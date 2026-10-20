import React, { useState } from 'react';
import type { ActiveTab } from '../../types';
import {
  FileSpreadsheet,
  ArrowLeftRight,
  ReceiptText,
  FileCode,
  ScanLine,
  ShieldCheck,
  FileBarChart,
  ArrowRight,
  Plus,
  X
} from 'lucide-react';

interface OverviewViewProps {
  onNavigateTab: (tab: ActiveTab) => void;
  onShowToast?: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
}

export const OverviewView: React.FC<OverviewViewProps> = ({
  onNavigateTab,
  onShowToast,
}) => {
  const [isNewTaskModalOpen, setIsNewTaskModalOpen] = useState(false);

  // The 7 Primary Automation Tools specified in the prompt
  const tools: {
    id: ActiveTab;
    name: string;
    description: string;
    icon: React.ComponentType<{ className?: string; strokeWidth?: number }>;
  }[] = [
    {
      id: 'excel',
      name: 'Excel Cleaner',
      description: 'Clean and standardize spreadsheets.',
      icon: FileSpreadsheet,
    },
    {
      id: 'bank-rec',
      name: 'Bank Reconciliation',
      description: 'Match bank transactions with ledger entries.',
      icon: ArrowLeftRight,
    },
    {
      id: 'gst-rec',
      name: 'GST Reconciliation',
      description: 'Compare purchase registers with GSTR-2B data.',
      icon: ReceiptText,
    },
    {
      id: 'pdf-excel',
      name: 'PDF to Excel',
      description: 'Extract structured data from supported PDFs.',
      icon: FileCode,
    },
    {
      id: 'invoices',
      name: 'Invoice Processing',
      description: 'Extract and validate invoice information.',
      icon: ScanLine,
    },
    {
      id: 'audit',
      name: 'Audit Tools',
      description: 'Manage checks and working papers.',
      icon: ShieldCheck,
    },
    {
      id: 'reports',
      name: 'Financial Reports',
      description: 'Prepare and export financial working papers.',
      icon: FileBarChart,
    },
  ];

  // Compact Recent Activity records
  const recentActivities = [
    {
      file: 'Tally_Raw_Sales_Ledger_Extract_Aug26.xlsx',
      task: 'Excel Cleaner',
      status: 'Cleaned',
      updated: 'Today, 11:20 AM',
      tab: 'excel' as ActiveTab,
    },
    {
      file: 'HDFC_Bank_Current_Account_Sept2026.pdf',
      task: 'Bank Reconciliation',
      status: 'Reconciled',
      updated: 'Today, 10:45 AM',
      tab: 'bank-rec' as ActiveTab,
    },
    {
      file: 'GSTR2B_Sept2026_Vs_PurchaseBooks.xlsx',
      task: 'GST Reconciliation',
      status: 'Exceptions (1)',
      updated: 'Yesterday, 04:15 PM',
      tab: 'gst-rec' as ActiveTab,
    },
    {
      file: 'Vendor_Invoices_Batch_Sept2026.pdf',
      task: 'Invoice Processing',
      status: 'Extracted',
      updated: '14 Oct 2026',
      tab: 'invoices' as ActiveTab,
    },
    {
      file: 'Statutory_Audit_Sampling_FY26.xlsx',
      task: 'Audit Tools',
      status: 'Sampled',
      updated: '12 Oct 2026',
      tab: 'audit' as ActiveTab,
    },
  ];

  return (
    <div className="space-y-6">
      {/* 1. Top Section */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-semibold text-[#111827] tracking-tight">Overview</h1>
          <p className="text-xs text-[#6B7280] mt-0.5">Your accounting workspace.</p>
        </div>

        <button
          onClick={() => setIsNewTaskModalOpen(true)}
          className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium text-white bg-gray-900 hover:bg-black rounded-md transition-colors"
        >
          <Plus className="w-3.5 h-3.5" />
          <span>New Task</span>
        </button>
      </div>

      {/* 2. Main Section: Automation Tools Grid */}
      <div>
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3.5">
          {tools.map((tool) => {
            const Icon = tool.icon;
            return (
              <div
                key={tool.id}
                onClick={() => onNavigateTab(tool.id)}
                className="bg-white border border-[#E5E7EB] rounded-lg p-4 transition-all duration-150 cursor-pointer group hover:border-gray-900 flex flex-col justify-between h-[130px]"
              >
                <div className="space-y-2">
                  <div className="flex items-center justify-between">
                    <div className="w-7 h-7 rounded bg-gray-50 border border-[#E5E7EB] flex items-center justify-center text-gray-900">
                      <Icon className="w-3.5 h-3.5" strokeWidth={2} />
                    </div>
                    <ArrowRight className="w-3.5 h-3.5 text-[#6B7280] group-hover:text-gray-900 group-hover:translate-x-0.5 transition-all" />
                  </div>
                  <div>
                    <h3 className="text-xs font-semibold text-[#111827] group-hover:text-gray-900 transition-colors">
                      {tool.name}
                    </h3>
                    <p className="text-xs text-[#6B7280] mt-0.5 leading-relaxed">
                      {tool.description}
                    </p>
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      </div>

      {/* 3. Secondary Section: Compact Recent Activity Table */}
      <div className="bg-white border border-[#E5E7EB] rounded-lg overflow-hidden">
        <div className="px-4 py-3 border-b border-[#E5E7EB] flex items-center justify-between">
          <h2 className="text-xs font-semibold text-[#111827]">Recent Activity</h2>
          <span className="text-[11px] text-[#6B7280]">5 most recent files</span>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs">
            <thead>
              <tr className="bg-gray-50 border-b border-[#E5E7EB] text-[#6B7280] text-[11px] font-medium">
                <th className="py-2.5 px-4 font-medium">File</th>
                <th className="py-2.5 px-4 font-medium">Task</th>
                <th className="py-2.5 px-4 font-medium">Status</th>
                <th className="py-2.5 px-4 font-medium text-right">Updated</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-[#E5E7EB] text-[#111827]">
              {recentActivities.map((act, index) => (
                <tr
                  key={index}
                  onClick={() => onNavigateTab(act.tab)}
                  className="hover:bg-gray-50 transition-colors cursor-pointer group"
                >
                  <td className="py-2.5 px-4">
                    <span className="font-medium text-[#111827] group-hover:text-gray-900 transition-colors">
                      {act.file}
                    </span>
                  </td>
                  <td className="py-2.5 px-4 text-[#6B7280]">
                    {act.task}
                  </td>
                  <td className="py-2.5 px-4">
                    <span className={`inline-flex items-center px-2 py-0.5 rounded text-[11px] font-medium ${
                      act.status.includes('Exceptions')
                        ? 'bg-[#FDF2F2] text-[#C43D43] border border-[#FDE8E8]'
                        : 'bg-gray-100 text-gray-900 border border-gray-200'
                    }`}>
                      {act.status}
                    </span>
                  </td>
                  <td className="py-2.5 px-4 text-right text-[#6B7280] font-mono text-[11px]">
                    {act.updated}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* New Task Selection Modal */}
      {isNewTaskModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/30 backdrop-blur-xs p-4">
          <div className="bg-white rounded-lg border border-[#E5E7EB] max-w-md w-full p-5 space-y-4 shadow-lg">
            <div className="flex items-center justify-between pb-3 border-b border-[#E5E7EB]">
              <div>
                <h3 className="text-sm font-semibold text-[#111827]">Select Tool for New Task</h3>
                <p className="text-xs text-[#6B7280] mt-0.5">Choose an accounting automation to start work.</p>
              </div>
              <button
                onClick={() => setIsNewTaskModalOpen(false)}
                className="text-[#6B7280] hover:text-[#111827] p-1 rounded"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="space-y-1.5 max-h-80 overflow-y-auto">
              {tools.map((t) => {
                const Icon = t.icon;
                return (
                  <button
                    key={t.id}
                    onClick={() => {
                      setIsNewTaskModalOpen(false);
                      onNavigateTab(t.id);
                    }}
                    className="w-full flex items-center justify-between p-2.5 rounded-md hover:bg-gray-50 border border-transparent hover:border-[#E5E7EB] transition-colors text-left"
                  >
                    <div className="flex items-center gap-2.5">
                      <div className="w-7 h-7 rounded bg-gray-50 border border-[#E5E7EB] flex items-center justify-center text-gray-900 shrink-0">
                        <Icon className="w-3.5 h-3.5" />
                      </div>
                      <div>
                        <span className="text-xs font-semibold text-[#111827] block">{t.name}</span>
                        <span className="text-[11px] text-[#6B7280] block">{t.description}</span>
                      </div>
                    </div>
                    <ArrowRight className="w-3.5 h-3.5 text-[#6B7280]" />
                  </button>
                );
              })}
            </div>

            <div className="pt-2 border-t border-[#E5E7EB] flex justify-end">
              <button
                onClick={() => setIsNewTaskModalOpen(false)}
                className="px-3 py-1.5 text-xs text-[#6B7280] hover:text-[#111827] rounded border border-[#E5E7EB] hover:bg-gray-50"
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
