import React, { useState } from 'react';
import type { ActiveTab, Client } from '../../types';
import {
  X,
  ArrowLeftRight,
  ReceiptText,
  FileSpreadsheet,
  ScanLine,
  FileBarChart,
  ArrowRight
} from 'lucide-react';

interface NewAutomationModalProps {
  isOpen: boolean;
  onClose: () => void;
  clients: Client[];
  selectedClientId: string;
  onStartAutomation: (module: ActiveTab, clientId: string) => void;
}

export const NewAutomationModal: React.FC<NewAutomationModalProps> = ({
  isOpen,
  onClose,
  clients,
  selectedClientId,
  onStartAutomation,
}) => {
  const [selectedModule, setSelectedModule] = useState<ActiveTab>('bank-rec');
  const [targetClientId, setTargetClientId] = useState<string>(
    selectedClientId === 'all' ? clients[0]?.id || '' : selectedClientId
  );
  const [period, setPeriod] = useState('September 2026 (Q2)');

  if (!isOpen) return null;

  const modules: {
    id: ActiveTab;
    title: string;
    description: string;
    icon: React.ComponentType<{ className?: string }>;
    inputTypes: string;
  }[] = [
    {
      id: 'bank-rec',
      title: 'Bank Reconciliation',
      description: 'Auto-match bank statements with accounting ledgers (Tally, SAP, Zoho)',
      icon: ArrowLeftRight,
      inputTypes: 'XLSX, XLS, CSV, PDF statements',
    },
    {
      id: 'gst-rec',
      title: 'GST Reconciliation (2B vs Books)',
      description: 'Validate GSTR-2B ITC eligibility against Purchase Register with rule 37A / 16(2)(aa)',
      icon: ReceiptText,
      inputTypes: 'JSON from GST Portal, Tally Purchase Register',
    },
    {
      id: 'excel',
      title: 'Excel Data Cleaning & Normalization',
      description: 'Normalize date formats, clean negative bracket amounts, deduplicate rows, map columns',
      icon: FileSpreadsheet,
      inputTypes: 'Any spreadsheet (XLSX, CSV, TSV)',
    },
    {
      id: 'invoices',
      title: 'Invoice Document Extraction',
      description: 'Optical extraction of GSTIN, tax rates, HSN codes and amounts with confidence ratings',
      icon: ScanLine,
      inputTypes: 'Scanned PDFs, Images, Vendor bills',
    },
    {
      id: 'reports',
      title: 'Trial Balance & Working Papers',
      description: 'Debit/Credit balance verification, P&L schedules, Schedule III Balance Sheet preview',
      icon: FileBarChart,
      inputTypes: 'Accounting Trial Balance extract',
    },
  ];

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    onStartAutomation(selectedModule, targetClientId);
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-[#111827]/40 p-4">
      <div className="bg-white rounded border border-gray-200 shadow-modal w-full max-w-xl overflow-hidden animate-in fade-in zoom-in-95 duration-150">
        {/* Header */}
        <div className="px-5 py-4 border-b border-gray-200 flex items-center justify-between bg-white">
          <div>
            <h3 className="text-sm font-semibold text-[#111827]">Start New Automation Workflow</h3>
            <p className="text-xs text-gray-600 mt-0.5">
              Select the client and accounting automation module to execute
            </p>
          </div>
          <button
            onClick={onClose}
            className="text-gray-600 hover:text-[#111827] p-1 rounded hover:bg-gray-50"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="p-5 space-y-4">
          {/* Client & Period Selection */}
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="block text-xs font-medium text-[#111827] mb-1">
                Client Entity
              </label>
              <select
                value={targetClientId}
                onChange={(e) => setTargetClientId(e.target.value)}
                className="w-full text-xs bg-white border border-gray-200 rounded px-3 py-2 text-[#111827] focus:outline-none focus:border-gray-300"
              >
                {clients.map((c) => (
                  <option key={c.id} value={c.id}>
                    {c.name} ({c.gstin.slice(0, 5)}...)
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-xs font-medium text-[#111827] mb-1">
                Accounting Period
              </label>
              <select
                value={period}
                onChange={(e) => setPeriod(e.target.value)}
                className="w-full text-xs bg-white border border-gray-200 rounded px-3 py-2 text-[#111827] focus:outline-none focus:border-gray-300"
              >
                <option value="September 2026 (Q2)">September 2026 (Q2)</option>
                <option value="August 2026">August 2026</option>
                <option value="July 2026">July 2026</option>
                <option value="Q1 FY 2026-27">Q1 FY 2026-27</option>
                <option value="Full FY 2025-26">Full FY 2025-26 (Statutory Audit)</option>
              </select>
            </div>
          </div>

          {/* Module Selector */}
          <div>
            <label className="block text-xs font-medium text-[#111827] mb-2">
              Select Automation Task
            </label>
            <div className="space-y-2">
              {modules.map((mod) => {
                const Icon = mod.icon;
                const isSelected = selectedModule === mod.id;

                return (
                  <div
                    key={mod.id}
                    onClick={() => setSelectedModule(mod.id)}
                    className={`p-3 rounded border text-left cursor-pointer transition-colors flex items-start gap-3 ${
                      isSelected
                        ? 'border-gray-300 bg-gray-50 ring-1 ring-[#111827]'
                        : 'border-gray-200 hover:border-gray-300 bg-white'
                    }`}
                  >
                    <div
                      className={`p-2 rounded shrink-0 ${
                        isSelected ? 'bg-gray-900 text-white' : 'bg-gray-50 text-gray-600'
                      }`}
                    >
                      <Icon className="w-4 h-4" />
                    </div>
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center justify-between">
                        <span className="text-xs font-semibold text-[#111827]">{mod.title}</span>
                        <span className="text-[10px] text-gray-600">{mod.inputTypes}</span>
                      </div>
                      <p className="text-[11px] text-gray-600 mt-0.5 leading-snug">{mod.description}</p>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          {/* Footer Action */}
          <div className="pt-3 border-t border-gray-200 flex items-center justify-end gap-2">
            <button
              type="button"
              onClick={onClose}
              className="px-3.5 py-1.5 text-xs text-gray-600 hover:text-[#111827] rounded border border-gray-200 hover:bg-gray-50"
            >
              Cancel
            </button>
            <button
              type="submit"
              className="px-4 py-1.5 text-xs font-medium bg-gray-900 hover:bg-black text-white rounded flex items-center gap-1.5 transition-colors shadow-sm"
            >
              <span>Open Automation Workspace</span>
              <ArrowRight className="w-3.5 h-3.5" />
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};
