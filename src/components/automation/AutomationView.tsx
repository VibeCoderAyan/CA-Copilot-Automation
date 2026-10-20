import React, { useState } from 'react';
import type { InvoiceExtracted } from '../../types';
import { ExcelAutomationView } from '../excel/ExcelAutomationView';
import { PdfToExcelView } from '../pdf/PdfToExcelView';
import { InvoiceProcessingView } from '../invoice/InvoiceProcessingView';
import { WorkflowOrchestratorView } from './WorkflowOrchestratorView';
import { AdvancedModulesHub } from './AdvancedModulesHub';
import { FileSpreadsheet, FileCode, ScanLine, GitMerge, Sparkles } from 'lucide-react';

interface AutomationViewProps {
  initialSubTab?: 'advanced-modules' | 'orchestrator' | 'excel' | 'pdf-excel' | 'invoices';
  invoices: InvoiceExtracted[];
  selectedClientId: string;
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
}

export const AutomationView: React.FC<AutomationViewProps> = ({
  initialSubTab = 'advanced-modules',
  invoices,
  selectedClientId,
  onShowToast,
}) => {
  const [activeSubTab, setActiveSubTab] = useState<'advanced-modules' | 'orchestrator' | 'excel' | 'pdf-excel' | 'invoices'>(initialSubTab);

  const tabs: { id: 'advanced-modules' | 'orchestrator' | 'excel' | 'pdf-excel' | 'invoices'; label: string; icon: React.ComponentType<{ className?: string }> }[] = [
    { id: 'advanced-modules', label: '7-Module Firm Suite', icon: Sparkles },
    { id: 'orchestrator', label: 'Workflow Engine', icon: GitMerge },
    { id: 'excel', label: 'Excel Cleaner', icon: FileSpreadsheet },
    { id: 'pdf-excel', label: 'PDF to Excel', icon: FileCode },
    { id: 'invoices', label: 'Invoice Processing', icon: ScanLine },
  ];

  return (
    <div className="space-y-4">
      {/* Page Header & Segmented Sub-Nav */}
      <div className="bg-white rounded-lg border border-[#E5E7EB] px-5 py-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <h1 className="text-base font-semibold text-[#111827] tracking-tight">Automation Engine & Workflows</h1>
          <p className="text-xs text-[#6B7280] mt-0.5">End-to-end accounting automation workflows, spreadsheet cleaning, and OCR invoice digitization.</p>
        </div>

        {/* Segmented Control */}
        <div className="inline-flex p-1 bg-gray-50 border border-[#E5E7EB] rounded-md shrink-0">
          {tabs.map((tab) => {
            const Icon = tab.icon;
            const isActive = activeSubTab === tab.id;
            return (
              <button
                key={tab.id}
                onClick={() => setActiveSubTab(tab.id)}
                className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded transition-all cursor-pointer ${
                  isActive
                    ? 'bg-gray-900 text-white font-semibold shadow-xs'
                    : 'text-[#6B7280] hover:text-[#111827]'
                }`}
              >
                <Icon className={`w-3.5 h-3.5 ${isActive ? 'text-white' : 'text-[#6B7280]'}`} />
                <span>{tab.label}</span>
              </button>
            );
          })}
        </div>
      </div>

      {/* Active Automation Workspace */}
      <div>
        {activeSubTab === 'advanced-modules' && (
          <AdvancedModulesHub
            selectedClientId={selectedClientId}
            onShowToast={onShowToast}
          />
        )}

        {activeSubTab === 'orchestrator' && (
          <WorkflowOrchestratorView
            selectedClientId={selectedClientId}
            onShowToast={onShowToast}
          />
        )}

        {activeSubTab === 'excel' && (
          <ExcelAutomationView
            onShowToast={onShowToast}
            selectedClientId={selectedClientId}
          />
        )}

        {activeSubTab === 'pdf-excel' && (
          <PdfToExcelView
            onShowToast={onShowToast}
            selectedClientId={selectedClientId}
          />
        )}

        {activeSubTab === 'invoices' && (
          <InvoiceProcessingView
            invoices={invoices}
            onShowToast={onShowToast}
            selectedClientId={selectedClientId}
          />
        )}
      </div>
    </div>
  );
};

