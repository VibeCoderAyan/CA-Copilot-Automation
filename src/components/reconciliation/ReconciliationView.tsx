import React, { useState } from 'react';
import type { BankTransaction, LedgerTransaction, GstRecord } from '../../types';
import { BankReconciliationView } from '../bank/BankReconciliationView';
import { GstReconciliationView } from '../gst/GstReconciliationView';
import { ArrowLeftRight, ReceiptText } from 'lucide-react';

interface ReconciliationViewProps {
  initialSubTab?: 'bank-rec' | 'gst-rec';
  bankTransactions: BankTransaction[];
  ledgerTransactions: LedgerTransaction[];
  gstRecords: GstRecord[];
  selectedClientId: string;
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
}

export const ReconciliationView: React.FC<ReconciliationViewProps> = ({
  initialSubTab = 'bank-rec',
  bankTransactions,
  ledgerTransactions,
  gstRecords,
  selectedClientId,
  onShowToast,
}) => {
  const [activeSubTab, setActiveSubTab] = useState<'bank-rec' | 'gst-rec'>(initialSubTab);

  const tabs: { id: 'bank-rec' | 'gst-rec'; label: string; icon: React.ComponentType<{ className?: string }> }[] = [
    { id: 'bank-rec', label: 'Bank Reconciliation', icon: ArrowLeftRight },
    { id: 'gst-rec', label: 'GST Reconciliation', icon: ReceiptText },
  ];

  return (
    <div className="space-y-4">
      {/* Header & Sub-Tab Switcher */}
      <div className="bg-white rounded-lg border border-[#E5E7EB] px-5 py-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <h1 className="text-base font-semibold text-[#111827] tracking-tight">Reconciliation Workspace</h1>
          <p className="text-xs text-[#6B7280] mt-0.5">Automated BRS ledger matching and GSTR-2B purchase register comparison.</p>
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
                className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded transition-all ${
                  isActive
                    ? 'bg-white text-gray-900 font-semibold shadow-xs border border-[#E5E7EB]'
                    : 'text-[#6B7280] hover:text-[#111827]'
                }`}
              >
                <Icon className={`w-3.5 h-3.5 ${isActive ? 'text-gray-900' : 'text-[#6B7280]'}`} />
                <span>{tab.label}</span>
              </button>
            );
          })}
        </div>
      </div>

      {/* Active Reconciliation Workspace */}
      <div>
        {activeSubTab === 'bank-rec' && (
          <BankReconciliationView
            bankTransactions={bankTransactions}
            ledgerTransactions={ledgerTransactions}
            selectedClientId={selectedClientId}
            onShowToast={onShowToast}
          />
        )}

        {activeSubTab === 'gst-rec' && (
          <GstReconciliationView
            gstRecords={gstRecords}
            selectedClientId={selectedClientId}
            onShowToast={onShowToast}
          />
        )}
      </div>
    </div>
  );
};
