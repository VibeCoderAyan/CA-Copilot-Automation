import React from 'react';
import type { ActiveTab } from '../../types';
import { Search } from 'lucide-react';

interface NavbarProps {
  activeTab: ActiveTab;
  searchQuery: string;
  setSearchQuery: (query: string) => void;
  selectedClientId: string;
  setSelectedClientId: (id: string) => void;
}

export const Navbar: React.FC<NavbarProps> = ({
  activeTab,
  searchQuery,
  setSearchQuery,
  selectedClientId,
  setSelectedClientId,
}) => {
  const titles: Record<string, string> = {
    overview: 'Overview',
    dashboard: 'Overview',
    automation: 'Automation',
    reconciliation: 'Reconciliation',
    excel: 'Excel Cleaner',
    'bank-rec': 'Bank Reconciliation',
    'gst-rec': 'GST Reconciliation',
    'pdf-excel': 'PDF to Excel',
    invoices: 'Invoice Processing',
    audit: 'Audit Tools',
    reports: 'Financial Reports',
    clients: 'Clients',
    settings: 'Settings',
  };

  const clients = [
    { id: '', name: 'Select / Enter Client...' },
    { id: 'client_active', name: 'Active Client Audit' },
  ];

  return (
    <header className="h-[52px] bg-white border-b border-[#E5E7EB] px-6 flex items-center justify-between sticky top-0 z-10 shrink-0">
      {/* 1. Breadcrumbs */}
      <div className="flex items-center gap-2 text-xs">
        <span className="text-[#6B7280]">LedgerFlow</span>
        <span className="text-[#6B7280]">/</span>
        <span className="font-semibold text-[#111827]">
          {titles[activeTab] || 'Workspace'}
        </span>
      </div>

      {/* 2. Global Search Input */}
      <div className="flex-1 max-w-sm mx-6 hidden sm:block">
        <div className="relative w-full">
          <Search className="w-3.5 h-3.5 text-[#6B7280] absolute left-3 top-1/2 -translate-y-1/2" />
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="Search records, vouchers, files..."
            className="w-full pl-8 pr-3 py-1.5 bg-gray-50 border border-[#E5E7EB] rounded-md text-xs text-[#111827] placeholder-[#6B7280] focus:outline-none focus:bg-white focus:border-gray-900 transition-colors"
          />
        </div>
      </div>

      {/* 3. Client Selector & User Profile */}
      <div className="flex items-center gap-3">
        {/* Client Selector */}
        <div className="flex items-center gap-1.5">
          <label htmlFor="client-select" className="text-[11px] text-[#6B7280] hidden md:inline">Client:</label>
          <select
            id="client-select"
            value={selectedClientId}
            onChange={(e) => setSelectedClientId(e.target.value)}
            className="text-xs bg-gray-50 border border-[#E5E7EB] rounded-md px-2.5 py-1.5 font-medium text-[#111827] focus:outline-none focus:border-gray-900 cursor-pointer max-w-[190px] truncate"
          >
            {clients.map((c) => (
              <option key={c.id} value={c.id}>
                {c.name}
              </option>
            ))}
          </select>
        </div>

        {/* User Profile */}
        <div className="flex items-center gap-2 pl-3 border-l border-[#E5E7EB]">
          <div className="w-7 h-7 rounded-full bg-gray-900 text-white flex items-center justify-center text-[11px] font-semibold shrink-0">
            CA
          </div>
          <div className="hidden lg:block text-left">
            <span className="text-xs font-medium text-[#111827] block leading-tight">CA Office</span>
            <span className="text-[10px] text-[#6B7280] block leading-tight">Partner</span>
          </div>
        </div>
      </div>
    </header>
  );
};
