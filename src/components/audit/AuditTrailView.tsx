import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import type { AuditLog } from '../../types';
import { Badge } from '../common/Badge';
import {
  ShieldCheck,
  Search,
  Download,
  Filter,
  CheckCircle2,
  Lock,
  Calendar,
  User,
  KeyRound
} from 'lucide-react';

import { exportStyledWorkbook } from '../../utils/excelStyler';

interface AuditTrailViewProps {
  logs: AuditLog[];
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
  selectedClientId?: string;
}

export const AuditTrailView: React.FC<AuditTrailViewProps> = ({
  logs,
  onShowToast,
  selectedClientId: _selectedClientId,
}) => {
  const [searchQuery, setSearchQuery] = useState('');
  const [moduleFilter, setModuleFilter] = useState('All');

  const filteredLogs = logs.filter((log) => {
    const matchesModule = moduleFilter === 'All' || log.module === moduleFilter;
    const matchesSearch =
      log.user.toLowerCase().includes(searchQuery.toLowerCase()) ||
      log.action.toLowerCase().includes(searchQuery.toLowerCase()) ||
      log.details.toLowerCase().includes(searchQuery.toLowerCase()) ||
      log.clientName.toLowerCase().includes(searchQuery.toLowerCase());

    return matchesModule && matchesSearch;
  });

  const handleExportAuditLog = () => {
    const exportData = filteredLogs.map((l) => ({
      'Log ID': l.id,
      'Timestamp (IST)': l.timestamp,
      'CA User': l.user,
      'User Role': l.userRole,
      'Action': l.action,
      'Module': l.module,
      'Client': l.clientName,
      'Audit Details': l.details,
      'IP Address': l.ipAddress,
      'SHA256 Hash': l.hash,
    }));

    exportStyledWorkbook(
      [
        {
          name: 'System_Audit_Trail',
          data: exportData,
          columnWidths: {
            'Log ID': 14,
            'Timestamp (IST)': 22,
            'CA User': 20,
            'User Role': 16,
            'Action': 22,
            'Module': 24,
            'Client': 26,
            'Audit Details': 45,
            'IP Address': 18,
            'SHA256 Hash': 34,
          },
        },
      ],
      'LedgerFlow_Immutable_Audit_Trail.xlsx'
    );

    onShowToast('Audit Log Exported', 'Downloaded regulatory audit record in formatted Excel.', 'success');
  };

  return (
    <div className="space-y-6">
      {/* Header Card */}
      <div className="bg-white rounded border border-gray-200 p-5">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-gray-200">
          <div>
            <div className="flex items-center gap-2">
              <ShieldCheck className="w-4 h-4 text-gray-900" />
              <h3 className="text-sm font-semibold text-[#111827]">Regulatory Immutable Audit Trail</h3>
              <Badge variant="success">Tamper-Proof SHA-256</Badge>
              <Badge variant="navy">ICAI SQC 1 Compliant</Badge>
            </div>
            <p className="text-xs text-gray-600 mt-1">
              Cryptographically verified timestamped event log for peer reviews, statutory scrutiny, and quality control
            </p>
          </div>

          <button
            onClick={handleExportAuditLog}
            className="flex items-center gap-1.5 px-3.5 py-1.5 bg-gray-900 hover:bg-black text-white rounded text-xs font-medium transition-colors shadow-sm self-start sm:self-auto"
          >
            <Download className="w-3.5 h-3.5" />
            <span>Export Audit Trail (.xlsx)</span>
          </button>
        </div>

        {/* Filter and Search Bar */}
        <div className="mt-4 flex flex-col md:flex-row md:items-center justify-between gap-3">
          <div className="flex items-center gap-2 overflow-x-auto pb-1 md:pb-0">
            {['All', 'Bank Reconciliation', 'GST Reconciliation', 'Excel Automation', 'Invoice Processing', 'Financial Reports'].map((mod) => (
              <button
                key={mod}
                onClick={() => setModuleFilter(mod)}
                className={`px-3 py-1.5 text-xs font-medium rounded border transition-colors shrink-0 ${
                  moduleFilter === mod
                    ? 'bg-gray-100 text-gray-900 border-gray-200'
                    : 'bg-white text-gray-600 border-gray-200 hover:bg-gray-50'
                }`}
              >
                {mod}
              </button>
            ))}
          </div>

          <div className="relative w-full md:w-64">
            <Search className="w-3.5 h-3.5 text-gray-600 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="Search user, action, hash..."
              className="w-full pl-8 pr-3 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs text-[#111827] placeholder-[#4B5563] focus:outline-none focus:bg-white focus:border-gray-300"
            />
          </div>
        </div>
      </div>

      {/* Audit Logs Table */}
      <div className="bg-white rounded border border-gray-200 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs border-collapse">
            <thead className="bg-gray-50 border-b border-gray-200 text-gray-600 font-medium">
              <tr>
                <th className="py-2.5 px-3 border-r border-gray-200 w-36">Timestamp (IST)</th>
                <th className="py-2.5 px-3 border-r border-gray-200">CA User & Role</th>
                <th className="py-2.5 px-3 border-r border-gray-200">Action & Module</th>
                <th className="py-2.5 px-3 border-r border-gray-200">Client Scope</th>
                <th className="py-2.5 px-3 border-r border-gray-200">Event Description</th>
                <th className="py-2.5 px-3 text-right">Integrity Hash</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-[#E5E7EB]">
              {filteredLogs.map((log) => (
                <tr key={log.id} className="hover:bg-gray-50 transition-colors">
                  <td className="py-3 px-3 border-r border-gray-200 font-mono text-[11px] text-gray-600">
                    {log.timestamp}
                  </td>
                  <td className="py-3 px-3 border-r border-gray-200">
                    <div className="font-medium text-[#111827]">{log.user}</div>
                    <div className="text-[10px] text-gray-600">{log.userRole}</div>
                  </td>
                  <td className="py-3 px-3 border-r border-gray-200">
                    <div className="font-semibold text-[#111827]">{log.action}</div>
                    <span className="text-[10px] px-1.5 py-0.2 rounded bg-gray-50 border border-gray-200 text-[#203B63] font-medium mt-0.5 inline-block">
                      {log.module}
                    </span>
                  </td>
                  <td className="py-3 px-3 border-r border-gray-200 font-medium text-[#111827]">
                    {log.clientName}
                  </td>
                  <td className="py-3 px-3 border-r border-gray-200 text-gray-600 text-[11px] max-w-xs">
                    {log.details}
                    <div className="text-[10px] text-[#8D9B93] mt-0.5 font-mono">IP: {log.ipAddress}</div>
                  </td>
                  <td className="py-3 px-3 text-right font-mono text-[10px] text-gray-900">
                    <div className="flex items-center justify-end gap-1">
                      <Lock className="w-3 h-3 text-gray-900" />
                      <span title={log.hash}>{log.hash.slice(0, 16)}...</span>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};
