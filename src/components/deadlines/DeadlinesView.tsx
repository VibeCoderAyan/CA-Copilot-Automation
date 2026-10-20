import React, { useState } from 'react';
import type { StatutoryDeadline, Client } from '../../types';
import { Badge } from '../common/Badge';
import {
  CalendarCheck,
  Calendar,
  AlertOctagon,
  Clock,
  CheckCircle2,
  Filter,
  Check,
  Building,
  UserCheck
} from 'lucide-react';

interface DeadlinesViewProps {
  deadlines: StatutoryDeadline[];
  clients: Client[];
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
  selectedClientId: string;
}

export const DeadlinesView: React.FC<DeadlinesViewProps> = ({
  deadlines: initialDeadlines,
  clients,
  onShowToast,
  selectedClientId,
}) => {
  const [deadlinesList, setDeadlinesList] = useState<StatutoryDeadline[]>(initialDeadlines);
  const [activeActFilter, setActiveActFilter] = useState('All');

  const filteredDeadlines = deadlinesList.filter((dl) => {
    const matchesAct = activeActFilter === 'All' || dl.act === activeActFilter;
    const matchesClient =
      selectedClientId === 'all' || dl.applicableClients.includes(selectedClientId);
    return matchesAct && matchesClient;
  });

  const handleMarkCompleted = (id: string) => {
    setDeadlinesList((prev) =>
      prev.map((dl) => (dl.id === id ? { ...dl, status: 'Filed', daysRemaining: 0 } : dl))
    );
    onShowToast('Statutory Filing Marked as Complete', 'Updated compliance tracker and audit trail.', 'success');
  };

  return (
    <div className="space-y-6">
      {/* Header Card */}
      <div className="bg-white rounded border border-gray-200 p-5">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-gray-200">
          <div>
            <div className="flex items-center gap-2">
              <CalendarCheck className="w-4 h-4 text-gray-900" />
              <h3 className="text-sm font-semibold text-[#111827]">
                Statutory Compliance & Return Filing Tracker
              </h3>
              <Badge variant="navy">October - December 2026</Badge>
            </div>
            <p className="text-xs text-gray-600 mt-1">
              Statutory deadlines across GST, Income Tax, TDS, and Ministry of Corporate Affairs (MCA)
            </p>
          </div>

          <div className="flex items-center gap-2 shrink-0">
            <span className="text-xs font-semibold text-[#C43D43] bg-[#FDF2F2] border border-[#FDE8E8] px-2.5 py-1 rounded">
              1 Overdue Task Requires Immediate Action
            </span>
          </div>
        </div>

        {/* Act Filters */}
        <div className="flex items-center gap-2 pt-4 overflow-x-auto">
          {['All', 'GST', 'Income Tax', 'TDS', 'MCA / Companies Act'].map((act) => (
            <button
              key={act}
              onClick={() => setActiveActFilter(act)}
              className={`px-3 py-1.5 text-xs font-medium rounded border transition-colors shrink-0 ${
                activeActFilter === act
                  ? 'bg-gray-100 text-gray-900 border-gray-200'
                  : 'bg-white text-gray-600 border-gray-200 hover:bg-gray-50'
              }`}
            >
              {act}
            </button>
          ))}
        </div>
      </div>

      {/* Deadlines List */}
      <div className="bg-white rounded border border-gray-200 overflow-hidden">
        <div className="divide-y divide-[#E5E7EB]">
          {filteredDeadlines.map((dl) => {
            const isOverdue = dl.status === 'Overdue';
            const isDueSoon = dl.status === 'Due Soon';
            const isFiled = dl.status === 'Filed';

            return (
              <div
                key={dl.id}
                className={`p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-4 transition-colors ${
                  isOverdue
                    ? 'bg-[#FDF2F2]/40 hover:bg-[#FDF2F2]/60'
                    : isDueSoon
                    ? 'bg-[#FEF9EE]/30 hover:bg-[#FEF9EE]/60'
                    : 'hover:bg-gray-50'
                }`}
              >
                <div className="flex items-start gap-3 min-w-0">
                  <div
                    className={`w-9 h-9 rounded flex items-center justify-center shrink-0 mt-0.5 ${
                      isOverdue
                        ? 'bg-[#FDF2F2] text-[#C43D43] border border-[#FDE8E8]'
                        : isDueSoon
                        ? 'bg-[#FEF9EE] text-[#9A6700] border border-[#FBEEC8]'
                        : isFiled
                        ? 'bg-gray-100 text-gray-900 border border-gray-200'
                        : 'bg-[#F4F7FB] text-[#203B63] border border-[#E2E8F0]'
                    }`}
                  >
                    {isOverdue ? (
                      <AlertOctagon className="w-4 h-4" />
                    ) : isFiled ? (
                      <CheckCircle2 className="w-4 h-4" />
                    ) : (
                      <Calendar className="w-4 h-4" />
                    )}
                  </div>

                  <div>
                    <div className="flex items-center gap-2 flex-wrap">
                      <span className="text-xs font-bold text-[#111827]">{dl.title}</span>
                      <Badge
                        variant={
                          isOverdue ? 'error' : isDueSoon ? 'warning' : isFiled ? 'success' : 'navy'
                        }
                      >
                        {dl.formNo}
                      </Badge>
                      <span className="text-[11px] text-gray-600">• {dl.act}</span>
                    </div>

                    <div className="flex items-center gap-3 text-[11px] text-gray-600 mt-1.5 flex-wrap">
                      <span>
                        Statutory Due Date: <strong className="text-[#111827] font-semibold">{dl.dueDate}</strong>
                      </span>
                      <span>•</span>
                      <span>Assigned: <strong>{dl.assignedPartner}</strong></span>
                      <span>•</span>
                      <span>
                        Applicable Clients:{' '}
                        <strong>
                          {dl.applicableClients
                            .map((cId: string) => clients.find((c) => c.id === cId)?.name || cId)
                            .join(', ')}
                        </strong>
                      </span>
                    </div>
                  </div>
                </div>

                <div className="flex items-center gap-3 shrink-0 self-end sm:self-center">
                  <div className="text-right">
                    {isFiled ? (
                      <span className="text-xs font-semibold text-gray-900 flex items-center gap-1">
                        <CheckCircle2 className="w-3.5 h-3.5" /> Filed & Compliant
                      </span>
                    ) : isOverdue ? (
                      <span className="text-xs font-bold text-[#C43D43] block">
                        {Math.abs(dl.daysRemaining)} Days Overdue
                      </span>
                    ) : isDueSoon ? (
                      <span className="text-xs font-bold text-[#9A6700] block">
                        Due in {dl.daysRemaining} days
                      </span>
                    ) : (
                      <span className="text-xs text-gray-600 block">
                        In {dl.daysRemaining} days
                      </span>
                    )}
                    <span className="text-[10px] text-gray-600 block mt-0.5">Penal interest applies</span>
                  </div>

                  {!isFiled && (
                    <button
                      onClick={() => handleMarkCompleted(dl.id)}
                      className="px-3 py-1.5 text-xs font-medium bg-gray-900 hover:bg-black text-white rounded flex items-center gap-1 shadow-sm transition-colors"
                    >
                      <Check className="w-3.5 h-3.5" />
                      <span>Mark Filed</span>
                    </button>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
};
