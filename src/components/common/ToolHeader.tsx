import React from 'react';
import { ArrowLeft } from 'lucide-react';
import { InnerToolBadge } from './InnerToolBadge';

export interface ToolHeaderProps {
  toolId: string;
  title: string;
  description: string;
  onBack: () => void;
  badge?: string;
  actionButtons?: React.ReactNode;
  hidePrintHeader?: boolean;
}

export const ToolHeader: React.FC<ToolHeaderProps> = ({
  toolId,
  title,
  description,
  onBack,
  badge: _badge,
  actionButtons,
  hidePrintHeader = false,
}) => {
  return (
    <>
      {/* 1. Official Chartered Accountant Print Letterhead (Rendered ONLY when printing) */}
      {!hidePrintHeader && (
        <div className="hidden print:block pb-4 mb-5 border-b-2 border-gray-900">
          <div className="flex items-start justify-between">
            <div>
              <div className="text-[10px] font-bold tracking-widest text-gray-600 uppercase">
                Chartered Accountants Statutory Audit & Compliance Working Paper
              </div>
              <h1 className="text-xl font-black text-gray-900 tracking-tight mt-1">{title}</h1>
              <p className="text-[11px] text-gray-600 mt-0.5">{description}</p>
            </div>
            <div className="text-right text-[10px] text-gray-700 space-y-0.5 shrink-0 pl-4">
              <div className="font-bold text-gray-900">ICAI Standards / SA-230 Documentation</div>
              <div>Generated: {new Date().toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' })}</div>
              <div className="text-[9px] text-gray-500">Classification: Official Working Paper</div>
            </div>
          </div>
        </div>
      )}

      {/* 2. Interactive Screen UI Header (Hidden on print) */}
      <div className="mb-6 print:hidden">
        {/* Top Back Navigation Button */}
        <button
          onClick={onBack}
          className="inline-flex items-center gap-2 px-3 py-1.5 rounded-lg text-xs font-semibold text-gray-700 bg-white border border-gray-200 shadow-2xs hover:bg-gray-50 hover:text-gray-900 hover:border-gray-300 transition-all cursor-pointer mb-5 group"
        >
          <ArrowLeft className="w-3.5 h-3.5 text-gray-400 group-hover:text-gray-900 group-hover:-translate-x-0.5 transition-transform" />
          <span>Back to All Tools</span>
        </button>

        {/* Main Tool Identity & Action Row */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-5 border-b border-gray-200">
          <div className="flex items-center gap-3.5">
            <InnerToolBadge toolId={toolId} />
            <div>
              <h1 className="text-xl font-bold text-gray-900 tracking-tight">{title}</h1>
              <p className="text-xs text-gray-500 mt-0.5">{description}</p>
            </div>
          </div>
          {actionButtons && (
            <div className="flex items-center gap-2 flex-wrap">
              {actionButtons}
            </div>
          )}
        </div>
      </div>
    </>
  );
};
